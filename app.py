import streamlit as st
import pandas as pd
import json
import uuid
import time
import re
import copy
from datetime import datetime, timedelta

# =====================================================================
# 页面配置
# =====================================================================
st.set_page_config(
    page_title="IATF 审计转换工具 (v69.0 文件清单修复版)",
    page_icon="🛡️",
    layout="wide"
)

# =====================================================================
# 侧边栏：模板与模式配置
# =====================================================================
with st.sidebar:
    st.header("⚙️ 全局配置")
    st.divider()
    
    st.markdown("### 🔍 提取模式选择")
    run_mode = st.radio(
        "请根据报告类型选择：",
        (
            "纯净标准模式 (无附属场所)", 
            "单提取：EMS 扩展场所", 
            "单提取：RL 支持场所",
            "全量综合模式 (提取 EMS + RL + 被支持场所)"
        ),
        index=0
    )
    st.divider()
    
    st.info("💡 请上传您的 JSON 模板。程序将把该文件作为完整的底层骨架。")
    user_template_file = st.file_uploader("上传基础 JSON 模板", type=["json"])
    
    base_template_data = None
    if user_template_file:
        try:
            base_template_data = json.load(user_template_file)
            st.success(f"✅ 已加载底座: {user_template_file.name}")
        except Exception as e:
            st.error(f"❌ 解析失败: {e}")
            st.stop()
    else:
        st.warning("👈 请先上传底座文件以启动程序。")
        st.stop()

# =====================================================================
# 通用辅助函数区
# =====================================================================
def ensure_path(d, path):
    current = d
    for key in path:
        if key not in current or not isinstance(current[key], dict):
            current[key] = {}
        current = current[key]
    return current

def safe_get(obj, key, default=""):
    if isinstance(obj, dict):
        return obj.get(key, default)
    return default

def extract_and_format_english_name(raw_val):
    clean_val = str(raw_val).replace("姓名:", "").replace("Name:", "").strip()
    if not clean_val: return ""
    eng_only = re.sub(r'[^a-zA-Z\s]', ' ', clean_val).strip()
    eng_only = re.sub(r'\s+', ' ', eng_only)
    if eng_only:
        parts = eng_only.split()
        if len(parts) >= 2 and parts[0].isupper() and not parts[1].isupper():
            return f"{parts[1]} {parts[0]}"
        else:
            return eng_only
    return clean_val

# =====================================================================
# 独立模块 1：EMS 扩展场所提取器 (F21:M25)
# =====================================================================
def extract_ems_sites(info_df):
    ems_sites = []
    if info_df.empty: return ems_sites
    header_r = -1
    col_map = {}
    row_start, row_end = 20, min(25, info_df.shape[0])
    col_start, col_end = 5, min(13, info_df.shape[1])

    for r in range(row_start, row_end):
        for c in range(col_start, col_end):
            val = str(info_df.iloc[r, c]).strip().upper()
            if "EMS扩展场所信息" in val or "扩展制造场所" in val or "扩展现场" in val:
                header_r = r
                for c_scan in range(col_start, col_end):
                    h_val = str(info_df.iloc[r, c_scan]).strip()
                    if "中文名称" in h_val: col_map['name_cn'] = c_scan
                    elif "英文名称" in h_val: col_map['name_en'] = c_scan
                    elif "中文地址" in h_val: col_map['addr_cn'] = c_scan
                    elif "英文地址" in h_val: col_map['addr_en'] = c_scan
                    elif "邮编" in h_val or "邮政编码" in h_val: col_map['zip'] = c_scan
                    elif "USI" in h_val.upper(): col_map['usi'] = c_scan
                    elif "人数" in h_val: col_map['emp'] = c_scan
                break
        if header_r != -1: break
            
    if header_r != -1:
        for r in range(header_r + 1, row_end):
            def safe_get_cell(row, col_idx):
                if col_idx == -1 or col_idx >= info_df.shape[1]: return ""
                v = str(info_df.iloc[row, col_idx]).strip()
                return "" if v.lower() == 'nan' else v

            name_cn = safe_get_cell(r, col_map.get('name_cn', -1))
            name_en = safe_get_cell(r, col_map.get('name_en', -1))
            addr_cn = safe_get_cell(r, col_map.get('addr_cn', -1))
            
            if not name_cn and not addr_cn: continue
            if "名称" in name_cn and "地址" in addr_cn: continue
            
            full_site_name = name_cn
            if name_en and name_en not in name_cn:
                full_site_name = f"{name_cn} {name_en}".strip()

            addr_en = safe_get_cell(r, col_map.get('addr_en', -1))
            zip_code = safe_get_cell(r, col_map.get('zip', -1))
            usi = safe_get_cell(r, col_map.get('usi', -1))
            emp = safe_get_cell(r, col_map.get('emp', -1))

            ems_street, ems_city, ems_state, ems_country = addr_en, "", "", ""
            if addr_en:
                clean_eng = addr_en.replace('，', ',')
                parts = [p.strip() for p in clean_eng.split(',') if p.strip()]
                if len(parts) >= 3:
                    ems_country = parts[-1]
                    ems_state = parts[-2]
                    ems_city = parts[-3]
                    ems_street = ", ".join(parts[:-3])
                else:
                    ems_street = addr_en

            site_obj = {
                "Id": str(uuid.uuid4()),
                "SiteName": full_site_name,
                "IATF_USI": usi,
                "Usi": usi,
                "TotalNumberEmployees": emp,
                "AddressNative": {"Street1": addr_cn, "City": "", "State": "", "Country": "中国", "PostalCode": zip_code},
                "Address": {"Street1": ems_street, "City": ems_city, "State": ems_state, "Country": ems_country, "PostalCode": zip_code}
            }
            ems_sites.append(site_obj)
    return ems_sites

# =====================================================================
# 独立模块 2：RL 支持场所提取器 (F27:N32)
# =====================================================================
def extract_rl_sites(info_df):
    support_sites = []
    if info_df.empty: return support_sites
    header_r = -1
    col_map = {}
    rl_row_start, rl_row_end = 26, min(32, info_df.shape[0])
    rl_col_start, rl_col_end = 5, min(14, info_df.shape[1])

    for r in range(rl_row_start, rl_row_end):
        for c in range(rl_col_start, rl_col_end):
            val = str(info_df.iloc[r, c]).strip().upper()
            if ("支持场所" in val or "RL" in val) and "被" not in val:
                header_r = r
                for c_scan in range(rl_col_start, rl_col_end):
                    h_val = str(info_df.iloc[r, c_scan]).strip()
                    if "中文名称" in h_val: col_map['name_cn'] = c_scan
                    elif "英文名称" in h_val: col_map['name_en'] = c_scan
                    elif "中文地址" in h_val: col_map['addr_cn'] = c_scan
                    elif "英文地址" in h_val: col_map['addr_en'] = c_scan
                    elif "邮编" in h_val or "邮政编码" in h_val: col_map['zip'] = c_scan
                    elif "USI" in h_val.upper(): col_map['usi'] = c_scan
                    elif "人数" in h_val: col_map['emp'] = c_scan
                    elif "支持功能" in h_val: col_map['func'] = c_scan
                break
        if header_r != -1: break
            
    if header_r != -1:
        for r in range(header_r + 1, rl_row_end):
            def safe_get_cell(row, col_idx):
                if col_idx == -1 or col_idx >= info_df.shape[1]: return ""
                v = str(info_df.iloc[row, col_idx]).strip()
                return "" if v.lower() == 'nan' else v

            name_cn = safe_get_cell(r, col_map.get('name_cn', -1))
            name_en = safe_get_cell(r, col_map.get('name_en', -1))
            addr_cn = safe_get_cell(r, col_map.get('addr_cn', -1))
            
            if not name_cn and not addr_cn: continue
            if "名称" in name_cn and "地址" in addr_cn: continue
            
            full_site_name = name_cn
            if name_en and name_en not in name_cn:
                full_site_name = f"{name_cn} {name_en}".strip()

            addr_en = safe_get_cell(r, col_map.get('addr_en', -1))
            zip_code = safe_get_cell(r, col_map.get('zip', -1))
            usi = safe_get_cell(r, col_map.get('usi', -1))
            emp = safe_get_cell(r, col_map.get('emp', -1))
            func = safe_get_cell(r, col_map.get('func', -1))

            rl_street, rl_city, rl_state, rl_country = addr_en, "", "", ""
            if addr_en:
                clean_eng = addr_en.replace('，', ',')
                parts = [p.strip() for p in clean_eng.split(',') if p.strip()]
                if len(parts) >= 3:
                    rl_country = parts[-1]
                    rl_state = parts[-2]
                    rl_city = parts[-3]
                    rl_street = ", ".join(parts[:-3])
                else:
                    rl_street = addr_en

            site_obj = {
                "Id": str(uuid.uuid4()),
                "SiteName": full_site_name,
                "Comments": func,
                "IATF_USI": usi,
                "Usi": usi,
                "TotalNumberEmployees": emp,
                "AddressNative": {"Street1": addr_cn, "City": "", "State": "", "Country": "中国", "PostalCode": zip_code},
                "Address": {"Street1": rl_street, "City": rl_city, "State": rl_state, "Country": rl_country, "PostalCode": zip_code}
            }
            support_sites.append(site_obj)
    return support_sites

# =====================================================================
# 独立模块 3：被支持场所提取器 (F34:N38)
# =====================================================================
def extract_receiving_sites(info_df):
    receiving_sites = []
    if info_df.empty: return receiving_sites
    header_r = -1
    col_map = {}
    
    rec_row_start, rec_row_end = 33, min(38, info_df.shape[0])
    rec_col_start, rec_col_end = 5, min(14, info_df.shape[1])

    for r in range(rec_row_start, rec_row_end):
        for c in range(rec_col_start, rec_col_end):
            val = str(info_df.iloc[r, c]).strip().upper()
            if "被支持场所" in val:
                header_r = r
                for c_scan in range(rec_col_start, rec_col_end):
                    h_val = str(info_df.iloc[r, c_scan]).strip()
                    if "中文名称" in h_val: col_map['name_cn'] = c_scan
                    elif "英文名称" in h_val: col_map['name_en'] = c_scan
                    elif "中文地址" in h_val: col_map['addr_cn'] = c_scan
                    elif "英文地址" in h_val: col_map['addr_en'] = c_scan
                    elif "邮编" in h_val or "邮政编码" in h_val: col_map['zip'] = c_scan
                    elif "USI" in h_val.upper(): col_map['usi'] = c_scan
                    elif "人数" in h_val: col_map['emp'] = c_scan
                    elif "支持功能" in h_val: col_map['func'] = c_scan
                break
        if header_r != -1: break
            
    if header_r != -1:
        for r in range(header_r + 1, rec_row_end):
            def safe_get_cell(row, col_idx):
                if col_idx == -1 or col_idx >= info_df.shape[1]: return ""
                v = str(info_df.iloc[row, col_idx]).strip()
                return "" if v.lower() == 'nan' else v

            name_cn = safe_get_cell(r, col_map.get('name_cn', -1))
            name_en = safe_get_cell(r, col_map.get('name_en', -1))
            addr_cn = safe_get_cell(r, col_map.get('addr_cn', -1))
            
            if not name_cn and not addr_cn: continue
            if "名称" in name_cn and "地址" in addr_cn: continue
            
            full_site_name = name_cn
            if name_en and name_en not in name_cn:
                full_site_name = f"{name_cn} {name_en}".strip()

            addr_en = safe_get_cell(r, col_map.get('addr_en', -1))
            zip_code = safe_get_cell(r, col_map.get('zip', -1))
            usi = safe_get_cell(r, col_map.get('usi', -1))
            emp = safe_get_cell(r, col_map.get('emp', -1))
            func = safe_get_cell(r, col_map.get('func', -1))

            rec_street, rec_city, rec_state, rec_country = addr_en, "", "", ""
            if addr_en:
                clean_eng = addr_en.replace('，', ',')
                parts = [p.strip() for p in clean_eng.split(',') if p.strip()]
                if len(parts) >= 3:
                    rec_country = parts[-1]
                    rec_state = parts[-2]
                    rec_city = parts[-3]
                    rec_street = ", ".join(parts[:-3])
                else:
                    rec_street = addr_en

            site_obj = {
                "Id": str(uuid.uuid4()),
                "SiteName": full_site_name,
                "Comments": func,
                "IATF_USI": usi,
                "Usi": usi,
                "TotalNumberEmployees": emp,
                "AddressNative": {"Street1": addr_cn, "City": "", "State": "", "Country": "中国", "PostalCode": zip_code},
                "Address": {"Street1": rec_street, "City": rec_city, "State": rec_state, "Country": rec_country, "PostalCode": zip_code}
            }
            receiving_sites.append(site_obj)
    return receiving_sites

# =====================================================================
# 主流程区：核心转换逻辑
# =====================================================================
def generate_json_logic(excel_file, base_data, mode):
    final_json = copy.deepcopy(base_data)
    
    try:
        xls = pd.ExcelFile(excel_file)
        db_df = pd.read_excel(xls, sheet_name='数据库', header=None) if '数据库' in xls.sheet_names else pd.read_excel(xls, sheet_name=0, header=None)
        proc_df = pd.read_excel(xls, sheet_name='过程清单') if '过程清单' in xls.sheet_names else pd.DataFrame()
        info_df = pd.read_excel(xls, sheet_name='信息', header=None) if '信息' in xls.sheet_names else pd.DataFrame()
        
        # 💥 优化：主动按名字寻找“文件清单”，找不到再用备用逻辑
        if '文件清单' in xls.sheet_names:
            doc_list_df = pd.read_excel(xls, sheet_name='文件清单', header=None)
        else:
            doc_list_df = pd.read_excel(xls, sheet_name=xls.sheet_names[8], header=None) if len(xls.sheet_names) >= 9 else pd.DataFrame()
    except Exception as e:
        raise ValueError(f"Excel 读取失败: {str(e)}")

    def find_val_by_key(df, keywords, col_offset=1):
        if df.empty: return ""
        for r in range(df.shape[0]):
            for c in range(df.shape[1]):
                cell_val = str(df.iloc[r, c]).strip()
                for k in keywords:
                    if k in cell_val:
                        if c + col_offset < df.shape[1]:
                            return str(df.iloc[r, c + col_offset]).strip()
        return ""
        
    def get_db_val(r, c):
        try:
            val = db_df.iloc[r, c]
            return str(val).strip() if pd.notna(val) else ""
        except: return ""

    raw_name_full = find_val_by_key(db_df, ["姓名", "Auditor Name"]) or get_db_val(5, 1)
    raw_name = raw_name_full.replace("姓名:", "").replace("Name:", "").strip() if raw_name_full else ""
    formatted_team_name = extract_and_format_english_name(raw_name_full)

    ccaa_raw = find_val_by_key(db_df, ["审核员CCAA", "CCAA"]) or get_db_val(4, 1)
    caa_no = ""
    if ccaa_raw:
        match = re.search(r'(?:CCAA[:：\s-])\s*(.*)', ccaa_raw, re.IGNORECASE | re.DOTALL)
        caa_no = match.group(1).strip() if match else ccaa_raw.strip()

    auditor_id = ""
    if not info_df.empty:
        for r in range(info_df.shape[0]):
            for c in range(info_df.shape[1]):
                cell_text = str(info_df.iloc[r, c])
                if "IATF Card" in cell_text or "IATF卡号" in cell_text:
                    if c + 1 < info_df.shape[1]:
                        raw_val = str(info_df.iloc[r, c + 1]).strip()
                        raw_val = raw_val.replace('\n', ' ').replace('\r', ' ')
                        auditor_id = re.sub(r'^IATF[:：\s-]*', '', raw_val, flags=re.IGNORECASE).strip()
                        if len(auditor_id) > 4: break
            if auditor_id and len(auditor_id) > 4: break

    start_date_raw = find_val_by_key(db_df, ["审核开始日期", "审核开始时间"]) or get_db_val(2, 1)
    end_date_raw = find_val_by_key(db_df, ["审核结束日期", "审核结束时间"]) or get_db_val(3, 1)
    
    def fmt_iso(val):
        try:
            clean_val = str(val).replace('年', '-').replace('月', '-').replace('日', '')
            dt = pd.to_datetime(clean_val, errors='coerce')
            if pd.notna(dt): return dt.strftime('%Y-%m-%d') + "T00:00:00.000Z"
        except: pass
        return ""
        
    start_iso, end_iso = fmt_iso(start_date_raw), fmt_iso(end_date_raw)
    
    next_audit_iso = ""
    try:
        clean_end = str(end_date_raw).replace('年', '-').replace('月', '-').replace('日', '')
        end_dt = pd.to_datetime(clean_end, errors='coerce')
        if pd.notna(end_dt): next_audit_iso = (end_dt + timedelta(days=45)).strftime('%Y-%m-%d') + "T00:00:00.000Z"
    except: pass

    customers_list = []
    if not info_df.empty:
        header_r = -1
        col_map = {'cust': -1, 'name': -1, 'date': -1, 'code': -1}
        for r in range(info_df.shape[0]):
            row_str = " ".join([str(x) for x in info_df.iloc[r, :]]).upper()
            if "CUSTOMER" in row_str and ("CSR" in row_str or "TITLE" in row_str):
                header_r = r
                for c in range(info_df.shape[1]):
                    val = str(info_df.iloc[r, c]).strip().upper()
                    if "CUSTOMER" in val or "客户" in val: col_map['cust'] = c
                    elif "CSR" in val or "TITLE" in val: col_map['name'] = c
                    elif "VERSION" in val or "DATE" in val or "版本" in val or "日期" in val: col_map['date'] = c
                    elif "供应商代码" in val or "SUPPLIER" in val or "CODE" in val: col_map['code'] = c
                break
                
        if header_r != -1:
            for r in range(header_r + 1, info_df.shape[0]):
                cust_val = str(info_df.iloc[r, col_map['cust']]).strip() if col_map['cust'] != -1 else ""
                if not cust_val or cust_val.lower() == 'nan': continue
                if "审核员" in cust_val or "AUDIT" in cust_val.upper() or "NAME" in cust_val.upper(): break
                    
                name_val = str(info_df.iloc[r, col_map['name']]).strip() if col_map['name'] != -1 else ""
                date_val = str(info_df.iloc[r, col_map['date']]).strip() if col_map['date'] != -1 else ""
                code_val = str(info_df.iloc[r, col_map['code']]).strip() if col_map['code'] != -1 else ""
                
                final_date = date_val.replace(" 00:00:00", "").strip()
                customers_list.append({
                    "Name": cust_val, "SupplierCode": code_val, "NameCSRDocument": name_val, "DateCSRDocument": final_date
                })

    if not customers_list:
        customer_name = find_val_by_key(db_df, ["顾客", "客户名称"]) or get_db_val(29, 1)
        supplier_code = find_val_by_key(db_df, ["供应商编码", "供应商代码"]) or get_db_val(30, 1)
        csr_name = find_val_by_key(db_df, ["CSR文件名称"]) or get_db_val(31, 1)
        csr_date_raw = find_val_by_key(db_df, ["CSR文件日期"]) or get_db_val(32, 1)
        csr_date = str(csr_date_raw).replace(" 00:00:00", "").strip()
        if csr_date.lower() == 'nan': csr_date = ""
        if customer_name or supplier_code or csr_name:
            customers_list.append({
                "Name": customer_name, "SupplierCode": supplier_code, "NameCSRDocument": csr_name, "DateCSRDocument": csr_date
            })

    # 主地址混合剥离扫描
    english_address = ""
    native_street = ""
    cands = []
    if not db_df.empty:
        for r_idx in range(9, 14):
            if r_idx < db_df.shape[0]:
                if 1 < db_df.shape[1]: cands.append(str(db_df.iloc[r_idx, 1]))
                if 4 < db_df.shape[1]: cands.append(str(db_df.iloc[r_idx, 4]))
                
    def get_anchored(df, keywords):
        res = []
        if df.empty: return res
        for r in range(df.shape[0]):
            for c in range(df.shape[1]):
                val = str(df.iloc[r, c]).strip().upper()
                if any(k in val for k in keywords):
                    res.append(str(df.iloc[r, c])) 
                    if c + 1 < df.shape[1]: res.append(str(df.iloc[r, c+1]))
                    if c + 2 < df.shape[1]: res.append(str(df.iloc[r, c+2]))
                    if r + 1 < df.shape[0]: res.append(str(df.iloc[r+1, c]))
                    if r + 1 < df.shape[0] and c+1 < df.shape[1]: res.append(str(df.iloc[r+1, c+1]))
        return res
        
    cands += get_anchored(info_df, ["审核地址", "AUDIT ADDRESS", "ADDRESS"])
    cands += get_anchored(db_df, ["地址", "ADDRESS"])
    
    en_parts, zh_parts = [], []
    for cand in cands:
        cand = str(cand).strip()
        if not cand or cand.lower() == 'nan': continue
        cand = re.sub(r'^(审核地址|组织地址|企业地址|地址|现场地址|AUDIT ADDRESS|ADDRESS)[\s:：]*', '', cand, flags=re.IGNORECASE).strip()
        if not cand: continue
        
        lines = cand.replace('\r', '\n').split('\n')
        for line in lines:
            line = line.strip()
            if not line: continue
            
            has_zh = bool(re.search(r'[\u4e00-\u9fff]', line))
            has_en = bool(re.search(r'[a-zA-Z]{3,}', line)) 
            
            if has_zh and has_en:
                en_str = re.sub(r'[\u4e00-\u9fff]', ' ', line)
                en_str = re.sub(r'[，。；（）]', ' ', en_str) 
                en_str = re.sub(r'\s+', ' ', en_str).strip(" ()-.,")
                zh_str = re.sub(r'[a-zA-Z]', '', line)
                zh_str = re.sub(r'\s+', ' ', zh_str).strip(" ()-.,")
                
                if len(en_str) > 10: en_parts.append(en_str)
                if len(zh_str) > 5: zh_parts.append(zh_str)
            elif has_zh: zh_parts.append(line)
            elif has_en: en_parts.append(line)

    english_address = max(en_parts, key=len) if en_parts else ""
    native_street = max(zh_parts, key=len) if zh_parts else ""

    street, city, state, country = english_address, "", "", ""
    if english_address:
        clean_eng = english_address.replace('，', ',')
        parts = [p.strip() for p in clean_eng.split(',') if p.strip()]
        if len(parts) >= 3:
            country = parts[-1]
            state = parts[-2]
            city = parts[-3]
            street = ", ".join(parts[:-3])
        else:
            street = english_address

    final_json["uuid"] = str(uuid.uuid4())
    final_json["created"] = int(time.time() * 1000)

    ensure_path(final_json, ["AuditData", "AuditDate"])
    if start_iso: final_json["AuditData"]["AuditDate"]["Start"] = start_iso
    if end_iso: final_json["AuditData"]["AuditDate"]["End"] = end_iso
    final_json["AuditData"]["CbIdentificationNo"] = find_val_by_key(db_df, ["认证机构标识号"]) or get_db_val(2, 4)
    final_json["AuditData"]["AuditorName"] = raw_name
    final_json["AuditData"]["auditorname"] = raw_name

    if "AuditTeam" not in final_json["AuditData"] or not isinstance(final_json["AuditData"]["AuditTeam"], list) or len(final_json["AuditData"]["AuditTeam"]) == 0:
        final_json["AuditData"]["AuditTeam"] = [{}]
        
    team = final_json["AuditData"]["AuditTeam"][0]
    if isinstance(team, dict):
        team.update({
            "Name": formatted_team_name, "CaaNo": caa_no, "AuditorId": auditor_id, 
            "AuditDaysPerformed": 1.5, "DatesOnSite": [{"Date": start_iso, "Day": 1}, {"Date": end_iso, "Day": 0.5}]
        })

    ensure_path(final_json, ["OrganizationInformation", "AddressNative"])
    ensure_path(final_json, ["OrganizationInformation", "Address"])
    org = final_json["OrganizationInformation"]
    
    org["OrganizationName"] = find_val_by_key(db_df, ["组织名称"]) or get_db_val(1, 4)
    org["IndustryCode"] = find_val_by_key(db_df, ["行业代码", "Industry Code"])
    org["IATF_USI"] = find_val_by_key(db_df, ["IATF USI", "USI"]) or get_db_val(3, 4)
    org["TotalNumberEmployees"] = find_val_by_key(db_df, ["包括扩展现场在内的员工总数", "员工总数"]) or get_db_val(27, 1)
    org["CertificateScope"] = find_val_by_key(db_df, ["证书范围"])
    org["Representative"] = find_val_by_key(db_df, ["组织代表", "管理者代表", "联系人", "Representative"]) or get_db_val(15, 1)
    org["Telephone"] = find_val_by_key(db_df, ["联系电话", "电话", "Telephone"]) or get_db_val(15, 4)
    extracted_email = find_val_by_key(db_df, ["电子邮箱", "邮箱", "Email", "E-mail"]) or get_db_val(16, 1)
    org["Email"] = "" if str(extracted_email).strip() == "0" else extracted_email
    
    if "LanguageByManufacturingPersonnel" in org:
        lang_node = org["LanguageByManufacturingPersonnel"]
        if isinstance(lang_node, list) and len(lang_node) > 0:
            if isinstance(lang_node[0], dict): lang_node[0]["Products"] = ""
        elif isinstance(lang_node, dict):
            if "0" in lang_node and isinstance(lang_node["0"], dict): lang_node["0"]["Products"] = ""
            else: lang_node["Products"] = ""
    
    if native_street:
        org["AddressNative"]["Street1"] = native_street
    org["AddressNative"]["Country"] = "中国"
    
    if english_address:
        org["Address"]["State"] = state
        org["Address"]["City"] = city
        org["Address"]["Country"] = country if country else "China"
        org["Address"]["Street1"] = street
        
    postal_code = find_val_by_key(db_df, ["邮政编码"]) or get_db_val(10, 4)
    if postal_code:
        org["AddressNative"]["PostalCode"] = postal_code
        org["Address"]["PostalCode"] = postal_code

    # =====================================================================
    # 根据模式拔插模块
    # =====================================================================
    if "全量综合模式" in mode:
        ems_sites = extract_ems_sites(info_df)
        if ems_sites:
            final_json["ExtendedManufacturingSites"] = ems_sites
            org["ExtendedManufacturingSite"] = "1"
        else:
            org["ExtendedManufacturingSite"] = "0"
            
        support_sites = extract_rl_sites(info_df)
        if support_sites:
            final_json["ProvidingSupportSites"] = support_sites
            
        receiving_sites = extract_receiving_sites(info_df)
        if receiving_sites:
            final_json["ReceivingSupportSites"] = receiving_sites
            
    elif "EMS" in mode:
        ems_sites = extract_ems_sites(info_df)
        if ems_sites:
            final_json["ExtendedManufacturingSites"] = ems_sites
            org["ExtendedManufacturingSite"] = "1"
        else:
            org["ExtendedManufacturingSite"] = "0"
            
    elif "RL" in mode:
        org["ExtendedManufacturingSite"] = "0"
        support_sites = extract_rl_sites(info_df)
        if support_sites:
            final_json["ProvidingSupportSites"] = support_sites
            
    else:
        org["ExtendedManufacturingSite"] = "0"

    ensure_path(final_json, ["CustomerInformation"])
    final_json["CustomerInformation"]["Customers"] = []
    for c_info in customers_list:
        cust_obj = {
            "Id": str(uuid.uuid4()), "Name": c_info["Name"], "SupplierCode": c_info["SupplierCode"],
            "Csrs": [{"Id": str(uuid.uuid4()), "Name": c_info["Name"], "SupplierCode": c_info["SupplierCode"],
                      "NameCSRDocument": c_info["NameCSRDocument"], "DateCSRDocument": c_info["DateCSRDocument"]}]
        }
        final_json["CustomerInformation"]["Customers"].append(cust_obj)

    # 💥💥💥 [重构：文件清单精准映射抓取逻辑] 💥💥💥
    doc_map = {}
    if not doc_list_df.empty:
        clause_col = -1
        doc_col = -1
        header_r = -1
        # 寻找“条款”列和“名称”列
        for r in range(min(10, doc_list_df.shape[0])):
            for c in range(doc_list_df.shape[1]):
                val = str(doc_list_df.iloc[r, c]).strip()
                if "条款" in val or "标准条款" in val:
                    clause_col = c
                if "公司内对应的程序文件" in val or "包含名称" in val or "文件名称" in val:
                    doc_col = c
            if clause_col != -1 and doc_col != -1:
                header_r = r
                break
        
        if header_r != -1:
            for r in range(header_r + 1, doc_list_df.shape[0]):
                clause_val = str(doc_list_df.iloc[r, clause_col]).strip()
                if not clause_val or clause_val.lower() == 'nan': continue
                
                # 仅提取最前面的数字和点 (如 "4.4.1.2产品安全" -> "4.4.1.2")
                match = re.match(r'^([\d\.]+)', clause_val)
                if match:
                    clause_no = match.group(1)
                    if clause_no.endswith('.'): clause_no = clause_no[:-1]
                    
                    doc_parts = []
                    # 动态合并被分在连续3列里的内容：名称、编号、版本
                    for dc in range(doc_col, min(doc_col + 3, doc_list_df.shape[1])):
                        part_val = str(doc_list_df.iloc[r, dc]).strip()
                        if part_val and part_val.lower() != 'nan':
                            doc_parts.append(part_val)
                    
                    if doc_parts:
                        # 用空格拼接，更符合阅读习惯
                        doc_map[clause_no] = " ".join(doc_parts)

    # 将映射后的文件清单精准注射进 JSON 骨架
    if doc_map and "Stage1DocumentedRequirements" in final_json and "IatfClauseDocuments" in final_json["Stage1DocumentedRequirements"]:
        clause_docs = final_json["Stage1DocumentedRequirements"]["IatfClauseDocuments"]
        for i in range(len(clause_docs)):
            if isinstance(clause_docs[i], dict):
                p_no = str(clause_docs[i].get("ProcessNo", ""))
                # 如果这个条款号在字典中找到了对应提取出来的文件信息，就更新覆盖
                if p_no in doc_map:
                    clause_docs[i]["DocumentName"] = doc_map[p_no]

    processes = []
    if not proc_df.empty:
        clause_cols = proc_df.columns[13:] if proc_df.shape[1] > 13 else []
        for idx, row in proc_df.iterrows():
            p_name = str(row.iloc[0]).strip()
            rep_name = str(row.iloc[2]).strip() if pd.notna(row.iloc[2]) else ""
            if not p_name or p_name.lower() == 'nan': continue
            proc_obj = {
                "Id": str(uuid.uuid4()), "ProcessName": p_name, "RepresentativeName": rep_name,
                "ManufacturingProcess": "0", "OnSiteProcess": "1", "RemoteProcess": "0",
                "AuditNotes": [{"Id": str(uuid.uuid4()), "AuditorId": auditor_id, "AuditorName": raw_name}]
            }
            for col in clause_cols:
                if str(row[col]).strip().upper() in ['X', 'TRUE']: proc_obj[col] = True
            processes.append(proc_obj)
    final_json["Processes"] = processes

    if "Results" not in final_json: final_json["Results"] = {}
    if "AuditReportFinal" not in final_json["Results"]: final_json["Results"]["AuditReportFinal"] = {}
    if end_iso: final_json["Results"]["AuditReportFinal"]["Date"] = end_iso
    if next_audit_iso: final_json["Results"]["DateNextScheduledAudit"] = next_audit_iso
    
    b6_raw_val = get_db_val(5, 1)
    b6_formatted_name = extract_and_format_english_name(b6_raw_val)
    final_json["Results"]["AuditReportFinal"]["AuditorName"] = b6_formatted_name

    # 附带回传提取到的文件个数以供界面显示
    return final_json, len(doc_map)

# =====================================================================
# 主界面展示区
# =====================================================================

st.title("🛡️ 多模板审计转换引擎 (v69.0 文件清单终极修复版)")
st.markdown(f"💡 **当前运行模式**: `{run_mode}`")

st.markdown("### 📥 上传数据源")
uploaded_files = st.file_uploader("支持批量上传 .xlsx 格式文件", type=["xlsx"], accept_multiple_files=True)

if uploaded_files:
    st.divider()
    
    for file in uploaded_files:
        try:
            res_json, mapped_doc_count = generate_json_logic(file, base_template_data, run_mode)
            st.success(f"✅ 解析成功：{file.name}")
            
            row_col1, row_col2 = st.columns([3, 1])
            
            with row_col1:
                with st.expander("👀 查看数据提取日志", expanded=True):
                     if "全量综合模式" in run_mode:
                         ems_count = len(res_json.get('ExtendedManufacturingSites', []))
                         rl_count = len(res_json.get('ProvidingSupportSites', []))
                         rec_count = len(res_json.get('ReceivingSupportSites', []))
                         st.code(f"""
[模块: 全量综合提取]
✅ EMS扩展场所提取: {ems_count} 个
✅ RL支持场所提取 : {rl_count} 个
✅ 被支持场所提取 : {rec_count} 个
✅ 文件清单精准映射: {mapped_doc_count} 条 (已对应填入JSON)
标志位(EMS): "{res_json.get('OrganizationInformation', {}).get('ExtendedManufacturingSite', '缺失')}"
                         """.strip(), language="yaml")
                         
                     elif "EMS" in run_mode:
                         try:
                             ems_sites = res_json.get('ExtendedManufacturingSites', [])
                             ems_count = len(ems_sites)
                             ems_sample = ems_sites[0] if ems_count > 0 else {}
                         except:
                             ems_count, ems_sample = 0, {}
                         st.code(f"""
[模块: EMS扩展场所]
提取数量: {ems_count} 个
场所名称: "{safe_get(ems_sample, 'SiteName', '无')}"
文件映射: {mapped_doc_count} 条
标志位: "{res_json.get('OrganizationInformation', {}).get('ExtendedManufacturingSite', '缺失')}"
                         """.strip(), language="yaml")
                         
                     elif "RL" in run_mode:
                         try:
                             rl_sites = res_json.get('ProvidingSupportSites', [])
                             rl_count = len(rl_sites)
                             rl_sample = rl_sites[0] if rl_count > 0 else {}
                         except:
                             rl_count, rl_sample = 0, {}
                         st.code(f"""
[模块: RL支持场所]
提取数量: {rl_count} 个
场所名称: "{safe_get(rl_sample, 'SiteName', '无')}"
文件映射: {mapped_doc_count} 条
                         """.strip(), language="yaml")
                         
                     else:
                         st.code(f"""
[模块: 纯净标准]
中文主地址: "{safe_get(res_json.get('OrganizationInformation', {}).get('AddressNative', {}), 'Street1', '缺失')}"
文件清单映射: {mapped_doc_count} 条目已准确写入
                         """.strip(), language="yaml")

            with row_col2:
                st.download_button(
                    label=f"📥 下载 JSON 文件",
                    data=json.dumps(res_json, indent=2, ensure_ascii=False),
                    file_name=file.name.replace(".xlsx", ".json"),
                    key=f"dl_{file.name}"
                )
        except Exception as e:
            st.error(f"❌ 解析 {file.name} 失败: {str(e)}")






















