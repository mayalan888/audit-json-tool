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
    page_title="IATF 审计转换工具 (v70.3 全量地址拆分版)",
    page_icon="🛡️",
    layout="wide"
)

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

def parse_chinese_address(addr_str):
    """
    专门解析中文地址的省、市、街道。
    支持直辖市、省、自治区等。
    """
    province, city, street = "", "", addr_str
    if not addr_str: return province, city, street

    clean_addr = re.sub(r'^中国', '', addr_str).strip()
    
    # 匹配省份/直辖市
    p_match = re.search(r'(.+?(省|自治区|北京|上海|天津|重庆))', clean_addr)
    if p_match:
        province = p_match.group(1).strip()
        if province in ["北京", "上海", "天津", "重庆"]: province += "市"
        remain_addr = clean_addr[len(p_match.group(1)):].strip()
        
        # 匹配市/州
        c_match = re.search(r'(.+?(市|地区|盟|自治州|州))', remain_addr)
        if c_match:
            city = c_match.group(1).strip()
            street = remain_addr[len(city):].strip()
        else:
            if "市" in province: city = province
            street = remain_addr
            
    return province, city, street

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
            def safe_get_cell(col_idx):
                if col_idx == -1 or col_idx >= info_df.shape[1]: return ""
                v = str(info_df.iloc[r, col_idx]).strip()
                return "" if v.lower() == 'nan' else v

            name_cn = safe_get_cell(col_map.get('name_cn', -1))
            name_en = safe_get_cell(col_map.get('name_en', -1))
            addr_cn = safe_get_cell(col_map.get('addr_cn', -1))
            
            if not name_cn and not addr_cn: continue
            if "名称" in name_cn and "地址" in addr_cn: continue
            
            full_site_name = name_cn
            if name_en and name_en not in name_cn:
                full_site_name = f"{name_cn} {name_en}".strip()

            addr_en = safe_get_cell(col_map.get('addr_en', -1))
            zip_code = safe_get_cell(col_map.get('zip', -1))
            usi = safe_get_cell(col_map.get('usi', -1))
            emp = safe_get_cell(col_map.get('emp', -1))

            # 💥 应用 EMS 中文地址拆分 💥
            ems_zh_p, ems_zh_c, ems_zh_s = parse_chinese_address(addr_cn)

            ems_street, ems_city, ems_state, ems_country = addr_en, "", "", ""
            if addr_en:
                parts = [p.strip() for p in addr_en.replace('，', ',').split(',') if p.strip()]
                if len(parts) >= 3:
                    ems_country, ems_state, ems_city = parts[-1], parts[-2], parts[-3]
                    ems_street = ", ".join(parts[:-3])

            site_obj = {
                "Id": str(uuid.uuid4()), "SiteName": full_site_name, "IATF_USI": usi, "Usi": usi, "TotalNumberEmployees": emp,
                "AddressNative": {"Street1": ems_zh_s, "City": ems_zh_c, "State": ems_zh_p, "Country": "中国", "PostalCode": zip_code},
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
            def safe_get_cell(col_idx):
                if col_idx == -1 or col_idx >= info_df.shape[1]: return ""
                v = str(info_df.iloc[r, col_idx]).strip()
                return "" if v.lower() == 'nan' else v

            name_cn = safe_get_cell(col_map.get('name_cn', -1))
            name_en = safe_get_cell(col_map.get('name_en', -1))
            addr_cn = safe_get_cell(col_map.get('addr_cn', -1))
            
            if not name_cn and not addr_cn: continue
            if "名称" in name_cn and "地址" in addr_cn: continue
            
            full_site_name = name_cn
            if name_en and name_en not in name_cn:
                full_site_name = f"{name_cn} {name_en}".strip()

            addr_en = safe_get_cell(col_map.get('addr_en', -1))
            zip_code = safe_get_cell(col_map.get('zip', -1))
            usi = safe_get_cell(col_map.get('usi', -1))
            emp = safe_get_cell(col_map.get('emp', -1))
            func = safe_get_cell(col_map.get('func', -1))

            # 💥 应用 RL 中文地址拆分 💥
            rl_zh_p, rl_zh_c, rl_zh_s = parse_chinese_address(addr_cn)

            rl_street, rl_city, rl_state, rl_country = addr_en, "", "", ""
            if addr_en:
                parts = [p.strip() for p in addr_en.replace('，', ',').split(',') if p.strip()]
                if len(parts) >= 3:
                    rl_country, rl_state, rl_city = parts[-1], parts[-2], parts[-3]
                    rl_street = ", ".join(parts[:-3])

            site_obj = {
                "Id": str(uuid.uuid4()), "SiteName": full_site_name, "Comments": func, "IATF_USI": usi, "Usi": usi, "TotalNumberEmployees": emp,
                "AddressNative": {"Street1": rl_zh_s, "City": rl_zh_c, "State": rl_zh_p, "Country": "中国", "PostalCode": zip_code},
                "Address": {"Street1": rl_street, "City": rl_city, "State": rl_state, "Country": rl_country, "PostalCode": zip_code}
            }
            support_sites.append(site_obj)
    return support_sites

# =====================================================================
# 独立模块 3：被支持场所提取器 (F34:N38) - 🔥 [修复漏缺，并增加地址拆分] 🔥
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
            def safe_get_cell(col_idx):
                if col_idx == -1 or col_idx >= info_df.shape[1]: return ""
                v = str(info_df.iloc[r, col_idx]).strip()
                return "" if v.lower() == 'nan' else v

            name_cn = safe_get_cell(col_map.get('name_cn', -1))
            name_en = safe_get_cell(col_map.get('name_en', -1))
            addr_cn = safe_get_cell(col_map.get('addr_cn', -1))
            
            if not name_cn and not addr_cn: continue
            if "名称" in name_cn and "地址" in addr_cn: continue
            
            full_site_name = name_cn
            if name_en and name_en not in name_cn:
                full_site_name = f"{name_cn} {name_en}".strip()

            addr_en = safe_get_cell(col_map.get('addr_en', -1))
            zip_code = safe_get_cell(col_map.get('zip', -1))
            usi = safe_get_cell(col_map.get('usi', -1))
            emp = safe_get_cell(col_map.get('emp', -1))
            func = safe_get_cell(col_map.get('func', -1))

            # 💥 应用被支持场所的中文地址拆分 💥
            rec_zh_p, rec_zh_c, rec_zh_s = parse_chinese_address(addr_cn)

            rec_street, rec_city, rec_state, rec_country = addr_en, "", "", ""
            if addr_en:
                parts = [p.strip() for p in addr_en.replace('，', ',').split(',') if p.strip()]
                if len(parts) >= 3:
                    rec_country, rec_state, rec_city = parts[-1], parts[-2], parts[-3]
                    rec_street = ", ".join(parts[:-3])

            site_obj = {
                "Id": str(uuid.uuid4()), "SiteName": full_site_name, "Comments": func, "IATF_USI": usi, "Usi": usi, "TotalNumberEmployees": emp,
                "AddressNative": {"Street1": rec_zh_s, "City": rec_zh_c, "State": rec_zh_p, "Country": "中国", "PostalCode": zip_code},
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
        perf_df = pd.read_excel(xls, sheet_name='过程绩效', header=None) if '过程绩效' in xls.sheet_names else pd.DataFrame()
        
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

    # 💥 KPI 防覆写提取逻辑 💥
    kpi_map = {}
    time_period = ""
    if not perf_df.empty:
        if perf_df.shape[0] > 1 and perf_df.shape[1] > 5:
            time_period = fmt_iso(str(perf_df.iloc[1, 5]).strip())
        
        header_r = -1
        col_map_kpi = {'proc': -1, 'kpi': -1, 'target': -1, 'result': -1, 'trend': -1}
        for r in range(min(10, perf_df.shape[0])):
            for c in range(perf_df.shape[1]):
                val = str(perf_df.iloc[r, c]).strip().upper()
                if val == "过程" or "KPI名称" in val:
                    header_r = r
                    for scan_c in range(perf_df.shape[1]):
                        h_val = str(perf_df.iloc[r, scan_c]).strip().upper()
                        if ("过程" == h_val or "PROCESS" in h_val) and col_map_kpi['proc'] == -1: col_map_kpi['proc'] = scan_c
                        elif ("KPI" in h_val or "指标" in h_val) and col_map_kpi['kpi'] == -1: col_map_kpi['kpi'] = scan_c
                        elif ("目标" in h_val or "TARGET" in h_val) and col_map_kpi['target'] == -1: col_map_kpi['target'] = scan_c
                        elif ("结果" in h_val or "RESULT" in h_val) and col_map_kpi['result'] == -1: col_map_kpi['result'] = scan_c
                        elif ("趋势" in h_val or "TREND" in h_val) and col_map_kpi['trend'] == -1: col_map_kpi['trend'] = scan_c
                    break
            if header_r != -1: break
            
        if header_r != -1:
            curr_proc = ""
            for r in range(header_r + 1, perf_df.shape[0]):
                p_v = str(perf_df.iloc[r, col_map_kpi['proc']]).strip() if col_map_kpi['proc'] != -1 else ""
                if p_v and p_v.lower() != 'nan': curr_proc = p_v
                k_v = str(perf_df.iloc[r, col_map_kpi['kpi']]).strip() if col_map_kpi['kpi'] != -1 else ""
                if not k_v or k_v.lower() == 'nan': continue
                
                t_v = str(perf_df.iloc[r, col_map_kpi['target']]).strip()
                r_v = str(perf_df.iloc[r, col_map_kpi['result']]).strip()
                tr_v = str(perf_df.iloc[r, col_map_kpi['trend']]).strip()
                
                trend_code = "0"
                if "积极" in tr_v or "1" == tr_v: trend_code = "1"
                elif "消极" in tr_v or "-1" == tr_v: trend_code = "-1"

                if curr_proc not in kpi_map: kpi_map[curr_proc] = []
                kpi_map[curr_proc].append({
                    "KPI": k_v, "CurrentTarget": t_v if t_v.lower() != 'nan' else "",
                    "Results": r_v if r_v.lower() != 'nan' else "", "TrendLastAudit": trend_code,
                    "TimePeriodFrom": time_period
                })

    cands = []
    if not db_df.empty:
        for r_idx in range(9, 14):
            if r_idx < db_df.shape[0]:
                if 1 < db_df.shape[1]: cands.append(str(db_df.iloc[r_idx, 1]))
                if 4 < db_df.shape[1]: cands.append(str(db_df.iloc[r_idx, 4]))
    
    en_parts, zh_parts = [], []
    for cand in cands:
        cand = str(cand).strip()
        if not cand or cand.lower() == 'nan': continue
        cand = re.sub(r'^(审核地址|组织地址|地址|现场地址|AUDIT ADDRESS|ADDRESS)[\s:：]*', '', cand, flags=re.IGNORECASE).strip()
        has_zh = bool(re.search(r'[\u4e00-\u9fff]', cand))
        has_en = bool(re.search(r'[a-zA-Z]{3,}', cand))
        if has_zh and has_en:
            zh_parts.append(re.sub(r'[a-zA-Z]', '', cand).strip(" ()-.,"))
            en_parts.append(re.sub(r'[\u4e00-\u9fff]', '', cand).strip(" ()-.,"))
        elif has_zh: zh_parts.append(cand)
        elif has_en: en_parts.append(cand)

    native_address_full = max(zh_parts, key=len) if zh_parts else ""
    english_address_full = max(en_parts, key=len) if en_parts else ""

    # 💥 主地址中文拆分 💥
    native_p, native_c, native_s = parse_chinese_address(native_address_full)

    en_street, en_city, en_state, en_country = english_address_full, "", "", "China"
    if english_address_full:
        parts = [p.strip() for p in english_address_full.replace('，', ',').split(',') if p.strip()]
        if len(parts) >= 3:
            en_country, en_state, en_city = parts[-1], parts[-2], parts[-3]
            en_street = ", ".join(parts[:-3])

    final_json["uuid"] = str(uuid.uuid4())
    ensure_path(final_json, ["AuditData", "AuditDate"])
    final_json["AuditData"]["AuditDate"].update({"Start": start_iso, "End": end_iso})
    final_json["AuditData"]["AuditorName"] = raw_name

    org = final_json["OrganizationInformation"]
    ensure_path(org, ["AddressNative"])
    ensure_path(org, ["Address"])
    
    org["AddressNative"].update({ "Street1": native_s, "City": native_c, "State": native_p, "Country": "中国" })
    org["Address"].update({ "Street1": en_street, "City": en_city, "State": en_state, "Country": en_country })
    
    postal = find_val_by_key(db_df, ["邮政编码"]) or get_db_val(10, 4)
    org["AddressNative"]["PostalCode"] = postal
    org["Address"]["PostalCode"] = postal

    # 💥 多模式场所数据挂载 (包含恢复的被支持场所) 💥
    if "全量综合模式" in mode:
        ems_sites = extract_ems_sites(info_df)
        if ems_sites:
            final_json["ExtendedManufacturingSites"] = ems_sites
            org["ExtendedManufacturingSite"] = "1"
        else:
            org["ExtendedManufacturingSite"] = "0"
            
        support_sites = extract_rl_sites(info_df)
        if support_sites: final_json["ProvidingSupportSites"] = support_sites
            
        receiving_sites = extract_receiving_sites(info_df)
        if receiving_sites: final_json["ReceivingSupportSites"] = receiving_sites
            
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
        if support_sites: final_json["ProvidingSupportSites"] = support_sites
            
    else:
        org["ExtendedManufacturingSite"] = "0"

    # KPI 映射
    processes = []
    total_kpi = 0
    if not proc_df.empty:
        for idx, row in proc_df.iterrows():
            p_name = str(row.iloc[0]).strip()
            if not p_name or p_name.lower() == 'nan': continue
            proc_obj = { "Id": str(uuid.uuid4()), "ProcessName": p_name, "ProcessPerformance": [] }
            clean_p = re.sub(r'\s+', '', p_name)
            for k, v_list in kpi_map.items():
                if clean_p in re.sub(r'\s+', '', k) or re.sub(r'\s+', '', k) in clean_p:
                    proc_obj["ProcessPerformance"] = copy.deepcopy(v_list)
                    total_kpi += len(v_list)
                    break
            processes.append(proc_obj)
    final_json["Processes"] = processes

    return final_json, total_kpi

# =====================================================================
# 侧边栏与主界面
# =====================================================================
with st.sidebar:
    st.header("⚙️ 全局配置")
    run_mode = st.radio(
        "请根据报告类型选择：",
        ("纯净标准模式 (无附属场所)", "单提取：EMS 扩展场所 (F21-M25)", "单提取：RL 支持场所 (F27-N32)", "全量综合模式 (提取 EMS + RL + 被支持场所)"),
        index=3
    )
    st.divider()
    user_template_file = st.file_uploader("上传基础 JSON 模板", type=["json"])
    base_template_data = None
    if user_template_file:
        base_template_data = json.load(user_template_file)
    else:
        st.warning("👈 请先上传底座文件以启动程序。")
        st.stop()

st.title("🛡️ IATF 转换引擎 v70.3 (全量地址分离版)")
uploaded_files = st.file_uploader("支持批量上传 .xlsx 格式文件", type=["xlsx"], accept_multiple_files=True)

if uploaded_files:
    for file in uploaded_files:
        try:
            res_json, kpi_count = generate_json_logic(file, base_template_data, run_mode)
            st.success(f"✅ 解析成功：{file.name}")
            
            # 统计信息显示
            col1, col2 = st.columns([3, 1])
            with col1:
                with st.expander("👀 查看数据提取日志", expanded=True):
                    log_text = f"✅ KPI映射成功: {kpi_count}条\n"
                    if "全量综合模式" in run_mode:
                        log_text += f"✅ EMS扩展场所: {len(res_json.get('ExtendedManufacturingSites', []))}个\n"
                        log_text += f"✅ RL支持场所: {len(res_json.get('ProvidingSupportSites', []))}个\n"
                        log_text += f"✅ 被支持场所: {len(res_json.get('ReceivingSupportSites', []))}个\n"
                    st.code(log_text, language="yaml")
                    
            with col2:
                st.download_button(
                    label=f"📥 下载 JSON",
                    data=json.dumps(res_json, indent=2, ensure_ascii=False),
                    file_name=file.name.replace(".xlsx", ".json"),
                    key=f"dl_{file.name}"
                )
        except Exception as e:
            st.error(f"❌ 解析失败: {e}")
