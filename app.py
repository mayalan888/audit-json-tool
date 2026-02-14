import streamlit as st
import pandas as pd
import json
import uuid
import time
import os
import copy
import re
from datetime import datetime, timedelta

# --- 页面配置 ---
st.set_page_config(
    page_title="IATF 审计转换工具 (v44.0)",
    page_icon="🛡️",
    layout="wide"
)

# --- 1. 侧边栏：模板加载与融合逻辑 ---
with st.sidebar:
    st.header("⚙️ 模板配置")
    
    base_template = None
    if os.path.exists('金磁.json'):
        try:
            with open('金磁.json', 'r', encoding='utf-8') as f:
                base_template = json.load(f)
            st.success("✅ 已加载标准底座: `金磁.json`")
        except Exception as e:
            st.error(f"❌ 读取 `金磁.json` 失败: {e}")
            st.stop()
    else:
        st.error("❌ 找不到标准底座 `金磁.json`！请确保它在项目根目录下。")
        st.stop()

    st.info("💡 请上传您的模板。程序将提取其中的 Stage1 节点来替换底座。")
    user_template_file = st.file_uploader("上传自定义 JSON 模板", type=["json"])
    
    user_template_data = None
    if user_template_file:
        try:
            user_template_data = json.load(user_template_file)
            st.success(f"✅ 成功加载自定义模板: {user_template_file.name}")
        except Exception as e:
            st.error(f"❌ 自定义模板解析失败: {e}")
            st.stop()
    else:
        st.warning("👈 请先在左侧上传 JSON 模板文件。")
        st.stop()

# --- 辅助函数：安全寻址 ---
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

# --- 核心转换逻辑 ---
def generate_json_logic(excel_file, base_data, user_data):
    final_json = copy.deepcopy(base_data)
    
    for key in ["Stage1Activities", "Stage1Part1", "Stage1Part2"]:
        if key in user_data:
            final_json[key] = copy.deepcopy(user_data[key])
    
    try:
        xls = pd.ExcelFile(excel_file)
        db_df = pd.read_excel(xls, sheet_name='数据库', header=None) if '数据库' in xls.sheet_names else pd.read_excel(xls, sheet_name=0, header=None)
        proc_df = pd.read_excel(xls, sheet_name='过程清单') if '过程清单' in xls.sheet_names else pd.DataFrame()
        info_df = pd.read_excel(xls, sheet_name='信息', header=None) if '信息' in xls.sheet_names else pd.DataFrame()
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

    # ================= 2. 数据提取 =================
    
    # [姓名]
    raw_name_full = find_val_by_key(db_df, ["姓名", "Auditor Name"]) or get_db_val(5, 1)
    raw_name = raw_name_full.replace("姓名:", "").replace("Name:", "").strip() if raw_name_full else ""
    auditor_name = raw_name
    english_part = re.sub(r'[\u4e00-\u9fff]', '', raw_name).strip()
    if english_part:
        parts = english_part.split()
        if len(parts) >= 2 and parts[0].isupper() and not parts[1].isupper(): auditor_name = f"{parts[1]} {parts[0]}"
        else: auditor_name = english_part

    # [CCAA]
    ccaa_raw = find_val_by_key(db_df, ["审核员CCAA", "CCAA"]) or get_db_val(4, 1)
    caa_no = ""
    if ccaa_raw:
        match = re.search(r'(?:CCAA[:：\s-])\s*(.*)', ccaa_raw, re.IGNORECASE | re.DOTALL)
        caa_no = match.group(1).strip() if match else ccaa_raw.strip()

    # [AuditorId]
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

    # [日期]
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

    # [顾客与 CSR]
    customer_name = find_val_by_key(db_df, ["顾客", "客户名称"]) or get_db_val(29, 1)
    supplier_code = find_val_by_key(db_df, ["供应商编码", "供应商代码"]) or get_db_val(30, 1)
    csr_name = find_val_by_key(db_df, ["CSR文件名称"]) or get_db_val(31, 1)
    csr_date_raw = find_val_by_key(db_df, ["CSR文件日期"]) or get_db_val(32, 1)
    csr_date = fmt_iso(csr_date_raw)

    # 💥 [中英文地址分离与智能提取（破案修复版）]
    def is_chinese(s): return bool(re.search(r'[\u4e00-\u9fff]', s))
    
    native_street = ""
    english_address = ""

    # 1. 优先扫描信息表（不会被其他变量干扰提前中断）
    if not info_df.empty:
        for r in range(info_df.shape[0]):
            for c in range(info_df.shape[1]):
                val = str(info_df.iloc[r, c]).strip()
                if "审核地址" in val or "Audit Address" in val:
                    if c + 1 < info_df.shape[1]:
                        rv = str(info_df.iloc[r, c+1]).strip()
                        if rv and rv.lower() != 'nan':
                            lines = rv.replace('\r', '\n').split('\n')
                            en_lines = [l.strip() for l in lines if not is_chinese(l) and l.strip()]
                            zh_lines = [l.strip() for l in lines if is_chinese(l) and l.strip()]
                            if en_lines: english_address = " ".join(en_lines)
                            if zh_lines: native_street = " ".join(zh_lines)
                        break

    # 2. 如果信息表里没找到，再去数据库兜底
    if not native_street or not english_address:
        db_candidates = [get_db_val(11, 1), get_db_val(11, 4)]
        zh_cands = [c for c in db_candidates if c and is_chinese(c)]
        en_cands = [c for c in db_candidates if c and not is_chinese(c)]
        if not native_street and zh_cands: native_street = max(zh_cands, key=len)
        if not english_address and en_cands: english_address = max(en_cands, key=len)

    # 倒序切分
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

    # ================= 3. 定点替换入 final_json =================

    final_json["uuid"] = str(uuid.uuid4())
    final_json["created"] = int(time.time() * 1000)

    # A. 审核数据
    ensure_path(final_json, ["AuditData", "AuditDate"])
    if start_iso: final_json["AuditData"]["AuditDate"]["Start"] = start_iso
    if end_iso: final_json["AuditData"]["AuditDate"]["End"] = end_iso
    final_json["AuditData"]["CbIdentificationNo"] = find_val_by_key(db_df, ["认证机构标识号"]) or get_db_val(2, 4)

    if "AuditTeam" not in final_json["AuditData"] or not isinstance(final_json["AuditData"]["AuditTeam"], list) or len(final_json["AuditData"]["AuditTeam"]) == 0:
        final_json["AuditData"]["AuditTeam"] = [{}]
        
    team = final_json["AuditData"]["AuditTeam"][0]
    if isinstance(team, dict):
        team.update({
            "Name": auditor_name,
            "CaaNo": caa_no,
            "AuditorId": auditor_id, 
            "AuditDaysPerformed": 1.5,
            "DatesOnSite": [{"Date": start_iso, "Day": 1}, {"Date": end_iso, "Day": 0.5}]
        })

    # B. 组织与地址信息 
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
            if isinstance(lang_node[0], dict):
                lang_node[0]["Products"] = ""
        elif isinstance(lang_node, dict):
            if "0" in lang_node and isinstance(lang_node["0"], dict):
                lang_node["0"]["Products"] = ""
            else:
                lang_node["Products"] = ""
    
    org["AddressNative"].update({
        "Street1": native_street,
        "State": "",
        "City": "",
        "Country": "中国",
        "PostalCode": find_val_by_key(db_df, ["邮政编码"]) or get_db_val(10, 4)
    })
    
    org["Address"].update({
        "State": state, 
        "City": city, 
        "Country": country, 
        "Street1": street,
        "PostalCode": find_val_by_key(db_df, ["邮政编码"]) or get_db_val(10, 4)
    })

    # C. 顾客与 CSR 
    ensure_path(final_json, ["CustomerInformation"])
    if "Customers" not in final_json["CustomerInformation"] or not isinstance(final_json["CustomerInformation"]["Customers"], list) or not final_json["CustomerInformation"]["Customers"]:
        final_json["CustomerInformation"]["Customers"] = [{}]
        
    cust = final_json["CustomerInformation"]["Customers"][0]
    if "Id" not in cust: cust["Id"] = str(uuid.uuid4())
    if customer_name: cust["Name"] = customer_name
    if supplier_code: cust["SupplierCode"] = supplier_code
    
    if "Csrs" not in cust or not isinstance(cust["Csrs"], list) or not cust["Csrs"]:
        cust["Csrs"] = [{}]
        
    csr = cust["Csrs"][0]
    if customer_name: csr["Name"] = customer_name
    if supplier_code: csr["SupplierCode"] = supplier_code
    if csr_name: csr["NameCSRDocument"] = csr_name
    if csr_date: csr["DateCSRDocument"] = csr_date

    # D. 文件清单定点替换
    docs_list = []
    if not doc_list_df.empty:
        for c in range(doc_list_df.shape[1]):
            for r in range(doc_list_df.shape[0]):
                cell_val = str(doc_list_df.iloc[r, c]).strip()
                if "公司内对应的程序文件" in cell_val or "包含名称、编号、版本" in cell_val:
                    for r2 in range(r + 1, doc_list_df.shape[0]):
                        val = str(doc_list_df.iloc[r2, c]).strip()
                        if val and val.lower() != 'nan':
                            docs_list.append(val)
                    break
            if docs_list: break

    if docs_list:
        ensure_path(final_json, ["Stage1DocumentedRequirements"])
        if "IatfClauseDocuments" not in final_json["Stage1DocumentedRequirements"] or not isinstance(final_json["Stage1DocumentedRequirements"]["IatfClauseDocuments"], list):
            final_json["Stage1DocumentedRequirements"]["IatfClauseDocuments"] = []
            
        clause_docs = final_json["Stage1DocumentedRequirements"]["IatfClauseDocuments"]
        for i, doc_name in enumerate(docs_list):
            if i < len(clause_docs):
                if isinstance(clause_docs[i], dict):
                    clause_docs[i]["DocumentName"] = doc_name
            else:
                clause_docs.append({"DocumentName": doc_name})

    # E. 过程清单重建
    processes = []
    if not proc_df.empty:
        clause_cols = proc_df.columns[13:] if proc_df.shape[1] > 13 else []
        for idx, row in proc_df.iterrows():
            p_name = str(row.iloc[0]).strip()
            rep_name = str(row.iloc[2]).strip() if pd.notna(row.iloc[2]) else ""
            
            if not p_name or p_name.lower() == 'nan': continue
            proc_obj = {
                "Id": str(uuid.uuid4()),
                "ProcessName": p_name,
                "RepresentativeName": rep_name,
                "ManufacturingProcess": "0",
                "OnSiteProcess": "1",
                "RemoteProcess": "0",
                "AuditNotes": [{
                    "Id": str(uuid.uuid4()),
                    "AuditorId": auditor_id
                }]
            }
            for col in clause_cols:
                if str(row[col]).strip().upper() in ['X', 'TRUE']: proc_obj[col] = True
            processes.append(proc_obj)
    final_json["Processes"] = processes

    # F. 结果日期
    if "Results" not in final_json: final_json["Results"] = {}
    if "AuditReportFinal" not in final_json["Results"]: final_json["Results"]["AuditReportFinal"] = {}
    
    if end_iso: final_json["Results"]["AuditReportFinal"]["Date"] = end_iso
    if next_audit_iso: final_json["Results"]["DateNextScheduledAudit"] = next_audit_iso

    return final_json

# ================= 主界面 =================
st.title("🛡️ 多模板审计转换引擎 (v44.0 地址防短路版)")
st.markdown("💡 **修复日志**：去除了导致英文地址读取提前短路中断的 BUG 代码，现在英文 Address 节点将完美生成。")

uploaded_files = st.file_uploader("📥 上传 Excel 数据表", type=["xlsx"], accept_multiple_files=True)

if uploaded_files:
    st.divider()
    for file in uploaded_files:
        try:
            res_json = generate_json_logic(file, base_template, user_template_data)
            st.success(f"✅ {file.name} 转换成功")
            
            with st.expander("👀 查看诊断面板", expanded=True):
                 st.code(f"""
【英文 Address 完美切分确认】
Street1: "{safe_get(res_json['OrganizationInformation']['Address'], 'Street1')}"
City:    "{safe_get(res_json['OrganizationInformation']['Address'], 'City')}"
State:   "{safe_get(res_json['OrganizationInformation']['Address'], 'State')}"
Country: "{safe_get(res_json['OrganizationInformation']['Address'], 'Country')}"
                 """.strip(), language="yaml")

            st.download_button(
                label=f"📥 下载 JSON ({file.name})",
                data=json.dumps(res_json, indent=2, ensure_ascii=False),
                file_name=file.name.replace(".xlsx", ".json"),
                key=f"dl_{file.name}"
            )
        except Exception as e:
            st.error(f"❌ {file.name} 核心处理失败: {str(e)}")








