import streamlit as st
import pandas as pd
import json
import uuid
import time
import os
import copy
import re
from datetime import datetime, timedelta

st.set_page_config(page_title="IATF 审计转换工具 (v25.0)", page_icon="🎯", layout="wide")

# ================= UI：强制要求上传模板 =================
with st.sidebar:
    st.header("⚙️ 模板配置")
    st.warning("⚠️ 必须上传 JSON 模板。程序不再提供默认模板。")
    user_template = st.file_uploader("上传自定义 JSON 模板", type=["json"])
    
    active_template = None
    template_name = ""

    if user_template:
        try:
            active_template = json.load(user_template)
            template_name = user_template.name
            st.success(f"✅ 成功加载模板: {template_name}")
        except Exception as e:
            st.error(f"❌ 模板解析失败: {e}")
            st.stop()
    else:
        st.stop() # 强制阻断，无模板不运行

# --- 辅助函数：安全更新深层节点，保留未提及的兄弟字段 ---
def ensure_path(d, path):
    current = d
    for key in path:
        if key not in current or not isinstance(current[key], dict):
            current[key] = {}
        current = current[key]
    return current

# --- 核心转换逻辑 ---
def generate_json_logic(excel_file, template_data):
    # 深度拷贝底稿：其余全部保留
    final_json = copy.deepcopy(template_data)
    
    try:
        xls = pd.ExcelFile(excel_file)
        # 指定读取的四张表
        db_df = pd.read_excel(xls, sheet_name='数据库', header=None) if '数据库' in xls.sheet_names else pd.read_excel(xls, sheet_name=0, header=None)
        proc_df = pd.read_excel(xls, sheet_name='过程清单') if '过程清单' in xls.sheet_names else pd.DataFrame()
        info_df = pd.read_excel(xls, sheet_name='信息', header=None) if '信息' in xls.sheet_names else pd.DataFrame()
        
        # 提取第九张子表 (文件清单)
        doc_list_df = pd.DataFrame()
        if len(xls.sheet_names) >= 9:
            doc_list_df = pd.read_excel(xls, sheet_name=xls.sheet_names[8], header=None)
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

    # --- 1. 基础数据提取 ---
    raw_name_full = find_val_by_key(db_df, ["姓名", "Auditor Name"])
    raw_name = raw_name_full.replace("姓名:", "").replace("Name:", "").strip() if raw_name_full else ""
    auditor_name = raw_name
    english_part = re.sub(r'[\u4e00-\u9fff]', '', raw_name).strip()
    if english_part:
        parts = english_part.split()
        if len(parts) >= 2 and parts[0].isupper() and not parts[1].isupper():
            auditor_name = f"{parts[1]} {parts[0]}"
        else:
            auditor_name = english_part

    ccaa_raw = find_val_by_key(db_df, ["审核员CCAA", "CCAA"])
    caa_no = ""
    if ccaa_raw:
        match = re.search(r'(?:CCAA[:：\s-])\s*(.*)', ccaa_raw, re.IGNORECASE | re.DOTALL)
        caa_no = match.group(1).strip() if match else ccaa_raw.strip()

    auditor_id = ""
    if not info_df.empty:
        for r in range(info_df.shape[0]):
            for c in range(info_df.shape[1]):
                cell_text = str(info_df.iloc[r, c])
                if "IATF Card" in cell_text:
                    if c + 1 < info_df.shape[1]:
                        raw_val = str(info_df.iloc[r, c + 1]).strip()
                        raw_val = raw_val.replace('\n', ' ').replace('\r', ' ')
                        clean_val = re.sub(r'^IATF[:：\s-]*', '', raw_val, flags=re.IGNORECASE).strip()
                        if len(clean_val) > 4:
                            auditor_id = clean_val
                            break
            if auditor_id: break

    # 日期处理与 +45天
    start_date_raw = find_val_by_key(db_df, ["审核开始日期"])
    end_date_raw = find_val_by_key(db_df, ["审核结束日期"])
    def fmt_iso(val):
        try:
            dt = pd.to_datetime(val, errors='coerce')
            if pd.notna(dt): return dt.strftime('%Y-%m-%d') + "T00:00:00.000Z"
        except: pass
        return ""
    start_iso, end_iso = fmt_iso(start_date_raw), fmt_iso(end_date_raw)
    
    next_audit_iso = ""
    try:
        end_dt = pd.to_datetime(end_date_raw, errors='coerce')
        if pd.notna(end_dt):
            next_audit_iso = (end_dt + timedelta(days=45)).strftime('%Y-%m-%d') + "T00:00:00.000Z"
    except: pass

    # 业务字段提取
    total_employees = find_val_by_key(db_df, ["包括扩展现场在内的员工总数"])
    certificate_scope = find_val_by_key(db_df, ["证书范围"])
    postal_code = find_val_by_key(db_df, ["邮政编码", "邮编"])
    zh_street1 = find_val_by_key(db_df, ["街道1"])
    
    customer_name = find_val_by_key(db_df, ["顾客"])
    supplier_code = find_val_by_key(db_df, ["供应商编码"])
    csr_name = find_val_by_key(db_df, ["CSR文件名称"])
    csr_date = fmt_iso(find_val_by_key(db_df, ["CSR文件日期"]))

    # 💥 英文地址拼接与切分逻辑 (信息表)
    en_state, en_city, en_country, en_street1 = "", "", "", ""
    en_addr_raw = ""
    if not info_df.empty:
        for r in range(info_df.shape[0]):
            for c in range(info_df.shape[1]):
                val = str(info_df.iloc[r, c]).strip()
                if "审核地址" in val or "Address" in val:
                    if c + 1 < info_df.shape[1]:
                        right_val = str(info_df.iloc[r, c+1]).strip()
                        if re.search(r'[a-zA-Z]', right_val):
                            en_addr_raw = right_val
                            break
            if en_addr_raw: break

    if en_addr_raw:
        # 缝合断行
        clean_addr = en_addr_raw.replace('\n', ' ').replace('\r', ' ').replace('，', ',')
        parts = [p.strip() for p in clean_addr.split(',') if p.strip()]
        en_parts = [re.sub(r'[\u4e00-\u9fff]', '', p).strip() for p in parts if re.search(r'[a-zA-Z]', p)]
        
        if en_parts:
            en_country = en_parts.pop(-1)
            street_accum = []
            for p in en_parts:
                if "PROVINCE" in p.upper(): en_state = p
                elif "CITY" in p.upper(): en_city = p
                else: street_accum.append(p)
            en_street1 = ", ".join(street_accum)

    # ================= 2. 定点替换 =================

    # 根节点
    final_json["uuid"] = str(uuid.uuid4())
    final_json["created"] = int(time.time() * 1000)

    # AuditData
    ensure_path(final_json, ["AuditData", "AuditDate"])
    final_json["AuditData"]["AuditDate"]["Start"] = start_iso
    final_json["AuditData"]["AuditDate"]["End"] = end_iso
    final_json["AuditData"]["CbIdentificationNo"] = find_val_by_key(db_df, ["认证机构标识号"])

    if "AuditTeam" in final_json["AuditData"] and len(final_json["AuditData"]["AuditTeam"]) > 0:
        team = final_json["AuditData"]["AuditTeam"][0]
        team.update({"Name": auditor_name, "CaaNo": caa_no, "AuditorId": auditor_id, "AuditDaysPerformed": 1.5})
        team["DatesOnSite"] = [{"Date": start_iso, "Day": 1}, {"Date": end_iso, "Day": 0.5}]

    # OrganizationInformation
    ensure_path(final_json, ["OrganizationInformation", "AddressNative"])
    ensure_path(final_json, ["OrganizationInformation", "Address"])
    org = final_json["OrganizationInformation"]
    org.update({"TotalNumberEmployees": total_employees, "CertificateScope": certificate_scope})
    org["AddressNative"].update({"Street1": zh_street1, "PostalCode": postal_code, "Country": "中国"})
    org["Address"].update({"State": en_state, "City": en_city, "Country": en_country, "Street1": en_street1, "PostalCode": postal_code})

    # CustomerInformation (嵌套 CSR)
    if "CustomerInformation" in final_json and "Customers" in final_json["CustomerInformation"] and len(final_json["CustomerInformation"]["Customers"]) > 0:
        cust = final_json["CustomerInformation"]["Customers"][0]
        cust.update({"Name": customer_name, "SupplierCode": supplier_code})
        if "Csrs" in cust and len(cust["Csrs"]) > 0:
            cust["Csrs"][0].update({"Name": customer_name, "SupplierCode": supplier_code, "NameCSRDocument": csr_name, "DateCSRDocument": csr_date})

    # 第九张表 (Stage1DocumentedRequirements)
    docs_list = []
    if not doc_list_df.empty:
        for c in range(doc_list_df.shape[1]):
            for r in range(doc_list_df.shape[0]):
                if "公司内对应的程序文件" in str(doc_list_df.iloc[r, c]):
                    for r2 in range(r + 1, doc_list_df.shape[0]):
                        val = str(doc_list_df.iloc[r2, c]).strip()
                        if val and val.lower() != 'nan': docs_list.append(val)
                    break

    if docs_list:
        ensure_path(final_json, ["Stage1DocumentedRequirements"])
        if "IatfClauseDocuments" not in final_json["Stage1DocumentedRequirements"]: final_json["Stage1DocumentedRequirements"]["IatfClauseDocuments"] = []
        clause_docs = final_json["Stage1DocumentedRequirements"]["IatfClauseDocuments"]
        for i, doc_name in enumerate(docs_list):
            if i < len(clause_docs): clause_docs[i]["DocumentName"] = doc_name
            else: clause_docs.append({"DocumentName": doc_name})

    # 过程清单
    processes = []
    if not proc_df.empty:
        clause_cols = proc_df.columns[13:] if proc_df.shape[1] > 13 else []
        for idx, row in proc_df.iterrows():
            p_name = str(row.iloc[12]).strip()
            if not p_name or p_name.lower() == 'nan': continue
            proc_obj = {
                "Id": str(uuid.uuid4()), "ProcessName": p_name,
                "AuditNotes": [{
                    "Id": str(uuid.uuid4()), "AuditorId": auditor_id,
                    "ManufacturingProcess": "0", "OnSiteProcess": "1", "RemoteProcess": "0"
                }]
            }
            for col in clause_cols:
                if str(row[col]).strip().upper() in ['X', 'TRUE']: proc_obj[col] = True
            processes.append(proc_obj)
    if processes: final_json["Processes"] = processes

    # Results
    ensure_path(final_json, ["Results", "AuditReportFinal"])
    final_json["Results"]["AuditReportFinal"]["Date"] = end_iso
    final_json["Results"]["DateNextScheduledAudit"] = next_audit_iso

    return final_json

# ================= UI =================
st.title("🚀 终极定点审计转换引擎 (v25.0)")
st.write(f"当前生效模板：**{template_name}**")
uploaded_files = st.file_uploader("上传 Excel 数据表", type=["xlsx"], accept_multiple_files=True)

if uploaded_files:
    st.divider()
    for file in uploaded_files:
        try:
            res_json = generate_json_logic(file, active_template)
            team = res_json["AuditData"]["AuditTeam"][0]
            with st.expander(f"📄 {file.name} - 处理结果", expanded=True):
                col1, col2 = st.columns([3, 1])
                with col1:
                    st.success("✅ 转换成功")
                    st.code(f"姓名: {team['Name']} | CCAA: {team['CaaNo'][:15]}... | IATF ID: {team['AuditorId']}", language="yaml")
                with col2:
                    st.download_button("📥 下载 JSON", data=json.dumps(res_json, indent=2, ensure_ascii=False), file_name=file.name.replace(".xlsx", ".json"), key=f"dl_{file.name}")
        except Exception as e:
            st.error(f"❌ {file.name} 处理失败: {str(e)}")





