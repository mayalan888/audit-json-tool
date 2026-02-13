import streamlit as st
import pandas as pd
import json
import uuid
import time
import os
import copy
import re
from datetime import datetime, timedelta

st.set_page_config(page_title="IATF 审计转换工具 (v23.0)", page_icon="🎯", layout="wide")

# ================= UI：强制要求上传模板 =================
with st.sidebar:
    st.header("⚙️ 模板配置")
    st.warning("⚠️ 必须上传 JSON 模板才能进行转换。")
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
        st.stop()

# --- 辅助函数：安全更新深层节点 ---
def ensure_path(d, path):
    current = d
    for key in path:
        if key not in current or not isinstance(current[key], dict):
            current[key] = {}
        current = current[key]
    return current

# --- 核心数据转换逻辑 ---
def generate_json_logic(excel_file, template_data):
    final_json = copy.deepcopy(template_data)
    
    try:
        xls = pd.ExcelFile(excel_file)
        db_df = pd.read_excel(xls, sheet_name='数据库', header=None) if '数据库' in xls.sheet_names else pd.read_excel(xls, sheet_name=0, header=None)
        proc_df = pd.read_excel(xls, sheet_name='过程清单') if '过程清单' in xls.sheet_names else pd.DataFrame()
        info_df = pd.read_excel(xls, sheet_name='信息', header=None) if '信息' in xls.sheet_names else pd.DataFrame()
        
        doc_list_df = pd.DataFrame()
        if '文件清单' in xls.sheet_names:
            doc_list_df = pd.read_excel(xls, sheet_name='文件清单', header=None)
        elif len(xls.sheet_names) >= 9:
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

    # --- 1. 数据提取 ---
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
                if "IATF Card" in cell_text or "IATF卡号" in cell_text:
                    if c + 1 < info_df.shape[1]:
                        raw_val = str(info_df.iloc[r, c + 1]).strip()
                        raw_val = raw_val.replace('\n', ' ').replace('\r', ' ')
                        clean_val = re.sub(r'^IATF[:：\s-]*', '', raw_val, flags=re.IGNORECASE).strip()
                        if len(clean_val) > 4:
                            auditor_id = clean_val
                            break
            if auditor_id: break

    # (B) 日期提取
    start_date_raw = find_val_by_key(db_df, ["审核开始日期", "审核开始时间"])
    end_date_raw = find_val_by_key(db_df, ["审核结束日期", "审核结束时间"])
    
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
            next_dt = end_dt + timedelta(days=45)
            next_audit_iso = next_dt.strftime('%Y-%m-%d') + "T00:00:00.000Z"
    except: pass

    # (C) 业务字段提取
    total_employees = find_val_by_key(db_df, ["包括扩展现场在内的员工总数"])
    certificate_scope = find_val_by_key(db_df, ["证书范围"])
    postal_code = find_val_by_key(db_df, ["邮政编码", "邮编"])
    
    customer_name = find_val_by_key(db_df, ["顾客", "客户名称"])
    supplier_code = find_val_by_key(db_df, ["供应商编码", "供应商代码"])
    csr_name = find_val_by_key(db_df, ["CSR文件名称"])
    csr_date = fmt_iso(find_val_by_key(db_df, ["CSR文件日期"]))

    # 中文地址提取
    zh_addr = ""
    addr_candidates = []
    if not db_df.empty:
        for r in range(db_df.shape[0]):
            for c in range(db_df.shape[1]):
                val = str(db_df.iloc[r, c])
                if "地址" in val or "Address" in val:
                    if c+1 < db_df.shape[1]: addr_candidates.append(str(db_df.iloc[r, c+1]))
                    if c+4 < db_df.shape[1]: addr_candidates.append(str(db_df.iloc[r, c+4]))
    for cand in addr_candidates:
        if cand and cand.lower() != 'nan' and re.search(r'[\u4e00-\u9fff]', cand):
            if len(cand) > len(zh_addr): zh_addr = cand

    # 💥 (D) 英文地址全局智能搜寻与断行修复
    en_state, en_city, en_country, en_street1 = "", "", "", ""
    en_addr_raw = ""
    
    # 扫描全表，锁定包含 PROVINCE/CITY 和 CHINA 的格子（这才是真正的英文地址）
    for df in [info_df, db_df]:
        if df.empty: continue
        for r in range(df.shape[0]):
            for c in range(df.shape[1]):
                val = str(df.iloc[r, c]).strip()
                val_upper = val.upper()
                if ("PROVINCE" in val_upper or "CITY" in val_upper) and "CHINA" in val_upper:
                    en_addr_raw = val
                    break
            if en_addr_raw: break
        if en_addr_raw: break

    if en_addr_raw:
        # 1. 把换行符全替换为空格！让被切断的 "LOUDI" 和 "CITY" 重新连在一起
        clean_addr = en_addr_raw.replace('\n', ' ').replace('\r', ' ').replace('，', ',')
        
        # 2. 按照逗号切割
        parts = [p.strip() for p in clean_addr.split(',') if p.strip()]
        
        # 3. 剔除所有中文字符并清洗多余空格
        en_parts = []
        for p in parts:
            if re.search(r'[a-zA-Z]', p):
                p_clean = re.sub(r'[\u4e00-\u9fff]', '', p).strip()
                p_clean = re.sub(r'\s+', ' ', p_clean) # 把多个连在一起的空格变成一个
                if p_clean:
                    en_parts.append(p_clean)

        if en_parts:
            # 最后一个必为 Country
            en_country = en_parts.pop(-1)
            
            street_parts = []
            for p in en_parts:
                p_upper = p.upper()
                if "PROVINCE" in p_upper:
                    en_state = p
                elif "CITY" in p_upper:
                    en_city = p
                else:
                    # 剩下的全归 Street1
                    street_parts.append(p)
            
            en_street1 = ", ".join(street_parts)

    # ================= 2. 定点替换逻辑 =================

    final_json["uuid"] = str(uuid.uuid4())
    final_json["created"] = int(time.time() * 1000)

    # A. 审核基础信息定点替换
    ensure_path(final_json, ["AuditData", "AuditDate"])
    final_json["AuditData"]["AuditDate"]["Start"] = start_iso
    final_json["AuditData"]["AuditDate"]["End"] = end_iso
    final_json["AuditData"]["CbIdentificationNo"] = find_val_by_key(db_df, ["认证机构标识号", "认证机构识别号"])

    if "AuditTeam" not in final_json["AuditData"] or not isinstance(final_json["AuditData"]["AuditTeam"], list):
        final_json["AuditData"]["AuditTeam"] = [{}]
    elif len(final_json["AuditData"]["AuditTeam"]) == 0:
        final_json["AuditData"]["AuditTeam"].append({})
        
    team = final_json["AuditData"]["AuditTeam"][0]
    team["Name"] = auditor_name
    team["CaaNo"] = caa_no
    team["AuditorId"] = auditor_id
    team["AuditDaysPerformed"] = 1.5
    team["DatesOnSite"] = [{"Date": start_iso, "Day": 1}, {"Date": end_iso, "Day": 0.5}]

    # B. 组织信息定点替换
    ensure_path(final_json, ["OrganizationInformation", "AddressNative"])
    ensure_path(final_json, ["OrganizationInformation", "Address"])
    org = final_json["OrganizationInformation"]
    org["OrganizationName"] = find_val_by_key(db_df, ["组织名称"])
    org["IATF_USI"] = find_val_by_key(db_df, ["IATF USI"])
    org["TotalNumberEmployees"] = total_employees
    org["CertificateScope"] = certificate_scope
    
    org["AddressNative"]["Street1"] = zh_addr
    org["AddressNative"]["PostalCode"] = postal_code
    org["AddressNative"]["Country"] = "中国"
    
    # 精准填入切分好的纯英文地址，同步更新 PostalCode
    org["Address"]["State"] = en_state
    org["Address"]["City"] = en_city
    org["Address"]["Country"] = en_country
    org["Address"]["Street1"] = en_street1
    org["Address"]["PostalCode"] = postal_code 

    # C. 顾客与 CSR 定点替换
    ensure_path(final_json, ["CustomerInformation"])
    if "Customers" not in final_json["CustomerInformation"] or not isinstance(final_json["CustomerInformation"]["Customers"], list):
        final_json["CustomerInformation"]["Customers"] = [{}]
    elif len(final_json["CustomerInformation"]["Customers"]) == 0:
        final_json["CustomerInformation"]["Customers"].append({})
        
    cust = final_json["CustomerInformation"]["Customers"][0]
    if "Id" not in cust: cust["Id"] = str(uuid.uuid4())
    cust["Name"] = customer_name
    cust["SupplierCode"] = supplier_code
    
    if "Csrs" not in cust or not isinstance(cust["Csrs"], list):
        cust["Csrs"] = [{}]
    elif len(cust["Csrs"]) == 0:
        cust["Csrs"].append({})
        
    csr = cust["Csrs"][0]
    csr["Name"] = customer_name
    csr["SupplierCode"] = supplier_code
    csr["NameCSRDocument"] = csr_name
    csr["DateCSRDocument"] = csr_date

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
                clause_docs[i]["DocumentName"] = doc_name
            else:
                clause_docs.append({"DocumentName": doc_name})

    # E. 过程清单处理
    processes = []
    if not proc_df.empty:
        clause_cols = proc_df.columns[13:] if proc_df.shape[1] > 13 else []
        for idx, row in proc_df.iterrows():
            p_name = str(row.iloc[12]).strip()
            if not p_name or p_name.lower() == 'nan': continue
            
            proc_obj = {
                "Id": str(uuid.uuid4()),
                "ProcessName": p_name,
                "AuditNotes": [{
                    "Id": str(uuid.uuid4()),
                    "AuditorId": auditor_id,
                    "ManufacturingProcess": "0",
                    "OnSiteProcess": "1",
                    "RemoteProcess": "0"
                }]
            }
            for col in clause_cols:
                if str(row[col]).strip().upper() in ['X', 'TRUE']: 
                    proc_obj[col] = True
            processes.append(proc_obj)
    if processes:
        final_json["Processes"] = processes

    # F. 最终结果日期定点替换
    ensure_path(final_json, ["Results", "AuditReportFinal"])
    final_json["Results"]["AuditReportFinal"]["Date"] = end_iso
    final_json["Results"]["DateNextScheduledAudit"] = next_audit_iso

    return final_json

# ================= 主界面展示 =================
st.title("🎯 IATF 审计数据转换工具 (v23.0)")
st.markdown(f"**当前套用模板**：`{template_name}`")

uploaded_files = st.file_uploader("📥 上传 Excel 数据表", type=["xlsx"], accept_multiple_files=True)

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
                    st.code(f"姓名: {team['Name']}\nCCAA: {team['CaaNo']}\nIATF ID: {team['AuditorId']}", language="yaml")
                with col2:
                    st.download_button("📥 下载生成后的 JSON", data=json.dumps(res_json, indent=2, ensure_ascii=False), file_name=file.name.replace(".xlsx", ".json"), key=f"dl_{file.name}")
        except Exception as e:
            st.error(f"❌ {file.name} 处理失败: {str(e)}")






