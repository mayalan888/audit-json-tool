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
    page_title="IATF 审计数据转换工具 v9",
    page_icon="📊",
    layout="wide"
)

# --- 核心转换逻辑函数 ---
def generate_json_logic(excel_file, template_data):
    final_json = copy.deepcopy(template_data)
    
    try:
        xls = pd.ExcelFile(excel_file)
        
        # 加载必要 Sheet
        db_df = pd.read_excel(xls, sheet_name='数据库', header=None) if '数据库' in xls.sheet_names else pd.read_excel(xls, sheet_name=0, header=None)
        proc_df = pd.read_excel(xls, sheet_name='过程清单') if '过程清单' in xls.sheet_names else pd.DataFrame()
        info_df = pd.read_excel(xls, sheet_name='信息', header=None) if '信息' in xls.sheet_names else pd.DataFrame()
        
        doc_list_df = pd.DataFrame()
        if '文件清单' in xls.sheet_names:
            doc_list_df = pd.read_excel(xls, sheet_name='文件清单')
        elif len(xls.sheet_names) >= 9:
            doc_list_df = pd.read_excel(xls, sheet_name=xls.sheet_names[8])
            
    except Exception as e:
        raise ValueError(f"Excel 读取失败: {str(e)}")

    def get_db_val(row, col):
        try:
            val = db_df.iloc[row, col]
            return str(val).strip() if pd.notna(val) else ""
        except:
            return ""

    # --- 1. 基础信息提取 ---
    report_name = get_db_val(1, 1)
    org_name = get_db_val(1, 4)
    start_date_raw = get_db_val(2, 1)
    cb_id = get_db_val(2, 4)
    end_date_raw = get_db_val(3, 1)
    usi_code = get_db_val(3, 4)

    # --- 2. 姓名与编号映射 (根据您的最新要求) ---
    
    # 姓名 (数据库中姓名一栏: 第5行第2列)
    raw_auditor_name = get_db_val(4, 1)
    # 自动清洗姓名逻辑 (保持英文字符顺序)
    auditor_name = raw_auditor_name
    english_part = re.sub(r'[\u4e00-\u9fff]', '', raw_auditor_name).strip()
    if english_part:
        parts = english_part.split()
        if len(parts) >= 2 and parts[0].isupper() and not parts[1].isupper():
            auditor_name = f"{parts[1]} {parts[0]}"
        else:
            auditor_name = english_part

    # AuditorId (信息表中 "IATF Card：" 后 "IATF:" 后的编号)
    auditor_id = ""
    if not info_df.empty:
        # 在整个“信息”表中搜索关键字
        for r in range(len(info_df)):
            for c in range(len(info_df.columns)):
                cell_str = str(info_df.iloc[r, c])
                if "IATF Card：" in cell_str:
                    # 尝试从当前单元格或右侧单元格提取
                    target_text = cell_str + " " + str(info_df.iloc[r, c+1] if c+1 < len(info_df.columns) else "")
                    match = re.search(r'IATF:\s*([\w-]+)', target_text)
                    if match:
                        auditor_id = match.group(1)
                        break
            if auditor_id: break

    # CaaNo (数据库中 "审核员CCAA-编号" 后 "CCAA:" 后的编号)
    caa_no = ""
    # 通常在第4行第1列 (索引 3, 0)
    ccaa_raw = get_db_val(3, 0)
    ccaa_match = re.search(r'CCAA:\s*([\w-]+)', ccaa_raw)
    if ccaa_match:
        caa_no = ccaa_match.group(1)

    # --- 3. 员工、客户与 CSR ---
    total_employees = get_db_val(27, 1)
    customer_name = get_db_val(29, 1)
    supplier_code = get_db_val(30, 1)
    if supplier_code == "无": supplier_code = ""
    csr_doc_name = get_db_val(31, 1)
    csr_doc_date_raw = get_db_val(32, 1)

    # 格式化日期
    def parse_to_dt(d): return pd.to_datetime(d, errors='coerce')
    def fmt_iso(dt): return dt.strftime('%Y-%m-%d') + "T00:00:00.000Z" if pd.notna(dt) else ""

    start_dt, end_dt = parse_to_dt(start_date_raw), parse_to_dt(end_date_raw)
    csr_date = f"{parse_to_dt(csr_doc_date_raw).year}/{parse_to_dt(csr_doc_date_raw).month}" if pd.notna(parse_to_dt(csr_doc_date_raw)) else str(csr_doc_date_raw)

    # --- 4. 地址与范围 ---
    certificate_scope = ""
    for idx, row in db_df.iterrows():
        if str(row[0]).strip() == "证书范围":
            certificate_scope = str(row[1]).strip() if pd.notna(row[1]) else ""
            break

    # 地址识别逻辑
    c1, c4 = get_db_val(11, 1), get_db_val(11, 4)
    zh_addr = max([c for c in [c1, c4] if re.search(r'[\u4e00-\u9fff]', c)], key=len, default="")
    en_addr = max([c for c in [c1, c4] if not re.search(r'[\u4e00-\u9fff]', c)], key=len, default="")
    
    # --- 5. 更新数据结构 ---
    
    # AuditData
    if "AuditData" not in final_json: final_json["AuditData"] = {}
    final_json["AuditData"].update({
        "AuditDate": {"Start": fmt_iso(start_dt), "End": fmt_iso(end_dt)},
        "ReportName": report_name,
        "CbIdentificationNo": cb_id,
        "AuditTeam": [{
            "Name": auditor_name,
            "CaaNo": caa_no,
            "AuditorId": auditor_id,
            "AuditDaysPerformed": 1.5,
            "DatesOnSite": [{"Date": fmt_iso(start_dt), "Day": 1}, {"Date": fmt_iso(end_dt), "Day": 0.5}],
            "PlanningTime": "0.0000"
        }]
    })

    # Org Info
    if "OrganizationInformation" not in final_json: final_json["OrganizationInformation"] = {}
    final_json["OrganizationInformation"].update({
        "OrganizationName": org_name,
        "AddressNative": {"Street1": zh_addr, "PostalCode": get_db_val(10, 4), "Country": "中国"},
        "IATF_USI": usi_code,
        "CertificateScope": certificate_scope,
        "TotalNumberEmployees": total_employees
    })

    # Customer & CSR
    if "CustomerInformation" not in final_json: final_json["CustomerInformation"] = {}
    final_json["CustomerInformation"]["Customers"] = [{
        "Id": str(int(time.time() * 1000)),
        "Name": customer_name,
        "SupplierCode": supplier_code,
        "Csrs": [{"Id": str(int(time.time() * 1000) + 1), "NameCSRDocument": csr_doc_name, "DateCSRDocument": csr_date}]
    }]

    # 文件清单 (Stage1)
    if not doc_list_df.empty and "Stage1DocumentedRequirements" in final_json:
        target_col = next((col for col in doc_list_df.columns if "公司内对应的程序文件" in str(col)), None)
        if target_col:
            iatf_docs = final_json["Stage1DocumentedRequirements"].get("IatfClauseDocuments", [])
            for i, doc_name in enumerate(doc_list_df[target_col]):
                if i < len(iatf_docs): iatf_docs[i]["DocumentName"] = str(doc_name).strip() if pd.notna(doc_name) else ""

    # 过程清单 (Processes)
    processes = []
    clause_cols = proc_df.columns[13:] if proc_df.shape[1] > 13 else []
    for idx, row in proc_df.iterrows():
        p_name = str(row.iloc[12]).strip()
        if not p_name or p_name.lower() == 'nan': continue
        proc_obj = {
            "Id": str(int(time.time() * 1000) + idx),
            "ProcessName": p_name,
            "RepresentativeName": str(row.iloc[2]),
            "AuditNotes": [{"Id": int(time.time() * 1000) + idx + 1000, "AuditorId": auditor_id}],
            "ManufacturingProcess": "0", "OnSiteProcess": "1", "RemoteProcess": "0"
        }
        for col in clause_cols:
            if str(row[col]).strip().upper() in ['X', 'TRUE']: proc_obj[col] = True
        processes.append(proc_obj)

    final_json["Processes"] = processes
    final_json["uuid"], final_json["created"] = str(uuid.uuid4()), int(time.time() * 1000)

    # 结果日期 (Results)
    if "Results" not in final_json: final_json["Results"] = {}
    final_json["Results"]["AuditReportFinal"] = {"Date": fmt_iso(end_dt)}
    final_json["Results"]["DateNextScheduledAudit"] = fmt_iso(end_dt + timedelta(days=45)) if pd.notna(end_dt) else ""

    return final_json

# --- Streamlit 界面 ---
st.title("📊 IATF 审计数据转换工具")
uploaded_files = st.file_uploader("上传 Excel 数据表", type=["xlsx"], accept_multiple_files=True)

if uploaded_files:
    # 加载默认模板
    with open('金磁.json', 'r', encoding='utf-8') as f:
        template = json.load(f)
        
    for file in uploaded_files:
        try:
            res_json = generate_json_logic(file, template)
            st.success(f"✅ {file.name} 转换成功")
            st.download_button(f"📥 下载 {file.name.replace('.xlsx', '.json')}", 
                               json.dumps(res_json, indent=2, ensure_ascii=False), 
                               file_name=file.name.replace(".xlsx", ".json"))
        except Exception as e:
            st.error(f"❌ {file.name} 处理失败: {e}")


