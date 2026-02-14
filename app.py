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
    page_title="IATF 审计转换工具 (v36.0 纯净提取版)",
    page_icon="🛡️",
    layout="wide"
)

# --- 1. 侧边栏：强制要求用户导入模板 ---
with st.sidebar:
    st.header("⚙️ 模板配置")
    st.warning("⚠️ 必须上传您的 JSON 模板。程序将仅提取指定的三个阶段数据。")
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
        st.info("请先在左侧上传 JSON 模板文件。")
        st.stop()

# --- 辅助函数：安全寻址 ---
def ensure_path(d, path):
    current = d
    for key in path:
        if key not in current or not isinstance(current[key], dict):
            current[key] = {}
        current = current[key]
    return current

# --- 核心转换逻辑 ---
def generate_json_logic(excel_file, template_data):
    # 💥 【核心修改】：不再全盘深拷贝，而是从零构建！
    final_json = {}
    
    # 仅取用上传的 json 文件中指定的这三个特定节点
    for key in ["Stage1Activities", "Stage1Part1", "Stage1Part2"]:
        if key in template_data:
            final_json[key] = copy.deepcopy(template_data[key])
    
    try:
        xls = pd.ExcelFile(excel_file)
        db_df = pd.read_excel(xls, sheet_name='数据库', header=None) if '数据库' in xls.sheet_names else pd.read_excel(xls, sheet_name=0, header=None)
        proc_df = pd.read_excel(xls, sheet_name='过程清单') if '过程清单' in xls.sheet_names else pd.DataFrame()
        info_df = pd.read_excel(xls, sheet_name='信息', header=None) if '信息' in xls.sheet_names else pd.DataFrame()
        doc_list_df = pd.read_excel(xls, sheet_name=xls.sheet_names[8], header=None) if len(xls.sheet_names) >= 9 else pd.DataFrame()
    except Exception as e:
        raise ValueError(f"Excel 读取失败: {str(e)}")

    # 智能游动扫描
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
        
    # 绝对坐标兜底 
    def get_db_val(r, c):
        try:
            val = db_df.iloc[r, c]
            return str(val).strip() if pd.notna(val) else ""
        except: return ""

    # ================= 1. 数据提取 (双重保险) =================
    
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

    # [日期深度修复]
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

    # [英文地址倒序切分]
    english_address = ""

    if not info_df.empty:
        for r in range(info_df.shape[0]):
            for c in range(info_df.shape[1]):
                val = str(info_df.iloc[r, c]).strip()
                if "审核地址" in val or "Audit Address" in val:
                    if c + 1 < info_df.shape[1]:
                        rv = str(info_df.iloc[r, c+1]).strip()
                        if rv and rv.lower() != 'nan':
                            lines = rv.replace('\r', '\n').split('\n')
                            en_lines = [l.strip() for l in lines if not re.search(r'[\u4e00-\u9fff]', l) and l.strip()]
                            if en_lines: english_address = " ".join(en_lines)
                        break
            if english_address: break

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

    # ================= 2. 核心结构强制组装 =================

    final_json["uuid"] = str(uuid.uuid4())
    final_json["created"] = int(time.time() * 1000)

    # A. 审核数据
    ensure_path(final_json, ["AuditData", "AuditData"])
    if start_iso: final_json["AuditData"]["AuditData"]["start"] = start_iso
    if end_iso: final_json["AuditData"]["AuditData"]["end"] = end_iso
    final_json["AuditData"]["CbIdentificationNo"] = find_val_by_key(db_df, ["认证机构标识号"]) or get_db_val(2, 4)

    # 因为是空白重建，必须确保 AuditTeam 存在
    if "AuditTeam" not in final_json["AuditData"] or not final_json["AuditData"]["AuditTeam"]:
        final_json["AuditData"]["AuditTeam"] = [{}]
        
    team = final_json["AuditData"]["AuditTeam"][0]
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
    
    # AddressNative 留空
    org["AddressNative"].update({
        "Street1": "",
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
                clause_docs[i]["DocumentName"] = doc_name
            else:
                clause_docs.append({"DocumentName": doc_name})

    # E. 过程清单重建
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
st.title("🛡️ 多模板审计转换引擎 (v36.0 纯净提取版)")
st.write(f"当前生效模板：**{template_name}**")

uploaded_files = st.file_uploader("📥 上传 Excel 数据表", type=["xlsx"], accept_multiple_files=True)

if uploaded_files:
    st.divider()
    for file in uploaded_files:
        try:
            res_json = generate_json_logic(file, active_template)
            
            # 安全检查节点是否存在，避免网页预览报错
            has_stage1 = "Stage1Activities" in res_json
            
            st.success(f"✅ {file.name} 转换成功")
            
            with st.expander("👀 查看诊断面板", expanded=True):
                 st.code(f"""
【模板提取确认】
成功从模板中剥离并嵌入 Stage1Activities: {has_stage1}
成功从模板中剥离并嵌入 Stage1Part1:      {"Stage1Part1" in res_json}
成功从模板中剥离并嵌入 Stage1Part2:      {"Stage1Part2" in res_json}

【AddressNative 确认】
Street1: "{res_json['OrganizationInformation']['AddressNative'].get('Street1', '')}"
                 """.strip(), language="yaml")
                 
            st.download_button(
                label=f"📥 下载 JSON ({file.name})",
                data=json.dumps(res_json, indent=2, ensure_ascii=False),
                file_name=file.name.replace(".xlsx", ".json"),
                key=f"dl_{file.name}"
            )
        except Exception as e:
            st.error(f"❌ {file.name} 处理失败: {str(e)}")





