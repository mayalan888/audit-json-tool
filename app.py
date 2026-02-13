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
    page_title="IATF 审计转换工具 (v13.0 业务修正版)",
    page_icon="🛡️",
    layout="wide"
)

# --- 核心逻辑 ---
def generate_json_logic(excel_file, template_data):
    # 使用深拷贝，确保其余未提及内容原样保留
    final_json = copy.deepcopy(template_data)
    
    try:
        xls = pd.ExcelFile(excel_file)
        # 1. 加载子表
        db_df = pd.read_excel(xls, sheet_name='数据库', header=None) if '数据库' in xls.sheet_names else pd.read_excel(xls, sheet_name=0, header=None)
        proc_df = pd.read_excel(xls, sheet_name='过程清单') if '过程清单' in xls.sheet_names else pd.DataFrame()
        info_df = pd.read_excel(xls, sheet_name='信息', header=None) if '信息' in xls.sheet_names else pd.DataFrame()
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

    # --- 1. 姓名重排逻辑 ---
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

    # --- 2. CCAA 编号 ---
    ccaa_raw = find_val_by_key(db_df, ["审核员CCAA", "CCAA"])
    caa_no = ""
    if ccaa_raw:
        match = re.search(r'(?:CCAA[:：\s-])\s*(.*)', ccaa_raw, re.IGNORECASE | re.DOTALL)
        caa_no = match.group(1).strip() if match else ccaa_raw.strip()

    # --- 3. IATF ID (AuditorId) 终极修复逻辑 ---
    auditor_id = ""
    if not info_df.empty:
        for r in range(info_df.shape[0]):
            for c in range(info_df.shape[1]):
                cell_text = str(info_df.iloc[r, c])
                if "IATF Card" in cell_text:
                    if c + 1 < info_df.shape[1]:
                        raw_val = str(info_df.iloc[r, c + 1]).strip()
                        raw_val = raw_val.replace('\n', ' ').replace('\r', ' ')
                        auditor_id = re.sub(r'^IATF[:：\s-]*', '', raw_val, flags=re.IGNORECASE).strip()
                        if len(auditor_id) > 4: 
                            break
            if auditor_id and len(auditor_id) > 4: break

    # --- 4. 日期处理 (更正为: 审核开始日期/审核结束日期) ---
    start_date_raw = find_val_by_key(db_df, ["审核开始日期"])
    end_date_raw = find_val_by_key(db_df, ["审核结束日期"])
    
    def fmt_iso(val):
        try:
            dt = pd.to_datetime(val, errors='coerce')
            if pd.notna(dt): return dt.strftime('%Y-%m-%d') + "T00:00:00.000Z"
        except: pass
        return ""
        
    start_iso, end_iso = fmt_iso(start_date_raw), fmt_iso(end_date_raw)

    # --- 5. 组装 JSON (定点替换提及字段) ---
    
    # 初始化 AuditData 容器
    if "AuditData" not in final_json: final_json["AuditData"] = {}
    
    # 【要求】：AuditData 下嵌套 AuditData 存放 start/end
    if "AuditData" not in final_json["AuditData"]: final_json["AuditData"]["AuditData"] = {}
    final_json["AuditData"]["AuditData"]["start"] = start_iso
    final_json["AuditData"]["AuditData"]["end"] = end_iso

    # 【要求】：CbIdentificationNo 对应 认证机构标识号
    final_json["AuditData"]["CbIdentificationNo"] = find_val_by_key(db_df, ["认证机构标识号"])

    # 填充团队信息 (AuditDaysPerformed 固定 1.5)
    final_json["AuditData"]["AuditTeam"] = [{
        "Name": auditor_name,
        "CaaNo": caa_no,
        "AuditorId": auditor_id,
        "AuditDaysPerformed": 1.5,
        "DatesOnSite": [{"Date": start_iso, "Day": 1}, {"Date": end_iso, "Day": 0.5}],
        "PlanningTime": "0.0000"
    }]

    # 填充组织信息
    if "OrganizationInformation" not in final_json: final_json["OrganizationInformation"] = {}
    final_json["OrganizationInformation"].update({
        "OrganizationName": find_val_by_key(db_df, ["组织名称"]),
        "IATF_USI": find_val_by_key(db_df, ["IATF USI"]),
        "TotalNumberEmployees": find_val_by_key(db_df, ["员工总数"]),
        "CertificateScope": find_val_by_key(db_df, ["证书范围"])
    })

    # 过程清单逻辑
    processes = []
    if not proc_df.empty:
        clause_cols = proc_df.columns[13:] if proc_df.shape[1] > 13 else []
        for idx, row in proc_df.iterrows():
            p_name = str(row.iloc[12]).strip()
            if not p_name or p_name.lower() == 'nan': continue
            proc_obj = {
                "Id": str(uuid.uuid4()), # 使用标准 UUID
                "ProcessName": p_name,
                "AuditNotes": [{"Id": str(uuid.uuid4()), "AuditorId": auditor_id}],
                "ManufacturingProcess": "0", "OnSiteProcess": "1", "RemoteProcess": "0"
            }
            for col in clause_cols:
                if str(row[col]).strip().upper() in ['X', 'TRUE']: proc_obj[col] = True
            processes.append(proc_obj)
    final_json["Processes"] = processes

    # 系统字段
    final_json["uuid"], final_json["created"] = str(uuid.uuid4()), int(time.time() * 1000)
    
    # 结果同步
    if "Results" not in final_json: final_json["Results"] = {}
    if "AuditReportFinal" not in final_json["Results"]: final_json["Results"]["AuditReportFinal"] = {}
    final_json["Results"]["AuditReportFinal"]["Date"] = end_iso

    return final_json

# --- UI ---
st.title("🛡️ 审计数据转换工具 (v13.0 业务修正版)")
st.caption("同步更新：审核日期路径嵌套、表头关键字更正（标识号/日期）")

uploaded_files = st.file_uploader("上传 Excel 文件", type=["xlsx"], accept_multiple_files=True)

if uploaded_files:
    template = {}
    # 优先寻找金磁.json作为底稿，保留未提及的内容
    if os.path.exists('金磁.json'):
        with open('金磁.json', 'r', encoding='utf-8') as f:
            template = json.load(f)

    for file in uploaded_files:
        try:
            res_json = generate_json_logic(file, template)
            team = res_json["AuditData"]["AuditTeam"][0]
            st.success(f"✅ {file.name} 转换成功")
            # 预览提取的核心字段
            st.code(f"姓名: {team['Name']}\nCCAA: {team['CaaNo'][:20]}...\nIATF ID: {team['AuditorId']}", language="yaml")
            st.download_button(label=f"📥 下载 JSON", data=json.dumps(res_json, indent=2, ensure_ascii=False), file_name=file.name.replace(".xlsx", ".json"), key=f"dl_{file.name}")
            st.divider()
        except Exception as e:
            st.error(f"❌ 处理失败: {e}")
            







