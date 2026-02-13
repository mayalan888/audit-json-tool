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
    page_title="IATF 审计转换工具 (v13.0)",
    page_icon="🛡️",
    layout="wide"
)

# --- 核心逻辑 ---
def generate_json_logic(excel_file, template_data):
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
                # 寻找包含 "IATF Card" 的格子
                if "IATF Card" in cell_text:
                    # 强制锁定右边那个单元格
                    if c + 1 < info_df.shape[1]:
                        raw_val = str(info_df.iloc[r, c + 1]).strip()
                        # A. 替换所有换行符为空格，防止截断
                        raw_val = raw_val.replace('\n', ' ').replace('\r', ' ')
                        # B. 精准剔除 "IATF:" 前缀
                        # 匹配字符串开头的 IATF 以及随后的冒号/空格，并将其替换为空
                        auditor_id = re.sub(r'^IATF[:：\s-]*', '', raw_val, flags=re.IGNORECASE).strip()
                        # 如果提取结果太短（比如只有 GZH），说明可能找错行了，继续搜寻
                        if len(auditor_id) > 4: 
                            break
            if auditor_id and len(auditor_id) > 4: break

    # --- 4. 组装 JSON ---
    start_date_raw = find_val_by_key(db_df, ["审核开始时间"])
    end_date_raw = find_val_by_key(db_df, ["审核结束时间"])
    def fmt_iso(val):
        try:
            dt = pd.to_datetime(val, errors='coerce')
            if pd.notna(dt): return dt.strftime('%Y-%m-%d') + "T00:00:00.000Z"
        except: pass
        return ""
    start_iso, end_iso = fmt_iso(start_date_raw), fmt_iso(end_date_raw)

    if "AuditData" not in final_json: final_json["AuditData"] = {}
    final_json["AuditData"].update({
        "AuditDate": {"Start": start_iso, "End": end_iso},
        "CbIdentificationNo": find_val_by_key(db_df, ["认证机构识别号"]),
        "AuditTeam": [{
            "Name": auditor_name,
            "CaaNo": caa_no,
            "AuditorId": auditor_id,        # 映射最终纯净编号
            "AuditDaysPerformed": 1.5,
            "DatesOnSite": [{"Date": start_iso, "Day": 1}, {"Date": end_iso, "Day": 0.5}],
            "PlanningTime": "0.0000"
        }]
    })

    # 填充其他信息
    if "OrganizationInformation" not in final_json: final_json["OrganizationInformation"] = {}
    final_json["OrganizationInformation"].update({
        "OrganizationName": find_val_by_key(db_df, ["组织名称"]),
        "IATF_USI": find_val_by_key(db_df, ["IATF USI"]),
        "TotalNumberEmployees": find_val_by_key(db_df, ["员工总数"]),
        "CertificateScope": find_val_by_key(db_df, ["证书范围"])
    })

    # 过程清单逻辑 (引用 AuditorId)
    processes = []
    if not proc_df.empty:
        clause_cols = proc_df.columns[13:] if proc_df.shape[1] > 13 else []
        for idx, row in proc_df.iterrows():
            p_name = str(row.iloc[12]).strip()
            if not p_name or p_name.lower() == 'nan': continue
            proc_obj = {
                "Id": str(int(time.time() * 1000) + idx),
                "ProcessName": p_name,
                "AuditNotes": [{"Id": int(time.time()*1000)+idx+200, "AuditorId": auditor_id}],
                "ManufacturingProcess": "0", "OnSiteProcess": "1", "RemoteProcess": "0"
            }
            for col in clause_cols:
                if str(row[col]).strip().upper() in ['X', 'TRUE']: proc_obj[col] = True
            processes.append(proc_obj)

    final_json["Processes"] = processes
    final_json["uuid"], final_json["created"] = str(uuid.uuid4()), int(time.time() * 1000)
    if "Results" not in final_json: final_json["Results"] = {}
    final_json["Results"]["AuditReportFinal"] = {"Date": end_iso}

    return final_json

# --- UI ---
st.title("🛡️ 审计数据转换工具 (v13.0)")
st.caption("修复：强制锁定『信息』表右侧单元格，解决 GZH 截断与识别偏移问题")

uploaded_files = st.file_uploader("上传 Excel 文件", type=["xlsx"], accept_multiple_files=True)

if uploaded_files:
    template = {}
    if os.path.exists('金磁.json'):
        with open('金磁.json', 'r', encoding='utf-8') as f:
            template = json.load(f)

    for file in uploaded_files:
        try:
            res_json = generate_json_logic(file, template)
            team = res_json["AuditData"]["AuditTeam"][0]
            st.success(f"✅ {file.name} 转换成功")
            st.code(f"姓名: {team['Name']}\nCCAA: {team['CaaNo']}\nIATF ID: {team['AuditorId']}", language="yaml")
            st.download_button(label=f"📥 下载 JSON", data=json.dumps(res_json, indent=2, ensure_ascii=False), file_name=file.name.replace(".xlsx", ".json"))
            st.divider()
        except Exception as e:
            st.error(f"❌ 处理失败: {e}")









