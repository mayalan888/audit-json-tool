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
    page_title="IATF 审计转换工具 (v14.0 修正版)",
    page_icon="📋",
    layout="wide"
)

# --- 辅助函数：安全寻址防止路径不存在时报错 ---
def ensure_path(d, path):
    current = d
    for key in path:
        if key not in current or not isinstance(current[key], dict):
            current[key] = {}
        current = current[key]
    return current

# --- 核心转换逻辑 ---
def generate_json_logic(excel_file, template_data):
    # 深度拷贝用户上传的模板：其余内容全部保留
    final_json = copy.deepcopy(template_data)
    
    try:
        xls = pd.ExcelFile(excel_file)
        # 加载核心子表
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

    # --- 1. 姓名与 ID 提取 ---
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
                        auditor_id = re.sub(r'^IATF[:：\s-]*', '', raw_val, flags=re.IGNORECASE).strip()
                        if len(auditor_id) > 4: break
            if auditor_id and len(auditor_id) > 4: break

    # --- 2. 日期处理 (更正为: 审核开始日期/审核结束日期) ---
    start_date_raw = find_val_by_key(db_df, ["审核开始日期"])
    end_date_raw = find_val_by_key(db_df, ["审核结束日期"])
    def fmt_iso(val):
        try:
            dt = pd.to_datetime(val, errors='coerce')
            if pd.notna(dt): return dt.strftime('%Y-%m-%d') + "T00:00:00.000Z"
        except: pass
        return ""
    start_iso, end_iso = fmt_iso(start_date_raw), fmt_iso(end_date_raw)

    # --- 3. 定点映射到 JSON ---
    
    # A. AuditData 模块 (修正路径为 AuditData -> AuditData -> start/end)
    ensure_path(final_json, ["AuditData", "AuditData"])
    final_json["AuditData"]["AuditData"]["start"] = start_iso
    final_json["AuditData"]["AuditData"]["end"] = end_iso
    
    # B. CbIdentificationNo (更正为: 认证机构标识号)
    final_json["AuditData"]["CbIdentificationNo"] = find_val_by_key(db_df, ["认证机构标识号"])

    # C. 团队信息
    if "AuditTeam" not in final_json["AuditData"]: final_json["AuditData"]["AuditTeam"] = [{}]
    team = final_json["AuditData"]["AuditTeam"][0]
    team.update({
        "Name": auditor_name,
        "CaaNo": caa_no,
        "AuditorId": auditor_id,
        "AuditDaysPerformed": 1.5,
        "DatesOnSite": [{"Date": start_iso, "Day": 1}, {"Date": end_iso, "Day": 0.5}],
        "PlanningTime": "0.0000"
    })

    # D. 组织信息
    ensure_path(final_json, ["OrganizationInformation"])
    final_json["OrganizationInformation"].update({
        "OrganizationName": find_val_by_key(db_df, ["组织名称"]),
        "IATF_USI": find_val_by_key(db_df, ["IATF USI"]),
        "TotalNumberEmployees": find_val_by_key(db_df, ["员工总数"]),
        "CertificateScope": find_val_by_key(db_df, ["证书范围"])
    })

    # E. 过程清单 (动态生成 ID)
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

    # F. 系统字段
    final_json["uuid"], final_json["created"] = str(uuid.uuid4()), int(time.time() * 1000)
    
    # G. 结果同步
    ensure_path(final_json, ["Results", "AuditReportFinal"])
    final_json["Results"]["AuditReportFinal"]["Date"] = end_iso

    return final_json

# --- 侧边栏：模板管理 (强制要求上传) ---
with st.sidebar:
    st.header("⚙️ 模板管理")
    st.warning("⚠️ 必须上传 JSON 模板。不再支持默认模板。")
    user_template_file = st.file_uploader("上传自定义模板 JSON", type=["json"])
    
    active_template = None
    if user_template_file:
        try:
            active_template = json.load(user_template_file)
            st.success(f"✅ 已加载模板: {user_template_file.name}")
        except Exception as e:
            st.error(f"❌ 模板解析失败: {e}")
            st.stop()
    else:
        st.stop() # 无模板则不显示主界面逻辑

# --- 主界面 ---
st.title("🚀 精准转换引擎 (v14.0 修正版)")
st.write(f"当前生效模板：**{user_template_file.name}**")

uploaded_files = st.file_uploader("上传 Excel 数据表 (可多选)", type=["xlsx"], accept_multiple_files=True)

if uploaded_files:
    st.divider()
    for excel_file in uploaded_files:
        try:
            res_json = generate_json_logic(excel_file, active_template)
            team = res_json["AuditData"]["AuditTeam"][0]
            
            with st.expander(f"📄 {excel_file.name} - 提取预览", expanded=True):
                col1, col2 = st.columns([3, 1])
                with col1:
                    st.success(f"转换成功！")
                    st.code(f"姓名: {team['Name']}\nCCAA: {team['CaaNo'][:20]}...\nIATF ID: {team['AuditorId']}", language="yaml")
                with col2:
                    st.download_button(
                        label="📥 下载 JSON",
                        data=json.dumps(res_json, indent=2, ensure_ascii=False),
                        file_name=excel_file.name.replace(".xlsx", ".json"),
                        key=f"dl_{excel_file.name}"
                    )
        except Exception as e:
            st.error(f"❌ {excel_file.name} 处理失败: {e}")






