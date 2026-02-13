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
    page_title="IATF 审计转换工具 (v26.0)",
    page_icon="🎯",
    layout="wide"
)

# --- 1. 侧边栏：强制要求导入用户自己的 JSON 模板 ---
with st.sidebar:
    st.header("⚙️ 模板配置")
    st.warning("⚠️ 必须上传您的 JSON 模板。程序不再提供默认模板。")
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
        # 如果用户没有导入模板，程序阻断运行
        st.info("请先在左侧上传 JSON 模板文件。")
        st.stop()

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
    # 深度拷贝用户导入的底稿：只替换提及部分，其余全部保留
    final_json = copy.deepcopy(template_data)
    
    try:
        xls = pd.ExcelFile(excel_file)
        # 读取约定的四张表
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

    # --- 数据提取 ---
    # 姓名提取与重排
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

    # IATF ID (AuditorId) 提取
    auditor_id = ""
    if not info_df.empty:
        for r in range(info_df.shape[0]):
            for c in range(info_df.shape[1]):
                if "IATF Card" in str(info_df.iloc[r, c]):
                    if c + 1 < info_df.shape[1]:
                        raw_val = str(info_df.iloc[r, c + 1]).strip().replace('\n', ' ')
                        auditor_id = re.sub(r'^IATF[:：\s-]*', '', raw_val, flags=re.IGNORECASE).strip()
                        if len(auditor_id) > 4: break
            if auditor_id: break

    # 日期处理 (要求：审核开始日期 / 审核结束日期)
    start_date_raw = find_val_by_key(db_df, ["审核开始日期"])
    end_date_raw = find_val_by_key(db_df, ["审核结束日期"])
    def fmt_iso(val):
        try:
            dt = pd.to_datetime(val, errors='coerce')
            if pd.notna(dt): return dt.strftime('%Y-%m-%d') + "T00:00:00.000Z"
        except: pass
        return ""
    start_iso, end_iso = fmt_iso(start_date_raw), fmt_iso(end_date_raw)
    
    # 下次审核日期 (+45天)
    next_audit_iso = ""
    try:
        end_dt = pd.to_datetime(end_date_raw, errors='coerce')
        if pd.notna(end_dt):
            next_audit_iso = (end_dt + timedelta(days=45)).strftime('%Y-%m-%d') + "T00:00:00.000Z"
    except: pass

    # 地址切分逻辑
    en_state, en_city, en_country, en_street1 = "", "", "", ""
    en_addr_raw = ""
    if not info_df.empty:
        for r in range(info_df.shape[0]):
            for c in range(info_df.shape[1]):
                val = str(info_df.iloc[r, c]).strip()
                if "审核地址" in val or "Address" in val:
                    if c + 1 < info_df.shape[1]:
                        rv = str(info_df.iloc[r, c+1]).strip()
                        if re.search(r'[a-zA-Z]', rv):
                            en_addr_raw = rv.replace('\n', ' ')
                            break
            if en_addr_raw: break

    if en_addr_raw:
        parts = [p.strip() for p in en_addr_raw.replace('，', ',').split(',') if p.strip()]
        en_parts = [re.sub(r'[\u4e00-\u9fff]', '', p).strip() for p in parts if re.search(r'[a-zA-Z]', p)]
        if en_parts:
            en_country = en_parts.pop(-1)
            streets = []
            for p in en_parts:
                if "PROVINCE" in p.upper(): en_state = p
                elif "CITY" in p.upper(): en_city = p
                else: streets.append(p)
            en_street1 = ", ".join(streets)

    # ================= 2. 定点替换逻辑 (仅替换提及字段) =================

    # 全局 UUID 与 时间戳
    final_json["uuid"] = str(uuid.uuid4())
    final_json["created"] = int(time.time() * 1000)

    # A. 审核数据定点替换 (要求路径：AuditData -> AuditData -> start/end)
    ensure_path(final_json, ["AuditData", "AuditData"])
    final_json["AuditData"]["AuditData"]["start"] = start_iso
    final_json["AuditData"]["AuditData"]["end"] = end_iso
    
    # 要求：CbIdentificationNo 对应 认证机构标识号
    final_json["AuditData"]["CbIdentificationNo"] = find_val_by_key(db_df, ["认证机构标识号"])

    # 团队信息
    if "AuditTeam" in final_json["AuditData"] and len(final_json["AuditData"]["AuditTeam"]) > 0:
        team = final_json["AuditData"]["AuditTeam"][0]
        team.update({
            "Name": auditor_name,
            "AuditorId": auditor_id,
            "AuditDaysPerformed": 1.5,
            "DatesOnSite": [{"Date": start_iso, "Day": 1}, {"Date": end_iso, "Day": 0.5}]
        })

    # B. 地址信息定点替换
    ensure_path(final_json, ["OrganizationInformation", "Address"])
    org_addr = final_json["OrganizationInformation"]["Address"]
    org_addr.update({"State": en_state, "City": en_city, "Country": en_country, "Street1": en_street1})
    org_addr["PostalCode"] = find_val_by_key(db_df, ["邮政编码"])

    # C. 组织信息
    org_info = final_json["OrganizationInformation"]
    org_info["TotalNumberEmployees"] = find_val_by_key(db_df, ["包括扩展现场在内的员工总数"])
    org_info["CertificateScope"] = find_val_by_key(db_df, ["证书范围"])

    # D. 过程清单 (每一部分 AuditNotes 加上制造/现场/远程标记)
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

    # E. 结果日期
    ensure_path(final_json, ["Results", "AuditReportFinal"])
    final_json["Results"]["AuditReportFinal"]["Date"] = end_iso
    final_json["Results"]["DateNextScheduledAudit"] = next_audit_iso

    return final_json

# ================= 主界面 =================
st.title("🚀 多模板审计转换引擎 (v26.0)")
st.write(f"当前生效模板：**{template_name}**")

uploaded_files = st.file_uploader("📥 上传 Excel 数据表", type=["xlsx"], accept_multiple_files=True)

if uploaded_files:
    st.divider()
    for file in uploaded_files:
        try:
            res_json = generate_json_logic(file, active_template)
            st.success(f"✅ {file.name} 转换成功")
            st.download_button(
                label=f"📥 下载 JSON ({file.name})",
                data=json.dumps(res_json, indent=2, ensure_ascii=False),
                file_name=file.name.replace(".xlsx", ".json"),
                key=f"dl_{file.name}"
            )
        except Exception as e:
            st.error(f"❌ {file.name} 处理失败: {str(e)}")
            








