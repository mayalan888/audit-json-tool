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
    page_title="IATF 审计转换工具 (v28.0)",
    page_icon="🎯",
    layout="wide"
)

# --- 1. 侧边栏：强制要求用户导入自己的 JSON 模板 ---
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

    # ================= 数据提取 =================
    
    # 1. 姓名提取与重排
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

    # 2. CCAA 提取
    ccaa_raw = find_val_by_key(db_df, ["审核员CCAA", "CCAA"])
    caa_no = ""
    if ccaa_raw:
        match = re.search(r'(?:CCAA[:：\s-])\s*(.*)', ccaa_raw, re.IGNORECASE | re.DOTALL)
        caa_no = match.group(1).strip() if match else ccaa_raw.strip()

    # 3. IATF ID (AuditorId) 提取
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

    # 4. 日期处理
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

    # 💥 5. 英文地址提取与按倒序位置拆分 (移植自旧版代码)
    candidates = []
    # 扫描数据库表中可能的地址候选
    for r in range(db_df.shape[0]):
        for c in range(db_df.shape[1]):
            val = str(db_df.iloc[r, c])
            if "地址" in val or "Address" in val:
                if c+1 < db_df.shape[1]: candidates.append(str(db_df.iloc[r, c+1]).strip())
                if c+4 < db_df.shape[1]: candidates.append(str(db_df.iloc[r, c+4]).strip())
                
    native_street = ""
    english_address = ""
    def is_chinese(s): return bool(re.search(r'[\u4e00-\u9fff]', s))

    # 选出中文地址
    zh_candidates = [c for c in candidates if c and is_chinese(c)]
    if zh_candidates: native_street = max(zh_candidates, key=len)
        
    # 选出纯英文地址
    en_candidates = [c for c in candidates if c and not is_chinese(c)]
    if en_candidates: english_address = max(en_candidates, key=len)
        
    # 如果没找到英文地址，去“信息”表里找
    if not english_address and not info_df.empty:
         for r in range(len(info_df)):
            for c in range(len(info_df.columns)):
                cell_val = str(info_df.iloc[r, c])
                if "审核地址" in cell_val or "Audit Address" in cell_val:
                    if c + 1 < len(info_df.columns):
                        candidate = str(info_df.iloc[r, c+1]).strip()
                        if candidate and not is_chinese(candidate):
                            english_address = candidate
                            break
            if english_address: break

    # 按倒序位置拆分地址
    street = english_address
    city = ""
    state = ""
    country = ""
    
    if english_address:
        # 将可能断开的换行符和全角逗号处理掉
        clean_eng = english_address.replace('\n', ' ').replace('\r', ' ').replace('，', ',')
        parts = [p.strip() for p in clean_eng.split(',') if p.strip()]
        
        # 倒序分配逻辑
        if len(parts) >= 3:
            country = parts[-1]
            state = parts[-2]
            city = parts[-3]
            street = ", ".join(parts[:-3])
        else:
            street = english_address

    # ================= 定点替换逻辑 =================

    final_json["uuid"] = str(uuid.uuid4())
    final_json["created"] = int(time.time() * 1000)

    # A. 审核数据
    ensure_path(final_json, ["AuditData", "AuditData"])
    final_json["AuditData"]["AuditData"]["start"] = start_iso
    final_json["AuditData"]["AuditData"]["end"] = end_iso
    final_json["AuditData"]["CbIdentificationNo"] = find_val_by_key(db_df, ["认证机构标识号"])

    if "AuditTeam" in final_json["AuditData"] and len(final_json["AuditData"]["AuditTeam"]) > 0:
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
    
    org.update({
        "TotalNumberEmployees": find_val_by_key(db_df, ["包括扩展现场在内的员工总数"]),
        "CertificateScope": find_val_by_key(db_df, ["证书范围"])
    })
    
    org["AddressNative"].update({
        "Street1": native_street,
        "Country": "中国",
        "PostalCode": find_val_by_key(db_df, ["邮政编码"])
    })
    
    org["Address"].update({
        "State": state, 
        "City": city, 
        "Country": country, 
        "Street1": street,
        "PostalCode": find_val_by_key(db_df, ["邮政编码"])
    })

    # C. 过程清单重建
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

    # D. 结果日期
    ensure_path(final_json, ["Results", "AuditReportFinal"])
    final_json["Results"]["AuditReportFinal"]["Date"] = end_iso
    final_json["Results"]["DateNextScheduledAudit"] = next_audit_iso

    return final_json

# ================= 主界面 =================
st.title("🚀 多模板审计转换引擎 (v28.0 倒序切分版)")
st.write(f"当前生效模板：**{template_name}**")

uploaded_files = st.file_uploader("📥 上传 Excel 数据表", type=["xlsx"], accept_multiple_files=True)

if uploaded_files:
    st.divider()
    for file in uploaded_files:
        try:
            res_json = generate_json_logic(file, active_template)
            team = res_json["AuditData"]["AuditTeam"][0]
            st.success(f"✅ {file.name} 转换成功")
            with st.expander("查看提取预览", expanded=False):
                 st.code(f"姓名: {team['Name']}\nID: {team['AuditorId']}\nCity: {res_json['OrganizationInformation']['Address']['City']}\nCountry: {res_json['OrganizationInformation']['Address']['Country']}", language="yaml")
            st.download_button(
                label=f"📥 下载 JSON ({file.name})",
                data=json.dumps(res_json, indent=2, ensure_ascii=False),
                file_name=file.name.replace(".xlsx", ".json"),
                key=f"dl_{file.name}"
            )
        except Exception as e:
            st.error(f"❌ {file.name} 处理失败: {str(e)}")










