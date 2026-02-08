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
    page_title="IATF 审计转换工具 (v12.0)",
    page_icon="🛡️",
    layout="wide"
)

# --- 核心转换逻辑 ---
def generate_json_logic(excel_file, template_data):
    final_json = copy.deepcopy(template_data)
    
    try:
        xls = pd.ExcelFile(excel_file)
        # 1. 读取工作表 (智能索引)
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

    # --- 2. CCAA 编号 (完整提取逗号前后内容) ---
    ccaa_raw = find_val_by_key(db_df, ["审核员CCAA", "CCAA"])
    caa_no = ""
    if ccaa_raw:
        # 匹配 CCAA: 之后的所有字符 (跨行匹配)
        match = re.search(r'(?:CCAA[:：\s-])\s*(.*)', ccaa_raw, re.IGNORECASE | re.DOTALL)
        caa_no = match.group(1).strip() if match else ccaa_raw.strip()

    # --- 3. IATF ID (AuditorId) 精准逻辑升级 ---
    auditor_id = ""
    if not info_df.empty:
        for r in range(info_df.shape[0]):
            for c in range(info_df.shape[1]):
                cell_text = str(info_df.iloc[r, c])
                
                # 步骤 A：定位包含 "IATF Card" 的格子
                if "IATF Card" in cell_text or "IATF卡号" in cell_text:
                    # 优先取右侧格子的内容
                    target_text = ""
                    if c + 1 < info_df.shape[1]:
                        val = str(info_df.iloc[r, c+1]).strip()
                        if val and val.lower() != 'nan':
                            target_text = val
                    
                    # 兜底：如果右侧为空，可能在当前格内
                    if not target_text:
                        target_text = cell_text
                    
                    # 步骤 B：寻找该字符串中【最后一次】出现 IATF 之后的内容
                    # 解决 GZH 这种截断问题，确保抓取 IATF 后面的长字符串
                    if "IATF" in target_text.upper():
                        # 找到 IATF 的位置 (不区分大小写)
                        pos = target_text.upper().rfind("IATF")
                        # 截取从 IATF 之后开始的所有字符
                        after_iatf = target_text[pos + 4:] 
                        # 清除开头的冒号、空格、横杠
                        auditor_id = re.sub(r'^[:：\s-]+', '', after_iatf).strip()
                    else:
                        # 如果没有 IATF 关键字，则取整个格子
                        auditor_id = target_text
                    break
            if auditor_id: break

    # --- 4. 日期处理 ---
    start_date_raw = find_val_by_key(db_df, ["审核开始时间"])
    end_date_raw = find_val_by_key(db_df, ["审核结束时间"])
    
    def fmt_iso(val):
        try:
            dt = pd.to_datetime(val, errors='coerce')
            if pd.notna(dt): return dt.strftime('%Y-%m-%d') + "T00:00:00.000Z"
        except: pass
        return ""

    start_iso = fmt_iso(start_date_raw)
    end_iso = fmt_iso(end_date_raw)

    # --- 5. 组装 JSON ---
    if "AuditData" not in final_json: final_json["AuditData"] = {}
    
    final_json["AuditData"].update({
        "AuditDate": {"Start": start_iso, "End": end_iso},
        "CbIdentificationNo": find_val_by_key(db_df, ["认证机构识别号"]),
        "AuditTeam": [{
            "Name": auditor_name,
            "CaaNo": caa_no,
            "AuditorId": auditor_id,        # 最终输出截图中的编号
            "AuditDaysPerformed": 1.5,
            "DatesOnSite": [{"Date": start_iso, "Day": 1}, {"Date": end_iso, "Day": 0.5}],
            "PlanningTime": "0.0000"
        }]
    })

    if "OrganizationInformation" not in final_json: final_json["OrganizationInformation"] = {}
    final_json["OrganizationInformation"].update({
        "OrganizationName": find_val_by_key(db_df, ["组织名称"]),
        "IATF_USI": find_val_by_key(db_df, ["IATF USI"]),
        "TotalNumberEmployees": find_val_by_key(db_df, ["员工总数"]),
        "CertificateScope": find_val_by_key(db_df, ["证书范围"])
    })

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
                if str(row[col]).strip().upper() in ['X', 'TRUE']:
                    proc_obj[col] = True
            processes.append(proc_obj)

    final_json["Processes"] = processes
    final_json["uuid"], final_json["created"] = str(uuid.uuid4()), int(time.time() * 1000)
    
    if "Results" not in final_json: final_json["Results"] = {}
    final_json["Results"]["AuditReportFinal"] = {"Date": end_iso}

    return final_json

# --- Streamlit UI ---
st.title("🛡️ 审计数据转换工具 (v12.0)")
st.caption("修复: 采用右侧单元格重定向逻辑，确保完整提取 IATF ID")

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
            st.code(f"""
姓名: {team['Name']}
CCAA: {team['CaaNo']}
IATF ID: {team['AuditorId']}
            """, language="yaml")
            st.download_button(label=f"📥 下载 JSON ({file.name})", data=json.dumps(res_json, indent=2, ensure_ascii=False), file_name=file.name.replace(".xlsx", ".json"))
            st.divider()
        except Exception as e:
            st.error(f"❌ 处理失败: {e}")










