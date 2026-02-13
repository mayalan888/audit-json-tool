import streamlit as st
import pandas as pd
import json
import uuid
import time
import os
import copy
import re
from datetime import datetime, timedelta

st.set_page_config(page_title="IATF 审计转换工具 (v27.0)", page_icon="🎯", layout="wide")

# --- 1. 侧边栏：强制要求上传模板 ---
with st.sidebar:
    st.header("⚙️ 模板配置")
    st.warning("⚠️ 必须上传 JSON 模板。程序将用生成的内容替换模板中的匹配条目。")
    user_template = st.file_uploader("上传自定义 JSON 模板", type=["json"])
    
    active_template = None
    if user_template:
        try:
            active_template = json.load(user_template)
            st.success(f"✅ 成功加载模板: {user_template.name}")
        except Exception as e:
            st.error(f"❌ 模板解析失败: {e}")
            st.stop()
    else:
        st.stop()

# --- 2. 深度替换函数：对于相同的条目，用 source 替换 target ---
def deep_patch_replace(target, source):
    """
    递归遍历：如果 source 中的路径在 target 中也存在，则替换 target 的值。
    """
    if isinstance(target, dict) and isinstance(source, dict):
        for key in source.keys():
            if key in target:
                if isinstance(source[key], dict) and isinstance(target[key], dict):
                    deep_patch_replace(target[key], source[key])
                elif isinstance(source[key], list) and isinstance(target[key], list):
                    # 对于列表（如团队、过程），通常业务上以数据表为准，直接替换匹配索引
                    for i in range(min(len(target[key]), len(source[key]))):
                        if isinstance(target[key][i], (dict, list)) and isinstance(source[key][i], (dict, list)):
                            deep_patch_replace(target[key][i], source[key][i])
                        else:
                            target[key][i] = copy.deepcopy(source[key][i])
                    # 如果生成的列表更长，追加剩余部分
                    if len(source[key]) > len(target[key]):
                        target[key].extend(copy.deepcopy(source[key][len(target[key]):]))
                else:
                    target[key] = copy.deepcopy(source[key])
    return target

# --- 3. 核心业务提取逻辑 (生成默认内容) ---
def generate_default_content(excel_file):
    xls = pd.ExcelFile(excel_file)
    db_df = pd.read_excel(xls, sheet_name='数据库', header=None) if '数据库' in xls.sheet_names else pd.read_excel(xls, sheet_name=0, header=None)
    proc_df = pd.read_excel(xls, sheet_name='过程清单') if '过程清单' in xls.sheet_names else pd.DataFrame()
    info_df = pd.read_excel(xls, sheet_name='信息', header=None) if '信息' in xls.sheet_names else pd.DataFrame()
    
    doc_list_df = pd.DataFrame()
    if len(xls.sheet_names) >= 9:
        doc_list_df = pd.read_excel(xls, sheet_name=xls.sheet_names[8], header=None)

    def find_val(df, keywords, offset=1):
        if df.empty: return ""
        for r in range(df.shape[0]):
            for c in range(df.shape[1]):
                if any(k in str(df.iloc[r, c]) for k in keywords):
                    return str(df.iloc[r, c + offset]).strip() if c + offset < df.shape[1] else ""
        return ""

    # (A) 姓名重排
    raw_name = find_val(db_df, ["姓名", "Auditor Name"]).replace("姓名:", "").replace("Name:", "").strip()
    aud_name = raw_name
    en_part = re.sub(r'[\u4e00-\u9fff]', '', raw_name).strip()
    if en_part:
        pts = en_part.split()
        aud_name = f"{pts[1]} {pts[0]}" if len(pts) >= 2 and pts[0].isupper() and not pts[1].isupper() else en_part

    # (B) IATF ID
    aud_id = ""
    for r in range(info_df.shape[0]):
        for c in range(info_df.shape[1]):
            if "IATF Card" in str(info_df.iloc[r, c]):
                rv = str(info_df.iloc[r, c+1]).strip().replace('\n', ' ')
                aud_id = re.sub(r'^IATF[:：\s-]*', '', rv, flags=re.IGNORECASE).strip()
                if len(aud_id) > 4: break
        if aud_id: break

    # (C) 日期与 +45天
    s_iso = fmt_iso = lambda v: pd.to_datetime(v, errors='coerce').strftime('%Y-%m-%d') + "T00:00:00.000Z" if pd.notna(pd.to_datetime(v, errors='coerce')) else ""
    start_iso = fmt_iso(find_val(db_df, ["审核开始日期"]))
    end_iso = fmt_iso(find_val(db_df, ["审核结束日期"]))
    next_iso = (pd.to_datetime(find_val(db_df, ["审核结束日期"])) + timedelta(days=45)).strftime('%Y-%m-%d') + "T00:00:00.000Z" if start_iso else ""

    # (D) 地址智能拆分
    e_state, e_city, e_country, e_street = "", "", "", ""
    ea_raw = ""
    for r in range(info_df.shape[0]):
        for c in range(info_df.shape[1]):
            if "审核地址" in str(info_df.iloc[r, c]) and re.search(r'[a-zA-Z]', str(info_df.iloc[r, c+1])):
                ea_raw = str(info_df.iloc[r, c+1]).strip().replace('\n', ' ')
                break
    if ea_raw:
        pts = [re.sub(r'[\u4e00-\u9fff]', '', p).strip() for p in ea_raw.replace('，', ',').split(',') if re.search(r'[a-zA-Z]', p)]
        if pts:
            e_country = pts.pop(-1)
            streets = []
            for p in pts:
                if "PROVINCE" in p.upper(): e_state = p
                elif "CITY" in p.upper(): e_city = p
                else: streets.append(p)
            e_street = ", ".join(streets)

    # (E) 构建默认生成的 JSON 内容 (Source)
    generated = {
        "uuid": str(uuid.uuid4()),
        "created": int(time.time() * 1000),
        "AuditData": {
            "AuditData": {"start": start_iso, "end": end_iso}, #
            "CbIdentificationNo": find_val(db_df, ["认证机构标识号"]), #
            "AuditTeam": [{"Name": aud_name, "AuditorId": aud_id, "AuditDaysPerformed": 1.5}]
        },
        "OrganizationInformation": {
            "TotalNumberEmployees": find_val(db_df, ["包括扩展现场在内的员工总数"]),
            "CertificateScope": find_val(db_df, ["证书范围"]),
            "Address": {"State": e_state, "City": e_city, "Country": e_country, "Street1": e_street, "PostalCode": find_val(db_df, ["邮政编码"])},
            "AddressNative": {"Street1": find_val(db_df, ["街道1"]), "Country": "中国", "PostalCode": find_val(db_df, ["邮政编码"])}
        },
        "Results": {
            "AuditReportFinal": {"Date": end_iso},
            "DateNextScheduledAudit": next_iso
        }
    }
    
    # 填充文件清单
    docs = []
    for c in range(doc_list_df.shape[1]):
        for r in range(doc_list_df.shape[0]):
            if "公司内对应的程序文件" in str(doc_list_df.iloc[r, c]):
                for r2 in range(r+1, doc_list_df.shape[0]):
                    v = str(doc_list_df.iloc[r2, c]).strip()
                    if v and v.lower() != 'nan': docs.append({"DocumentName": v})
                break
    generated["Stage1DocumentedRequirements"] = {"IatfClauseDocuments": docs}

    # 填充过程清单
    procs = []
    if not proc_df.empty:
        cols = proc_df.columns[13:]
        for i, row in proc_df.iterrows():
            if str(row.iloc[12]).strip().lower() in ['', 'nan']: continue
            p_obj = {"Id": str(uuid.uuid4()), "ProcessName": str(row.iloc[12]), "AuditNotes": [{"Id": str(uuid.uuid4()), "AuditorId": aud_id, "ManufacturingProcess": "0", "OnSiteProcess": "1", "RemoteProcess": "0"}]}
            for col in cols:
                if str(row[col]).strip().upper() in ['X', 'TRUE']: p_obj[col] = True
            procs.append(p_obj)
    generated["Processes"] = procs

    return generated

# --- 4. 主界面 ---
st.title("🚀 IATF 审计数据深度补丁转换引擎 (v27.0)")
st.write(f"当前合并底稿：**{user_template.name}**")

files = st.file_uploader("📥 上传 Excel 数据表", type=["xlsx"], accept_multiple_files=True)

if files:
    st.divider()
    for f in files:
        try:
            # 1. 先生成默认内容
            generated_content = generate_default_content(f)
            # 2. 将上传的模板作为 target，进行深度合并替换
            final_output = deep_patch_replace(copy.deepcopy(active_template), generated_content)
            
            with st.expander(f"📄 {f.name} - 合并成功", expanded=True):
                col1, col2 = st.columns([3, 1])
                with col1:
                    st.success("已完成深度替换")
                    team = generated_content["AuditData"]["AuditTeam"][0]
                    st.code(f"姓名: {team['Name']} | ID: {team['AuditorId']}", language="yaml")
                with col2:
                    st.download_button("📥 下载 JSON", data=json.dumps(final_output, indent=2, ensure_ascii=False), file_name=f.name.replace(".xlsx", ".json"), key=f"dl_{f.name}")
        except Exception as e:
            st.error(f"❌ {f.name} 处理失败: {str(e)}")









