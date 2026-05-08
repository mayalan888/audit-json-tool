import streamlit as st
import pandas as pd
import json
import uuid
import time
import re
import copy
from datetime import datetime, timedelta

# =====================================================================
# 页面配置
# =====================================================================
st.set_page_config(
    page_title="IATF 审计转换工具 (v70.8 智能 Key 匹配修复版)",
    page_icon="🛡️",
    layout="wide"
)

# =====================================================================
# 辅助函数区
# =====================================================================
def ensure_path(d, path):
    current = d
    for key in path:
        if key not in current or not isinstance(current[key], dict):
            current[key] = {}
        current = current[key]
    return current

def safe_get(obj, key, default=""):
    if isinstance(obj, dict):
        return obj.get(key, default)
    return default

def extract_and_format_english_name(raw_val):
    clean_val = str(raw_val).replace("姓名:", "").replace("Name:", "").strip()
    if not clean_val: return ""
    eng_only = re.sub(r'[^a-zA-Z\s]', ' ', clean_val).strip()
    eng_only = re.sub(r'\s+', ' ', eng_only)
    if eng_only:
        parts = eng_only.split()
        if len(parts) >= 2 and parts[0].isupper() and not parts[1].isupper():
            return f"{parts[1]} {parts[0]}"
        else:
            return eng_only
    return clean_val

def parse_chinese_address(addr_str):
    province, city, street = "", "", addr_str
    if not addr_str: return province, city, street
    clean_addr = re.sub(r'^中国', '', addr_str).strip()
    p_match = re.search(r'(.+?(省|自治区|北京|上海|天津|重庆))', clean_addr)
    if p_match:
        province = p_match.group(1).strip()
        if province in ["北京", "上海", "天津", "重庆"]: province += "市"
        remain_addr = clean_addr[len(p_match.group(1)):].strip()
        c_match = re.search(r'(.+?(市|地区|盟|自治州|州))', remain_addr)
        if c_match:
            city = c_match.group(1).strip()
            street = remain_addr[len(city):].strip()
        else:
            if "市" in province: city = province
            street = remain_addr
    return province, city, street

# =====================================================================
# 场所提取逻辑 (继承自 v70.5 纯净版)
# =====================================================================
def extract_ems_sites(info_df):
    ems_sites = []
    if info_df.empty: return ems_sites
    row_start, row_end = 20, min(25, info_df.shape[0])
    col_map = {}
    header_r = -1
    for r in range(row_start, row_end):
        for c in range(5, min(13, info_df.shape[1])):
            val = str(info_df.iloc[r, c]).strip().upper()
            if "EMS" in val or "扩展" in val:
                header_r = r
                for cs in range(5, min(13, info_df.shape[1])):
                    h = str(info_df.iloc[r, cs]).strip()
                    if "中文名称" in h: col_map['name_cn'] = cs
                    elif "英文名称" in h: col_map['name_en'] = cs
                    elif "中文地址" in h: col_map['addr_cn'] = cs
                    elif "英文地址" in h: col_map['addr_en'] = cs
                    elif "邮编" in h: col_map['zip'] = cs
                    elif "USI" in h.upper(): col_map['usi'] = cs
                    elif "人数" in h: col_map['emp'] = cs
                break
        if header_r != -1: break
    if header_r != -1:
        for r in range(header_r + 1, row_end):
            name_cn = str(info_df.iloc[r, col_map.get('name_cn', -1)]).strip() if 'name_cn' in col_map else ""
            addr_cn = str(info_df.iloc[r, col_map.get('addr_cn', -1)]).strip() if 'addr_cn' in col_map else ""
            if name_cn == "nan" or not name_cn: continue
            p, city, s = parse_chinese_address(addr_cn)
            site = {
                "Id": str(uuid.uuid4()), "SiteName": name_cn, "TotalNumberEmployees": str(info_df.iloc[r, col_map.get('emp', -1)]),
                "AddressNative": {"Street1": s, "City": city, "State": p, "Country": "中国"}
            }
            ems_sites.append(site)
    return ems_sites

def extract_rl_sites(info_df):
    sites = []
    if info_df.empty: return sites
    row_start, row_end = 26, min(32, info_df.shape[0])
    col_map = {}
    header_r = -1
    for r in range(row_start, row_end):
        for c in range(5, min(14, info_df.shape[1])):
            val = str(info_df.iloc[r, c]).strip().upper()
            if "RL" in val or "支持场所" in val:
                header_r = r
                for cs in range(5, min(14, info_df.shape[1])):
                    h = str(info_df.iloc[r, cs]).strip()
                    if "中文名称" in h: col_map['name_cn'] = cs
                    elif "中文地址" in h: col_map['addr_cn'] = cs
                    elif "人数" in h: col_map['emp'] = cs
                break
        if header_r != -1: break
    if header_r != -1:
        for r in range(header_r + 1, row_end):
            name_cn = str(info_df.iloc[r, col_map.get('name_cn', -1)]).strip() if 'name_cn' in col_map else ""
            addr_cn = str(info_df.iloc[r, col_map.get('addr_cn', -1)]).strip() if 'addr_cn' in col_map else ""
            if name_cn == "nan" or not name_cn: continue
            p, city, s = parse_chinese_address(addr_cn)
            sites.append({
                "Id": str(uuid.uuid4()), "SiteName": name_cn, "TotalNumberEmployees": str(info_df.iloc[r, col_map.get('emp', -1)]),
                "AddressNative": {"Street1": s, "City": city, "State": p, "Country": "中国"}
            })
    return sites

# =====================================================================
# 主逻辑
# =====================================================================
def generate_json_logic(excel_file, base_data, mode):
    final_json = copy.deepcopy(base_data)
    try:
        xls = pd.ExcelFile(excel_file)
        db_df = pd.read_excel(xls, sheet_name='数据库', header=None) if '数据库' in xls.sheet_names else pd.read_excel(xls, sheet_name=0, header=None)
        proc_df = pd.read_excel(xls, sheet_name='过程清单') if '过程清单' in xls.sheet_names else pd.DataFrame()
        info_df = pd.read_excel(xls, sheet_name='信息', header=None) if '信息' in xls.sheet_names else pd.DataFrame()
        perf_df = pd.read_excel(xls, sheet_name='过程绩效', header=None) if '过程绩效' in xls.sheet_names else pd.DataFrame()
        scope_df = pd.read_excel(xls, sheet_name='范围', header=None) if '范围' in xls.sheet_names else pd.DataFrame()
    except Exception as e:
        raise ValueError(f"Excel 读取失败: {str(e)}")

    def find_val_by_key(df, keywords, col_offset=1):
        for r in range(df.shape[0]):
            for c in range(df.shape[1]):
                if any(k in str(df.iloc[r, c]) for k in keywords):
                    return str(df.iloc[r, c + col_offset]).strip() if c + col_offset < df.shape[1] else ""
        return ""

    def get_db_val(r, c):
        try: return str(db_df.iloc[r, c]).strip() if pd.notna(db_df.iloc[r, c]) else ""
        except: return ""

    # [基础信息提取]
    raw_name_full = find_val_by_key(db_df, ["姓名"]) or get_db_val(5, 1)
    formatted_team_name = extract_and_format_english_name(raw_name_full)
    start_iso = pd.to_datetime(find_val_by_key(db_df, ["开始日期"]) or get_db_val(2,1)).strftime('%Y-%m-%dT00:00:00.000Z')
    end_iso = pd.to_datetime(find_val_by_key(db_df, ["结束日期"]) or get_db_val(3,1)).strftime('%Y-%m-%dT00:00:00.000Z')

    # 💥💥💥 [核心修复：中英文范围智能 Key 匹配逻辑] 💥💥💥
    cert_scope_en, cert_scope_cn = "", ""
    if not scope_df.empty:
        for r in range(scope_df.shape[0]):
            for c in range(scope_df.shape[1]):
                v = str(scope_df.iloc[r, c])
                if "审核范围" in v:
                    val = str(scope_df.iloc[r, c+1]).strip() if c+1 < scope_df.shape[1] else ""
                    if "英" in v or "En" in v: cert_scope_en = val
                    else: cert_scope_cn = val

    org = ensure_path(final_json, ["OrganizationInformation"])
    
    # 智能搜索底座中已存在的 Key，防止生成新节点
    target_eng_key, target_native_key = "CertificateScope", "CertificateScopeNative"
    for k in org.keys():
        k_low = k.lower()
        if "scope" in k_low:
            if "native" in k_low or "zh" in k_low or "cn" in k_low: target_native_key = k
            else: target_eng_key = k

    if cert_scope_en: org[target_eng_key] = cert_scope_en
    if cert_scope_cn: org[target_native_key] = cert_scope_cn
    # ==============================================================

    # [过程绩效 KPI 提取 - 带锁]
    kpi_map = {}
    if not perf_df.empty:
        col_m = {'proc':-1, 'kpi':-1, 'target':-1, 'result':-1, 'trend':-1}
        for r in range(min(10, perf_df.shape[0])):
            for c in range(perf_df.shape[1]):
                v = str(perf_df.iloc[r, c]).upper()
                if "过程" in v or "KPI" in v:
                    for sc in range(perf_df.shape[1]):
                        hv = str(perf_df.iloc[r, sc]).upper()
                        if "过程" in hv and col_m['proc']==-1: col_m['proc']=sc
                        elif "KPI" in hv and col_m['kpi']==-1: col_m['kpi']=sc
                        elif "目标" in hv and col_m['target']==-1: col_m['target']=sc
                        elif "结果" in hv and col_m['result']==-1: col_m['result']=sc
                    break
        if col_m['proc']!=-1:
            curr_p = ""
            for r in range(1, perf_df.shape[0]):
                p_val = str(perf_df.iloc[r, col_m['proc']]).strip()
                if p_val and p_val != "nan": curr_p = p_val
                k_val = str(perf_df.iloc[r, col_m['kpi']]).strip()
                if k_val and k_val != "nan":
                    if curr_p not in kpi_map: kpi_map[curr_p] = []
                    kpi_map[curr_p].append({"KPI": k_val, "CurrentTarget": str(perf_df.iloc[r, col_m['target']]), "Results": str(perf_df.iloc[r, col_m['result']])})

    # [过程深度融合逻辑 - 继承自 v70.6]
    if not proc_df.empty:
        new_procs = []
        base_proc_map = {re.sub(r'\s+', '', p.get("ProcessName", "")): p for p in final_json.get("Processes", []) if isinstance(p, dict)}
        for idx, row in proc_df.iterrows():
            name = str(row.iloc[0]).strip()
            if not name or name=="nan": continue
            clean_name = re.sub(r'\s+', '', name)
            proc_obj = base_proc_map.get(clean_name, {"Id": str(uuid.uuid4()), "ProcessName": name, "AuditNotes": []})
            proc_obj["ProcessName"] = name
            proc_obj["RepresentativeName"] = str(row.iloc[2])
            # 注入匹配的 KPI
            for k, v in kpi_map.items():
                if clean_name in re.sub(r'\s+', '', k):
                    proc_obj["ProcessPerformance"] = v
                    break
            new_procs.append(proc_obj)
        final_json["Processes"] = new_procs

    # [主地址拆分]
    addr_native = find_val_by_key(db_df, ["地址"]) or get_db_val(10, 1)
    p, city, s = parse_chinese_address(addr_native)
    org_addr = ensure_path(org, ["AddressNative"])
    org_addr.update({"Street1": s, "City": city, "State": p, "Country": "中国"})

    # [场所处理]
    if "全量" in mode:
        final_json["ExtendedManufacturingSites"] = extract_ems_sites(info_df)
        final_json["ProvidingSupportSites"] = extract_rl_sites(info_df)

    final_json["uuid"] = str(uuid.uuid4())
    return final_json

# =====================================================================
# 界面
# =====================================================================
st.title("🛡️ IATF 转换引擎 v70.8 (智能 Key 匹配修复版)")
with st.sidebar:
    run_mode = st.radio("模式", ("纯净标准模式", "全量综合模式"), index=1)
    u_file = st.file_uploader("底座 JSON", type=["json"])
    if not u_file: st.stop()
    base_template = json.load(u_file)

uploaded_files = st.file_uploader("上传 Excel", type=["xlsx"], accept_multiple_files=True)
if uploaded_files:
    for f in uploaded_files:
        try:
            res = generate_json_logic(f, base_template, run_mode)
            st.success(f"✅ {f.name} 解析成功")
            st.download_button(f"📥 下载 {f.name}.json", json.dumps(res, indent=2, ensure_ascii=False))
        except Exception as e:
            st.error(f"❌ 失败: {e}")
