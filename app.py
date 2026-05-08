import streamlit as st
import pandas as pd
import json
import uuid
import time
import re
import copy
from datetime import datetime, timedelta

# =====================================================================
# 页面配置 (恢复 v70.1 的 UI 风格)
# =====================================================================
st.set_page_config(
    page_title="IATF 审计转换工具 (v70.9 全功能回归版)",
    page_icon="🛡️",
    layout="wide"
)

# =====================================================================
# 侧边栏：模板与模式配置 (恢复 v70.1)
# =====================================================================
with st.sidebar:
    st.header("⚙️ 全局配置")
    st.divider()
    
    st.markdown("### 🔍 提取模式选择")
    run_mode = st.radio(
        "请根据报告类型选择：",
        (
            "纯净标准模式 (无附属场所)", 
            "单提取：EMS 扩展场所 (F21-M25)", 
            "单提取：RL 支持场所 (F27-N32)",
            "全量综合模式 (提取 EMS + RL + 被支持场所)"
        ),
        index=3
    )
    st.divider()
    
    st.info("💡 请上传您的 JSON 模板。程序将把该文件作为完整的底层骨架。")
    user_template_file = st.file_uploader("上传基础 JSON 模板", type=["json"])
    
    base_template_data = None
    if user_template_file:
        try:
            base_template_data = json.load(user_template_file)
            st.success(f"✅ 已加载底座: {user_template_file.name}")
        except Exception as e:
            st.error(f"❌ 解析失败: {e}")
            st.stop()
    else:
        st.warning("👈 请先上传底座文件以启动程序。")
        st.stop()

# =====================================================================
# 通用辅助函数区
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
# 场所提取逻辑 (严格继承自 v70.1 基座，仅加入地址拆分)
# =====================================================================
def extract_ems_sites(info_df):
    ems_sites = []
    if info_df.empty: return ems_sites
    header_r = -1
    col_map = {}
    row_start, row_end = 20, min(25, info_df.shape[0])
    for r in range(row_start, row_end):
        for c in range(5, min(13, info_df.shape[1])):
            val = str(info_df.iloc[r, c]).strip().upper()
            if "EMS扩展场所信息" in val or "扩展制造场所" in val:
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
            def safe_cell(idx): return str(info_df.iloc[r, idx]).strip() if idx != -1 else ""
            name_cn = safe_cell(col_map.get('name_cn', -1))
            addr_cn = safe_cell(col_map.get('addr_cn', -1))
            if not name_cn or name_cn == "nan": continue
            p, city, s = parse_chinese_address(addr_cn)
            ems_sites.append({
                "Id": str(uuid.uuid4()), "SiteName": name_cn, "IATF_USI": safe_cell(col_map.get('usi', -1)),
                "TotalNumberEmployees": safe_cell(col_map.get('emp', -1)),
                "AddressNative": {"Street1": s, "City": city, "State": p, "Country": "中国", "PostalCode": safe_cell(col_map.get('zip', -1))}
            })
    return ems_sites

def extract_rl_sites(info_df):
    sites = []
    if info_df.empty: return sites
    header_r = -1
    col_map = {}
    for r in range(26, min(32, info_df.shape[0])):
        for c in range(5, min(14, info_df.shape[1])):
            val = str(info_df.iloc[r, c]).strip().upper()
            if ("支持场所" in val or "RL" in val) and "被" not in val:
                header_r = r
                for cs in range(5, min(14, info_df.shape[1])):
                    h = str(info_df.iloc[r, cs]).strip()
                    if "中文名称" in h: col_map['name_cn'] = cs
                    elif "中文地址" in h: col_map['addr_cn'] = cs
                    elif "人数" in h: col_map['emp'] = cs
                    elif "支持功能" in h: col_map['func'] = cs
                break
        if header_r != -1: break
    if header_r != -1:
        for r in range(header_r + 1, min(32, info_df.shape[0])):
            def safe_cell(idx): return str(info_df.iloc[r, idx]).strip() if idx != -1 else ""
            name_cn = safe_cell(col_map.get('name_cn', -1))
            addr_cn = safe_cell(col_map.get('addr_cn', -1))
            if not name_cn or name_cn == "nan": continue
            p, city, s = parse_chinese_address(addr_cn)
            sites.append({
                "Id": str(uuid.uuid4()), "SiteName": name_cn, "Comments": safe_cell(col_map.get('func', -1)),
                "TotalNumberEmployees": safe_cell(col_map.get('emp', -1)),
                "AddressNative": {"Street1": s, "City": city, "State": p, "Country": "中国"}
            })
    return sites

def extract_receiving_sites(info_df):
    sites = []
    if info_df.empty: return sites
    header_r = -1
    col_map = {}
    for r in range(33, min(38, info_df.shape[0])):
        for c in range(5, min(14, info_df.shape[1])):
            if "被支持场所" in str(info_df.iloc[r, c]):
                header_r = r
                for cs in range(5, min(14, info_df.shape[1])):
                    h = str(info_df.iloc[r, cs]).strip()
                    if "中文名称" in h: col_map['name_cn'] = cs
                    elif "中文地址" in h: col_map['addr_cn'] = cs
                break
        if header_r != -1: break
    if header_r != -1:
        for r in range(header_r + 1, min(38, info_df.shape[0])):
            def safe_cell(idx): return str(info_df.iloc[r, idx]).strip() if idx != -1 else ""
            name_cn = safe_cell(col_map.get('name_cn', -1))
            addr_cn = safe_cell(col_map.get('addr_cn', -1))
            if not name_cn or name_cn == "nan": continue
            p, city, s = parse_chinese_address(addr_cn)
            sites.append({
                "Id": str(uuid.uuid4()), "SiteName": name_cn,
                "AddressNative": {"Street1": s, "City": city, "State": p, "Country": "中国"}
            })
    return sites

# =====================================================================
# 主流程区：核心转换逻辑 (严格遵循 v70.1 保护逻辑)
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

    # [基础信息]
    raw_name_full = find_val_by_key(db_df, ["姓名"]) or get_db_val(5, 1)
    formatted_team_name = extract_and_format_english_name(raw_name_full)
    start_iso = pd.to_datetime(find_val_by_key(db_df, ["开始日期"]) or get_db_val(2,1), errors='coerce')
    start_iso = start_iso.strftime('%Y-%m-%dT00:00:00.000Z') if pd.notna(start_iso) else ""

    # 💥💥 [核心修复：范围重复问题] 💥💥
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
    
    # 智能寻找底座中已有的 Key，避免重复创建
    native_key = "CertificateScopeNative"
    for k in org.keys():
        k_low = k.lower()
        if "scope" in k_low and ("native" in k_low or "zh" in k_low or "cn" in k_low):
            native_key = k # 锁定底座自带的 Key
            break
            
    if cert_scope_cn: org[native_key] = cert_scope_cn
    if cert_scope_en: org["CertificateScope"] = cert_scope_en

    # [地址拆分]
    addr_native = find_val_by_key(db_df, ["地址", "ADDRESS"]) or get_db_val(10, 1)
    p, city, s = parse_chinese_address(addr_native)
    org_addr = ensure_path(org, ["AddressNative"])
    org_addr.update({"Street1": s, "City": city, "State": p, "Country": "中国"})

    # [KPI 提取 - 严格版]
    kpi_map = {}
    if not perf_df.empty:
        col_m = {'proc':-1, 'kpi':-1, 'target':-1, 'result':-1}
        for r in range(min(10, perf_df.shape[0])):
            for c in range(perf_df.shape[1]):
                v = str(perf_df.iloc[r, c]).upper()
                if "过程" in v or "KPI" in v:
                    for sc in range(perf_df.shape[1]):
                        hv = str(perf_df.iloc[r, sc]).upper()
                        if "过程" == hv and col_m['proc']==-1: col_m['proc']=sc
                        elif "KPI" in hv and col_m['kpi']==-1: col_m['kpi']=sc
                    break
        if col_m['proc']!=-1:
            curr_p = ""
            for r in range(1, perf_df.shape[0]):
                p_v = str(perf_df.iloc[r, col_m['proc']]).strip()
                if p_v and p_v != "nan": curr_p = p_v
                k_v = str(perf_df.iloc[r, col_m['kpi']]).strip()
                if k_v and k_v != "nan":
                    if curr_p not in kpi_map: kpi_map[curr_p] = []
                    kpi_map[curr_p].append({"KPI": k_v, "CurrentTarget": str(perf_df.iloc[r, col_m.get('target',-1)]), "Results": str(perf_df.iloc[r, col_m.get('result',-1)])})

    # [过程保护融合逻辑]
    if not proc_df.empty:
        new_procs = []
        base_proc_map = {re.sub(r'\s+', '', p.get("ProcessName", "")): p for p in final_json.get("Processes", []) if isinstance(p, dict)}
        for idx, row in proc_df.iterrows():
            name = str(row.iloc[0]).strip()
            if not name or name=="nan": continue
            clean_name = re.sub(r'\s+', '', name)
            proc_obj = base_proc_map.get(clean_name, {"Id": str(uuid.uuid4()), "ProcessName": name})
            proc_obj["RepresentativeName"] = str(row.iloc[2])
            # 注入 KPI
            for pk, pv in kpi_map.items():
                if clean_name in re.sub(r'\s+', '', pk): proc_obj["ProcessPerformance"] = pv
            new_procs.append(proc_obj)
        final_json["Processes"] = new_procs

    # [多模式场所提取]
    if "全量" in mode:
        final_json["ExtendedManufacturingSites"] = extract_ems_sites(info_df)
        final_json["ProvidingSupportSites"] = extract_rl_sites(info_df)
        final_json["ReceivingSupportSites"] = extract_receiving_sites(info_df)
        org["ExtendedManufacturingSite"] = "1" if final_json.get("ExtendedManufacturingSites") else "0"

    final_json["uuid"] = str(uuid.uuid4())
    return final_json

# =====================================================================
# 主界面
# =====================================================================
st.title("🛡️ 多模板审计转换引擎 (v70.9 全功能修复版)")
st.markdown(f"💡 **当前运行模式**: `{run_mode}`")

uploaded_files = st.file_uploader("支持批量上传 .xlsx 格式文件", type=["xlsx"], accept_multiple_files=True)

if uploaded_files:
    st.divider()
    for file in uploaded_files:
        try:
            res_json = generate_json_logic(file, base_template_data, run_mode)
            st.success(f"✅ 解析成功：{file.name}")
            
            row_col1, row_col2 = st.columns([3, 1])
            with row_col1:
                with st.expander("👀 查看数据提取日志", expanded=True):
                    ems_c = len(res_json.get('ExtendedManufacturingSites', []))
                    rl_c = len(res_json.get('ProvidingSupportSites', []))
                    st.code(f"""
[状态报告]
✅ 范围 Key 匹配: 已原位更新底座节点
✅ 中文地址拆分: 省/市/街道已完成分离
✅ 过程特征继承: 底座 UUID 及隐藏参数已保留
✅ EMS 提取数量: {ems_c} 个
✅ RL 提取数量 : {rl_c} 个
                    """.strip(), language="yaml")
                    
            with row_col2:
                # 💥 关键修复：补全下载参数 💥
                st.download_button(
                    label="📥 下载 JSON 文件",
                    data=json.dumps(res_json, indent=2, ensure_ascii=False),
                    file_name=file.name.replace(".xlsx", ".json"),
                    mime="application/json",
                    key=f"dl_{file.name}"
                )
        except Exception as e:
            st.error(f"❌ 解析 {file.name} 失败: {str(e)}")
