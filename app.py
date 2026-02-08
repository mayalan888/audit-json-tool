import streamlit as st
import pandas as pd
import json
import uuid
import time
import copy
import re
from datetime import datetime, timedelta

# --- 页面配置 ---
st.set_page_config(
    page_title="IATF 审计数据转换工具 (智能修复版)",
    page_icon="🔍",
    layout="wide"
)

# --- 核心转换逻辑 ---
def generate_json_logic(excel_file, template_data):
    final_json = copy.deepcopy(template_data)
    
    try:
        xls = pd.ExcelFile(excel_file)
        
        # 1. 读取工作表 (增加容错)
        # 尝试读取 '数据库'，如果没有则读取第1张表
        if '数据库' in xls.sheet_names:
            db_df = pd.read_excel(xls, sheet_name='数据库', header=None)
        else:
            db_df = pd.read_excel(xls, sheet_name=0, header=None)

        # 尝试读取 '过程清单'
        proc_df = pd.DataFrame()
        if '过程清单' in xls.sheet_names:
            proc_df = pd.read_excel(xls, sheet_name='过程清单')

        # 尝试读取 '信息'
        info_df = pd.DataFrame()
        if '信息' in xls.sheet_names:
            info_df = pd.read_excel(xls, sheet_name='信息', header=None)
        
        # 尝试读取 '文件清单' (通常是第9张表)
        doc_list_df = pd.DataFrame()
        if '文件清单' in xls.sheet_names:
            doc_list_df = pd.read_excel(xls, sheet_name='文件清单')
        elif len(xls.sheet_names) >= 9:
            doc_list_df = pd.read_excel(xls, sheet_name=xls.sheet_names[8])
            
    except Exception as e:
        raise ValueError(f"Excel 读取失败: {str(e)}")

    # --- 智能搜索函数 ---
    def find_val_by_key(df, keywords, col_offset=1):
        """
        在 DataFrame 中搜索关键词，找到后返回其右侧(col_offset)单元格的值。
        """
        if df.empty: return ""
        # 遍历所有单元格
        for r in range(df.shape[0]):
            for c in range(df.shape[1]):
                cell_val = str(df.iloc[r, c]).strip()
                # 检查单元格是否包含任一关键词
                for k in keywords:
                    if k in cell_val:
                        # 找到关键词，尝试获取右侧单元格
                        if c + col_offset < df.shape[1]:
                            target_val = str(df.iloc[r, c + col_offset]).strip()
                            if target_val and target_val.lower() != 'nan':
                                return target_val
        return ""

    # --- 2. 提取数据 ---

    # (A) 姓名 Name
    # 优先找 "姓名"，如果找不到找 "Name"
    auditor_name = find_val_by_key(db_df, ["姓名", "Auditor Name"])
    # 如果名字里混入了英文 (比如 "Zhang San")，尝试把中文去掉，只留英文名(可选)
    # 这里保持原样或做简单清洗
    if auditor_name:
        # 去掉可能存在的 "Name:" 前缀
        auditor_name = auditor_name.replace("Name:", "").replace("姓名:", "").strip()

    # (B) CCAA 编号 (CaaNo)
    # 在数据库表中找 "审核员CCAA" 或 "CCAA"
    ccaa_raw = find_val_by_key(db_df, ["审核员CCAA", "CCAA-编号"])
    caa_no = ""
    if ccaa_raw:
        # 正则提取 CCAA: 后面的内容 (兼容中英文冒号)
        match = re.search(r'CCAA[:：]\s*([A-Za-z0-9-]+)', ccaa_raw)
        if match:
            caa_no = match.group(1)
        else:
            # 如果没有 CCAA: 前缀，可能整个单元格就是编号
            caa_no = ccaa_raw

    # (C) IATF 编号 (AuditorId)
    # 在信息表中找 "IATF Card"
    iatf_raw = find_val_by_key(info_df, ["IATF Card", "IATF卡号"])
    auditor_id = "" # 默认空
    if iatf_raw:
        # 正则提取 IATF: 后面的内容
        match = re.search(r'IATF[:：]\s*([0-9-]+)', iatf_raw)
        if match:
            auditor_id = match.group(1)
    
    # 如果信息表没找到，尝试在数据库表的 CCAA 那个格子里找 (有时候写在一起)
    if not auditor_id and "IATF" in ccaa_raw:
        match = re.search(r'IATF[:：-]?\s*([0-9-]+)', ccaa_raw)
        if match:
            auditor_id = match.group(1)
    
    # 兜底：如果还是没找到，给一个默认占位符，防止报错
    if not auditor_id:
        auditor_id = "UNKNOWN-ID"

    # (D) 其他基础信息 (尝试智能搜索，失败则回退到固定坐标)
    report_name = find_val_by_key(db_df, ["报告名称"]) or str(db_df.iloc[1, 1]) if db_df.shape[0]>1 else ""
    org_name = find_val_by_key(db_df, ["组织名称"]) or str(db_df.iloc[1, 4]) if db_df.shape[0]>1 else ""
    
    start_date_raw = find_val_by_key(db_df, ["审核开始时间"]) or str(db_df.iloc[2, 1]) if db_df.shape[0]>2 else ""
    end_date_raw = find_val_by_key(db_df, ["审核结束时间"]) or str(db_df.iloc[3, 1]) if db_df.shape[0]>3 else ""
    
    cb_id = find_val_by_key(db_df, ["认证机构识别号"]) or str(db_df.iloc[2, 4]) if db_df.shape[0]>2 else ""
    usi_code = find_val_by_key(db_df, ["IATF USI"]) or str(db_df.iloc[3, 4]) if db_df.shape[0]>3 else ""
    
    total_employees = find_val_by_key(db_df, ["员工总数"])
    certificate_scope = find_val_by_key(db_df, ["证书范围"])
    
    customer_name = find_val_by_key(db_df, ["客户名称"])
    supplier_code = find_val_by_key(db_df, ["供应商代码"])
    if supplier_code == "无": supplier_code = ""

    # (E) 地址处理
    # 搜索包含 "地址" 的行
    zh_addr = ""
    en_addr = ""
    
    # 简单的全表扫描找地址
    addr_candidates = []
    if not db_df.empty:
        for r in range(db_df.shape[0]):
            for c in range(db_df.shape[1]):
                val = str(db_df.iloc[r, c])
                if "地址" in val or "Address" in val:
                    # 收集这一行所有可能的地址文本
                    if c+1 < db_df.shape[1]: addr_candidates.append(str(db_df.iloc[r, c+1]))
                    if c+4 < db_df.shape[1]: addr_candidates.append(str(db_df.iloc[r, c+4]))

    # 区分中英文地址
    for cand in addr_candidates:
        if not cand or cand.lower() == 'nan': continue
        if re.search(r'[\u4e00-\u9fff]', cand):
            if len(cand) > len(zh_addr): zh_addr = cand
        else:
            if len(cand) > len(en_addr): en_addr = cand

    # (F) 日期格式化
    def fmt_iso(val):
        try:
            dt = pd.to_datetime(val, errors='coerce')
            if pd.notna(dt):
                return dt.strftime('%Y-%m-%d') + "T00:00:00.000Z"
        except: pass
        return ""

    start_iso = fmt_iso(start_date_raw)
    end_iso = fmt_iso(end_date_raw)

    # --- 3. 组装 JSON ---
    if "AuditData" not in final_json: final_json["AuditData"] = {}
    
    # 填充 AuditTeam
    final_json["AuditData"].update({
        "AuditDate": {"Start": start_iso, "End": end_iso},
        "ReportName": report_name,
        "CbIdentificationNo": cb_id,
        "AuditTeam": [{
            "Name": auditor_name,      # 修复：现在是真正的姓名
            "CaaNo": caa_no,           # 修复：正则提取后的纯数字
            "AuditorId": auditor_id,   # 修复：正则提取后的纯数字
            "AuditDaysPerformed": 1.5,
            "DatesOnSite": [{"Date": start_iso, "Day": 1}, {"Date": end_iso, "Day": 0.5}],
            "PlanningTime": "0.0000"
        }]
    })

    # Organization Info
    if "OrganizationInformation" not in final_json: final_json["OrganizationInformation"] = {}
    final_json["OrganizationInformation"].update({
        "OrganizationName": org_name,
        "AddressNative": {"Street1": zh_addr, "PostalCode": "", "Country": "中国"},
        "Address": {"Street1": en_addr, "City": "", "Country": "China"}, # 简易填充
        "IATF_USI": usi_code,
        "CertificateScope": certificate_scope,
        "TotalNumberEmployees": total_employees
    })

    # Customer Info
    if "CustomerInformation" not in final_json: final_json["CustomerInformation"] = {}
    final_json["CustomerInformation"]["Customers"] = [{
        "Id": str(int(time.time() * 1000)),
        "Name": customer_name,
        "SupplierCode": supplier_code,
        "Csrs": [] # 简化处理，如有 CSR 可在此添加逻辑
    }]

    # 4. 过程清单 (Processes)
    processes = []
    # 假设从第13列开始是条款 (Index 13)
    clause_cols = proc_df.columns[13:] if not proc_df.empty and proc_df.shape[1] > 13 else []
    
    if not proc_df.empty:
        for idx, row in proc_df.iterrows():
            # 过程名称通常在第13列 (Index 12)
            p_name = str(row.iloc[12]).strip() if proc_df.shape[1] > 12 else ""
            if not p_name or p_name.lower() == 'nan': continue
            
            proc_obj = {
                "Id": str(int(time.time() * 1000) + idx),
                "ProcessName": p_name,
                "AuditNotes": [{"Id": int(time.time()*1000)+idx+999, "AuditorId": auditor_id}],
                "ManufacturingProcess": "0", "OnSiteProcess": "1", "RemoteProcess": "0"
            }
            # 勾选条款
            for col in clause_cols:
                if str(row[col]).strip().upper() in ['X', 'TRUE']:
                    proc_obj[col] = True
            processes.append(proc_obj)

    final_json["Processes"] = processes
    final_json["uuid"] = str(uuid.uuid4())
    final_json["created"] = int(time.time() * 1000)

    # 结果日期
    if "Results" not in final_json: final_json["Results"] = {}
    final_json["Results"]["AuditReportFinal"] = {"Date": end_iso}
    
    return final_json

# --- 界面 ---
st.title("🔎 智能审计转换工具 v9.1")
st.caption("已启用智能搜索模式：自动定位姓名、CCAA编号及IATF卡号")

uploaded_files = st.file_uploader("请上传 Excel 文件", type=["xlsx"], accept_multiple_files=True)

if uploaded_files:
    # 尝试加载默认模板
    template = {}
    if os.path.exists('金磁.json'):
        with open('金磁.json', 'r', encoding='utf-8') as f:
            template = json.load(f)
    else:
        st.warning("⚠️ 未找到 '金磁.json' 模板，将使用空对象生成，可能导致部分字段缺失。")

    for file in uploaded_files:
        try:
            res_json = generate_json_logic(file, template)
            
            # 预览关键提取结果，方便核对
            team = res_json.get("AuditData", {}).get("AuditTeam", [{}])[0]
            st.success(f"✅ {file.name} 处理完成")
            st.markdown(f"""
            - **提取姓名**: `{team.get('Name')}`
            - **CCAA编号**: `{team.get('CaaNo')}`
            - **IATF卡号**: `{team.get('AuditorId')}`
            """)
            
            st.download_button(
                label=f"📥 下载 JSON ({file.name})",
                data=json.dumps(res_json, indent=2, ensure_ascii=False),
                file_name=file.name.replace(".xlsx", ".json")
            )
            st.divider()
            
        except Exception as e:
            st.error(f"❌ {file.name} 失败: {e}")



