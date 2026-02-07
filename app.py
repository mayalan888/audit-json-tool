import streamlit as st
import pandas as pd
import json
import uuid
import time
import copy
import re
from datetime import datetime, timedelta
from io import BytesIO

# --- 页面配置 ---
st.set_page_config(
    page_title="IATF 审计数据转换工具",
    page_icon="📊",
    layout="wide"
)

# --- 核心转换逻辑函数 ---
def generate_json_logic(excel_file, template_data):
    """
    核心处理逻辑 v8：
    包含：地址智能识别、日期自动计算、证书范围提取、文件清单填充。
    """
    # 深拷贝模板，防止修改原数据
    final_json = copy.deepcopy(template_data)
    
    # 1. 读取 Excel 所有必要的 Sheet
    # 注意：Streamlit 上传的文件是类似 BytesIO 的对象，可以直接被 pandas 读取
    try:
        xls = pd.ExcelFile(excel_file)
        
        # 读取数据库表 (Sheet1)
        if '数据库' in xls.sheet_names:
            db_df = pd.read_excel(xls, sheet_name='数据库', header=None)
        else:
            # 尝试读取第一个表作为数据库
            db_df = pd.read_excel(xls, sheet_name=0, header=None)

        # 读取过程清单 (Sheet2)
        if '过程清单' in xls.sheet_names:
            proc_df = pd.read_excel(xls, sheet_name='过程清单')
        else:
            proc_df = pd.DataFrame() # 空表防止报错

        # 读取信息表 (Sheet? - 用于地址兜底)
        if '信息' in xls.sheet_names:
            info_df = pd.read_excel(xls, sheet_name='信息', header=None)
        else:
            info_df = pd.DataFrame()

        # 读取文件清单 (通常是第9张表，或者名字叫'文件清单')
        doc_list_df = pd.DataFrame()
        if '文件清单' in xls.sheet_names:
            doc_list_df = pd.read_excel(xls, sheet_name='文件清单')
        elif len(xls.sheet_names) >= 9:
            # 尝试读取第9张表 (index 8)
            doc_list_df = pd.read_excel(xls, sheet_name=xls.sheet_names[8])
            
    except Exception as e:
        raise ValueError(f"Excel 读取失败: {str(e)}")

    # 2. 辅助函数：获取单元格值
    def get_db_val(row, col):
        try:
            val = db_df.iloc[row, col]
            return str(val).strip() if pd.notna(val) else ""
        except:
            return ""

    # 3. 提取基础信息
    report_name = get_db_val(1, 1)
    org_name = get_db_val(1, 4)
    start_date_raw = get_db_val(2, 1)
    cb_id = get_db_val(2, 4)
    end_date_raw = get_db_val(3, 1)
    usi_code = get_db_val(3, 4)
    caa_no = get_db_val(4, 1)

    # 4. 提取新增字段 (员工数、客户、CSR)
    total_employees = get_db_val(27, 1)
    customer_name = get_db_val(29, 1)
    supplier_code = get_db_val(30, 1)
    if supplier_code == "无": supplier_code = ""
    csr_doc_name = get_db_val(31, 1)
    csr_doc_date_raw = get_db_val(32, 1)

    # 格式化 CSR 日期
    csr_doc_date = str(csr_doc_date_raw)
    try:
        dt = pd.to_datetime(csr_doc_date_raw, errors='coerce')
        if not pd.isna(dt):
            csr_doc_date = f"{dt.year}/{dt.month}"
    except:
        pass

    # 5. 提取证书范围
    certificate_scope = ""
    for idx, row in db_df.iterrows():
        if str(row[0]).strip() == "证书范围":
            certificate_scope = str(row[1]).strip() if pd.notna(row[1]) else ""
            break

    # 6. 审核员姓名清洗
    raw_auditor_name = get_db_val(5, 1)
    auditor_name = raw_auditor_name
    english_part = re.sub(r'[\u4e00-\u9fff]', '', raw_auditor_name).strip()
    if english_part:
        parts = english_part.split()
        if len(parts) >= 2 and parts[0].isupper() and not parts[1].isupper():
            auditor_name = f"{parts[1]} {parts[0]}"
        else:
            auditor_name = english_part

    # 7. 智能地址识别 (中文/英文)
    candidates = []
    c1 = get_db_val(11, 1)
    c4 = get_db_val(11, 4)
    if c1: candidates.append(c1)
    if c4: candidates.append(c4)

    native_street = ""
    english_address = ""

    def is_chinese(s): return bool(re.search(r'[\u4e00-\u9fff]', s))

    # 取最长的中文作为 Native
    zh_candidates = [c for c in candidates if is_chinese(c)]
    if zh_candidates: native_street = max(zh_candidates, key=len)

    # 取最长的英文作为 Address
    en_candidates = [c for c in candidates if not is_chinese(c)]
    if en_candidates: english_address = max(en_candidates, key=len)

    # 兜底：如果没找到英文，去信息表找
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

    postal_code = get_db_val(10, 4)

    # 英文地址拆分
    street = english_address
    city = ""
    state = ""
    country = ""
    if english_address:
        parts = [p.strip() for p in english_address.split(',') if p.strip()]
        if len(parts) >= 3:
            country = parts[-1]
            state = parts[-2]
            city = parts[-3]
            street = ", ".join(parts[:-3])
        else:
            street = english_address

    # 构建 AddressNative 对象
    address_native_obj = {
        "Street1": native_street,
        "PostalCode": postal_code,
        "Country": "中国"
    }

    # 联系人
    representative = get_db_val(15, 1)
    telephone = get_db_val(15, 4)
    email = get_db_val(16, 1)
    if email == "0": email = ""

    # 8. 日期处理与计算
    def parse_to_datetime(d):
        try: return pd.to_datetime(d, errors='coerce')
        except: return pd.NaT

    def fmt_date_str(dt):
        if pd.isna(dt): return ""
        return dt.strftime('%Y-%m-%d') + "T00:00:00.000Z"

    start_dt = parse_to_datetime(start_date_raw)
    end_dt = parse_to_datetime(end_date_raw)

    audit_start_fmt = fmt_date_str(start_dt)
    audit_end_fmt = fmt_date_str(end_dt)
    dates_on_site = [{"Date": audit_start_fmt, "Day": 1}, {"Date": audit_end_fmt, "Day": 0.5}]

    # 9. 更新 JSON: AuditData
    if "AuditData" not in final_json: final_json["AuditData"] = {}
    final_json["AuditData"].update({
        "AuditDate": {"Start": audit_start_fmt, "End": audit_end_fmt},
        "ReportName": report_name,
        "CbIdentificationNo": cb_id,
        "AuditTeam": [{
            "Name": auditor_name,
            "CaaNo": caa_no,
            "AuditorId": "6-AUD-C-2410-1773-2260",
            "AuditDaysPerformed": 1.5,
            "DatesOnSite": dates_on_site,
            "PlanningTime": "0.0000"
        }]
    })

    # 10. 更新 JSON: OrganizationInformation
    if "OrganizationInformation" not in final_json: final_json["OrganizationInformation"] = {}
    final_json["OrganizationInformation"].update({
        "OrganizationName": org_name,
        "Address": {
            "Street1": street, "City": city, "State": state, "Country": country, "PostalCode": postal_code
        },
        "AddressNative": address_native_obj,
        "Representative": representative,
        "Telephone": telephone,
        "Email": email,
        "IATF_USI": usi_code,
        "CertificateScope": certificate_scope,
        "TotalNumberEmployees": total_employees
    })

    # 11. 更新 JSON: CustomerInformation
    if "CustomerInformation" not in final_json: final_json["CustomerInformation"] = {}
    new_customer = {
        "Id": str(int(time.time() * 1000) + 9999),
        "Name": customer_name,
        "SupplierCode": supplier_code,
        "Kpis": [],
        "Csrs": [{"Id": str(int(time.time() * 1000) + 8888), "NameCSRDocument": csr_doc_name, "DateCSRDocument": csr_doc_date}]
    }
    final_json["CustomerInformation"]["Customers"] = [new_customer]

    # 12. 更新文件清单 (Stage1DocumentedRequirements)
    if not doc_list_df.empty and "Stage1DocumentedRequirements" in final_json:
        target_col = None
        # 寻找包含特定关键词的列
        for col in doc_list_df.columns:
            if "公司内对应的程序文件" in str(col):
                target_col = col
                break
        
        if target_col:
            iatf_docs = final_json["Stage1DocumentedRequirements"].get("IatfClauseDocuments", [])
            for i, doc_name in enumerate(doc_list_df[target_col]):
                if i < len(iatf_docs):
                    val = str(doc_name).strip()
                    if val.lower() == "nan": val = ""
                    iatf_docs[i]["DocumentName"] = val

    # 13. 重构过程清单 (Processes)
    processes = []
    # 假设条款列从第13列开始 (index 13)
    clause_columns = proc_df.columns[13:] if proc_df.shape[1] > 13 else []

    for idx, row in proc_df.iterrows():
        rep_name = str(row.iloc[2]) if pd.notna(row.iloc[2]) else ""
        proc_name = str(row.iloc[12]) if pd.notna(row.iloc[12]) else ""
        if not proc_name or proc_name.lower() == "nan": continue

        audit_note_obj = {
            "Id": int(time.time() * 1000) + idx + 10000,
            "AuditorId": "6-AUD-C-2410-1773-2260"
        }

        proc_obj = {
            "Id": str(int(time.time() * 1000) + idx),
            "ProcessName": proc_name,
            "RepresentativeName": rep_name,
            "ProcessPerformance": [],
            "Activities": [],
            "CustomerCsrReference": [],
            "Shifts": [],
            "ExtendedShifts": [],
            "AuditNotes": [audit_note_obj],
            "ManufacturingProcess": "0",
            "OnSiteProcess": "1",
            "RemoteProcess": "0",
            "ProcessAuditTeam": [],
        }
        for col in clause_columns:
            val = row[col]
            is_checked = False
            if pd.notna(val):
                s_val = str(val).strip().upper()
                if s_val == "X" or s_val == "TRUE" or val is True:
                    is_checked = True
            if is_checked:
                proc_obj[col] = True
        processes.append(proc_obj)

    final_json["Processes"] = processes
    final_json["uuid"] = str(uuid.uuid4())
    final_json["created"] = int(time.time() * 1000)

    # 14. 更新结果日期 (Results)
    if "Results" not in final_json: final_json["Results"] = {}
    if "AuditReportFinal" not in final_json["Results"]: final_json["Results"]["AuditReportFinal"] = {}
    
    final_json["Results"]["AuditReportFinal"]["Date"] = audit_end_fmt

    if not pd.isna(end_dt):
        next_audit_dt = end_dt + timedelta(days=45)
        final_json["Results"]["DateNextScheduledAudit"] = fmt_date_str(next_audit_dt)
    else:
        final_json["Results"]["DateNextScheduledAudit"] = ""

    return final_json

# --- 网页界面 ---

st.title("🚀 审计数据 Excel 转 JSON 工具")
st.markdown("""
本工具会自动识别中英文地址、提取证书范围、自动计算审核周期，并填充文件清单。
""")
st.markdown("---")

# 1. 侧边栏：模板配置
with st.sidebar:
    st.header("⚙️ 设置")
    st.info("默认使用服务器上的json文件作为模板。如果您有新的模板结构，可以在此上传覆盖。")
    uploaded_template = st.file_uploader("上传自定义模板 JSON (可选)", type=["json"])

# 2. 加载模板
template_data = None
try:
    if uploaded_template:
        template_data = json.load(uploaded_template)
        st.sidebar.success("✅ 已加载自定义模板")
    else:
        with open('金磁.json', 'r', encoding='utf-8') as f:
            template_data = json.load(f)
        # st.sidebar.success("✅ 已加载默认模板 (金磁.json)")
except FileNotFoundError:
    st.error("❌ 严重错误：未找到默认模板 `金磁.json`。请确保该文件已上传到 GitHub 仓库根目录。")
    st.stop()
except Exception as e:
    st.error(f"❌ 模板加载失败: {e}")
    st.stop()

# 3. 主区域：文件上传
uploaded_files = st.file_uploader("📤 请上传 Excel 数据表 (支持多文件)", type=["xlsx"], accept_multiple_files=True)

# 4. 转换按钮与逻辑
if uploaded_files:
    if st.button("开始转换", type="primary"):
        st.markdown("---")
        st.subheader("📋 处理结果")
        
        for excel_file in uploaded_files:
            col1, col2 = st.columns([3, 1])
            
            with col1:
                st.write(f"正在处理: **{excel_file.name}**...")
            
            try:
                # 调用核心转换逻辑
                result_json = generate_json_logic(excel_file, template_data)
                
                # 生成文件名
                output_filename = excel_file.name.replace(".xlsx", ".json")
                if output_filename == excel_file.name:
                    output_filename += ".json"
                
                # 转换为 JSON 字符串用于下载
                json_str = json.dumps(result_json, indent=2, ensure_ascii=False)
                
                with col1:
                    st.success("转换成功！")
                    # 显示一些关键信息验证
                    emp_num = result_json.get('OrganizationInformation', {}).get('TotalNumberEmployees', 'N/A')
                    native_addr = result_json.get('OrganizationInformation', {}).get('AddressNative', {}).get('Street1', 'N/A')
                    st.caption(f"员工数: {emp_num} | 中文地址: {native_addr[:20]}...")

                with col2:
                    st.download_button(
                        label="📥 下载 JSON",
                        data=json_str,
                        file_name=output_filename,
                        mime="application/json"
                    )
            
            except Exception as e:
                with col1:
                    st.error(f"处理失败: {e}")
            
            st.divider()
else:
    st.info("👋 请在上方上传 Excel 文件开始工作。")

