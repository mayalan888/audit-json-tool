import streamlit as st
import pandas as pd
import json
import uuid
import time
import os
import copy
import re
from datetime import datetime, timedelta

st.set_page_config(page_title="IATF 审计转换工具 (v20.0)", page_icon="🎯", layout="wide")

# ================= UI：强制要求上传模板 =================
with st.sidebar:
    st.header("⚙️ 模板配置")
    st.warning("⚠️ 必须上传 JSON 模板才能进行转换。")
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
        st.stop()

# --- 辅助函数：安全更新深层节点 ---
def ensure_path(d, path):
    current = d
    for key in path:
        if key not in current or not isinstance(current[key], dict):
            current[key] = {}
        current = current[key]
    return current

# --- 核心数据转换逻辑 ---
def generate_json_logic(excel_file, template_data):
    final_json = copy.deepcopy(template_data)
    
    try:
        xls = pd.ExcelFile(excel_file)
        db_df = pd.read_excel(xls, sheet_name='数据库', header=None) if '数据库' in xls.sheet_names else pd.read_excel(xls, sheet_name=0, header=None)
        proc_df = pd.read_excel(xls, sheet_name='过程清单') if '过程清单' in xls.sheet_names else pd.DataFrame()
        info_df = pd.read_excel(xls, sheet_name='信息', header=None) if '信息' in xls.sheet_names else pd.DataFrame()
        
        doc_list_df = pd.DataFrame()
        if '文件清单' in xls.sheet_names:
            doc_list_df = pd.read_excel(xls, sheet_name='文件清单', header=None)
        elif len(xls.sheet_names) >= 9:
            doc_list_df = pd.read_excel(xls, sheet_name=xls.sheet_names[8], header=None)
    except Exception as e:




