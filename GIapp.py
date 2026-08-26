# ==========================================
# 非寿险精算对标报告系统 - 完整版
# 包含：登录认证、官网监控、智能页码定位、图表生成、AI分析、PDF导出
# ==========================================

import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import numpy as np
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import plotly.express as px
import openpyxl
import io
import re
import time
import json
import base64
from concurrent.futures import ThreadPoolExecutor, as_completed
import requests
from bs4 import BeautifulSoup
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill
from step8 import show_step_8_content

# ==========================================
# 0. 页面基础配置（必须在所有 st 命令之前）
# ==========================================
st.set_page_config(
    page_title="财险数智年报平台 | 对标报告系统",
    page_icon="Digi.png",
    layout="wide"
)

# ==========================================
# KPMG 官方色板（源自官方色卡）
# ==========================================
KPMG_COLORS = [
    "#00338D",  # 深蓝
    "#1E49E2",  # 亮蓝
    "#76D2FF",  # 淡蓝
    "#00B8F5",  # 青蓝
    "#ACEAFF",  # 极淡蓝
    "#510DBC",  # 深紫
    "#B497FF",  # 淡紫
    "#7213EA",  # 紫
    "#AB0D82",  # 玫红
    "#FD349C",  # 粉红
    "#FB8E7E",  # 橙红
    "#00C0AE",  # 青绿
]

# 保活脚本
components.html("""
<script>
    setInterval(() => {
        window.parent.document.dispatchEvent(new Event('mousemove'));
    }, 300000);
</script>
""", height=0, width=0)

# 左上角 Logo（已注释，可自行恢复）
# try:
#     st.logo("Digi.png", icon_image="Digi.png")
# except Exception as e:
#     st.error(f"Logo 加载失败，请检查文件名。错误: {e}")

# ==========================================
# 0.1 CSS 全量美化
# ==========================================
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Noto+Sans+SC:wght@300;400;500;600;700&display=swap');
    html, body, [class*="css"] { font-family: "PingFang SC", "HarmonyOS Sans", "Microsoft YaHei", "Noto Sans SC", sans-serif !important; color: #4A5568; }
    .stApp { background: linear-gradient(160deg, #F8FAFC 0%, #F1F5F9 50%, #E2E8F0 100%); }
    ::-webkit-scrollbar { width: 6px; height: 6px; }
    ::-webkit-scrollbar-track { background: transparent; }
    ::-webkit-scrollbar-thumb { background: #B0BEC5; border-radius: 3px; }
    ::-webkit-scrollbar-thumb:hover { background: #90A4AE; }
    ::selection { background: #B8C5D6; color: #1E293B; }
    [data-testid="stSidebar"] { background: linear-gradient(180deg, #E2E8F0 0%, #CBD5E1 100%) !important; border-right: 1px solid #94A3B8; }
    [data-testid="stSidebar"] .stTextInput label, [data-testid="stSidebar"] .stMarkdown { color: #334155 !important; font-weight: 500; }
    h1 { color: #111827 !important; font-weight: 600 !important; letter-spacing: 2px; border-bottom: none !important; padding-bottom: 10px; background: transparent !important; }
    h3 { color: #1E293B !important; font-size: 17px !important; font-weight: 600 !important; letter-spacing: 1px; background: rgba(255, 255, 255, 0.45) !important; backdrop-filter: blur(10px); -webkit-backdrop-filter: blur(10px); border: 1px solid rgba(255, 255, 255, 0.8); border-radius: 8px !important; padding: 8px 18px !important; box-shadow: 0 4px 12px rgba(0, 0, 0, 0.03), inset 0 1px 0 rgba(255, 255, 255, 0.5) !important; width: fit-content; margin-bottom: 18px !important; margin-top: 10px !important; }
    .stButton > button { font-size: 14px !important; font-weight: 600 !important; letter-spacing: 1.5px; background-color: #94A3B8; color: #FFFFFF; border: 1px solid #64748B; border-radius: 4px; padding: 8px 24px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1), 0 2px 4px -1px rgba(0,0,0,0.06); transition: all 0.2s ease; }
    .stButton > button:hover { background-color: #64748B; border-color: #475569; color: #F8FAFC; transform: translateY(-1px); box-shadow: 0 10px 15px -3px rgba(0,0,0,0.1); }
    .stButton > button:active { transform: translateY(1px); box-shadow: none; }
    .stButton > button:disabled { background-color: #CBD5E1 !important; color: #94A3B8 !important; border-color: #CBD5E1 !important; box-shadow: none !important; cursor: not-allowed; }
    .stTabs [data-baseweb="tab-list"] { gap: 4px; background-color: transparent; border-bottom: 2px solid #CBD5E1; }
    .stTabs [data-baseweb="tab"] { font-size: 14px; letter-spacing: 0.5px; background-color: #EDF2F7; color: #64748B; border-radius: 6px 6px 0 0; border: 1px solid #CBD5E1; border-bottom: none; padding: 10px 22px; transition: all 0.15s ease; }
    .stTabs [data-baseweb="tab"]:hover { background-color: #E2E8F0; color: #475569; }
    .stTabs [aria-selected="true"] { background-color: #94A3B8 !important; color: #FFFFFF !important; font-weight: 600 !important; border-color: #64748B !important; }
    .stTextInput > div > div > input, .stNumberInput > div > div > input { background-color: #FFFFFF !important; border: 1px solid #CBD5E1 !important; border-radius: 4px !important; color: #1E293B !important; font-weight: 500; transition: border-color 0.2s ease; }
    .stTextInput > div > div > input:focus, .stNumberInput > div > div > input:focus { border-color: #64748B !important; box-shadow: 0 0 0 2px rgba(100,116,139,0.15) !important; }
    .stMultiSelect [data-baseweb="tag"] { background-color: #CBD5E1 !important; color: #1E293B !important; border-radius: 4px !important; font-size: 13px !important; }
    [data-testid="stFileUploader"] section { border: 2px dashed #B0BEC5 !important; border-radius: 8px !important; background-color: #FAFBFC !important; transition: border-color 0.2s ease; }
    [data-testid="stFileUploader"] section:hover { border-color: #78909C !important; background-color: #F5F7FA !important; }
    hr { border: none !important; height: 1px !important; background: linear-gradient(to right, transparent, #B0BEC5, transparent) !important; margin: 20px 0 !important; }
    .info-card { background: #FFFFFF; border: 1px solid #E2E8F0; border-left: 4px solid #94A3B8; border-radius: 6px; padding: 20px 24px; margin: 12px 0; box-shadow: 0 1px 3px rgba(0,0,0,0.04); }
    .info-card h4 { color: #475569; margin: 0 0 8px 0; font-weight: 600; font-size: 15px; }
    .info-card p, .info-card li { color: #64748B; font-size: 14px; line-height: 1.8; }
    .info-card.pink { border-left-color: #D4A5A5; }
    .info-card.green { border-left-color: #A5C4B5; }
    .info-card.blue { border-left-color: #94A3B8; }
    .sidebar-brand { text-align: center; padding: 8px 0 16px 0; border-bottom: 1px solid #B0BEC5; margin-bottom: 20px; }
    .sidebar-brand .logo-text { font-size: 20px; font-weight: 700; color: #475569; letter-spacing: 3px; }
    .sidebar-brand .logo-sub { font-size: 11px; color: #94A3B8; letter-spacing: 4px; margin-top: 4px; }
    .placeholder-section { text-align: center; padding: 40px 20px; color: #94A3B8; }
    .placeholder-section .big-icon { font-size: 48px; margin-bottom: 16px; opacity: 0.6; }
    .placeholder-section p { font-size: 14px; line-height: 2; color: #64748B; }
</style>
""", unsafe_allow_html=True)

# ==========================================
# 0.2 登录认证 + Step0~Step1（财险适配版）
# ==========================================

# ---------- 1. 初始化系统状态变量 ----------
for k, v in {
    'logged_in': False,
    'user_role': None,
    'api_key': "",
    'base_url': "https://api.deepseek.com",
    'model_name': "deepseek-chat"
}.items():
    if k not in st.session_state:
        st.session_state[k] = v

# ---------- 2. 独立登录界面 ----------
if not st.session_state['logged_in']:
    
    st.markdown("""<style>
    header {visibility: hidden;}
    .stApp {background: linear-gradient(135deg, #040B16 0%, #0A1931 50%, #002266 100%); color: #E2E8F0;}
    [data-testid="column"]:nth-of-type(2) {background: rgba(255,255,255,0.08); backdrop-filter: blur(25px); -webkit-backdrop-filter: blur(25px); border-radius: 20px; border: 1px solid rgba(0,243,255,0.3); box-shadow: 0 15px 35px rgba(0,0,0,0.5), inset 0 0 15px rgba(0,243,255,0.15); padding: 40px 30px; margin-top: 6vh;}
    .title-glow {background: linear-gradient(90deg, #FFFFFF 0%, #76D2FF 50%, #00F3FF 100%); -webkit-background-clip: text; -webkit-text-fill-color: transparent; font-weight: 900; filter: drop-shadow(0 0 15px rgba(0,243,255,0.6));}
    div[data-baseweb="input"]>div, div[data-baseweb="select"]>div {background: rgba(0,0,0,0.4)!important; border: 1px solid rgba(0,243,255,0.3)!important; border-radius: 8px!important;}
    div[data-baseweb="input"]>div:focus-within, div[data-baseweb="select"]>div:focus-within {border-color: #00F3FF!important; box-shadow: 0 0 12px rgba(0,243,255,0.5)!important;}
    div[data-testid="stRadio"] label p {color: #FFFFFF!important; font-weight: bold!important; font-size: 14px!important;}
    label p, .stSelectbox label p, .stTextInput label p {color: #00F3FF!important; letter-spacing: 1px!important;}
    input, .stSelectbox span {color: #FFFFFF!important; font-size: 14px!important;}
    button[kind="primary"] {background: linear-gradient(90deg, #0044CC, #0088FF)!important; border: 1px solid rgba(0,243,255,0.5)!important; box-shadow: 0 0 15px rgba(0,136,255,0.4)!important; color: white!important; font-weight: bold!important; letter-spacing: 2px!important; border-radius: 8px!important; margin-top: 5px!important;}
    button[kind="primary"]:hover {box-shadow: 0 0 25px rgba(0,243,255,0.8)!important; transform: scale(1.02);}
    [data-testid="stPopover"] {display: flex; justify-content: flex-end;}
    [data-testid="stPopover"] button {background: transparent!important; border: none!important; box-shadow: none!important; color: #94A3B8!important; font-size: 12px!important; font-weight: normal!important; padding: 0!important; min-height: 0!important; text-decoration: underline; margin-top: 8px; margin-bottom: 15px;}
    [data-testid="stPopover"] button:hover {color: #00F3FF!important;}
    </style>""", unsafe_allow_html=True)

    _, col2, _ = st.columns([1, 1.5, 1])
    with col2:
        st.markdown("<div style='text-align:center; margin-bottom:30px;'><h1 style='font-size:32px; margin:0;'><span class='title-glow'>财险数智年报平台</span></h1><p style='color:#76D2FF; font-size:11px; letter-spacing:4px; margin-top:8px; font-weight:bold;'>P&C INTELLIGENCE</p></div>", unsafe_allow_html=True)
        
        u_type = st.radio("访问权限", ["普通用户", "项目组成员"], label_visibility="collapsed", horizontal=True)
        sec_code = st.text_input("安全验证", type="password", placeholder="请输入内部安全码")
        
        ai_map = {
            "阿里云百炼 (通义千问)": ("https://dashscope.aliyuncs.com/compatible-mode/v1", "qwen-plus"),
            "DeepSeek (深度求索)": ("https://api.deepseek.com", "deepseek-v4-flash"),
            "月之暗面 (Kimi)": ("https://api.moonshot.cn/v1", "moonshot-v1-8k"),
            "智谱AI (GLM-4)": ("https://open.bigmodel.cn/api/paas/v4", "glm-4"),
            "OpenAI (ChatGPT)": ("https://api.openai.com/v1", "gpt-4o")
        }
        ai_pr = st.selectbox("选择您将使用的AI", list(ai_map.keys()) + ["自定义私有化节点"])
        
        d_url, d_mod = ai_map.get(ai_pr, ("https://api.deepseek.com", "deepseek-chat"))
        if ai_pr == "自定义私有化节点":
            d_url, d_mod = st.text_input("RPC 接口地址", "https://api.deepseek.com"), st.text_input("指定模型版本", "deepseek-chat")

        api_input = st.text_input("请填写您使用AI的API Key", type="password", placeholder=" sk-... (不填则仅开放年报检测、数据合并及可视化本地分析功能)")
        
        if "OpenAI" in ai_pr:
            st.markdown("<p style='font-size:11px; color:#F87171; text-align:right; margin-top:4px; margin-bottom:0;'>⚠️ 严禁向境外节点传输涉密数据</p>", unsafe_allow_html=True)
        
        with st.popover("如何获取API key?"):
            st.markdown("**1.** 前往各大模型官方开放平台注册账号。\n\n**2.** 在控制台生成 `sk-` 开头的密钥。\n\n**3.** 复制填入上方输入框即可体验 AI 功能。")

        if st.button("启 动 系 统", type="primary", use_container_width=True):
            if u_type == "普通用户" and sec_code != "KPMG1234":
                st.error("❌ 拒绝访问：普通用户安全码错误")
            elif u_type == "项目组成员" and sec_code != "KPMG666":
                st.error("❌ 拒绝访问：项目组成员安全码错误")
            else:
                st.session_state.update({
                    'logged_in': True,
                    'user_role': u_type,
                    'api_key': api_input,
                    'base_url': d_url,
                    'model_name': d_mod
                })
                st.rerun()

        st.markdown("<div style='text-align:center; color:#94A3B8; font-size:11px; margin-top:30px; letter-spacing:1px;'>系统版本：v3.0 (Alpha) © 2026<br>Developed by 曾萍Polly@KPMG</div>", unsafe_allow_html=True)

    st.stop()

# ==========================================
# 🆕 在这里添加全局变量初始化（必须放在所有 def 函数之前）
# ==========================================
notes_dict = {}
ordered_modules = []

# ---------- 3. 导入 PDF 处理依赖 ----------
import fitz
import pdfplumber
from openai import OpenAI

# ---------- 4. 辅助函数：单位嗅探 ----------
def get_report_unit(text):
    """从文本中嗅探金额单位"""
    if not text:
        return None
    m1 = re.search(r'(?:金额)?单位[：:\s]*((?:人民币)?[元十百千万亿]+)', text)
    if m1:
        return m1.group(1).strip()
    m2 = re.search(r'(人民币[十百千万亿]+元)', text)
    if m2:
        return m2.group(1).strip()
    if "人民币元" in text:
        return "人民币元"
    return None
def normalize_field(s):
    if not isinstance(s, str):
        return ""
    # 去除常见空白字符
    s = s.replace('\xa0', '').replace('\u3000', '').replace(' ', '').strip()
    # 全角括号转半角
    s = s.replace('（', '(').replace('）', ')')
    return s
# ---------- 5. PDF 视觉提取引擎 ----------
def extract_single_page_vision(pdf_bytes, page_num, expected_name, api_key, base_url, model_name):
    """将 PDF 页面转为高清图片，使用视觉大模型提取表格"""
    try:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        if page_num < 1 or page_num > len(doc):
            return None, ""
        page = doc.load_page(page_num - 1)
        
        zoom_matrix = fitz.Matrix(3.0, 3.0)
        pix = page.get_pixmap(matrix=zoom_matrix, alpha=False)
        img_data = pix.tobytes("png")
        base64_image = base64.b64encode(img_data).decode('utf-8')
        doc.close()

        prompt = f"""你是一个四大会计师事务所的资深财险数字化审计专家。
我为你提供了一张【保险公司年报的单页高清截图】，目标是精准提取：【{expected_name}】。
表格中的【数字是由图片构成的】，请发挥强大的视觉识别能力，提取所有文字和表格。
【强制要求】：
1. 所有的科目名、数字金额之间，必须被 "|" 隔开。严禁使用连续空格。
2. 精准对齐多级表头！如果有留白单元格，必须用 "|" 占位补齐！
3. 绝不能漏掉任何一行数据，保留金额里的逗号和括号。
4. 纯文本输出，不要使用 Markdown 代码块。不需要输出 SHEET_NAME 标签，直接排版。"""

        client = OpenAI(api_key=api_key, base_url=base_url)
        response = client.chat.completions.create(
            model=model_name,
            messages=[
                {
                    "role": "user",
                    "content": [
                        {"type": "text", "text": prompt},
                        {"type": "image_url", "image_url": {"url": f"data:image/png;base64,{base64_image}"}}
                    ]
                }
            ],
            temperature=0.0
        )
        
        result_text = response.choices[0].message.content.strip()
        result_text = re.sub(r'^```(csv|text)?\n?', '', result_text, flags=re.MULTILINE).replace("```", "").strip()
        
        lines = result_text.split('\n')
        parsed_data = []
        max_cols = 0
        for row in lines:
            clean_row = row.strip()
            if not clean_row:
                continue
            if re.match(r'^[\s\|-]+$', clean_row) and '-' in clean_row:
                continue
            if clean_row.startswith('|'):
                clean_row = clean_row[1:]
            if clean_row.endswith('|'):
                clean_row = clean_row[:-1]
            cols = [col.strip() for col in clean_row.split('|')] if '|' in clean_row else [clean_row]
            parsed_data.append(cols)
            max_cols = max(max_cols, len(cols))
            
        for row in parsed_data:
            if len(row) < max_cols:
                row.extend([''] * (max_cols - len(row)))
                
        return pd.DataFrame(parsed_data), "【提示：该页已使用 OCR 视觉引擎处理，原文本为图片结构】"
        
    except Exception as e:
        return None, f"图像处理失败: {str(e)}"

# ---------- 6. PDF 文本提取引擎 ----------
def extract_single_page(pdf_bytes, page_num, expected_name, api_key, base_url, model_name):
    """使用 pdfplumber 提取文本，并用 LLM 转换为结构化数据"""
    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            if page_num < 1 or page_num > len(pdf.pages):
                return None, ""
            page = pdf.pages[page_num - 1]
            raw_text = page.extract_text(layout=True)
            
        if not raw_text:
            return None, ""

        prompt = f"""你是一个四大会计师事务所的资深财险数字化审计专家。
当前任务：提取【{expected_name}】。
请将以下纯文本中的【段落文字】原样保留为一行，将【财务表格】转化为用 "|" 严格分隔的标准网格格式。
【强制要求】：
1. 所有的科目名、附注编号、数字金额之间，必须被 "|" 隔开。严禁使用连续空格。
2. 精准对齐多级表头！如果某个单元格是空的，必须用 "|" 占位补齐！
3. 绝不能漏掉任何一行数据，保留金额里的逗号和括号。
4. 纯文本输出，不要使用 Markdown 代码块。不需要输出 SHEET_NAME，直接输出转化后的数据。

以下是原始文本，请开始输出：
{raw_text}"""

        client = OpenAI(api_key=api_key, base_url=base_url)
        response = client.chat.completions.create(
            model=model_name,
            messages=[{"role": "user", "content": prompt}],
            temperature=0.0
        )
        
        result_text = response.choices[0].message.content.strip()
        result_text = re.sub(r'^```(csv|text)?\n?', '', result_text, flags=re.MULTILINE)
        result_text = result_text.replace("```", "").strip()
        
        lines = result_text.split('\n')
        parsed_data = []
        max_cols = 0
        for row in lines:
            clean_row = row.strip()
            if not clean_row:
                continue
            if re.match(r'^[\s\|-]+$', clean_row) and '-' in clean_row:
                continue
            if clean_row.startswith('|'):
                clean_row = clean_row[1:]
            if clean_row.endswith('|'):
                clean_row = clean_row[:-1]
            if '|' not in clean_row:
                cols = [clean_row]
            else:
                cols = [col.strip() for col in clean_row.split('|')]
            parsed_data.append(cols)
            max_cols = max(max_cols, len(cols))
            
        for row in parsed_data:
            if len(row) < max_cols:
                row.extend([''] * (max_cols - len(row)))
            
        return pd.DataFrame(parsed_data), raw_text
        
    except Exception as e:
        return None, f"提取失败: {str(e)}"

# ---------- 7. AI 页码定位引擎（财险版） ----------
def ai_find_pages(pdf_bytes, api_key, target_tables, base_url, model_name, company_name=""):
    """混合双引擎：AI负责读目录推算主表，Python雷达负责深潜寻找附注表（财险版）"""
    try:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        
        # 1. 提取前 60 页（用于AI目录识别）
        toc_text = ""
        for i in range(min(60, len(doc))):
            page = doc.load_page(i)
            page_text = page.get_text("text").replace("\n", " ")
            page_text = " ".join(page_text.split())
            toc_text += f"---第{i+1}页---\n{page_text[:1200]}\n"
            
        # ==========================================
        # 财险专属附注雷达特征矩阵（其他表保持不变）
        # ==========================================
        feature_matrix = {
            "保险业务收入表 (附注)": [
                ["保险业务收入", "已赚保费", "签单保费"],
                ["原保险", "再保险", "分入"],
                ["分保费收入", "分保费用", "保费收入"]
            ],
            "综合成本率拆解表 (附注)": [
                ["综合成本率", "COR", "成本率"],
                ["赔付率", "费用率", "综合赔付率", "综合费用率"],
                ["已发生赔款", "手续费及佣金", "业务及管理费"]
            ],
            "准备金评估表 (附注)": [
                ["未到期责任准备金", "未决赔款准备金"],
                ["IBNR", "已发生未报告"],
                ["准备金", "负债", "准备金余额"]
            ],
            "费用结构表 (附注)": [
                ["业务及管理费", "手续费及佣金"],
                ["职工薪酬", "折旧", "租赁", "办公费", "宣传费"],
                ["获取费用", "维持费用"]
            ],
            "投资收益表 (附注)": [
                ["投资收益", "投资资产", "净投资回报"],
                ["公允价值变动", "利息收入", "股息收入"],
                ["债券", "股票", "基金", "长期股权投资"]
            ],
            "偿付能力表 (附注)": [
                ["偿付能力", "综合偿付能力充足率", "核心偿付能力充足率"],
                ["实际资本", "最低资本"],
                ["风险", "资本"]
            ],
            "业务及管理费 (附注)": [
                ["业务及管理费", "管理费用"],
                ["职工薪酬", "折旧", "租赁", "办公费", "宣传费"]
            ],
            "主要经营指标": [
                ["偿付能力充足率", "核心偿付能力", "综合偿付能力"],
                ["实际资本", "最低资本", "风险资本"],
                ["保险业务收入", "净利润", "净资产"]
            ]
        }
        # ==========================================
        # 保险合同负债及资产 专属识别规则
        # ==========================================
        insurance_rules = {
            "平安": {
                "title_words": ["保险合同负债/资产", "保险合同负债／资产", "保险合同负债及资产"],
                "title_require": [],
                "include": ["保险合同负债", "保险合同资产", "合同服务边际"],
                "exclude": ["再保险合同", "持有的再保险", "分出的再保险"],
                "require_any": []
            },
            "阳光": {
                "title_words": ["保险合同负债及资产"],
                "title_require": [],
                "include": ["保险合同负债", "保险合同资产", "合同服务边际"],
                "exclude": ["再保险合同", "持有的再保险"],
                "require_any": []
            },
            "众安": {
                "title_words": ["保险合同资产和负债", "保险合同负债和资产", "保险合同资产及负债", "保险合同负债及资产"],
                "title_require": [],
                "include": ["签发的保险合同", "保费分配法"],
                "exclude": ["持有的再保险合同", "再保险合同"],
                "require_any": ["保费分配法"]
            },
            "人保": {
                "title_words": ["保险合同"],
                "title_require": [
                    "本公司签发的保险合同采用保费分配法计量",
                    "本集团和本公司签发的保险合同采用通用模型计量"
                ],
                "include": [
                    "本公司签发的保险合同采用保费分配法计量",
                    "本集团和本公司签发的保险合同采用通用模型计量",
                    "通用模型计量"
                ],
                "exclude": [],
                "require_any": []
            },
            "太保": {
                "title_words": ["保险合同负债/资产", "保险合同负债／资产", "保险合同负债及资产"],
                "title_require": [],
                "include": ["保险合同负债", "保险合同资产", "合同服务边际"],
                "exclude": ["再保险合同", "持有的再保险"],
                "require_any": []
            },
            "太平": {
                "title_words": ["保险合同资产及负债", "保险合同负债及资产"],
                "title_require": [],
                "include": ["保险合同资产及负债", "合同服务边际", "保险合同负债"],
                "exclude": ["再保险合同", "持有的再保险"],
                "require_any": []
            }
        }

        # 未知/自动识别公司时的兜底规则
        default_insurance_rule = {
            "title_words": [
                "保险合同负债及资产", "保险合同资产及负债", "保险合同负债/资产",
                "保险合同负债／资产", "保险合同资产和负债", "保险合同负债和资产"
            ],
            "title_require": [],
            "include": ["保险合同负债", "保险合同资产", "合同服务边际"],
            "exclude": ["再保险合同", "持有的再保险"],
            "require_any": []
        }
        
        simple_title_OR = {
            "利润表 (附注)": ["利润表", "损益表", "利润总额"],
            "资产负债表 (附注)": ["资产负债表", "资产", "负债", "所有者权益"],
            "现金流量表 (附注)": ["现金流量表", "现金流", "经营活动"],
            "股东/所有者权益变动表 (附注)": ["股东权益变动表", "权益变动", "所有者权益变动表"],
            "公司利润表": ["利润表", "损益表", "利润总额", "公司利润表"],
            "公司资产负债表": ["资产负债表", "公司资产负债表"],
            "业务及管理费": ["业务及管理费", "管理费用"],
            "主要经营指标": ["主要经营指标", "经营指标", "偿付能力指标"]
        }
        
        context_anchors = ["项目注释", "财务报表附注", "报表附注", "项目附注"]
        
        found_hints = {table: [] for table in target_tables if table in feature_matrix or table in simple_title_OR or table == "保险合同负债及资产" or table == "保险产品经营信息"}

        # ====== 新增：初始化 page_type_map ======
        page_type_map = {}

        # ==========================================================
        # 从第0页开始扫描
        # ==========================================================
        for i in range(0, len(doc)):
            page = doc.load_page(i)
            text_raw = page.get_text("text")
            text_clean = text_raw.replace("\n", "").replace(" ", "")
            header_text = text_clean[:800]
            
            # 检测表格
            try:
                tables = page.find_tables()
                has_actual_table = len(tables.tables) > 0 if tables else False
            except:
                has_actual_table = False
            is_continued_table = "(续)" in text_clean or "（续）" in text_clean
            is_table_page = is_continued_table or has_actual_table
            
            is_notes_page = any(anchor in text_clean for anchor in context_anchors)
            
            # ---------- 1. 匹配 feature_matrix（其他表） ----------
            if is_notes_page and is_table_page:
                for table, condition_groups in feature_matrix.items():
                    if table in target_tables:
                        matched = all(any(kw in text_clean for kw in group) for group in condition_groups)
                        if matched:
                            page_num = i + 1
                            if page_num not in found_hints[table]:
                                found_hints[table].append(page_num)
            
            # ---------- 2. 匹配 simple_title_OR（其他表） ----------
            for table, any_of_words in simple_title_OR.items():
                if table in target_tables:
                    if any(kw in header_text for kw in any_of_words):
                        page_num = i + 1
                        if "合并" in header_text:
                            page_type_map[page_num] = 'group'
                        else:
                            page_type_map[page_num] = 'company'
                        if page_num not in found_hints[table]:
                            found_hints[table].append(page_num)
            
            # ---------- 3. 特殊处理“保险合同负债及资产” ----------
            
            if "保险合同负债及资产" in target_tables:
                # 3.0 公司归一化匹配（支持"平安产险"/"太保财险"等全称命中短键）
                company_key = ""
                if company_name:
                    for _key in insurance_rules:
                        if _key in company_name:
                            company_key = _key
                            break
                rule = insurance_rules.get(company_key, default_insurance_rule)

                # 3.1 数据页判断（动态年份，支持2023等历史报告；续表页豁免年份要求）
                if not is_continued_table and not re.search(r'20\d{2}', text_clean):
                    continue

                # 3.2 排除资产负债表误命中（需"资产负债表"标题 或 ≥2个强信号；
                #     注意："负债合计"是"保险合同负债合计"的子串，单独出现不能判为资产负债表）
                if "资产负债表" in text_clean:
                    continue
                bs_strong = ["资产总计", "所有者权益合计", "流动资产", "非流动资产", "流动负债", "非流动负债"]
                if sum(1 for w in bs_strong if w in text_clean) >= 2:
                    continue

                # 3.3 表名称匹配（按公司差异化表名）
                score = 0   
                title_words = rule.get("title_words", default_insurance_rule["title_words"])
                title_match = any(word in text_clean for word in title_words)
                # 表名较泛的公司（如人保仅"保险合同"），额外要求命中具体子表头
                title_require = rule.get("title_require", [])
                if title_match and title_require:
                    if not any(word in text_clean for word in title_require):
                        title_match = False
                if not title_match:
                    continue
                score += 10

                # 3.4 硬性必含字段（如众安仅取"保费分配法"部分）
                require_any = rule.get("require_any", [])
                if require_any:
                    if not any(word in text_clean for word in require_any):
                        continue

                # 3.5 IFRS17核心字段加分（新增"通用模型"）
                indicators = ["合同服务边际","未来现金流量","非金融风险调整","履约现金流",
                              "亏损部分","非亏损部分","未到期责任负债","已发生赔款负债",
                              "保费分配法","通用模型"]
                for word in indicators:
                    if word in text_clean:
                        score += 1

                # 3.6 公司差异规则（用归一化后的 rule，不再用 company_name 直接取）
                for word in rule.get("include", []):
                    if word in text_clean:
                        score += 3
                for word in rule.get("exclude", []):
                    if word in text_clean:
                        score -= 8

                # 3.7 续表识别 / 3.8 最终判断（不变）
                if "(续)" in text_clean or "（续）" in text_clean:
                    score += 3
                if score >= 10:
                    page_num = i + 1
                    if page_num not in found_hints["保险合同负债及资产"]:
                        found_hints["保险合同负债及资产"].append(page_num)

            # ---------- 4. 特殊处理“保险产品经营信息” ----------
            if "保险产品经营信息" in target_tables:
                # 使用更精准的标题关键词（只保留最独特的标识）
                if any(kw in header_text for kw in ["保险产品经营信息", "原保险保费收入居前5位的险种"]):
                    # 必须包含实际表格（find_tables 检测到表格，或为续表）
                    if has_actual_table or is_continued_table:
                        # 排除利润表标题（避免误匹配）
                        if not any(kw in header_text for kw in ["利润表", "损益表", "利润总额", "公司利润表"]):
                            page_num = i + 1
                            if page_num not in found_hints.get("保险产品经营信息", []):
                                found_hints.setdefault("保险产品经营信息", []).append(page_num)

        # ==========================================================
        # 公司口径过滤
        # ==========================================================
        for table in ["公司利润表", "公司资产负债表"]:
            if table in found_hints and found_hints[table]:
                company_pages = [p for p in found_hints[table] if page_type_map.get(p) == 'company']
                # 如果没有公司口径，则保留所有（包括合并）
                found_hints[table] = company_pages if company_pages else found_hints[table]
                if company_pages:
                    found_hints[table] = company_pages

        doc.close()
        
        # 构建提示文本（供AI参考）
        hint_text = "\n\n【⚡ Python程序附注雷达线索】\n"
        for table, pages in found_hints.items():
            if pages:
                pages = sorted(list(set(pages)))
                hint_text += f"- {table} 真实表格物理页码位于: {pages[:4]}\n"

        # AI 目录驱动提示（保持原样）
        prompt = f"""你是一个顶级的财险审计专家，任务是定位 PDF 中的核心财务报表页码。

【关键定位逻辑（双模式自适应）】
模式 A：目录驱动（适用于有目录的报表）
- 如果前几页存在"目录"或"Contents"，请识别目标报表对应的页码。
- 计算物理偏移量（封面/说明页占用的页数），得出真实的物理页码。

模式 B：标题驱动（适用于无目录或目录不全的报表）
- 如果没有明确目录，请直接扫描文本中 `---第N页---` 标记下方的页眉或正文大标题。
- 例如：若在 `---第8页---` 下方看到"合并利润表"，则物理页码即为 8。

【通用执行准则】
1. 🎯 如果需求表单名称包含"（合并）"，优先寻找：
   合并资产负债表
   合并利润表
   合并现金流量表
   合并股东权益变动表

2. 🎯 如果需求表单名称包含"（公司）"，优先寻找：
   公司资产负债表
   公司利润表
   公司现金流量表
   公司股东权益变动表

3. 🎯 如果年报中没有区分"合并"和"公司"，仅出现：
   资产负债表
   利润表
   现金流量表
   股东权益变动表
   则同一页码同时返回给对应的（合并）和（公司）

4. 🎯 跨页逻辑：由于财务报表通常包含"(续)"，请务必返回所有相关的物理页码数组。

5. 🎯 附注表线索（最高优先级）：对于带有"(附注)"字样的复杂底表，请【无脑完全信任】下方提供的【Python程序附注雷达线索】，直接返回线索中的页码。

需求表单列表：
{json.dumps(target_tables, ensure_ascii=False)}

【待扫描文本内容（前 60 页）】：
{toc_text}

【⚡ Python程序附注雷达辅助线索】：
{hint_text}

【强制输出格式】
只输出合法 JSON，键为表单名，值为物理页码数组。找不到填 [0]。
格式示例：{{"合并利润表": [8, 9], "资产负债表 (合并优先)": [2, 3, 4]}}"""

        client = OpenAI(api_key=api_key, base_url=base_url)
        response = client.chat.completions.create(
            model=model_name,
            messages=[{"role": "user", "content": prompt}],
            temperature=0.1
        )
        
        result_text = response.choices[0].message.content.strip()
        result_text = result_text.replace("```json", "").replace("```", "").strip()
        return json.loads(result_text)
        
    except Exception as e:
        return {"error": str(e)}
# ==========================================
# 以下为原有财险对标报告系统（第1~10段）—— 所有函数定义移至主逻辑之前
# ==========================================
# 0.添加公司边框和标题
def add_company_borders(fig, companies, x_labels, top_margin=60, row=None, col=None):
    """
    为单图或子图添加公司边框和标题。
    如果 row 和 col 不为空，则视为子图模式，直接覆盖整个子图区域。
    """
    # ===== 子图模式 =====
    if row is not None and col is not None:
        if len(companies) == 1:
            co = companies[0]
            # 使用 x domain / y domain 限定在子图内部
            fig.add_shape(
                type="rect",
                xref="x domain", yref="y domain",
                x0=0, x1=1,
                y0=0, y1=1.0,          # 覆盖整个子图区域（从底部到顶部）
                fillcolor="rgba(200,200,200,0.05)",
                line=dict(color="#CCCCCC", width=1.2),
                layer="below",
                row=row, col=col
            )
            # 公司名称放在子图顶部内部（y=0.98）
            fig.add_annotation(
                x=0.5,
                y=0.95,
                text=f"<b>{co}</b>",
                showarrow=False,
                font=dict(size=12, color="#888888"),
                xanchor="center",
                yanchor="bottom",
                xref="x domain",
                yref="y domain",
                row=row, col=col
            )
        return fig
        
    # ===== 单图模式（原有逻辑） =====
    company_ranges = {}
    for co in companies:
        indices = [i for i, label in enumerate(x_labels) if co in label]
        if indices:
            company_ranges[co] = (min(indices), max(indices))
    if not company_ranges:
        for co in companies:
            if co in x_labels:
                idx = x_labels.index(co)
                company_ranges[co] = (idx, idx)
    if not company_ranges:
        return fig

    for co, (start, end) in company_ranges.items():
        width_extend = 0.55 if (end - start) > 0 else 0.45
        x0 = start - width_extend
        x1 = end + width_extend
        fig.add_shape(
            type="rect",
            xref="x", yref="paper",
            x0=x0, x1=x1,
            y0=0, y1=1.03,
            fillcolor="rgba(200,200,200,0.05)",
            line=dict(color="#CCCCCC", width=1.2),
            layer="below"
        )
        fig.add_annotation(
            x=(start + end) / 2,
            y=0.985,
            text=f"<b>{co}</b>",
            showarrow=False,
            font=dict(size=12, color="#888888"),
            xanchor="center",
            yanchor="bottom",
            xref="x",
            yref="paper"
        )
    fig.update_layout(margin=dict(t=top_margin))
    return fig
    
# 1.全局颜色工具
def get_color_map(all_cos):
    current_selection_key = tuple(sorted(all_cos))
    if 'company_color_map' not in st.session_state or st.session_state.get('_last_color_selection') != current_selection_key:
        PRESET_COLORS = ["#C00000", "#0865EE", "#FEAED7", "#92D050", "#7030A0", "#EF9867", "#61CBF4", "#C7A0F7"]
        st.session_state['company_color_map'] = {co: PRESET_COLORS[i % len(PRESET_COLORS)] for i, co in enumerate(all_cos)}
        st.session_state['_last_color_selection'] = current_selection_key
    return st.session_state['company_color_map']

# 2.辅助显示函数（AI引擎、笔记、图表渲染等）
@st.cache_data(show_spinner=False)
def _call_llm_api_cached(prompt, field_name, latest_year, unit_str, user_api_key, api_base, api_model):
    try:
        client = OpenAI(api_key=user_api_key, base_url=api_base)
        res = client.chat.completions.create(
            model=api_model,
            messages=[{"role": "user", "content": prompt}],
            temperature=0.3
        )
        return f"<b><span style='color:#00338D'></span></b> {res.choices[0].message.content}"
    except Exception as e:
        return f"<span style='color:#C00000; font-size:12px;'>⚠️ AI报错: {str(e)}</span>"

def generate_ai_insight(df, field_name, is_pct=False, template=""):
    """
    生成 AI 分析文本
    - template: 参考话术模板（来自“分析内容-默认”列）
    """
    enable_ai = st.session_state.get('enable_ai', False)
    latest_year = st.session_state.get('latest_year', 2025)
    selected_cos = st.session_state.get('selected_cos_cache', [])
    divisor = st.session_state.get('divisor', 1)
    unit_label = st.session_state.get('unit_label', '百万元')

    if not enable_ai or df is None or df.empty or not field_name:
        return ""

    try:
        col_field = '字段名' if '字段名' in df.columns else '指标名称'
        # 提取最新年份数据
        d_sub = df[(df[col_field].astype(str).str.contains(field_name, na=False)) &
                   (df['报告年份'].astype(str) == str(latest_year)) &
                   (df['公司'].isin(selected_cos))].copy()
        if d_sub.empty:
            return ""

        val_col = "(百万)人民币" if "(百万)人民币" in d_sub.columns else d_sub.columns[-1]
        # 上年数据
        prev_year = st.session_state.get('prev_year', 2024)
        d_prev = df[(df[col_field].astype(str).str.contains(field_name, na=False)) &
                    (df['报告年份'].astype(str) == str(prev_year)) &
                    (df['公司'].isin(selected_cos))].copy()

        # ---- 构建数据事实 ----
        data_dict = {}
        for _, r in d_sub.iterrows():
            co = r['公司']
            val = r[val_col]
            data_dict[co] = val

        # 格式化数据
        data_summary = []
        for co, val in data_dict.items():
            if is_pct:
                data_summary.append(f"{co}: {val:.1%}")
            else:
                data_summary.append(f"{co}: {val/divisor:.1f}{unit_label}")

        # 排名和变化
        sorted_items = sorted(data_dict.items(), key=lambda x: x[1], reverse=True)
        top_co, top_val = sorted_items[0]
        bottom_co, bottom_val = sorted_items[-1]

        changes = []
        for co in selected_cos:
            curr_val = data_dict.get(co)
            prev_vals = d_prev[d_prev['公司'] == co][val_col]
            prev_val = prev_vals.iloc[0] if not prev_vals.empty else None
            if curr_val is not None and prev_val is not None and prev_val != 0:
                pct_change = (curr_val - prev_val) / abs(prev_val)
                changes.append((co, pct_change))
        changes.sort(key=lambda x: x[1], reverse=True)
        top_gain = changes[0] if changes else None
        top_loss = changes[-1] if changes else None

        # 构建事实文本
        facts = f"最新年（{latest_year}）数据：{', '.join(data_summary)}。"
        facts += f"最高为{top_co}（{top_val:.2f}），最低为{bottom_co}（{bottom_val:.2f}）。"
        if top_gain:
            facts += f"同比增幅最大为{top_gain[0]}（{top_gain[1]:.1%}），降幅最大为{top_loss[0]}（{top_loss[1]:.1%}）。"

        # ---- 构建提示词 ----
        if template and template.strip():
            prompt = f"""
你是一位精算分析专家。请根据以下数据事实，参考给定的“话术模板”风格，生成一段简洁通顺的分析文字（不超过80字）。
数据事实：{facts}
话术模板：{template}
要求：仿照模板的表述方式，但必须基于数据事实，不要出现模板中的固定公司名或无关内容。
输出直接是分析文字，不要前缀。
"""
        else:
            prompt = f"你是一位精算分析专家。请根据以下数据事实，用简洁通顺的中文总结一段分析（不超过60字）：{facts}"

        # ---- 调用AI ----
        api_key = st.session_state.get('api_key', "").strip()
        if api_key:
            return _call_llm_api_cached(prompt, field_name, latest_year, unit_label, api_key, st.session_state.get('base_url'), st.session_state.get('model_name'))
        else:
            # 无API时返回数据事实
            return f"<b>📊 {facts}</b>"

    except Exception as e:
        return f"<span style='color:#C00000;'>⚠️ AI分析出错: {e}</span>"

def display_notes(module_id, ai_df=None, ai_field=None, is_pct=False):
    global notes_dict
    md = notes_dict.get(module_id, {})
    an_default = md.get('analysis_default', "")   # 话术模板
    an_custom  = md.get('analysis_custom', "")
    nt         = md.get('note', "")
    ai_txt     = generate_ai_insight(ai_df, ai_field, is_pct, template=an_default)

    # 如果启用AI，优先使用AI生成的文本，否则使用预设
    if ai_txt and st.session_state.get('enable_ai', False):
        display_text = ai_txt
    else:
        display_text = an_default
        if an_custom:
            display_text += "<br>" + an_custom if display_text else an_custom
        if ai_txt and not st.session_state.get('enable_ai', False):
            display_text += "<br>" + ai_txt if display_text else ai_txt

    if display_text:
        html = '<div style="text-align:left; background:#F0F4FA; border-left:4px solid #00338D; padding:3px 10px; margin-bottom:6px; border-radius:4px; font-family:Microsoft YaHei, 微软雅黑, sans-serif;">'
        html += f'<p style="margin:2px 0; color:#0A1F5C; font-size:14px; line-height:1.4;">{display_text}</p>'
        st.markdown(html + "</div>", unsafe_allow_html=True)

    return (an_default or an_custom), nt

def display_bottom_note(nt_text):
    if nt_text and str(nt_text).lower() != 'nan':
        st.markdown(f'<div style="text-align: left;margin-top: 1px; margin-bottom: 0px; padding-left: 5px;text-align: left;"><p style="margin: 0; color: #888; font-size: 12px; font-style: italic;">* 注释：{nt_text}</p></div>', unsafe_allow_html=True)

def show_chart(fig, p_mode, m_id=None):
    if not fig:
        return
    if p_mode:
        # 打印模式：自动适应宽度，通过 flex 居中
        fig.update_layout(
            autosize=True,
            height=500,
            margin=dict(t=30, b=30, l=50, r=20)
        )
        st.markdown('<div style="display: flex; justify-content: center; width: 100%;">', unsafe_allow_html=True)
        st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})
        st.markdown('</div>', unsafe_allow_html=True)
    else:
        # 非打印模式保持不变
        if m_id in ["claim_ratio_trend", "expense_ratio_trend"]:
            fig.update_layout(
                autosize=False,
                width=900,
                height=550,
                margin=dict(t=50, b=100, l=60, r=40)
            )
            st.plotly_chart(fig, use_container_width=False, config={"displayModeBar": False})
        else:
            fig.update_layout(autosize=True)
            st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})
            
# 3.文本披露卡片
def display_textual_disclosures(df, cos, cy):
    st.markdown("#### 📄 关键会计政策与精算假设披露")
    for co in cos:
        df_co = df[df['公司'] == co]
        def get_text(y, kw):
            s = df_co[(df_co['报告年份'].astype(str) == str(y)) & 
                      (df_co['字段名'].astype(str).str.contains(kw, na=False))]['值']
            if s.empty:
                return "未披露"
            val = s.iloc[0]
            return str(val) if pd.notna(val) and str(val).strip() != "" else "未披露"
        discount_rate = get_text(cy, '折现率')
        risk_margin = get_text(cy, '风险边际') or get_text(cy, '非金融风险调整')
        policy_method = get_text(cy, '计量方法') or get_text(cy, '会计政策')
        if all(v == "未披露" for v in [discount_rate, risk_margin, policy_method]):
            continue
        with st.expander(f"📌 {co} - 精算假设与政策", expanded=False):
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("计量方法（政策）", policy_method)
            with col2:
                st.metric("折现率假设", discount_rate)
            with col3:
                st.metric("风险边际（RA）", risk_margin)


# 4.综合成本率拆解（多因子分组柱状图）
def create_cor_breakdown_chart(df, cos, year, divisor, highlight_co="无"):
    factors = [
        "当期发生赔款及理赔费用",
        "已发生赔款负债履约现金流变动",
        "亏损合同损益",
        "承保财务损益",
        "再保净成本",
        "提取保费准备金"
    ]
    color_2024 = "#1E49E2"
    color_2025 = "#00338D"
    
    prev_year = st.session_state.get('prev_year', 2024)
    latest_year = st.session_state.get('latest_year', 2025)
    years = [str(prev_year), str(latest_year)]
    
    raw_df = st.session_state.get('integrated_data', None)
    if raw_df is None:
        st.error("❌ 未找到集成数据，请先完成 Step 5 数据集成。")
        return [go.Figure() for _ in factors]
    
    raw_df = df.copy()
    raw_df.columns = raw_df.columns.str.strip()
    raw_df['公司'] = raw_df['公司'].astype(str).str.strip()
    raw_df['字段名'] = raw_df['字段名'].astype(str).str.strip()
    raw_df['报告年份'] = raw_df['报告年份'].astype(str).str.replace('.0', '', regex=False).str.strip()
    raw_df = raw_df[raw_df['报告年份'] != '']
    raw_df = raw_df[~raw_df['报告年份'].str.lower().isin(['nan', 'none'])]
    
    companies = [c.strip() for c in cos]
    
    service_revenue = {}
    for co in companies:
        for yr in years:
            mask = (raw_df['公司'] == co) & (raw_df['报告年份'] == yr) & (raw_df['字段名'] == '保险服务收入')
            rev_series = raw_df.loc[mask, '(百万)人民币']
            rev = rev_series.sum() if not rev_series.empty else 0
            if pd.isna(rev) or rev == 0:
                mask2 = (raw_df['公司'] == co) & (raw_df['报告年份'] == yr) & (raw_df['字段名'] == '保险业务收入')
                rev_series2 = raw_df.loc[mask2, '(百万)人民币']
                rev = rev_series2.sum() if not rev_series2.empty else 0
            if rev == 0:
                st.warning(f"⚠️ 公司 {co} 在 {yr} 年未找到保险服务收入或保险业务收入，分母设为 1")
                rev = 1
            service_revenue[(co, yr)] = rev
    
    data = {}
    for f in factors:
        data[f] = {}
        for co in companies:
            df_co = raw_df[raw_df['公司'] == co]
            for yr in years:
                mask_f = (df_co['报告年份'] == yr) & (df_co['字段名'] == f)
                v_series = df_co.loc[mask_f, '(百万)人民币']
                v = v_series.sum() if not v_series.empty else 0
                denom = service_revenue.get((co, yr), 1)
                ratio = v / denom if denom != 0 else 0
                data[f][(co, yr)] = ratio if pd.notna(ratio) else 0
    
    figures = []
    for f in factors:
        y_2024 = [data[f].get((co, years[0]), 0) for co in companies]
        y_2025 = [data[f].get((co, years[1]), 0) for co in companies]
        all_vals = y_2024 + y_2025
        
        if all_vals:
            max_val = max(all_vals)
            min_val = min(all_vals)
            abs_max = max(abs(max_val), abs(min_val))
            
            if abs_max < 0.001:
                y_range = [-0.001, 0.001]
            else:
                range_min = min_val * 1.2 if min_val < 0 else 0
                range_max = max_val * 1.3 if max_val > 0 else 0
                if range_max - range_min < 0.002:
                    margin = (0.002 - (range_max - range_min)) / 2
                    range_min -= margin
                    range_max += margin
                y_range = [range_min, range_max]
        else:
            y_range = [-0.1, 0.1]
        
        fig = go.Figure()
        fig.add_trace(go.Bar(
            name=f"{years[0]}YE",
            x=companies,
            y=y_2024,
            marker_color=color_2024,
            width=0.35,
            offset=-0.2,
            text=[f"{v:.2%}" for v in y_2024],  # 显示两位小数百分比
            textposition='outside',
        ))
        fig.add_trace(go.Bar(
            name=f"{years[1]}YE",
            x=companies,
            y=y_2025,
            marker_color=color_2025,
            width=0.35,
            offset=0.2,
            text=[f"{v:.2%}" for v in y_2025],
            textposition='outside',
        ))
        
        fig.update_layout(
            barmode='group',
            title=f"<b>{f}</b>",
            height=350,
            margin=dict(t=50, b=40, l=50, r=20),
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="center", x=0.5),
            plot_bgcolor='rgba(0,0,0,0)',
            paper_bgcolor='rgba(0,0,0,0)',
            yaxis=dict(
                tickformat=".2%",
                title="占保险服务收入比例",
                range=y_range,
                zeroline=True,
                zerolinecolor='gray',
                zerolinewidth=1,
            ),
            xaxis=dict(
                tickangle=-30,
                tickfont=dict(size=10),
            )
        )
        figures.append(fig)
    
    return figures
    
# 5.通用多分类构成分析引擎
def create_kpmg_multi_composition_chart(df, field_map, color_map, title_prefix, show_labels, 
                                        label_size, bar_width, co_font_size, highlight_co="无",
                                        add_zero_line=False, years_to_show=None):
    # ===== 从 session_state 读取配置 =====
    selected_cos = st.session_state.get('selected_cos_cache', [])
    HL_BOX_FILL = st.session_state.get('HL_BOX_FILL', "rgba(0,51,141,0.03)")
    HL_BOX_LINE = st.session_state.get('HL_BOX_LINE', "rgba(0,51,141,0.35)")
    fields = list(field_map.keys())
    
    # 如果指定了年份过滤，则先过滤 df
    if years_to_show is not None:
        df = df[df['报告年份'].astype(str).isin([str(y) for y in years_to_show])].copy()
    
    d = df[df['公司'].isin(selected_cos)].drop_duplicates(subset=['公司', '报告年份', '字段名'], keep='first').copy()
    d['报告年份'] = d['报告年份'].astype(str).str.replace(".0", "", regex=False)
    
    d_p = d[d['字段名'].isin(fields)].pivot_table(
        index=['公司', '报告年份'], columns='字段名', values='(百万)人民币', aggfunc='first'
    ).fillna(0) if not d.empty else pd.DataFrame()
    
    valid_fields = [f for f in fields if f in d_p.columns]
    if valid_fields:
        d_p['Total'] = d_p[valid_fields].abs().sum(axis=1).replace(0, 1)
    else:
        d_p['Total'] = 1.0
    
    for f in fields:
        if f in d_p.columns:
            d_p[field_map[f]] = d_p[f] / d_p['Total'] * 100
        else:
            d_p[field_map[f]] = 0
    
    all_yrs = sorted(d['报告年份'].unique())
    av_cos = [co for co in selected_cos if co in d['公司'].unique()]
    if not av_cos:
        fig = go.Figure()
        fig.add_annotation(
            text="⚠️ 无数据：请检查公司或字段",
            x=0.5, y=0.5, showarrow=False,
            font=dict(size=18, color="red")
        )
        fig.update_layout(
            height=300,
            paper_bgcolor='rgba(0,0,0,0)',
            plot_bgcolor='rgba(0,0,0,0)',
            xaxis=dict(showticklabels=False),
            yaxis=dict(showticklabels=False)
        )
        return fig, pd.DataFrame()
    
    titles = [f"<b>{co}</b>" for co in av_cos]
    fig = make_subplots(rows=1, cols=len(av_cos), shared_yaxes=True, 
                        horizontal_spacing=0.015, column_titles=titles)
    
    for i, co in enumerate(av_cos):
        d_co = d[d['公司']==co].pivot(index='报告年份', columns='字段名', values='(百万)人民币').reindex(all_yrs).fillna(0)
        valid_co_fields = [f for f in fields if f in d_co.columns]
        raw_total = d_co[valid_co_fields].abs().sum(axis=1) if valid_co_fields else pd.Series(0, index=d_co.index)
        d_co['Total_raw'] = raw_total
        
        for f, d_n in field_map.items():
            if f in d_co.columns:
                val = d_co[f] / d_co['Total_raw'].replace(0, 1) * 100
            else:
                val = pd.Series(0, index=d_co.index)
            
            clr = color_map.get(f, "#CCCCCC")
            is_dark = any(x in clr for x in ["00338D", "510DBC", "1E49E2", "0A2B5E", "FD349C"])
            txt_c = "white" if is_dark else "black"
            
            fig.add_trace(go.Bar(
                x=[f"{y}YE" for y in d_co.index], 
                y=val,
                name=d_n if i==0 else None,
                marker_color=clr,
                text=[f"{v:.0f}%" if abs(v) >= 1 else "" for v in val] if show_labels else None,
                textangle=0, textposition='inside', insidetextanchor='middle',
                textfont=dict(size=label_size, color=txt_c),
                constraintext='none', cliponaxis=False,
                width=bar_width, showlegend=(i==0), legendgroup=f
            ), row=1, col=i+1)
        
        missing_y = [100 if v == 0 else 0 for v in raw_total]
        missing_t = ["未披露" if v == 0 else "" for v in raw_total]
        fig.add_trace(go.Bar(
            x=[f"{y}YE" for y in d_co.index],
            y=missing_y,
            marker_color="#CDCDCD",
            text=missing_t,
            textangle=0, textposition='inside', insidetextanchor='middle',
            textfont=dict(size=12, color="white"),
            constraintext='none', cliponaxis=False,
            width=bar_width, showlegend=False, hoverinfo='skip'
        ), row=1, col=i+1)
        
        is_hl = (str(co).strip() == str(highlight_co).strip())
        bg_fill = HL_BOX_FILL if is_hl else "rgba(0,0,0,0)"
        line_dict = dict(color=HL_BOX_LINE, width=1.5) if is_hl else dict(color="rgba(0,0,0,0)", width=0)
        fig.add_shape(
            type="rect", xref="x domain", yref="y domain",
            x0=-0.06, x1=1.06, y0=-0.1, y1=1.08,
            fillcolor=bg_fill, line=line_dict, layer="above", row=1, col=i+1
        )
    
    df_avg = d_p[[field_map[f] for f in fields if f in d_p.columns]].groupby('报告年份').mean()
    df_avg = df_avg.reindex(all_yrs[-2:] if len(all_yrs)>=2 else all_yrs)
    df_avg.index = [f"{y}YE" for y in df_avg.index]
    
    fig.update_layout(
        barmode='relative',
        height=400,
        margin=dict(t=50, b=80, l=20, r=20),
        legend=dict(orientation="h", yanchor="top", y=-0.17, xanchor="center", x=0.5, font=dict(size=10)),
        plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)'
    )
    
    if add_zero_line:
        fig.add_hline(y=0, line_color="orange", line_width=1.5, layer="below")
    
    for ann in fig.layout.annotations:
        if "<b>" in str(ann.text):
            ann.update(y=1, font=dict(size=co_font_size, color="#00338D"))
    
    return fig, df_avg

#5.0 概览表格
def create_overview_table(df, cos, latest_year, prev_year, unit_label="百万元"):
    """
    生成概览表格：按保险服务收入排序，展示各公司关键指标同比变化（%）
    返回 Plotly 表格，表头单元格包含三行：指标名称、%变化、25YE/24YE-1
    """
    # 提取所需指标（内部字段名）
    metrics = ["总资产", "净资产", "保险服务收入", "承保利润", "净利润"]
    # 对应的显示名称（将“承保利润”显示为“保险服务业绩”）
    display_names = ["总资产", "净资产", "保险服务收入", "保险服务业绩", "净利润"]
    
    # 获取最近两年数据
    df_latest = df[df['报告年份'].astype(str) == str(latest_year)]
    df_prev = df[df['报告年份'].astype(str) == str(prev_year)]
    
    # 计算各公司保险服务收入（用于排序）
    premium_latest = {}
    for co in cos:
        val = df_latest[(df_latest['公司'] == co) & (df_latest['字段名'] == '保险服务收入')]['(百万)人民币']
        premium_latest[co] = val.iloc[0] if not val.empty else 0
    
    # 按保险服务收入降序排序
    sorted_cos = sorted(cos, key=lambda x: premium_latest.get(x, 0), reverse=True)
    
    data = {'公司': sorted_cos}
    for metric in metrics:
        pct_changes = []
        for co in sorted_cos:
            curr_val = df_latest[(df_latest['公司'] == co) & (df_latest['字段名'] == metric)]['(百万)人民币']
            prev_val = df_prev[(df_prev['公司'] == co) & (df_prev['字段名'] == metric)]['(百万)人民币']
            curr = curr_val.iloc[0] if not curr_val.empty else None
            prev = prev_val.iloc[0] if not prev_val.empty else None
            if curr is not None and prev is not None and prev != 0:
                pct = (curr - prev) / abs(prev)
            else:
                pct = None
            pct_changes.append(pct)
        data[metric] = pct_changes
    
    df_table = pd.DataFrame(data)
    # 格式化百分比
    def fmt_pct(x):
        if pd.isna(x):
            return "N/A"
        return f"{x:.1%}"
    for m in metrics:
        df_table[m] = df_table[m].apply(fmt_pct)
    
    # 构建表头：第一列"公司"也使用三行（补两个空行），其他列显示三行
    header_values = ['公司<br><br>']
    for display_name in display_names:
        header_values.append(f"{display_name}<br>%变化<br>25YE/24YE-1")
    
    # 使用 Plotly 绘制表格
    fig = go.Figure(data=[go.Table(
        header=dict(
            values=header_values,
            fill_color='#00338D',
            align='center',
            font=dict(color='white', size=13)
        ),
        cells=dict(
            values=[df_table['公司']] + [df_table[m] for m in metrics],
            fill_color=[['#F8F9FA', 'white'] * (len(df_table)//2 + 1)],
            align='center',
            font=dict(size=12)
        )
    )])
    fig.update_layout(
        title="关键指标同比变化（%）",
        margin=dict(l=10, r=10, t=50, b=10),
        paper_bgcolor='rgba(0,0,0,0)',
        plot_bgcolor='rgba(0,0,0,0)'
    )
    return fig

# 5.1 通用单指标柱状图
def create_kpmg_chart(df, field_name, title_prefix, show_labels, pct_font_size, global_gap,
                      sort_by_value=False, is_percentage=False):
    # ===== 从 session_state 读取配置 =====
    selected_cos = st.session_state.get('selected_cos_cache', [])
    divisor = st.session_state.get('divisor', 1)
    prev_year = st.session_state.get('prev_year', 2024)
    latest_year = st.session_state.get('latest_year', 2025)
    highlight_co = st.session_state.get('highlight_co', "无")
    unit_label = st.session_state.get('unit_label', '百万元')
    HL_BOX_FILL = st.session_state.get('HL_BOX_FILL', "rgba(0,51,141,0.03)")
    HL_BOX_LINE = st.session_state.get('HL_BOX_LINE', "rgba(0,51,141,0.35)")

    # 清洗字段名和公司名
    df_clean = df.copy()
    df_clean['字段名'] = df_clean['字段名'].astype(str).str.strip()
    df_clean['公司'] = df_clean['公司'].astype(str).str.strip()
    field_name_clean = field_name.strip()
    selected_cos = [c.strip() for c in selected_cos]

    # 过滤数据
    d = df_clean[df_clean['公司'].isin(selected_cos)].copy()
    d['报告年份'] = d['报告年份'].astype(str).str.replace('.0', '', regex=False).str.strip()
    val_col = '(百万)人民币' if '(百万)人民币' in d.columns else d.columns[-1]

    # 模糊匹配字段名
    norm_field = normalize_field(field_name_clean)
    df_plot = d[d['字段名'].apply(lambda x: normalize_field(x) == norm_field)].copy()

    if df_plot.empty:
        fig = go.Figure()
        fig.add_annotation(text="⚠️ 无数据：请检查字段名或数据源", x=0.5, y=0.5, showarrow=False, font=dict(size=18, color="red"))
        fig.update_layout(height=700, paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)',
                          xaxis=dict(showticklabels=False), yaxis=dict(showticklabels=False))
        return fig

    # 数值缩放
    df_plot['value'] = df_plot[val_col]
    if not is_percentage:
        df_plot['value'] = df_plot['value'] / divisor

    y_old, y_new = str(prev_year), str(latest_year)

    # 排序
    if sort_by_value:
        df_new = df_plot[df_plot['报告年份'] == y_new].set_index('公司')['value']
        sorted_cos = sorted(selected_cos, key=lambda co: df_new.get(co, -float('inf')), reverse=True)
    else:
        sorted_cos = selected_cos

    fig = go.Figure()
    color_map = {y_old: "#1E49E2", y_new: "#00338D"}

    valid_vals = df_plot['value'].dropna()
    all_max = valid_vals.max() if not valid_vals.empty else 100
    abs_max = valid_vals.abs().max() if not valid_vals.empty else 100
    fixed_offset = all_max * 0.2
    placeholder_h = (abs_max * 0.15 if abs_max > 0 else 10) / divisor

    def fmt(val):
        if is_percentage:
            return f"{val:.1%}" if pd.notna(val) else ""
        else:
            if pd.notna(val):
                if divisor >= 1000:
                    return f"{val:,.2f}"
                elif divisor >= 100:
                    return f"{val:,.2f}"
                else:
                    return f"{val:,.0f}"
            return ""

    # 绘制柱状图
    for yr in [y_old, y_new]:
        df_yr = df_plot[df_plot['报告年份'] == yr].set_index('公司').reindex(sorted_cos).reset_index()
        y_vals, m_colors, t_texts, t_colors, t_pos = [], [], [], [], []
        for v in df_yr['value']:
            if pd.isna(v):
                y_vals.append(placeholder_h)
                m_colors.append("#CDCDCD")
                t_texts.append("未披露")
                t_colors.append("white")
                t_pos.append("inside")
            else:
                y_vals.append(v)
                m_colors.append(color_map[yr])
                t_texts.append(fmt(v) if show_labels and v != 0 else "")
                t_colors.append("#333333")
                t_pos.append("outside")
        fig.add_trace(go.Bar(
            name=f"{yr}YE",
            x=df_yr['公司'],
            y=y_vals,
            marker_color=m_colors,
            text=t_texts,
            textposition=t_pos,
            textfont=dict(size=12, color=t_colors),
            cliponaxis=False,
            textangle=0,
            constraintext='none'
        ))

    # 标注增长率
    df_old_series = df_plot[df_plot['报告年份'] == y_old].set_index('公司')['value']
    df_new_series = df_plot[df_plot['报告年份'] == y_new].set_index('公司')['value']
    for co in sorted_cos:
        v_old = df_old_series.get(co, np.nan)
        v_new = df_new_series.get(co, np.nan)
        if pd.notna(v_old) and pd.notna(v_new) and v_old != 0:
            if is_percentage:
                diff = (v_new - v_old) * 100
                label = f"↗ +{diff:.1f}pp" if diff >= 0 else f"↘ {diff:.1f}pp"
                color = "#FD349C" if diff >= 0 else "#269924"
            else:
                pct = (v_new - v_old) / abs(v_old)
                label = f"↗ {pct:.1%}" if pct >= 0 else f"↘ {pct:.1%}"
                color = "#FD349C" if pct >= 0 else "#269924"
            fig.add_annotation(
                x=co,
                y=max(v_old, v_new) + fixed_offset,
                text=f"<b>{label}</b>",
                showarrow=False,
                font=dict(color=color, size=pct_font_size),
                xshift=15
            )

    # 高亮框（特定追踪公司）
    if highlight_co in sorted_cos:
        idx = sorted_cos.index(highlight_co)
        fig.add_shape(
            type="rect",
            xref="x", yref="paper",
            x0=idx - 0.45, x1=idx + 0.45,
            y0=-0.08, y1=1,
            fillcolor=HL_BOX_FILL,
            line=dict(color=HL_BOX_LINE, width=1.5),
            layer="below"
        )

    # ===== 🆕 添加灰色公司边框 =====
    fig = add_company_borders(fig, sorted_cos, sorted_cos, top_margin=70)
    # 隐藏 x 轴下方的公司名称（避免重复）
    fig.update_xaxes(showticklabels=False)

    # ===== 布局设置 =====
    fig.update_layout(
        barmode='group',
        paper_bgcolor='rgba(0,0,0,0)',
        plot_bgcolor='rgba(0,0,0,0)',
        bargroupgap=0.0,
        bargap=global_gap,
        legend=dict(
            orientation="v",        # 改为垂直
            yanchor="middle",
            y=0.5,
            xanchor="right",
            x=-0.1,                # 图例紧贴图表左侧
            font=dict(size=12)
        ),
        margin=dict(t=70, b=40, l=20, r=20),
        height=700
    )
    fig.update_xaxes(showgrid=False, zeroline=False, showline=False)

    # ===== 设置纵轴 =====
    y_vals = df_plot['value'].dropna()
    if not y_vals.empty:
        y_min = y_vals.min()
        y_max = y_vals.max()
        padding = (y_max - y_min) * 0.2 if y_max != y_min else abs(y_max) * 0.3
        y_range = [y_min - padding if y_min < 0 else 0,
                   y_max + padding if y_max > 0 else 0]
        if y_range[0] == 0 and y_range[1] == 0:
            y_range = [0, 1]

        all_max_abs = y_vals.abs().max()
        fixed_offset = all_max_abs * 0.2
        max_label_y = y_max + fixed_offset if y_max > 0 else y_max
        y_range[1] = max(y_range[1], max_label_y) * 1.05

        data_range = y_range[1] - y_range[0]
        if data_range > 0:
            raw_step = data_range / 6
            magnitude = 10 ** (len(str(int(raw_step))) - 1) if raw_step >= 1 else 10 ** (len(str(int(raw_step * 10))) - 2)
            for base in [1, 2, 5]:
                candidate = base * magnitude
                if candidate >= raw_step / 2:
                    step = candidate
                    break
            else:
                step = magnitude
        else:
            step = None
    else:
        y_range = [0, 1]
        step = 1

    fig.update_yaxes(
        showgrid=False,
        zeroline=True,
        zerolinecolor="#E0E0E0",
        zerolinewidth=1,
        showline=False,
        range=y_range,
        dtick=step,
        tickformat=',.0f' if divisor <= 1 else ',.2f',
        title_text="百分比" if is_percentage else f"{unit_label}人民币",
        title_font=dict(size=12, color="#333333"),
        exponentformat='none',
    )

    return fig
                          
# 5.2 保险业务构成（两种方法占比堆叠图）
def create_method_composition_chart(df, cos, year, divisor=1, unit_label="百万元"):
    """
    绘制每家公司保险服务收入采用保费分配法与未采用保费分配法的占比堆叠图
    横轴为公司，每个公司显示两个年份（去年和今年）的堆叠柱，纵轴为百分比。
    """
    # 筛选字段（原始名称，用于匹配）
    field_未采用_raw = "保险业务构成(未采用保费分配法)"
    field_采用_raw = "保险业务构成（采用保费分配法）"
    
    # 获取年份列表（去年和今年）
    prev_year = st.session_state.get('prev_year', 2024)
    latest_year = st.session_state.get('latest_year', 2025)
    years = [str(prev_year), str(latest_year)]
    
    # 提取数据，清洗年份（去除 .0 和可能的 YE 后缀）
    df_year = df.copy()
    df_year['报告年份'] = df_year['报告年份'].astype(str).str.replace('.0', '', regex=False).str.strip()
    df_year['报告年份'] = df_year['报告年份'].str.replace('YE', '', regex=False).str.strip()
    df_year = df_year[df_year['报告年份'].isin(years)]
    df_year = df_year[df_year['公司'].isin(cos)]
    
    # 检查数据是否存在（模糊匹配）
    available_fields = df_year['字段名'].unique()
    norm_available = [normalize_field(f) for f in available_fields]
    target_未 = normalize_field(field_未采用_raw)
    target_采 = normalize_field(field_采用_raw)
    
    actual_未 = None
    actual_采 = None
    for f in available_fields:
        if normalize_field(f) == target_未:
            actual_未 = f
        if normalize_field(f) == target_采:
            actual_采 = f
    
    if actual_未 is None or actual_采 is None:
        fig = go.Figure()
        fig.add_annotation(
            text="未找到保费分配法构成数据",
            x=0.5, y=0.5, showarrow=False,
            font=dict(size=16, color="red")
        )
        fig.update_layout(height=300, paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)')
        return fig
    
    # 构建数据结构：{(公司, 年份): {未采用占比, 采用占比}}
    data = {}
    for co in cos:
        for yr in years:
            # 获取未采用占比
            val_un = df_year[(df_year['公司'] == co) & (df_year['报告年份'] == yr) & (df_year['字段名'] == actual_未)]['(百万)人民币']
            val_un = val_un.iloc[0] if not val_un.empty else 0
            # 获取采用占比
            val_ad = df_year[(df_year['公司'] == co) & (df_year['报告年份'] == yr) & (df_year['字段名'] == actual_采)]['(百万)人民币']
            val_ad = val_ad.iloc[0] if not val_ad.empty else 0
            # 如果两个都是0，跳过
            if val_un == 0 and val_ad == 0:
                data[(co, yr)] = None
            else:
                # 转换为百分比（小数转百分数）
                data[(co, yr)] = {
                    "未采用": val_un * 100,
                    "采用": val_ad * 100
                }
    
    # 如果没有有效数据
    if not any(v is not None for v in data.values()):
        fig = go.Figure()
        fig.add_annotation(
            text="所有公司均无有效构成数据",
            x=0.5, y=0.5, showarrow=False,
            font=dict(size=16, color="red")
        )
        fig.update_layout(height=300, paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)')
        return fig
    
    # 创建子图布局：每个公司一列，共享y轴
    num_companies = len(cos)
    fig = make_subplots(
        rows=1, cols=num_companies,
        shared_yaxes=True,
        horizontal_spacing=0.02
    )
    
    # 颜色
    colors = {"未采用": "#00338D", "采用": "#0865EE"}
    
    for i, co in enumerate(cos):
        col_idx = i + 1
        x_vals = []
        y_un = []
        y_ad = []
        # 按年份顺序（去年、今年）
        for yr in years:
            x_vals.append(f"{yr}YE")
            d = data.get((co, yr))
            if d is None:
                y_un.append(0)
                y_ad.append(0)
            else:
                y_un.append(d["未采用"])
                y_ad.append(d["采用"])
        
        # 未采用 trace
        fig.add_trace(
            go.Bar(
                x=x_vals,
                y=y_un,
                name="未采用保费分配法",
                legendgroup="未采用保费分配法",
                marker_color=colors["未采用"],
                text=[f"{v:.1f}%" if v > 0 else "" for v in y_un],
                textposition='inside',
                insidetextanchor='middle',
                textfont=dict(color="white", size=11),
                width=0.6,
                showlegend=(i == 0)      # 只在第一个子图显示图例
            ),
            row=1, col=col_idx
        )
        # 采用 trace
        fig.add_trace(
            go.Bar(
                x=x_vals,
                y=y_ad,
                name="采用保费分配法",
                legendgroup="采用保费分配法",
                marker_color=colors["采用"],
                text=[f"{v:.1f}%" if v > 0 else "" for v in y_ad],
                textposition='inside',
                insidetextanchor='middle',
                textfont=dict(color="white", size=11),
                width=0.6,
                showlegend=(i == 0)
            ),
            row=1, col=col_idx
        )
        
        # ===== 🆕 为当前子图添加公司边框 =====
        # x_vals 是该子图的横轴标签（年份）
        fig = add_company_borders(fig, [co], x_vals, top_margin=40, row=1, col=col_idx)
    
    fig.update_layout(
        barmode='relative',  # 堆叠模式
        height=400,
        margin=dict(t=40, b=40, l=40, r=20),
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="center",
            x=0.5
        ),
        plot_bgcolor='rgba(0,0,0,0)',
        paper_bgcolor='rgba(0,0,0,0)',
        yaxis=dict(
            title="占比（%）",
            tickformat=".0f",
            range=[0, 105]
        )
    )
    
    # 更新每个子图的x轴和y轴
    for i in range(1, num_companies + 1):
        fig.update_xaxes(
            tickangle=0,
            showline=False,
            showgrid=False,
            zeroline=False,
            row=1, col=i
        )
        fig.update_yaxes(
            showgrid=False,
            zeroline=False,
            showline=False,
            row=1, col=i
        )
    
    return fig

# 5.3 利润构成图（含公司边框）
def create_profit_composition_chart(df, cos, year, divisor=1, unit_label="百万元"):
    """
    绘制每家公司承保利润与投资利润的分组条形图（实际金额），并标注占比。
    添加灰色公司边框，隐藏x轴公司名称。
    """
    field_uw = "承保利润"
    field_inv = "投资利润"
    prev_year = st.session_state.get('prev_year', 2024)
    latest_year = st.session_state.get('latest_year', 2025)
    years = [str(prev_year), str(latest_year)]

    df_year = df.copy()
    df_year['报告年份'] = df_year['报告年份'].astype(str).str.replace('.0', '', regex=False).str.strip()
    df_year['报告年份'] = df_year['报告年份'].str.replace('YE', '', regex=False).str.strip()
    df_year = df_year[df_year['报告年份'].isin(years)]
    df_year = df_year[df_year['公司'].isin(cos)]

    available_fields = df_year['字段名'].unique()
    norm_available = [normalize_field(f) for f in available_fields]
    target_uw = normalize_field(field_uw)
    target_inv = normalize_field(field_inv)

    actual_uw = None
    actual_inv = None
    for f in available_fields:
        if normalize_field(f) == target_uw:
            actual_uw = f
        if normalize_field(f) == target_inv:
            actual_inv = f

    if actual_uw is None or actual_inv is None:
        fig = go.Figure()
        fig.add_annotation(text="未找到承保利润或投资利润数据", x=0.5, y=0.5, showarrow=False, font=dict(size=16, color="red"))
        fig.update_layout(height=400, paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)')
        return fig

    # 提取各公司各年份的原始值（已换算为百万元）
    raw_data = {}
    for co in cos:
        for yr in years:
            val_uw = df_year[(df_year['公司'] == co) & (df_year['报告年份'] == yr) & (df_year['字段名'] == actual_uw)]['(百万)人民币']
            val_uw = val_uw.iloc[0] if not val_uw.empty else 0
            val_inv = df_year[(df_year['公司'] == co) & (df_year['报告年份'] == yr) & (df_year['字段名'] == actual_inv)]['(百万)人民币']
            val_inv = val_inv.iloc[0] if not val_inv.empty else 0
            raw_data[(co, yr)] = {"承保利润": val_uw / divisor, "投资利润": val_inv / divisor}

    yr_str = str(year)
    companies = cos
    data = {co: raw_data.get((co, yr_str), {"承保利润": 0, "投资利润": 0}) for co in companies}

    labels = ["承保利润", "投资利润"]
    colors = ["#00338D", "#0865EE"]
    fig = go.Figure()

    # 收集所有值用于确定y轴范围
    all_vals = []
    for co in companies:
        all_vals.append(data[co]["承保利润"])
        all_vals.append(data[co]["投资利润"])

    y_max = max(all_vals) if all_vals else 1
    y_min = min(all_vals) if all_vals else 0
    y_range = [min(0, y_min * 1.2), max(0, y_max * 1.3)]

    # 绘制分组条形
    for i, (label, color) in enumerate(zip(labels, colors)):
        values = [data[co][label] for co in companies]
        ratios = []
        for co in companies:
            total = data[co]["承保利润"] + data[co]["投资利润"]
            if total != 0:
                ratios.append(data[co][label] / total)
            else:
                ratios.append(0)
        texts = []
        for val, rat in zip(values, ratios):
            if val == 0:
                texts.append("")
            else:
                texts.append(f"{val:.1f}\n({rat:.1%})")
        fig.add_trace(go.Bar(
            name=label,
            x=companies,
            y=values,
            marker_color=color,
            text=texts,
            textposition='outside',
            textfont=dict(size=10),
            width=0.35,
            offsetgroup=i,
            cliponaxis=False
        ))

    fig.add_hline(y=0, line_dash="dash", line_color="gray", line_width=1.5, opacity=0.7)

    # ===== 🆕 添加灰色公司边框 =====
    fig = add_company_borders(fig, companies, companies, top_margin=70)
    # 隐藏 x 轴下方的公司名称（避免重复）
    fig.update_xaxes(showticklabels=False)

    fig.update_layout(
        barmode='group',
        title=f"各公司利润构成（最新年 {year}YE）",
        xaxis_title="公司",
        yaxis_title=f"金额（{unit_label}）",
        yaxis=dict(range=y_range, tickformat=",.0f", zeroline=True, zerolinecolor='gray'),
        height=450,
        margin=dict(t=70, b=60, l=40, r=20),  # 顶部边距从 50 改为 70
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="center", x=0.5),
        plot_bgcolor='rgba(0,0,0,0)',
        paper_bgcolor='rgba(0,0,0,0)',
        bargap=0.15,
        bargroupgap=0.1,
    )
    return fig
    
# 5.4 综合赔付率拆解堆叠图（自动提取年份，优化配色 + 公司边框 + 年份标注）
def create_cor_breakdown_stacked_chart(df, cos, latest_year, prev_year, divisor=1, unit_label="百万元", highlight_co="无"):
    """
    绘制综合赔付率拆解的堆叠柱状图（各因子占保险服务收入比例）
    自动从数据中提取最近两个年份，每个柱子下方显示年份标签。
    """
    factors = [
        "当期发生赔款及理赔费用",
        "已发生赔款负债履约现金流变动",
        "亏损合同损益",
        "承保财务损益",
        "再保净成本",
        "提取保费准备金"
    ]
    
    factor_colors = {
        "当期发生赔款及理赔费用": "#00338D",
        "已发生赔款负债履约现金流变动": "#510DBC",
        "亏损合同损益": "#B0BEC5",
        "承保财务损益": "#76D2FF",
        "再保净成本": "#FD349C",
        "提取保费准备金": "#00C0AE"
    }
    
    raw = df.copy()
    raw['报告年份'] = raw['报告年份'].astype(str).str.replace('.0', '', regex=False)
    raw['公司'] = raw['公司'].astype(str).str.strip()
    raw['字段名'] = raw['字段名'].astype(str).str.strip()
    
    # 提取实际存在的年份
    available_years = sorted(
        [int(y) for y in raw['报告年份'].unique() if y.isdigit()],
        reverse=True
    )
    if len(available_years) >= 2:
        years = [str(available_years[1]), str(available_years[0])]
        year_display = f"{available_years[1]}YE vs {available_years[0]}YE"
    elif len(available_years) == 1:
        years = [str(available_years[0]), str(available_years[0])]
        year_display = f"{available_years[0]}YE"
    else:
        years = [str(prev_year), str(latest_year)]
        year_display = f"{prev_year}YE vs {latest_year}YE"
    
    # 计算分母（保险服务收入）
    service_revenue = {}
    for co in cos:
        for yr in years:
            rev = raw[(raw['公司'] == co) & (raw['报告年份'] == yr) & (raw['字段名'] == '保险服务收入')]['(百万)人民币']
            service_revenue[(co, yr)] = rev.sum() if not rev.empty else 1.0
    
    # 构建数据
    rows = []
    for co in cos:
        for yr in years:
            row = {'公司': co, '年份': yr}
            for f in factors:
                val = raw[(raw['公司'] == co) & (raw['报告年份'] == yr) & (raw['字段名'] == f)]['(百万)人民币']
                val_sum = val.sum() if not val.empty else 0
                ratio = val_sum / service_revenue[(co, yr)] * 100
                row[f] = ratio
            rows.append(row)
    df_plot = pd.DataFrame(rows)
    
    # ===== 🆕 修改：x轴标签只显示年份 =====
    df_plot['x_label'] = df_plot['年份'] + 'YE'  # 只显示年份，不包含公司名
    x_labels = df_plot['x_label'].unique()  # 实际是 ["2024YE", "2025YE"] 循环
    
    # 创建图表
    fig = go.Figure()
    for f in factors:
        fig.add_trace(go.Bar(
            x=df_plot['x_label'],
            y=df_plot[f],
            name=f,
            marker_color=factor_colors[f],
            legendgroup=f,
            text=[f"{v:.1f}%" if abs(v) > 0.5 else "" for v in df_plot[f]],
            textposition='inside',
            insidetextanchor='middle',
            textfont=dict(size=9, color='white'),
            hovertemplate=f"{f}: %{{y:.1f}}%<extra>%{{x}}</extra>"
        ))
    
    fig.add_hline(y=0, line_dash="dash", line_color="gray", line_width=1, opacity=0.5)
    
    # ===== 添加公司边框（公司名称在框线上方） =====
    df_plot['x_label'] = df_plot['公司'] + '<br>' + df_plot['年份'] + 'YE'
    x_labels = df_plot['x_label'].unique()
    
    # 创建图表（使用原始x_label）
    fig = go.Figure()
    for f in factors:
        fig.add_trace(go.Bar(
            x=df_plot['x_label'],
            y=df_plot[f],
            name=f,
            marker_color=factor_colors[f],
            legendgroup=f,
            text=[f"{v:.1f}%" if abs(v) > 0.5 else "" for v in df_plot[f]],
            textposition='inside',
            insidetextanchor='middle',
            textfont=dict(size=9, color='white'),
            hovertemplate=f"{f}: %{{y:.1f}}%<extra>%{{x}}</extra>"
        ))
    
    fig.add_hline(y=0, line_dash="dash", line_color="gray", line_width=1, opacity=0.5)
    
    # 添加公司边框
    fig = add_company_borders(fig, cos, x_labels, top_margin=70)
    
    # ===== 设置x轴显示年份标签 =====
    # 获取所有唯一的x_label（公司+年份）的列表，按顺序
    all_labels = df_plot['x_label'].tolist()
    # 生成年份列表
    year_labels = []
    for label in all_labels:
        # 提取年份部分，格式为"公司<br>年份YE"
        if '<br>' in label:
            year_part = label.split('<br>')[1]
        else:
            year_part = label
        year_labels.append(year_part)
    
    # 更新x轴：使用年份作为显示文本
    fig.update_xaxes(
        tickvals=all_labels,  # 实际的标签值
        ticktext=year_labels,  # 显示为年份
        showticklabels=True,
        tickangle=0
    )
    
    # 高亮框（如有指定公司）
    if highlight_co != "无":
        highlight_indices = [i for i, label in enumerate(all_labels) if highlight_co in label]
        for idx in highlight_indices:
            fig.add_shape(
                type="rect",
                xref="x", yref="y",
                x0=idx - 0.48, x1=idx + 0.48,
                y0=0, y1=1,
                fillcolor="rgba(0,51,141,0.05)",
                line=dict(color="rgba(0,51,141,0.8)", width=1.5),
                layer="below"
            )
    
    # 布局设置
    fig.update_layout(
        barmode='relative',
        title=f"综合赔付率拆解（{year_display}）",
        xaxis_title="",
        yaxis_title="占保险服务收入比例（%）",
        legend=dict(
            orientation="v",
            yanchor="middle",
            y=0.5,
            xanchor="right",
            x=-0.15,
            font=dict(size=11)
        ),
        height=550,
        margin=dict(l=20, r=20, t=70, b=60),
        plot_bgcolor='rgba(0,0,0,0)',
        paper_bgcolor='rgba(0,0,0,0)',
        bargap=0.15,
        bargroupgap=0.1,
        hovermode='x unified'
    )
    
    return fig
    
# 5.5 保险服务收入业务构成堆叠图（保费贡献） - 按全局公司顺序
def create_premium_stacked_chart(df, cos, year, divisor=1, unit_label="百万元", highlight_co="无"):
    """
    绘制各公司保险服务收入业务构成堆叠图（按险种）
    """
    field_type = "保险服务收入"
    prefix = f"{field_type}-"
    all_fields = df['字段名'].unique()
    business_fields = [f for f in all_fields if isinstance(f, str) and f.startswith(prefix)]
    business_fields = [f for f in business_fields if not f.endswith("合计")]

    if not business_fields:
        fig = go.Figure()
        fig.add_annotation(text=f"未找到{field_type}业务构成数据", x=0.5, y=0.5, showarrow=False, font=dict(size=16, color="red"))
        fig.update_layout(height=400)
        return fig

    df_year = df[df['报告年份'].astype(str) == str(year)]
    df_year = df_year[df_year['公司'].isin(cos)]

    # 构建数据透视
    pivot_data = []
    for co in cos:
        row = {'公司': co}
        for f in business_fields:
            val_series = df_year[(df_year['公司'] == co) & (df_year['字段名'] == f)]['(百万)人民币']
            val = val_series.iloc[0] if not val_series.empty else 0
            label = f.replace(prefix, "")
            row[label] = val / divisor
        pivot_data.append(row)
    df_plot = pd.DataFrame(pivot_data)
    display_labels = [c for c in df_plot.columns if c != '公司']

    if not df_plot[display_labels].map(lambda x: x != 0).any().any():
        fig = go.Figure()
        fig.add_annotation(text="所有公司业务构成均为零或未披露", x=0.5, y=0.5, showarrow=False, font=dict(size=16, color="red"))
        fig.update_layout(height=400)
        return fig

    # 按传入的 cos 顺序确保 df_plot 的顺序
    df_plot = df_plot.set_index('公司').reindex(cos).reset_index()

    # 计算每个公司的总金额（用于百分比）
    totals = df_plot[display_labels].sum(axis=1)

    # ===== 唯一颜色分配 =====
    import plotly.express as px
    base_colors = KPMG_COLORS
    if len(display_labels) > len(base_colors):
        extra_colors = px.colors.qualitative.Plotly
        extra_unique = [c for c in extra_colors if c not in base_colors]
        all_colors = base_colors + extra_unique
    else:
        all_colors = base_colors
    if len(display_labels) > len(all_colors):
        all_colors = all_colors + px.colors.qualitative.Alphabet
    color_map = {label: all_colors[i % len(all_colors)] for i, label in enumerate(display_labels)}

    fig = go.Figure()

    for i, label in enumerate(display_labels):
        values = df_plot[label].fillna(0).tolist()
        ratios = []
        for val, total in zip(values, totals):
            if total != 0:
                ratios.append(val / total * 100)
            else:
                ratios.append(0)
        texts = []
        for val, rat in zip(values, ratios):
            if val == 0:
                texts.append("")
            else:
                texts.append(f"{rat:.1f}%")
        fig.add_trace(go.Bar(
            name=label,
            x=df_plot['公司'],
            y=ratios,
            marker_color=color_map[label],
            text=texts,
            textposition='inside',
            insidetextanchor='middle',
            textfont=dict(size=10, color='white' if i % 2 == 0 else 'black'),
            hovertemplate=f"{label}: %{{y:.1f}}%<extra>%{{x}}</extra>",
            showlegend=True,
            legendgroup=label,
        ))

    fig.add_hline(y=0, line_dash="dash", line_color="gray", line_width=1, opacity=0.5)

    if highlight_co != "无" and highlight_co in df_plot['公司'].values:
        idx = df_plot['公司'].tolist().index(highlight_co)
        fig.add_shape(
            type="rect",
            xref="x", yref="y",
            x0=idx - 0.45, x1=idx + 0.45,
            y0=0, y1=1,
            fillcolor="rgba(0,51,141,0.05)",
            line=dict(color="rgba(0,51,141,0.8)", width=1.5),
            layer="below"
        )

    fig.update_layout(
        barmode='relative',
        title=f"保险服务收入业务构成（{year}YE）",
        xaxis_title="公司",
        yaxis_title="占比（%）",
        legend=dict(
            orientation="v",
            yanchor="middle",
            y=0.5,
            xanchor="right",
            x=-0.15,
            font=dict(size=11)
        ),
        height=550,
        margin=dict(l=20, r=20, t=50, b=50),
        plot_bgcolor='rgba(0,0,0,0)',
        paper_bgcolor='rgba(0,0,0,0)',
        bargap=0.15,
        bargroupgap=0.1,
        hovermode='x unified'
    )
    return fig

# 5.6 承保利润业务构成堆叠图 - 按全局公司顺序
def create_profit_contribution_stacked_chart(df, cos, year, divisor=1, unit_label="百万元", highlight_co="无"):
    """
    绘制各公司承保利润业务构成堆叠图（按险种，显示占比）
    支持正负值堆叠（正数向上，负数向下）
    """
    field_type = "承保利润"
    prefix = f"{field_type}-"
    all_fields = df['字段名'].unique()
    business_fields = [f for f in all_fields if isinstance(f, str) and f.startswith(prefix)]
    business_fields = [f for f in business_fields if not f.endswith("合计")]

    if not business_fields:
        fig = go.Figure()
        fig.add_annotation(text=f"未找到{field_type}业务构成数据", x=0.5, y=0.5, showarrow=False, font=dict(size=16, color="red"))
        fig.update_layout(height=400)
        return fig

    df_year = df[df['报告年份'].astype(str) == str(year)]
    df_year = df_year[df_year['公司'].isin(cos)]

    pivot_data = []
    for co in cos:
        row = {'公司': co}
        for f in business_fields:
            val_series = df_year[(df_year['公司'] == co) & (df_year['字段名'] == f)]['(百万)人民币']
            val = val_series.iloc[0] if not val_series.empty else 0
            label = f.replace(prefix, "")
            row[label] = val / divisor
        pivot_data.append(row)
    df_plot = pd.DataFrame(pivot_data)
    display_labels = [c for c in df_plot.columns if c != '公司']

    if not df_plot[display_labels].map(lambda x: x != 0).any().any():
        fig = go.Figure()
        fig.add_annotation(text="所有公司业务构成均为零或未披露", x=0.5, y=0.5, showarrow=False, font=dict(size=16, color="red"))
        fig.update_layout(height=400)
        return fig

    # 按传入的 cos 顺序确保 df_plot 的顺序
    df_plot = df_plot.set_index('公司').reindex(cos).reset_index()

    # 计算每个公司的绝对值总和（用于归一化）
    abs_sums = {}
    for co in cos:
        vals = [df_plot[df_plot['公司'] == co][label].iloc[0] for label in display_labels]
        total = sum(abs(v) for v in vals)
        abs_sums[co] = total if total != 0 else 1.0

    # 归一化为百分比（保留符号）
    for co in cos:
        for label in display_labels:
            val = df_plot[df_plot['公司'] == co][label].iloc[0]
            pct = val / abs_sums[co] * 100
            df_plot.loc[df_plot['公司'] == co, label] = pct

    # ===== 唯一颜色分配 =====
    import plotly.express as px
    base_colors = KPMG_COLORS
    if len(display_labels) > len(base_colors):
        extra_colors = px.colors.qualitative.Plotly
        extra_unique = [c for c in extra_colors if c not in base_colors]
        all_colors = base_colors + extra_unique
    else:
        all_colors = base_colors
    if len(display_labels) > len(all_colors):
        all_colors = all_colors + px.colors.qualitative.Alphabet
    color_map = {label: all_colors[i % len(all_colors)] for i, label in enumerate(display_labels)}

    fig = go.Figure()

    for i, label in enumerate(display_labels):
        values = []
        for co in cos:
            val = df_plot[df_plot['公司'] == co][label].iloc[0]
            values.append(val)
        texts = []
        for idx, co in enumerate(cos):
            val = values[idx]
            if val == 0:
                texts.append("")
            else:
                texts.append(f"{val:.1f}%")
        fig.add_trace(go.Bar(
            name=label,
            x=cos,
            y=values,
            marker_color=color_map[label],
            text=texts,
            textposition='inside',
            insidetextanchor='middle',
            textfont=dict(size=10, color='white' if i % 2 == 0 else 'black'),
            hovertemplate=f"{label}: %{{y:.1f}}%<extra>%{{x}}</extra>",
            showlegend=True,
            legendgroup=label,
        ))

    fig.add_hline(y=0, line_dash="dash", line_color="gray", line_width=1, opacity=0.5)

    if highlight_co != "无" and highlight_co in cos:
        idx = cos.index(highlight_co)
        fig.add_shape(
            type="rect",
            xref="x", yref="y",
            x0=idx - 0.45, x1=idx + 0.45,
            y0=-100, y1=100,
            fillcolor="rgba(0,51,141,0.05)",
            line=dict(color="rgba(0,51,141,0.8)", width=1.5),
            layer="below"
        )

    fig.update_layout(
        barmode='relative',
        title=f"承保利润业务构成（{year}YE）",
        xaxis_title="公司",
        yaxis_title="占比（%）",
        legend=dict(
            orientation="v",
            yanchor="middle",
            y=0.5,
            xanchor="right",
            x=-0.15,
            font=dict(size=11)
        ),
        height=550,
        margin=dict(l=20, r=20, t=50, b=50),
        plot_bgcolor='rgba(0,0,0,0)',
        paper_bgcolor='rgba(0,0,0,0)',
        bargap=0.15,
        bargroupgap=0.1,
        hovermode='x unified'
    )
    return fig

# 6.汇总表
def create_nonlife_summary_table(df, cos, highlight_co="无"):
    cy, py = latest_year, prev_year
    col_names, c1, c2, c3, c4, c5, c6, c7, c8, c9, c10 = [], [], [], [], [], [], [], [], [], [], []
    
    for co in cos:
        df_co = df[df['公司'] == co]

        def get_val(y, kw):
            s = df_co[(df_co['报告年份'].astype(str) == str(y)) & 
                      (df_co['字段名'].astype(str).str.contains(kw, na=False))].drop_duplicates(
                          subset=['公司','报告年份','字段名'])['(百万)人民币']
            if s.empty: return None
            v = s.sum(min_count=1)
            return None if pd.isna(v) else v

        def calc(n, d, is_growth=False, is_ratio=False):
            if n is None or d is None: return "未披露"
            if d == 0: return "-"
            if is_ratio:
                return f"{n/d:.1%}"
            ratio = (n / d) - 1 if is_growth else (n / d)
            if is_growth:
                is_positive = n > d
            else:
                is_positive = ratio >= 0
            return f"{abs(ratio):.0%}" if is_positive else f"-{abs(ratio):.0%}"

        prem_cy = get_val(cy, '保险业务收入') or get_val(cy, '已赚保费')
        prem_py = get_val(py, '保险业务收入') or get_val(py, '已赚保费')
        
        invest_comp_cy = get_val(cy, '投资成分')
        inv_ratio_cy = invest_comp_cy / prem_cy if (invest_comp_cy is not None and prem_cy) else None

        claim_cy = get_val(cy, '已发生赔款') or get_val(cy, '赔付支出')
        comm_cy = get_val(cy, '手续费及佣金支出')
        adm_cy  = get_val(cy, '业务及管理费')
        if prem_cy is None or prem_cy == 0:
            cor_cy = None
        else:
            total_cost = (claim_cy or 0) + (comm_cy or 0) + (adm_cy or 0)
            cor_cy = total_cost / prem_cy

        loss_ratio_cy = claim_cy / prem_cy if (claim_cy is not None and prem_cy) else None
        exp_ratio_cy = ((comm_cy or 0) + (adm_cy or 0)) / prem_cy if prem_cy else None

        current_claim_cy = get_val(cy, '当期发生赔款') or get_val(cy, '当期赔款及理赔费用')
        curr_claim_ratio = current_claim_cy / prem_cy if (current_claim_cy is not None and prem_cy) else None

        comm_ratio = comm_cy / prem_cy if (comm_cy is not None and prem_cy) else None

        loss_comp_cy = get_val(cy, '亏损部分') or get_val(cy, '亏损合同')
        lrc_nonloss_cy = get_val(cy, '未到期责任负债-非亏损部分') or get_val(cy, 'LRC非亏损部分')
        if loss_comp_cy is not None and lrc_nonloss_cy is not None:
            loss_comp_ratio_cy = loss_comp_cy / (loss_comp_cy + lrc_nonloss_cy)
        else:
            loss_comp_ratio_cy = None

        back_dev_cy = get_val(cy, '回溯偏差') or get_val(cy, '已发生赔款负债回溯偏差')
        back_dev_ratio = back_dev_cy / claim_cy if (back_dev_cy is not None and claim_cy) else None

        col_names.append(co)
        c1.append(prem_cy)
        c2.append(calc(prem_cy, prem_py, True))
        c3.append(calc(inv_ratio_cy, 1, False, True) if inv_ratio_cy is not None else "未披露")
        c4.append(calc(cor_cy, 1, False, True) if cor_cy is not None else "未披露")
        c5.append(calc(loss_ratio_cy, 1, False, True) if loss_ratio_cy is not None else "未披露")
        c6.append(calc(exp_ratio_cy, 1, False, True) if exp_ratio_cy is not None else "未披露")
        c7.append(calc(curr_claim_ratio, 1, False, True) if curr_claim_ratio is not None else "未披露")
        c8.append(calc(comm_ratio, 1, False, True) if comm_ratio is not None else "未披露")
        c9.append(calc(loss_comp_ratio_cy, 1, False, True) if loss_comp_ratio_cy is not None else "未披露")
        c10.append(calc(back_dev_ratio, 1, False, True) if back_dev_ratio is not None else "未披露")

    headers = [
        "公司名称", "保险业务收入", "保费增长率<br>(新旧准则)",
        "投资成分占比", "综合成本率<br>(COR)", "综合赔付率",
        "综合费用率", "当期赔款/理赔费用<br>占比", "手续费占比",
        "亏损成分占比", "回溯偏差"
    ]
    
    current_hl = str(highlight_co).strip()
    html = "<table style='width:100%; border-collapse:collapse; font-family:sans-serif; font-size:10px; margin-bottom:15px;'>"
    html += "<tr style='background-color:#00338D; color:white; text-align:center; font-weight:bold;'>"
    for idx, h in enumerate(headers):
        html += f"<th style='padding:4px 2px; {'text-align:left;' if idx==0 else 'text-align:center;'} border:1.5px solid white;'>{h}</th>"
    html += "</tr>"

    for i, co in enumerate(col_names):
        row_vals = [c1[i], c2[i], c3[i], c4[i], c5[i], c6[i], c7[i], c8[i], c9[i], c10[i]]
        is_hl = (str(co).strip() == current_hl)
        bg = HL_BOX_FILL if is_hl else ("white" if i % 2 == 0 else "#F8F9FA")
        base = (f"background-color:{bg}; padding:3px; font-size:10px; "
                + ("border-top:1.5px solid #00338D; border-bottom:1.5px solid #00338D; font-weight:bold;"
                   if is_hl else "border:1px solid #EAEAEA;"))
        s_first = base + f"text-align:left; color:#333333; {'border-left:1.5px solid #00338D;' if is_hl else ''}"
        s_mid   = base + f"text-align:center; {'color:#00338D; border-left:none; border-right:none;' if is_hl else 'color:#444444;'}"
        s_last  = base + f"text-align:center; {'color:#00338D; border-right:1.5px solid #00338D; border-left:none;' if is_hl else 'color:#444444;'}"

        html += f"<tr><td style='{s_first}'>{co}</td>"
        for idx, v in enumerate(row_vals):
            style = s_last if idx == len(row_vals)-1 else s_mid
            if v == "未披露":
                cell_style = style.replace(f"background-color:{bg};", "background-color:#CDCDCD;") \
                                  .replace("color:#444444;", "color:white;") \
                                  .replace("color:#00338D;", "color:white;") 
                html += f"<td style='{cell_style}'>未披露</td>"
            else:
                html += f"<td style='{style}'>{v}</td>"
        html += "</tr>"

    return html + "</table>"

# 7.主页面函数 show_step_7_content
def show_step_7_content():
    global notes_dict, ordered_modules
    # ----- 样式与前端 -----
    st.markdown("""
    <style>
    [data-testid="stSidebar"] { background: rgba(255,255,255,0.95) !important; border-right: 1px solid #EAEAEA !important; box-shadow: 2px 0px 15px rgba(0,0,0,0.08) !important; }
    .nav-floating-sign { position: fixed; left: 0; top: 50%; transform: translateY(-50%); background: rgba(0, 51, 141, 0.85); color: white; padding: 20px 8px; border-radius: 0 12px 12px 0; writing-mode: vertical-rl; text-orientation: mixed; font-size: 22px; font-weight: bold; letter-spacing: 3px; z-index: 9999; cursor: pointer; box-shadow: 3px 3px 12px rgba(0,0,0,0.25); transition: all 0.2s; }
    .nav-floating-sign:hover { background: rgba(0, 51, 141, 1); padding-left: 15px; }
    .stPlotlyChart { width: 100% !important; min-width: 0 !important; }
    .print-only { display: none !important; }
    .cover-page { position: relative !important; width: 338.67mm !important; height: 190.5mm !important; margin: 0 !important; padding: 0 !important; page-break-after: always !important; overflow: hidden !important; background: transparent !important; }
    .cover-page img { width: 100% !important; height: 100% !important; object-fit: cover !important; display: block !important; }
    .block-container { padding-top:0 !important; padding-right:10px !important; padding-left:10px !important; margin-top:0 !important; }
    .cover-text { forced-color-adjust: none !important; -webkit-print-color-adjust: exact !important; print-color-adjust: exact !important; color: white !important; -webkit-text-fill-color: white !important; }
    .element-container:first-child{ margin-top:0 !important; padding-top:0 !important; }
    /* 新增：保证每个模块整体不跨页 */
    .page-break-container { 
        page-break-inside: avoid; 
        break-inside: avoid; 
        margin: 0 !important; 
        padding: 0 !important; 
    }
    .page-break-title { 
        page-break-before: always; 
        break-before: page; 
        padding-top: 10px !important; 
        margin-top: 0 !important; 
        text-align: left !important; 
    }
    @media print {
        /* 页面设置：保留动态脚本覆盖，但默认边距稍大，避免内容溢出 */
        @page { 
            size: A4 portrait; 
            margin: 15mm 15mm 15mm 15mm; 
            @bottom-center { 
                content: counter(page); 
                font-size: 12px; 
                color: #666; 
                font-family: Microsoft YaHei, sans-serif; 
            } 
        }
        .print-only { display: block !important; }
        /* 移除固定宽高限制，允许内容自然展开，避免截断 */
        html, body { 
            width: 100% !important; 
            height: auto !important; 
            overflow: visible !important; 
            zoom: 100% !important; 
        }
        .main .block-container { 
            width: 100% !important; 
            max-width: 100% !important; 
            min-width: 0 !important; 
            padding: 0 10mm !important; 
        }
        .block-container { padding-top: 0rem !important; }
        .no-print, h1, .nav-floating-sign, [data-testid="collapsedControl"], header, footer, [data-testid="stHeader"], [data-testid="stSidebar"], section[data-testid="stSidebar"], [data-testid="stToolbar"], button[kind="secondary"], input, .stSlider, [data-testid="stSelectbox"], [data-testid="stRadio"], [data-testid="stExpander"], .stAlert, button[role="tab"], div[role="tablist"], [data-baseweb="tab-list"], hr { display: none !important; }
        /* 每个模块容器自动分页，内部不拆分 */
        .page-break-container { 
            page-break-inside: avoid !important; 
            break-inside: avoid !important; 
            margin: 0 !important; 
            padding: 0 !important; 
            padding-bottom: 5mm !important; 
        }
        .stApp { max-width: 100% !important; width: 100% !important; }
        .keep-columns [data-testid="stHorizontalBlock"]{ display:flex!important; flex-wrap:nowrap!important; align-items:flex-start!important; justify-content:space-between!important; gap:0!important; width:100%!important; }
        .keep-columns [data-testid="stHorizontalBlock"]>div{ width:49%!important; min-width:49%!important; max-width:49%!important; flex:0 0 49%!important; overflow:hidden!important; page-break-inside:avoid!important; break-inside:avoid!important; }
        /* 标题强制换页 */
        .page-break-title { 
            page-break-before: always !important; 
            break-before: page !important; 
            padding-top: 10px !important; 
            margin-top: 0 !important; 
            text-align: left !important; 
        }
        h2 { display: block !important; text-align: left !important; color: #00338D !important; font-size: 30px !important; font-weight: bold !important; border-bottom: 2px solid #00338D !important; padding-bottom: 6px !important; margin: 14px 0 10px 0 !important; }
        h3:not(.no-print) { display: block !important; text-align: left !important; color: #00338D !important; font-size: 30px !important; font-weight: bold !important; margin: 10px 0 8px 0 !important; page-break-after: avoid !important; }
        /* 图表自适应，不溢出，强制避免分页 */
        .plotly-graph-div,
        .stPlotlyChart {
            width: 100% !important;
            max-width: 100% !important;
            page-break-inside: avoid !important;
            break-inside: avoid !important;
            display: block !important;
        }
        div[data-testid="stDataFrame"], div[data-testid="stTable"] { zoom: 0.65 !important; margin: 0 auto 20px auto !important; max-width: 100% !important; page-break-inside: auto !important; }
        div[data-testid="stTable"] tr { page-break-inside: avoid !important; }
        .element-container { 
            page-break-inside: avoid !important; 
            break-inside: avoid !important; 
            width: 100% !important; 
        }
        .pdf-page-break { break-before: page !important; page-break-before: always !important; height: 0 !important; margin: 0 !important; padding: 0 !important; }
        table { page-break-inside: auto !important; }
        tr { page-break-inside: avoid !important; page-break-after: auto !important; }
        td, th { page-break-inside: avoid !important; }
        thead { display: table-header-group !important; }
    }
    @media print and (orientation: portrait) { .stPlotlyChart { margin-bottom: 10mm !important; } }
    @media print and (orientation: landscape) { .stPlotlyChart { margin-bottom: 6mm !important; } }
    .stPlotlyChart, div[data-testid="stDataFrame"] { display: flex !important; justify-content: center !important; }
    .highlight-blue-box { border: 1.5px solid rgba(0,51,141,0.85) !important; border-radius: 12px !important; padding: 10px !important; background: rgba(0,51,141,0.02) !important; box-shadow: 0px 4px 12px rgba(0,51,141,0.12) !important; margin-bottom: 25px !important; }
    </style>
    <div class="nav-floating-sign" id="custom-nav-trigger">展开导航栏</div>
    """, unsafe_allow_html=True)

    components.html("""<script>let t = setInterval(() => { const d = window.parent.document; const b = d.getElementById("custom-nav-trigger"); const c = d.querySelector('[data-testid="collapsedControl"]') || d.querySelector('button[kind="header"]'); if(b && c) { b.onclick = () => c.click(); clearInterval(t); } }, 500);</script>""", height=0, width=0)
    # ==========================================
    # 加载注释表（导航配置）
    # ==========================================

    st.markdown("<div class='no-print'>", unsafe_allow_html=True)
    with st.expander("📥 公司内容分析与注释输入", expanded=False):
        st.markdown("""
            <div style='background:linear-gradient(135deg,#00338D,#0865EE); 
            border-radius:10px; padding:14px 18px; margin-bottom:16px;
            display:flex; align-items:center; justify-content:space-between;'>
                <span style='color:white; font-size:15px; font-weight:bold;'>
                    获取官方注释表模板
                </span>
                <span style='color:rgba(255,255,255,0.8); font-size:12px;'>
                    需要安全码验证
                </span>
            </div>
        """, unsafe_allow_html=True)
    
        col_pwd, col_btn = st.columns([3, 1])
        with col_pwd:
            pwd_input = st.text_input("", placeholder="请输入安全码...", 
                                       type="password", key="template_pwd",
                                       label_visibility="collapsed")
        with col_btn:
            check_btn = st.button("🔓验证下载", use_container_width=True)
    
        if check_btn:
            if 'filled_notes_excel' in st.session_state:
                del st.session_state['filled_notes_excel']
                
            if pwd_input == "KPMG666":  # ✅ 保持与登录密码一致
                try:
                    import requests
                    from io import BytesIO
                    import openpyxl
        
                    # 🔧 修改点：替换为财险注释表模板的 URL
                    url = "https://github.com/z-xylym/my-actuary-tool/raw/refs/heads/main/step7-%E8%B4%A2%E9%99%A9%E6%A0%87%E6%B3%A8%E8%A1%A8_0806.xlsx"
                    r = requests.get(url, timeout=15)
        
                    if r.status_code == 200:
                        wb = openpyxl.load_workbook(BytesIO(r.content))
                        ws = wb.active
                        header = {cell.value: cell.column for cell in ws[1]}
                        mid_col = header.get('模块ID')
                        custom_col = header.get('分析内容-自定义')
        
                        if mid_col and custom_col and 'integrated_data' in st.session_state:
                            df_for_gen = st.session_state['integrated_data'].copy()
                            _cos = st.session_state.get('selected_cos_cache', [])
                            _valid_years = sorted([y for y in df_for_gen['报告年份'].dropna().astype(str).str.replace(".0","",regex=False).unique() if y.isdigit()])
                            _cy = int(_valid_years[-1]) if _valid_years else 2025
                            _py = int(_valid_years[-2]) if len(_valid_years) > 1 else 2024
        
                            # 遍历 Excel 行，填充分析内容（如果有自定义生成逻辑）
                            for row in ws.iter_rows(min_row=2):
                                m_id_val = row[mid_col - 1].value
                                if not m_id_val: continue
                                # 如果有自动生成话术的需求，可在这里调用
                                # 目前财险系统中暂无 generate_custom_analysis，所以跳过自动填充
                                # 后续如需自动填充，可在此添加逻辑
                                pass
        
                        output = BytesIO()
                        wb.save(output)
                        output.seek(0)
                        
                        st.session_state['filled_notes_excel'] = output.getvalue()
                        st.success("✅ 验证成功，请点击下方按钮下载模板")
                    else:
                        st.error("❌ 文件获取失败")
                except Exception as e:
                    st.error(f"❌ 错误：{e}")
            else:
                st.error("❌ 安全码错误")
        
        # 下载按钮
        if 'filled_notes_excel' in st.session_state:
            st.download_button(
                label="🔓点击下载注释表模板",
                data=st.session_state['filled_notes_excel'],
                file_name=f"注释表_{st.session_state.get('selected_cos_cache', [''])[0] if st.session_state.get('selected_cos_cache') else ''}_{2025}年.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
    
        st.divider()        
        use_default = st.toggle("使用默认注释表", value=True, key="use_default_notes")
        df_notes = None
        if use_default:
            try:
                # 🔧 修改点：替换为财险注释表模板的 URL
                df_notes = pd.read_excel("https://github.com/Polly031021/Annual-Report-System/raw/main/step7-%E8%B4%A2%E9%99%A9%E6%A0%87%E5%87%86%E6%B3%A8%E9%87%8A%E8%A1%A8_0824.xlsx")
                st.success("✅ 内置默认注释表加载成功")
            except Exception as e:
                st.error(f"❌ 加载失败：{e}")
        else:
            st.info("💡 请上传包含【图片文件名】【模块ID】【一级分类】【二级分类】【对应图表名称】【分析内容】【注释内容】的 Excel")
            notes_file = st.file_uploader("上传 Excel", type=['xlsx', 'xls'])
            if notes_file: 
                df_notes = pd.read_excel(notes_file)

        # ---- 解析注释表 ----
        if df_notes is not None:
            # 统一清洗文本前后的空格
            for col in ['一级分类', '二级分类', '对应图表名称', '模块ID']:
                if col in df_notes.columns:
                    df_notes[col] = df_notes[col].astype(str).str.strip().replace(['nan', 'NaN', 'NAN', 'None'], '')
            
            # 如果二级分类为空白，强制塞入"全部"
            if '二级分类' in df_notes.columns:
                df_notes['二级分类'] = df_notes['二级分类'].apply(lambda x: "全部" if str(x).strip() == "" else str(x).strip())
            
            # 填充 notes_dict 和 ordered_modules
            has_img = '图片文件名' in df_notes.columns
            for _, r in df_notes.iterrows():
                m_id = str(r.get('模块ID', '')).strip()
                if not m_id: 
                    continue
                img_val = str(r.get('图片文件名', '')).strip() if has_img and pd.notna(r.get('图片文件名')) else ''
                notes_dict[m_id] = {
                    'title': str(r.get('对应图表名称', '')).strip(),
                    'analysis_default': str(r.get('分析内容-默认', '')).strip() if pd.notna(r.get('分析内容-默认')) else '',
                    'analysis_custom': str(r.get('分析内容-自定义', '')).strip() if pd.notna(r.get('分析内容-自定义')) else '',
                    'note': str(r.get('注释内容', '')).strip() if pd.notna(r.get('注释内容')) else '',
                    'image_file': img_val if img_val.lower() != 'nan' else ''
                }
                if m_id not in ordered_modules:
                    ordered_modules.append(m_id)
            
            # 存入 session_state
            st.session_state['df_notes'] = df_notes
    
    st.markdown("</div>", unsafe_allow_html=True)
    
    # ----- 数据检查与年份提取 -----
    if 'integrated_data' not in st.session_state or st.session_state['integrated_data'] is None:
        # 判断用户角色
        if st.session_state.get('user_role') == "普通用户":
            st.info("📂 请上传已集成好的数据文件（Excel 或 CSV），格式需包含 '公司'、'报告年份'、'字段名'、'(百万)人民币' 等列。")
            
            # 提供示例模板
            with st.expander("📄 查看所需数据格式示例"):
                sample_df = pd.DataFrame({
                    "公司": ["示例公司A", "示例公司A"],
                    "报告年份": [2024, 2025],
                    "字段名": ["总资产", "总资产"],
                    "(百万)人民币": [100000, 110000]
                })
                st.dataframe(sample_df)
                st.caption("请确保你的数据包含以上四列，且每一行对应一个公司、一个年份、一个指标。")
            
            uploaded_file = st.file_uploader("上传集成后的数据表", type=["xlsx", "csv"], key="step7_upload")
            if uploaded_file is not None:
                try:
                    if uploaded_file.name.endswith('.csv'):
                        df = pd.read_csv(uploaded_file)
                    else:
                        df = pd.read_excel(uploaded_file)
                    # 检查必需列
                    required_cols = ['公司', '报告年份', '字段名', '(百万)人民币']
                    if all(col in df.columns for col in required_cols):
                        # 将年份列转换为字符串并去除 .0
                        df['报告年份'] = df['报告年份'].astype(str).str.replace('.0', '', regex=False)
                        st.session_state['integrated_data'] = df
                        st.success("✅ 数据加载成功！正在刷新...")
                        st.rerun()
                    else:
                        st.error(f"❌ 数据格式不正确，必须包含列：{required_cols}，你的数据列有：{df.columns.tolist()}")
                except Exception as e:
                    st.error(f"❌ 读取文件失败：{e}")
            # 阻止继续执行报告渲染（因为数据还不存在）
            st.stop()
        else:
            # 项目组成员：提示去 Step 6
            st.warning("⚠️ 请先在 Step 6 完成数据集成。")
            st.stop()
    df_raw = st.session_state['integrated_data'].copy()
    valid_years = sorted([y for y in df_raw['报告年份'].dropna().astype(str).str.replace(".0", "", regex=False).unique() if y.isdigit()])
    latest_year = int(valid_years[-1]) if valid_years else 2025
    prev_year = int(valid_years[-2]) if len(valid_years) > 1 else 2023

    # ----- 侧边栏导航 -----
    print_mode, active_m_id = False, None
    with st.sidebar:
        st.markdown("<h3 style='color: #00338D; font-size: 18px;'>公司报告导航</h3>", unsafe_allow_html=True)
        if 'df_notes' in st.session_state and st.session_state['df_notes'] is not None and not st.session_state['df_notes'].empty:
            df_n = st.session_state['df_notes']
            for _, r in df_n.iterrows():
                m_id = str(r.get('模块ID', '')).strip()
                if m_id:
                    notes_dict[m_id] = {
                        'title': str(r.get('对应图表名称', '')).strip(),
                        'analysis_default': str(r.get('分析内容-默认', '')).strip() if pd.notna(r.get('分析内容-默认')) else '',
                        'analysis_custom': str(r.get('分析内容-自定义', '')).strip() if pd.notna(r.get('分析内容-自定义')) else '',
                        'note': str(r.get('注释内容', '')).strip() if pd.notna(r.get('注释内容')) else '',
                        'image_file': str(r.get('图片文件名', '')).strip() if pd.notna(r.get('图片文件名')) else ''
                    }
                    if m_id not in ordered_modules:
                        ordered_modules.append(m_id)
            first_levels = [x for x in df_n['一级分类'].unique() if str(x).strip() != ""]
            main_nav = st.radio("📁 一级模块", first_levels + ["🖨️ 一键显示全部 (打印/导出)"], key="s7_m")

            if main_nav == "🖨️ 一键显示全部 (打印/导出)":
                print_mode = True
                if not st.session_state.get("pdf_export_mode", False):
                    st.info('竖版适合文字多的页面，横版适合宽图表。勾选"背景图形"以保留颜色。')
                    components.html("""
                    <div style="display:flex; flex-direction:column; gap:8px;">
                        <button onclick="printAs('portrait')" style="width:100%; padding:11px; background:#00338D; color:white; border:none; border-radius:6px; cursor:pointer; font-weight:bold; font-size:13px;">
                            🖨️ 导出竖版 A4 PDF
                        </button>
                        <button onclick="printAs('widescreen')" style="width:100%; padding:11px; background:#008578; color:white; border:none; border-radius:6px; cursor:pointer; font-weight:bold; font-size:13px;">
                            🖨️ 导出横版 16:9 PDF
                        </button>
                    </div>
                    <script>
                    function printAs(mode) {
                        const doc = window.parent.document;
                    
                        // 移除旧打印样式
                        const old = doc.getElementById('dynamic-print-style');
                        if (old) old.remove();
                    
                        const style = doc.createElement('style');
                        style.id = 'dynamic-print-style';
                    
                        if (mode === 'widescreen') {
                            style.innerHTML = `
                                @page {
                                    size: landscape;
                                    margin: 5mm 8mm;
                                }
                    
                                @page {
                                    -webkit-size: landscape;
                                }
                    
                                .main .block-container {
                                    width: 100% !important;
                                    max-width: 100% !important;
                                }
                            `;
                        } else {
                            style.innerHTML = `
                                @page {
                                    size: A4 portrait;
                                    margin: 10mm;
                                }
                    
                                @page {
                                    -webkit-size: A4 portrait;
                                }
                            `;
                        }
                    
                        doc.head.appendChild(style);
                    
                        // 强制重新计算布局
                        window.parent.getComputedStyle(doc.body).width;
                    
                        // ==========================================
                        // 等待 Plotly 图表完成渲染
                        // ==========================================
                        const startTime = Date.now();
                    
                        function checkChartsReady() {
                    
                            const charts = Array.from(
                                doc.querySelectorAll('.plotly-graph-div')
                            );
                    
                            // 检查所有 Plotly 图表
                            const allReady =
                                charts.length === 0 ||
                                charts.every(chart => {
                    
                                    const svg = chart.querySelector('.main-svg');
                                    const rect = chart.getBoundingClientRect();
                    
                                    return (
                                        svg &&
                                        rect.width > 100 &&
                                        rect.height > 100
                                    );
                                });
                    
                            // 图表全部完成
                            if (allReady) {
                    
                                // 再给浏览器两帧时间完成最终布局
                                requestAnimationFrame(() => {
                                    requestAnimationFrame(() => {
                    
                                        setTimeout(() => {
                                            window.parent.print();
                                        }, 300);
                    
                                    });
                                });
                    
                                return;
                            }
                    
                            // 最多等待 8 秒
                            if (Date.now() - startTime > 8000) {
                    
                                console.warn(
                                    'Plotly 图表等待超过 8 秒，继续打印'
                                );
                    
                                window.parent.print();
                                return;
                            }
                    
                            // 100ms 后再次检查
                            setTimeout(checkChartsReady, 100);
                        }
                    
                        checkChartsReady();
                    }
                    </script>
                    """, height=100)
            else:
                df_sub1 = df_n[df_n['一级分类'] == main_nav]
                raw_sec_levels = [x for x in df_sub1['二级分类'].unique() if x != "全部"]
                if len(raw_sec_levels) == 0:
                    charts = [x for x in df_sub1['对应图表名称'].unique() if x]
                    chart_nav = st.radio("具体图表", charts, key="s7_c")
                    matched = df_sub1[df_sub1['对应图表名称'] == chart_nav]
                    active_m_id = matched.iloc[0]['模块ID'] if not matched.empty else None
                else:
                    sub_nav = st.radio("📂 二级模块", ["全部"] + raw_sec_levels, key="s7_s")
                    if sub_nav != "全部":
                        df_sub2 = df_sub1[df_sub1['二级分类'] == sub_nav]
                        charts = [x for x in df_sub2['对应图表名称'].unique() if x]
                        chart_nav = st.radio("具体图表", charts, key="s7_c")
                        matched = df_sub2[df_sub2['对应图表名称'] == chart_nav]
                        active_m_id = matched.iloc[0]['模块ID'] if not matched.empty else None
                    else:
                        charts = [x for x in df_sub1['对应图表名称'].unique() if x]
                        chart_nav = st.radio("具体图表 (当前二级：全部)", charts, key="s7_c_all")
                        matched = df_sub1[df_sub1['对应图表名称'] == chart_nav]
                        active_m_id = matched.iloc[0]['模块ID'] if not matched.empty else None
        else:
            st.warning("⚠️ 请先加载包含层级信息的注释表")

    # ----- 配置与图片上传 -----
    st.markdown("<div class='no-print'>", unsafe_allow_html=True)
    with st.expander("⚙️ 公司级图表设置与图片覆盖", expanded=False):
        c0, c1, c2, c3, c4 = st.columns([1, 2, 1, 1, 1])   
        with c0:
            type_options = sorted([
                str(x).strip() for x in df_raw["公司类型"].dropna().unique()
                if str(x).strip() not in ["", "nan", "None"]
            ])
            if "s7_company_types" not in st.session_state:
                st.session_state["s7_company_types"] = ["全部"]
            selected_types_raw = st.multiselect(
                "公司类型",
                options=["全部"] + type_options,
                default=st.session_state["s7_company_types"],
                key="s7_company_types"
            )
            if (not selected_types_raw) or ("全部" in selected_types_raw):
                selected_types = ["全部"]
                df_filtered = df_raw.copy()
                st.write("df_filtered 行数:", len(df_filtered))
                st.dataframe(df_filtered.head(3))
            else:
                selected_types = selected_types_raw
                df_filtered = df_raw[df_raw["公司类型"].astype(str).str.strip().isin(selected_types)].copy()
        with c1: 
            raw_ordered_cos = list(dict.fromkeys(df_filtered['公司'].dropna().tolist()))
            selected_cos = st.multiselect("展示公司", options=raw_ordered_cos, default=raw_ordered_cos)
        with c2: 
            unit_label = st.selectbox("显示单位", ["十亿元", "亿元", "百万元", "十万元"])
            divisor = {"十亿元": 1000, "亿元": 100, "百万元": 1, "十万元": 0.1}[unit_label]
        with c3: highlight_co = st.selectbox("特定追踪", ["无"] + selected_cos)
        with c4: enable_ai = st.toggle("一键AI分析", value=False)
        
        st.markdown("<hr style='margin: 5px 0 15px 0;'>", unsafe_allow_html=True)
        sc1, sc2, sc3 = st.columns([2, 2, 3])
        with sc1:
            available_fields = sorted([str(x) for x in df_filtered['字段名'].unique() if str(x).strip() not in ['nan', 'None', '']])
            sort_field = st.selectbox("📊 图表展示顺序依据", ["默认（按列表原始顺序）"] + available_fields)
        with sc2:
            sort_order = st.radio("排序方向", ["降序 (从大到小) ⬇️", "升序 (从小到大) ⬆️"], horizontal=True)
        # ----- 统一公司排序逻辑 -----
        # 如果用户选择“默认”，则按保险服务收入降序排列（与“新准则保险服务收入排名”一致）
        if sort_field == "默认（按列表原始顺序）":
            # 使用保险服务收入作为排序依据
            sort_field_actual = "保险服务收入"
            is_desc = True  # 降序
            if sort_field_actual in available_fields:
                df_sort = df_filtered[(df_filtered['字段名'] == sort_field_actual) & (df_filtered['报告年份'].astype(str) == str(latest_year))]
                val_col = "(百万)人民币" if "(百万)人民币" in df_sort.columns else df_sort.columns[-1]
                sort_map = {}
                for _, r in df_sort.iterrows():
                    val = pd.to_numeric(r.get(val_col, np.nan), errors='coerce')
                    sort_map[str(r['公司']).strip()] = val
                default_val = float('-inf') if is_desc else float('inf')
                selected_cos = sorted(selected_cos, key=lambda x: sort_map.get(x) if pd.notna(sort_map.get(x)) else default_val, reverse=is_desc)
            else:
                # 若没有保险服务收入，则按公司名称排序
                selected_cos = sorted(selected_cos)
        elif sort_field != "默认（按列表原始顺序）" and selected_cos:
            df_sort = df_filtered[(df_filtered['字段名'] == sort_field) & (df_filtered['报告年份'].astype(str) == str(latest_year))]
            val_col = "(百万)人民币" if "(百万)人民币" in df_sort.columns else df_sort.columns[-1]
            is_desc = "降序" in sort_order
            sort_map = {}
            for _, r in df_sort.iterrows():
                val = pd.to_numeric(r.get(val_col, np.nan), errors='coerce')
                sort_map[str(r['公司']).strip()] = val
            default_val = float('-inf') if is_desc else float('inf')
            selected_cos = sorted(selected_cos, key=lambda x: sort_map.get(x) if pd.notna(sort_map.get(x)) else default_val, reverse=is_desc)
        st.session_state['selected_cos_cache'] = selected_cos
        global HL_BOX_FILL, HL_BOX_LINE
        HL_BOX_FILL = "rgba(0,51,141,0.08)"   # 适当加深背景
        HL_BOX_LINE = "rgba(0,51,141,0.8)"    # 加深边框
                
        st.markdown("---")
        st.caption("📸 手动上传图片（png或jpg）")
        if 'manual_upload_images' not in st.session_state:
            st.session_state.manual_upload_images = {}
        uploaded_files = st.file_uploader("拖入截图文件", type=['png', 'jpg'], accept_multiple_files=True)

        if uploaded_files and ordered_modules:
            cols = st.columns(2)
            for i, file in enumerate(uploaded_files):
                with cols[i % 2]:
                    def get_name(m):
                        if m == "不匹配/跳过":
                            return "不匹配/跳过"
                        return notes_dict.get(m, {}).get('title', m)
                    sel_mid = st.selectbox(
                        f"图片 {file.name} 对应：",
                        options=["不匹配/跳过"] + ordered_modules,
                        format_func=get_name,
                        key=f"up_{i}"
                    )
                    if sel_mid != "不匹配/跳过":
                        st.session_state.manual_upload_images[sel_mid] = file
                    st.image(file, use_column_width=True)
    st.markdown("</div>", unsafe_allow_html=True)


    # ----- 获取配置参数 -----
    st.session_state['df_filtered'] = df_filtered
    st.session_state['selected_cos_cache'] = selected_cos
    st.session_state['divisor'] = divisor
    st.session_state['latest_year'] = latest_year
    st.session_state['highlight_co'] = highlight_co
    st.session_state['enable_ai'] = enable_ai
    st.session_state['unit_label'] = unit_label
    st.session_state['print_mode'] = print_mode
    st.session_state['active_m_id'] = active_m_id
    st.session_state['prev_year'] = prev_year
    st.session_state['divisor'] = divisor
    st.session_state['HL_BOX_FILL'] = HL_BOX_FILL
    st.session_state['HL_BOX_LINE'] = HL_BOX_LINE
    st.session_state['highlight_co'] = highlight_co

    if not selected_cos:
        selected_cos = list(df_filtered['公司'].unique())
        st.session_state['selected_cos_cache'] = selected_cos

    st.markdown("<h3 class='no-print' style='font-weight:700;'>📊 公司级对标报告</h3>", unsafe_allow_html=True)

    # ----- 打印模式（一键显示全部）----- 
    if print_mode:
        # ===== 🆕 封面页（使用 Base64 内嵌图片） =====
        # 将你复制的 Base64 代码粘贴到下面（替换整个字符串）
        cover_image_base64 = "data:image/jpeg;base64,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xSUm6is3IdhaKTPFJUXGLxTqZT6QBTqbTt1SwClWkpRiloA6lWm5FOVqkYYFOFJuFKMUALxThikwKVRQAcU5aTbTlWkMXIpaTaKdtFAg2+9OxSbRTttILiYp1FKKdhgFNOxRQTQAoFG2haSsxiilpFpakBy0tItLTAeop1MWn0AOApaaKVqAFpy00U4dKAFpVpKUUgFpy02lHWkOw6lWkpVp3EOpVpKFouA6lFJS0rjQ6lWkFKtADqdTadSActLTVp1AxR1paKWgBacoplOHSgY7GKVetN4p69aAHUu6koqxjweKXdTRS0gFp60ylVjTAfS/hSUZoAdS0imlpDHUUUUgHL0paRelLQAopabRTGPWlpFNLkVIhaKTNLQUgp47Uynr2qkJi80tFLTAKX0pKMUxiinDNIKeKdhCcilpeaOTVWAbSgUvIpcUrAFO20mDTqaATiloopggpR0pKctAw20m33p1JVALtopwoppAJ+FLS4oxVpCuJRS0lMQq06kWnAUWAPSlo20oFWgFpabzS5GOasBeKOKOKVWHpVJCG8UCncdhRj2osAnrTafj2pMD0osMT60lLj2pNppgKKMUbTS81IBTlpNppyofWtEmIKKdtpOa0UQuAp1Ax/dp24dhWqiSR7TRineZ7UufarshajDRTqMGjlEMop2088Zo2n0q7DG0jdKXac9aCtSxjR707jtSbTTljNNXAcKUUBSKTFbxJHbj6fpTt59P0pm335+tLz/k1rdkNIduNG403mjdV3FYccmkzjpSbqKoVh24+lHNN5oyc1V2FkG05p2Kbk0pzVoBcUhWlXPpQa0JYzb70u33opwpi1G7c07yxxQM07PStEhajGjHrTfL9KlZhTS3FWF2REUvFK1NJxQAppuD7Ucf5FLtPrQAjUlOI980HpVodxtL+FJT+a1RJGVOelLtOOlSbfYUFCegFUK5Hz3op2089KNvqRVCGmmN3qXb75pu0c0ikRUtLsNHlmqGxtKKXyzSY29TQriDgGilKjijir1ENNG7608gHtQFHpVWAb+lHrT9ntRt46U7MQ2jv2pQvqcUbB6/pWiTEJt60eX704J15o24qkhCCOjA74/Onfhml3f7PNUDI9o9aKd/wGgr6VQg4xSUvln0pNp9KroA7imNjFPwcdDSbTg/LVIQxcelLx6U4Rn0x+NGw1Yhp25pNw7U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"
    
        st.markdown(f"""
        <style>
        .cover-page {{
            width: 100%;
            height: 100vh;
            background-image: url('{cover_image_base64}');
            background-size: cover;
            background-position: center;
            position: relative;
            page-break-after: always;
            display: flex;
            flex-direction: column;
            justify-content: center;
            align-items: center;
            text-align: center;
            color: white;
            font-family: "Microsoft YaHei", "PingFang SC", sans-serif;
            margin: 0;
            padding: 0;
            -webkit-print-color-adjust: exact;
            print-color-adjust: exact;
        }}
        .cover-page::before {{
            content: "";
            position: absolute;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            background: rgba(0, 0, 0, 0.3);
            z-index: 1;
        }}
        .cover-title {{
            font-size: 56px;
            font-weight: 700;
            letter-spacing: 4px;
            text-shadow: 2px 2px 12px rgba(0,0,0,0.6);
            z-index: 2;
            margin-bottom: 20px;
        }}
        .cover-subtitle {{
            font-size: 28px;
            font-weight: 300;
            letter-spacing: 8px;
            text-shadow: 1px 1px 8px rgba(0,0,0,0.5);
            z-index: 2;
            margin-bottom: 30px;
        }}
        .cover-date {{
            font-size: 20px;
            font-weight: 300;
            letter-spacing: 6px;
            opacity: 0.9;
            text-shadow: 1px 1px 4px rgba(0,0,0,0.5);
            z-index: 2;
        }}
        .cover-page > * {{
            position: relative;
            z-index: 2;
        }}
        @media print {{
            .cover-page {{
                height: 100%;
                width: 100%;
                position: relative;
                margin: 0;
                padding: 0;
                page-break-after: always;
            }}
            body, html {{
                margin: 0;
                padding: 0;
            }}
        }}
        </style>
        <div class="cover-page">
            <div class="cover-title">2025年新会计准则业绩表现和洞察</div>
            <div class="cover-subtitle">保险公司</div>
            <div class="cover-date">2026年8月</div>
        </div>
        """, unsafe_allow_html=True)
        
        # 然后渲染所有模块
        for idx, m_id in enumerate(ordered_modules):
            render_report_module(m_id, print_mode, is_first=(idx == 0))
        return
        return

    # ----- 单模块模式 -----
    if not active_m_id:
        st.info("请在左侧导航栏中选择一个具体图表模块。")
        return

    render_report_module(active_m_id, print_mode, is_first=True)
    return
    
# 7.1 路由引擎
def render_pure_chart_entity(m_id, print_mode):
    global notes_dict
    
    # ====== 先读取所有必要的 session_state 变量 ======
    df_filtered = st.session_state.get('df_filtered', pd.DataFrame())
    selected_cos = st.session_state.get('selected_cos_cache', [])
    latest_year = st.session_state.get('latest_year', 2025)
    prev_year = st.session_state.get('prev_year', 2024)
    divisor = st.session_state.get('divisor', 1)
    unit_label = st.session_state.get('unit_label', "百万元")
    highlight_co = st.session_state.get('highlight_co', "无")
    global HL_BOX_FILL, HL_BOX_LINE
    HL_BOX_FILL = st.session_state.get('HL_BOX_FILL', "rgba(0,51,141,0.03)")
    HL_BOX_LINE = st.session_state.get('HL_BOX_LINE', "rgba(0,51,141,0.35)")
    current_hl = highlight_co if highlight_co else "无"
    
    # ====== 概览模块 ======
    if m_id in ["概览", "overview"]:
        fig = create_overview_table(df_filtered, selected_cos, latest_year, prev_year, unit_label)
        # 替换原来的 st.plotly_chart
        show_chart(fig, print_mode, m_id)
        display_notes(m_id, df_filtered, "概览")
        display_bottom_note(notes_dict.get(m_id, {}).get('note', ''))
        return
    
    # ==========================================
    # 1. 文本类披露（已在主调度中单独处理，此处留空以防误调用）
    # ==========================================
    if m_id in ["policy_method", "discount_rate", "risk_margin"]:
        display_textual_disclosures(df_filtered, selected_cos, latest_year)
        display_notes(m_id, df_filtered, "")
        display_bottom_note(notes_dict.get(m_id, {}).get('note', ''))
        return

    # ==========================================
    # 2. 单指标柱状图（两年对比 + 增长率标注）
    # ==========================================
    single_metric_map = {
        "premium_ranking": ("保险服务收入", True, False),
        "new_old_ratio": ("新旧准则比值", False, True),
        "investment_component": ("投资成分占比", False, True),    # ← 改为“投资成分占比”
        "premium_growth": ("保费增长率（旧准则）", False, True),  # ← 改为“保费增长率（旧准则）”
        "loss_ratio": ("综合赔付率", False, True),
        "expense_ratio": ("综合费用率", False, True),
        "loss_component": ("亏损成分占比", False, True),
        "backtest_deviation": ("回溯偏差率", False, True),
    }

    # ==========================================
    # 综合成本率变化 + 综合费用率/赔付率 breakdown
    # ==========================================
    if m_id == "cor_trend":
        if not print_mode:
            c1, c2, c3 = st.columns(3)
            with c1:
                show_labels = st.toggle("显示标签", True, key=f"lab_{m_id}")
            with c2:
                pct_sz = st.slider("涨幅字号", 8, 24, 14, key=f"psz_{m_id}")   # 不再使用
            with c3:
                gap = st.slider("柱间距", 0.1, 0.8, 0.15, key=f"gap_{m_id}")   # 不再使用
        else:
            show_labels = st.session_state.get(f"lab_{m_id}", True)
            pct_sz = st.session_state.get(f"psz_{m_id}", 14)
            gap = st.session_state.get(f"gap_{m_id}", 0.15)
        
        # ---- 只保留第二张图：综合费用率与赔付率 breakdown ----
        raw_df2 = df_filtered.copy()
        cos_list = st.session_state.get('selected_cos_cache', [])
        prev_y = st.session_state.get('prev_year', 2024)
        latest_y = st.session_state.get('latest_year', 2025)
        years = [str(prev_y), str(latest_y)]
        
        # 提取数据
        metrics = ["综合赔付率", "综合费用率"]
        data_dict = {}
        for co in cos_list:
            data_dict[co] = {}
            for yr in years:
                data_dict[co][yr] = {}
                for m in metrics:
                    val = raw_df2[
                        (raw_df2['公司'].astype(str).str.strip() == str(co).strip()) &
                        (raw_df2['报告年份'].astype(str).str.replace('.0', '', regex=False).str.strip() == yr) &
                        (raw_df2['字段名'].astype(str).str.strip() == m)
                    ]['(百万)人民币']
                    if not val.empty:
                        num_val = pd.to_numeric(val.iloc[0], errors='coerce')
                        data_dict[co][yr][m] = 0 if pd.isna(num_val) else float(num_val)
                    else:
                        data_dict[co][yr][m] = 0
        
        # 构造 x 轴位置和数值
        x_vals = []
        y_exp = []   # 费用率
        y_loss = []  # 赔付率
        total_vals = []  # 综合成本率
        for i, co in enumerate(cos_list):
            base_idx = 2 * i
            # 2024
            x_vals.append(base_idx)
            exp = data_dict[co][str(prev_y)]["综合费用率"]
            loss = data_dict[co][str(prev_y)]["综合赔付率"]
            y_exp.append(exp)
            y_loss.append(loss)
            total_vals.append(exp + loss)
            # 2025
            x_vals.append(base_idx + 1)
            exp = data_dict[co][str(latest_y)]["综合费用率"]
            loss = data_dict[co][str(latest_y)]["综合赔付率"]
            y_exp.append(exp)
            y_loss.append(loss)
            total_vals.append(exp + loss)
        
        # 年份标签（每个柱子对应一个年份）
        year_labels = []
        for _ in cos_list:
            year_labels.append("2024YE")
            year_labels.append("2025YE")
        
        fig2 = go.Figure()
        fig2.add_trace(go.Bar(
            name="综合费用率",
            x=x_vals,
            y=y_exp,
            marker_color="#1E49E2",
            text=[f"{v:.1%}" if show_labels and v != 0 else "" for v in y_exp],
            textposition='inside',
            insidetextanchor='middle',
            textfont=dict(color="white", size=11),
            width=0.4,
        ))
        fig2.add_trace(go.Bar(
            name="综合赔付率",
            x=x_vals,
            y=y_loss,
            marker_color="#00338D",
            text=[f"{v:.1%}" if show_labels and v != 0 else "" for v in y_loss],
            textposition='inside',
            insidetextanchor='middle',
            textfont=dict(color="white", size=11),
            width=0.4,
        ))
        
        # 添加综合成本率数值（柱顶）
        for i, (x, total) in enumerate(zip(x_vals, total_vals)):
            fig2.add_annotation(
                x=x,
                y=total,
                text=f"{total:.1%}",
                showarrow=False,
                font=dict(size=12, color="#333"),
                yshift=5,
            )
        
        fig2.update_layout(
            barmode='relative',
            title="综合费用率与赔付率 breakdown",
            xaxis=dict(
                tickvals=x_vals,
                ticktext=year_labels,
                tickangle=-30,
                tickfont=dict(size=11, family="Microsoft YaHei", color="#333", style="normal"),
            ),
            yaxis=dict(title="占保险服务收入比例", tickformat=".1%"),
            height=400,
            bargap=0.15,
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="center", x=0.5),
            plot_bgcolor='rgba(0,0,0,0)',
            paper_bgcolor='rgba(0,0,0,0)',
        )
        
        # 添加公司名 annotations（每组柱子中间下方）
        for i, co in enumerate(cos_list):
            fig2.add_annotation(
                x=2*i + 0.5,
                y=-0.08,  # 相对于图表高度的比例，放在 x 轴标签下方
                text=co,
                showarrow=False,
                font=dict(size=12, family="Microsoft YaHei", color="#333", style="normal"),
                xref='x',
                yref='paper',
                yanchor='top',
                xanchor='center'
            )
        
        show_chart(fig2, print_mode, m_id)
        
        display_notes(m_id, df_filtered, "综合成本率")
        display_bottom_note(notes_dict.get(m_id, {}).get('note', ''))
        return

    
    if m_id in single_metric_map:
        field_name, sort_by, is_pct = single_metric_map[m_id]
        
        # ----- 获取 UI 控件值（必须保留）-----
        if not print_mode:
            c1, c2, c3 = st.columns(3)
            with c1:
                show_labels = st.toggle("显示标签", True, key=f"lab_{m_id}")
            with c2:
                pct_sz = st.slider("涨幅字号", 8, 24, 14, key=f"psz_{m_id}")
            with c3:
                gap = st.slider("柱间距", 0.1, 0.8, 0.15, key=f"gap_{m_id}")
        else:
            show_labels = st.session_state.get(f"lab_{m_id}", True)
            pct_sz = st.session_state.get(f"psz_{m_id}", 14)
            gap = st.session_state.get(f"gap_{m_id}", 0.15)
        
        # ----- 对“保费分析”指标强制限定“保费收入”类别 -----
        premium_analysis_ids = ["premium_ranking", "new_old_ratio", "investment_component", "premium_growth"]
        if m_id in premium_analysis_ids:
            if '类别' in df_filtered.columns:
                df_plot = df_filtered[df_filtered['类别'] == "保费收入"].copy()
            else:
                df_plot = df_filtered.copy()
            if df_plot.empty:
                df_plot = df_filtered.copy()
        else:
            df_plot = df_filtered.copy()
        
        fig = create_kpmg_chart(df_plot, field_name, "", show_labels, pct_sz, gap, 
                                sort_by_value=sort_by, is_percentage=is_pct)
        show_chart(fig, print_mode, m_id)
        display_notes(m_id, df_filtered, field_name, is_pct)
        display_bottom_note(notes_dict.get(m_id, {}).get('note', ''))
        return

    # ==========================================
    # 3. 综合成本率拆解（多因子分组柱状图）
    # ==========================================
    if m_id == "cor_components":
        fig = create_cor_breakdown_stacked_chart(
            df_filtered, selected_cos, latest_year, divisor, unit_label, current_hl
        )
        show_chart(fig, print_mode, m_id)
        display_notes(m_id, df_filtered, "综合赔付率")
        display_bottom_note(notes_dict.get(m_id, {}).get('note', ''))
        return
    # ==========================================
    # 4. 费用分类构成（使用通用多分类引擎）
    # ==========================================
    if m_id == "expense_classification":
        if not print_mode:
            c1, c2, c3 = st.columns(3)
            with c1:
                show_labels = st.toggle("显示标签", True, key=f"lab_{m_id}")
            with c2:
                label_size = st.slider("标签字号", 8, 20, 12, key=f"lsz_{m_id}")
            with c3:
                bar_width = st.slider("柱宽", 0.2, 1.0, 0.6, key=f"wid_{m_id}")
        else:
            show_labels = st.session_state.get(f"lab_{m_id}", True)
            label_size = st.session_state.get(f"lsz_{m_id}", 12)
            bar_width = st.session_state.get(f"wid_{m_id}", 0.6)
        
        # 只显示最近两年
        prev_year = st.session_state.get('prev_year', 2024)
        latest_year = st.session_state.get('latest_year', 2025)
        years_to_show = [str(prev_year), str(latest_year)]
        
        field_map = {
            "获取费用占比": "获取费用",
            "维持费用占比": "维持费用",
            "非履约费用占比": "非履约费用",
            "履约费用占比": "履约费用"
        }
        color_map = {
            "获取费用占比": KPMG_COLORS[0],  # 深蓝
            "维持费用占比": KPMG_COLORS[1],  # 亮蓝
            "非履约费用占比": KPMG_COLORS[10], # 橙红
            "履约费用占比": KPMG_COLORS[5],  # 深紫
        }

        fig, df_avg = create_kpmg_multi_composition_chart(
            df_filtered, field_map, color_map, "",
            show_labels=show_labels, label_size=label_size, bar_width=bar_width, 
            co_font_size=13, highlight_co=current_hl, add_zero_line=False,
            years_to_show=years_to_show
        )
        show_chart(fig, print_mode, m_id)
        display_notes(m_id, df_filtered, "费用分类")
        display_bottom_note(notes_dict.get(m_id, {}).get('note', ''))
        return

    # ==========================================
    # 5. 税务分析：税前/净利润及有效税率
    # ==========================================
    if m_id == "tax_analysis":
        if not print_mode:
            c1, c2, c3, c4, c5 = st.columns(5)
            with c1:
                show_labels = st.toggle("显示标签", True, key=f"lab_{m_id}")
            with c2:
                bar_width = st.slider("柱宽", 0.1, 0.8, 0.35, key=f"wid_{m_id}")
            with c3:
                label_size = st.slider("字号", 8, 16, 12, key=f"lsz_{m_id}")
            with c4:
                co_font_size = st.slider("公司名字号", 10, 24, 14, key=f"cfs_{m_id}")
            with c5:
                co_y_offset = st.slider("标题偏移", 1.0, 1.2, 1.08, step=0.01, key=f"off_{m_id}")
        else:
            show_labels = st.session_state.get(f"lab_{m_id}", True)
            bar_width = st.session_state.get(f"wid_{m_id}", 0.35)
            label_size = st.session_state.get(f"lsz_{m_id}", 12)
            co_font_size = st.session_state.get(f"cfs_{m_id}", 14)
            co_y_offset = st.session_state.get(f"off_{m_id}", 1.08)
        
        # 数据预处理
        df_tax_sub = df_filtered[(df_filtered['字段名'].isin(['税前利润总额', '净利润'])) & (df_filtered['公司'].isin(selected_cos))].drop_duplicates(subset=['公司', '报告年份', '字段名']).copy()
        if not df_tax_sub.empty:
            df_tax_pivot = df_tax_sub.pivot_table(index=['公司', '报告年份'], columns='字段名', values='(百万)人民币').fillna(0).reset_index()
            denom = df_tax_pivot['税前利润总额'].replace(0, np.nan)
            df_tax_pivot['有效税率'] = np.where(
                df_tax_pivot['税前利润总额'] != 0,
                (df_tax_pivot['税前利润总额'] - df_tax_pivot['净利润']) / denom,
                0
            )
            df_tax_pivot['报告年份'] = df_tax_pivot['报告年份'].astype(str).str.replace(".0", "", regex=False) + "YE"
            fig = create_tax_subplot_chart(
                df_tax_pivot, selected_cos, 
                show_labels=show_labels, bar_width=bar_width, label_size=label_size, 
                co_font_size=co_font_size, co_y_offset=co_y_offset, highlight_co=current_hl
            )
        else:
            fig = go.Figure()
            fig.add_annotation(text="无税前/净利润数据", x=0.5, y=0.5, showarrow=False)
        show_chart(fig, print_mode, m_id)
        display_notes(m_id, df_filtered, "税前利润")
        display_bottom_note(notes_dict.get(m_id, {}).get('note', ''))
        return

    # ==========================================
    # 6. 净资产变动明细表（纯 HTML 表格）
    # ==========================================
    if m_id == "equity_change_detail":
        html_table = create_equity_change_detail_table(
            df_raw=df_filtered,
            target_year=latest_year,
            cos=selected_cos,
            divisor=divisor,
            unit_label=unit_label,
            highlight_co=current_hl
        )
        st.markdown(html_table, unsafe_allow_html=True)
        display_notes(m_id, df_filtered, "净资产")
        display_bottom_note(notes_dict.get(m_id, {}).get('note', ''))
        return

    # ==========================================
    # 7. 总资产 / 净资产趋势（带明细表）
    # ==========================================
    if m_id == "asset_trend":
        if not print_mode:
            c1, c2, c3 = st.columns(3)
            with c1:
                show_labels = st.toggle("显示标签", True, key=f"lab_{m_id}")
            with c2:
                pct_sz = st.slider("涨幅字号", 8, 24, 14, key=f"psz_{m_id}")
            with c3:
                gap = st.slider("柱间距", 0.1, 0.8, 0.15, key=f"gap_{m_id}")
        else:
            show_labels = st.session_state.get(f"lab_{m_id}", True)
            pct_sz = st.session_state.get(f"psz_{m_id}", 14)
            gap = st.session_state.get(f"gap_{m_id}", 0.15)
        
        fig = create_kpmg_chart(df_filtered, "总资产", "", show_labels, pct_sz, gap, 
                                sort_by_value=False, is_percentage=False)
        show_chart(fig, print_mode, m_id)
        display_notes(m_id, df_filtered, "总资产")
        display_bottom_note(notes_dict.get(m_id, {}).get('note', ''))
        return

    if m_id == "equity_trend":
        if not print_mode:
            c1, c2, c3 = st.columns(3)
            with c1:
                show_labels = st.toggle("显示标签", True, key=f"lab_{m_id}")
            with c2:
                pct_sz = st.slider("涨幅字号", 8, 24, 14, key=f"psz_{m_id}")
            with c3:
                gap = st.slider("柱间距", 0.1, 0.8, 0.15, key=f"gap_{m_id}")
        else:
            show_labels = st.session_state.get(f"lab_{m_id}", True)
            pct_sz = st.session_state.get(f"psz_{m_id}", 14)
            gap = st.session_state.get(f"gap_{m_id}", 0.15)
        
        fig = create_kpmg_chart(df_filtered, "期末股东权益", "", show_labels, pct_sz, gap, 
                                sort_by_value=False, is_percentage=False)
        show_chart(fig, print_mode, m_id)
        # 下方补充净资产变动明细表
        html_table = create_equity_change_detail_table(
            df_raw=df_filtered,
            target_year=latest_year,
            cos=selected_cos,
            divisor=divisor,
            unit_label=unit_label,
            highlight_co=current_hl
        )
        st.markdown(html_table, unsafe_allow_html=True)
        display_notes(m_id, df_filtered, "净资产")
        display_bottom_note(notes_dict.get(m_id, {}).get('note', ''))
        return

    # ==========================================
    # 8. 费用结构拆解图（堆叠柱状图）
    # ==========================================
    if m_id == "expense_structure":
        if not print_mode:
            c1, c2, c3 = st.columns(3)
            with c1:
                show_labels = st.toggle("显示标签", True, key=f"lab_{m_id}")
            with c2:
                label_size = st.slider("标签字号", 8, 20, 12, key=f"lsz_{m_id}")
            with c3:
                bar_width = st.slider("柱宽", 0.2, 1.0, 0.5, key=f"wid_{m_id}")
        else:
            show_labels = st.session_state.get(f"lab_{m_id}", True)
            label_size = st.session_state.get(f"lsz_{m_id}", 12)
            bar_width = st.session_state.get(f"wid_{m_id}", 0.5)
        
        fig, _ = create_expense_structure_chart(
            df_filtered, selected_cos,
            show_labels=show_labels, label_size=label_size, bar_width=bar_width, 
            co_font_size=13, highlight_co=current_hl
        )
        show_chart(fig, print_mode, m_id)
        display_notes(m_id, df_filtered, "费用")
        display_bottom_note(notes_dict.get(m_id, {}).get('note', ''))
        return

    # ==========================================
    # 9. 六维度散点图矩阵
    # ==========================================
    if m_id == "six_dimensions":
        if not print_mode:
            c1, c2, c3 = st.columns(3)
            with c1:
                show_labels = st.toggle("显示标签", True, key=f"lab_{m_id}")
            with c2:
                label_size = st.slider("字号", 8, 16, 11, key=f"lsz_{m_id}")
            with c3:
                dot_size = st.slider("点大小", 4, 16, 11, key=f"dsz_{m_id}")
        else:
            show_labels = st.session_state.get(f"lab_{m_id}", True)
            label_size = st.session_state.get(f"lsz_{m_id}", 11)
            dot_size = st.session_state.get(f"dsz_{m_id}", 11)
        
        figs = create_nonlife_six_dimensional_charts(
            df_raw=df_filtered,
            target_year=latest_year,
            cos=selected_cos,
            divisor=divisor,
            unit_label=unit_label,
            highlight_co=current_hl,
            label_size=label_size,
            show_labels=show_labels,
            dot_size=dot_size
        )
        # 以 3 列网格展示 6 个子图
        for i, fig in enumerate(figs):
            if i % 3 == 0:
                cols = st.columns(3)
            with cols[i % 3]:
                st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})
        display_notes(m_id, df_filtered, "六维度")
        display_bottom_note(notes_dict.get(m_id, {}).get('note', ''))
        return

    # ==========================================
    # 🆕 10. 综合赔付率趋势图（新增）
    # ==========================================
    if m_id == "claim_ratio_trend":
        if not print_mode:
            show_labels = st.toggle("显示标签", True, key=f"lab_{m_id}")
        else:
            show_labels = st.session_state.get(f"lab_{m_id}", True)
        figs = create_ratio_trend_chart(
            df_filtered, selected_cos,
            factor_list=[
                "当期发生赔款及理赔费用",
                "已发生赔款负债履约现金流变动",
                "亏损合同损益",
                "再保净成本"
            ],
            ratio_field="综合赔付率",
            title_prefix="综合赔付率拆解趋势",
            divisor=divisor,
            unit_label=unit_label,
            highlight_co=current_hl,
            show_labels=show_labels,
            ratio_first=True   # 综合图放在首位
        )
        for fig in figs:
            show_chart(fig, print_mode, m_id)
        display_notes(m_id, df_filtered, "综合赔付率")
        display_bottom_note(notes_dict.get(m_id, {}).get('note', ''))
        return

    # ==========================================
    # 🆕 11. 综合费用率趋势图（新增）
    # ==========================================
    if m_id == "expense_ratio_trend":
        if not print_mode:
            show_labels = st.toggle("显示标签", True, key=f"lab_{m_id}")
        else:
            show_labels = st.session_state.get(f"lab_{m_id}", True)
        
        figs = create_ratio_trend_chart(
            df_filtered, selected_cos,
            factor_list=[
                "获取费用摊销",
                "维持费用"
            ],
            ratio_field="综合费用率",
            title_prefix="综合费用率拆解趋势",
            divisor=divisor,
            unit_label=unit_label,
            highlight_co=current_hl,
            show_labels=show_labels,
            ratio_first=True   # 综合图放在首位
        )
        
        # 逐个显示图表（每个因子一张图）
        for fig in figs:
            show_chart(fig, print_mode, m_id)
        
        # 显示注释和底部说明
        display_notes(m_id, df_filtered, "综合费用率")
        display_bottom_note(notes_dict.get(m_id, {}).get('note', ''))
        return
    # ==========================================
    # 保险服务收入业务构成（堆叠图）
    # ==========================================
    if m_id == "premium_composition":
        fig = create_premium_stacked_chart(
            df_filtered, selected_cos, latest_year, divisor, unit_label, current_hl
        )
        show_chart(fig, print_mode, m_id)
        display_notes(m_id, df_filtered, "保险服务收入构成")
        display_bottom_note(notes_dict.get(m_id, {}).get('note', ''))
        return

    # ==========================================
    # 承保利润业务构成（堆叠图）
    # ==========================================
    if m_id == "profit_composition":
        fig = create_profit_contribution_stacked_chart(
            df_filtered, selected_cos, latest_year, divisor, unit_label, current_hl
        )
        show_chart(fig, print_mode, m_id)
        display_notes(m_id, df_filtered, "承保利润构成")
        display_bottom_note(notes_dict.get(m_id, {}).get('note', ''))
        return

    # ==========================================
    # 保险业务构成（两种方法占比堆叠图）
    # ==========================================
    if m_id == "premium_method_composition":
        fig = create_method_composition_chart(
            df_filtered, selected_cos, latest_year, divisor, unit_label
        )
        show_chart(fig, print_mode, m_id)
        display_notes(m_id, df_filtered, "保险业务构成")
        display_bottom_note(notes_dict.get(m_id, {}).get('note', ''))
        return

    # ==========================================
    # 关键指标 - 总资产与净资产（柱状图）
    # ==========================================
    if m_id == "key_assets_equity":
        # 设置UI控件（仅非打印模式）
        if not print_mode:
            c1, c2, c3 = st.columns(3)
            with c1:
                show_labels = st.toggle("显示标签", True, key=f"lab_{m_id}")
            with c2:
                pct_sz = st.slider("涨幅字号", 8, 24, 14, key=f"psz_{m_id}")
            with c3:
                gap = st.slider("柱间距", 0.1, 0.8, 0.15, key=f"gap_{m_id}")
        else:
            show_labels = st.session_state.get(f"lab_{m_id}", True)
            pct_sz = st.session_state.get(f"psz_{m_id}", 14)
            gap = st.session_state.get(f"gap_{m_id}", 0.15)
        
        # 绘制总资产图
        fig_asset = create_kpmg_chart(
            df_filtered, "总资产", "", show_labels, pct_sz, gap,
            sort_by_value=False, is_percentage=False
        )
        # 绘制净资产图
        fig_equity = create_kpmg_chart(
            df_filtered, "净资产", "", show_labels, pct_sz, gap,
            sort_by_value=False, is_percentage=False
        )
        
        # 显示两个图表（上下排列）
        show_chart(fig_asset, print_mode, m_id)
        show_chart(fig_equity, print_mode, m_id)
        
        display_notes(m_id, df_filtered, "总资产与净资产")
        display_bottom_note(notes_dict.get(m_id, {}).get('note', ''))
        return

    # ==========================================
    # 关键指标 - 承保利润（柱状图）
    # ==========================================
    if m_id == "key_underwriting_profit":
        if not print_mode:
            c1, c2, c3 = st.columns(3)
            with c1:
                show_labels = st.toggle("显示标签", True, key=f"lab_{m_id}")
            with c2:
                pct_sz = st.slider("涨幅字号", 8, 24, 14, key=f"psz_{m_id}")
            with c3:
                gap = st.slider("柱间距", 0.1, 0.8, 0.15, key=f"gap_{m_id}")
        else:
            show_labels = st.session_state.get(f"lab_{m_id}", True)
            pct_sz = st.session_state.get(f"psz_{m_id}", 14)
            gap = st.session_state.get(f"gap_{m_id}", 0.15)
        
        fig = create_kpmg_chart(df_filtered, "承保利润", "", show_labels, pct_sz, gap,
                                sort_by_value=False, is_percentage=False)
        show_chart(fig, print_mode, m_id)
        display_notes(m_id, df_filtered, "承保利润")
        display_bottom_note(notes_dict.get(m_id, {}).get('note', ''))
        return

    # ==========================================
    # 关键指标 - 投资利润（柱状图）
    # ==========================================
    if m_id == "key_investment_profit":
        if not print_mode:
            c1, c2, c3 = st.columns(3)
            with c1:
                show_labels = st.toggle("显示标签", True, key=f"lab_{m_id}")
            with c2:
                pct_sz = st.slider("涨幅字号", 8, 24, 14, key=f"psz_{m_id}")
            with c3:
                gap = st.slider("柱间距", 0.1, 0.8, 0.15, key=f"gap_{m_id}")
        else:
            show_labels = st.session_state.get(f"lab_{m_id}", True)
            pct_sz = st.session_state.get(f"psz_{m_id}", 14)
            gap = st.session_state.get(f"gap_{m_id}", 0.15)
        
        fig = create_kpmg_chart(df_filtered, "投资利润", "", show_labels, pct_sz, gap,
                                sort_by_value=False, is_percentage=False)
        show_chart(fig, print_mode, m_id)
        display_notes(m_id, df_filtered, "投资利润")
        display_bottom_note(notes_dict.get(m_id, {}).get('note', ''))
        return

    # ==========================================
    # 关键指标 - 净利润（柱状图）
    # ==========================================
    if m_id == "key_net_profit":
        if not print_mode:
            c1, c2, c3 = st.columns(3)
            with c1:
                show_labels = st.toggle("显示标签", True, key=f"lab_{m_id}")
            with c2:
                pct_sz = st.slider("涨幅字号", 8, 24, 14, key=f"psz_{m_id}")
            with c3:
                gap = st.slider("柱间距", 0.1, 0.8, 0.15, key=f"gap_{m_id}")
        else:
            show_labels = st.session_state.get(f"lab_{m_id}", True)
            pct_sz = st.session_state.get(f"psz_{m_id}", 14)
            gap = st.session_state.get(f"gap_{m_id}", 0.15)
        
        fig = create_kpmg_chart(df_filtered, "净利润", "", show_labels, pct_sz, gap,
                                sort_by_value=False, is_percentage=False)
        show_chart(fig, print_mode, m_id)
        display_notes(m_id, df_filtered, "净利润")
        display_bottom_note(notes_dict.get(m_id, {}).get('note', ''))
        return

    # ==========================================
    # 关键指标 - 利润构成（承保 vs 投资 堆叠图）
    # ==========================================
    if m_id == "key_profit_composition":
        fig = create_profit_composition_chart(
            df_filtered, selected_cos, latest_year, divisor, unit_label
        )
        show_chart(fig, print_mode, m_id)
        display_notes(m_id, df_filtered, "利润构成")
        display_bottom_note(notes_dict.get(m_id, {}).get('note', ''))
        return
    # ==========================================
    # 未匹配
    # ==========================================
    st.warning(f"未识别的模块 ID：{m_id}，请检查 Excel 模板。")

# 7.2 统一渲染出口
def render_report_module(m_id, print_mode, is_first=False):
    global notes_dict  # ✅ 添加这一行
    """
    统一渲染出口：标题 + 手动截图覆盖 + 注释框 + 图表 + 底部注释
    内部调用 render_pure_chart_entity 绘制图表
    """
    mod_data = notes_dict.get(m_id, {})
    
    # ---- 1. 生成完整标题（从 notes_dict 读取） ----
    full_title = mod_data.get('title', m_id)
    if 'df_notes' in st.session_state and isinstance(st.session_state['df_notes'], pd.DataFrame):
        df_n = st.session_state['df_notes']
        if '模块ID' in df_n.columns:
            match = df_n[df_n['模块ID'] == m_id]
            if not match.empty:
                r = match.iloc[0]
                title_parts = []
                for field in ['一级分类', '二级分类', '对应图表名称']:
                    val = str(r.get(field, '')).strip()
                    if val and val.lower() != 'nan' and val != '全部':
                        title_parts.append(val)
                if title_parts:
                    full_title = " - ".join(title_parts)
    
    # ---- 2. 打印模式容器 ----
    if print_mode:
        st.markdown("<div class='page-break-container' style='margin:0;padding:0;page-break-inside:avoid;'>", unsafe_allow_html=True)
    if not print_mode:
        st.markdown(
            "<div class='no-print' style='height:2px; background:linear-gradient(to right, #00338D, #0865EE, #00338D); opacity:0.3; margin:0;'></div>",
            unsafe_allow_html=True
        )
    
    # ---- 3. 标题（打印和网页都执行） ----
    title_cls = "page-break-title" if (print_mode and not is_first) else ""
    mt = "0px" if (print_mode and is_first) else "20px"
    font_size = "35px" if print_mode else "30px"
    st.markdown(
        f"<h3 class='{title_cls}' style='"
        f"text-align:left; color:#00338D; font-size:{font_size}; font-weight:900; "
        f"font-family:Microsoft YaHei, 微软雅黑, sans-serif; "
        f"margin-top:{mt}; margin-bottom:20px; border:none; padding-bottom:0px;'>"
        f"{full_title}</h3>",
        unsafe_allow_html=True
    )
    
    # ---- 4. 手动截图覆盖 ----
    if 'manual_upload_images' in st.session_state and m_id in st.session_state.manual_upload_images:
        if print_mode:
            st.image(st.session_state.manual_upload_images[m_id], use_column_width=True)
        else:
            img_col_left, img_col_center, img_col_right = st.columns([1, 8, 1])
            with img_col_center:
                st.image(st.session_state.manual_upload_images[m_id], use_column_width=True)
        # 显示注释（display_notes 内部已包含注释显示，无需再单独调用 display_bottom_note）
        display_notes(m_id)
        if print_mode:
            st.markdown("</div>", unsafe_allow_html=True)
        return
    
    # ---- 5. AI / 注释框（预调用，获取注释文本，并显示注释） ----
    # display_notes 内部会生成 AI 洞察并显示分析内容，同时负责显示注释
    display_notes(m_id, ai_df=st.session_state.get('df_filtered', pd.DataFrame()), ai_field=mod_data.get('title', m_id))
    
    # ---- 6. 图表（区分打印/网页模式） ----
    if print_mode:
        render_pure_chart_entity(m_id, print_mode)
    else:
        chart_col_left, chart_col_center, chart_col_right = st.columns([1, 10, 1])
        with chart_col_center:
            render_pure_chart_entity(m_id, print_mode)
    
    # ---- 7. 不再单独调用 display_bottom_note，避免重复 ----
    # 注释已由第5步的 display_notes 统一显示
    
    if print_mode:
        st.markdown("</div>", unsafe_allow_html=True)
        
# 8.税务分析图表
def create_tax_subplot_chart(df_pivot, selected_cos, show_labels, bar_width, label_size, co_font_size, co_y_offset, highlight_co="无"):
    av_cos = [c for c in selected_cos if c in df_pivot['公司'].unique()]
    if not av_cos:
        return go.Figure()
    
    cols = ['税前利润总额', '净利润']
    g_max = df_pivot[cols].max().max() / divisor
    g_min = df_pivot[cols].min().min() / divisor
    y_top = g_max * 1.25 if pd.notna(g_max) else 100
    y_bot = min(0, g_min * 1.2) if pd.notna(g_min) else 0
    lbl_y = g_max * 1.15 if pd.notna(g_max) else 100
    ph_h = g_max * 0.4 if pd.notna(g_max) else 10
    all_yrs = sorted(df_pivot['报告年份'].astype(str).unique())
    
    n = len(av_cos)
    fig = make_subplots(rows=1, cols=n, shared_yaxes=True, 
                        column_titles=[f"<b><span style='color:#00338D;'>{c}</span></b>" for c in av_cos], 
                        horizontal_spacing=0.03 if n <= 1 else min(0.03, 0.8 / (n - 1)))
    
    for nm, c in zip(cols, ['#1E49E2', '#C7A0F7']):
        fig.add_trace(go.Scatter(x=[None], y=[None], mode="markers", 
                                 marker=dict(symbol="square", size=12, color=c), 
                                 name=nm, showlegend=True), row=1, col=1)

    for i, co in enumerate(av_cos):
        ci = i + 1
        d = df_pivot[df_pivot['公司'] == co].sort_values('报告年份')
        x = d['报告年份'].astype(str).tolist()
        yp = (d[cols[0]] / divisor).tolist()
        yn = (d[cols[1]] / divisor).tolist()
        tr = d['有效税率'].tolist()
        
        pr, nr, pt, nt, my, mt = [], [], [], [], [], []
        for vp, vn in zip(yp, yn):
            if (pd.isna(vp) or vp == 0) and (pd.isna(vn) or vn == 0):
                pr.append(0); nr.append(0); pt.append(""); nt.append(""); my.append(ph_h); mt.append("未披露")
            else:
                pr.append(vp); nr.append(vn)
                pt.append(f"{vp:.1f}" if show_labels and pd.notna(vp) else "")
                nt.append(f"{vn:.1f}" if show_labels and pd.notna(vn) else "")
                my.append(0); mt.append("")
        
        fig.add_trace(go.Bar(x=x, y=pr, marker_color='#1E49E2', text=pt, 
                             textposition='outside', textfont=dict(size=label_size), 
                             width=bar_width, offsetgroup=1, showlegend=False, cliponaxis=False), row=1, col=ci)
        fig.add_trace(go.Bar(x=x, y=nr, marker_color='#C7A0F7', text=nt, 
                             textposition='outside', textfont=dict(size=label_size), 
                             width=bar_width, offsetgroup=2, showlegend=False, cliponaxis=False), row=1, col=ci)
        fig.add_trace(go.Bar(x=x, y=my, marker_color="#CDCDCD", text=mt, 
                             textposition='inside', insidetextanchor='middle', 
                             textfont=dict(size=12, color="white"), 
                             width=bar_width * 2, showlegend=False, cliponaxis=False, hoverinfo='skip'), row=1, col=ci)
        
        for j, yr in enumerate(x):
            if pd.notna(tr[j]) and mt[j] == "":
                fig.add_annotation(x=yr, y=lbl_y, text=f"<b>{tr[j]:.0%}</b>", 
                                   showarrow=False, font=dict(size=label_size + 2, color="#97014F" if tr[j] >= 0 else "#269924"),
                                   xref=f"x{ci}" if ci > 1 else "x", yref="y1")
        
        hl = (str(co).strip() == str(highlight_co).strip())
        fig.add_shape(type="rect", xref=f"x{ci} domain" if ci > 1 else "x domain", yref="paper",
                      x0=-0.05, x1=1.05, y0=-0.12, y1=1.1,
                      line=dict(color=HL_BOX_LINE if hl else "rgba(200,200,200,0.3)", width=1.5 if hl else 1),
                      fillcolor=HL_BOX_FILL if hl else "rgba(0,0,0,0)", layer="above")

    fig.update_layout(height=550, margin=dict(t=40, b=100, l=20, r=20),
                      plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)',
                      legend=dict(orientation="h", yanchor="top", y=-0.15, xanchor="center", x=0.5))
    
    for i in range(1, n + 1):
        fig.update_xaxes(type='category', categoryorder='array', categoryarray=all_yrs,
                         showline=True, linecolor='lightgray', linewidth=1,
                         showgrid=False, tickfont=dict(size=10), zeroline=False, ticks="", ticklen=0, row=1, col=i)
        fig.update_yaxes(showline=False, showgrid=False, showticklabels=False,
                         zeroline=True, zerolinecolor="#E0E0E0", range=[y_bot, y_top], row=1, col=i)
    
    for a in fig.layout.annotations:
        if "span" in str(a.text):
            a.update(y=co_y_offset, font_size=co_font_size)
    
    return fig

# 9.财务与投资分析图表
def create_asset_composition_chart(df, selected_cos, field_map, color_map, title_prefix, show_labels, label_size, bar_width, co_font_size, highlight_co="无"):
    fields = list(field_map.keys())
    d_struct = df[(df['公司'].isin(selected_cos)) & (df['字段名'].isin(fields))].copy()
    d_struct['报告年份'] = d_struct['报告年份'].astype(str).str.replace(".0", "", regex=False) + "YE"
    all_yrs = [f"{prev_year}YE", f"{latest_year}YE"]
    av_cos = [co for co in selected_cos if co in d_struct['公司'].unique()]
    if not av_cos:
        return go.Figure()

    hl_co = str(highlight_co).strip()
    fig = make_subplots(rows=1, cols=len(av_cos), shared_yaxes=True, 
                        column_titles=[f"<span style='color:#00338D;'><b>{co}</b></span>" for co in av_cos], 
                        horizontal_spacing=0.015)
    for i, co in enumerate(av_cos):
        d_co = d_struct[d_struct['公司'] == co].pivot(index='报告年份', columns='字段名', values='(百万)人民币').reindex(all_yrs).fillna(0)
        for f in fields:
            if f not in d_co.columns:
                d_co[f] = 0
        raw_total = d_co.sum(axis=1)
        d_co['T'] = raw_total.replace(0, 1)
        
        for fn in fields:
            val_pct = d_co[fn] / d_co['T'] * 100
            is_dark = any(x in color_map.get(fn, "") for x in ["00338D", "510DBC", "1E49E2"])
            txt_c = "white" if is_dark else "black"
            fig.add_trace(go.Bar(
                x=d_co.index, y=val_pct,
                name=field_map[fn] if i == 0 else None,
                marker_color=color_map.get(fn, "#CCCCCC"),
                text=[f"{v:.0f}%" if show_labels and raw_total.iloc[idx] > 0 else "" for idx, v in enumerate(val_pct)],
                textposition='inside', insidetextanchor='middle',
                textfont=dict(size=label_size, color=txt_c),
                constraintext='none', textangle=0, cliponaxis=False,
                width=bar_width, showlegend=(i == 0), legendgroup=fn, hoverinfo="skip"
            ), row=1, col=i + 1)
        
        fig.add_trace(go.Bar(
            x=d_co.index, y=[100 if t == 0 else 0 for t in raw_total],
            name="未披露", marker_color="#CDCDCD",
            text=["未披露" if t == 0 else "" for t in raw_total],
            textposition='inside', insidetextanchor='middle',
            textfont=dict(size=label_size, color="white"),
            constraintext='none', textangle=0, cliponaxis=False,
            width=bar_width, showlegend=False, hoverinfo="skip"
        ), row=1, col=i + 1)

        is_hl = (str(co).strip() == hl_co)
        fig.add_shape(
            type="rect", xref="x domain", yref="y domain",
            x0=-0.06, x1=1.06, y0=-0.12, y1=1.12,
            fillcolor=HL_BOX_FILL if is_hl else "rgba(0,0,0,0)",
            line=dict(color=HL_BOX_LINE, width=1.5) if is_hl else dict(color="rgba(0,0,0,0)", width=0),
            layer="above", row=1, col=i + 1
        )

    fig.update_layout(
        barmode='stack', height=550,
        paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)',
        uniformtext=dict(minsize=label_size, mode='show'),
        margin=dict(t=50, b=120, l=40, r=40),
        legend=dict(orientation="h", yanchor="top", y=-0.2, xanchor="center", x=0.5,
                    font=dict(size=11), itemsizing="constant")
    )
    for i in range(1, len(av_cos) + 1):
        fig.update_xaxes(showgrid=False, showline=False, zeroline=False,
                         tickfont=dict(size=10), ticks="", ticklen=0, row=1, col=i)
        fig.update_yaxes(showgrid=False, range=[0, 101], showline=False, zeroline=False,
                         tickvals=[0, 25, 50, 75, 100], ticktext=["0%", "25%", "50%", "75%", "100%"],
                         showticklabels=(i == 1), row=1, col=i)
    for ann in fig.layout.annotations:
        ann.update(y=ann.y + 0.05, font_size=co_font_size)
    return fig

def create_oci_chart(df_raw, year, title, show_labels, co_font_size, bar_gap, selected_cos, highlight_co="无"):
    oci_fields = ['可转损益OCI合计', '不可转损益OCI合计', '净利润', '综合收益总额']
    d_sub = df_raw[(df_raw['字段名'].isin(oci_fields)) & 
                   (df_raw['报告年份'].astype(str) == str(year).replace(".0", ""))].drop_duplicates(
                       subset=['公司', '报告年份', '字段名'], keep='last').copy()
    if d_sub.empty:
        return go.Figure().update_layout(paper_bgcolor='rgba(0,0,0,0)')

    df_pivot = d_sub.pivot_table(index='公司', columns='字段名', values='(百万)人民币', aggfunc='sum').fillna(0)
    for f in oci_fields:
        if f not in df_pivot.columns:
            df_pivot[f] = 0.0
    df_pivot = df_pivot / divisor
    df_pivot['other'] = df_pivot['综合收益总额'] - df_pivot['净利润'] - df_pivot['可转损益OCI合计'] - df_pivot['不可转损益OCI合计']
    mc = {
        "净利润": {"c": "rgb(0,176,240)", "n": "净利润"},
        "可转损益OCI合计": {"c": "rgb(253,52,156)", "n": "可转损益OCI变动"},
        "不可转损益OCI合计": {"c": "rgb(114,19,234)", "n": "不可转损益OCI变动"},
        "other": {"c": "rgb(127,127,127)", "n": "其他"},
        "综合收益总额": {"c": "rgb(172,234,255)", "n": "综合收益总额"}
    }

    av_cos = [c for c in selected_cos if c in df_pivot.index]
    hl_co = str(highlight_co).strip()
    if not av_cos:
        return go.Figure()
    all_vals = df_pivot[['净利润', '可转损益OCI合计', '不可转损益OCI合计', 'other', '综合收益总额']].values.flatten()
    max_val = np.nanmax(all_vals)
    min_val = np.nanmin(all_vals)
    buffer = (max_val - min_val) * 0.8 if max_val != min_val else abs(max_val) * 0.3
    y_range = [min_val - buffer if min_val < 0 else min_val - abs(min_val) * 0.3, max_val + buffer]

    fig = make_subplots(rows=1, cols=len(av_cos), shared_yaxes=True, horizontal_spacing=0.015,
                        subplot_titles=[f"<b><span style='color:#00338D;'>{co}</span></b>" for co in av_cos])
    for col_idx, co in enumerate(av_cos):
        for m_key, m_info in mc.items():
            val = df_pivot.loc[co].get(m_key, 0)
            fig.add_trace(go.Bar(
                name=m_info["n"], x=[m_key], y=[val],
                text=[f"{val:.0f}" if (show_labels and val != 0) else ""],
                textposition='outside', textfont=dict(size=11, color='#00338D'),
                marker_color=m_info["c"], width=0.8,
                legendgroup=m_key, showlegend=(col_idx == 0),
                cliponaxis=False, constraintext='none'
            ), row=1, col=col_idx + 1)

        is_hl = (str(co).strip() == hl_co)
        bg_fill = HL_BOX_FILL if is_hl else "rgba(240, 240, 240, 0.35)"
        line_dict = dict(color=HL_BOX_LINE, width=1.5) if is_hl else dict(color="rgba(210, 210, 210, 0.6)", width=1)
        fig.add_shape(
            type="rect", xref="x domain", yref="y domain",
            x0=-0.04, x1=1.04, y0=-0.04, y1=1.13,
            fillcolor=bg_fill, line=line_dict, layer="above" if is_hl else "below",
            row=1, col=col_idx + 1
        )
        fig.update_xaxes(showticklabels=False, showline=False, zeroline=False,
                         showgrid=False, ticks="", ticklen=0, row=1, col=col_idx + 1)

    fig.update_layout(
        paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)',
        barmode='group', bargap=bar_gap,
        height=420,
        margin=dict(t=50, b=40, l=40, r=30),
        legend=dict(orientation="h", yanchor="top", y=-0.2, xanchor="center", x=0.5, font=dict(size=11))
    )
    fig.update_yaxes(range=y_range, showline=False, zeroline=False, showgrid=False,
                     gridcolor='rgba(0,0,0,0)', gridwidth=0, row=1, col="all")
    for ann in fig.layout.annotations:
        ann.update(y=1.03, font_size=co_font_size)
    return fig

def create_asset_liab_oci_chart(df_raw, selected_cos, bar_gap, co_font_size, show_labels, highlight_co="无"):
    yrs = [str(prev_year), str(latest_year)]
    fns = ['可转损益的负债OCI', 'FVOCI债券公允价值']
    d = df_raw[df_raw['字段名'].isin(fns)].copy()
    d['v'] = d['(百万)人民币'] / divisor
    av_cos = [c for c in selected_cos if c in d['公司'].unique()]
    if not av_cos:
        return go.Figure()

    def gv(c, y, f):
        return d[(d['公司'] == c) & (d['报告年份'].astype(str) == y) & (d['字段名'] == f)]['v'].sum()

    all_v = [gv(c, y, f) for c in av_cos for y in yrs for f in fns]
    g_max = max(all_v + [0.1])
    g_min = min(all_v + [-0.1])
    ph_h = max(abs(g_max), abs(g_min)) * 0.4
    n = len(av_cos)
    colors = {fns[0]: "rgb(0, 184, 245)", fns[1]: "rgb(253, 52, 156)"}

    fig = make_subplots(rows=1, cols=n, shared_yaxes=True,
                        horizontal_spacing=0.01 if n <= 1 else min(0.01, 0.8 / (n - 1)),
                        subplot_titles=[f"<b><span style='color:#00338D;'>{c}</span></b>" for c in av_cos])
    for fn, c in colors.items():
        fig.add_trace(go.Scatter(x=[None], y=[None], mode="markers",
                                 marker=dict(symbol="square", size=12, color=c),
                                 name=fn, showlegend=True), row=1, col=1)

    for i, co in enumerate(av_cos):
        ci = i + 1
        x_lbl = [f"{y}YE" for y in yrs]
        v1, v2, t1, t2, my = [], [], [], [], []
        for y in yrs:
            val1, val2 = gv(co, y, fns[0]), gv(co, y, fns[1])
            m = (val1 == 0 and val2 == 0)
            v1.append(0 if m else val1)
            v2.append(0 if m else val2)
            my.append(ph_h if m else 0)
            t1.append("" if m else (f"{val1:.0f}" if show_labels and val1 != 0 else ""))
            t2.append("" if m else (f"{val2:.0f}" if show_labels and val2 != 0 else ""))
            if m:
                fig.add_annotation(x=f"{y}YE", y=ph_h / 2, text="未披露", showarrow=False,
                                   font=dict(color="white", size=12),
                                   xref=f"x{ci}" if ci > 1 else "x", yref="y1",
                                   xanchor="center", yanchor="middle")

        c1_list = ["#ACEAFF", colors[fns[0]]]
        c2_list = ["#FFD6E8", colors[fns[1]]]

        fig.add_trace(go.Bar(x=x_lbl, y=v1, marker_color=c1_list, text=t1,
                             textposition='outside', textfont=dict(size=11, color='#00338D'),
                             width=0.4, offsetgroup=1, showlegend=False,
                             cliponaxis=False, constraintext='none'), row=1, col=ci)
        fig.add_trace(go.Bar(x=x_lbl, y=v2, marker_color=c2_list, text=t2,
                             textposition='outside', textfont=dict(size=11, color='#00338D'),
                             width=0.4, offsetgroup=2, showlegend=False,
                             cliponaxis=False, constraintext='none'), row=1, col=ci)
        fig.add_trace(go.Bar(x=x_lbl, y=my, marker_color="#CDCDCD",
                             width=0.8, offset=-0.4, showlegend=False,
                             cliponaxis=False, hoverinfo='skip'), row=1, col=ci)

        hl = (str(co).strip() == str(highlight_co).strip())
        xref_str = f"x{ci} domain" if ci > 1 else "x domain"
        fig.add_shape(type="rect", xref=xref_str, yref="y domain",
                      x0=-0.06, x1=1.06, y0=0, y1=1.08,
                      fillcolor=HL_BOX_FILL if hl else "rgba(245,245,245,0.5)",
                      line=dict(color=HL_BOX_LINE if hl else "#CCCCCC", width=1.5 if hl else 1),
                      layer="below")

    fig.update_layout(barmode='group', bargap=bar_gap,
                      paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)',
                      height=420,
                      margin=dict(t=50, b=80, l=20, r=20),
                      legend=dict(orientation="h", yanchor="top", y=-0.28, x=0.5, xanchor="center"))
    y_rng = [g_min - abs(g_min) * 0.3, g_max + abs(g_max) * 0.3]

    for i in range(1, n + 1):
        fig.update_xaxes(type='category', categoryorder='array', categoryarray=x_lbl,
                         showline=False, zeroline=False, showgrid=False, ticks="", ticklen=0, row=1, col=i)
        fig.update_yaxes(range=y_rng, showline=False, zeroline=True, zerolinecolor="#E0E0E0",
                         zerolinewidth=1.02, showgrid=False, gridcolor='rgba(0,0,0,0)',
                         gridwidth=0, row=1, col=i)
    for ann in fig.layout.annotations:
        if "span" in str(ann.text):
            ann.update(y=1.08, font_size=co_font_size)

    return fig

def create_equity_change_detail_table(df_raw, target_year, cos, divisor=1, unit_label="", highlight_co="无"):
    df = df_raw.copy()
    df["报告年份"] = df["报告年份"].astype(str).str.replace(".0", "", regex=False).str.strip()
    df["公司"] = df["公司"].astype(str).str.strip()
    df["字段名"] = df["字段名"].astype(str).str.strip()

    year_str = str(target_year)

    fields = [
        ("1", "净利润", "净利润"),
        ("2", "其他综合收益", "其他综合收益"),
        ("3", "股东权益分配", "股东权益分配"),
        ("4", "股东增资", "股东增资"),
        ("5", "发债", "发债"),
        ("6", "子公司合并", "子公司合并"),
        ("7", "其他权益变动", "其他权益变动"),
        ("8=sum(1-7)", "净资产变动值", "净资产变动值"),
    ]

    val_col = "(百万)人民币" if "(百万)人民币" in df.columns else df.columns[-1]

    d = df[
        (df["报告年份"] == year_str) &
        (df["公司"].isin(cos)) &
        (df["字段名"].isin([x[2] for x in fields]))
    ].drop_duplicates(subset=["公司", "报告年份", "字段名"], keep="first").copy()

    def to_num(v):
        if pd.isna(v):
            return np.nan
        s = str(v).strip()
        if s.lower() in ["", "nan", "none", "null", "未披露", "—", "-"]:
            return np.nan
        try:
            return float(
                s.replace(",", "")
                 .replace("(", "-")
                 .replace(")", "")
            )
        except:
            return np.nan

    data_map = {}
    for _, r in d.iterrows():
        data_map[(str(r["公司"]).strip(), str(r["字段名"]).strip())] = to_num(r.get(val_col, np.nan))

    for co in cos:
        co = str(co).strip()
        raw_net_change = data_map.get((co, "净资产变动值"), np.nan)
        if pd.isna(raw_net_change):
            vals = [data_map.get((co, f[2]), np.nan) for f in fields[:7]]
            vals_valid = [x for x in vals if pd.notna(x)]
            if vals_valid:
                data_map[(co, "净资产变动值")] = sum(vals_valid)

    def fmt(v):
        if pd.isna(v):
            return "未披露", True
        v = v / divisor
        if abs(v) < 1e-9:
            return "0.0", False
        return f"{v:,.1f}", False

    current_hl = str(highlight_co).strip()

    html = (
        "<style>"
        "@media print {"
        ".equity-detail-wrap { margin-top:4px!important; margin-bottom:6px!important; }"
        ".equity-detail-title { font-size:10px!important; margin-bottom:3px!important; }"
        ".equity-detail-table { font-size:8.5px!important; }"
        ".equity-detail-table th { padding:3px 4px!important; line-height:1.15!important; }"
        ".equity-detail-table td { padding:2.5px 4px!important; line-height:1.15!important; }"
        "}"
        "</style>"
        "<div class='equity-detail-wrap' style='margin-top:8px; margin-bottom:10px;'>"
        f"<div class='equity-detail-title' style='font-size:13px; font-weight:bold; color:#00338D; margin-bottom:5px;'>"
        f"{year_str}年净资产变动明细表"
        + (f"（单位：{unit_label}）" if unit_label else "")
        + "</div>"
    )

    html += (
        "<table class='equity-detail-table' style='width:100%; border-collapse:collapse; "
        "font-family:Microsoft YaHei, sans-serif; font-size:11px; table-layout:fixed;'>"
    )

    html += (
        "<tr style='background-color:#00338D; color:white; text-align:center; font-weight:bold;'>"
        "<th style='padding:5px 6px; border:1px solid white; width:10%;'>序号</th>"
        "<th style='padding:5px 6px; border:1px solid white; width:16%;'>项目</th>"
    )
    for co in cos:
        co_str = str(co).strip()
        html += f"<th style='padding:5px 6px; border:1px solid white;'>{co_str}</th>"
    html += "</tr>"

    for row_idx, (no, display_name, field_name) in enumerate(fields):
        row_bg = "#F8F9FA" if row_idx % 2 == 0 else "white"

        html += f"<tr style='background-color:{row_bg};'>"
        html += f"<td style='padding:4px 6px; text-align:center; border:1px solid #EAEAEA; color:#333;'>{no}</td>"
        html += f"<td style='padding:4px 6px; text-align:left; border:1px solid #EAEAEA; font-weight:bold; color:#333;'>{display_name}</td>"

        for co in cos:
            co_str = str(co).strip()
            val, missing = fmt(data_map.get((co_str, field_name), np.nan))

            if co_str == current_hl and not missing:
                cell_style = (
                    "background-color:rgba(0,51,141,0.08); color:#00338D; "
                    "font-weight:bold; border:1px solid #00338D;"
                )
            elif missing:
                cell_style = "background-color:#CDCDCD; color:white; border:1px solid #EAEAEA;"
            else:
                cell_style = "color:#333; border:1px solid #EAEAEA;"

            html += f"<td style='padding:4px 6px; text-align:center; {cell_style}'>{val}</td>"

        html += "</tr>"

    html += "</table></div>"
    return html

def create_expense_structure_chart(df, selected_cos, show_labels, label_size, bar_width, co_font_size, highlight_co="无"):
    prev_year = st.session_state.get('prev_year', 2024)
    latest_year = st.session_state.get('latest_year', 2025)
    years = [str(prev_year), str(latest_year)]
    field_map = {
        "获取费用摊销": "获取费用摊销",
        "维持费用": "维持费用",
    }
    fields = list(field_map.keys())
    
    raw_df = df.copy()
    raw_df.columns = raw_df.columns.str.strip()
    raw_df['公司'] = raw_df['公司'].astype(str).str.strip()
    raw_df['字段名'] = raw_df['字段名'].astype(str).str.strip()
    # 统一处理年份，去除 .0
    raw_df['报告年份'] = raw_df['报告年份'].astype(str).str.replace('.0', '', regex=False).str.strip()
    # 过滤无效年份
    raw_df = raw_df[raw_df['报告年份'] != '']
    raw_df = raw_df[~raw_df['报告年份'].str.lower().isin(['nan', 'none'])]
    
    service_revenue = {}
    for co in selected_cos:
        co_clean = co.strip()
        for yr in years:
            mask = (raw_df['公司'] == co_clean) & \
                   (raw_df['报告年份'] == yr) & \
                   (raw_df['字段名'] == '保险服务收入')
            rev_series = raw_df.loc[mask, '(百万)人民币']
            rev = rev_series.sum() if not rev_series.empty else 0
            if pd.isna(rev) or rev == 0:
                mask2 = (raw_df['公司'] == co_clean) & \
                        (raw_df['报告年份'] == yr) & \
                        (raw_df['字段名'] == '保险业务收入')
                rev_series2 = raw_df.loc[mask2, '(百万)人民币']
                rev = rev_series2.sum() if not rev_series2.empty else 0
            if rev == 0:
                st.warning(f"⚠️ 公司 {co} 在 {yr} 年未找到保险服务收入或保险业务收入！")
                rev = 1
            service_revenue[(co, yr)] = rev
    
    fig = make_subplots(rows=1, cols=len(selected_cos), horizontal_spacing=0.02,
                        subplot_titles=[f"<span style='color:#00338D;'><b>{co}</b></span>" for co in selected_cos])
    
    for i, co in enumerate(selected_cos):
        co_clean = co.strip()
        cd = raw_df[raw_df['公司'] == co_clean]
        vals = {}
        for f in fields:
            vals[f] = {}
            for yr in years:
                mask_f = (cd['报告年份'] == yr) & (cd['字段名'] == f)
                v_series = cd.loc[mask_f, '(百万)人民币']
                v = v_series.sum() if not v_series.empty else 0
                denominator = service_revenue.get((co, yr), 1)
                ratio = v / denominator * 100 if denominator != 0 else 0
                vals[f][yr] = ratio if pd.notna(ratio) else 0
        
        x_axis = [f"{y}YE" for y in years]
        for f in fields:
            y_vals = [vals[f].get(y, 0) for y in years]
            fig.add_trace(go.Bar(
                x=x_axis, y=y_vals,
                name=field_map[f],
                marker_color="#00338D" if f == "获取费用摊销" else "#0865EE",
                width=bar_width,
                showlegend=(i == 0),
                text=[f"{v:.1f}%" if show_labels and v != 0 else "" for v in y_vals],
                textposition='inside',
                textfont=dict(size=label_size, color="white"),
                constraintext='none'
            ), row=1, col=i+1)
        
        for yr in years:
            total = sum(vals[f].get(yr, 0) for f in fields)
            if total > 0:
                fig.add_annotation(
                    x=f"{yr}YE", y=total + 2,
                    text=f"<b>{total:.1f}%</b>",
                    showarrow=False,
                    font=dict(size=label_size, color="#222"),
                    row=1, col=i+1
                )
    
    fig.update_layout(
        barmode='stack',
        height=550,
        paper_bgcolor='rgba(0,0,0,0)',
        plot_bgcolor='rgba(0,0,0,0)',
        margin=dict(t=60, b=80, l=10, r=10),
        legend=dict(orientation="h", yanchor="top", y=-0.15, xanchor="center", x=0.5)
    )
    fig.update_yaxes(tickformat=".0f", title_text="占保险服务收入比例", row=1, col=1)
    return fig, (0,0,0)

def create_ratio_trend_chart(df, cos, factor_list, ratio_field, title_prefix,
                             divisor=1, unit_label="", highlight_co="无",
                             show_labels=True, line_width=2, ratio_first=False):
    """
    每个指标生成一个独立折线图，每条线代表一家公司。
    只绘制有分母（保险服务收入或保险业务收入）的年份，确保比率有意义。
    """
    raw = df.copy()
    raw['报告年份'] = raw['报告年份'].astype(str).str.replace('.0', '', regex=False)
    raw['公司'] = raw['公司'].astype(str).str.strip()
    raw['字段名'] = raw['字段名'].astype(str).str.strip()

    needed = factor_list + [ratio_field, '保险服务收入', '保险业务收入']
    raw = raw[raw['公司'].isin(cos) & raw['字段名'].isin(needed)]

    years_raw = raw['报告年份'].unique()
    years = sorted(
        [y for y in years_raw if str(y).strip() not in ['', 'nan', 'NaN', 'None']],
        key=lambda x: int(x)
    )
    if not years:
        return []

    # 计算分母
    denom = {}
    for co in cos:
        for yr in years:
            v = raw[(raw['公司']==co) & (raw['报告年份']==yr) & (raw['字段名']=='保险服务收入')]['(百万)人民币']
            if not v.empty and v.iloc[0] != 0:
                denom[(co, yr)] = v.iloc[0]
            else:
                v2 = raw[(raw['公司']==co) & (raw['报告年份']==yr) & (raw['字段名']=='保险业务收入')]['(百万)人民币']
                if not v2.empty and v2.iloc[0] != 0:
                    denom[(co, yr)] = v2.iloc[0]
                else:
                    denom[(co, yr)] = None

    # 构建数据
    all_items = [ratio_field] + factor_list if ratio_first else factor_list + [ratio_field]
    metric_data = {item: {co: {} for co in cos} for item in all_items}

    for co in cos:
        for yr in years:
            den = denom.get((co, yr))
            for f in factor_list:
                if den is None:
                    metric_data[f][co][yr] = np.nan
                else:
                    s = raw[(raw['公司']==co) & (raw['报告年份']==yr) & (raw['字段名']==f)]['(百万)人民币']
                    val = s.iloc[0] if not s.empty else 0
                    metric_data[f][co][yr] = val / den if den != 0 else np.nan
            s_ratio = raw[(raw['公司']==co) & (raw['报告年份']==yr) & (raw['字段名']==ratio_field)]['(百万)人民币']
            metric_data[ratio_field][co][yr] = s_ratio.iloc[0] if not s_ratio.empty else np.nan

    # 颜色：使用自定义强对比色（从 KPMG 色板中精选）
    custom_colors = ["#00338D", "#510DBC", "#FD349C", "#00C0AE", "#FB8E7E", "#7213EA"]
    company_colors = {co: custom_colors[i % len(custom_colors)] for i, co in enumerate(cos)}
    marker_symbols = ['circle', 'square', 'diamond', 'triangle-up', 'star', 'pentagon', 'x', 'cross']

    figs = []
    for item in all_items:   # 按调整后的顺序循环
        fig = go.Figure()
        is_ratio = (item == ratio_field)
        title = f"{title_prefix} - {item}" if title_prefix else item
        if is_ratio:
            title = f"{title_prefix} - {item}（综合）" if title_prefix else f"{item}（综合）"

        all_y = []

        for idx, co in enumerate(cos):
            data = metric_data[item][co]
            valid_years = [yr for yr in years if not np.isnan(data.get(yr, np.nan))]
            if not valid_years:
                continue
            x_vals = valid_years
            y_vals = [data[yr] for yr in valid_years]
            all_y.extend(y_vals)

            color = company_colors[co]
            is_hl = (str(co).strip() == str(highlight_co).strip())
            width = line_width + 2 if is_hl else line_width
            marker_size = 10 if is_hl else 8
            symbol = marker_symbols[idx % len(marker_symbols)]

            fig.add_trace(go.Scatter(
                x=x_vals,
                y=y_vals,
                mode='lines+markers' + ('+text' if show_labels else ''),
                name=co,
                line=dict(width=width, color=color),
                marker=dict(size=marker_size, color=color, symbol=symbol,
                            line=dict(width=1.5, color='white')),
            ))

        if not all_y:
            fig.update_layout(title=title, height=400)
            figs.append(fig)
            continue

        y_min = min(all_y)
        y_max = max(all_y)
        range_width = y_max - y_min
        if range_width == 0:
            range_width = 1
        padding = range_width * 0.1
        y_lower = y_min - padding
        y_upper = y_max + padding

        fig.update_layout(
            title=title,
            xaxis_title="年份",
            yaxis_title="比率",
            xaxis=dict(
                type='category',
                tickmode='array',
                tickvals=years,
                ticktext=years,
                gridcolor='rgba(0,0,0,0)',   # 去掉网格线
                gridwidth=0,
            ),
            yaxis=dict(
                tickformat=".2%",
                range=[y_lower, y_upper],
                gridcolor='rgba(0,0,0,0)',   # 去掉网格线
                gridwidth=0,
                zeroline=True,
                zerolinecolor='gray',
                zerolinewidth=1,
            ),
            height=400,
            margin=dict(t=50, b=50, l=40, r=20),
            legend=dict(orientation="h", yanchor="top", y=-0.15, xanchor="center", x=0.5,
                        font=dict(size=11)),
            plot_bgcolor='rgba(0,0,0,0)',
            paper_bgcolor='rgba(0,0,0,0)'
        )
        figs.append(fig)

    return figs


def create_nonlife_six_dimensional_charts(df_raw, target_year, cos, divisor=1, unit_label="百万元", 
                                          highlight_co="无", label_size=11, show_labels=False, dot_size=11):
    df = df_raw.copy()
    df[["字段名", "公司", "报告年份"]] = df[["字段名", "公司", "报告年份"]].astype(str).apply(
        lambda x: x.str.strip().str.replace(".0", "", regex=False))
    
    needed = ["净利润", "期初股东权益", "期末股东权益", "总资产", "投资收益率", "综合成本率", "保险业务收入"]
    needed_prev = ["保险业务收入"]
    
    df = df[(df["报告年份"] == str(target_year)) & df["公司"].isin(cos) & df["字段名"].isin(needed)].drop_duplicates(
        subset=["公司", "报告年份", "字段名"])
    
    df_prev = df_raw[(df_raw["报告年份"] == str(int(target_year)-1)) & df_raw["公司"].isin(cos) & 
                     df_raw["字段名"].isin(needed_prev)].drop_duplicates(subset=["公司", "报告年份", "字段名"])
    
    if df.empty:
        return []
    
    a_cols = ["(百万)人民币", "(亿元)人民币", "(百万)原币", "(亿元)原币", "人民币", "原币", "数值", "值"]
    r_cols = ["(%)原币", "(%)", "百分比", "比例", "占比", "比率"]
    
    def get_v(r, is_ratio):
        for c in (r_cols + a_cols if is_ratio else a_cols + r_cols):
            if c in r.index and pd.notna(r[c]) and str(r[c]).strip().lower() not in ["", "nan", "none", "null", "-"]:
                s = str(r[c]).strip().replace(",", "")
                try:
                    v = float(s[:-1])/100.0 if s.endswith("%") else float(s)
                    return v/100.0 if is_ratio and ("%" in str(r[c]) or abs(v) > 1) else v
                except:
                    continue
        return np.nan
    
    records = [{"公司": r["公司"], "字段名": r["字段名"], "数值": get_v(r, r["字段名"] in ["投资收益率", "综合成本率"])} 
               for _, r in df.iterrows()]
    df_p = pd.DataFrame(records).dropna(subset=["数值"]).pivot_table(
        index="公司", columns="字段名", values="数值", aggfunc="first").reindex(cos)
    df_p.index.name = None
    
    def get_col(n):
        return pd.to_numeric(df_p[n], errors="coerce") if n in df_p.columns else pd.Series(np.nan, index=df_p.index)
    
    net_profit = get_col("净利润")
    eq_b = get_col("期初股东权益")
    eq_e = get_col("期末股东权益")
    ta = get_col("总资产")
    inv_return = get_col("投资收益率")
    cor = get_col("综合成本率")
    
    prem_curr = get_col("保险业务收入")
    prem_prev = pd.Series(np.nan, index=cos)
    for _, r in df_prev.iterrows():
        prem_prev[str(r["公司"]).strip()] = get_v(r, False)
    prem_growth = (prem_curr - prem_prev) / prem_prev.abs() if not prem_prev.empty else pd.Series(np.nan, index=cos)
    
    pd_data = pd.DataFrame({
        "公司": df_p.index,
        "净利润": net_profit,
        "净资产": eq_e,
        "保费增长率": prem_growth,
        "投资收益率": inv_return,
        "综合成本率": cor,
        "总资产": ta,
        "期初净资产": eq_b
    })
    
    pd_data["承保利润率"] = np.where((eq_b + eq_e) != 0, net_profit / ((eq_b + eq_e) / 2), np.nan)
    pd_data["财务杠杆率"] = np.where(ta != 0, eq_e / ta, np.nan)
    pd_data["净资产增长率"] = np.where(eq_b != 0, (eq_e - eq_b) / eq_b, np.nan)
    
    cfgs = [
        ("承保利润率", "承保利润率", "净资产", ".1%"),
        ("保费增长", "保费增长率", "净资产", ".1%"),
        ("财务杠杆", "财务杠杆率", "净资产", ".1%"),
        ("投资能力", "投资收益率", "净资产", ".1%"),
        ("财务稳定", "净资产增长率", "净资产", ".1%"),
        ("承保效益", "综合成本率", "净资产", ".1%")
    ]
    
    cmap = {co: px.colors.qualitative.Plotly[i % 10] for i, co in enumerate(cos)}
    hl = str(highlight_co).strip()
    rd = divisor or 1
    figs = []
    
    for title_text, y_col, x_col, y_fmt in cfgs:
        d_plt = pd_data[["公司", y_col, x_col]].copy().replace([np.inf, -np.inf], np.nan).dropna()
        
        display_y = y_col
        if y_col == "承保利润率":
            display_y = "承保利润率（净利润/净资产平均值）"
        elif y_col == "财务杠杆率":
            display_y = "财务杠杆率（净资产/总资产）"
        elif y_col == "承保效益":
            display_y = "综合成本率（COR，越低越好）"
        
        title_html = (f"<span style='font-size:13px'><b>{title_text}</b></span><br>"
                      f"<span style='font-size:11px;color:#666'>Y轴={display_y}，X轴={x_col}</span>")
        
        fig = go.Figure()
        
        if d_plt.empty:
            fig.update_layout(title=dict(text=title_html, x=0.02), height=250,
                              margin=dict(l=20, r=15, t=40, b=10),
                              paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)")
            figs.append(fig)
            continue
        
        d_plt["x_p"] = pd.to_numeric(d_plt[x_col], errors="coerce") / rd
        
        for _, r in d_plt.sort_values("公司").iterrows():
            co_name = str(r["公司"])
            is_hl = (co_name == hl)
            fig.add_trace(go.Scatter(
                x=[r["x_p"]], y=[r[y_col]],
                mode="markers+text" if show_labels else "markers",
                name=co_name,
                text=[co_name] if show_labels else None,
                textposition="top center",
                textfont=dict(size=label_size, color="#333"),
                marker=dict(
                    size=dot_size * 1.45 if is_hl else dot_size,
                    color=cmap.get(co_name, "#1f77b4"),
                    line=dict(color="white", width=1.8 if is_hl else 1.2),
                    opacity=0.95
                ),
                customdata=[[f"{r['x_p']:,.2f}", f"{r[y_col]:{y_fmt}}", f"{r[x_col]:,.2f}"]],
                hovertemplate=(f"<b>{co_name}</b><br>{x_col}: %{{customdata[0]}}<br>"
                              f"{y_col}: %{{customdata[1]}}<br>{x_col}: %{{customdata[2]}}<extra></extra>")
            ))
        
        fig.update_layout(
            title=dict(text=title_html, x=0.02),
            width=260, height=250,
            margin=dict(l=20, r=20, t=45, b=15),
            paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
            showlegend=False,
            hovermode="closest",
            xaxis=dict(showgrid=True, gridcolor="rgba(180,180,180,0.2)",
                      zeroline=True, zerolinecolor="rgba(150,150,150,0.3)"),
            yaxis=dict(tickformat=y_fmt, showgrid=True, gridcolor="rgba(180,180,180,0.2)",
                      zeroline=True, zerolinecolor="rgba(150,150,150,0.3)")
        )
        
        fig.add_hline(y=0, line_dash="dash", line_color="rgba(120,120,120,0.7)", line_width=1)
        fig.add_vline(x=0, line_dash="dash", line_color="rgba(120,120,120,0.7)", line_width=1)
        figs.append(fig)
    
    return figs
# ==========================================
# 登录后主系统界面（主逻辑）
# ==========================================
api_key = st.session_state.get('api_key', "")
base_url = st.session_state.get('base_url', "https://api.deepseek.com")
model_name = st.session_state.get('model_name', "deepseek-chat")
user_role = st.session_state.get('user_role', "普通用户")

# 顶部状态栏
top_col1, top_col2 = st.columns([8, 1])
with top_col1:
    st.markdown(
        f"<div class='no-print' style='color: #6c757d; font-size: 14px; margin-top: 8px;'>"
        f"当前身份：<b>{user_role}</b> | 欢迎使用财险数智分析平台</div>",
        unsafe_allow_html=True
    )
with top_col2:
    if st.button("退出登录"):
        st.session_state['logged_in'] = False
        st.rerun()

st.markdown("<hr style='margin-top: 5px; margin-bottom: 15px;'>", unsafe_allow_html=True)
st.markdown("<h1 class='no-print' style='font-weight:700; padding-top:10px;'>财险数智・年报处理平台</h1>", unsafe_allow_html=True)

# ---------- 9. 财险公司默认列表 ----------
DEFAULT_COMPANIES = [
    {"类别": "上市", "公司": "平安产险", "类型": "财险", "链接地址": "https://property.pingan.com/gongkaixinxipilu/nianduxinxipilubaogao.shtml"},
    {"类别": "上市", "公司": "人保产险", "类型": "财险", "链接地址": "https://property.picc.com.cn/gkxx/ndxx/"},
    {"类别": "上市", "公司": "太平产险", "类型": "财险", "链接地址": "https://caixian.cntaiping.com/info-ndxxpl/"},
    {"类别": "上市", "公司": "众安产险", "类型": "财险", "链接地址": "https://www.zhongan.com/channel/public/publicInfo_ndxx2018.html"},
    {"类别": "上市", "公司": "阳光产险", "类型": "财险", "链接地址": "https://www.4000-000-000.com/#/pltable?menuId=628&type=menu"},
    {"类别": "上市", "公司": "太保财险", "类型": "财险", "链接地址": "https://property.cpic.com.cn/xccbx/gkxxbl/ndxx/?subMenu=2&inSub=1"},
    {"类别": "上市", "公司": "平安产险（偿付能力报告）", "类型": "财险", "链接地址": "https://property.pingan.com/gongkaixinxipilu/changfunenglixinxipilubaogao.shtml"},
    {"类别": "上市", "公司": "人保产险（偿付能力报告）", "类型": "财险", "链接地址": "https://property.picc.com.cn/gkxx/zxxx/cfnl/"},
    {"类别": "上市", "公司": "太平产险（偿付能力报告）", "类型": "财险", "链接地址": "https://caixian.cntaiping.com/info-cfnljdbgzy/"},
    {"类别": "上市", "公司": "众安产险（偿付能力报告）", "类型": "财险", "链接地址": "http://zhongan.com/channel/public/publicInfo_cfnl2018.html"},
    {"类别": "上市", "公司": "阳光产险（偿付能力报告）", "类型": "财险", "链接地址": "https://www.4000-000-000.com/#/pltable?menuId=633&type=menu"},
    {"类别": "上市", "公司": "太保财险（偿付能力报告）", "类型": "财险", "链接地址": "https://property.cpic.com.cn/xccbx/gkxxbl/cfnlxxzq/?subMenu=4&inSub=3"},
    {"类别": "非上市", "公司": "华泰财险", "类型": "年报", "链接地址": "https://pc.ehuatai.com/annual_info.html"}
]

# ---------- 10. Step 0~8 Tabs ----------
if st.session_state['user_role'] == "项目组成员":
    tab0, tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8 = st.tabs([
        " 🌐 Step 0 ／ 官网年报监控 ",
        " 📑 Step 1 ／ 智能页码检索 ",
        " ⚡ Step 2 ／ 表格智能转换 ",
        " 📝 Step 3 ／ 目标表标准填报 ",
        " 🔍 Step 4 ／ 数据勾稽检查 ",
        " ⛓️‍💥 Step 5 ／ 多公司合并 ",
        " 📊 Step 6 ／ 自定义对标分析 ",
        " 🖼️ Step 7 ／ 公司级对标报告 ",
        " 📈 Step 8 ／ 行业分类统计分析 "
    ])
    
    # ----- Step 0：官网年报监控（财险版）-----
    with tab0:
        st.markdown("### 🌐 财险公司官网年报监控")
        
        col_t0_1, _ = st.columns([1, 2])
        with col_t0_1:
            target_year = st.number_input("📅 请选择监控年份", min_value=2010, max_value=2050, value=2025, step=1)
        
        st.markdown(f"""
        <div class="info-card green">
            <h4>功能说明</h4>
            <p>系统将自动扫描各财险公司官网，检测网页中是否出现 <b>{target_year}</b> 字样。您可以<b>双击下方表格修改网址</b>。</p>
        </div>
        """, unsafe_allow_html=True)
    
        if 'company_df' not in st.session_state:
            st.session_state.company_df = pd.DataFrame(DEFAULT_COMPANIES)
        
        st.markdown("#### 🏢 监控目标名单 (双击下方链接可修改)")
        edited_df = st.data_editor(
            st.session_state.company_df,
            column_config={
                "链接地址": st.column_config.TextColumn("🌐 网页链接地址 (支持双击编辑)", max_chars=1000, width="large"),
                "类别": st.column_config.TextColumn(disabled=True),
                "公司": st.column_config.TextColumn(disabled=True),
                "类型": st.column_config.TextColumn(disabled=True),
            },
            use_container_width=True,
            hide_index=True,
            key="company_data_editor_step0"
        )
        st.session_state.company_df = edited_df
    
        st.markdown("---")
        
        col_btn1, col_btn2 = st.columns([1, 1])
        with col_btn1:
            if not edited_df.empty:
                target_company = st.selectbox("🎯 单家快速检索：", options=edited_df['公司'].tolist())
                single_scan = st.button(f"🔍 检索【{target_company}】", use_container_width=True)
            else:
                st.warning("请先在表格中添加财险公司数据")
                single_scan = False
                target_company = None
        with col_btn2:
            st.write("")
            batch_scan = st.button("🚀 启动批量扫描", type="primary", use_container_width=True)
    
        if single_scan or batch_scan:
            if edited_df.empty:
                st.warning("⚠️ 请先在表格中添加财险公司数据。")
            else:
                tasks = edited_df.to_dict('records') if batch_scan else edited_df[edited_df['公司'] == target_company].to_dict('records')
                results = []
                my_bar = st.progress(0, text="准备启动扫描...")
                total_tasks = len(tasks)
                headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'}
    
                for index, row in enumerate(tasks):
                    company_name = row['公司']
                    url = str(row['链接地址'])
                    my_bar.progress((index + 1) / total_tasks, text=f"正在扫描 ({index+1}/{total_tasks}): {company_name}")
                    
                    status = "🔴 未更新 / 无法访问"
                    if url.lower().endswith('.pdf'):
                        status = "⚠️ 直接PDF链接 (需手动核实)"
                    elif "http" in url:
                        try:
                            response = requests.get(url, headers=headers, timeout=8)
                            response.encoding = response.apparent_encoding
                            soup = BeautifulSoup(response.text, 'html.parser')
                            page_text = soup.get_text()
                            if str(target_year) in page_text:
                                status = "🟢 极可能已更新!"
                        except Exception as e:
                            status = "🟡 网站拦截/超时，需手动查看"
                    else:
                        status = "无效链接"
                    
                    results.append({
                        "公司名称": company_name,
                        "监控状态": status,
                        "检测结果描述": f"网页中已检索到 {target_year} 字样" if "🟢" in status else "未发现关键字",
                        "直达链接": url
                    })
                    time.sleep(0.3)
    
                my_bar.empty()
                st.success(f"🎉 {target_year}年度检测扫描完成！")
                df_result = pd.DataFrame(results)
                st.data_editor(
                    df_result,
                    column_config={
                        "直达链接": st.column_config.LinkColumn("点击前往网页"),
                        "监控状态": st.column_config.TextColumn("状态", width="medium")
                    },
                    hide_index=True,
                    use_container_width=True,
                    key="scan_result_display"
                )
                st.info("💡 提示：请点击标记为 🟢 的公司链接下载 PDF，下载后请前往 [Step 1] 进行页码定位。")
    
    # ----- Step 1：智能页码检索（财险版）-----
    with tab1:
        st.markdown("### 📑 智能页码检索")
        uploaded_file = st.file_uploader(
            "拖拽或选择一份已经下载好的年报 PDF 文件",
            type="pdf",
            help="推荐上传一份财险公司的完整年报"
        )
    
        if uploaded_file:
            if 'pdf_bytes' not in st.session_state or st.session_state.get('pdf_name') != uploaded_file.name:
                st.session_state['pdf_bytes'] = uploaded_file.read()
                st.session_state['pdf_name'] = uploaded_file.name
                if 'found_pages' in st.session_state:
                    del st.session_state['found_pages']
                if 'edited_pages' in st.session_state:
                    del st.session_state['edited_pages']
            
            pdf_bytes = st.session_state['pdf_bytes']
            doc = fitz.open(stream=pdf_bytes, filetype="pdf")
            total_pages = len(doc)
            
            st.caption(f"当前加载文件：{uploaded_file.name}　|　文档共 {total_pages} 页")
            st.markdown("---")
            
            col_left, col_spacer, col_right = st.columns([1, 0.05, 1.2])
    
            with col_left:
                st.markdown("#### 检索目标设定")
                # 公司选择：决定"保险合同负债及资产"的差异化识别规则
                company_options = ["平安产险", "人保产险", "太平产险", "众安产险", "阳光产险", "太保财险", "其他/自动识别"]
                default_company_idx = 0
                if target_company:
                    for _ci, _co in enumerate(company_options):
                        if _co in target_company or target_company in _co:
                            default_company_idx = _ci
                            break
                selected_company = st.selectbox("🏢 当前年报对应公司：", options=company_options, index=default_company_idx, help="用于精准匹配《保险合同负债及资产》表，请选择与上传PDF一致的公司")
                target_tables = st.multiselect(
                    "请选择需要定位的报表：",
                    [
                        "保险合同负债及资产",
                        "公司利润表",
                        "业务及管理费",
                        "公司资产负债表",
                        "主要经营指标",
                        "保险产品经营信息"   # 新增
                    ],
                    default=["保险合同负债及资产", "公司利润表", "公司资产负债表"]
                )
                
                st.markdown("")
                btn_search = st.button("启动智能检索", use_container_width=True)
                
                if btn_search:
                    current_api_key = st.session_state.get('api_key', "").strip()
                    if not current_api_key:
                        st.error("⚠️ 未检测到 API Key！请刷新页面返回登录界面填写。")
                    elif not target_tables:
                        st.error("请至少选择一张报表。")
                    else:
                        with st.spinner("需要一些时间，请稍等~"):
                            try:
                                result = ai_find_pages(pdf_bytes, current_api_key, target_tables, base_url, model_name, target_company if single_scan else "")
                                st.write("当前公司名称:", target_company)
                                main_table_pairs = [
                                    ("资产负债表（合并）", "资产负债表（公司）"),
                                    ("利润表（合并）", "利润表（公司）"),
                                    ("现金流量表（合并）", "现金流量表（公司）"),
                                    ("股东/所有者权益变动表（合并）", "股东/所有者权益变动表（公司）")
                                ]
                                for merge_key, company_key in main_table_pairs:
                                    merge_pages = result.get(merge_key, [0])
                                    company_pages = result.get(company_key, [0])
                                    if merge_pages != [0] and company_pages == [0]:
                                        result[company_key] = merge_pages
                                    elif company_pages != [0] and merge_pages == [0]:
                                        result[merge_key] = company_pages
                                st.session_state['found_pages'] = result
                                st.success("检索完成！请在下方核对物理页码。")
                            except Exception as e:
                                st.error(f"引擎出现异常：{e}")
    
                if 'found_pages' in st.session_state:
                    st.markdown("---")
                    st.markdown("#### 结果核对")
                    st.caption("提示：若表格跨页，请以英文逗号分隔页码（如 64, 65）。可结合右侧预览进行校准。")
                    
                    edited_pages = {}
                    for table_name in target_tables:
                        page_data = st.session_state['found_pages'].get(table_name, [0])
                        if isinstance(page_data, int):
                            page_data = [page_data]
                        str_val = ", ".join(map(str, page_data))
                        user_input = st.text_input(f"{table_name}", value=str_val, key=f"page_{table_name}")
                        try:
                            edited_pages[table_name] = [int(x.strip()) for x in user_input.split(",") if x.strip().isdigit()]
                        except:
                            edited_pages[table_name] = page_data
                    
                    st.session_state['edited_pages'] = edited_pages
                    
                    st.markdown("")
                    if st.button("确认页码，进入下一步", use_container_width=True):
                        st.success("页码已确认！请前往 Step 2 进行表格转换。")
    
            with col_right:
                st.markdown("#### 页面预览")
                if 'found_pages' in st.session_state and 'edited_pages' in st.session_state:
                    table_to_view = st.selectbox("选择要预览的报表：", options=list(st.session_state['edited_pages'].keys()))
                    pages_to_preview = st.session_state['edited_pages'].get(table_to_view, [0])
                    if not pages_to_preview:
                        pages_to_preview = [0]
                    
                    if len(pages_to_preview) > 1:
                        current_page = st.radio("该报表包含多页，请切换预览：", options=pages_to_preview, horizontal=True)
                    else:
                        current_page = pages_to_preview[0]
                    
                    preview_idx = current_page - 1
                    if 0 <= preview_idx < total_pages:
                        page = doc.load_page(preview_idx)
                        pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
                        st.image(pix.tobytes("png"), caption=f"当前预览：第 {current_page} 页 / 共 {total_pages} 页", use_column_width=True)
                    elif current_page == 0:
                        st.info("尚未识别到页码，请在左侧输入框中手动填入。")
                    else:
                        st.warning("该页码超出了文档总页数，请检查左侧修改。")
                else:
                    st.info("上传文件并启动检索后，此处将显示对应页面。")
    # ----- Step 2：表格智能转换 -----
    with tab2:
        st.markdown("### ⚡ 表格智能转换")
        st.markdown("""
        <div class="info-card blue">
            <h4>功能说明</h4>
            <p>系统将调用大模型对多页 PDF 进行网格化对齐重构，并自动拼接同表跨页数据，生成标准化 Excel。<b>如果出现提取失败或错位</b>，请开启下方的【图片扫描模式】重试。</p>
        </div>
        """, unsafe_allow_html=True)
        
        st.markdown("")
        
        st.markdown("#### PDF类型选择")
        use_vision = st.toggle("📸 开启图片扫描模式 (适用于扫描版、纯图片PDF)", value=False)
        if use_vision:
            st.caption("提示：图片扫描模式请确保您登录时选择的模型支持视觉功能（如 GPT-4o, 阿里云通义千问等）。DeepSeek 暂不支持直接读图。")
        st.markdown("---")
        
        if 'edited_pages' not in st.session_state or 'pdf_bytes' not in st.session_state:
            st.warning("⚠️ 请先在 Step 1 中上传文件并完成【页码确认】。")
            st.button("开始提取结构化数据", disabled=True, use_container_width=True)
        else:
            current_api_key = st.session_state.get('api_key', "").strip()
            current_base_url = st.session_state.get('base_url', "https://api.deepseek.com")
            current_model_name = st.session_state.get('model_name', "deepseek-chat")
            
            can_run = True
            if not current_api_key:
                st.error("⚠️ 未检测到 API Key！请刷新页面返回登录界面填写授权密钥。")
                can_run = False
    
            if not can_run:
                st.button("开始提取结构化数据", disabled=True, use_container_width=True)
            else:
                valid_tasks = {k: v for k, v in st.session_state['edited_pages'].items() if v != [0]}
                
                if not valid_tasks:
                    st.warning("⚠️ 没有找到任何有效的页码，无需提取。")
                else:
                    if st.button("开始提取", use_container_width=True):
                        with st.spinner("☕ 趁现在喝口水吧，正在后台拼命打工中..."):
                            extracted_sheets = {}
                            global_unit = "未能自动提取，需人工核对"
                            all_tasks = []
                            for table_name, pages in valid_tasks.items():
                                for page_num in pages:
                                    all_tasks.append({"table": table_name, "page": page_num})
                            
                            total_pages_to_process = len(all_tasks)
                            pages_done = 0
                            
                            progress_bar = st.progress(0)
                            status_box = st.status(f"正在转化 {total_pages_to_process} 个任务...", expanded=True)
                            temp_results = {table_name: {} for table_name in valid_tasks.keys()}
                            
                            with status_box:
                                workers = 2 if use_vision else 5
                                
                                with ThreadPoolExecutor(max_workers=workers) as executor:
                                    future_to_task = {}
                                    
                                    for task in all_tasks:
                                        if use_vision:
                                            future = executor.submit(
                                                extract_single_page_vision, 
                                                st.session_state['pdf_bytes'], 
                                                task["page"], 
                                                task["table"], 
                                                current_api_key,
                                                current_base_url,
                                                current_model_name
                                            )
                                        else:
                                            future = executor.submit(
                                                extract_single_page, 
                                                st.session_state['pdf_bytes'], 
                                                task["page"], 
                                                task["table"], 
                                                current_api_key,
                                                current_base_url,
                                                current_model_name
                                            )
                                        future_to_task[future] = task
                                    
                                    for future in as_completed(future_to_task):
                                        task = future_to_task[future]
                                        t_name = task["table"]
                                        p_num = task["page"]
                                        try:
                                            df, raw_text = future.result()
                                            if df is not None and not df.empty:
                                                temp_results[t_name][p_num] = df
                                                st.write(f"✅ [{t_name}] - 第 {p_num} 页提取完成！")
                                                if global_unit == "未能自动提取，需人工核对":
                                                    unit = get_report_unit(raw_text)
                                                    if unit:
                                                        global_unit = unit
                                                        st.write(f"&nbsp;&nbsp;&nbsp;&nbsp;🔎 识别到该公司报表单位：【{global_unit}】")
                                            else:
                                                st.write(f"⚠️ [{t_name}] - 第 {p_num} 页未提取到有效数据。")
                                        except Exception as e:
                                            st.error(f"❌ [{t_name}] - 第 {p_num} 页提取失败: {str(e)}")
                                        
                                        pages_done += 1
                                        progress_bar.progress(pages_done / total_pages_to_process)
                            
                            for table_name, pages in valid_tasks.items():
                                table_dfs = []
                                for p_num in pages: 
                                    if p_num in temp_results[table_name]:
                                        table_dfs.append(temp_results[table_name][p_num])
                                
                                if table_dfs:
                                    merged_df = pd.concat(table_dfs, ignore_index=True)
                                    safe_sheet_name = re.sub(r'[\\/*?:\[\]]', '', table_name)[:30]
                                    extracted_sheets[safe_sheet_name] = merged_df
                                        
                            status_box.update(label="🎉 所有任务提取完成！", state="complete", expanded=False)
                            st.session_state['extracted_data'] = extracted_sheets
                            st.session_state['global_unit'] = global_unit
                        
        if 'extracted_data' in st.session_state:
            st.markdown("---")
            st.markdown("#### 📊 提取结果预览")
            extracted_data = st.session_state['extracted_data']
            company_name = st.session_state.get('pdf_name', '未命名').replace(".pdf", "")
            
            unit_df = pd.DataFrame([
                {"项目": "源文件名称", "信息": company_name},
                {"项目": "探测到的报表单位", "信息": st.session_state.get('global_unit', '')}
            ])
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                unit_df.to_excel(writer, sheet_name="基本信息_单位", index=False)
                for sheet_name, df in extracted_data.items():
                    df.to_excel(writer, sheet_name=sheet_name, index=False, header=False)
            excel_data = output.getvalue()
            
            st.download_button(
                label="⬇️ 一键下载结构化提取表 (Excel)",
                data=excel_data,
                file_name=f"{company_name}_数据提取.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
            
            st.markdown("")
            sheet_names = list(extracted_data.keys())
            if sheet_names:
                tabs = st.tabs(sheet_names)
                for i, tab in enumerate(tabs):
                    with tab:
                        st.dataframe(extracted_data[sheet_names[i]], use_container_width=True)
   
    # ----- Step 3：目标表标准填报 -----
    with tab3:
        st.markdown("### 📝 目标表标准填报")
        st.markdown('<div class="info-card pink"><h4>功能说明</h4><p>AI 将自动寻找科目数据填充数值，生成 Excel 计算公式。支持内置模板或上传自定义模板。</p></div>', unsafe_allow_html=True)
        
        # 从财险公司列表获取类型映射
        COMPANY_TYPE_MAP = {item["公司"]: item["类别"] for item in DEFAULT_COMPANIES}
        
        col_t1, col_t2 = st.columns([1, 1])
        with col_t1: use_default = st.toggle("使用系统默认模板", value=True, help="开启后直接使用内置的财险样例表")
        
        template_file = (
"https://github.com/Polly031021/Annual-Report-System/raw/main/%E6%8C%87%E6%A0%87%E6%A0%B7%E4%BE%8B%E8%A1%A8.xlsx" 
            if use_default 
            else st.file_uploader("上传自定义目标表模板 (.xlsx)", type="xlsx", key="unique_template_uploader")
        )
    
        if template_file:
            COL_COMPANY, COL_CATEGORY, COL_FIELD_NAME, COL_FIELD_TYPE, COL_NOTE, COL_RULE, COL_CO_TYPE = "公司", "类别", "字段名", "字段类型", "注释", "计算规则", "公司类型"
    
            if 'extracted_data' not in st.session_state:
                st.warning("⚠️ 尚未找到提取的数据，请先完成 Step 1 和 Step 2。")
            elif st.button("启动智能填报与公式生成", use_container_width=True):
                with st.status("此过程略慢，请稍等~", expanded=True) as status_box:
                    st.write("📄 正在获取并解析模板表结构...")
                    
                    try:
                        if use_default:
                            import requests, io
                            res = requests.get(template_file, timeout=15)
                            res.raise_for_status()
                            template_df = pd.read_excel(io.BytesIO(res.content))
                        else:
                            template_df = pd.read_excel(template_file)
                    except Exception as e:
                        st.error(f"❌ 读取模板失败，请检查网络连接或文件格式：{e}")
                        st.stop()
                    
                    year_cols = sorted([c for c in template_df.columns if re.search(r'20\d{2}', str(c))])
                    if len(year_cols) < 2:
                        st.error("❌ 错误：模板中未能自动识别出两个或以上的年份数据列！请检查表头规范。")
                        st.stop()
                    
                    col_prev, col_curr = year_cols[-2], year_cols[-1]
                    st.session_state['col_prev'] = col_prev
                    st.session_state['col_curr'] = col_curr
                    
                    raw_name = st.session_state.get('pdf_name', '未命名').replace(".pdf", "")
                    company_short = re.sub(r'202\d年.*', '', raw_name).strip()
                    matched_type = next((c_type for c_name, c_type in COMPANY_TYPE_MAP.items() if c_name[:2] in company_short or company_short[:2] in c_name), "其他")
                    working_df = template_df.copy()   # ⬅️ 添加这一行
                    if COL_COMPANY in working_df.columns: working_df[COL_COMPANY] = company_short
                    if COL_CO_TYPE in working_df.columns: working_df[COL_CO_TYPE] = matched_type
                    
                    st.write("正在准备目标指标清单...")
                    input_items = working_df[working_df[COL_FIELD_TYPE].astype(str).str.strip() == "输入"]
                    ai_target_list = []
                    for _, r in input_items.drop_duplicates(subset=[COL_FIELD_NAME]).iterrows():
                        item = {
                            "类别": str(r.get(COL_CATEGORY, "")),
                            "标准字段名": str(r.get(COL_FIELD_NAME, "")),
                            "别名参考": str(r.get(COL_NOTE, "")) if pd.notna(r.get(COL_NOTE)) else "",
                            "提取说明": str(r.get(COL_NOTE, "")) if pd.notna(r.get(COL_NOTE)) else ""
                        }
                        ai_target_list.append(item)                    
                        
                    st.write("正在分析上下文...")
                    extracted_data = st.session_state['extracted_data']
                    context_text = "".join([f"\n[表名: {name}]\n{df.to_csv(index=False, sep='|')}\n" for name, df in extracted_data.items()])

                    # =========================================================
                    # 🌟 财险版系统提示词（已改造，强调提取说明）
                    # =========================================================
                    SYSTEM_PROMPT = """你是一个资深的财险精算审计专家。任务：将 PDF 提取的财务明细精准填入目标底稿，并严格区分上一年度和最新年份。

【通用执行准则：绝对原样提取】
1. ⚠️ 绝对指令：除了特殊的加总需求外，所有指标必须【原封不动地照抄】原文里的文本！
- 必须保留括号（表示负数）、逗号、正负号。例如原文是 "(295,992.00)"，必须输出为 "(295,992.00)"。
- 严禁擅自删除符号、严禁自行进行四舍五入。
2. 别名匹配：若标准名找不到，请查看"别名参考"列，或利用你的精算知识寻找同义词。
3. 空值处理：若原文为"—"或"无"请输出 "0"。若完全找不到该指标，请输出 null。
4. 自定义提取说明（最高优先级）：对于每个指标，如果“提取说明”不为空，请严格按照该说明提取数据。说明中可能包含“两个值相加”、“四个值相加”等操作，你需要根据说明从原文中定位相关数值并进行算术运算后填入。在匹配表格时，如果表格标题不明确，请根据表格内容中的关键描述词（如“分出的再保险合同”、“保险合同负债”等）来识别正确的表格。对于行匹配，请使用包含（contains）方式，即只要行文本中包含您指定的关键词即可，不必完全一致。如果说明为空，则使用“别名参考”或默认知识进行匹配。在定位数值时，务必以该行与指定列标题交叉处的单元格为准，不要取该列下方或其他行的汇总数。在提取数值时，务必严格区分“已发生赔款负债”列与“合计”列。只取指定列在该行的单元格值，不要取任何合计列的值，除非注释中明确要求取合计。
5. 在匹配表格时，如果表格标题不明确，请根据表格内容中的关键描述词（如“分出的再保险合同”、“保险合同负债”等）来识别正确的表格。对于行匹配，请使用包含（contains）方式，即只要行文本中包含您指定的关键词即可，不必完全一致。

【特殊指令：财险利润表科目映射】
当提取财险利润表科目时，请使用以下映射规则：
1. 【保险业务收入】：提取"保险业务收入"或"已赚保费"或"签单保费"。
2. 【分出保费】：提取"分出保费"或"分保费用"。
3. 【提取未到期责任准备金】：提取"未到期责任准备金"或"UPR"。
4. 【已发生赔款】：提取"已发生赔款"或"赔付支出"或"赔款支出"。
5. 【手续费及佣金支出】：提取"手续费及佣金支出"。
6. 【业务及管理费】：提取"业务及管理费"或"管理费用"。
7. 【承保利润】：若无法直接提取，可用"保险业务收入 - 分出保费 - 提取未到期责任准备金 - 已发生赔款 - 手续费及佣金支出 - 业务及管理费"计算。

【特殊指令：财险资产负债表科目映射】
当提取财险资产负债表科目时，请使用以下映射规则：
1. 【应收保费】：提取"应收保费"。
2. 【应收分保账款】：提取"应收分保账款"或"应收分保款项"。
3. 【应付保费】：提取"应付保费"或"预收保费"。
4. 【未到期责任准备金】：提取"未到期责任准备金"。
5. 【未决赔款准备金】：提取"未决赔款准备金"或"IBNR"。
6. 【总资产】：提取"总资产"或"资产总计"。
7. 【总负债】：提取"总负债"或"负债合计"。
8. 【股东权益】：提取"股东权益"或"所有者权益"或"净资产"。

【特殊指令：财险综合成本率拆解】
当提取综合成本率相关科目时：
1. 【综合成本率】：若有直接披露，优先提取；否则用"赔付率 + 费用率"计算。
2. 【综合赔付率】：提取"赔付率"或"综合赔付率"。
3. 【综合费用率】：提取"费用率"或"综合费用率"。

【时间维度与排版防反转绝对指令】
中国企业财报的标准排版规则是：【左边的数据列为当期（最新年份），右边的数据列为上期（历史年份）】。
- 当填报要求中包含"较晚/最新年份"（如当期/本年/年末）的键时：必须从原表表头的"本期"、"本年"、"年末"或【最左侧的数据列】取数。
- 当填报要求中包含"较早/历史年份"（如上期/上年/年初）的键时：必须从原表表头的"上期"、"上年"、"年初"或【紧挨着它的右侧数据列】取数。
⚠️ 警告：大语言模型极易犯"线性思维"错误（误以为小年份排在左边，大年份排在右边）。要求你必须打破定势，仔细核对表头！永远记住：左边的列才是最新当期！

【提取说明优先级】
- 如果“提取说明”字段有内容，必须优先按照该说明执行，忽略其他默认映射。
- 如果“提取说明”为空，则按“别名参考”或标准映射提取。

仅输出合法的 JSON 格式，严禁带有任何 Markdown 标记或文字说明。
格式要求：
- 顶层键为“公司名称”（与模板表中的“公司”列一致）。
- 每个公司下为字段名到数值的映射。
- 每个字段包含 "prev"（上一年）和 "curr"（最新年）两个键。
格式示例：{"平安产险": {"保险服务收入-车险": {"prev": "0", "curr": "228494928861"}, "营业利润-车险": {"prev": "0", "curr": "123456"}}, "人保产险": {...}}"""

                    st.write("正在呼叫大模型进行语义映射与精准取数...")
                    
                    current_api_key = st.session_state.get('api_key', "").strip()
                    current_base_url = st.session_state.get('base_url', "https://api.deepseek.com")
                    current_model_name = st.session_state.get('model_name', "deepseek-chat")
                    
                    client = OpenAI(api_key=current_api_key, base_url=current_base_url)
                    def process_final_val(val):
                        """将 AI 返回的数值字符串转为浮点数，支持括号负数、逗号分隔、空值处理"""
                        if val is None:
                            return 0.0
                        if isinstance(val, (int, float)):
                            return float(val)
                        s = str(val).strip()
                        if s in ['', 'null', 'None', 'nan', 'NaN']:
                            return 0.0
                        # 处理 (123,456.78) -> -123456.78
                        if s.startswith('(') and s.endswith(')'):
                            s = '-' + s[1:-1]
                        s = s.replace(',', '').replace(' ', '')
                        try:
                            return float(s)
                        except ValueError:
                            return 0.0
                    try:
                        # 收集所有需要提取的公司
                        companies_to_extract = list(set([
                            str(r.get(COL_COMPANY, '')).strip()
                            for _, r in working_df.iterrows()
                            if str(r.get(COL_COMPANY, '')).strip() != ''
                        ]))
                    
                        # 构建用户消息，明确列出每个指标的提取说明
                        user_content = f"""目标公司列表：{companies_to_extract}
                    
                    目标指标列表（含提取说明）：
                    {json.dumps(ai_target_list, ensure_ascii=False, indent=2)}
                    
                    请按照以下要求提取数据：
                    1. 对于每个指标，如果“提取说明”不为空，必须严格按照说明中的描述进行提取。
                    2. 如果“提取说明”为空，则根据“别名参考”或你的精算知识寻找匹配项。
                    3. 注意时间维度：最新年份对应左侧数据列，历史年份对应右侧数据列。
                    4. 输出格式必须为合法 JSON，顶层键为公司名称，每个公司下是字段名到 {{"prev": ..., "curr": ...}} 的映射。
                    
                    数据内容：
                    {context_text}"""
                    
                        response = client.chat.completions.create(
                            model=current_model_name,
                            messages=[
                                {"role": "system", "content": SYSTEM_PROMPT},
                                {"role": "user", "content": user_content}
                            ],
                            temperature=0.0
                        )
                        ai_res = response.choices[0].message.content.strip()
                        ai_res = re.sub(r'^```(json)?\n?', '', ai_res, flags=re.MULTILINE).replace("```", "").strip()
                        ai_data = json.loads(ai_res)
                        st.write("✅ 数据映射成功！正在回填表格...")
                        
                        # ----- 1. 回填输入数据（按公司和字段名双重匹配） -----
                        for idx, row in working_df.iterrows():
                            ftype = str(row.get(COL_FIELD_TYPE, "")).strip()
                            if ftype != "输入":
                                continue
                            fname = str(row.get(COL_FIELD_NAME, "")).strip()
                            if not fname:
                                continue
                            company_name = str(row.get(COL_COMPANY, "")).strip()
                            
                            if company_name:
                                # 有公司名，精确匹配
                                if company_name in ai_data and fname in ai_data[company_name]:
                                    item_data = ai_data[company_name][fname]
                                    if isinstance(item_data, dict):
                                        prev_val = process_final_val(item_data.get("prev"))
                                        curr_val = process_final_val(item_data.get("curr"))
                                        # 对“获取费用”和“维持费用”取绝对值
                                        if fname in ["获取费用", "维持费用"]:
                                            if isinstance(prev_val, (int, float)):
                                                prev_val = abs(prev_val)
                                            if isinstance(curr_val, (int, float)):
                                                curr_val = abs(curr_val)
                                        working_df.at[idx, col_prev] = prev_val
                                        working_df.at[idx, col_curr] = curr_val
                            else:
                                # 没有公司名，从所有公司中查找该字段
                                found_companies = [co for co, data in ai_data.items() if fname in data]
                                if len(found_companies) == 1:
                                    co = found_companies[0]
                                    item_data = ai_data[co][fname]
                                    if isinstance(item_data, dict):
                                        prev_val = process_final_val(item_data.get("prev"))
                                        curr_val = process_final_val(item_data.get("curr"))
                                        if fname in ["获取费用", "维持费用"]:
                                            if isinstance(prev_val, (int, float)):
                                                prev_val = abs(prev_val)
                                            if isinstance(curr_val, (int, float)):
                                                curr_val = abs(curr_val)
                                        working_df.at[idx, col_prev] = prev_val
                                        working_df.at[idx, col_curr] = curr_val
                                elif len(found_companies) > 1:
                                    # 多个公司都有该字段，默认取第一个公司的值（可根据需要调整）
                                    co = found_companies[0]
                                    item_data = ai_data[co][fname]
                                    if isinstance(item_data, dict):
                                        prev_val = process_final_val(item_data.get("prev"))
                                        curr_val = process_final_val(item_data.get("curr"))
                                        if fname in ["获取费用", "维持费用"]:
                                            if isinstance(prev_val, (int, float)):
                                                prev_val = abs(prev_val)
                                            if isinstance(curr_val, (int, float)):
                                                curr_val = abs(curr_val)
                                        working_df.at[idx, col_prev] = prev_val
                                        working_df.at[idx, col_curr] = curr_val
                                # 如果找不到，则留空
                        # ----- 2. 计算“计算”类型字段的值（用于预览）-----
                        # 构建字段名->数值的映射（用于计算）
                        field_prev_vals = {}
                        field_curr_vals = {}
                        for _, r in working_df.iterrows():
                            fname = str(r.get(COL_FIELD_NAME, "")).strip()
                            if fname:
                                field_prev_vals[fname] = r[col_prev] if pd.notna(r[col_prev]) else 0.0
                                field_curr_vals[fname] = r[col_curr] if pd.notna(r[col_curr]) else 0.0
                        
                        # 遍历每一行，对“计算”类型执行计算
                        for idx, row in working_df.iterrows():
                            ftype = str(row.get(COL_FIELD_TYPE, "")).strip()
                            if ftype != "计算":
                                continue
                            rule = str(row.get(COL_RULE, "")).strip()
                            if not rule or rule in ["nan", "None", ""]:
                                continue
                            # 替换 [字段名] 为数值
                            def replace_field(match):
                                field = match.group(1)
                                # 注意：计算规则中可能包含当前年份或其他字段，我们分别计算prev和curr
                                return f"field_prev_vals['{field}']"  # 但需要在eval时使用局部变量
                            # 更安全的方法是构造表达式字符串，然后用eval计算
                            # 我们分别计算prev和curr
                            try:
                                # 提取所有 [字段名]
                                fields = re.findall(r'\[(.*?)\]', rule)
                                # 构造prev表达式
                                prev_expr = rule
                                curr_expr = rule
                                for f in fields:
                                    prev_expr = prev_expr.replace(f"[{f}]", str(field_prev_vals.get(f, 0.0)))
                                    curr_expr = curr_expr.replace(f"[{f}]", str(field_curr_vals.get(f, 0.0)))
                                # 安全计算（只允许基本运算）
                                # 使用eval，但限制命名空间
                                safe_dict = {"abs": abs, "min": min, "max": max, "round": round}
                                # 对prev和curr分别计算
                                try:
                                    prev_result = eval(prev_expr, {"__builtins__": None}, safe_dict)
                                    working_df.at[idx, col_prev] = prev_result if isinstance(prev_result, (int, float)) else 0.0
                                except:
                                    working_df.at[idx, col_prev] = 0.0
                                try:
                                    curr_result = eval(curr_expr, {"__builtins__": None}, safe_dict)
                                    working_df.at[idx, col_curr] = curr_result if isinstance(curr_result, (int, float)) else 0.0
                                except:
                                    working_df.at[idx, col_curr] = 0.0
                            except Exception as e:
                                # 如果计算失败，留空
                                pass
                    
                        st.write("⚙️ 正在生成可追溯的单元格公式...")
                        
                        # 生成Excel（含公式）
                        temp_buffer = io.BytesIO()
                        working_df.to_excel(temp_buffer,index=False)
                        temp_buffer.seek(0)
                        
                        wb = openpyxl.load_workbook(temp_buffer)
                        ws = wb.active
                        
                        cols = {cell.value:cell.column for cell in ws[1]}
                        
                        prev_idx = cols.get(col_prev)
                        curr_idx = cols.get(col_curr)
                        
                        # 确保原有的公式单元格不被破坏
                        for r in range(2,ws.max_row+1):
                            for c_idx in [prev_idx,curr_idx]:
                                if not c_idx:
                                    continue
                                cell = ws.cell(r,c_idx)
                                if isinstance(cell.value,str) and cell.value.startswith("="):
                                    cell.value = cell.value
                        
                        field_to_row = {
                            str(ws.cell(r,cols[COL_FIELD_NAME]).value).strip():r
                            for r in range(2,ws.max_row+1)
                            if ws.cell(r,cols[COL_FIELD_NAME]).value
                        }
                        
                        sorted_fields = sorted(field_to_row.keys(),key=len,reverse=True)
                        
                        prev_letter = get_column_letter(prev_idx) if prev_idx else ""
                        curr_letter = get_column_letter(curr_idx) if curr_idx else ""
                        
                        for r in range(2,ws.max_row+1):
                            ftype = str(ws.cell(r,cols[COL_FIELD_TYPE]).value).strip()
                            rule = str(ws.cell(r,cols[COL_RULE]).value).strip()
                            if ftype!="计算" or rule in ["nan","None",""]:
                                continue
                            prev_formula = rule
                            curr_formula = rule
                            for f in sorted_fields:
                                if f in rule:
                                    row_num = field_to_row[f]
                                    if prev_letter:
                                        prev_formula = prev_formula.replace(f,f"{prev_letter}{row_num}")
                                    if curr_letter:
                                        curr_formula = curr_formula.replace(f,f"{curr_letter}{row_num}")
                            try:
                                if prev_idx:
                                    ws.cell(r,prev_idx).value = f"={prev_formula}"
                                if curr_idx:
                                    ws.cell(r,curr_idx).value = f"={curr_formula}"
                            except:
                                pass
                        
                        # 格式化数字
                        for r in range(2,ws.max_row+1):
                            for c_idx in [prev_idx,curr_idx]:
                                if not c_idx:
                                    continue
                                cell = ws.cell(r,c_idx)
                                cell.number_format = '#,##0_ ;(#,##0);"-"'
                                if cell.value and not str(cell.value).startswith("="):
                                    try:
                                        clean_val = str(cell.value).replace(',','').replace('(','-').replace(')','')
                                        cell.value = float(clean_val)
                                    except:
                                        pass
                        
                        final_buffer = io.BytesIO()
                        wb.save(final_buffer)
                        final_buffer.seek(0)
                        
                        st.session_state['filled_excel'] = final_buffer.getvalue()
                        st.session_state['final_df'] = working_df
                        st.session_state['col_prev'] = col_prev
                        st.session_state['col_curr'] = col_curr
                    
                    except Exception as e:
                        st.error(f"❌ 流程中断: {str(e)}")
                        status_box.update(label="处理失败", state="error", expanded=True)

        if 'filled_excel' in st.session_state:
            st.markdown("---")
            st.markdown("#### 📋 智能填报结果预览")
            company_name = st.session_state.get('pdf_name', '未命名').replace(".pdf", "")
            
            st.download_button(
                label="⬇️ 下载已填报的底稿 (含公式)",
                data=st.session_state['filled_excel'],
                file_name=f"{company_name}_自动填报表.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
            st.markdown("💡 *预览表仅展示数值，下载后的 Excel 已包含自动关联公式。*")
            st.dataframe(st.session_state['final_df'], use_container_width=True)
    # ----- Step 4 分析与校验（财险版）-----
    with tab4:
        # 🎨 注入自定义 CSS（保持不变）
        st.markdown("""
            <style>
            div[data-testid="stPopover"] > button {
                padding: 2px 8px !important;
                min-height: 26px !important;
                height: 26px !important;
                font-size: 12px !important;
                background-color: #f8f9fa !important;
                border: 1px solid #ddd !important;
            }
            div[data-testid="stColumn"] button[kind="secondary"] {
                padding: 2px 10px !important;
                height: 26px !important;
                font-size: 12px !important;
                white-space: nowrap !important;
                min-width: 70px !important;
                display: flex !important;
                align-items: center !important;
                justify-content: center !important;
            }
            .stSubheader { margin-bottom: -10px !important; }
            </style>
        """, unsafe_allow_html=True)
    
        # 1. 顶部工具栏布局
        col_title, col_btn_sample, col_btn_refresh, col_manual_area = st.columns([3.5, 1.8, 1.1, 1.8])
        
        with col_title:
            st.subheader("🏁 精算数据勾稽关系检查")
            
        with col_btn_sample:
            with st.popover("校验表获取", use_container_width=True):
                st.caption("将人工校验两列复制到需要校对的表格对应位置")
                import requests
                # 🔧 修改点：替换为财险勾稽规则样例表 URL
                url = "https://github.com/Polly031021/Annual-Report-System/raw/main/%E5%8B%BE%E7%A8%BD%E8%A7%84%E5%88%99%E6%A0%B7%E4%BE%8B%E8%A1%A8.xlsx"
                try:
                    with st.spinner("获取中..."):
                        response = requests.get(url, timeout=10)
                        response.raise_for_status()
                    st.download_button(
                        label="⬇️ 下载样例文件", 
                        data=response.content, 
                        file_name="财险人工校验样例.xlsx", 
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                except Exception as e:
                    st.error("网络请求失败，请检查网络或 GitHub 链接是否有效！")
        
        with col_btn_refresh:
            if st.button("🔄 刷新", use_container_width=True):
                st.rerun() 
    
        # --- 📖 核对重点与语音（财险版）---
        with col_manual_area:
            c1, c2 = st.columns([1, 0.4])
            
            # ========== 🔧 修改点：财险版核对指南 ==========
            manual_content = """
            财险数据人工核对指南
            
            1. 模板准备： 请将样例表中的"人工校验"两列手动复制到当前需要核对的表格中。
            2. 年度对齐： 重点检查 最新年 和 上一年 的数据是否有填反现象。
            3. 综合成本率： 检查 COR = 综合赔付率 + 综合费用率 是否成立。
            4. 承保利润： 检查 承保利润 = 已赚保费 - 赔付支出 - 手续费及佣金 - 业务及管理费。
            5. 赔付率拆解： 检查 当期赔款占比、回溯偏差率等因子是否合理。
            6. 费用分类： 检查 手续费及佣金、业务及管理费 的分类是否准确。
            7. 偿付能力： 如数据中有偿付能力充足率，需与最低资本、实际资本勾稽。
            8. 准备金评估： 未到期责任准备金、未决赔款准备金（含 IBNR）需与相关科目勾稽。
            """
    
            with c1:
                with st.popover("🕶️核对指南", use_container_width=True):
                    html_code = f"""
    <div style="font-family: 'Microsoft YaHei', sans-serif; padding: 5px;">
        <h4 style="color: #1f4e79; margin: 0 0 10px 0; font-size: 1rem; border-bottom: 2px solid #e1e4e8; padding-bottom: 5px;">
            财险数据人工核对指南
        </h4>
        <div style="line-height: 1.6; font-size: 0.85rem; color: #333333;">
            <p style="margin-bottom: 8px;">
                <strong style="color: #d32f2f;">1. 模板准备：</strong> 
                将样例表的"人工校验"两列复制到核对表中。
            </p>
            <p style="margin-bottom: 8px;">
                <strong style="color: #d32f2f;">2. 年度对齐：</strong> 
                检查 <span style="background-color: #fff3cd; padding: 1px 3px;">去年</span> 和 <span style="background-color: #fff3cd; padding: 1px 3px;">最新年</span> 数据是否有填反。
            </p>
            <p style="margin-bottom: 8px;">
                <strong style="color: #d32f2f;">3. 综合成本率：</strong> 
                检查 COR = 综合赔付率 + 综合费用率 是否成立。
            </p>
            <p style="margin-bottom: 8px;">
                <strong style="color: #d32f2f;">4. 承保利润：</strong> 
                承保利润 = 已赚保费 - 赔付支出 - 手续费 - 管理费。
            </p>
            <p style="margin-bottom: 8px;">
                <strong style="color: #d32f2f;">5. 赔付率拆解：</strong> 
                检查当期赔款占比、回溯偏差率等因子是否合理。
            </p>
            <p style="margin-bottom: 8px;">
                <strong style="color: #d32f2f;">6. 费用分类：</strong> 
                检查手续费及佣金、业务及管理费的分类是否准确。
            </p>
            <p style="margin-bottom: 8px;">
                <strong style="color: #d32f2f;">7. 偿付能力：</strong> 
                偿付能力充足率需与最低资本、实际资本勾稽。
            </p>
            <p style="margin-bottom: 0;">
                <strong style="color: #d32f2f;">8. 准备金评估：</strong> 
                未到期责任准备金、未决赔款准备金（含 IBNR）需与相关科目勾稽。
            </p>
        </div>
    </div>
    """
                    st.markdown(html_code, unsafe_allow_html=True)
                    
            with c2:
                clean_voice_text = manual_content.replace('\n', '。').replace('**', '')
                tts_html = f"""
                <div style="display: flex; align-items: center; height: 26px;">
                    <button id="tts-btn" onclick="toggleSpeak()" style="
                        border: 1px solid #ddd;
                        background-color: #f8f9fa;
                        border-radius: 4px;
                        cursor: pointer;
                        width: 26px;
                        height: 26px;
                        display: flex;
                        align-items: center;
                        justify-content: center;
                        font-size: 12px;
                    ">📢</button>
                </div>
                <script>
                var msg = new SpeechSynthesisUtterance();
                msg.text = "{clean_voice_text}";
                msg.lang = 'zh-CN';
                msg.rate = 1.2;
                function toggleSpeak() {{
                    const btn = document.getElementById('tts-btn');
                    if (window.speechSynthesis.speaking) {{
                        window.speechSynthesis.cancel();
                        btn.style.backgroundColor = "#f8f9fa";
                        btn.innerText = "📢";
                    }} else {{
                        window.speechSynthesis.speak(msg);
                        btn.style.backgroundColor = "#ffebeb"; 
                        btn.innerText = "⏹";
                        msg.onend = function() {{
                            btn.style.backgroundColor = "#f8f9fa";
                            btn.innerText = "📢";
                        }};
                    }}
                }}
                </script>
                """
                st.components.v1.html(tts_html, height=30)
    
        # ----- 校验执行逻辑（保持不变，框架完全通用）-----
        if 'final_df' in st.session_state:
            df = st.session_state['final_df'].copy()
            
            with st.expander("🛠️ 勾稽规则配置 (点击展开/修改公式)"):
                st.info("curr代表当前年度数据；prev代表上一年度数据。")
                
                # 🔧 修改点：使用空列表或财险规则，不再使用寿险 DEFAULT_RULES
                # ===== 财险默认校验规则（预置 5 条核心规则） =====
                DEFAULT_RULES_PC = [
                    {
                        "规则名称": "1. 综合成本率 = 赔付率 + 费用率",
                        "公式": "abs(curr['综合成本率'] - (curr['综合赔付率'] + curr['综合费用率'])) < 0.001",
                        "类型": "single",
                        "描述": "COR 应等于赔付率与费用率之和（误差<0.1%）"
                    },
                    {
                        "规则名称": "2. 承保利润勾稽",
                        "公式": "abs(curr['承保利润'] - (curr['保险业务收入'] - curr['已发生赔款'] - curr['手续费及佣金支出'] - curr['业务及管理费'])) < 5",
                        "类型": "single",
                        "描述": "承保利润 = 收入 - 赔款 - 手续费 - 管理费"
                    },
                    {
                        "规则名称": "3. 总资产 = 总负债 + 股东权益",
                        "公式": "abs(curr['总资产'] - (curr['总负债'] + curr['股东权益'])) < 5",
                        "类型": "single",
                        "描述": "资产负债表恒等式"
                    },
                    {
                        "规则名称": "4. 净资产变动（跨年）",
                        "公式": "abs(curr['期末股东权益'] - prev['期末股东权益'] - curr['净利润'] - curr['其他综合收益']) < 10",
                        "类型": "cross",
                        "描述": "期末净资产 ≈ 上期净资产 + 净利润 + OCI"
                    },
                    {
                        "规则名称": "5. 新旧准则比值合理",
                        "公式": "abs(curr['新旧准则比值'] - (curr['保险服务收入'] / curr['保险业务收入'])) < 0.001",
                        "类型": "single",
                        "描述": "新旧准则比值 = 保险服务收入 / 保险业务收入"
                    }
                ]
                
                # 使用预置规则作为默认值
                rules_df = st.data_editor(
                    pd.DataFrame(DEFAULT_RULES_PC),
                    num_rows="dynamic",
                    key="rules_editor_v3"
                )
                # 方案2：预置财险规则（取消下面注释）
                # DEFAULT_RULES_PC = [
                #     {"规则名称": "1. 综合成本率 = 赔付率 + 费用率", 
                #      "公式": "abs(curr['综合成本率'] - (curr['赔付率'] + curr['费用率'])) < 0.001", 
                #      "类型": "single"},
                #     ...
                # ]
                # rules_df = st.data_editor(pd.DataFrame(DEFAULT_RULES_PC), num_rows="dynamic", key="rules_editor_v3")
    
                target_prev = str(st.session_state.get('col_prev'))
                target_curr = str(st.session_state.get('col_curr'))
                
                col_prev_name = next(
                    (col for col in df.columns if str(target_prev) in str(col)),
                    None
                )
                
                col_curr_name = next(
                    (col for col in df.columns if str(target_curr) in str(col)),
                    None
                )
                
                if not col_prev_name or not col_curr_name:
                    st.error(
                        f"❌ 表格中找不到包含 '{target_prev}' 和 '{target_curr}' 的列。"
                    )
                    st.stop()
    
            def parse_finance_num(val):
                if pd.isnull(val): return 0.0
                v_str = str(val).strip()
                if v_str in ['-', '']: return 0.0
                if v_str.startswith('(') and v_str.endswith(')'): v_str = '-' + v_str[1:-1]
                v_str = v_str.replace(',', '')
                try: return float(v_str)
                except ValueError: return 0.0
    
            # SmartDict（保持不变）
            class SmartDict(dict):
                def _clean(self, k):
                    if not isinstance(k, str): return k
                    return k.replace('（','(').replace('）',')').replace('：',':').replace(' ','').replace('\n','').replace('\xa0','')
                
                def __init__(self, mapping):
                    super().__init__(mapping)
                    self._map = {self._clean(k): k for k in mapping.keys()}
                    
                def __getitem__(self, k):
                    clean_k = self._clean(k)
                    if clean_k in self._map:
                        return super().__getitem__(self._map[clean_k])
                    raise KeyError(k)
    
            base_prev = {
                str(k).strip(): parse_finance_num(v)
                for k, v in zip(df['字段名'], df[col_prev_name])
            }
            
            base_curr = {
                str(k).strip(): parse_finance_num(v)
                for k, v in zip(df['字段名'], df[col_curr_name])
            }
            
            prev_year_map = SmartDict(base_prev)
            curr_year_map = SmartDict(base_curr)
            
            check_results = []
            
            for _, rule in rules_df.iterrows():
                rule_name = str(rule.get('规则名称', ''))
                formula = str(rule.get('公式', ''))
                rule_type = str(rule.get('类型', 'single')).strip().lower()
            
                targets = (
                    [(f"{target_prev}-{target_curr}", None)]
                    if rule_type == "cross"
                    else [
                        (target_prev, prev_year_map),
                        (target_curr, curr_year_map)
                    ]
                )
            
                for year_label, curr_map in targets:
                    ctx = {
                        'curr': curr_map,
                        'prev': prev_year_map,
                        'prev_year': prev_year_map,
                        'curr_year': curr_year_map,
                        'abs': abs,
                        'min': min,
                        'max': max,
                        'round': round
                    }
            
                    try:
                        is_pass = bool(eval(formula, {"__builtins__": None}, ctx))
                        status = "✅ PASS" if is_pass else "❌ FALSE"
                        detail = "勾稽无误" if is_pass else "差额超过允许阈值"
            
                    except KeyError as e:
                        status = "⚠️ ERROR"
                        detail = f"缺失字段: {str(e)}"
            
                    except Exception as e:
                        status = "⚠️ ERROR"
                        detail = f"公式错误: {str(e)}"
            
                    check_results.append({
                        "年度": year_label,
                        "规则名称": rule_name,
                        "检查结果": status,
                        "诊断详情": detail
                    })
    
            res_display_df = pd.DataFrame(check_results)
            def color_status(val):
                if '✅' in str(val):
                    return 'background-color: #C6EFCE; color: #006100;'
                elif '❌' in str(val):
                    return 'background-color: #FFC7CE; color: #9C0006;'
                elif '⚠️' in str(val):
                    return 'background-color: #FFEB9C; color: #9C5700;'
                return ''
            # 在 st.dataframe 之前插入
            if res_display_df.empty:
                res_display_df = pd.DataFrame(columns=["年度", "规则名称", "检查结果", "诊断详情"])
            else:
                # 可选：确保列都存在
                for col in ["年度", "规则名称", "检查结果", "诊断详情"]:
                    if col not in res_display_df.columns:
                        res_display_df[col] = ""
            
            # 然后正常显示
            st.dataframe(
                res_display_df.style.map(color_status, subset=['检查结果']),
                use_container_width=True, hide_index=True
            )
            
            def color_status(val):
                if '✅' in str(val): return 'background-color: #C6EFCE; color: #006100;'
                elif '❌' in str(val): return 'background-color: #FFC7CE; color: #9C0006;'
                elif '⚠️' in str(val): return 'background-color: #FFEB9C; color: #9C5700;'
                return ''
    
            st.dataframe(
                res_display_df.style.map(color_status, subset=['检查结果']),
                use_container_width=True, hide_index=True
            )
    
            # 导出逻辑（保持不变）
            if st.button("📥 导出带勾稽结果的报告"):
                if 'filled_excel' not in st.session_state:
                    st.error("❌ 找不到填报后的数据，请先重新运行 Step 3")
                    st.stop()
                    
                input_buffer = io.BytesIO(st.session_state['filled_excel'])
                workbook = openpyxl.load_workbook(input_buffer)
                
                ws_data = workbook.active
                ws_data.title = "数据明细"
                
                if "勾稽报告" in workbook.sheetnames:
                    del workbook["勾稽报告"]
                ws_res = workbook.create_sheet("勾稽报告")
                
                headers = list(res_display_df.columns)
                for c_idx, header in enumerate(headers, 1):
                    ws_res.cell(row=1, column=c_idx).value = header
                
                from openpyxl.styles import PatternFill
                fill_green = PatternFill(start_color='C6EFCE', fill_type='solid')
                fill_red = PatternFill(start_color='FFC7CE', fill_type='solid')
                fill_yellow = PatternFill(start_color='FFEB9C', fill_type='solid')
                
                for r_idx, row_data in enumerate(res_display_df.values, 2):
                    for c_idx, value in enumerate(row_data, 1):
                        cell = ws_res.cell(row=r_idx, column=c_idx)
                        cell.value = value
                        if c_idx == 3:
                            val_str = str(value)
                            if "PASS" in val_str: cell.fill = fill_green
                            elif "FALSE" in val_str: cell.fill = fill_red
                            elif "ERROR" in val_str: cell.fill = fill_yellow
                
                ws_res.column_dimensions['B'].width = 30
                ws_res.column_dimensions['C'].width = 15
                ws_res.column_dimensions['D'].width = 40
    
                output = io.BytesIO()
                workbook.save(output)
                
                st.download_button(
                    label="点击下载精算复核底稿", 
                    data=output.getvalue(), 
                    file_name=f"勾稽复核底稿_{st.session_state.get('pdf_name','未命名').replace('.pdf','')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        else:
            st.info("💡 提取结果将在此处显示。请先在上方点击“开始提取”按钮。")    
    # ----- Step 5：多公司数据集成 -----
    with tab5:
        col_title, col_btn = st.columns([4, 1])
        with col_title:
            st.subheader("⛓️‍💥 多公司数据集成与汇率/单位转换")
        st.info("功能说明：支持上传单文件多Sheet或多文件。系统将自动提取所有公司，请在下方为不同公司配置对应的汇率和原始表格的单位。")
    
        uploaded_files = st.file_uploader("请上传已完成勾稽检查的底稿 (支持多文件或单文件多Sheet)", type="xlsx", accept_multiple_files=True)
    
        @st.cache_data(show_spinner=False)
        def run_data_integration(temp_data_list, rate_cfg, unit_cfg):
            # =========================================================
            # 🌟 修改点：财险版豁免字段列表（比率/文本类不换算）
            # =========================================================
            exact_exempt_fields = [
                # 比率类（已为百分比，无需换算）
                "综合成本率",
                "综合赔付率",
                "综合费用率",
                "赔付率",
                "费用率",
                "投资收益率",
                "综合偿付能力充足率",
                "核心偿付能力充足率",
                # 文本类（无需换算）
                "折现率假设",
                "风险边际",
                "非金融风险调整",
                "计量方法",
                "会计政策",
            ]
            
            combined_list = []
            
            def clean_to_float(val):
                try:
                    if isinstance(val, (int, float)): return float(val)
                    if isinstance(val, str):
                        val = val.replace(',','').replace('(','-').replace(')','').strip()
                        return 0.0 if val in ['-',''] else float(val)
                    return 0.0
                except: return 0.0
    
            for item in temp_data_list:
                df_single, comp_name = item["df"], item["comp"]
                rate, unit_mult = rate_cfg[comp_name], unit_cfg[comp_name]
    
                year_cols = {}
                for c in df_single.columns:
                    import re
                    m = re.search(r'(20\d{2})', str(c))
                    if m: year_cols[m.group(1)] = c
    
                if len(year_cols) < 1: continue
    
                base_cols = ["公司类型","公司","类别","字段名","字段类型"]
                existing_base = [c for c in base_cols if c in df_single.columns]
    
                for year_label, col_name in year_cols.items():
                    df_year = df_single[existing_base + [col_name]].copy()
                    df_year["公司"] = comp_name
                    df_year["报告年份"] = year_label
                    df_year["汇率"] = rate
                    df_year["(百万)原币"] = None
                    df_year["(百万)人民币"] = None
                    df_year["汇率"] = df_year["汇率"].astype(object)
    
                    for idx in df_year.index:
                        raw_val = df_year.loc[idx, col_name]
                        f_name = str(df_year.loc[idx, "字段名"]).strip()
                        # 获取字段类型（若存在）
                        field_type = str(df_year.loc[idx, "字段类型"]).strip() if "字段类型" in df_year.columns else ""
                        # 如果是“计算”字段，直接保留原值，不做任何换算
                        if field_type == "计算" and f_name != "再保净成本":
                            df_year.at[idx, "(百万)原币"] = raw_val
                            df_year.at[idx, "(百万)人民币"] = raw_val
                            df_year.at[idx, "汇率"] = "豁免换算"
                            continue
                                                
                        # 判断是否为豁免字段（比率/文本）
                        if f_name in exact_exempt_fields:
                            # 文本字段直接保留原值
                            if any(x in f_name for x in ["假设", "方法", "政策", "边际"]):
                                df_year.at[idx,"(百万)原币"] = str(raw_val) if pd.notna(raw_val) else ""
                                df_year.at[idx,"(百万)人民币"] = str(raw_val) if pd.notna(raw_val) else ""
                                df_year.at[idx,"汇率"] = "豁免换算"
                            else:
                                # 比率字段：去除%符号后转为小数
                                dec_val = float(raw_val.replace('%','').strip())/100.0 if isinstance(raw_val,str) and '%' in raw_val else clean_to_float(raw_val)
                                df_year.at[idx,"(百万)原币"] = dec_val
                                df_year.at[idx,"(百万)人民币"] = dec_val
                                df_year.at[idx,"汇率"] = "豁免换算"
                        else:
                            c_val = clean_to_float(raw_val)
                            df_year.at[idx,"(百万)原币"] = c_val * unit_mult
                            df_year.at[idx,"(百万)人民币"] = c_val * unit_mult * rate
    
                    final_cols = ["公司类型","公司","类别","字段名","字段类型","报告年份","(百万)原币","汇率","(百万)人民币"]
                    actual_cols = [c for c in final_cols if c in df_year.columns]
                    combined_list.append(df_year[actual_cols])
    
            return pd.concat(combined_list, ignore_index=True) if combined_list else pd.DataFrame()
    
        # ✅ 从列名自动识别单位的辅助函数
        def detect_unit_from_col(col_name):
            """从列名里读取单位关键词，返回对应的下拉选项文字"""
            s = str(col_name)
            if "十亿" in s: return "原表为: 十亿元 (× 1,000)"
            if "亿元" in s: return "原表为: 亿元 (× 100)"
            if "百万" in s: return "原表为: 百万元 (无需转换)"
            if "万元" in s or "万" in s: return "原表为: 万元 (÷ 100)"
            if "千元" in s: return "原表为: 千元 (÷ 1,000)"
            if "元" in s:   return "原表为: 元 (÷ 1,000,000)"
            return "原表为: 元 (÷ 1,000,000)"  # ← 默认改为“元”
    
        unit_multipliers = {
            "原表为: 百万元 (无需转换)": 1.0,
            "原表为: 万元 (÷ 100)": 0.01,
            "原表为: 元 (÷ 1,000,000)": 0.000001,
            "原表为: 千元 (÷ 1,000)": 0.001,
            "原表为: 亿元 (× 100)": 100.0,
            "原表为: 十亿元 (× 1,000)": 1000.0,
        }
    
        if uploaded_files:
            all_temp_data, found_companies = [], {}
    
            for file in uploaded_files:
                xl = pd.ExcelFile(file)
                for sheet_name in xl.sheet_names:
                    df_raw = pd.read_excel(file, sheet_name=sheet_name)
                    current_company = str(df_raw["公司"].iloc[0]) if "公司" in df_raw.columns and not df_raw["公司"].empty else sheet_name
                    if "基本信息" in current_company or "Sheet" in current_company: continue
    
                    # ✅ 从列名自动检测默认单位
                    import re
                    detected_unit = "原表为: 元 (÷ 1,000,000)"  # 先设默认值
                    for c in df_raw.columns:
                        if re.search(r'20\d{2}', str(c)):
                            detected_unit = detect_unit_from_col(c)
                            break
    
                    found_companies[current_company] = detected_unit
                    all_temp_data.append({"comp": current_company, "df": df_raw, "source": f"{file.name} - {sheet_name}"})
    
            st.markdown("#### 💵 汇率与数值单位配置盘")
            st.caption("目标表统一要求以【百万元人民币】展示。系统已根据列名自动识别默认单位，请人工核对后调整。")
    
            rate_config, unit_config = {}, {}
            rate_cols = st.columns(3)
            for i, (comp, default_unit) in enumerate(sorted(found_companies.items())):
                with rate_cols[i % 3]:
                    with st.container(border=True):
                        st.markdown(f"**🏢 {comp}**")
                        rate_config[comp] = st.number_input("汇率 (相对于RMB)", value=1.0, step=0.0001, format="%.4f", key=f"rate_{comp}")
                        unit_options = list(unit_multipliers.keys())
                        default_idx = unit_options.index(default_unit) if default_unit in unit_options else 0
                        unit_choice = st.selectbox("原表数值单位", unit_options, index=default_idx, key=f"unit_{comp}")
                        unit_config[comp] = unit_multipliers[unit_choice]
    
            if st.button("开始集成并换算数据", type="primary", use_container_width=True):
                with st.spinner("正在后台执行极速数据合并与换算..."):
                    final_all_df = run_data_integration(all_temp_data, rate_config, unit_config)
    
                if not final_all_df.empty:
                    st.session_state['integrated_data'] = final_all_df
                    years_found = sorted(final_all_df['报告年份'].unique().tolist())
                    st.success(f"✅ 集成与换算完毕！共处理 {len(found_companies)} 家公司，识别到年份：{' / '.join(years_found)}，生成 {len(final_all_df)} 条对标数据。")
                    st.dataframe(final_all_df, use_container_width=True)
    
                    import io
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        final_all_df.to_excel(writer, index=False, sheet_name='行业集成分析表')
                    st.download_button(label="下载行业集成目标表 (长表格式)", data=output.getvalue(), file_name="行业集成目标数据表.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
                else:
                    st.error("未能从上传的文件中提取到有效数据，请检查列名中是否包含年份信息（如 2024YE、2025YE 等）。")

    # ----- Step 6 可视化分析面板 -----
        with tab6:
            st.markdown("### 📊 自定义对标分析")
        
            # ── 数据源选择 ───────────────────────────────────────────────────────
            if 'integrated_data' not in st.session_state:
                st.session_state['integrated_data'] = None
        
            source_choice = st.radio(
                "数据源选择", ["直接引用集成后的数据", "上传集成表 Excel"], horizontal=True
            )
            df_raw = None
        
            if source_choice == "上传集成表 Excel":
                viz_file = st.file_uploader("上传行业集成目标表", type=["xlsx"])
                if viz_file:
                    df_raw = pd.read_excel(viz_file)
                    st.session_state['integrated_data'] = df_raw
            else:
                df_raw = st.session_state.get('integrated_data')
        
            if df_raw is None:
                st.info("💡 请先完成数据集成或上传目标底稿。")
                st.stop()
        
            # ==========================================
            # 🌟 提速秘籍 1：锁定数据预处理 (只算一次)
            # ==========================================
            @st.cache_data(show_spinner=False)
            def prepare_viz_data(df_in):
                df_clean = df_in.copy()
                df_clean.columns = (
                    df_clean.columns.astype(str)
                    .str.strip()
                    .str.replace('\n', '', regex=False)
                    .str.replace('\r', '', regex=False)
                    .str.replace('\ufeff', '', regex=False)
                )
                df_clean['报告年份'] = df_clean['报告年份'].astype(str).str.replace('.0', '', regex=False)
                val_col = "(百万)人民币" if "(百万)人民币" in df_clean.columns else df_clean.columns[-1]
        
                # 去重
                dedup_cols = ['公司', '报告年份', '字段名']
                if '类别' in df_clean.columns:
                    dedup_cols.append('类别')
                df_clean = df_clean.drop_duplicates(subset=dedup_cols, keep='first')
        
                # 透视表
                df_pivot = (
                    df_clean.groupby(['公司', '报告年份', '字段名'])[val_col]
                    .sum().unstack('字段名').reset_index()
                )
        
                # 提取选项列表
                all_fields = sorted([x for x in df_clean['字段名'].unique().tolist() if isinstance(x, str) and x.strip()])
                all_types    = sorted(df_clean['类别'].dropna().astype(str).unique().tolist()) if '类别'    in df_clean.columns else []
                all_co_types = sorted(df_clean['公司类型'].dropna().astype(str).unique().tolist()) if '公司类型' in df_clean.columns else []
        
                all_years_sorted = sorted(
                    [y for y in df_clean['报告年份'].unique() if str(y).isdigit()],
                    key=lambda x: int(x)
                )
                
                return df_clean, df_pivot, val_col, all_fields, all_types, all_co_types, all_years_sorted
        
            # 调用缓存的数据处理函数
            df_clean, df_pivot, val_col, all_fields, all_types, all_co_types, all_years_sorted = prepare_viz_data(df_raw)
        
            # ── 动态年份颜色计算 (极速，无需缓存) ──────────────────────────────────
            KPMG_COLOR_MAP = {
                "Lightest": "#BFE8FF",
                "Light": "#76D2FF",
                "Primary": "#1E49E2",
                "Dark": "#00338D",
            }
            
            n_years = len(all_years_sorted)
            dynamic_year_colors = {}
            target_colors = []
            
            if n_years == 3:
                target_colors = [KPMG_COLOR_MAP["Light"], KPMG_COLOR_MAP["Primary"], KPMG_COLOR_MAP["Dark"]]
            elif n_years == 2:
                target_colors = [KPMG_COLOR_MAP["Primary"], KPMG_COLOR_MAP["Dark"]]
            else:
                target_colors = [KPMG_COLOR_MAP["Lightest"], KPMG_COLOR_MAP["Light"], KPMG_COLOR_MAP["Primary"], KPMG_COLOR_MAP["Dark"]]
            
            num_colors_to_use = len(target_colors)
            if n_years > 0:
                for idx, year in enumerate(all_years_sorted):
                    color_index = 0
                    if n_years > 1:
                        proportion = idx / (n_years - 1)
                        color_index = int(proportion * (num_colors_to_use - 1))
                    color_index = min(color_index, num_colors_to_use - 1)
                    dynamic_year_colors[str(year)] = target_colors[color_index]

            # ── KPMG 色卡 ────────────────────────────────────────────────────────
            # 定义 KPMG 官方色卡（可根据需要增删颜色分类）
            KPMG_CATEGORIES = {
                "KPMG 品牌色": {
                    "深蓝": "#00338D",
                    "浅蓝": "#0865EE",
                    "红色": "#C00000"
                },
                "辅助色": {
                    "绿色": "#92D050",
                    "紫色": "#7030A0",
                    "橙色": "#EF9867",
                    "青色": "#61CBF4",
                    "粉色": "#FEAED7"
                }
            }
            with st.expander("🎨 查看 KPMG 官方色卡"):
                for cat_name, cat_colors in KPMG_CATEGORIES.items():
                    st.markdown(f"**{cat_name}**")
                    html_str = "".join([
                        f'<div style="display:inline-block;margin-right:15px;margin-bottom:8px;">'
                        f'<div style="width:14px;height:14px;background-color:{c};display:inline-block;'
                        f'border-radius:3px;vertical-align:middle;border:1px solid #ddd;"></div>'
                        f'<span style="font-size:13px;vertical-align:middle;"> {n} <b>({c})</b></span></div>'
                        for n, c in cat_colors.items()
                    ])
                    st.markdown(html_str, unsafe_allow_html=True)
        
            st.divider()
            # ==========================================
            # 🌟 提速秘籍 2：锁定 UI 与绘图更新范围 (片段刷新)
            # ==========================================
            @st.fragment
            def render_viz_dashboard(df_clean, df_pivot, val_col, all_fields, all_types, all_co_types, dynamic_year_colors):
                import re, plotly.graph_objects as go, plotly.express as px
            
                def safe_num(s):
                    return pd.to_numeric(s.astype(str).str.strip().str.replace(",", "", regex=False).str.replace("(", "-", regex=False).str.replace(")", "", regex=False).replace(["", "nan", "None", "未披露", "—", "-"], np.nan), errors="coerce")
            
                def safe_formula_eval(wide_df, formula):
                    expr, local_dict = str(formula).strip(), {}
                    for i, field in enumerate(re.findall(r"\[(.*?)\]", expr)):
                        key = f"V{i}"
                        local_dict[key] = safe_num(wide_df[field]) if field in wide_df.columns else pd.Series(np.nan, index=wide_df.index)
                        expr = expr.replace(f"[{field}]", key)
                    try:
                        return pd.Series(pd.eval(expr, local_dict=local_dict, engine="python"), index=wide_df.index).replace([np.inf, -np.inf], np.nan)
                    except Exception:
                        return pd.Series(np.nan, index=wide_df.index)
            
                with st.expander("🛠️ 核心配置面板", expanded=True):
                    r1c1, r1c2, r1c3 = st.columns([1.5, 1, 1])
                    is_pct_stack_mode = False
            
                    with r1c1:
                        chart_type = st.selectbox("📈 1. 图表类型", ["簇状柱状图", "堆积柱状图", "折线对比图", "饼图", "内外环结构对比图", "散点图"])
            
                        if chart_type == "散点图":
                            calc_mode, plot_df_base = "散点图", pd.DataFrame()
                            st.caption("先选择字段；如需运算，开启高级公式，用 [字段名] 引用指标。")
                            x_field = st.selectbox("X轴字段", all_fields, index=0, key="scatter_x_field")
                            y_field = st.selectbox("Y轴字段", all_fields, index=min(1, len(all_fields)-1), key="scatter_y_field")
                            use_formula = st.checkbox("高级公式模式", value=False, key="scatter_formula_on")
                            x_formula = st.text_input("X轴公式", value=f"[{x_field}]", key="scatter_x_formula") if use_formula else f"[{x_field}]"
                            y_formula = st.text_input("Y轴公式", value=f"[{y_field}]", key="scatter_y_formula") if use_formula else f"[{y_field}]"
                            x_axis_title_custom = st.text_input("X轴显示名称", value=x_field, key="scatter_x_title")
                            y_axis_title_custom = st.text_input("Y轴显示名称", value=y_field, key="scatter_y_title")
            
                            all_cos_list = sorted(df_pivot['公司'].unique().tolist())
                            selected_cos = st.multiselect("🏗️ 选择公司", all_cos_list, default=all_cos_list[:min(8, len(all_cos_list))], key="scatter_selected_cos")
                            years_sorted = sorted(df_pivot['报告年份'].unique().tolist(), key=lambda x: int(x) if str(x).isdigit() else 9999)
                            scatter_years = st.multiselect("📅 选择年份", years_sorted, default=years_sorted[-1:], key="scatter_years")
            
                        elif "环" in chart_type or chart_type == "饼图":
                            calc_mode = "结构分析"
                            selected_multi_fields = st.multiselect("🎯 选择构成指标", all_fields, default=all_fields[:min(3, len(all_fields))])
                            selected_cos = st.selectbox("🏗️ 选择展示公司", sorted(df_pivot['公司'].unique().tolist()))
                            plot_df_base = df_clean[(df_clean['字段名'].isin(selected_multi_fields)) & (df_clean['公司'] == selected_cos)].copy().rename(columns={val_col: 'final_val'})
            
                        else:
                            if chart_type == "堆积柱状图":
                                stack_sub = st.radio("堆积模式", ["单指标 / 公式", "多指标占比"], horizontal=True, key="stack_sub_mode")
                                is_pct_stack_mode = (stack_sub == "多指标占比")
            
                            if is_pct_stack_mode:
                                calc_mode = "多指标占比"
                                pct_field_pool = sorted(df_clean[df_clean['类别'] == st.selectbox("🗂️ 按类别筛选指标", ["全部类别"] + all_types, key="pct_type_filter")]['字段名'].unique().tolist()) if all_types and st.session_state.get("pct_type_filter") not in [None, "全部类别"] else all_fields
                                selected_pct_fields = st.multiselect("🎯 选择堆积指标", pct_field_pool, default=pct_field_pool[:min(3, len(pct_field_pool))], key="pct_stack_fields")
                                plot_df_base = df_clean[df_clean['字段名'].isin(selected_pct_fields)][['公司', '报告年份', '字段名', val_col]].copy().rename(columns={val_col: 'final_val'}) if selected_pct_fields else pd.DataFrame(columns=['公司', '报告年份', '字段名', 'final_val'])
                            else:
                                calc_mode = st.radio("数据模式", ["单指标直显", "自定义公式运算"], horizontal=True)
                                if calc_mode == "单指标直显":
                                    filtered_fields = sorted(df_clean[df_clean['类别'] == st.selectbox("🗂️ 按类别筛选指标", ["全部类别"] + all_types, key="type_filter_single")]['字段名'].unique().tolist()) if all_types and st.session_state.get("type_filter_single") not in [None, "全部类别"] else all_fields
                                    target_field = st.selectbox("🎯 选择显示指标", filtered_fields)
                                    plot_df_base = df_pivot[['公司', '报告年份', target_field]].rename(columns={target_field: 'final_val'}).copy()
                                else:
                                    v1, v2 = st.columns(2)
                                    var_a = v1.selectbox("变量 A", all_fields, index=0)
                                    var_b = v2.selectbox("变量 B", ["无"] + all_fields, index=1 if len(all_fields) > 1 else 0)
                                    formula_input = st.text_input("✏️ 自定义公式 (使用 A 和 B)", value="A - B")
                                    try:
                                        A_data, B_data = safe_num(df_pivot[var_a]).fillna(0), (safe_num(df_pivot[var_b]).fillna(0) if var_b != "无" else 0)
                                        df_pivot['final_val'] = pd.Series(pd.eval(formula_input, local_dict={'A': A_data, 'B': B_data}), index=df_pivot.index).replace([np.inf, -np.inf], np.nan).fillna(0)
                                        plot_df_base = df_pivot[['公司', '报告年份', 'final_val']].copy()
                                    except Exception as e:
                                        st.error(f"公式无效: {e}")
                                        plot_df_base = pd.DataFrame(columns=['公司', '报告年份', 'final_val'])
            
                            all_cos_list = sorted(df_pivot['公司'].unique().tolist())
                            if all_co_types:
                                co_type_sel = st.selectbox("🏢 按公司类型快速选择", ["不按公司类型筛选"] + all_co_types, key="co_type_sel")
                                if co_type_sel != "不按公司类型筛选" and st.session_state.get('_prev_co_type_sel') != co_type_sel:
                                    st.session_state['selected_cos_ms'] = sorted(df_clean[df_clean['公司类型'] == co_type_sel]['公司'].unique().tolist())
                                    st.session_state['_prev_co_type_sel'] = co_type_sel
                            selected_cos = st.multiselect("🏗️ 选择对比公司", all_cos_list, default=all_cos_list[:min(2, len(all_cos_list))], key="selected_cos_ms")
            
                    with r1c2:
                        x_axis_mode = "散点视角" if chart_type == "散点图" else (st.radio("🔍 布局视角", ["以公司为横轴", "以年份为横轴"]) if "环" not in chart_type else "结构视角")
                        decimals = st.number_input("🔢 小数位数", 0, 4, 0)
                        show_value = st.toggle("✅ 显示数据标签", value=True)
            
                    with r1c3:
                        unit_options = {"原始数值": 1.0, "亿元": 0.01, "十亿元": 0.001, "百分比(%)": 100.0}
                        selected_unit = st.selectbox("📏 数值单位换算", list(unit_options.keys()))
                        multiplier = unit_options[selected_unit]
                        y_axis_title = st.text_input("📝 Y轴单位显示修改", value=f"单位: {selected_unit.split(' ')[0]}")
                        is_transparent = st.toggle("🌈 开启透明背景模式")
                        show_avg = st.checkbox("平均值线") if "柱" in chart_type else False
                        avg_color = st.color_picker("基准线颜色", value="#ED2124") if show_avg else "#ED2124"
            
                if chart_type == "散点图":
                    scatter_df = df_pivot.copy()
                    scatter_df["X"], scatter_df["Y"] = safe_formula_eval(scatter_df, x_formula) * multiplier, safe_formula_eval(scatter_df, y_formula) * multiplier
                    plot_df = scatter_df[scatter_df["公司"].isin(selected_cos) & scatter_df["报告年份"].isin(scatter_years)].dropna(subset=["X", "Y"]).copy()
                    if plot_df.empty:
                        st.warning("散点图暂无可用数据，请检查字段、公式、公司或年份。")
                        st.stop()
                    color_val = "公司"
                else:
                    if plot_df_base.empty:
                        st.warning("暂无数据，请检查筛选条件。")
                        st.stop()
                    if calc_mode in ("结构分析", "多指标占比"):
                        cos_filter = [selected_cos] if isinstance(selected_cos, str) else selected_cos
                        plot_df, color_val = plot_df_base[plot_df_base['公司'].isin(cos_filter)].copy(), "字段名"
                    else:
                        plot_df = plot_df_base[plot_df_base['公司'].isin(selected_cos)].copy()
                        color_val = "报告年份" if x_axis_mode == "以公司为横轴" else "公司"
                    plot_df['绘制金额'] = safe_num(plot_df['final_val']).fillna(0) * multiplier
            
                st.markdown("#### 🎨 自定义图例标签与颜色")
                legend_col = "公司" if chart_type == "散点图" else color_val
                unique_items = sorted(plot_df[legend_col].dropna().unique().tolist())
                rename_map, color_map, c_cols = {}, {}, st.columns(4)
                for i, item in enumerate(unique_items):
                    with c_cols[i % 4]:
                        st.caption(f"原始值: {item}")
                        new_label = st.text_input("显示名称", value=str(item), key=f"rename_{chart_type}_{item}")
                        default_c = dynamic_year_colors.get(str(item), DEFAULT_COLORS[i % len(DEFAULT_COLORS)] if 'DEFAULT_COLORS' in globals() else "#1E49E2")
                        new_color = st.color_picker("选择颜色", value=default_c, key=f"color_{chart_type}_{item}")
                        rename_map[item], color_map[new_label] = new_label, new_color
                if chart_type != "散点图":
                    plot_df[color_val] = plot_df[color_val].map(rename_map)
                else:
                    plot_df["_legend_show"] = plot_df[color_val].map(rename_map).fillna(plot_df[color_val])
            
                fig, fmt = go.Figure(), f'.{decimals}f'
                suffix, label_fmt = ("%" if selected_unit == "百分比(%)" else ""), f'%{{y:.{decimals}f}}{"%" if selected_unit == "百分比(%)" else ""}'
            
                if chart_type == "散点图":
                    for item in rename_map.values():
                        d = plot_df[plot_df["_legend_show"] == item]
                        if d.empty: continue
                        fig.add_trace(go.Scatter(
                            x=d["X"], y=d["Y"], name=str(item),
                            mode="markers+text" if show_value else "markers",
                            text=d["公司"] if show_value else None,
                            textposition="top center",
                            marker=dict(size=14, color=color_map.get(item, "#1E49E2"), line=dict(width=1.4, color="white")),
                            customdata=np.stack([d["公司"], d["报告年份"]], axis=-1),
                            hovertemplate="<b>%{customdata[0]}</b><br>年份：%{customdata[1]}<br>X=%{x}<br>Y=%{y}<extra></extra>"
                        ))
                    fig.update_xaxes(title=x_axis_title_custom, zeroline=True, zerolinecolor="#999", showgrid=True, gridcolor="#f0f0f0")
                    fig.update_yaxes(title=y_axis_title_custom, zeroline=True, zerolinecolor="#999", showgrid=True, gridcolor="#f0f0f0")
            
                elif "柱状图" in chart_type:
                    barmode = 'group' if "簇状" in chart_type else 'relative'
                    if is_pct_stack_mode:
                        plot_df['x_label'] = plot_df['公司'] + ' ' + plot_df['报告年份']
                        plot_df['_total'] = plot_df.groupby('x_label')['绘制金额'].transform('sum').replace(0, np.nan)
                        plot_df['绘制占比'] = (plot_df['绘制金额'] / plot_df['_total'] * 100).fillna(0)
                        for item in rename_map.values():
                            d = plot_df[plot_df[color_val] == item]
                            fig.add_trace(go.Bar(x=d['x_label'], y=d['绘制占比'], name=str(item), marker_color=color_map[item],
                                text=d.apply(lambda r: f"{r['绘制占比']:.{decimals}f}%<br>{r['绘制金额']:.{decimals}f}", axis=1) if show_value else None, textposition='inside'))
                        fig.update_yaxes(range=[0, 105]); barmode = 'relative'
                    else:
                        for item in rename_map.values():
                            d = plot_df[plot_df[color_val] == item]
                            fig.add_trace(go.Bar(x=d["公司" if color_val == "报告年份" else "报告年份"], y=d["绘制金额"], name=str(item), marker_color=color_map[item],
                                text=d["绘制金额"] if show_value else None, texttemplate=label_fmt if show_value else None, textposition='outside'))
                    fig.update_layout(barmode=barmode)
            
                elif "折线对比图" in chart_type:
                    for item in rename_map.values():
                        d = plot_df[plot_df[color_val] == item]
                        fig.add_trace(go.Scatter(x=d["公司" if color_val == "报告年份" else "报告年份"], y=d["绘制金额"], name=str(item),
                            mode='lines+markers+text' if show_value else 'lines+markers', marker_color=color_map[item],
                            text=d["绘制金额"], texttemplate=label_fmt, textposition="top center"))
            
                elif chart_type == "饼图":
                    d = plot_df[plot_df['报告年份'] == plot_df['报告年份'].max()]
                    fig = px.pie(d, values='绘制金额', names=color_val, hole=0.4, color=color_val, color_discrete_map=color_map)
                    fig.update_traces(textinfo='percent+label' if show_value else 'percent')
            
                elif chart_type == "内外环结构对比图":
                    years_ring = sorted(plot_df['报告年份'].unique().tolist())
                    if len(years_ring) < 2: st.warning("环形图对比需要至少两年的数据。")
                    else:
                        d_outer, d_inner = plot_df[plot_df['报告年份'] == years_ring[-1]], plot_df[plot_df['报告年份'] == years_ring[0]]
                        fig.add_trace(go.Pie(labels=d_outer[color_val], values=d_outer['绘制金额'], hole=0.7, name=years_ring[-1], marker=dict(colors=[color_map[f] for f in d_outer[color_val]]), textinfo='percent+label' if show_value else 'percent'))
                        fig.add_trace(go.Pie(labels=d_inner[color_val], values=d_inner['绘制金额'], hole=0.4, name=years_ring[0], domain={'x': [0.15, 0.85], 'y': [0.15, 0.85]}, marker=dict(colors=[color_map[f] for f in d_inner[color_val]]), textinfo='percent' if show_value else 'none'))
                        fig.update_layout(annotations=[dict(text=f'内:{years_ring[0]}<br>外:{years_ring[-1]}', x=0.5, y=0.5, font_size=12, showarrow=False)])
            
                bg_color = "rgba(0,0,0,0)" if is_transparent else "white"
                if show_avg and "柱" in chart_type and not is_pct_stack_mode:
                    avg_v = plot_df['绘制金额'].mean()
                    fig.add_hline(y=avg_v, line_dash="dash", line_color=avg_color, annotation_text=f"平均: {avg_v:{fmt}}{suffix}", annotation_font=dict(color=avg_color))
            
                fig.update_layout(font_family="Microsoft YaHei", plot_bgcolor=bg_color, paper_bgcolor=bg_color,
                    margin=dict(t=120, l=10, r=10, b=20),
                    legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
                    annotations=[] if chart_type == "散点图" else [dict(x=0, y=1.18, xref='paper', yref='paper', text=f"<b>{y_axis_title}</b>", showarrow=False, font=dict(size=14, color="#333333"), xanchor='left')])
                fig.update_xaxes(showgrid=True if chart_type == "散点图" else False)
                fig.update_yaxes(showgrid=True, gridcolor="#f0f0f0")
            
                show_chart(fig, print_mode, m_id)
                with st.expander("📄 查看底层数据明细"):
                    st.dataframe(plot_df, use_container_width=True)
        
            # ==========================================
            # 在页面上调用封装好的交互展示区
            # ==========================================
            render_viz_dashboard(df_clean, df_pivot, val_col, all_fields, all_types, all_co_types, dynamic_year_colors)
        
    # ----- Step 7 公司级对标报告 -----
    with tab7:
        show_step_7_content()
    
    # ----- Step 8 行业分类统计分析 -----
    with tab8:
        show_step_8_content()

else:
    # 普通用户权限
    st.info("💡 普通用户仅可查看 Step 7 公司级对标报告。")
    show_step_7_content()

# ==================== 页脚 ====================
st.markdown("")
