#左下角搜索并打开 Anaconda Prompt（那个黑框框）。
#输入 D: 并按回车。
#输入 cd D:\实习\实习公司\毕马威\网页实现 并按回车。
#输入 streamlit run app.py 并按回车。
import streamlit as st
import requests
from bs4 import BeautifulSoup
import plotly.express as px
import plotly.graph_objects as go
import numpy as np
from pptx import Presentation
from pptx.util import Inches, Pt
from plotly.subplots import make_subplots
import streamlit.components.v1 as components
import uuid
import math

def show_step_8_content():
    """行业统计分析 - 财险版（完全适配财险指标）"""
    import pandas as pd
    import numpy as np
    import plotly.graph_objects as go
    import plotly.express as px
    from plotly.subplots import make_subplots
    import math
    import re
    import io
    import json
    import requests

    # ---------- 工具函数：解析 bins ----------
    def parse_bins_input(bins_str):
        if bins_str is None:
            return None
        bins_str = str(bins_str).strip()
        if not bins_str:
            return None
        try:
            bins = [float(x.strip()) for x in bins_str.split(",") if str(x).strip() != ""]
            bins = sorted(list(dict.fromkeys(bins)))
            return bins if len(bins) >= 2 else None
        except Exception:
            return None

    # ---------- 工具函数：从 URL 或文件加载自定义刻度 ----------
    def load_custom_bins_excel(uploaded_file):
        if isinstance(uploaded_file, str):
            resp = requests.get(uploaded_file, timeout=15)
            resp.raise_for_status()
            uploaded_file = io.BytesIO(resp.content)
        df = pd.read_excel(uploaded_file)
        custom_map = {}
        for _, row in df.iterrows():
            m_id = str(row["m_id"]).strip()
            if not m_id or m_id.lower() == "nan":
                continue
            if "enable" in df.columns and str(row["enable"]).strip() in ["0", "False", "false", "否"]:
                continue
            x_bins = parse_bins_input(row["x_bins"])
            y_bins = parse_bins_input(row["y_bins"])
            if x_bins is None and y_bins is None:
                continue
            custom_map[m_id] = {"x_bins_custom": x_bins, "y_bins_custom": y_bins}
        return custom_map

    # ---------- 工具函数：自动生成漂亮刻度 ----------
    def generate_nice_bins(values, target_bins=6):
        v = pd.Series(values).dropna()
        if len(v) < 2:
            return [0, 1]
        mn, mx = v.min(), v.max()
        if mn == mx:
            return [0, mx * 1.2] if mx else [0, 1]
        raw_step = (mx - mn) / target_bins
        mag = 10 ** math.floor(math.log10(abs(raw_step)))
        step = (1 if raw_step / mag <= 1 else 2 if raw_step / mag <= 2 else 5 if raw_step / mag <= 5 else 10) * mag
        lo = math.floor(mn / step) * step
        hi = math.ceil(mx / step) * step
        bins = np.arange(lo, hi + step, step).tolist()
        return bins if len(bins) >= 2 else [lo, hi] if lo != hi else [lo, lo + 1]

    # ============================================================
    # 1️⃣ 样式注入（与寿险版完全一致，仅调整导航按钮类名）
    # ============================================================
    st.markdown("""
    <style>
    [data-testid="stSidebar"] { background: rgba(255,255,255,0.95) !important; border-right: 1px solid #EAEAEA !important; box-shadow: 2px 0px 15px rgba(0,0,0,0.08) !important; }
    .nav-floating-sign-s8 { position: fixed; left: 0; top: 50%; transform: translateY(-50%); background: rgba(0, 133, 120, 0.85); color: white; padding: 20px 8px; border-radius: 0 12px 12px 0; writing-mode: vertical-rl; text-orientation: mixed; font-size: 22px; font-weight: bold; letter-spacing: 3px; z-index: 9999; cursor: pointer; box-shadow: 3px 3px 12px rgba(0,0,0,0.25); transition: all 0.2s; }
    .nav-floating-sign-s8:hover { background: rgba(0, 133, 120, 1); padding-left: 15px; }
    .stPlotlyChart { width: 100% !important; min-width: 0 !important; }
    .print-only { display: none !important; }
    .cover-page { position: relative !important; width: 338.67mm !important; height: 190.5mm !important; margin: 0 !important; padding: 0 !important; page-break-after: always !important; overflow: hidden !important; background: transparent !important; }
    .cover-page img { width: 100% !important; height: 100% !important; object-fit: cover !important; display: block !important; }
    .block-container { padding-top:0 !important; padding-right:10px !important; padding-left:10px !important; margin-top:0 !important; }
    .cover-text { forced-color-adjust: none !important; -webkit-print-color-adjust: exact !important; print-color-adjust: exact !important; color: white !important; -webkit-text-fill-color: white !important; }
    .element-container:first-child{ margin-top:0 !important; padding-top:0 !important; }
    @media print {
        .print-only { display: block !important; }
        html,body{ width:338.67mm!important; height:190.5mm!important; overflow:hidden!important; zoom:100%!important; }
        .main .block-container{ max-width:100%!important; padding-top:0!important; padding-bottom:0!important; }
        .block-container { padding-top: 0rem !important; }
        .no-print, h1, .nav-floating-sign-s8,
        [data-testid="collapsedControl"], header, footer,
        [data-testid="stHeader"], [data-testid="stSidebar"],
        section[data-testid="stSidebar"],
        [data-testid="stToolbar"], button[kind="secondary"],
        input, .stSlider, [data-testid="stSelectbox"],
        [data-testid="stRadio"], [data-testid="stExpander"],
        .stAlert,
        button[role="tab"],
        div[role="tablist"],
        [data-baseweb="tab-list"],
        hr { display: none !important; }
        .page-break-container { break-inside: avoid !important; margin: 0 !important; padding: 0 !important; padding-bottom: 0mm !important; }
        .stApp { max-width: 100% !important; width: 100% !important; }
        .keep-columns [data-testid="stHorizontalBlock"]{ display:flex!important; flex-wrap:nowrap!important; align-items:flex-start!important; justify-content:space-between!important; gap:0!important; width:100%!important; }
        .keep-columns [data-testid="stHorizontalBlock"]>div{ width:49%!important; min-width:49%!important; max-width:49%!important; flex:0 0 49%!important; overflow:hidden!important; page-break-inside:avoid!important; break-inside:avoid!important; }
        .page-break-title { break-before: page !important; padding-top: 10px !important; margin-top: 0 !important; text-align: left !important; }
        h2 { display: block !important; text-align: left !important; color: #00338D !important; font-size: 30px !important; font-weight: bold !important; border-bottom: 2px solid #00338D !important; padding-bottom: 6px !important; margin: 14px 0 10px 0 !important; }
        h3:not(.no-print) { display: block !important; text-align: left !important; color: #00338D !important; font-size: 30px !important; font-weight: bold !important; margin: 10px 0 8px 0 !important; page-break-after: avoid !important; }
        .plotly-graph-div, .stPlotlyChart { width: 100% !important; max-width: 100% !important; height: auto !important; page-break-inside: avoid !important; display: block !important; }
        div[data-testid="stDataFrame"], div[data-testid="stTable"] { zoom: 0.65 !important; margin: 0 auto 20px auto !important; max-width: 100% !important; page-break-inside: auto !important; }
        div[data-testid="stTable"] tr { page-break-inside: avoid !important; }
        .element-container { page-break-inside: avoid !important; width: 100% !important; }
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
    <div class="nav-floating-sign-s8" id="custom-nav-trigger-s8">展开行业导航栏 </div>
    """, unsafe_allow_html=True)

    components.html("""<script>let t = setInterval(() => { const d = window.parent.document; const b = d.getElementById("custom-nav-trigger-s8"); const c = d.querySelector('[data-testid="collapsedControl"]') || d.querySelector('button[kind="header"]'); if(b && c) { b.onclick = () => c.click(); clearInterval(t); } }, 500);</script>""", height=0, width=0)

    # ---------- 数据加载 ----------
    if 'integrated_data' not in st.session_state or st.session_state['integrated_data'] is None:
        st.warning("⚠️ 请先在 Step 6 完成数据集成。")
        return
    df_raw = st.session_state['integrated_data'].copy()
    # 清理年份
    df_raw['报告年份'] = df_raw['报告年份'].astype(str).str.replace('.0', '', regex=False)
    valid_years = sorted([int(y) for y in df_raw['报告年份'].unique() if y.isdigit()])
    if len(valid_years) < 2:
        st.warning("⚠️ 年份数量不足，至少需要两个年份。")
        return
    latest_year, prev_year = valid_years[-1], valid_years[-2]
    year_list = [prev_year, latest_year]

    st.markdown("### 📈 行业统计分析")

    # ---------- 公司类型顺序和颜色 ----------
    DEFAULT_TYPE_ORDER = ["头部", "银行系", "外资", "养老健康", "小型"]  # 仅作默认
    KPMG_CATEGORIES = globals().get('KPMG_CATEGORIES', {
        "Primary Colors": {"KPMG Blue": "#00338D", "Cobalt Blue": "#1E49E2",
                           "Light Blue": "#ACEAFF", "Pacific Blue": "#00B8F5",
                           "Purple": "#7213EA", "Pink": "#FD349C"},
        "Traffic Light": {"Red": "#ED2124", "Amber": "#F1C44D", "Positive Green": "#269924"}
    })
    primary = KPMG_CATEGORIES["Primary Colors"]
    traffic = KPMG_CATEGORIES["Traffic Light"]
    COMPANY_TYPE_COLORS = {
        "头部": primary["KPMG Blue"],
        "银行系": primary["Pacific Blue"],
        "外资": traffic["Positive Green"],
        "养老健康": traffic["Amber"],
        "小型": primary["Purple"]
    }
    _FALLBACK_COLORS = [primary["Cobalt Blue"], primary["Pink"], traffic["Red"], primary["Light Blue"]]

    def get_company_color(ct):
        return COMPANY_TYPE_COLORS.get(ct, _FALLBACK_COLORS[hash(ct) % len(_FALLBACK_COLORS)])

    # ---------- 公司排序字段 ----------
    COMPANY_SORT_FIELD_OPTIONS = ["总资产", "净资产"]

    def get_company_sort_map(df_raw, latest_year, sort_field):
        df_year = df_raw[df_raw['报告年份'].astype(str) == str(latest_year)]
        sort_map = {}
        for co, g in df_year.groupby('公司'):
            rows = g[g['字段名'] == sort_field]
            sort_map[co] = rows['(百万)人民币'].sum() if not rows.empty else np.nan
        return sort_map

    # ============================================================
    # 2️⃣ 财险版数据计算函数（替换所有寿险 calc_*）
    # ============================================================

    # ---------- 散点图数据计算（用于 SCATTER_AXIS_META） ----------
    def calc_scatter_data(df, selected_types, latest_year, prev_year, field_name, display_name):
        """
        计算散点图数据（各公司某指标当前值、前一年值、增长率）
        返回 DataFrame 包含：公司, 公司类型, display_name(当前值), f"{display_name}增长率"
        """
        df_curr = df[df['报告年份'] == str(latest_year)].copy()
        df_prev = df[df['报告年份'] == str(prev_year)].copy()
        company_type_map = dict(df_curr.groupby('公司')['公司类型'].first().dropna())
        result = []
        for co, ct in company_type_map.items():
            if ct not in selected_types:
                continue
            # 提取当前值
            row_curr = df_curr[(df_curr['公司'] == co) & (df_curr['字段名'] == field_name)]
            if row_curr.empty:
                continue
            curr_val = row_curr['(百万)人民币'].iloc[0] / 100  # 转为亿元
            # 提取前一年值
            row_prev = df_prev[(df_prev['公司'] == co) & (df_prev['字段名'] == field_name)]
            if row_prev.empty:
                continue
            prev_val = row_prev['(百万)人民币'].iloc[0] / 100
            if prev_val == 0:
                growth_rate = np.nan
            else:
                growth_rate = (curr_val - prev_val) / abs(prev_val) * 100
            result.append({
                '公司': co,
                '公司类型': ct,
                display_name: curr_val,
                f'{display_name}增长率': growth_rate
            })
        df_result = pd.DataFrame(result)
        if not df_result.empty:
            df_result.sort_values(display_name, ascending=False, inplace=True)
            df_result.reset_index(drop=True, inplace=True)
            df_result['id'] = range(1, len(df_result) + 1)
        return df_result

    # ---------- 堆叠分布图计算函数 ----------
    def calc_cor_ratio(df_co):
        """综合成本率（COR）"""
        cor = df_co[df_co['字段名'] == '综合成本率']['(百万)人民币'].sum()
        return cor if cor != 0 else np.nan

    def calc_loss_ratio(df_co):
        """综合赔付率"""
        val = df_co[df_co['字段名'] == '综合赔付率']['(百万)人民币'].sum()
        return val if val != 0 else np.nan

    def calc_expense_ratio(df_co):
        """综合费用率"""
        val = df_co[df_co['字段名'] == '综合费用率']['(百万)人民币'].sum()
        return val if val != 0 else np.nan

    # 堆叠分布图配置（与 STACK_DIST_META 对应）
    STACK_DIST_META = {
        "industry_cor_dist": {
            "name": "综合成本率分布",
            "calc_func": calc_cor_ratio,
            "default_x": "80,85,90,95,100,105,110,115,120",
            "y_label": "公司数量"
        },
        "industry_loss_ratio_dist": {
            "name": "综合赔付率分布",
            "calc_func": calc_loss_ratio,
            "default_x": "50,55,60,65,70,75,80,85,90,95",
            "y_label": "公司数量"
        },
        "industry_expense_ratio_dist": {
            "name": "综合费用率分布",
            "calc_func": calc_expense_ratio,
            "default_x": "10,15,20,25,30,35,40,45,50,55",
            "y_label": "公司数量"
        },
    }

    # ---------- 构成图计算函数 ----------
    def calc_industry_expense_composition(df, company_type, year):
        """费用构成（获取费用、维持费用、非履约费用）"""
        fields = ["获取费用", "维持费用", "非履约费用"]
        mask = (df['公司类型'] == company_type) & (df['报告年份'].astype(str) == str(year))
        df_f = df[mask].copy()
        companies = df_f['公司'].unique()
        company_ratios, company_totals = [], []
        for co in companies:
            df_co = df_f[df_f['公司'] == co]
            vals = {f: df_co[df_co['字段名'] == f]['(百万)人民币'].sum() for f in fields}
            total = sum(vals.values())
            if total == 0:
                continue
            ratios = {f: vals[f] / total * 100 for f in fields}
            company_ratios.append(ratios)
            company_totals.append(total)
        if not company_ratios:
            return None
        ratios_df = pd.DataFrame(company_ratios)
        return {'ratios': {f: ratios_df[f].mean() for f in fields}, 'avg_total': np.mean(company_totals)}

    def calc_industry_asset_composition(df, company_type, year):
        """资产构成（AC/FVOCI/FVTPL/指定FVOCI）"""
        fields = ["债权投资", "其他债权投资", "交易性金融资产", "其他权益工具投资"]
        # 映射到显示名
        field_display = {
            "债权投资": "AC（债权投资）",
            "其他债权投资": "FVOCI（其他债权投资）",
            "交易性金融资产": "FVTPL（交易性金融资产）",
            "其他权益工具投资": "指定FVOCI（其他权益工具）"
        }
        mask = (df['公司类型'] == company_type) & (df['报告年份'].astype(str) == str(year))
        df_f = df[mask].copy()
        companies = df_f['公司'].unique()
        company_ratios = []
        for co in companies:
            df_co = df_f[df_f['公司'] == co]
            total = sum(df_co[df_co['字段名'] == f]['(百万)人民币'].sum() for f in fields)
            if total == 0:
                continue
            ratios = {field_display[f]: df_co[df_co['字段名'] == f]['(百万)人民币'].sum() / total * 100 for f in fields}
            company_ratios.append(ratios)
        if not company_ratios:
            return None
        return pd.DataFrame(company_ratios).mean().to_dict()

    def calc_industry_profit_composition(df, company_type, year):
        """财险承保利润构成（保费收入、赔付、费用等对承保利润的贡献）"""
        # 承保利润 = 已赚保费 - 赔付支出 - 手续费 - 管理费
        mask = (df['公司类型'] == company_type) & (df['报告年份'].astype(str) == str(year))
        df_f = df[mask].copy()
        companies = df_f['公司'].unique()
        results = []
        for co in companies:
            df_co = df_f[df_f['公司'] == co]
            prem = df_co[df_co['字段名'] == '已赚保费']['(百万)人民币'].sum()
            claim = df_co[df_co['字段名'] == '赔付支出']['(百万)人民币'].sum()
            comm = df_co[df_co['字段名'] == '手续费及佣金']['(百万)人民币'].sum()
            adm = df_co[df_co['字段名'] == '业务及管理费']['(百万)人民币'].sum()
            if prem == 0:
                continue
            underwriting_profit = prem - claim - comm - adm
            # 计算各项目占比（用绝对值之和作为分母，类似寿险的构成逻辑）
            total_abs = abs(prem) + abs(claim) + abs(comm) + abs(adm) + abs(underwriting_profit) if underwriting_profit != 0 else 1
            results.append({
                '已赚保费': prem / total_abs * 100,
                '赔付支出': -claim / total_abs * 100,
                '手续费及佣金': -comm / total_abs * 100,
                '业务及管理费': -adm / total_abs * 100,
                '承保利润': underwriting_profit / total_abs * 100 if underwriting_profit != 0 else 0
            })
        if not results:
            return None
        df_res = pd.DataFrame(results)
        return df_res.mean().to_dict()

    # ============================================================
    # 3️⃣ 配置表（SCATTER_AXIS_META、COMPOSITION_FIELD_MAP）
    # ============================================================

    SCATTER_AXIS_META = {
        "industry_premium": {
            "title": "已赚保费",
            "x_field": "已赚保费",
            "y_field": "已赚保费增长率",
            "x_label": "已赚保费区间（亿元）",
            "y_label": "已赚保费增长率区间（%）",
            "default_x": "0,50,100,200,500,1000,2000",
            "default_y": "-100,-50,0,50,100,200,300"
        },
        "industry_cor": {
            "title": "综合成本率",
            "x_field": "综合成本率",
            "y_field": "综合成本率增长率",
            "x_label": "综合成本率区间（%）",
            "y_label": "综合成本率增长率区间（%）",
            "default_x": "80,85,90,95,100,105,110,115,120",
            "default_y": "-20,-10,-5,0,5,10,20,30"
        },
        "industry_loss_ratio": {
            "title": "综合赔付率",
            "x_field": "综合赔付率",
            "y_field": "综合赔付率增长率",
            "x_label": "综合赔付率区间（%）",
            "y_label": "综合赔付率增长率区间（%）",
            "default_x": "50,55,60,65,70,75,80,85,90,95",
            "default_y": "-20,-10,-5,0,5,10,20"
        },
        "industry_expense_ratio": {
            "title": "综合费用率",
            "x_field": "综合费用率",
            "y_field": "综合费用率增长率",
            "x_label": "综合费用率区间（%）",
            "y_label": "综合费用率增长率区间（%）",
            "default_x": "10,15,20,25,30,35,40,45,50",
            "default_y": "-20,-10,-5,0,5,10,20"
        },
        "industry_net_profit": {
            "title": "净利润",
            "x_field": "净利润",
            "y_field": "净利润增长率",
            "x_label": "净利润区间（亿元）",
            "y_label": "净利润增长率区间（%）",
            "default_x": "0,10,20,50,100,200,500",
            "default_y": "-100,-50,-20,0,20,50,100,200"
        },
        "industry_asset": {
            "title": "总资产",
            "x_field": "总资产",
            "y_field": "总资产增长率",
            "x_label": "总资产区间（亿元）",
            "y_label": "总资产增长率区间（%）",
            "default_x": "0,100,500,1000,2000,5000,10000",
            "default_y": "-20,-10,0,10,20,30,50"
        },
        "industry_equity": {
            "title": "净资产",
            "x_field": "净资产",
            "y_field": "净资产增长率",
            "x_label": "净资产区间（亿元）",
            "y_label": "净资产增长率区间（%）",
            "default_x": "0,50,100,200,500,1000,2000",
            "default_y": "-30,-20,-10,0,10,20,30,50"
        }
    }

    COMPOSITION_FIELD_MAP = {
        "industry_expense_struct": {
            "type": "expense",
            "fields": ["获取费用", "维持费用", "非履约费用"],
            "display": {"获取费用": "获取费用", "维持费用": "维持费用", "非履约费用": "非履约费用"}
        },
        "industry_asset_struct": {
            "type": "asset",
            "fields": ["债权投资", "其他债权投资", "交易性金融资产", "其他权益工具投资"]
        },
        "industry_profit_comp": {
            "type": "profit",
            "fields": ["已赚保费", "赔付支出", "手续费及佣金", "业务及管理费", "承保利润"]
        },
        "industry_cor_comp": {
            "type": "cor_breakdown",
            "fields": ["综合赔付率", "综合费用率"]
        }
    }

    # ============================================================
    # 4️⃣ 图表创建函数（与寿险版结构类似，但数据来源为财险计算）
    # ============================================================

    # ----- 散点图（热力图）-----
    def create_scatter_chart(df_scatter, config, x_label, y_label):
        if df_scatter.empty:
            st.warning("没有有效数据")
            return

        all_types_in_data = df_scatter['公司类型'].dropna().unique().tolist()
        _type_order = st.session_state.get('step8_type_order', DEFAULT_TYPE_ORDER)
        category_order = [t for t in _type_order if t in all_types_in_data]
        category_order += [t for t in all_types_in_data if t not in _type_order]
        category_colors = {ct: get_company_color(ct) for ct in category_order}

        x_col, y_col = config['x_col'], config['y_col']

        # 排序（按配置的排序字段）
        company_sort_map = get_company_sort_map(df_raw, latest_year, st.session_state.get('step8_company_sort_field', '总资产'))
        df_scatter = df_scatter.copy()
        df_scatter['_sort_val'] = df_scatter['公司'].map(company_sort_map)
        sort_col = '_sort_val' if df_scatter['_sort_val'].notna().any() else x_col
        existing_types = [ct for ct in category_order if ct in df_scatter['公司类型'].values]
        df_sorted = pd.concat([df_scatter[df_scatter['公司类型'] == ct].sort_values(sort_col, ascending=False) for ct in existing_types])
        other_types = [ct for ct in df_scatter['公司类型'].unique() if ct not in existing_types]
        for ct in other_types:
            df_sorted = pd.concat([df_sorted, df_scatter[df_scatter['公司类型'] == ct].sort_values(sort_col, ascending=False)])
        df_scatter = df_sorted.drop(columns='_sort_val').reset_index(drop=True)
        df_scatter['id'] = range(1, len(df_scatter) + 1)

        x_values, y_values = df_scatter[x_col].values, df_scatter[y_col].values
        min_x, max_x, min_y, max_y = min(x_values), max(x_values), min(y_values), max(y_values)

        # 分箱
        if config.get('x_bins_custom') and len(config['x_bins_custom']) > 1:
            bins_x = list(config['x_bins_custom'])
            x_labels = [f"({int(bins_x[i])}, {int(bins_x[i+1])}]" for i in range(len(bins_x) - 1)]
        else:
            step = max(10, int((max_x - min_x) / 4 / 10) * 10) if max_x > 10 else max(1, int((max_x - min_x) / 4))
            bins_x = list(np.arange(min_x - 5, max_x + step, step))
            if len(bins_x) < 2: bins_x = [min_x - 5, max_x + 5]
            bins_x = sorted(set(bins_x))
            x_labels = [f"({int(bins_x[i])}, {int(bins_x[i+1])}]" for i in range(len(bins_x) - 1)]

        if config.get('y_bins_custom') and len(config['y_bins_custom']) > 1:
            bins_y = list(config['y_bins_custom'])
            y_labels = [f"({int(bins_y[i])}, {int(bins_y[i+1])}]" for i in range(len(bins_y) - 1)]
        else:
            y_range = max_y - min_y
            step = max(10, int(y_range / 5)) if y_range != 0 else 10
            bins_y = list(np.arange(min_y - 10, max_y + step, step))
            if len(bins_y) < 2: bins_y = [min_y - 10, max_y + 10]
            bins_y = sorted(set(bins_y))
            y_labels = [f"({int(bins_y[i])}%, {int(bins_y[i+1])}%]" for i in range(len(bins_y) - 1)]

        nx, ny = len(x_labels), len(y_labels)

        def get_bin_idx(val, bins):
            if val <= bins[0]: return 0
            for i, (low, high) in enumerate(zip(bins[:-1], bins[1:])):
                if low < val <= high: return i
            return len(bins) - 2

        df_scatter['x_bin'] = df_scatter[x_col].apply(lambda v: get_bin_idx(v, bins_x))
        df_scatter['y_bin'] = df_scatter[y_col].apply(lambda v: get_bin_idx(v, bins_y))

        density = np.zeros((ny, nx))
        for _, row in df_scatter.iterrows():
            if 0 <= row['x_bin'] < nx and 0 <= row['y_bin'] < ny:
                density[int(row['y_bin']), int(row['x_bin'])] += 1

        def get_position_in_bin(val, bins, bin_idx):
            low, high = bins[bin_idx], bins[bin_idx + 1]
            if high > low:
                frac = max(0.1, min(0.9, (val - low) / (high - low)))
            else:
                frac = 0.5
            return bin_idx + frac

        df_scatter['x_pos'] = df_scatter.apply(lambda r: get_position_in_bin(r[x_col], bins_x, int(r['x_bin'])), axis=1)
        df_scatter['y_pos'] = df_scatter.apply(lambda r: get_position_in_bin(r[y_col], bins_y, int(r['y_bin'])), axis=1)

        # 抖动
        np.random.seed(42)
        for y in range(ny):
            for x in range(nx):
                mask = (df_scatter['x_bin'] == x) & (df_scatter['y_bin'] == y)
                indices = df_scatter[mask].index.tolist()
                n = len(indices)
                if n > 1:
                    for i, idx in enumerate(indices):
                        angle = (i / n) * 2 * np.pi
                        radius = 0.13 * (1 + 0.15 * i)
                        df_scatter.loc[idx, 'x_pos'] += radius * np.cos(angle)
                        df_scatter.loc[idx, 'y_pos'] += radius * np.sin(angle)
                        df_scatter.loc[idx, 'x_pos'] = np.clip(df_scatter.loc[idx, 'x_pos'], x + 0.08, x + 0.92)
                        df_scatter.loc[idx, 'y_pos'] = np.clip(df_scatter.loc[idx, 'y_pos'], y + 0.08, y + 0.92)

        # 表格数据
        df_table = df_scatter.copy()
        rows_per_group = int(np.ceil(len(df_table) / 2))
        left_df = df_table.iloc[:rows_per_group].copy()
        right_df = df_table.iloc[rows_per_group:].copy()
        for df in [left_df, right_df]:
            while len(df) < rows_per_group:
                df.loc[len(df)] = {"id": "", "公司": "", "公司类型": ""}
        table_values = [left_df["id"].astype(str).tolist(), left_df["公司"].tolist(),
                        right_df["id"].astype(str).tolist(), right_df["公司"].tolist()]

        n_rows = rows_per_group
        TABLE_ROW_HEIGHT, TABLE_HEADER_H = 18, 22
        table_body_h = n_rows * TABLE_ROW_HEIGHT + TABLE_HEADER_H + 60
        total_height = table_body_h + 40

        fig = make_subplots(rows=1, cols=2, column_widths=[0.65, 0.35],
                            shared_yaxes=False, horizontal_spacing=0.015,
                            specs=[[{"type": "xy"}, {"type": "table"}]])

        max_count = max(1, int(density.max()))

        def density_to_color(count, max_count):
            if count == 0: return None
            stops = [(0.00, (255, 255, 255)), (0.15, (253, 235, 236)), (0.40, (249, 199, 200)),
                     (0.70, (243, 154, 156)), (1.00, (237, 33, 36))]
            t = count / max_count
            for k in range(len(stops) - 1):
                t0, c0 = stops[k]
                t1, c1 = stops[k + 1]
                if t0 <= t <= t1:
                    ratio = (t - t0) / (t1 - t0)
                    r = int(c0[0] + ratio * (c1[0] - c0[0]))
                    g = int(c0[1] + ratio * (c1[1] - c0[1]))
                    b = int(c0[2] + ratio * (c1[2] - c0[2]))
                    return f"rgb({r},{g},{b})"
            return "rgb(237,33,36)"

        for i in range(ny):
            for j in range(nx):
                count = int(density[i, j])
                color = density_to_color(count, max_count)
                if color:
                    fig.add_shape(type="rect", x0=j, x1=j+1, y0=i, y1=i+1,
                                  fillcolor=color, line=dict(width=0), layer="below", row=1, col=1)
                if count > 0:
                    fig.add_annotation(x=j+0.85, y=i+0.85, text=str(count), showarrow=False,
                                       font=dict(size=9, color="#6B7280"), row=1, col=1)

        point_count = len(df_scatter)
        marker_size = max(9, min(14, int(150 / max(1, point_count / 22))))
        for ct in category_order:
            df_ct = df_scatter[df_scatter['公司类型'] == ct]
            if df_ct.empty: continue
            color = category_colors.get(ct, "#1E49E2")
            fig.add_trace(go.Scatter(x=df_ct['x_pos'], y=df_ct['y_pos'], mode='markers', name=ct,
                                     marker=dict(size=marker_size, color=color, line=dict(width=0.8, color='white')),
                                     hoverinfo='skip', showlegend=True), row=1, col=1)
            fig.add_trace(go.Scatter(x=df_ct['x_pos'], y=df_ct['y_pos'], mode='text', name=ct + "_text",
                                     text=df_ct['id'].astype(str), textposition='middle center',
                                     textfont=dict(size=max(6, int(marker_size * 0.52)), color='white'),
                                     showlegend=False, hoverinfo='skip'), row=1, col=1)

        for i in range(nx + 1):
            fig.add_shape(type="line", x0=i, x1=i, y0=0, y1=ny, line=dict(color="#D9E2EC", width=0.4), layer='above', row=1, col=1)
        for i in range(ny + 1):
            fig.add_shape(type="line", x0=0, x1=nx, y0=i, y1=i, line=dict(color="#D9E2EC", width=0.4), layer='above', row=1, col=1)

        left_colors = [category_colors.get(t, "#1F2937") for t in left_df["公司类型"]]
        right_colors = [category_colors.get(t, "#1F2937") for t in right_df["公司类型"]]
        fig.add_trace(go.Table(
            columnwidth=[10, 40, 10, 40],
            header=dict(values=["编号", "公司", "编号", "公司"], fill_color="#00338D",
                       font=dict(color="white", size=8), align="center", height=TABLE_HEADER_H),
            cells=dict(values=table_values, fill_color="white",
                      font=dict(color=[["#0F172A"] * rows_per_group, left_colors,
                                       ["#0F172A"] * rows_per_group, right_colors], size=7),
                      align="center", height=TABLE_ROW_HEIGHT)
        ), row=1, col=2)

        fig.update_layout(showlegend=True, height=total_height, width=1000,
                         plot_bgcolor='white', paper_bgcolor='white',
                         margin=dict(l=50, r=30, t=0, b=0),
                         legend=dict(orientation="h", yanchor="bottom", y=1, xanchor="center", x=0.33,
                                    font=dict(size=9)))
        fig.update_xaxes(title=dict(text=x_label, font=dict(size=10)), tickmode='array',
                        tickvals=[i + 0.5 for i in range(nx)], ticktext=x_labels, tickangle=-20,
                        tickfont=dict(size=9, color="#475569"),
                        showgrid=False, showline=False, row=1, col=1)
        fig.update_yaxes(title=dict(text=y_label, font=dict(size=10)), tickmode='array',
                        tickvals=[i + 0.5 for i in range(ny)], ticktext=y_labels,
                        tickfont=dict(size=9, color="#475569"),
                        showgrid=False, showline=False, row=1, col=1)
        fig.update_xaxes(visible=False, row=1, col=2)
        fig.update_yaxes(visible=False, row=1, col=2)

        fig.layout.xaxis.range = [0, nx]
        fig.layout.xaxis.autorange = False
        fig.layout.yaxis.range = [0, ny]
        fig.layout.yaxis.autorange = False
        fig.add_shape(type="rect", xref="x", yref="y", x0=0, x1=nx, y0=0, y1=ny,
                      fillcolor="rgba(0,0,0,0)", line=dict(color="#CBD5E1", width=0.8),
                      layer="above", row=1, col=1)

        st.markdown('<div style="display: flex; justify-content: center; margin: 0 auto;">', unsafe_allow_html=True)
        st.plotly_chart(fig, use_container_width=False)
        st.markdown('</div>', unsafe_allow_html=True)

    # ----- 堆叠分布图（与寿险版相似，传入数据不同）-----
    def calc_stack_distribution(df_raw, selected_types, latest_year, calc_func, x_bins_custom=None):
        """
        计算各公司类型在分箱区间的分布
        返回：distribution_df, labels, company_names_in_bin
        """
        df_year = df_raw[df_raw['报告年份'].astype(str) == str(latest_year)].copy()
        company_ratios = {}
        company_type_map = {}
        for co in df_year['公司'].unique():
            df_co = df_year[df_year['公司'] == co]
            ratio = calc_func(df_co)
            if ratio is None or np.isnan(ratio):
                continue
            company_ratios[co] = ratio
            ct = df_co['公司类型'].iloc[0] if not df_co.empty else None
            if ct:
                company_type_map[co] = ct

        if not company_ratios:
            return pd.DataFrame(), [], {}

        ratios_list = list(company_ratios.values())
        if x_bins_custom and len(x_bins_custom) > 1:
            bins = sorted(set(x_bins_custom))
            actual_min, actual_max = min(ratios_list), max(ratios_list)
            if actual_min < bins[0]:
                bins = [actual_min] + bins
            if actual_max > bins[-1]:
                bins = bins + [actual_max]
        else:
            actual_min, actual_max = min(ratios_list), max(ratios_list)
            n_bins = min(6, max(2, len(set(ratios_list))))
            if actual_max == actual_min:
                bins = [actual_min - 1, actual_max + 1]
            else:
                step = (actual_max - actual_min) / n_bins
                bins = [actual_min + i * step for i in range(n_bins + 1)]

        def fmt(v):
            return f"{int(v)}" if v == int(v) else f"{v:.1f}"
        labels = [f"({fmt(bins[i])}%, {fmt(bins[i+1])}%]" for i in range(len(bins) - 1)]

        distribution = {ct: {lbl: 0 for lbl in labels} for ct in selected_types}
        company_names_in_bin = {ct: {lbl: [] for lbl in labels} for ct in selected_types}

        def get_bin_idx(val, bins):
            if val < bins[0] or val > bins[-1]:
                return -1
            if val == bins[0]:
                return 0
            for i in range(len(bins) - 1):
                if bins[i] < val <= bins[i + 1]:
                    return i
            return -1

        for co, ratio in company_ratios.items():
            ct = company_type_map.get(co)
            if ct in selected_types:
                idx = get_bin_idx(ratio, bins)
                if idx != -1:
                    distribution[ct][labels[idx]] += 1
                    company_names_in_bin[ct][labels[idx]].append(co)

        distribution_df = pd.DataFrame(distribution).T.reindex(selected_types).fillna(0)
        return distribution_df, labels, company_names_in_bin

    def create_stack_chart_and_table(distribution_df, labels, metric_name, target_year, show_labels, label_size, company_names_in_bin=None):
        ct_list = [ct for ct in distribution_df.index if distribution_df.loc[ct].sum() > 0]
        if not ct_list:
            st.warning("没有有效数据")
            return

        fig = go.Figure()
        for ct in ct_list:
            values = distribution_df.loc[ct].values
            color = COMPANY_TYPE_COLORS.get(ct, "#94A3B8")
            hover_texts = []
            names_per_bin = (company_names_in_bin or {}).get(ct, {})
            for lbl in labels:
                names = names_per_bin.get(lbl, [])
                hover_texts.append("、".join(names) if names else "（无公司）")
            fig.add_trace(go.Bar(
                x=labels, y=values, name=ct, marker_color=color, width=0.45,
                text=[f"{int(v)}" if show_labels and v > 0 else "" for v in values],
                textposition='inside', textfont=dict(size=label_size, color="white"), textangle=0,
                customdata=hover_texts,
                hovertemplate="%{customdata}<extra>%{fullData.name}</extra>"
            ))

        total_per_bin = distribution_df.sum(axis=0).values
        max_total = int(max(total_per_bin)) if len(total_per_bin) > 0 else 1
        y_max = max_total + 1
        tick_step = 1 if y_max <= 10 else 2 if y_max <= 20 else 5 if y_max <= 50 else 10
        tick_vals = list(range(0, y_max + 1, tick_step))
        if tick_vals[-1] != y_max:
            tick_vals.append(y_max)
        tick_text = [f"{i}家" for i in tick_vals]
        tick_angle = -15 if len(labels) > 5 else 0

        fig.update_layout(
            title=dict(text=f"{metric_name}分布 - {target_year}年", x=0.5, xanchor='center', font=dict(size=14, color="#00338D")),
            barmode='stack', bargap=0.02, bargroupgap=0, height=280,
            width=900, autosize=False,
            paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)',
            xaxis=dict(tickfont=dict(size=9), tickangle=tick_angle, showgrid=True, gridcolor="#E8ECF1",
                       categoryorder='array', categoryarray=labels, title=""),
            yaxis=dict(title="公司数量", range=[0, y_max], tickvals=tick_vals, ticktext=tick_text,
                       showgrid=True, gridcolor="#E8ECF1", zeroline=True, zerolinecolor="#ccc",
                       title_font=dict(size=10)),
            legend=dict(orientation="h", yanchor="top", y=-0.18, xanchor="center", x=0.5, font=dict(size=9)),
            margin=dict(t=40, b=0, l=40, r=20)
        )
        col1, col2, col3 = st.columns([1, 2.5, 1])
        with col2:
            st.plotly_chart(fig, use_container_width=False, config={'displayModeBar': False})

        # 明细表格
        html_rows = []
        for i, ct in enumerate(ct_list):
            color = COMPANY_TYPE_COLORS.get(ct, "#94A3B8")
            values = distribution_df.loc[ct].values
            total = int(sum(values))
            row_bg = "#F8FAFC" if i % 2 == 0 else "white"
            row_cells = f'<td style="padding: 3px 5px; border: 1px solid #EAEAEA; background-color: {color}; color: white; font-weight: bold; font-size: 9px; text-align: left;">{ct}</td>'
            for v in values:
                row_cells += f'<td style="padding: 3px 5px; text-align: center; border: 1px solid #EAEAEA; background-color: {row_bg}; font-size: 9px;">{int(v)}家</td>'
            row_cells += f'<td style="padding: 3px 5px; text-align: center; border: 1px solid #EAEAEA; background-color: {row_bg}; font-weight: bold; font-size: 9px;">{total}家</td>'
            html_rows.append(f"<tr>{row_cells}</tr>")

        html_table = f"""
        <div style="margin-top: 0px; max-height: 300px; overflow-y: auto; display: flex; justify-content: center;">
            <div>
                <p style="font-size: 11px; font-weight: bold; margin-bottom: 3px; margin-top: 0px; text-align: left;">分布数据明细</p>
                <table style="width: 1000px; border-collapse: collapse; font-family: sans-serif;">
                    <thead>
                        <tr style="background-color: #00338D; color: white;">
                            <th style="padding: 4px 5px; border: 1px solid white; text-align: left; font-size: 9px;">公司类型</th>
                            {''.join([f'<th style="padding: 4px 5px; border: 1px solid white; text-align: center; font-size: 9px;">{label}</th>' for label in labels])}
                            <th style="padding: 4px 5px; border: 1px solid white; text-align: center; font-size: 9px;">合计</th>
                        </tr>
                    </thead>
                    <tbody>
                        {''.join(html_rows)}
                    </tbody>
                </table>
            </div>
        </div>
        """
        st.markdown(html_table, unsafe_allow_html=True)

    # ----- 构成图（费用结构、资产结构、利润构成）-----
    def create_composition_chart(results_by_year, year_list, selected_types, config, field_mapping, color_mapping=None):
        """
        通用构成图（可处理费用、资产、利润等）
        results_by_year: {year: {company_type: {field: value}}}
        field_mapping: {原始字段名: 显示名}
        """
        from plotly.subplots import make_subplots
        all_cts = [ct for ct in selected_types if ct in results_by_year.get(year_list[0], {})]
        if not all_cts:
            return go.Figure()

        fig = make_subplots(rows=1, cols=len(all_cts), shared_yaxes=True,
                            horizontal_spacing=0.015,
                            column_titles=[f"<b>{ct}</b>" for ct in all_cts])

        dark_colors = {"rgb(30,73,226)", "rgb(114,19,234)", "rgb(0,163,161)"}
        for i, ct in enumerate(all_cts):
            col_idx = i + 1
            x_vals = [f"{year}年" for year in year_list]
            cumulative = [0.0] * len(year_list)
            for field, display_name in field_mapping.items():
                y_vals = [results_by_year.get(year, {}).get(ct, {}).get(display_name, 0) for year in year_list]
                color = color_mapping.get(field, "#888") if color_mapping else "#888"
                is_dark = color in dark_colors
                fig.add_trace(go.Bar(
                    x=x_vals, y=y_vals, name=display_name,
                    marker_color=color, width=config.get('bar_width', 0.6),
                    text=[f"{v:.1f}%" if config.get('show_labels', True) and abs(v) > 0.5 else "" for v in y_vals],
                    textposition='inside', insidetextanchor='middle',
                    textfont=dict(size=config.get('label_size', 10), color="white" if is_dark else "black"),
                    showlegend=(i == 0), legendgroup=display_name
                ), row=1, col=col_idx)

            # 灰色边框
            fig.add_shape(type="rect", xref="x domain" if col_idx == 1 else f"x{col_idx} domain",
                          yref="y domain", x0=-0.06, x1=1.06, y0=-0.1, y1=1.15,
                          fillcolor="rgba(0,0,0,0)", line=dict(color="#E0E0E0", width=1),
                          layer="below", row=1, col=col_idx)

        fig.update_layout(
            barmode='stack', bargap=0.05, height=380,
            margin=dict(t=100, b=20, l=150, r=150),
            legend=dict(orientation="h", yanchor="top", y=-0.4, xanchor="center", x=0.5, font=dict(size=10)),
            plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)'
        )
        for i in range(1, len(all_cts) + 1):
            fig.update_xaxes(showgrid=False, showline=False, zeroline=False, ticks="", ticklen=0, row=1, col=i)
            fig.update_yaxes(showgrid=False, range=[0, 100], tickvals=[0, 25, 50, 75, 100],
                            ticktext=["0%", "25%", "50%", "75%", "100%"], zeroline=True, zerolinecolor="#E0E0E0", row=1, col=i)
        for ann in fig.layout.annotations:
            if "<b>" in str(ann.text):
                ann.update(y=1.08, font=dict(size=config.get('co_font_size', 12), color="#00338D"))
        return fig

    # ----- 利润构成图（类似寿险的 create_profit_composition_chart）-----
    def create_profit_comp_chart(results_by_ct, config):
        """财险承保利润构成图（已赚保费、赔付、费用、承保利润的贡献）"""
        import plotly.graph_objects as go
        if not results_by_ct:
            return go.Figure().add_annotation(text="无有效数据", x=0.5, y=0.5, showarrow=False), None

        display_mapping = [
            ("已赚保费", "已赚保费", "rgb(30,73,226)"),
            ("赔付支出", "赔付支出", "rgb(253,52,156)"),
            ("手续费及佣金", "手续费及佣金", "rgb(254,174,215)"),
            ("业务及管理费", "业务及管理费", "rgb(0,163,161)"),
            ("承保利润", "承保利润", "rgb(9,142,126)")
        ]
        ct_list = list(results_by_ct.keys())
        x_indices = list(range(len(ct_list)))
        show_labels, label_size, bar_width, co_font_size = config.get('show_labels', True), config.get('label_size', 11), config.get('bar_width', 0.35), config.get('co_font_size', 12)
        dark_colors = {"rgb(30,73,226)", "rgb(9,142,126)"}

        all_data = {ct: [results_by_ct[ct].get(name, 0) for name, _, _ in display_mapping] for ct in ct_list}
        pos_sums = [sum(v for v in all_data[ct] if v > 0) for ct in ct_list]
        neg_sums = [sum(v for v in all_data[ct] if v < 0) for ct in ct_list]
        y_max = (max(pos_sums) + 20) if pos_sums else 120
        y_min = (min(neg_sums) - 20) if neg_sums else -20

        fig = go.Figure()
        for idx, (_, legend_name, color) in enumerate(display_mapping):
            fig.add_trace(go.Bar(name=legend_name, x=x_indices, y=[all_data[ct][idx] for ct in ct_list],
                                 width=bar_width, marker_color=color,
                                 hovertemplate="%{fullData.name}<br>%{y:.1f}%<extra></extra>"))

        if show_labels:
            for i, ct in enumerate(ct_list):
                pos_cursor, neg_cursor = 0, 0
                for j, (_, _, color) in enumerate(display_mapping):
                    v = all_data[ct][j]
                    txt_color = "white" if color in dark_colors else "black"
                    if v >= 0:
                        center_y = pos_cursor + v / 2
                        pos_cursor += v
                    else:
                        center_y = neg_cursor + v / 2
                        neg_cursor += v
                    if abs(v) >= 1:
                        fig.add_annotation(x=i, y=center_y, text=f"{v:.1f}%", showarrow=False,
                                           xanchor="center", yanchor="middle",
                                           font=dict(size=label_size, color=txt_color))

        for i in range(len(ct_list)):
            fig.add_shape(type="rect", xref="x", yref="paper", x0=i - 0.46, x1=i + 0.46,
                          y0=0, y1=1, fillcolor="rgba(0,0,0,0)", line=dict(color="#E0E0E0", width=1), layer="below")

        fig.update_layout(
            barmode="relative", bargap=0.05, height=380,
            paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
            margin=dict(t=50, b=20, l=200, r=200),
            legend=dict(orientation="v", yanchor="middle", y=0.5, xanchor="right", x=-0.05,
                        traceorder="reversed", font=dict(size=11, color="#00338D")),
            yaxis=dict(side="right", showgrid=False, range=[y_min, y_max],
                       tickformat=".0f", ticksuffix="%", tickmode="auto", nticks=8,
                       zeroline=True, zerolinecolor="#F7860C", zerolinewidth=2)
        )
        fig.update_xaxes(tickvals=x_indices, ticktext=[f"<span style='font-size:{co_font_size}px;color:#00338D;'><b>{ct}</b></span>" for ct in ct_list],
                         side="top", showgrid=False, zeroline=False)
        return fig, None

    # ============================================================
    # 5️⃣ 注释表加载（与寿险版结构一致，仅替换默认链接）
    # ============================================================
    notes_dict_8, ordered_modules, first_levels = {}, [], []
    df_notes = None

    def generate_custom_analysis(m_id, df, cos, cy, py):
        """财险版自动生成分析文字（简化版本）"""
        # 这里可根据需要实现类似寿险的文字生成，为简洁起见，返回空或简单描述
        # 实际项目中可在此处调用计算函数生成文字
        return ""

    with st.expander("📥 行业分析注释输入", expanded=False):
        use_default_bins = st.toggle("使用默认坐标轴刻度表", value=True, key="step8_use_default_bins")
        if use_default_bins:
            try:
                default_bins_url = "https://raw.githubusercontent.com/z-xylym/my-actuary-tool/main/%E5%9D%90%E6%A0%87%E8%BD%B4%E5%88%BB%E5%BA%A6%E8%A1%A8-step8%20%E7%9A%84%E5%89%AF%E6%9C%AC.xlsx"  # 可替换为财险版
                custom_bins_map = load_custom_bins_excel(default_bins_url)
                st.session_state["custom_bins_map"] = custom_bins_map
                st.success(f"✅ 已从云端加载默认刻度表（共 {len(custom_bins_map)} 个图）")
            except Exception as e:
                st.error(f"❌ 云端下载失败：{e}")
        else:
            uploaded_bins_file = st.file_uploader("📊 上传自定义坐标轴刻度表（Excel）", type=["xlsx", "xls"], key="custom_bins_uploader")
            if uploaded_bins_file is not None:
                try:
                    custom_bins_map = load_custom_bins_excel(uploaded_bins_file)
                    st.session_state["custom_bins_map"] = custom_bins_map
                    st.success(f"✅ 已加载 {len(custom_bins_map)} 个图的自定义刻度配置")
                except Exception as e:
                    st.error(f"❌ 刻度表读取失败：{e}")

        use_default = st.toggle("使用默认注释表", value=True, key="step8_use_default")
        if use_default:
            try:
                default_url = "https://raw.githubusercontent.com/z-xylym/my-actuary-tool/main/RD-%E5%9B%BE%E7%89%87%E5%86%85%E5%AE%B9%E5%88%86%E6%9E%90%E5%92%8C%E6%B3%A8%E9%87%8A%E6%A8%A1%E6%9D%BF-step8.xlsx"
                import requests, io
                resp = requests.get(default_url, timeout=15)
                resp.raise_for_status()
                df_notes = pd.read_excel(io.BytesIO(resp.content))
                st.success("✅ 内置默认注释表从云端加载成功！")
            except Exception as e:
                st.error(f"❌ 云端下载失败：{e}")
        else:
            notes_file = st.file_uploader("上传 Excel 分析注释表", type=['xlsx', 'xls'], key="step8_notes")
            if notes_file:
                try:
                    df_notes = pd.read_excel(notes_file)
                    st.success("✅ 自定义注释表上传成功！")
                except Exception as e:
                    st.error(f"❌ 上传文件解析失败: {e}")

        if df_notes is not None:
            # 处理注释表，填充 notes_dict_8
            if '分析内容-自定义' in df_notes.columns:
                df_notes['分析内容-自定义'] = df_notes['分析内容-自定义'].astype('object')
            if '分析内容-默认' in df_notes.columns:
                df_notes['分析内容-默认'] = df_notes['分析内容-默认'].astype('object')

            for idx, row in df_notes.iterrows():
                mid = str(row.get('模块ID', '')).strip()
                if not mid or mid == 'nan':
                    continue
                raw_val = df_notes.loc[idx, '分析内容-自定义']
                if pd.isna(raw_val) or str(raw_val).strip() in ('', 'nan', 'None'):
                    generated = generate_custom_analysis(mid, df_raw, st.session_state.get('step8_selected_types', []), latest_year, prev_year)
                    if generated:
                        df_notes.loc[idx, '分析内容-自定义'] = str(generated)

            for col in df_notes.columns:
                df_notes[col] = df_notes[col].astype(str).str.strip()

            required_cols = ['模块ID', '一级分类', '二级分类', '对应图表名称', '分析内容-默认', '分析内容-自定义', '注释内容']
            for col in required_cols:
                if col not in df_notes.columns:
                    df_notes[col] = ''

            for col in ['一级分类', '二级分类', '对应图表名称', '模块ID']:
                if col in df_notes.columns:
                    df_notes[col] = df_notes[col].replace(['nan', 'NaN', 'NAN', 'None'], '')

            if '二级分类' in df_notes.columns:
                df_notes['二级分类'] = df_notes['二级分类'].apply(lambda x: "全部" if str(x).strip() == "" else str(x).strip())

            for _, r in df_notes.iterrows():
                m_id = str(r.get('模块ID', '')).strip()
                if not m_id:
                    continue
                notes_dict_8[m_id] = {
                    'title': str(r.get('对应图表名称', '')).strip(),
                    'analysis_default': str(r.get('分析内容-默认', '')).strip(),
                    'analysis_custom': str(r.get('分析内容-自定义', '')).strip(),
                    'note': str(r.get('注释内容', '')).strip(),
                    '一级分类': str(r.get('一级分类', '')).strip(),
                    '二级分类': str(r.get('二级分类', '')).strip(),
                }
                if m_id not in ordered_modules:
                    ordered_modules.append(m_id)

            first_levels = [x for x in df_notes['一级分类'].unique() if x and x != '']
            st.session_state['step8_notes_dict'] = notes_dict_8
            st.session_state['step8_ordered_modules'] = ordered_modules
            st.session_state['step8_df_notes'] = df_notes.copy()

    if df_notes is None and 'step8_df_notes' in st.session_state:
        df_notes = st.session_state['step8_df_notes'].copy()
        notes_dict_8 = st.session_state.get('step8_notes_dict', {})
        ordered_modules = st.session_state.get('step8_ordered_modules', [])
        first_levels = [x for x in df_notes['一级分类'].unique() if x and x != ''] if df_notes is not None else []

    # ---------- 侧边栏导航 ----------
    print_mode = False
    active_m_id = None
    active_chart_name = None
    with st.sidebar:
        st.markdown("<h3 style='color: #00338D; font-size: 18px;'>行业分析导航</h3>", unsafe_allow_html=True)
        if first_levels:
            main_nav = st.radio("📁 一级模块", first_levels + ["🖨️ 一键显示全部 (打印/导出)"], key="step8_main")
            if main_nav == "🖨️ 一键显示全部 (打印/导出)":
                print_mode = True
                st.info("💡 点击下方按钮导出 PDF")
                components.html("""<button onclick="window.parent.print()" style="width:100%; padding:12px; background:#00338D; color:white; border:none; border-radius:6px; cursor:pointer;">立即导出 PDF 报告</button>""", height=60)
            else:
                df_sub1 = df_notes[df_notes['一级分类'] == main_nav]
                sec_levels = [x for x in df_sub1['二级分类'].unique() if x and x != '']
                if len(sec_levels) == 0:
                    charts = [x for x in df_sub1['对应图表名称'].unique() if x and x != '']
                    chart_nav = st.radio("📊 具体图表", charts, key="step8_chart")
                    row = df_sub1[df_sub1['对应图表名称'] == chart_nav].iloc[0]
                    active_m_id = row['模块ID']
                else:
                    sub_nav = st.radio("📂 二级模块", ["全部"] + sec_levels, key="step8_sub")
                    if sub_nav != "全部":
                        df_sub2 = df_sub1[df_sub1['二级分类'] == sub_nav]
                        charts = [x for x in df_sub2['对应图表名称'].unique() if x and x != '']
                        chart_nav = st.radio("📊 具体图表", charts, key="step8_chart")
                        row = df_sub2[df_sub2['对应图表名称'] == chart_nav].iloc[0]
                        active_m_id = row['模块ID']
                    else:
                        charts = [x for x in df_sub1['对应图表名称'].unique() if x and x != '']
                        chart_nav = st.radio("📊 具体图表", charts, key="step8_chart")
                        row = df_sub1[df_sub1['对应图表名称'] == chart_nav].iloc[0]
                        active_m_id = row['模块ID']
        else:
            st.warning("⚠️ 请先上传包含层级信息的注释表")
            return

    # ---------- 行业分析配置（公司类型选择） ----------
    with st.expander("⚙️ 行业分析配置", expanded=False):
        c1, c2, c3 = st.columns(3)
        with c1:
            all_types = sorted([str(x) for x in df_raw['公司类型'].dropna().unique()])
            default_order = [t for t in DEFAULT_TYPE_ORDER if t in all_types]
            default_order += [t for t in all_types if t not in default_order]
            selected_types = st.multiselect("🏢 选择公司类型（可多选）", default_order, default=default_order)
            st.session_state['step8_selected_types'] = selected_types
        with c2:
            st.markdown("**📊 同类型公司展示顺序**")
            sort_field = st.selectbox("排序依据字段", COMPANY_SORT_FIELD_OPTIONS, index=0, key="step8_company_sort_field_selector")
            st.session_state['step8_company_sort_field'] = sort_field
        with c3:
            st.markdown("**📋 公司类型显示顺序**")
            type_order = st.multiselect("拖拽或重新选择顺序", default_order, default=default_order, key="step8_type_order_selector")
            full_order = type_order + [t for t in default_order if t not in type_order]
            st.session_state['step8_type_order'] = full_order

    company_sort_map = get_company_sort_map(df_raw, latest_year, st.session_state.get('step8_company_sort_field', COMPANY_SORT_FIELD_OPTIONS[0]))

    # ---------- 辅助函数 ----------
    def show_chart(fig, p_mode):
        if fig:
            if p_mode:
                fig.update_layout(width=1500, autosize=False)
                st.plotly_chart(fig, use_container_width=False)
            else:
                st.plotly_chart(fig, use_container_width=True)

    def get_uploaded_bins_by_mid(m_id):
        custom_map = st.session_state.get("custom_bins_map", {})
        cfg = custom_map.get(m_id, {})
        return cfg.get("x_bins_custom") or [], cfg.get("y_bins_custom") or []

    def bins_to_str(bins):
        if not bins:
            return ""
        return ",".join(str(int(x)) if x == int(x) else str(x) for x in bins)

    def parse_bins_input_safe(text, fallback):
        try:
            result = parse_bins_input(text)
            return result if result else fallback
        except Exception:
            return fallback

    def render_axis_bin_popover(m_id, title, default_x_bins, default_y_bins):
        uploaded_x, uploaded_y = get_uploaded_bins_by_mid(m_id)
        x_text_key = f"x_bins_text_{m_id}"
        y_text_key = f"y_bins_text_{m_id}"
        if x_text_key not in st.session_state:
            st.session_state[x_text_key] = bins_to_str(uploaded_x or default_x_bins)
        if y_text_key not in st.session_state:
            st.session_state[y_text_key] = bins_to_str(uploaded_y or default_y_bins)
        with st.popover(f"⚙️ 当前图参数设置：{title}", use_container_width=True):
            if uploaded_x or uploaded_y:
                st.info(f"📊 已使用上传表的刻度，不可手动修改\nX = {bins_to_str(uploaded_x) or '（空）'}；Y = {bins_to_str(uploaded_y) or '（空）'}")
            else:
                with st.form(key=f"bins_form_{m_id}"):
                    x_input = st.text_input("X 轴刻度（英文逗号分隔）", value=st.session_state[x_text_key], placeholder=bins_to_str(default_x_bins))
                    y_input = st.text_input("Y 轴刻度（英文逗号分隔）", value=st.session_state[y_text_key], placeholder=bins_to_str(default_y_bins))
                    if st.form_submit_button("应用", use_container_width=True):
                        st.session_state[x_text_key] = x_input
                        st.session_state[y_text_key] = y_input
                        st.rerun()
        if uploaded_x or uploaded_y:
            x_bins = uploaded_x or default_x_bins
            y_bins = uploaded_y or default_y_bins
        else:
            x_bins = parse_bins_input_safe(st.session_state[x_text_key], default_x_bins)
            y_bins = parse_bins_input_safe(st.session_state[y_text_key], default_y_bins)
        return x_bins, y_bins

    def get_axis_bins_for_mid(m_id, default_x_str, default_y_str):
        uploaded_x, uploaded_y = get_uploaded_bins_by_mid(m_id)
        use_upload_key = f"use_uploaded_bins_{m_id}"
        x_text_key = f"x_bins_text_{m_id}"
        y_text_key = f"y_bins_text_{m_id}"
        default_x = parse_bins_input(default_x_str)
        default_y = parse_bins_input(default_y_str)
        if st.session_state.get(use_upload_key) and (uploaded_x or uploaded_y):
            x_bins = uploaded_x or default_x
            y_bins = uploaded_y or default_y
        else:
            x_bins = parse_bins_input_safe(st.session_state.get(x_text_key, ""), default_x)
            y_bins = parse_bins_input_safe(st.session_state.get(y_text_key, ""), default_y)
        return x_bins or default_x, y_bins or default_y

    # ============================================================
    # 6️⃣ 主渲染函数 render_pure_chart_entity（路由所有 m_id）
    # ============================================================
    def render_pure_chart_entity(m_id, print_mode):
        years = sorted([int(y) for y in df_raw['报告年份'].dropna().astype(str).str.replace(".0", "", regex=False).unique() if y.isdigit()])
        latest_year, prev_year = years[-1], years[-2] if len(years) > 1 else years[-1] - 1
        selected_types = st.session_state.get('step8_selected_types', [])

        # 1) 散点图（SCATTER_AXIS_META 中定义的指标）
        if m_id in SCATTER_AXIS_META:
            meta = SCATTER_AXIS_META[m_id]
            df_scatter = calc_scatter_data(df_raw, selected_types, latest_year, prev_year, meta["x_field"], meta["title"])
            if df_scatter.empty:
                st.warning(f"缺少 {meta['x_field']} 字段")
                return
            x_col, y_col = meta["title"], f"{meta['title']}增长率"
            auto_x_bins = generate_nice_bins(df_scatter[x_col])
            auto_y_bins = generate_nice_bins(df_scatter[y_col])
            if not print_mode:
                _, btn_col = st.columns([6, 1])
                with btn_col:
                    x_bins, y_bins = render_axis_bin_popover(m_id, meta["title"], auto_x_bins, auto_y_bins)
            else:
                uploaded_x_print, uploaded_y_print = get_uploaded_bins_by_mid(m_id)
                if uploaded_x_print or uploaded_y_print:
                    st.session_state[f"use_uploaded_bins_{m_id}"] = True
                x_bins, y_bins = get_axis_bins_for_mid(m_id, bins_to_str(auto_x_bins), bins_to_str(auto_y_bins))
            config = {"x_col": x_col, "y_col": y_col, "x_bins_custom": x_bins, "y_bins_custom": y_bins}
            create_scatter_chart(df_scatter, config, meta["x_label"], meta["y_label"])
            return

        # 2) 堆叠分布图（STACK_DIST_META）
        elif m_id in STACK_DIST_META:
            meta = STACK_DIST_META[m_id]
            uploaded_x, _ = get_uploaded_bins_by_mid(m_id)
            x_text_key = f"x_bins_text_{m_id}"
            if x_text_key not in st.session_state:
                st.session_state[x_text_key] = bins_to_str(uploaded_x) if uploaded_x else meta["default_x"]
            if not print_mode:
                _, btn_col = st.columns([6, 1])
                with btn_col:
                    with st.popover(f"⚙️ {meta['name']} 参数设置", use_container_width=True):
                        show_labels = st.toggle("显示数值标签", value=True, key=f"lab_{m_id}")
                        label_size = st.slider("标签大小", 8, 16, 11, key=f"sz_{m_id}")
                        st.markdown("---")
                        if uploaded_x:
                            st.info(f"📊 已使用上传 Excel 的刻度，不可手动修改\nX = {bins_to_str(uploaded_x)}")
                            st.session_state[f"use_uploaded_bins_{m_id}"] = True
                        else:
                            with st.form(key=f"dist_bins_form_{m_id}"):
                                x_input = st.text_input("X 轴区间（英文逗号分隔）", value=st.session_state[x_text_key], placeholder=meta["default_x"])
                                if st.form_submit_button("应用", use_container_width=True):
                                    st.session_state[x_text_key] = x_input
                                    st.rerun()
                # 计算实际 bins
                df_year_tmp = df_raw[df_raw['报告年份'].astype(str) == str(latest_year)].copy()
                ratios_tmp = []
                for co in df_year_tmp['公司'].unique():
                    r = meta["calc_func"](df_year_tmp[df_year_tmp['公司'] == co])
                    if r is not None and not np.isnan(r):
                        ratios_tmp.append(r)
                if ratios_tmp:
                    min_r, max_r = min(ratios_tmp), max(ratios_tmp)
                    if max_r == min_r:
                        auto_bins = [min_r - 1, max_r + 1]
                    else:
                        n = min(6, max(2, len(set(ratios_tmp))))
                        step = (max_r - min_r) / n
                        auto_bins = [min_r + i * step for i in range(n + 1)]
                    auto_bins_str = bins_to_str(auto_bins)
                else:
                    auto_bins_str = meta["default_x"]
                x_bins, _ = get_axis_bins_for_mid(m_id, auto_bins_str, "")
            else:
                # 打印模式
                uploaded_x_print, _ = get_uploaded_bins_by_mid(m_id)
                if uploaded_x_print:
                    st.session_state[f"use_uploaded_bins_{m_id}"] = True
                    st.session_state[x_text_key] = bins_to_str(uploaded_x_print)
                    x_bins = uploaded_x_print
                else:
                    df_year_tmp = df_raw[df_raw['报告年份'].astype(str) == str(latest_year)].copy()
                    ratios_tmp = []
                    for co in df_year_tmp['公司'].unique():
                        r = meta["calc_func"](df_year_tmp[df_year_tmp['公司'] == co])
                        if r is not None and not np.isnan(r):
                            ratios_tmp.append(r)
                    if ratios_tmp:
                        min_r, max_r = min(ratios_tmp), max(ratios_tmp)
                        if max_r == min_r:
                            auto_bins = [min_r - 1, max_r + 1]
                        else:
                            n = min(6, max(2, len(set(ratios_tmp))))
                            step = (max_r - min_r) / n
                            auto_bins = [min_r + i * step for i in range(n + 1)]
                        auto_bins_str = bins_to_str(auto_bins)
                    else:
                        auto_bins_str = meta["default_x"]
                    x_bins, _ = get_axis_bins_for_mid(m_id, auto_bins_str, "")
                show_labels, label_size = True, 11

            distribution_df, labels, company_names_in_bin = calc_stack_distribution(
                df_raw, selected_types, latest_year, meta["calc_func"], x_bins_custom=x_bins
            )
            if distribution_df.empty or not labels:
                st.warning(f"无法计算 {meta['name']} 分布，数据不足")
                return
            st.session_state[f"stack_dist_labels_{m_id}"] = labels
            st.session_state[f"stack_dist_df_{m_id}"] = distribution_df
            create_stack_chart_and_table(distribution_df, labels, meta['name'], latest_year, show_labels, label_size, company_names_in_bin=company_names_in_bin)
            return

        # 3) 构成图：费用结构
        elif m_id == "industry_expense_struct":
            # 收集数据
            results_by_year = {}
            for year in year_list:
                year_results = {}
                for ct in selected_types:
                    res = calc_industry_expense_composition(df_raw, ct, year)
                    if res:
                        year_results[ct] = res['ratios']
                if year_results:
                    results_by_year[year] = year_results
            if not results_by_year:
                st.warning("无法计算费用结构")
                return
            if not print_mode:
                c1, c2, c3 = st.columns(3)
                with c1: show_labels = st.toggle("显示标签", value=True, key=f"lab_{m_id}")
                with c2: label_size = st.slider("标签大小", 8, 16, 10, key=f"sz_{m_id}")
                with c3: bar_width = st.slider("柱宽", 0.2, 0.8, 0.35, key=f"wid_{m_id}")
            else:
                show_labels, label_size, bar_width = True, 10, 0.35
            field_mapping = {"获取费用": "获取费用", "维持费用": "维持费用", "非履约费用": "非履约费用"}
            color_mapping = {"获取费用": "rgb(30,73,226)", "维持费用": "rgb(118,210,255)", "非履约费用": "rgb(114,19,234)"}
            config = {'show_labels': show_labels, 'label_size': label_size, 'bar_width': bar_width, 'co_font_size': 12}
            fig = create_composition_chart(results_by_year, year_list, selected_types, config, field_mapping, color_mapping)
            st.markdown("<p style='text-align:right;font-size:12px;color:#666;'>单位：百分比 (%)，顶部数字为总费用（亿元）</p>", unsafe_allow_html=True)
            show_chart(fig, print_mode)

        # 4) 构成图：资产结构
        elif m_id == "industry_asset_struct":
            results_by_year = {}
            for year in year_list:
                year_results = {}
                for ct in selected_types:
                    res = calc_industry_asset_composition(df_raw, ct, year)
                    if res:
                        year_results[ct] = res
                if year_results:
                    results_by_year[year] = year_results
            if not results_by_year:
                st.warning("无法计算资产结构")
                return
            if not print_mode:
                c1, c2, c3 = st.columns(3)
                with c1: show_labels = st.toggle("显示标签", value=True, key=f"lab_{m_id}")
                with c2: label_size = st.slider("标签大小", 8, 16, 10, key=f"sz_{m_id}")
                with c3: bar_width = st.slider("柱宽", 0.2, 0.8, 0.35, key=f"wid_{m_id}")
            else:
                show_labels, label_size, bar_width = True, 10, 0.35
            field_mapping = {
                "AC（债权投资）": "AC（债权投资）",
                "FVOCI（其他债权投资）": "FVOCI（其他债权投资）",
                "FVTPL（交易性金融资产）": "FVTPL（交易性金融资产）",
                "指定FVOCI（其他权益工具）": "指定FVOCI（其他权益工具）"
            }
            color_mapping = {
                "AC（债权投资）": "rgb(0, 184, 245)",
                "FVOCI（其他债权投资）": "rgb(114, 19, 234)",
                "FVTPL（交易性金融资产）": "rgb(253, 52, 156)",
                "指定FVOCI（其他权益工具）": "rgb(181, 2, 95)"
            }
            config = {'show_labels': show_labels, 'label_size': label_size, 'bar_width': bar_width, 'co_font_size': 12}
            fig = create_composition_chart(results_by_year, year_list, selected_types, config, field_mapping, color_mapping)
            st.markdown("<p style='text-align:right;font-size:12px;color:#666;'>单位：百分比 (%)</p>", unsafe_allow_html=True)
            show_chart(fig, print_mode)

        # 5) 利润构成（承保利润）
        elif m_id == "industry_profit_comp":
            target_year = latest_year
            results_by_ct = {}
            for ct in selected_types:
                res = calc_industry_profit_composition(df_raw, ct, target_year)
                if res:
                    results_by_ct[ct] = res
            if not results_by_ct:
                st.warning("无法计算承保利润构成")
                return
            if not print_mode:
                c1, c2, c3 = st.columns(3)
                with c1: show_labels = st.toggle("显示标签", value=True, key=f"lab_{m_id}")
                with c2: label_size = st.slider("标签字号", 8, 16, 11, key=f"psz_{m_id}")
                with c3: bar_width = st.slider("柱宽", 0.2, 0.8, 0.4, key=f"wid_{m_id}")
            else:
                show_labels, label_size, bar_width = True, 10, 0.35
            config = {'show_labels': show_labels, 'label_size': label_size, 'bar_width': bar_width, 'co_font_size': 14}
            fig, _ = create_profit_comp_chart(results_by_ct, config)
            st.markdown("<p style='text-align:right;font-size:12px;color:#666;'>单位：百分比 (%)</p>", unsafe_allow_html=True)
            show_chart(fig, print_mode)

        # 6) 综合成本率拆解（赔付率 vs 费用率）
        elif m_id == "industry_cor_comp":
            # 计算各类型平均赔付率和费用率
            results_by_year = {}
            for year in year_list:
                year_results = {}
                for ct in selected_types:
                    mask = (df_raw['公司类型'] == ct) & (df_raw['报告年份'].astype(str) == str(year))
                    df_f = df_raw[mask].copy()
                    avg_loss = df_f[df_f['字段名'] == '综合赔付率']['(百万)人民币'].mean()
                    avg_exp = df_f[df_f['字段名'] == '综合费用率']['(百万)人民币'].mean()
                    if not pd.isna(avg_loss) and not pd.isna(avg_exp):
                        year_results[ct] = {'综合赔付率': avg_loss, '综合费用率': avg_exp}
                if year_results:
                    results_by_year[year] = year_results
            if not results_by_year:
                st.warning("无法计算综合成本率拆解")
                return
            field_mapping = {"综合赔付率": "综合赔付率", "综合费用率": "综合费用率"}
            color_mapping = {"综合赔付率": "rgb(253,52,156)", "综合费用率": "rgb(0,163,161)"}
            config = {'show_labels': True, 'label_size': 10, 'bar_width': 0.35, 'co_font_size': 12}
            fig = create_composition_chart(results_by_year, year_list, selected_types, config, field_mapping, color_mapping)
            st.markdown("<p style='text-align:right;font-size:12px;color:#666;'>单位：百分比 (%)</p>", unsafe_allow_html=True)
            show_chart(fig, print_mode)

        else:
            st.info(f"⏳ 模块 [{m_id}] 尚未配置底层绘图代码")

    # ============================================================
    # 7️⃣ 报告包装器 render_report_module（与寿险版一致）
    # ============================================================
    def render_report_module(m_id, print_mode, is_first=False):
        mod_data = notes_dict_8.get(m_id, {})
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

        if print_mode:
            st.markdown("<div class='page-break-container' style='margin:0;padding:0;'>", unsafe_allow_html=True)
        if not print_mode:
            st.markdown("<div class='no-print' style='height:2px; background:linear-gradient(to right, #00338D, #005EBB, #FFFFFF); margin-bottom:15px;'></div>", unsafe_allow_html=True)

        title_cls = "page-break-title" if (print_mode and not is_first) else ""
        mt = "0px" if (print_mode and is_first) else "20px"
        font_size = "35px" if print_mode else "30px"
        st.markdown(
            f"<h3 class='{title_cls}' style='text-align:left; color:#00338D; font-size:{font_size}; font-weight:900; "
            f"font-family:Microsoft YaHei, 微软雅黑, sans-serif; margin-top:{mt}; margin-bottom:20px; border:none; padding-bottom:0px;'>"
            f"{full_title}</h3>", unsafe_allow_html=True
        )

        def clean_note(val):
            if pd.isna(val): return ""
            val_str = str(val).strip()
            return "" if val_str.lower() in ['nan', 'none', 'null', ''] else val_str

        analysis_default = clean_note(mod_data.get('analysis_default', ''))
        analysis_custom = clean_note(mod_data.get('analysis_custom', ''))
        if analysis_default or analysis_custom:
            html = '<div style="background-color:#F4F7FC; border-left:4px solid #00338D; padding:5px 5px; margin-bottom:5px; text-align:left; border-radius:3px;">'
            if analysis_default:
                html += f'<p style="margin:0; color:#0A1F5C; font-size:12px; line-height:1.4;">{analysis_default}</p>'
            if analysis_custom:
                mt_space = "4px" if analysis_default else "0px"
                html += f'<p style="margin:{mt_space} 0 0 0; color:#002678; font-size:13px; line-height:1.4;">{analysis_custom}</p>'
            html += '</div>'
            st.markdown(html, unsafe_allow_html=True)

        if print_mode:
            render_pure_chart_entity(m_id, print_mode)
        else:
            chart_col_left, chart_col_center, chart_col_right = st.columns([1, 10, 1])
            with chart_col_center:
                render_pure_chart_entity(m_id, print_mode)

        note_text = clean_note(mod_data.get('note', ''))
        if note_text:
            st.markdown(
                f'<div style="margin-top:2px; margin-bottom:20px; text-align:left;">'
                f'<p style="margin:0; color:#888; font-size:12px; font-style:italic; line-height:1.4;">注：{note_text}</p></div>',
                unsafe_allow_html=True
            )
        if print_mode:
            st.markdown("</div>", unsafe_allow_html=True)

    # ============================================================
    # 8️⃣ 最终执行器（网页/打印模式）
    # ============================================================
    if not print_mode:
        st.markdown("<hr class='no-print' style='border:none;border-top:1px solid #EAEAEA;margin:10px 0;'>", unsafe_allow_html=True)

    if print_mode:
        if 'ordered_modules' not in locals() or not ordered_modules:
            st.warning("⚠️ 报告顺序由【模块ID】的先后顺序决定，请先在上方传入有模块ID的注释表文件。")
        else:
            import datetime
            today = datetime.date.today()
            date_str = f"{today.year}年{today.month}月"
            type_str = st.session_state.get('step8_selected_type', '')
            if type_str == "全部":
                type_str = ""
            cover_url = "https://raw.githubusercontent.com/z-xylym/my-actuary-tool/main/%E6%A0%87%E9%A2%98%E9%A1%B5.png"
            back_url = "https://raw.githubusercontent.com/z-xylym/my-actuary-tool/main/%E5%B0%81%E5%BA%95%E9%A1%B5.png"

            st.markdown(f"""
            <div style="position:relative; width:338.67mm; height:175.5mm; page-break-after:always; overflow:hidden; margin:0; padding:0;
                -webkit-print-color-adjust:exact; print-color-adjust:exact; forced-color-adjust:none;">
                <img src="{cover_url}" style="width:100%; height:100%; object-fit:cover; display:block;"/>
                <div style="position:absolute; top:0; left:0; width:100%; height:100%;
                    display:flex; flex-direction:column; justify-content:center; align-items:flex-start;
                    padding:0 8%; box-sizing:border-box; margin-top:-30px; z-index:10;
                    forced-color-adjust:none; -webkit-print-color-adjust:exact; print-color-adjust:exact;">
                    <div style="font-size:52px; font-weight:900; line-height:1.4; margin-bottom:16px;
                        font-family:Microsoft YaHei,微软雅黑,sans-serif;
                        color:white; -webkit-text-fill-color:white;
                        text-shadow:2px 2px 4px rgba(0,0,0,0.5), 0 0 20px rgba(0,0,0,0.3);
                        forced-color-adjust:none; -webkit-print-color-adjust:exact;">
                        {latest_year}年行业表现和洞察<br>{type_str}财险公司<br>
                    </div>
                    <div style="font-size:22px; font-weight:500; margin:0;
                        font-family:Microsoft YaHei,微软雅黑,sans-serif;
                        color:white; -webkit-text-fill-color:white;
                        text-shadow:1px 1px 3px rgba(0,0,0,0.5);
                        forced-color-adjust:none; -webkit-print-color-adjust:exact;">{date_str}</div>
                </div>
            </div>
            """, unsafe_allow_html=True)

            for i, mod in enumerate(ordered_modules):
                render_report_module(mod, print_mode=True, is_first=(i == 0))

            st.markdown(f"""
            <div style="position:relative; width:338.67mm; height:175.5mm; page-break-before:always; overflow:hidden; margin:0; padding:0;
                -webkit-print-color-adjust:exact; print-color-adjust:exact; forced-color-adjust:none;">
                <img src="{back_url}" style="width:100%; height:100%; object-fit:cover; display:block;"/>
            </div>
            """, unsafe_allow_html=True)
    else:
        if active_m_id:
            render_report_module(active_m_id, print_mode=False, is_first=True)
        else:
            st.info("💡 请从左侧导航栏选择要查看的行业分析模块")