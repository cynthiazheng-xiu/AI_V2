# app.py - AI价到 - 小微外贸智能报价助手 (表格版)

import streamlit as st
import pandas as pd
import math
import subprocess
import os
import time
import hashlib
import json
from datetime import datetime, timedelta, timezone
from pathlib import Path

# -------------------- 页面配置 - 必须放在最前面 --------------------
st.set_page_config(
    page_title="AI价到 - 小微外贸智能出口报价助手",
    page_icon="💰",
    layout="wide",
    initial_sidebar_state="expanded"
)

# -------------------- 常量定义 --------------------
BASE_DATA_PATH = "C:\\Basic Information"
EXCEL_FILE = os.path.join(BASE_DATA_PATH, "Data.xlsx")  # 使用os.path.join确保正确的路径分隔符

# 工作表名称常量
SHEET_PORTS = "港口信息表"
SHEET_RATES = "汇率表"
SHEET_HS = "HS表"
SHEET_PRODUCTS = "商品信息表"
SHEET_CUSTOMERS = "客户信息表"

# 默认汇率（作为备用）
DEFAULT_RATES = {
    "USD": 6.9257, "EUR": 8.1863, "GBP": 9.3729, "JPY": 0.044775,
    "HKD": 0.8858, "AUD": 4.9092, "CAD": 5.0734, "CHF": 8.9762, "SGD": 5.4721
}

# 集装箱参数表
CONTAINER_SPECS = {
    "20'GP": {"type": "普通", "code": "GP", "volume": 33, "weight": 25000, "tare": 2275, "display": "20' 普通"},
    "40'GP": {"type": "普通", "code": "GP", "volume": 67, "weight": 29000, "tare": 3760, "display": "40' 普通"},
    "40'HC": {"type": "普通", "code": "HC", "volume": 76, "weight": 29000, "tare": 3950, "display": "40' 高箱"},
    "20'RF": {"type": "冷冻", "code": "RF", "volume": 27, "weight": 21000, "tare": 2900, "display": "20' 冷冻"},
    "40'RF": {"type": "冷冻", "code": "RF", "volume": 58, "weight": 26000, "tare": 4330, "display": "40' 冷冻"},
    "40'RH": {"type": "冷冻", "code": "RH", "volume": 66, "weight": 26000, "tare": 4560, "display": "40' 冷冻高箱"}
}

# 简化的集装箱列表（用于运费设置）
CONTAINER_TYPES = [
    {"name": "20'GP", "display": "20' 普通", "volume": 33, "weight": 25000},
    {"name": "40'GP", "display": "40' 普通", "volume": 67, "weight": 29000},
    {"name": "40'HC", "display": "40' 高箱", "volume": 76, "weight": 29000},
    {"name": "20'RF", "display": "20' 冷冻", "volume": 27, "weight": 21000},
    {"name": "40'RF", "display": "40' 冷冻", "volume": 58, "weight": 26000},
    {"name": "40'RH", "display": "40' 冷冻高箱", "volume": 66, "weight": 26000}
]

# 2020版国际贸易术语
INCOTERMS_2020 = [
    {"code": "EXW", "name": "EXW (工厂交货)", "full_name": "Ex Works", "description": "卖方在其所在地或其他指定地点将货物交给买方处置时即完成交货。"},
    {"code": "FOB", "name": "FOB (船上交货)", "full_name": "Free On Board", "description": "卖方在指定装运港将货物装到买方指定的船上即完成交货。"},
    {"code": "CIF", "name": "CIF (成本、保险费加运费)", "full_name": "Cost, Insurance and Freight", "description": "卖方支付将货物运至指定目的港的运费和保险费。"}
]

# -------------------- 自定义CSS样式 --------------------
st.markdown("""
<style>
    .main-header {
        background: linear-gradient(135deg, #0A174E 0%, #1D2B5E 100%);
        padding: 1.5rem 2rem;
        border-radius: 20px;
        margin-bottom: 1rem;
        text-align: center;
    }
    .main-header h1 {
        color: white !important;
        margin: 0;
        font-size: 2.5rem;
        font-weight: bold;
        text-shadow: 2px 2px 4px rgba(0,0,0,0.3);
    }
    .main-header .subtitle {
        color: #FFD700 !important;
        font-size: 1.1rem;
        margin-top: 0.5rem;
    }
    .budget-table {
        width: 100%;
        border-collapse: collapse;
        margin: 20px 0;
        font-size: 0.95rem;
        box-shadow: 0 2px 5px rgba(0,0,0,0.1);
    }
    .budget-table th {
        background-color: #0A174E;
        color: white;
        padding: 12px;
        text-align: left;
        font-weight: bold;
        border: 1px solid #1D2B5E;
    }
    .budget-table td {
        padding: 10px;
        border: 1px solid #dee2e6;
    }
    .budget-table tr:nth-child(even) {
        background-color: #f8f9fa;
    }
    .budget-table tr:hover {
        background-color: #f0f2f6;
    }
    .budget-total {
        background-color: #FFD700 !important;
        font-weight: bold;
    }
    .budget-highlight {
        background-color: #d4edda !important;
        font-weight: bold;
    }
    .cargo-info-table {
        width: 100%;
        border-collapse: collapse;
        margin-bottom: 1rem;
        background-color: #f0f2f6;
    }
    .cargo-info-table th {
        background-color: #1D2B5E;
        color: white;
        padding: 0.5rem;
        text-align: center;
        border: 1px solid #1D2B5E;
    }
    .cargo-info-table td {
        padding: 0.5rem;
        text-align: center;
        border: 1px solid #dee2e6;
        font-weight: bold;
    }
    .shipping-option {
        background-color: #e7f3ff;
        padding: 1rem;
        border-radius: 5px;
        margin-bottom: 0.5rem;
        border-left: 3px solid #0066cc;
    }
    .best-option {
        background-color: #d4edda;
        border-left: 3px solid #28a745;
        font-weight: bold;
    }
    .container-table {
        background-color: #f8f9fa;
        padding: 0.5rem;
        border-radius: 5px;
        margin-bottom: 0.5rem;
        font-size: 0.9rem;
    }
    .container-table table {
        width: 100%;
        border-collapse: collapse;
    }
    .container-table th {
        background-color: #0A174E;
        color: white;
        padding: 0.3rem;
        text-align: center;
    }
    .container-table td {
        padding: 0.3rem;
        text-align: center;
        border-bottom: 1px solid #dee2e6;
    }
</style>
""", unsafe_allow_html=True)

# -------------------- 辅助函数 --------------------
def get_beijing_time():
    utc_now = datetime.utcnow().replace(tzinfo=timezone.utc)
    beijing_now = utc_now.astimezone(timezone(timedelta(hours=8)))
    return beijing_now

def format_beijing_time(format_str='%Y-%m-%d %H:%M:%S'):
    return get_beijing_time().strftime(format_str)

# -------------------- 检查Excel文件 --------------------
def check_excel_file():
    """检查Excel文件是否存在"""
    file_exists = os.path.exists(EXCEL_FILE)
    return file_exists, EXCEL_FILE

# -------------------- 从Excel加载数据 --------------------
def load_data_from_excel():
    """从Excel文件加载所有数据"""
    data = {
        "ports": {},
        "rates": DEFAULT_RATES.copy(),
        "hs_info": {},
        "customer": None,
        "product": None
    }
    
    if not os.path.exists(EXCEL_FILE):
        return data
    
    try:
        excel_file = pd.ExcelFile(EXCEL_FILE)
        
        # 读取港口信息表
        if SHEET_PORTS in excel_file.sheet_names:
            df = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_PORTS)
            if not df.empty:
                for _, row in df.iterrows():
                    if len(row) >= 2:
                        country = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
                        port = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ""
                        if country and port and country != 'nan' and port != 'nan':
                            data["ports"][country] = port
        
        # 读取汇率表
        if SHEET_RATES in excel_file.sheet_names:
            df = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_RATES)
            if not df.empty:
                for _, row in df.iterrows():
                    if len(row) >= 3:
                        currency = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
                        try:
                            rate = float(row.iloc[2]) if pd.notna(row.iloc[2]) else 0
                            if currency and currency != 'nan' and rate > 0:
                                data["rates"][currency] = rate
                        except:
                            pass
        
        # 读取HS信息表
        if SHEET_HS in excel_file.sheet_names:
            df = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_HS)
            if not df.empty:
                for _, row in df.iterrows():
                    if len(row) >= 2:
                        hs_code = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
                        if hs_code and hs_code != 'nan':
                            data["hs_info"][hs_code] = {
                                "description": str(row.iloc[1]) if len(row) > 1 and pd.notna(row.iloc[1]) else "",
                                "tax_rate": float(row.iloc[2]) if len(row) > 2 and pd.notna(row.iloc[2]) else 0
                            }
        
        # 读取客户信息表
        if SHEET_CUSTOMERS in excel_file.sheet_names:
            df = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_CUSTOMERS)
            if not df.empty:
                latest = df.iloc[-1]
                data["customer"] = {
                    "customer_name": str(latest.iloc[0]) if len(df.columns) > 0 and pd.notna(latest.iloc[0]) else "",
                    "customer_country": str(latest.iloc[2]) if len(df.columns) > 2 and pd.notna(latest.iloc[2]) else "",
                    "customer_email": str(latest.iloc[3]) if len(df.columns) > 3 and pd.notna(latest.iloc[3]) else ""
                }
        
        # 读取商品信息表
        if SHEET_PRODUCTS in excel_file.sheet_names:
            df = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_PRODUCTS)
            if not df.empty:
                latest = df.iloc[-1]
                data["product"] = {
                    "product_code": str(latest.iloc[0]) if len(df.columns) > 0 and pd.notna(latest.iloc[0]) else "N003",
                    "product_name": str(latest.iloc[2]) if len(df.columns) > 2 and pd.notna(latest.iloc[2]) else "蓝宝石",
                    "hs_code": str(latest.iloc[6]) if len(df.columns) > 6 and pd.notna(latest.iloc[6]) else "7103910000",
                    "quantity": float(latest.iloc[8]) if len(df.columns) > 8 and pd.notna(latest.iloc[8]) else 0,
                    "price_per_ct": float(latest.iloc[9]) if len(df.columns) > 9 and pd.notna(latest.iloc[9]) else 0,
                    "unit_conversion": str(latest.iloc[11]) if len(df.columns) > 11 and pd.notna(latest.iloc[11]) else "1000CT/CARTON",
                    "gross_weight": float(latest.iloc[12]) if len(df.columns) > 12 and pd.notna(latest.iloc[12]) else 0.70,
                    "volume_per_pack": float(latest.iloc[14]) if len(df.columns) > 14 and pd.notna(latest.iloc[14]) else 0.0400
                }
                
    except Exception as e:
        st.error(f"读取Excel文件时出错: {e}")
    
    return data

# -------------------- 计算运输方案 --------------------
def calculate_shipping_options(total_volume, total_weight, freight_rates, lcl_rate_cbm, lcl_rate_kg):
    options = []
    
    # LCL散货
    lcl_volume_cost = total_volume * lcl_rate_cbm
    lcl_weight_cost = total_weight * lcl_rate_kg / 1000
    lcl_cost = max(lcl_volume_cost, lcl_weight_cost)
    
    if lcl_volume_cost > lcl_weight_cost:
        calculation = f"{total_volume:.2f} CBM × ${lcl_rate_cbm:.2f}/CBM = ${lcl_volume_cost:,.2f}"
    else:
        calculation = f"{total_weight:.2f} KG ÷ 1000 × ${lcl_rate_kg:.2f}/吨 = ${lcl_weight_cost:,.2f}"
    
    options.append({
        "name": "LCL散货",
        "cost_usd": lcl_cost,
        "calculation": calculation
    })
    
    # 整箱运输
    for container in CONTAINER_TYPES:
        if container['name'] in freight_rates:
            num = max(math.ceil(total_volume / container['volume']), 
                     math.ceil(total_weight / container['weight']))
            if num <= 5:
                cost = num * freight_rates[container['name']]
                calculation = f"{num} × ${freight_rates[container['name']]:,.2f} = ${cost:,.2f}"
                options.append({
                    "name": f"{num}×{container['display']}",
                    "cost_usd": cost,
                    "calculation": calculation
                })
    
    options.sort(key=lambda x: x["cost_usd"])
    return options, 0 if options else None

# -------------------- 显示出口预算表 --------------------
def display_budget_table(budget, selected_term):
    """显示出口预算表"""
    
    # 计算6 = 1-2+3+4+5
    total = (budget.get('采购成本', 0) - budget.get('出口退税', 0) + 
             budget.get('国内费用合计', 0) + budget.get('银行费用合计', 0) + 
             budget.get('其他费用', 0))
    
    html = f"""
    <table class="budget-table">
        <tr>
            <th style="width: 30%">项目</th>
            <th style="width: 40%">明细</th>
            <th style="width: 30%">金额 (CNY)</th>
        </tr>
        <tr>
            <td><strong>1. 采购成本</strong></td>
            <td>含税收入价</td>
            <td style="text-align: right">{budget.get('采购成本', 0):,.2f}</td>
        </tr>
        <tr>
            <td><strong>2. 退税收入</strong></td>
            <td>退税额</td>
            <td style="text-align: right">{budget.get('出口退税', 0):,.2f}</td>
        </tr>
        <tr>
            <td rowspan="3"><strong>3. 国内费用</strong></td>
            <td>出口国内运费</td>
            <td style="text-align: right">{budget.get('内陆运费', 6348.89):,.2f}</td>
        </tr>
        <tr>
            <td>国际运费</td>
            <td style="text-align: right">{budget.get('海运费', 13.85):,.2f}</td>
        </tr>
        <tr>
            <td>合计</td>
            <td style="text-align: right">{budget.get('国内费用合计', 0):,.2f}</td>
        </tr>
        <tr>
            <td rowspan="2"><strong>4. 银行费用</strong></td>
            <td>信用证费用</td>
            <td style="text-align: right">{budget.get('信用证费', 969.40):,.2f}</td>
        </tr>
        <tr>
            <td>合计</td>
            <td style="text-align: right">{budget.get('银行费用合计', 969.40):,.2f}</td>
        </tr>
        <tr>
            <td><strong>5. 其他费用</strong></td>
            <td></td>
            <td style="text-align: right">{budget.get('其他费用', 0):,.2f}</td>
        </tr>
        <tr class="budget-total">
            <td><strong>6 = 1-2+3+4+5</strong></td>
            <td>贸易术语</td>
            <td style="text-align: right"><strong>{selected_term}</strong></td>
        </tr>
        <tr class="budget-highlight">
            <td><strong>对外报价</strong></td>
            <td>{budget.get('对外报价', 0):,.2f} {budget.get('currency', 'USD')}</td>
            <td style="text-align: right">¥{budget.get('对外报价CNY', 0):,.2f}</td>
        </tr>
        <tr>
            <td><strong>预期盈亏额</strong></td>
            <td></td>
            <td style="text-align: right">¥{budget.get('预期盈亏额', 0):,.2f}</td>
        </tr>
        <tr>
            <td><strong>预期盈亏率</strong></td>
            <td></td>
            <td style="text-align: right">{budget.get('预期盈亏率', 0):.2f}%</td>
        </tr>
    </table>
    """
    
    return html

# -------------------- Session State 初始化 --------------------
if 'selected_currency' not in st.session_state:
    st.session_state.selected_currency = "USD"
if 'customer_data' not in st.session_state:
    st.session_state.customer_data = {}
if 'product_data' not in st.session_state:
    st.session_state.product_data = {}
if 'customer_fetched' not in st.session_state:
    st.session_state.customer_fetched = False
if 'product_fetched' not in st.session_state:
    st.session_state.product_fetched = False
if 'inland_freight' not in st.session_state:
    st.session_state.inland_freight = 6348.89
if 'lc_fee' not in st.session_state:
    st.session_state.lc_fee = 969.40
if 'freight_rates' not in st.session_state:
    st.session_state.freight_rates = {
        "20'GP": 1200.0, "40'GP": 1800.0, "40'HC": 2000.0,
        "20'RF": 2500.0, "40'RF": 3500.0, "40'RH": 3800.0
    }
if 'lcl_rate_cbm' not in st.session_state:
    st.session_state.lcl_rate_cbm = 50.0
if 'lcl_rate_kg' not in st.session_state:
    st.session_state.lcl_rate_kg = 2000.0
if 'budget' not in st.session_state:
    st.session_state.budget = None

# 加载Excel数据
file_exists, file_path = check_excel_file()
if file_exists:
    excel_data = load_data_from_excel()
    st.session_state.ports = excel_data.get("ports", {})
    st.session_state.exchange_rates = excel_data.get("rates", DEFAULT_RATES)
    st.session_state.hs_info = excel_data.get("hs_info", {})
    
    if excel_data.get("customer"):
        st.session_state.customer_data = excel_data["customer"]
        st.session_state.customer_fetched = True
    
    if excel_data.get("product"):
        st.session_state.product_data = excel_data["product"]
        st.session_state.product_fetched = True
else:
    st.session_state.ports = {}
    st.session_state.exchange_rates = DEFAULT_RATES
    st.session_state.hs_info = {}

# ==================== 页面内容开始 ====================

# 顶部标题
st.markdown("""
<div class="main-header">
    <h1>💰 AI价到 - 小微外贸智能折扣助手</h1>
    <div class="subtitle">智能报价 · 精准计算 · 一键成交</div>
</div>
""", unsafe_allow_html=True)

# 显示文件状态
file_exists, file_path = check_excel_file()
if file_exists:
    st.success(f"✅ Excel文件已找到: {file_path}")
else:
    st.error(f"⚠️ Excel文件不存在: {file_path}")
    st.info("请确认文件路径是否正确，或使用默认数据")

st.markdown("---")

# -------------------- 侧边栏 --------------------
with st.sidebar:
    st.markdown('<p class="sidebar-header">💱 汇率</p>', unsafe_allow_html=True)
    
    if file_exists:
        st.markdown("✅ <span class='status-badge status-success'>Excel数据已连接</span>", unsafe_allow_html=True)
    else:
        st.markdown("⚠️ <span class='status-badge status-warning'>使用默认汇率</span>", unsafe_allow_html=True)
    
    available_currencies = list(st.session_state.exchange_rates.keys())
    target_currency = st.selectbox("报价货币", available_currencies, 
                                  index=available_currencies.index("USD") if "USD" in available_currencies else 0)
    st.session_state.selected_currency = target_currency
    st.metric(f"1 {target_currency} = ", f"{st.session_state.exchange_rates[target_currency]:.4f} CNY")
    
    st.markdown("---")
    
    st.markdown('<p class="sidebar-header">🚢 物流信息</p>', unsafe_allow_html=True)
    
    # 集装箱参数表
    st.markdown("""
    <div class="container-table">
        <table>
            <tr><th>箱型</th><th>代码</th><th>体积</th><th>重量</th></tr>
    """, unsafe_allow_html=True)
    for name, spec in CONTAINER_SPECS.items():
        st.markdown(f"<tr><td>{name}</td><td>{spec['code']}</td><td>{spec['volume']}</td><td>{spec['weight']}</td></tr>", unsafe_allow_html=True)
    st.markdown("</table></div>", unsafe_allow_html=True)
    
    # 运费设置
    st.markdown("**集装箱运费 (USD)**")
    freight_rates = st.session_state.freight_rates.copy()
    for container in CONTAINER_TYPES:
        freight_rates[container['name']] = st.number_input(
            container['display'], 
            value=float(freight_rates.get(container['name'], 1200.0)), 
            step=50.0, 
            key=f"freight_{container['name']}"
        )
    
    st.markdown("**LCL散货费率**")
    lcl_rate_cbm = st.number_input("LCL(M) (USD/CBM)", value=st.session_state.lcl_rate_cbm, step=5.0)
    lcl_rate_kg = st.number_input("LCL(W) (USD/吨)", value=st.session_state.lcl_rate_kg, step=100.0)
    
    if st.button("更新运费设置"):
        st.session_state.freight_rates = freight_rates
        st.session_state.lcl_rate_cbm = lcl_rate_cbm
        st.session_state.lcl_rate_kg = lcl_rate_kg
        st.success("运费设置已更新")
        st.rerun()

# -------------------- 主区域 --------------------
col_left, col_right = st.columns(2)

with col_left:
    st.markdown("### 🏢 本公司信息")
    company_name = st.text_input("公司名称", "ABC International Trading CO. Ltd")

with col_right:
    st.markdown("### 👥 客户信息")
    if st.session_state.customer_fetched:
        st.success("✅ 已加载客户数据")
    
    default_customer = st.session_state.customer_data
    customer = st.text_input("客户名称", value=default_customer.get("customer_name", "") if default_customer else "")
    
    port_map = st.session_state.ports if st.session_state.ports else {"China": "Shanghai", "USA": "Los Angeles"}
    countries = list(port_map.keys())
    country = st.selectbox("目的国家", countries)
    st.text_input("目的港口", value=port_map.get(country, ""), disabled=True)

st.markdown("---")

# -------------------- 商品信息 --------------------
st.markdown("### 💎 商品信息")
if st.session_state.product_fetched:
    st.success("✅ 已加载商品数据")

default_product = st.session_state.product_data if st.session_state.product_data else {}

col1, col2, col3, col4 = st.columns(4)
with col1:
    product_code = st.text_input("商品编号", value=default_product.get("product_code", "N003") if default_product else "N003")
with col2:
    product_name = st.text_input("商品名称", value=default_product.get("product_name", "蓝宝石") if default_product else "蓝宝石")
with col3:
    hs_code = st.text_input("HS编码", value=default_product.get("hs_code", "7103910000") if default_product else "7103910000")
with col4:
    quantity = st.number_input("数量 (克拉)", value=float(default_product.get("quantity", 0)) if default_product else 0.0, step=100.0)

col5, col6, col7, col8 = st.columns(4)
with col5:
    price_per_ct = st.number_input("采购单价 (￥/克拉)", value=float(default_product.get("price_per_ct", 0.0)) if default_product else 0.0, step=1.0)
with col6:
    unit_conversion = st.text_input("单位换算", value=default_product.get("unit_conversion", "1000CT/CARTON") if default_product else "1000CT/CARTON")
with col7:
    gross_weight = st.number_input("毛重 (KGS/箱)", value=float(default_product.get("gross_weight", 0.70)) if default_product else 0.70, format="%.2f")
with col8:
    volume_per_pack = st.number_input("体积 (CBM/箱)", value=float(default_product.get("volume_per_pack", 0.0400)) if default_product else 0.0400, format="%.4f")

st.markdown("---")

# -------------------- 贸易术语 --------------------
st.markdown("### 📋 贸易术语")
term_options = [term["name"] for term in INCOTERMS_2020]
selected_term = st.selectbox("贸易术语", term_options, index=2)
profit_margin = st.slider("利润率 (%)", 5, 100, 20)
tax_rate = st.slider("出口退税率 (%)", 0, 17, 13)

st.markdown("---")

# -------------------- 开始报价 --------------------
if st.button("开始报价", type="primary", use_container_width=True):
    if quantity > 0 and price_per_ct > 0:
        # 计算总成本
        total_cost_cny = quantity * price_per_ct
        
        # 计算箱数
        try:
            ct_per_carton = float(unit_conversion.split('/')[0].replace('CT', '').strip())
            total_packages = math.ceil(quantity / ct_per_carton)
        except:
            total_packages = math.ceil(quantity / 1000)
        
        total_volume = total_packages * volume_per_pack
        total_weight = total_packages * gross_weight
        
        # 计算运费
        shipping_options, best_index = calculate_shipping_options(
            total_volume, total_weight,
            st.session_state.freight_rates,
            st.session_state.lcl_rate_cbm,
            st.session_state.lcl_rate_kg
        )
        
        best_option = shipping_options[best_index]
        freight_usd = best_option["cost_usd"]
        freight_cny = freight_usd * st.session_state.exchange_rates[st.session_state.selected_currency]
        
        # 计算各项费用
        tax_refund = total_cost_cny * (tax_rate / 100)
        domestic_fee_total = st.session_state.inland_freight + freight_cny
        bank_fee_total = st.session_state.lc_fee
        
        # 对外报价
        quoted_price_target = (total_cost_cny / st.session_state.exchange_rates[st.session_state.selected_currency]) * (1 + profit_margin/100)
        quoted_price_cny = quoted_price_target * st.session_state.exchange_rates[st.session_state.selected_currency]
        
        # 盈亏计算
        total_cost_with_freight = total_cost_cny - tax_refund + domestic_fee_total + bank_fee_total
        expected_profit = quoted_price_cny - total_cost_with_freight
        expected_profit_rate = (expected_profit / total_cost_with_freight) * 100 if total_cost_with_freight != 0 else 0
        
        # 保存预算数据
        st.session_state.budget = {
            "采购成本": total_cost_cny,
            "出口退税": tax_refund,
            "内陆运费": st.session_state.inland_freight,
            "海运费": freight_cny,
            "国内费用合计": domestic_fee_total,
            "信用证费": st.session_state.lc_fee,
            "银行费用合计": bank_fee_total,
            "其他费用": 0,
            "对外报价": quoted_price_target,
            "对外报价CNY": quoted_price_cny,
            "预期盈亏额": expected_profit,
            "预期盈亏率": expected_profit_rate,
            "currency": st.session_state.selected_currency,
            "总箱数": total_packages,
            "总体积": total_volume,
            "总毛重": total_weight,
            "shipping_options": shipping_options,
            "best_shipping_index": best_index,
            "selected_shipping": best_option["name"],
            "shipping_calculation": best_option.get("calculation", ""),
            "freight_usd": freight_usd
        }
        
        st.success("计算完成！")
        st.balloons()
    else:
        st.warning("请填写商品数量和单价")

# -------------------- 显示结果 --------------------
if st.session_state.budget:
    b = st.session_state.budget
    
    # 整批货物信息
    st.markdown("#### 📦 整批货物信息")
    st.markdown(f"""
    <table class="cargo-info-table">
        <tr><th>总箱数</th><th>总体积 (CBM)</th><th>总毛重 (KG)</th></tr>
        <tr><td>{b.get('总箱数', 0):,} 箱</td><td>{b.get('总体积', 0):.2f}</td><td>{b.get('总毛重', 0):.2f}</td></tr>
    </table>
    """, unsafe_allow_html=True)
    
    # 最终运费方案
    st.markdown("#### 🚢 最终运费方案")
    st.markdown(f"""
    <div class="shipping-option best-option">
        <strong>✅ 选中方案: {b.get('selected_shipping', '未知')}</strong><br>
        <strong>计算原理:</strong> {b.get('shipping_calculation', '')}<br>
        <strong>运费:</strong> ${b.get('freight_usd', 0):,.2f} | ￥{b.get('海运费', 0):,.2f}
    </div>
    """, unsafe_allow_html=True)
    
    # 所有方案对比
    with st.expander("查看所有运输方案对比"):
        for i, option in enumerate(b.get("shipping_options", [])):
            if i == b.get("best_shipping_index"):
                st.markdown(f"**✅ 最佳方案 #{i+1}: {option['name']}** - {option.get('calculation', '')} - ${option['cost_usd']:,.2f}")
            else:
                st.markdown(f"**方案 #{i+1}: {option['name']}** - {option.get('calculation', '')} - ${option['cost_usd']:,.2f}")
    
    # 出口预算表
    st.markdown("### 📊 出口预算表")
    budget_html = display_budget_table(b, selected_term)
    st.markdown(budget_html, unsafe_allow_html=True)

st.markdown("---")
st.markdown("© 2026 AI价到团队 | Excel数据源: C:\\Basic Information\\Data.xlsx")



























