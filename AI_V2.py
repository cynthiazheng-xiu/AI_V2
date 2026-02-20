# app.py - AI价到 - 小微外贸智能报价助手 (完整版)

import streamlit as st
import pandas as pd
import math
import subprocess
import os
import time
import json
from datetime import datetime, timedelta, timezone

# -------------------- 页面配置 - 必须放在最前面 --------------------
st.set_page_config(
    page_title="AI价到 - 小微外贸智能出口报价助手",
    page_icon="💰",
    layout="wide",
    initial_sidebar_state="expanded"
)

# -------------------- 常量定义 --------------------
# 使用原始字符串避免转义问题
BASE_DATA_PATH = r"C:\Basic Information"

# 使用os.path.join构建完整路径，确保路径分隔符正确
EXCEL_FILE_XLSX = os.path.join(BASE_DATA_PATH, "Data.xlsx")
EXCEL_FILE_XLS = os.path.join(BASE_DATA_PATH, "Data.xls")

# 检查哪个文件存在
if os.path.exists(EXCEL_FILE_XLSX):
    EXCEL_FILE = EXCEL_FILE_XLSX
elif os.path.exists(EXCEL_FILE_XLS):
    EXCEL_FILE = EXCEL_FILE_XLS
else:
    EXCEL_FILE = EXCEL_FILE_XLSX  # 默认使用xlsx

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

# 集装箱参数表（用于后台计算）
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
    {"code": "EXW", "name": "EXW (工厂交货)", "description": "卖方在其所在地交货"},
    {"code": "FOB", "name": "FOB (船上交货)", "description": "卖方在指定装运港将货物装到买方指定的船上"},
    {"code": "CIF", "name": "CIF (成本、保险费加运费)", "description": "卖方支付运费和保险费至目的港"}
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
    .sidebar-header {
        color: #0A174E;
        font-weight: bold;
        font-size: 1.1rem;
        margin-top: 1rem;
        margin-bottom: 0.5rem;
    }
    .status-badge {
        padding: 0.25rem 0.75rem;
        border-radius: 20px;
        font-size: 0.8rem;
        font-weight: bold;
        display: inline-block;
    }
    .status-success {
        background-color: #d4edda;
        color: #155724;
    }
    .status-warning {
        background-color: #fff3cd;
        color: #856404;
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
    file_info = {
        "exists": False,
        "path": EXCEL_FILE,
        "xlsx_exists": os.path.exists(EXCEL_FILE_XLSX),
        "xls_exists": os.path.exists(EXCEL_FILE_XLS),
        "dir_exists": os.path.exists(BASE_DATA_PATH),
        "files_in_dir": []
    }
    
    if os.path.exists(BASE_DATA_PATH):
        try:
            file_info["files_in_dir"] = os.listdir(BASE_DATA_PATH)
        except:
            pass
    
    if os.path.exists(EXCEL_FILE):
        file_info["exists"] = True
        file_info["path"] = EXCEL_FILE
        file_info["size"] = os.path.getsize(EXCEL_FILE)
        mod_time = os.path.getmtime(EXCEL_FILE)
        file_info["modified"] = datetime.fromtimestamp(mod_time).strftime('%Y-%m-%d %H:%M:%S')
        
        try:
            excel_file = pd.ExcelFile(EXCEL_FILE)
            file_info["sheets"] = excel_file.sheet_names
        except:
            file_info["sheets"] = ["无法读取工作表"]
    
    return file_info

# -------------------- 从Excel查找商品信息 --------------------
def search_product_from_excel(search_term, search_by="code"):
    """根据商品编号或英文名称从Excel查找商品信息"""
    
    if not os.path.exists(EXCEL_FILE):
        return None
    
    try:
        df = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_PRODUCTS)
        if df.empty:
            return None
        
        result = None
        if search_by == "code":
            # 按商品编号查找（第1列）
            for _, row in df.iterrows():
                cell_value = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
                if cell_value == str(search_term).strip():
                    result = row
                    break
        elif search_by == "name":
            # 按英文名称查找（第4列）
            for _, row in df.iterrows():
                cell_value = str(row.iloc[3]).strip().lower() if pd.notna(row.iloc[3]) else ""
                if cell_value == str(search_term).strip().lower():
                    result = row
                    break
        
        if result is not None:
            return {
                "product_code": str(result.iloc[0]) if len(df.columns) > 0 and pd.notna(result.iloc[0]) else "",
                "goods_type": str(result.iloc[1]) if len(df.columns) > 1 and pd.notna(result.iloc[1]) else "",
                "product_name": str(result.iloc[2]) if len(df.columns) > 2 and pd.notna(result.iloc[2]) else "",
                "product_name_en": str(result.iloc[3]) if len(df.columns) > 3 and pd.notna(result.iloc[3]) else "",
                "specification_cn": str(result.iloc[4]) if len(df.columns) > 4 and pd.notna(result.iloc[4]) else "",
                "specification_en": str(result.iloc[5]) if len(df.columns) > 5 and pd.notna(result.iloc[5]) else "",
                "hs_code": str(result.iloc[6]) if len(df.columns) > 6 and pd.notna(result.iloc[6]) else "",
                "sales_unit": str(result.iloc[7]) if len(df.columns) > 7 and pd.notna(result.iloc[7]) else "",
                "quantity": float(result.iloc[8]) if len(df.columns) > 8 and pd.notna(result.iloc[8]) else 0,
                "price_per_ct": float(result.iloc[9]) if len(df.columns) > 9 and pd.notna(result.iloc[9]) else 0,
                "package_unit": str(result.iloc[10]) if len(df.columns) > 10 and pd.notna(result.iloc[10]) else "",
                "unit_conversion": str(result.iloc[11]) if len(df.columns) > 11 and pd.notna(result.iloc[11]) else "",
                "gross_weight": float(result.iloc[12]) if len(df.columns) > 12 and pd.notna(result.iloc[12]) else 0,
                "net_weight": float(result.iloc[13]) if len(df.columns) > 13 and pd.notna(result.iloc[13]) else 0,
                "volume_per_pack": float(result.iloc[14]) if len(df.columns) > 14 and pd.notna(result.iloc[14]) else 0,
                "legal_unit": str(result.iloc[15]) if len(df.columns) > 15 and pd.notna(result.iloc[15]) else "",
                "customs_supervision": str(result.iloc[16]) if len(df.columns) > 16 and pd.notna(result.iloc[16]) else "",
                "inspection_category": str(result.iloc[17]) if len(df.columns) > 17 and pd.notna(result.iloc[17]) else "",
                "transport_notes": str(result.iloc[18]) if len(df.columns) > 18 and pd.notna(result.iloc[18]) else "",
                "description": str(result.iloc[19]) if len(df.columns) > 19 and pd.notna(result.iloc[19]) else ""
            }
        return None
    except Exception as e:
        st.error(f"查找商品信息时出错: {e}")
        return None

# -------------------- 运行PAD查找商品函数 --------------------
def run_pad_search_product(search_term, search_by):
    """模拟运行PAD查找商品"""
    with st.spinner(f"正在从Excel查找商品: {search_term}..."):
        time.sleep(1)
        product_info = search_product_from_excel(search_term, search_by)
        if product_info:
            return {"success": True, "data": product_info, "message": "找到商品信息"}
        else:
            return {"success": False, "message": "未找到匹配的商品信息"}

# -------------------- 运行PAD流程函数 --------------------
def run_pad_flow(flow_name):
    """调用Power Automate Desktop运行指定流程"""
    try:
        pad_path = "C:\\Program Files (x86)\\Power Automate Desktop\\PAD.Console.exe"
        if os.path.exists(pad_path):
            return {"success": True, "message": f"已启动PAD流程: {flow_name}"}
        else:
            return {"success": True, "message": f"模拟运行成功"}
    except Exception as e:
        return {"success": False, "message": str(e)}

# -------------------- 计算运输方案函数 --------------------
def calculate_shipping_options(total_volume, total_weight, freight_rates, lcl_rate_cbm, lcl_rate_kg):
    """计算运输方案"""
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
        "calculation": calculation,
        "description": "散货拼箱"
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
                    "calculation": calculation,
                    "description": f"{num}个{container['display']}"
                })
    
    options.sort(key=lambda x: x["cost_usd"])
    return options, 0 if options else None

# -------------------- 显示出口预算表函数 --------------------
def display_budget_table(budget, selected_term):
    """显示出口预算表"""
    
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
            <td style="text-align: right">{budget.get('海运费', 0):,.2f}</td>
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

# -------------------- 获取国家港口映射 --------------------
def get_country_port_map():
    """获取国家港口映射"""
    return {
        "China": "Shanghai", "USA": "Los Angeles", "Germany": "Hamburg",
        "UK": "Felixstowe", "Japan": "Tokyo", "Australia": "Sydney"
    }

# -------------------- Session State 初始化 --------------------
if 'selected_currency' not in st.session_state:
    st.session_state.selected_currency = "USD"
if 'exchange_rates' not in st.session_state:
    st.session_state.exchange_rates = DEFAULT_RATES
if 'product_data' not in st.session_state:
    st.session_state.product_data = {}
if 'product_fetched' not in st.session_state:
    st.session_state.product_fetched = False
if 'customer_data' not in st.session_state:
    st.session_state.customer_data = {}
if 'customer_fetched' not in st.session_state:
    st.session_state.customer_fetched = False
if 'inland_freight' not in st.session_state:
    st.session_state.inland_freight = 6348.89
if 'lc_fee' not in st.session_state:
    st.session_state.lc_fee = 969.40
if 'freight_rates' not in st.session_state:
    st.session_state.freight_rates = {
        "20'GP": 1200.0, "40'GP": 1800.0, "40'HC": 2000.0,
        "20'RF": 2500.0, "40'RF": 3500.0, "40'RH": 3800.0
    }
if 'lcl_rate_cbm_normal' not in st.session_state:
    st.session_state.lcl_rate_cbm_normal = 50.0
if 'lcl_rate_kg_normal' not in st.session_state:
    st.session_state.lcl_rate_kg_normal = 2000.0
if 'lcl_rate_cbm_frozen' not in st.session_state:
    st.session_state.lcl_rate_cbm_frozen = 75.0
if 'lcl_rate_kg_frozen' not in st.session_state:
    st.session_state.lcl_rate_kg_frozen = 3000.0
if 'budget' not in st.session_state:
    st.session_state.budget = None

# ==================== 页面内容开始 ====================

# -------------------- 顶部标题 --------------------
st.markdown("""
<div class="main-header">
    <h1>💰 AI价到 - 小微外贸智能出口报价助手</h1>
    <div class="subtitle">智能报价 · 精准计算 · 一键成交</div>
</div>
""", unsafe_allow_html=True)

# -------------------- 检查并显示文件状态 --------------------
file_info = check_excel_file()

# 显示详细的文件状态
st.markdown("### 📁 文件状态")
col1, col2 = st.columns(2)
with col1:
    st.write(f"基础路径: {BASE_DATA_PATH}")
    st.write(f"目录存在: {'✅' if file_info['dir_exists'] else '❌'}")
    if file_info['dir_exists'] and file_info['files_in_dir']:
        st.write(f"目录中的文件: {', '.join(file_info['files_in_dir'])}")

with col2:
    st.write(f"Data.xlsx 存在: {'✅' if file_info['xlsx_exists'] else '❌'}")
    st.write(f"Data.xls 存在: {'✅' if file_info['xls_exists'] else '❌'}")

if file_info["exists"]:
    st.success(f"✅ 找到Excel文件: {file_info['path']}")
    if "sheets" in file_info:
        st.info(f"📊 工作表: {', '.join(file_info['sheets'])}")
else:
    st.error(f"❌ Excel文件不存在: {EXCEL_FILE}")
    st.info("请确认:")
    st.info("1. 目录 C:\\Basic Information 是否存在")
    st.info("2. 文件名是 Data.xlsx 还是 Data.xls")
    st.info("3. 如果是.xls文件，请确保已安装openpyxl库: pip install openpyxl")

st.markdown("---")

# -------------------- 侧边栏 --------------------
with st.sidebar:
    st.markdown('<p class="sidebar-header">💱 汇率</p>', unsafe_allow_html=True)
    
    if file_info["exists"]:
        st.markdown("✅ <span class='status-badge status-success'>Excel数据可用</span>", unsafe_allow_html=True)
    else:
        st.markdown("⚠️ <span class='status-badge status-warning'>使用默认汇率</span>", unsafe_allow_html=True)
    
    available_currencies = list(st.session_state.exchange_rates.keys())
    target_currency = st.selectbox("报价货币", available_currencies, 
                                  index=available_currencies.index("USD") if "USD" in available_currencies else 0)
    st.session_state.selected_currency = target_currency
    st.metric(f"1 {target_currency} = ", f"{st.session_state.exchange_rates[target_currency]:.4f} CNY")
    
    st.markdown("---")
    
    st.markdown('<p class="sidebar-header">🚢 物流信息</p>', unsafe_allow_html=True)
    
    col_from, col_to = st.columns(2)
    with col_from:
        departure_port = st.text_input("起运港", value="Shanghai")
    with col_to:
        port_map = get_country_port_map()
        destination_port = st.text_input("目的港", value="Los Angeles")
    
    st.markdown("### 运费设置")
    
    # 普柜运费
    st.markdown("#### 普柜")
    col1, col2, col3 = st.columns([2, 1, 1])
    with col1:
        st.markdown("**类型**")
    with col2:
        st.markdown("**单价**")
    with col3:
        st.markdown("**单位**")
    
    # LCL(M) 普柜
    col1, col2, col3 = st.columns([2, 1, 1])
    with col1:
        st.markdown("LCL(M)")
    with col2:
        lcl_cbm_normal = st.number_input("", value=st.session_state.lcl_rate_cbm_normal, step=5.0, key="lcl_cbm_normal", label_visibility="collapsed")
    with col3:
        st.markdown("USD/CBM")
    
    # LCL(W) 普柜
    col1, col2, col3 = st.columns([2, 1, 1])
    with col1:
        st.markdown("LCL(W)")
    with col2:
        lcl_kg_normal = st.number_input("", value=st.session_state.lcl_rate_kg_normal, step=100.0, key="lcl_kg_normal", label_visibility="collapsed")
    with col3:
        st.markdown("USD/吨")
    
    # 20'普柜
    col1, col2, col3 = st.columns([2, 1, 1])
    with col1:
        st.markdown("20'")
    with col2:
        freight_20 = st.number_input("", value=st.session_state.freight_rates["20'GP"], step=50.0, key="freight_20", label_visibility="collapsed")
    with col3:
        st.markdown("USD")
    
    # 40'普柜
    col1, col2, col3 = st.columns([2, 1, 1])
    with col1:
        st.markdown("40'")
    with col2:
        freight_40 = st.number_input("", value=st.session_state.freight_rates["40'GP"], step=50.0, key="freight_40", label_visibility="collapsed")
    with col3:
        st.markdown("USD")
    
    # 40'高普柜
    col1, col2, col3 = st.columns([2, 1, 1])
    with col1:
        st.markdown("40'高")
    with col2:
        freight_40hc = st.number_input("", value=st.session_state.freight_rates["40'HC"], step=50.0, key="freight_40hc", label_visibility="collapsed")
    with col3:
        st.markdown("USD")
    
    st.markdown("---")
    
    # 冻柜运费
    st.markdown("#### 冻柜")
    
    # LCL(M) 冻柜
    col1, col2, col3 = st.columns([2, 1, 1])
    with col1:
        st.markdown("LCL(M)")
    with col2:
        lcl_cbm_frozen = st.number_input("", value=st.session_state.lcl_rate_cbm_frozen, step=5.0, key="lcl_cbm_frozen", label_visibility="collapsed")
    with col3:
        st.markdown("USD/CBM")
    
    # LCL(W) 冻柜
    col1, col2, col3 = st.columns([2, 1, 1])
    with col1:
        st.markdown("LCL(W)")
    with col2:
        lcl_kg_frozen = st.number_input("", value=st.session_state.lcl_rate_kg_frozen, step=100.0, key="lcl_kg_frozen", label_visibility="collapsed")
    with col3:
        st.markdown("USD/吨")
    
    # 20'冻柜
    col1, col2, col3 = st.columns([2, 1, 1])
    with col1:
        st.markdown("20'")
    with col2:
        freight_20rf = st.number_input("", value=st.session_state.freight_rates["20'RF"], step=50.0, key="freight_20rf", label_visibility="collapsed")
    with col3:
        st.markdown("USD")
    
    # 40'冻柜
    col1, col2, col3 = st.columns([2, 1, 1])
    with col1:
        st.markdown("40'")
    with col2:
        freight_40rf = st.number_input("", value=st.session_state.freight_rates["40'RF"], step=50.0, key="freight_40rf", label_visibility="collapsed")
    with col3:
        st.markdown("USD")
    
    # 40'高冻柜
    col1, col2, col3 = st.columns([2, 1, 1])
    with col1:
        st.markdown("40'高")
    with col2:
        freight_40rh = st.number_input("", value=st.session_state.freight_rates["40'RH"], step=50.0, key="freight_40rh", label_visibility="collapsed")
    with col3:
        st.markdown("USD")
    
    if st.button("更新运费设置", use_container_width=True):
        st.session_state.freight_rates["20'GP"] = freight_20
        st.session_state.freight_rates["40'GP"] = freight_40
        st.session_state.freight_rates["40'HC"] = freight_40hc
        st.session_state.freight_rates["20'RF"] = freight_20rf
        st.session_state.freight_rates["40'RF"] = freight_40rf
        st.session_state.freight_rates["40'RH"] = freight_40rh
        st.session_state.lcl_rate_cbm_normal = lcl_cbm_normal
        st.session_state.lcl_rate_kg_normal = lcl_kg_normal
        st.session_state.lcl_rate_cbm_frozen = lcl_cbm_frozen
        st.session_state.lcl_rate_kg_frozen = lcl_kg_frozen
        st.success("运费设置已更新")

# -------------------- 主区域 --------------------
col_left, col_right = st.columns(2)

with col_left:
    st.markdown("### 🏢 本公司信息")
    company_name = st.text_input("公司名称", "ABC International Trading CO. Ltd")

with col_right:
    st.markdown("### 👥 客户信息")
    customer_name = st.text_input("客户名称")
    country = st.selectbox("目的国家", ["China", "USA", "Germany", "UK", "Japan"])

st.markdown("---")

# -------------------- 商品信息 --------------------
st.markdown("### 💎 商品信息")

# 搜索区域
col_search1, col_search2, col_search3 = st.columns([3, 2, 1])
with col_search1:
    search_term = st.text_input("输入商品编号或英文名称", placeholder="例如: N003 或 Sapphires")
with col_search2:
    search_by = st.selectbox("搜索方式", ["商品编号", "英文名称"])
with col_search3:
    search_button = st.button("🔍 查找商品", type="primary", use_container_width=True)

if search_button and search_term:
    search_by_code = "code" if search_by == "商品编号" else "name"
    result = run_pad_search_product(search_term, search_by_code)
    if result["success"]:
        st.session_state.product_data = result["data"]
        st.session_state.product_fetched = True
        st.success(f"✅ {result['message']}")
        st.rerun()
    else:
        st.error(f"❌ {result['message']}")

if st.session_state.product_fetched:
    st.success(f"✅ 已加载商品: {st.session_state.product_data.get('product_name', '')}")

# 商品信息输入
default_product = st.session_state.product_data if st.session_state.product_data else {}

col1, col2, col3, col4 = st.columns(4)
with col1:
    product_code = st.text_input("商品编号", value=default_product.get("product_code", ""), 
                                disabled=bool(st.session_state.product_data))
with col2:
    product_name = st.text_input("商品名称", value=default_product.get("product_name", ""), 
                                disabled=bool(st.session_state.product_data))
with col3:
    product_name_en = st.text_input("英文名称", value=default_product.get("product_name_en", ""), 
                                   disabled=bool(st.session_state.product_data))
with col4:
    hs_code = st.text_input("HS编码", value=default_product.get("hs_code", ""), 
                           disabled=bool(st.session_state.product_data))

col5, col6, col7, col8 = st.columns(4)
with col5:
    quantity = st.number_input("数量 (克拉)", value=float(default_product.get("quantity", 0)), 
                              step=100.0, disabled=bool(st.session_state.product_data))
with col6:
    price_per_ct = st.number_input("采购单价 (￥/克拉)", value=float(default_product.get("price_per_ct", 0)), 
                                  step=1.0, disabled=bool(st.session_state.product_data))
with col7:
    unit_conversion = st.text_input("单位换算", value=default_product.get("unit_conversion", "1000CT/CARTON"), 
                                   disabled=bool(st.session_state.product_data))
with col8:
    gross_weight = st.number_input("毛重 (KGS/箱)", value=float(default_product.get("gross_weight", 0.70)), 
                                  format="%.2f", disabled=bool(st.session_state.product_data))

col9, col10, col11, col12 = st.columns(4)
with col9:
    volume_per_pack = st.number_input("体积 (CBM/箱)", value=float(default_product.get("volume_per_pack", 0.0400)), 
                                     format="%.4f", disabled=bool(st.session_state.product_data))

if st.session_state.product_data:
    if st.button("清除搜索结果"):
        st.session_state.product_data = {}
        st.session_state.product_fetched = False
        st.rerun()

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
        exchange_rate = st.session_state.exchange_rates[st.session_state.selected_currency]
        
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
            st.session_state.lcl_rate_cbm_normal,
            st.session_state.lcl_rate_kg_normal
        )
        
        best_option = shipping_options[best_index]
        freight_usd = best_option["cost_usd"]
        freight_cny = freight_usd * exchange_rate
        
        # 计算各项费用
        tax_refund = total_cost_cny * (tax_rate / 100)
        domestic_fee_total = st.session_state.inland_freight + freight_cny
        bank_fee_total = st.session_state.lc_fee
        
        # 对外报价
        quoted_price_target = (total_cost_cny / exchange_rate) * (1 + profit_margin/100)
        quoted_price_cny = quoted_price_target * exchange_rate
        
        # 盈亏计算
        total_cost_with_freight = total_cost_cny - tax_refund + domestic_fee_total + bank_fee_total
        expected_profit = quoted_price_cny - total_cost_with_freight
        expected_profit_rate = (expected_profit / total_cost_with_freight) * 100 if total_cost_with_freight != 0 else 0
        
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
        <tr>
            <th>总箱数</th>
            <th>总体积 (CBM)</th>
            <th>总毛重 (KG)</th>
        </tr>
        <tr>
            <td>{b.get('总箱数', 0):,} 箱</td>
            <td>{b.get('总体积', 0):.2f}</td>
            <td>{b.get('总毛重', 0):.2f}</td>
        </tr>
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

































