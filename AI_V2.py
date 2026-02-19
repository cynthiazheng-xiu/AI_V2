# app.py - AI价到 - 小微外贸智能报价助手 (完整版)

import streamlit as st
import pandas as pd
import math
import subprocess
import os
import time
from datetime import datetime, timedelta, timezone

# -------------------- 页面配置 - 必须放在最前面 --------------------
st.set_page_config(
    page_title="AI价到 - 小微外贸智能折扣助手",
    page_icon="💰",
    layout="wide",
    initial_sidebar_state="expanded"
)

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
    .section-header {
        background-color: #f0f2f6;
        padding: 0.5rem 1rem;
        border-radius: 8px;
        margin-bottom: 1rem;
        font-weight: bold;
        color: #0A174E;
        font-size: 1.2rem;
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
    .term-card {
        background-color: #f8f9fa;
        border-left: 4px solid #0A174E;
        padding: 0.5rem 1rem;
        border-radius: 4px;
        margin-bottom: 0.5rem;
    }
    .term-title {
        font-weight: bold;
        color: #0A174E;
    }
    .shipping-option {
        background-color: #e7f3ff;
        padding: 0.5rem;
        border-radius: 5px;
        margin-bottom: 0.3rem;
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
    .shipping-comparison {
        background-color: #fff3cd;
        padding: 0.5rem;
        border-radius: 5px;
        margin-bottom: 0.5rem;
        border-left: 3px solid #856404;
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

# 国家港口映射
country_port_map = {
    "Chile": "San Antonio", "USA": "Los Angeles", "Germany": "Hamburg",
    "Philippines": "Manila", "China": "Shanghai", "Japan": "Tokyo",
    "UK": "Felixstowe", "France": "Le Havre", "Italy": "Genoa",
    "Australia": "Sydney", "Brazil": "Santos", "India": "Mumbai"
}

# 默认汇率
DEFAULT_RATES = {
    "USD": 6.9257, "EUR": 8.1863, "GBP": 9.3729, "JPY": 0.044775,
    "HKD": 0.8858, "AUD": 4.9092, "CAD": 5.0734, "CHF": 8.9762, "SGD": 5.4721
}

# 集装箱参数表
CONTAINER_SPECS = {
    "20'": {"volume": 33, "weight": 25000},
    "40'": {"volume": 67, "weight": 29000},
    "40'高": {"volume": 76, "weight": 29000}
}

# 2020版国际贸易术语完整列表（简化版）
INCOTERMS_2020 = [
    {
        "code": "EXW",
        "name": "EXW (工厂交货)",
        "full_name": "Ex Works",
        "description": "卖方在其所在地或其他指定地点将货物交给买方处置时即完成交货。",
        "responsibility_seller": "在指定地点提供货物",
        "responsibility_buyer": "所有运输、保险、进出口清关",
        "risk_transfer": "卖方将货物交给买方处置时",
        "transport": "任何运输方式"
    },
    {
        "code": "FOB",
        "name": "FOB (船上交货)",
        "full_name": "Free On Board",
        "description": "卖方在指定装运港将货物装到买方指定的船上即完成交货。",
        "responsibility_seller": "出口清关、将货物装上船",
        "responsibility_buyer": "主运输、保险、进口清关",
        "risk_transfer": "货物装上船时",
        "transport": "海运和内河水运"
    },
    {
        "code": "CIF",
        "name": "CIF (成本、保险费加运费)",
        "full_name": "Cost, Insurance and Freight",
        "description": "卖方支付将货物运至指定目的港的运费和保险费。",
        "responsibility_seller": "出口清关、将货物装上船、支付运费和保险费",
        "responsibility_buyer": "进口清关、目的港卸货费",
        "risk_transfer": "货物装上船时",
        "transport": "海运和内河水运"
    }
]

# -------------------- 数据加载函数 --------------------
def load_rates_from_excel():
    """从PAD生成的Excel文件加载汇率数据"""
    excel_path = "C:\\ExchangeRates\\rates.xlsx"
    rates = DEFAULT_RATES.copy()
    rate_info = {
        "rates": rates,
        "publish_time": "未知",
        "fetch_time": "未知",
        "file_exists": False,
        "file_time": None
    }
    if os.path.exists(excel_path):
        try:
            df = pd.read_excel(excel_path)
            mod_time = os.path.getmtime(excel_path)
            mod_time_beijing = datetime.fromtimestamp(mod_time) + timedelta(hours=8)
            rate_info["file_time"] = mod_time_beijing.strftime('%Y-%m-%d %H:%M:%S')
            rate_info["file_exists"] = True
            for index, row in df.iterrows():
                currency_code = str(row.iloc[0]).strip()
                rate_value = row.iloc[2]
                if currency_code in rates:
                    try:
                        rates[currency_code] = float(rate_value)
                    except:
                        pass
            if len(df) > 0:
                rate_info["publish_time"] = str(df.iloc[0, 3]) if pd.notna(df.iloc[0, 3]) else "未知"
                rate_info["fetch_time"] = str(df.iloc[0, 4]) if pd.notna(df.iloc[0, 4]) else "未知"
            rate_info["rates"] = rates
        except Exception as e:
            st.error(f"读取汇率文件时出错: {e}")
    return rate_info

def load_customer_data_from_excel():
    """从PAD抓取的客户数据Excel加载"""
    excel_path = "C:\\PAD_Data\\customers.xlsx"
    if os.path.exists(excel_path):
        try:
            df = pd.read_excel(excel_path)
            if len(df) > 0:
                latest = df.iloc[-1]
                return {
                    "success": True,
                    "data": {
                        "customer_name": str(latest.iloc[0]) if pd.notna(latest.iloc[0]) else "",
                        "customer_rep": str(latest.iloc[1]) if pd.notna(latest.iloc[1]) else "",
                        "customer_country": str(latest.iloc[2]) if pd.notna(latest.iloc[2]) else "",
                        "customer_email": str(latest.iloc[3]) if pd.notna(latest.iloc[3]) else "",
                        "customer_address": str(latest.iloc[4]) if pd.notna(latest.iloc[4]) else "",
                        "payment_terms": str(latest.iloc[5]) if pd.notna(latest.iloc[5]) else "",
                        "fetch_time": str(latest.iloc[6]) if len(df.columns) > 6 and pd.notna(latest.iloc[6]) else format_beijing_time()
                    }
                }
        except Exception as e:
            st.error(f"读取客户数据失败: {e}")
    return {"success": False, "data": None}

def load_product_data_from_excel():
    """从PAD抓取的商品数据Excel加载"""
    excel_path = "C:\\PAD_Data\\products.xlsx"
    if os.path.exists(excel_path):
        try:
            df = pd.read_excel(excel_path)
            if len(df) > 0:
                latest = df.iloc[-1]
                return {
                    "success": True,
                    "data": {
                        "product_code": str(latest.iloc[0]) if len(df.columns) > 0 and pd.notna(latest.iloc[0]) else "N003",
                        "goods_type": str(latest.iloc[1]) if len(df.columns) > 1 and pd.notna(latest.iloc[1]) else "宝石或半宝石",
                        "product_name": str(latest.iloc[2]) if len(df.columns) > 2 and pd.notna(latest.iloc[2]) else "蓝宝石",
                        "product_name_en": str(latest.iloc[3]) if len(df.columns) > 3 and pd.notna(latest.iloc[3]) else "Sapphires",
                        "specification_cn": str(latest.iloc[4]) if len(df.columns) > 4 and pd.notna(latest.iloc[4]) else "已加工，未镶嵌，天然，无等级，刚玉",
                        "specification_en": str(latest.iloc[5]) if len(df.columns) > 5 and pd.notna(latest.iloc[5]) else "Processed,not inlaid,natural,no grade,corundum",
                        "hs_code": str(latest.iloc[6]) if len(df.columns) > 6 and pd.notna(latest.iloc[6]) else "7103910000",
                        "sales_unit": str(latest.iloc[7]) if len(df.columns) > 7 and pd.notna(latest.iloc[7]) else "克拉（CT）",
                        "quantity": float(latest.iloc[8]) if len(df.columns) > 8 and pd.notna(latest.iloc[8]) else 0,
                        "price_per_ct": float(latest.iloc[9]) if len(df.columns) > 9 and pd.notna(latest.iloc[9]) else 0,
                        "package_unit": str(latest.iloc[10]) if len(df.columns) > 10 and pd.notna(latest.iloc[10]) else "纸箱（CARTON）",
                        "unit_conversion": str(latest.iloc[11]) if len(df.columns) > 11 and pd.notna(latest.iloc[11]) else "1000CT/CARTON",
                        "gross_weight": float(latest.iloc[12]) if len(df.columns) > 12 and pd.notna(latest.iloc[12]) else 0.70,
                        "net_weight": float(latest.iloc[13]) if len(df.columns) > 13 and pd.notna(latest.iloc[13]) else 0.20,
                        "volume_per_pack": float(latest.iloc[14]) if len(df.columns) > 14 and pd.notna(latest.iloc[14]) else 0.0400,
                        "legal_unit": str(latest.iloc[15]) if len(df.columns) > 15 and pd.notna(latest.iloc[15]) else "克拉（CT）",
                        "customs_supervision": str(latest.iloc[16]) if len(df.columns) > 16 and pd.notna(latest.iloc[16]) else "无",
                        "inspection_category": str(latest.iloc[17]) if len(df.columns) > 17 and pd.notna(latest.iloc[17]) else "无",
                        "transport_notes": str(latest.iloc[18]) if len(df.columns) > 18 and pd.notna(latest.iloc[18]) else "无",
                        "description": str(latest.iloc[19]) if len(df.columns) > 19 and pd.notna(latest.iloc[19]) else "",
                        "fetch_time": str(latest.iloc[20]) if len(df.columns) > 20 and pd.notna(latest.iloc[20]) else format_beijing_time()
                    }
                }
        except Exception as e:
            st.error(f"读取商品数据失败: {e}")
    return {"success": False, "data": None}

def run_pad_flow(flow_name):
    """调用Power Automate Desktop运行指定流程"""
    try:
        pad_path = "C:\\Program Files (x86)\\Power Automate Desktop\\PAD.Console.exe"
        if os.path.exists(pad_path):
            result = subprocess.run([pad_path, flow_name], capture_output=True, text=True, timeout=10)
            return {"success": True, "message": f"已启动PAD流程: {flow_name}"}
        else:
            st.info(f"模拟运行PAD流程: {flow_name}")
            return {"success": True, "message": f"模拟运行成功"}
    except Exception as e:
        return {"success": False, "message": str(e)}

def calculate_shipping_options(total_volume, total_weight, freight_20, freight_40, freight_40hq, lcl_rate_cbm, lcl_rate_kg):
    """
    根据总体积和总重量计算所有可能的运输方案
    1. LCL方案：取体积计费和重量计费的较大值
    2. 整箱方案：尝试各种集装箱组合，找出满足体积和重量要求的最便宜组合
    
    返回: (所有方案列表, 最佳方案索引, 最佳运费)
    """
    options = []
    
    # ========== 方案1: LCL散货 ==========
    lcl_volume_cost = total_volume * lcl_rate_cbm
    lcl_weight_cost = total_weight * lcl_rate_kg / 1000  # 转换为吨计费
    lcl_cost = max(lcl_volume_cost, lcl_weight_cost)
    
    # 判断是轻货还是重货
    if lcl_volume_cost > lcl_weight_cost:
        cargo_type = "轻货 (按体积计费)"
    else:
        cargo_type = "重货 (按重量计费)"
    
    options.append({
        "name": "LCL散货",
        "description": f"散货拼箱 - {cargo_type}",
        "type": "lcl",
        "containers": {},
        "cost_usd": lcl_cost,
        "volume_cost": lcl_volume_cost,
        "weight_cost": lcl_weight_cost,
        "details": f"体积计费: ${lcl_volume_cost:,.2f} | 重量计费: ${lcl_weight_cost:,.2f} | 取大值: ${lcl_cost:,.2f}"
    })
    
    # ========== 方案2: 整箱运输 ==========
    container_types = [
        {"name": "20'", "volume": 33, "weight": 25000, "freight": freight_20},
        {"name": "40'", "volume": 67, "weight": 29000, "freight": freight_40},
        {"name": "40'高", "volume": 76, "weight": 29000, "freight": freight_40hq}
    ]
    
    # 计算所需集装箱的最大数量（向上取整）
    max_20_by_volume = math.ceil(total_volume / 33)
    max_20_by_weight = math.ceil(total_weight / 25000)
    max_20 = max(max_20_by_volume, max_20_by_weight)
    
    max_40_by_volume = math.ceil(total_volume / 67)
    max_40_by_weight = math.ceil(total_weight / 29000)
    max_40 = max(max_40_by_volume, max_40_by_weight)
    
    max_40hq_by_volume = math.ceil(total_volume / 76)
    max_40hq_by_weight = math.ceil(total_weight / 29000)
    max_40hq = max(max_40hq_by_volume, max_40hq_by_weight)
    
    # 限制搜索范围，避免组合爆炸
    max_containers = min(max(max_20, max_40, max_40hq), 5)  # 最多搜索到5个柜子
    
    # 遍历所有可能的组合
    fcl_options = []
    
    for num_20 in range(max_containers + 1):
        for num_40 in range(max_containers + 1):
            for num_40hq in range(max_containers + 1):
                # 跳过全零组合
                if num_20 == 0 and num_40 == 0 and num_40hq == 0:
                    continue
                
                # 计算总容量
                total_container_volume = num_20 * 33 + num_40 * 67 + num_40hq * 76
                total_container_weight = num_20 * 25000 + num_40 * 29000 + num_40hq * 29000
                
                # 检查是否满足体积和重量要求
                if total_container_volume >= total_volume and total_container_weight >= total_weight:
                    # 计算集装箱总费用
                    container_cost = num_20 * freight_20 + num_40 * freight_40 + num_40hq * freight_40hq
                    
                    # 计算利用率
                    volume_utilization = (total_volume / total_container_volume) * 100
                    weight_utilization = (total_weight / total_container_weight) * 100
                    
                    # 构建方案名称
                    containers_desc = []
                    if num_20 > 0:
                        containers_desc.append(f"{num_20}×20'")
                    if num_40 > 0:
                        containers_desc.append(f"{num_40}×40'")
                    if num_40hq > 0:
                        containers_desc.append(f"{num_40hq}×40'高")
                    
                    scheme_name = " + ".join(containers_desc)
                    
                    # 计算剩余空间
                    remaining_volume = total_container_volume - total_volume
                    remaining_weight = total_container_weight - total_weight
                    
                    fcl_options.append({
                        "name": scheme_name,
                        "description": f"{scheme_name} - 体积利用率: {volume_utilization:.1f}%, 重量利用率: {weight_utilization:.1f}%",
                        "type": "fcl",
                        "containers": {"20'": num_20, "40'": num_40, "40'高": num_40hq},
                        "cost_usd": container_cost,
                        "volume_utilization": volume_utilization,
                        "weight_utilization": weight_utilization,
                        "remaining_volume": remaining_volume,
                        "remaining_weight": remaining_weight,
                        "details": f"总运费: ${container_cost:,.2f} | 剩余体积: {remaining_volume:.1f}CBM | 剩余重量: {remaining_weight:.1f}KG"
                    })
    
    # 去重（基于方案名称和成本）
    unique_fcl = {}
    for opt in fcl_options:
        key = f"{opt['name']}_{opt['cost_usd']:.2f}"
        if key not in unique_fcl:
            unique_fcl[key] = opt
    
    # 按成本排序
    fcl_options = list(unique_fcl.values())
    fcl_options.sort(key=lambda x: x["cost_usd"])
    
    # 添加整箱方案到总方案列表
    options.extend(fcl_options)
    
    # 按成本排序所有方案
    options.sort(key=lambda x: x["cost_usd"])
    
    # 找出最佳方案（最低成本）
    if options:
        best_index = 0  # 排序后第一个就是最便宜的
    else:
        best_index = 0
    
    return options, best_index

# -------------------- Session State 初始化 --------------------
if 'quote_history' not in st.session_state:
    st.session_state.quote_history = []
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
if 'handling_fee' not in st.session_state:
    st.session_state.handling_fee = 100.0
if 'inspection_fee' not in st.session_state:
    st.session_state.inspection_fee = 200.0
if 'document_fee' not in st.session_state:
    st.session_state.document_fee = 300.0
if 'insurance_rate' not in st.session_state:
    st.session_state.insurance_rate = 0.3
if 'freight_20' not in st.session_state:
    st.session_state.freight_20 = 1200.0
if 'freight_40' not in st.session_state:
    st.session_state.freight_40 = 1800.0
if 'freight_40hq' not in st.session_state:
    st.session_state.freight_40hq = 2000.0
if 'lcl_rate_cbm' not in st.session_state:
    st.session_state.lcl_rate_cbm = 50.0   # LCL(M) 按体积费率 (USD/CBM)
if 'lcl_rate_kg' not in st.session_state:
    st.session_state.lcl_rate_kg = 2000.0  # LCL(W) 按重量费率 (USD/吨)

# 加载汇率数据
rate_info = load_rates_from_excel()
exchange_rates = rate_info["rates"]

# ==================== 页面内容开始 ====================

# -------------------- 顶部公司信息及PAD按钮 --------------------
st.markdown("""
<div class="main-header">
    <h1>💰 AI价到 - 小微外贸智能折扣助手</h1>
    <div class="subtitle">智能报价 · 精准计算 · 一键成交</div>
</div>
""", unsafe_allow_html=True)

# 第一行：PAD抓取按钮和北京时间
col_pad1, col_pad2, col_pad3, col_pad4 = st.columns(4)
with col_pad1:
    if st.button("🤖 抓取客户信息 (PAD)", use_container_width=True):
        with st.spinner("正在启动Power Automate Desktop抓取客户信息..."):
            result = run_pad_flow("FetchCustomerFromAlibaba")
            if result["success"]:
                st.success(result["message"])
                time.sleep(3)
                customer_result = load_customer_data_from_excel()
                if customer_result["success"]:
                    st.session_state.customer_data = customer_result["data"]
                    st.session_state.customer_fetched = True
                    st.rerun()
                else:
                    st.warning("未找到客户数据文件，请确保PAD流程已正确运行")
            else:
                st.error(f"启动失败: {result['message']}")
with col_pad2:
    if st.button("📦 抓取商品信息 (PAD)", use_container_width=True):
        with st.spinner("正在启动Power Automate Desktop抓取商品信息..."):
            result = run_pad_flow("FetchProductFromMarket")
            if result["success"]:
                st.success(result["message"])
                time.sleep(3)
                product_result = load_product_data_from_excel()
                if product_result["success"]:
                    st.session_state.product_data = product_result["data"]
                    st.session_state.product_fetched = True
                    st.rerun()
                else:
                    st.warning("未找到商品数据文件，请确保PAD流程已正确运行")
            else:
                st.error(f"启动失败: {result['message']}")
with col_pad3:
    if st.button("📊 刷新汇率 (PAD)", use_container_width=True):
        with st.spinner("正在启动Power Automate Desktop更新汇率..."):
            result = run_pad_flow("FetchBOERates")
            if result["success"]:
                st.success(result["message"])
                time.sleep(2)
                st.rerun()
            else:
                st.error(f"启动失败: {result['message']}")
with col_pad4:
    st.markdown(f"<div style='text-align: center; padding: 0.5rem; background-color: #e9ecef; border-radius: 5px;'>🕒 {format_beijing_time()}</div>", unsafe_allow_html=True)

st.markdown("---")

# -------------------- 侧边栏：汇率、HS信息、物流信息 --------------------
with st.sidebar:
    st.markdown('<p class="sidebar-header">💱 汇率</p>', unsafe_allow_html=True)
    # 汇率状态
    if rate_info["file_exists"]:
        st.markdown("✅ <span class='status-badge status-success'>PAD汇率数据已连接</span>", unsafe_allow_html=True)
    else:
        st.markdown("⚠️ <span class='status-badge status-warning'>使用默认汇率</span>", unsafe_allow_html=True)
    st.caption(f"牌价时间: {rate_info['publish_time']}")
    st.caption(f"抓取时间: {rate_info['fetch_time']}")
    
    # 货币选择
    available_currencies = list(exchange_rates.keys())
    target_currency = st.selectbox("报价货币", available_currencies, 
                                  index=available_currencies.index(st.session_state.selected_currency) 
                                  if st.session_state.selected_currency in available_currencies else 0,
                                  key="sidebar_currency")
    st.session_state.selected_currency = target_currency
    current_rate = exchange_rates[target_currency]
    st.metric(f"1 {target_currency} = ", f"{current_rate:.4f} CNY")
    
    st.markdown("---")
    
    st.markdown('<p class="sidebar-header">📋 HS编码信息</p>', unsafe_allow_html=True)
    hs_code_display = st.session_state.product_data.get("hs_code", "未获取") if st.session_state.product_fetched else "未填写"
    st.text_input("HS编码", value=hs_code_display, disabled=True, key="hs_code_sidebar")
    st.markdown("**海关总署查询** [点击访问](http://www.customs.gov.cn)")
    st.checkbox("是否享受优惠税率", key="preferential_tax")
    
    st.markdown("---")
    
    st.markdown('<p class="sidebar-header">🚢 物流信息</p>', unsafe_allow_html=True)
    
    # 显示集装箱参数表
    st.markdown("""
    <div class="container-table">
        <table>
            <tr>
                <th>类型</th>
                <th>最大体积(CBM)</th>
                <th>最大重量(KG)</th>
            </tr>
            <tr>
                <td>LCL(M)</td>
                <td>-</td>
                <td>-</td>
            </tr>
            <tr>
                <td>LCL(W)</td>
                <td>-</td>
                <td>-</td>
            </tr>
            <tr>
                <td>20'</td>
                <td>33</td>
                <td>25000</td>
            </tr>
            <tr>
                <td>40'</td>
                <td>67</td>
                <td>29000</td>
            </tr>
            <tr>
                <td>40'高</td>
                <td>76</td>
                <td>29000</td>
            </tr>
        </table>
    </div>
    """, unsafe_allow_html=True)
    
    col_from, col_to = st.columns(2)
    with col_from:
        departure_port = st.text_input("起运港", value="Shanghai", key="departure_port")
    with col_to:
        default_dest = country_port_map.get(st.session_state.customer_data.get("customer_country", ""), "")
        destination_port = st.text_input("目的港", value=default_dest, key="destination_port")
    
    st.markdown("**集装箱运费估算 (USD)**")
    
    freight_20 = st.number_input("20'", value=float(st.session_state.freight_20), step=50.0, key="freight_20_input")
    freight_40 = st.number_input("40'", value=float(st.session_state.freight_40), step=50.0, key="freight_40_input")
    freight_40hq = st.number_input("40'高", value=float(st.session_state.freight_40hq), step=50.0, key="freight_40hq_input")
    
    st.markdown("**LCL散货费率**")
    st.caption("LCL运费 = max(体积×LCL(M)单价, 重量×LCL(W)单价)")
    lcl_rate_cbm = st.number_input("LCL(M) (USD/CBM)", value=float(st.session_state.lcl_rate_cbm), step=5.0, key="lcl_rate_cbm_input")
    lcl_rate_kg = st.number_input("LCL(W) (USD/吨)", value=float(st.session_state.lcl_rate_kg), step=100.0, key="lcl_rate_kg_input")
    
    if st.button("更新运费设置", key="update_freight"):
        st.session_state.freight_20 = freight_20
        st.session_state.freight_40 = freight_40
        st.session_state.freight_40hq = freight_40hq
        st.session_state.lcl_rate_cbm = lcl_rate_cbm
        st.session_state.lcl_rate_kg = lcl_rate_kg
        st.success("运费设置已更新")
        st.rerun()
    
    st.caption("数据来源: 环球运费网 / PAD抓取")

# -------------------- 主区域：左右公司/客户信息 --------------------
col_left, col_right = st.columns(2, gap="large")

with col_left:
    st.markdown("### 🏢 本公司信息")
    with st.container():
        company_name = st.text_input("公司名称", "ABC International Trading CO. Ltd", key="company_name")
        company_phone = st.text_input("联系电话", "+86 21 1234 5678", key="company_phone")
        company_email = st.text_input("联系邮箱", "info@abctrading.com", key="company_email")
        company_website = st.text_input("网站", "www.abctrading.com", key="company_website")

with col_right:
    st.markdown("### 👥 客户信息")
    if st.session_state.customer_fetched:
        st.success("✅ 已从PAD抓取客户数据")
    else:
        st.info("⏳ 可点击上方抓取按钮获取")
    default_customer = st.session_state.customer_data if st.session_state.customer_data else {}
    
    col_cust1, col_cust2 = st.columns(2)
    
    with col_cust1:
        customer = st.text_input(
            "客户名称", 
            value=default_customer.get("customer_name", ""), 
            key="customer_name_input"
        )
        rep = st.text_input(
            "客户代表", 
            value=default_customer.get("customer_rep", ""), 
            key="customer_rep_input"
        )
        country_index = 0
        if default_customer.get("customer_country") and default_customer["customer_country"] in country_port_map:
            countries = list(country_port_map.keys())
            country_index = countries.index(default_customer["customer_country"])
        country = st.selectbox(
            "目的国家", 
            list(country_port_map.keys()), 
            index=country_index, 
            key="customer_country_input"
        )
        port = country_port_map.get(country, "San Antonio")
        st.text_input("目的港口", value=port, disabled=True, key="customer_port_input")
    
    with col_cust2:
        email = st.text_input(
            "邮箱", 
            value=default_customer.get("customer_email", ""), 
            key="customer_email_input"
        )
        address = st.text_area(
            "公司地址", 
            value=default_customer.get("customer_address", ""), 
            key="customer_address_input", 
            height=100
        )
        payment_options = ["T/T 30% deposit", "L/C at sight", "D/P", "D/A", "T/T 100% in advance"]
        payment_index = 0
        if default_customer.get("payment_terms") in payment_options:
            payment_index = payment_options.index(default_customer["payment_terms"])
        payment_terms = st.selectbox(
            "付款方式", 
            payment_options, 
            index=payment_index, 
            key="payment_terms_input"
        )

st.markdown("---")

# -------------------- 商品信息 --------------------
st.markdown("### 💎 商品信息")
if st.session_state.product_fetched:
    st.success("✅ 已从PAD抓取商品数据")
else:
    st.info("⏳ 可点击上方抓取按钮获取或手动输入")

default_product = st.session_state.product_data if st.session_state.product_data else {}

col1, col2, col3, col4 = st.columns(4)
with col1:
    product_code = st.text_input("商品编号", value=default_product.get("product_code", "N003"), key="product_code")
with col2:
    goods_type = st.text_input("货物类型", value=default_product.get("goods_type", "宝石或半宝石"), key="goods_type")
with col3:
    product_name = st.text_input("商品名称", value=default_product.get("product_name", "蓝宝石"), key="product_name")
with col4:
    product_name_en = st.text_input("英文名称", value=default_product.get("product_name_en", "Sapphires"), key="product_name_en")

col5, col6, col7, col8 = st.columns(4)
with col5:
    specification_cn = st.text_input("规格型号（中文）", value=default_product.get("specification_cn", "已加工，未镶嵌，天然，无等级，刚玉"), key="spec_cn")
with col6:
    specification_en = st.text_input("规格型号（英文）", value=default_product.get("specification_en", "Processed,not inlaid,natural,no grade,corundum"), key="spec_en")
with col7:
    hs_code = st.text_input("HS编码", value=default_product.get("hs_code", "7103910000"), key="hs_code")
with col8:
    sales_unit = st.text_input("销售单位", value=default_product.get("sales_unit", "克拉（CT）"), key="sales_unit")

col9, col10, col11, col12 = st.columns(4)
with col9:
    quantity = st.number_input(
        "数量 (克拉)", 
        value=float(default_product.get("quantity", 0)), 
        step=100.0,
        min_value=0.0,
        format="%.0f",
        key="quantity"
    )
with col10:
    price_per_ct = st.number_input(
        "采购单价 (￥/克拉)", 
        value=float(default_product.get("price_per_ct", 0.0)), 
        step=1.0,
        min_value=0.0,
        format="%.2f",
        key="price"
    )
with col11:
    package_unit = st.text_input("包装单位", value=default_product.get("package_unit", "纸箱（CARTON）"), key="package_unit")
with col12:
    unit_conversion = st.text_input("单位换算", value=default_product.get("unit_conversion", "1000CT/CARTON"), key="unit_conversion")

col13, col14, col15, col16 = st.columns(4)
with col13:
    gross_weight = st.number_input(
        "毛重 (KGS/纸箱)", 
        value=float(default_product.get("gross_weight", 0.70)),
        format="%.2f",
        min_value=0.0,
        step=0.1,
        key="gross_weight"
    )
with col14:
    net_weight = st.number_input(
        "净重 (KGS/纸箱)", 
        value=float(default_product.get("net_weight", 0.20)),
        format="%.2f",
        min_value=0.0,
        step=0.1,
        key="net_weight"
    )
with col15:
    volume_per_pack = st.number_input(
        "体积 (CBM/纸箱)", 
        value=float(default_product.get("volume_per_pack", 0.0400)),
        format="%.4f",
        min_value=0.0,
        step=0.001,
        key="volume"
    )
with col16:
    legal_unit = st.text_input("法定单位", value=default_product.get("legal_unit", "克拉（CT）"), key="legal_unit")

col17, col18, col19, col20 = st.columns(4)
with col17:
    customs_supervision = st.text_input("海关监管条件", value=default_product.get("customs_supervision", "无"), key="customs_supervision")
with col18:
    inspection_category = st.text_input("检验检疫类别", value=default_product.get("inspection_category", "无"), key="inspection_category")
with col19:
    transport_notes = st.text_input("运输说明", value=default_product.get("transport_notes", "无"), key="transport_notes")
with col20:
    description = st.text_input("商品描述", value=default_product.get("description", ""), key="description")

st.markdown("---")

# -------------------- 贸易术语 & 支付方式 --------------------
st.markdown("### 📋 贸易术语 & 支付方式")
col_term, col_pay = st.columns(2)
with col_term:
    term_options = [term["name"] for term in INCOTERMS_2020]
    selected_term = st.selectbox("贸易术语 (Incoterms 2020)", term_options, index=1, key="selected_term")
    selected_term_detail = next((term for term in INCOTERMS_2020 if term["name"] == selected_term), INCOTERMS_2020[0])
    st.caption(f"{selected_term_detail['description'][:100]}...")
    
    profit_margin = st.slider("利润率 (%)", min_value=5, max_value=100, value=20, step=5, key="profit_margin")
    tax_rate = st.slider("出口退税率 (%)", min_value=0, max_value=17, value=13, step=1, key="tax_rate")

with col_pay:
    st.markdown("**付款方式**")
    payment_method = st.selectbox("付款方式", ["T/T", "L/C", "D/P", "D/A"], key="payment_method")
    payment_bank = st.text_input("付款银行", value="Bank of China", key="payment_bank")
    payment_terms_detail = st.text_input("付款条件", value="30% deposit, 70% against B/L", key="payment_terms_detail")

st.markdown("---")

# -------------------- 开始报价按钮和出口预算表 --------------------
col_btn1, col_btn2 = st.columns([1, 5])
with col_btn1:
    calculate_pressed = st.button("开始报价", type="primary", use_container_width=True)

st.markdown("### 📊 出口预算表")

# 获取当前汇率
current_exchange_rate = exchange_rates[st.session_state.selected_currency]

if calculate_pressed:
    if quantity > 0 and price_per_ct > 0:
        # 计算总成本
        total_cost_cny = quantity * price_per_ct
        exchange_rate = current_exchange_rate
        total_cost_target = total_cost_cny / exchange_rate
        
        # 计算箱数、总体积和总重量
        try:
            conversion_parts = unit_conversion.split('/')
            if len(conversion_parts) == 2:
                ct_per_carton = float(conversion_parts[0].replace('CT', '').strip())
                total_packages = math.ceil(quantity / ct_per_carton)
            else:
                total_packages = math.ceil(quantity / 1000)
        except:
            total_packages = math.ceil(quantity / 1000)
        
        total_volume = total_packages * volume_per_pack
        total_weight = total_packages * gross_weight
        
        # 计算运输方案
        shipping_options, best_index = calculate_shipping_options(
            total_volume, 
            total_weight,
            st.session_state.freight_20,
            st.session_state.freight_40,
            st.session_state.freight_40hq,
            st.session_state.lcl_rate_cbm,
            st.session_state.lcl_rate_kg
        )
        
        best_option = shipping_options[best_index]
        freight_usd = best_option["cost_usd"]
        freight_cny = freight_usd * exchange_rate
        
        # 其他费用
        domestic_fee = st.session_state.handling_fee + st.session_state.inspection_fee + st.session_state.document_fee
        bank_charges = 200
        insurance_fee = total_cost_cny * (st.session_state.insurance_rate / 100)
        tax_refund = total_cost_cny * (tax_rate / 100)
        
        # 对外报价
        quoted_price_target = total_cost_target * (1 + profit_margin/100)
        quoted_price_cny = quoted_price_target * exchange_rate
        
        budget = {
            "采购成本": total_cost_cny,
            "国内费用": domestic_fee,
            "银行费用": bank_charges,
            "海运费": freight_cny,
            "保险费": insurance_fee,
            "出口退税": tax_refund,
            "总成本": total_cost_cny + domestic_fee + bank_charges + freight_cny + insurance_fee - tax_refund,
            "对外报价": quoted_price_target,
            "对外报价CNY": quoted_price_cny,
            "预期盈亏额": quoted_price_cny - (total_cost_cny + domestic_fee + bank_charges + freight_cny + insurance_fee - tax_refund),
            "预期盈亏率": 0,
            "总箱数": total_packages,
            "总体积": total_volume,
            "总重量": total_weight,
            "shipping_options": shipping_options,
            "best_shipping_index": best_index,
            "selected_shipping": best_option["name"]
        }
        budget["预期盈亏率"] = (budget["预期盈亏额"] / budget["总成本"]) * 100 if budget["总成本"] != 0 else 0
        
        st.session_state.budget = budget
        st.success("计算完成！")
        st.balloons()
    else:
        st.warning("请填写商品数量和单价")

# 显示预算表
if st.session_state.budget:
    b = st.session_state.budget
    
    # 显示运输方案对比
    with st.expander("🚢 查看运输方案对比", expanded=True):
        st.markdown("#### 所有运输方案（按成本排序）")
        for i, option in enumerate(b.get("shipping_options", [])):
            if i == b.get("best_shipping_index", 0):
                st.markdown(f"""
                <div class="shipping-option best-option">
                    <strong>✅ 最佳方案 #{i+1}: {option['name']}</strong><br>
                    {option['description']}<br>
                    运费: ${option['cost_usd']:,.2f} | ￥{option['cost_usd'] * current_exchange_rate:,.2f}<br>
                    <small>{option['details']}</small>
                </div>
                """, unsafe_allow_html=True)
            else:
                st.markdown(f"""
                <div class="shipping-option">
                    <strong>方案 #{i+1}: {option['name']}</strong><br>
                    {option['description']}<br>
                    运费: ${option['cost_usd']:,.2f} | ￥{option['cost_usd'] * current_exchange_rate:,.2f}
                </div>
                """, unsafe_allow_html=True)
    
    # 显示成本明细
    col_b1, col_b2 = st.columns(2)
    with col_b1:
        st.markdown("**📥 成本项**")
        st.metric("采购成本", f"￥{b['采购成本']:,.2f}")
        st.metric("国内费用", f"￥{b['国内费用']:,.2f}")
        st.metric("银行费用", f"￥{b['银行费用']:,.2f}")
        st.metric("海运费", f"￥{b['海运费']:,.2f}")
        st.metric("保险费", f"￥{b['保险费']:,.2f}")
    
    with col_b2:
        st.markdown("**📤 收入与退税**")
        st.metric("出口退税", f"￥{b['出口退税']:,.2f}")
        st.markdown("**💰 盈亏分析**")
        st.metric("总成本", f"￥{b['总成本']:,.2f}")
        st.metric("对外报价", f"{b['对外报价']:,.2f} {st.session_state.selected_currency}")
        st.metric("预期盈亏额", f"￥{b['预期盈亏额']:,.2f}", delta=f"{b['预期盈亏率']:.1f}%")
        
        # 新报价单价
        new_price_per_ct_target = b['对外报价'] / quantity if quantity > 0 else 0
        new_price_per_ct_cny = new_price_per_ct_target * current_exchange_rate
        st.metric("新报价单价", f"{new_price_per_ct_target:.2f} {st.session_state.selected_currency}/CT")
        
        st.markdown("**📦 货运信息**")
        st.metric("总箱数", f"{b['总箱数']:,} 箱")
        st.metric("总体积", f"{b['总体积']:.2f} CBM")
        st.metric("总重量", f"{b['总重量']:.2f} KG")
        st.metric("推荐运输方式", b.get("selected_shipping", "未计算"))

st.markdown("---")

# -------------------- 报价历史 --------------------
with st.sidebar:
    st.markdown("---")
    st.markdown('<p class="sidebar-header">📜 报价历史</p>', unsafe_allow_html=True)
    if st.session_state.quote_history:
        for i, quote in enumerate(st.session_state.quote_history[-5:]):
            st.markdown(f"**{quote['date']}** - {quote['product']}")
            st.markdown(f"金额: {quote['amount']}")
            st.markdown("---")
    else:
        st.info("暂无报价历史")

# -------------------- 底部版权 --------------------
col_footer1, col_footer2, col_footer3 = st.columns(3)
with col_footer1:
    st.markdown(f"© 2026 {company_name}")
with col_footer2:
    st.markdown("技术支持: AI价到团队")
with col_footer3:
    st.markdown("PAD数据源: 阿里巴巴国际站")






















