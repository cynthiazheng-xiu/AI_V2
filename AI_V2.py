import streamlit as st
import pandas as pd
import math
import subprocess
import os
import time
from datetime import datetime, timedelta, timezone

# 页面配置
st.set_page_config(
    page_title="AI价到 - 小微外贸智能报价助手",
    page_icon="💰",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 自定义CSS样式
st.markdown("""
<style>
    .main-header {
        background: linear-gradient(135deg, #0A174E 0%, #1D2B5E 100%);
        padding: 1.5rem 2rem;
        border-radius: 20px;
        color: white;
        margin-bottom: 1rem;
    }
    .company-info {
        background: linear-gradient(135deg, #1e3c72 0%, #2a5298 100%);
        padding: 1rem 2rem;
        border-radius: 15px;
        color: white;
        margin-bottom: 1.5rem;
    }
    .section-header {
        background-color: #f0f2f6;
        padding: 0.5rem 1rem;
        border-radius: 8px;
        margin-bottom: 1rem;
        font-weight: bold;
        color: #0A174E;
    }
    .fetch-button {
        background-color: #28a745;
        color: white;
        padding: 0.5rem 1rem;
        border-radius: 5px;
        text-align: center;
        font-weight: bold;
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

# 获取北京时间的函数
def get_beijing_time():
    """返回当前的北京时间"""
    utc_now = datetime.utcnow().replace(tzinfo=timezone.utc)
    beijing_now = utc_now.astimezone(timezone(timedelta(hours=8)))
    return beijing_now

def format_beijing_time(format_str='%Y-%m-%d %H:%M:%S'):
    """格式化北京时间"""
    return get_beijing_time().strftime(format_str)

# 国家港口映射
country_port_map = {
    "Chile": "San Antonio", "USA": "Los Angeles", "Germany": "Hamburg",
    "Philippines": "Manila", "China": "Shanghai", "Japan": "Tokyo",
    "UK": "Felixstowe", "France": "Le Havre", "Italy": "Genoa",
    "Australia": "Sydney", "Brazil": "Santos", "India": "Mumbai"
}

# 默认汇率（当Excel文件不存在时使用）
DEFAULT_RATES = {
    "USD": 6.9257,  # 美元
    "EUR": 8.1863,  # 欧元
    "GBP": 9.3729,  # 英镑
    "JPY": 0.044775, # 日元
    "HKD": 0.8858,  # 港币
    "AUD": 4.9092,  # 澳元
    "CAD": 5.0734,  # 加元
    "CHF": 8.9762,  # 瑞士法郎
    "SGD": 5.4721   # 新加坡元
}

# 商品预设数据
product_presets = {
    "蓝宝石 (Sapphires)": {
        "hs_code": "7103910000",
        "price_per_ct": 50.0,
        "volume_per_pack": 0.04,
        "weight_per_pack": 0.7,
        "description": "天然蓝宝石，优质切割",
        "category": "宝石"
    },
    "红宝石 (Rubies)": {
        "hs_code": "7103910000",
        "price_per_ct": 80.0,
        "volume_per_pack": 0.035,
        "weight_per_pack": 0.6,
        "description": "缅甸红宝石，色泽鲜艳",
        "category": "宝石"
    },
    "祖母绿 (Emeralds)": {
        "hs_code": "7103910000",
        "price_per_ct": 120.0,
        "volume_per_pack": 0.045,
        "weight_per_pack": 0.8,
        "description": "哥伦比亚祖母绿，高净度",
        "category": "宝石"
    },
    "钻石 (Diamonds)": {
        "hs_code": "7102390000",
        "price_per_ct": 500.0,
        "volume_per_pack": 0.03,
        "weight_per_pack": 0.5,
        "description": "天然钻石，1克拉以上",
        "category": "宝石"
    },
    "水晶 (Crystals)": {
        "hs_code": "7103999000",
        "price_per_ct": 15.0,
        "volume_per_pack": 0.05,
        "weight_per_pack": 1.0,
        "description": "天然水晶，多种颜色",
        "category": "半宝石"
    },
    "玛瑙 (Agate)": {
        "hs_code": "7103999000",
        "price_per_ct": 8.0,
        "volume_per_pack": 0.06,
        "weight_per_pack": 1.2,
        "description": "巴西玛瑙，天然纹路",
        "category": "半宝石"
    },
    "自定义商品": {
        "hs_code": "",
        "price_per_ct": 0.0,
        "volume_per_pack": 0.0,
        "weight_per_pack": 0.0,
        "description": "手动输入商品信息",
        "category": "其他"
    }
}

# 从Excel加载汇率数据的函数
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

# 从客户数据Excel加载的函数
def load_customer_data_from_excel():
    """从PAD抓取的客户数据Excel加载"""
    excel_path = "C:\\PAD_Data\\customers.xlsx"
    
    if os.path.exists(excel_path):
        try:
            df = pd.read_excel(excel_path)
            if len(df) > 0:
                # 取最新一条记录
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

# 从商品数据Excel加载的函数
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
                        "product_name": str(latest.iloc[0]) if pd.notna(latest.iloc[0]) else "",
                        "hs_code": str(latest.iloc[1]) if pd.notna(latest.iloc[1]) else "",
                        "quantity": float(latest.iloc[2]) if pd.notna(latest.iloc[2]) else 0,
                        "price_per_ct": float(latest.iloc[3]) if pd.notna(latest.iloc[3]) else 0,
                        "volume_per_pack": float(latest.iloc[4]) if pd.notna(latest.iloc[4]) else 0,
                        "weight_per_pack": float(latest.iloc[5]) if pd.notna(latest.iloc[5]) else 0,
                        "description": str(latest.iloc[6]) if len(df.columns) > 6 and pd.notna(latest.iloc[6]) else "",
                        "fetch_time": str(latest.iloc[7]) if len(df.columns) > 7 and pd.notna(latest.iloc[7]) else format_beijing_time()
                    }
                }
        except Exception as e:
            st.error(f"读取商品数据失败: {e}")
    
    return {"success": False, "data": None}

# 启动PAD流程的函数
def run_pad_flow(flow_name):
    """调用Power Automate Desktop运行指定流程"""
    try:
        # PAD的命令行调用方式（需要根据实际安装路径调整）
        pad_path = "C:\\Program Files (x86)\\Power Automate Desktop\\PAD.Console.exe"
        
        if os.path.exists(pad_path):
            # 方法1：直接调用PAD控制台
            result = subprocess.run([pad_path, flow_name], capture_output=True, text=True, timeout=10)
            return {"success": True, "message": f"已启动PAD流程: {flow_name}"}
        else:
            # 方法2：如果没有PAD，模拟成功（用于演示）
            st.info(f"模拟运行PAD流程: {flow_name}")
            return {"success": True, "message": f"模拟运行成功"}
            
    except Exception as e:
        return {"success": False, "message": str(e)}

# 初始化session state
if 'current_product' not in st.session_state:
    st.session_state.current_product = "蓝宝石 (Sapphires)"
if 'quote_history' not in st.session_state:
    st.session_state.quote_history = []
if 'selected_currency' not in st.session_state:
    st.session_state.selected_currency = "USD"
if 'customer_data' not in st.session_state:
    st.session_state.customer_data = {}
if 'product_data' not in st.session_state:
    st.session_state.product_data = {}

# 加载汇率数据
rate_info = load_rates_from_excel()
exchange_rates = rate_info["rates"]

# ==================== 顶部区域 ====================
st.markdown("""
<div class="main-header">
    <h1 style="margin:0;">💰 AI价到 - 小微外贸智能报价助手</h1>
</div>
""", unsafe_allow_html=True)

# 公司信息行
col_company1, col_company2, col_company3, col_company4 = st.columns(4)
with col_company1:
    company_name = st.text_input("公司名称", "ABC International Trading CO. Ltd", key="company_name")
with col_company2:
    company_phone = st.text_input("联系电话", "+86 21 1234 5678", key="company_phone")
with col_company3:
    company_email = st.text_input("联系邮箱", "info@abctrading.com", key="company_email")
with col_company4:
    company_website = st.text_input("网站", "www.abctrading.com", key="company_website")

# PAD抓取按钮行
col_pad1, col_pad2, col_pad3, col_pad4 = st.columns(4)
with col_pad1:
    if st.button("🤖 抓取客户信息 (PAD)", use_container_width=True):
        with st.spinner("正在启动Power Automate Desktop抓取客户信息..."):
            result = run_pad_flow("FetchCustomerFromAlibaba")
            if result["success"]:
                st.success(result["message"])
                # 模拟等待PAD完成（实际需要监控文件变化）
                time.sleep(3)
                # 加载抓取的数据
                customer_result = load_customer_data_from_excel()
                if customer_result["success"]:
                    st.session_state.customer_data = customer_result["data"]
                    st.rerun()
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
                    st.session_state.current_product = "自定义商品"
                    st.rerun()
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

# 汇率状态行
col_rate_status1, col_rate_status2, col_rate_status3, col_rate_status4 = st.columns(4)
with col_rate_status1:
    if rate_info["file_exists"]:
        st.markdown("✅ <span class='status-badge status-success'>PAD汇率数据已连接</span>", unsafe_allow_html=True)
    else:
        st.markdown("⚠️ <span class='status-badge status-warning'>使用默认汇率</span>", unsafe_allow_html=True)
with col_rate_status2:
    st.markdown(f"📊 **牌价时间:** {rate_info['publish_time']}")
with col_rate_status3:
    st.markdown(f"🔄 **抓取时间:** {rate_info['fetch_time']}")
with col_rate_status4:
    st.markdown(f"💱 **当前货币:** {st.session_state.selected_currency}")

st.markdown("---")

# ==================== 侧边栏 ====================
with st.sidebar:
    st.markdown("### ⚙️ 系统设置")
    
    # 汇率设置
    with st.expander("💱 汇率设置（中国银行牌价）", expanded=True):
        available_currencies = list(exchange_rates.keys())
        target_currency = st.selectbox("报价货币", available_currencies, 
                                      index=available_currencies.index(st.session_state.selected_currency) 
                                      if st.session_state.selected_currency in available_currencies else 0)
        st.session_state.selected_currency = target_currency
        
        current_rate = exchange_rates[target_currency]
        st.metric(f"1 {target_currency} = ", f"{current_rate:.4f} CNY")
        
        with st.expander("查看所有汇率"):
            for currency, rate in exchange_rates.items():
                st.text(f"{currency}: {rate:.4f}")
    
    # 报价设置
    with st.expander("📊 报价设置", expanded=True):
        profit_margin = st.slider("默认利润率 (%)", min_value=5, max_value=100, value=20, step=5)
        tax_rate = st.slider("出口退税率 (%)", min_value=0, max_value=17, value=13, step=1)
        incoterms = st.selectbox("贸易术语", ["FOB", "CIF", "CFR", "EXW", "DDP"])
        
        col_s1, col_s2 = st.columns(2)
        with col_s1:
            handling_fee = st.number_input("操作费", value=100, step=10)
            inspection_fee = st.number_input("商检费", value=200, step=10)
        with col_s2:
            document_fee = st.number_input("文件费", value=300, step=10)
            insurance_rate = st.number_input("保险费率 (%)", value=0.3, step=0.1, format="%.1f")
    
    # 报价历史
    with st.expander("📜 报价历史", expanded=False):
        if st.session_state.quote_history:
            for i, quote in enumerate(st.session_state.quote_history[-5:]):
                st.markdown(f"**{quote['date']}** - {quote['product']}")
                st.markdown(f"客户: {quote['customer']} | 金额: {quote['amount']}")
                st.markdown("---")
        else:
            st.info("暂无报价历史")
    
    # PAD使用说明
    with st.expander("📖 PAD流程说明", expanded=False):
        st.markdown("""
        **Power Automate Desktop 流程：**
        
        1. **FetchCustomerFromAlibaba**
           - 从阿里巴巴询价页抓取客户信息
           - URL: `http://gjmystu.dianyuesoft.com/#/Alibaba/Inquiry/Detail?id=...`
           - 输出: `C:\\PAD_Data\\customers.xlsx`
        
        2. **FetchProductFromMarket**
           - 从国内采购市场抓取商品信息
           - URL: `http://gjmystu.dianyuesoft.com/#/Practical/Purchase/market`
           - 输出: `C:\\PAD_Data\\products.xlsx`
        
        3. **FetchBOERates**
           - 从中国银行抓取汇率
           - URL: `https://www.boc.cn/sourcedb/whpj/`
           - 输出: `C:\\ExchangeRates\\rates.xlsx`
        """)

# ==================== 客户信息和商品信息左右并列 ====================
col_left, col_right = st.columns(2, gap="large")

# 左侧：客户信息
with col_left:
    st.markdown("""
    <div class="section-header">
        📋 客户信息
    </div>
    """, unsafe_allow_html=True)
    
    # 如果session中有抓取的数据，使用它
    default_customer = st.session_state.customer_data if st.session_state.customer_data else {}
    
    customer = st.text_input("客户名称", 
                            value=default_customer.get("customer_name", "Antonia Continental Commerce Ltd."), 
                            key="customer_name_input")
    rep = st.text_input("客户代表", 
                       value=default_customer.get("customer_rep", "Alfredo Mariani"), 
                       key="customer_rep_input")
    
    country_index = 0
    if default_customer.get("customer_country"):
        countries = list(country_port_map.keys())
        if default_customer["customer_country"] in countries:
            country_index = countries.index(default_customer["customer_country"])
    
    country = st.selectbox("目的国家", list(country_port_map.keys()), 
                          index=country_index, key="customer_country_input")
    port = country_port_map.get(country, "San Antonio")
    st.text_input("目的港口", value=port, disabled=True, key="customer_port_input")
    
    email = st.text_input("邮箱", 
                         value=default_customer.get("customer_email", "16203962@yahoo.com"), 
                         key="customer_email_input")
    address = st.text_area("公司地址", 
                          value=default_customer.get("customer_address", "4 Talcahuano Court, Talcahuano, Chile"), 
                          key="customer_address_input", height=100)
    
    payment_options = ["T/T 30% deposit", "L/C at sight", "D/P", "D/A", "T/T 100% in advance"]
    payment_index = 0
    if default_customer.get("payment_terms") in payment_options:
        payment_index = payment_options.index(default_customer["payment_terms"])
    
    payment_terms = st.selectbox("付款方式", payment_options, 
                                index=payment_index, key="payment_terms_input")
    
    if default_customer.get("fetch_time"):
        st.info(f"📌 PAD抓取时间: {default_customer['fetch_time']}")

# 右侧：商品信息
with col_right:
    st.markdown("""
    <div class="section-header">
        💎 商品信息
    </div>
    """, unsafe_allow_html=True)
    
    # 商品快速切换
    st.markdown("**快速选择商品：**")
    
    categories = {}
    for product_name, product_data in product_presets.items():
        category = product_data.get("category", "其他")
        if category not in categories:
            categories[category] = []
        categories[category].append(product_name)
    
    for category, products in categories.items():
        if category != "其他":
            st.markdown(f"**{category}**")
            cols = st.columns(min(4, len(products)))
            for i, product_name in enumerate(products[:4]):
                with cols[i]:
                    if st.button(product_name.split()[0], key=f"btn_{product_name}", use_container_width=True):
                        st.session_state.current_product = product_name
                        st.rerun()
    
    st.markdown("---")
    
    # 如果session中有抓取的商品数据，使用它
    default_product = st.session_state.product_data if st.session_state.product_data else {}
    
    if st.session_state.current_product == "自定义商品" and default_product:
        product_name = st.text_input("商品名称", 
                                    value=default_product.get("product_name", "请输入商品名称"), 
                                    key="product_name_custom")
        hs_code = st.text_input("HS编码", 
                               value=default_product.get("hs_code", ""), 
                               key="hs_code_custom")
    else:
        selected_preset = product_presets[st.session_state.current_product]
        product_name = st.text_input("商品名称", st.session_state.current_product, disabled=True, key="product_name")
        hs_code = st.text_input("HS编码", selected_preset["hs_code"], key="hs_code")
    
    # 数量和单价
    col_q1, col_q2 = st.columns(2)
    with col_q1:
        if st.session_state.current_product == "自定义商品" and default_product.get("quantity"):
            quantity = st.number_input("数量 (克拉)", 
                                      value=default_product["quantity"], 
                                      step=100, key="quantity_custom")
        else:
            quantity = st.number_input("数量 (克拉)", value=5000, step=100, key="quantity")
    
    with col_q2:
        if st.session_state.current_product == "自定义商品":
            if default_product.get("price_per_ct"):
                price_per_ct = st.number_input("采购单价 (￥/克拉)", 
                                              value=default_product["price_per_ct"], 
                                              step=1.0, key="price_custom")
            else:
                price_per_ct = st.number_input("采购单价 (￥/克拉)", value=0.0, step=1.0, key="price_custom")
        else:
            price_per_ct = st.number_input("采购单价 (￥/克拉)", 
                                          value=selected_preset["price_per_ct"], 
                                          step=1.0, key="price")
    
    # 包装信息
    st.markdown("**包装信息**")
    col_p1, col_p2 = st.columns(2)
    with col_p1:
        if st.session_state.current_product == "自定义商品":
            if default_product.get("volume_per_pack"):
                volume_per_pack = st.number_input("单箱体积 (CBM)", 
                                                 value=default_product["volume_per_pack"], 
                                                 format="%.3f", key="volume_custom")
            else:
                volume_per_pack = st.number_input("单箱体积 (CBM)", value=0.0, format="%.3f", key="volume_custom")
        else:
            volume_per_pack = st.number_input("单箱体积 (CBM)", 
                                             value=selected_preset["volume_per_pack"], 
                                             format="%.3f", key="volume")
    
    with col_p2:
        if st.session_state.current_product == "自定义商品":
            if default_product.get("weight_per_pack"):
                weight_per_pack = st.number_input("单箱毛重 (KG)", 
                                                 value=default_product["weight_per_pack"], 
                                                 format="%.2f", key="weight_custom")
            else:
                weight_per_pack = st.number_input("单箱毛重 (KG)", value=0.0, format="%.2f", key="weight_custom")
        else:
            weight_per_pack = st.number_input("单箱毛重 (KG)", 
                                             value=selected_preset["weight_per_pack"], 
                                             format="%.2f", key="weight")
    
    # 显示商品描述
    if st.session_state.current_product != "自定义商品":
        st.info(f"📝 {selected_preset['description']}")
    elif default_product.get("description"):
        st.info(f"📝 {default_product['description']}")
    
    if default_product.get("fetch_time"):
        st.info(f"📌 PAD抓取时间: {default_product['fetch_time']}")

# ==================== 计算结果区域 ====================
st.markdown("---")
st.markdown("### 📊 报价计算结果")

col_btn1, col_btn2, col_btn3 = st.columns([1, 1, 2])
with col_btn1:
    if st.button("开始计算", type="primary", use_container_width=True):
        total_cost_cny = quantity * price_per_ct
        total_cost_target = total_cost_cny / exchange_rates[st.session_state.selected_currency]
        
        quote_record = {
            "date": format_beijing_time(),
            "product": product_name,
            "customer": customer,
            "amount": f"{total_cost_target:,.2f} {st.session_state.selected_currency}"
        }
        st.session_state.quote_history.append(quote_record)
        
        st.success("计算完成！")
        st.balloons()

with col_btn2:
    if st.button("📧 发送报价", use_container_width=True):
        st.info("报价单已准备发送！")

# 显示计算结果
col_result1, col_result2, col_result3, col_result4 = st.columns(4)

with col_result1:
    total_cost_cny = quantity * price_per_ct
    st.metric("采购总成本 (CNY)", f"￥{total_cost_cny:,.2f}")

with col_result2:
    total_cost_target = total_cost_cny / exchange_rates[st.session_state.selected_currency]
    st.metric(f"采购总成本 ({st.session_state.selected_currency})", 
              f"{total_cost_target:,.2f} {st.session_state.selected_currency}")

with col_result3:
    suggested_price = total_cost_target * (1 + profit_margin/100)
    st.metric(f"建议报价 ({st.session_state.selected_currency})", 
              f"{suggested_price:,.2f} {st.session_state.selected_currency}",
              delta=f"{profit_margin}% 利润率")

with col_result4:
    total_packages = math.ceil(quantity/100) if quantity > 0 else 0
    st.metric("总箱数", f"{total_packages:,} 箱")

# 显示详细商品信息
st.markdown("### 📦 商品详情")
col_detail1, col_detail2, col_detail3, col_detail4 = st.columns(4)
with col_detail1:
    st.info(f"**HS编码:** {hs_code}")
with col_detail2:
    st.info(f"**总箱数:** {total_packages:,} 箱")
with col_detail3:
    total_volume = total_packages * volume_per_pack if total_packages > 0 else 0
    st.info(f"**总体积:** {total_volume:.2f} CBM")
with col_detail4:
    total_weight = total_packages * weight_per_pack if total_packages > 0 else 0
    st.info(f"**总毛重:** {total_weight:.2f} KG")

st.markdown("---")

# 底部版权
col_footer1, col_footer2, col_footer3 = st.columns(3)
with col_footer1:
    st.markdown(f"© 2026 {company_name}")
with col_footer2:
    st.markdown("技术支持: AI价到团队")
with col_footer3:
    st.markdown("PAD数据源: 阿里巴巴询价页 / 国内采购市场 / 中国银行")










