import streamlit as st
import pandas as pd
import math
from datetime import datetime
import requests
import json
import time

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
        margin-bottom: 2rem;
    }
    .card {
        background: white;
        border: 1px solid #E5E9F0;
        border-radius: 16px;
        padding: 1.5rem;
        box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05);
        margin-bottom: 1rem;
    }
    .profit-card {
        background: linear-gradient(135deg, #059669 0%, #10B981 100%);
        color: white;
        padding: 1.5rem;
        border-radius: 16px;
        text-align: center;
    }
    .product-badge {
        background-color: #F3F4F6;
        padding: 0.25rem 0.75rem;
        border-radius: 20px;
        font-size: 0.9rem;
        color: #374151;
        display: inline-block;
        margin-right: 0.5rem;
        margin-bottom: 0.5rem;
        cursor: pointer;
    }
    .product-badge:hover {
        background-color: #E5E7EB;
    }
    .sidebar-info {
        background-color: #f8f9fa;
        padding: 1rem;
        border-radius: 10px;
        margin-bottom: 1rem;
    }
    .section-header {
        background-color: #f0f2f6;
        padding: 0.5rem 1rem;
        border-radius: 8px;
        margin-bottom: 1rem;
        font-weight: bold;
    }
    .fetch-button {
        background-color: #4CAF50;
        color: white;
        padding: 0.5rem 1rem;
        border-radius: 5px;
        text-align: center;
        margin-bottom: 1rem;
    }
    .status-success {
        color: #28a745;
        font-weight: bold;
    }
    .status-warning {
        color: #ffc107;
        font-weight: bold;
    }
    .status-error {
        color: #dc3545;
        font-weight: bold;
    }
</style>
""", unsafe_allow_html=True)

# 国家港口映射
country_port_map = {
    "Chile": "San Antonio", "USA": "Los Angeles", "Germany": "Hamburg",
    "Philippines": "Manila", "China": "Shanghai", "Japan": "Tokyo",
    "UK": "Felixstowe", "France": "Le Havre", "Italy": "Genoa",
    "Australia": "Sydney", "Brazil": "Santos", "India": "Mumbai"
}

# 汇率映射（示例数据）
exchange_rates = {
    "USD": 7.25,  # 美元
    "EUR": 7.85,  # 欧元
    "GBP": 9.15,  # 英镑
    "JPY": 0.048, # 日元
    "AUD": 4.75,  # 澳元
    "CAD": 5.35   # 加元
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

# 模拟从API抓取数据的函数（不使用BeautifulSoup）
def fetch_customer_data_from_api(source="power_automate"):
    """
    模拟从API或Power Automate Desktop抓取客户数据
    """
    try:
        # 这里模拟从不同来源获取数据
        if source == "power_automate":
            # 模拟从Power Automate Desktop获取的数据
            time.sleep(2)  # 模拟网络延迟
            return {
                "success": True,
                "data": {
                    "customer_name": "Global Trade Imports Ltd (从Power Automate抓取)",
                    "customer_rep": "John Smith",
                    "customer_country": "USA",
                    "customer_email": "john.smith@globaltrade.com",
                    "customer_address": "123 Trade Center, New York, NY 10001, USA",
                    "payment_terms": "L/C at sight",
                    "source": "Power Automate Desktop - 客户数据库"
                }
            }
        elif source == "alibaba":
            # 模拟从阿里巴巴国际站API获取
            time.sleep(1.5)
            return {
                "success": True,
                "data": {
                    "customer_name": "Alibaba Import Co., Ltd",
                    "customer_rep": "Li Wei",
                    "customer_country": "China",
                    "customer_email": "li.wei@alibaba.com",
                    "customer_address": "969 West Wen'er Road, Hangzhou, China",
                    "payment_terms": "T/T 30% deposit",
                    "source": "Alibaba.com API"
                }
            }
        elif source == "made_in_china":
            time.sleep(1.8)
            return {
                "success": True,
                "data": {
                    "customer_name": "Made-in-China Importer",
                    "customer_rep": "Wang Fang",
                    "customer_country": "Germany",
                    "customer_email": "fang.wang@made-in-china.com",
                    "customer_address": "Berlin Trade Center, Germany",
                    "payment_terms": "D/P",
                    "source": "Made-in-China.com API"
                }
            }
        else:
            return {
                "success": False,
                "error": "未知的数据来源"
            }
    except Exception as e:
        return {
            "success": False,
            "error": str(e)
        }

def fetch_product_data_from_api(source="power_automate"):
    """
    模拟从API或Power Automate Desktop抓取商品数据
    """
    try:
        if source == "power_automate":
            time.sleep(2.5)
            return {
                "success": True,
                "data": {
                    "product_name": "Premium Sapphires (从Power Automate抓取)",
                    "hs_code": "7103910000",
                    "quantity": 8000,
                    "price_per_ct": 55.0,
                    "volume_per_pack": 0.045,
                    "weight_per_pack": 0.75,
                    "description": "高级蓝宝石，通过Power Automate抓取的商品数据",
                    "source": "Power Automate Desktop - 商品管理系统"
                }
            }
        elif source == "gemstone_database":
            time.sleep(2)
            return {
                "success": True,
                "data": {
                    "product_name": "Burma Rubies",
                    "hs_code": "7103910000",
                    "quantity": 3000,
                    "price_per_ct": 85.0,
                    "volume_per_pack": 0.038,
                    "weight_per_pack": 0.65,
                    "description": "缅甸红宝石，来自宝石数据库",
                    "source": "Gemstone Database API"
                }
            }
        elif source == "supplier_portal":
            time.sleep(1.8)
            return {
                "success": True,
                "data": {
                    "product_name": "Colombian Emeralds",
                    "hs_code": "7103910000",
                    "quantity": 2000,
                    "price_per_ct": 130.0,
                    "volume_per_pack": 0.05,
                    "weight_per_pack": 0.85,
                    "description": "哥伦比亚祖母绿，来自供应商门户",
                    "source": "Supplier Portal API"
                }
            }
        else:
            return {
                "success": False,
                "error": "未知的数据来源"
            }
    except Exception as e:
        return {
            "success": False,
            "error": str(e)
        }

# 初始化session state
if 'current_product' not in st.session_state:
    st.session_state.current_product = "蓝宝石 (Sapphires)"
if 'quote_history' not in st.session_state:
    st.session_state.quote_history = []
if 'dark_mode' not in st.session_state:
    st.session_state.dark_mode = False
if 'selected_currency' not in st.session_state:
    st.session_state.selected_currency = "USD"
if 'customer_data_fetched' not in st.session_state:
    st.session_state.customer_data_fetched = False
if 'product_data_fetched' not in st.session_state:
    st.session_state.product_data_fetched = False
if 'fetch_source' not in st.session_state:
    st.session_state.fetch_source = "power_automate"

# ==================== 侧边栏 ====================
with st.sidebar:
    st.markdown("### ⚙️ 系统设置")
    
    # 公司信息
    with st.expander("🏢 公司信息", expanded=True):
        company_name = st.text_input("公司名称", "ABC International Trading CO. Ltd")
        company_phone = st.text_input("联系电话", "+86 21 1234 5678")
        company_email = st.text_input("联系邮箱", "info@abctrading.com")
        company_website = st.text_input("网站", "www.abctrading.com")
    
    # 数据抓取设置
    with st.expander("🔄 数据抓取设置", expanded=True):
        st.markdown("**Power Automate Desktop 设置**")
        st.session_state.fetch_source = st.selectbox(
            "数据来源",
            ["power_automate", "alibaba", "made_in_china", "gemstone_database", "supplier_portal"],
            format_func=lambda x: {
                "power_automate": "Power Automate Desktop",
                "alibaba": "阿里巴巴国际站",
                "made_in_china": "Made-in-China.com",
                "gemstone_database": "宝石数据库",
                "supplier_portal": "供应商门户"
            }.get(x, x)
        )
        
        power_automate_path = st.text_input("Power Automate 脚本路径", "C:\\PowerAutomate\\scripts\\fetch_data.ps1")
        api_endpoint = st.text_input("API 端点", "http://localhost:5000/api/fetch-data")
        auto_refresh = st.checkbox("自动刷新数据", value=False)
        
        if auto_refresh:
            refresh_interval = st.slider("刷新间隔(秒)", 30, 300, 60)
    
    # 汇率设置
    with st.expander("💱 汇率设置", expanded=True):
        base_currency = st.selectbox("基础货币", ["CNY"], disabled=True)
        target_currency = st.selectbox("报价货币", list(exchange_rates.keys()), 
                                      index=list(exchange_rates.keys()).index(st.session_state.selected_currency))
        st.session_state.selected_currency = target_currency
        
        # 显示当前汇率
        current_rate = exchange_rates[target_currency]
        st.info(f"1 {target_currency} = {current_rate:.2f} CNY")
        
        # 手动调整汇率
        if st.checkbox("手动调整汇率"):
            manual_rate = st.number_input(f"1 {target_currency} =  CNY", 
                                         value=current_rate, step=0.01, format="%.2f")
            exchange_rates[target_currency] = manual_rate
    
    # 报价设置
    with st.expander("📊 报价设置", expanded=True):
        profit_margin = st.slider("默认利润率 (%)", min_value=5, max_value=100, value=20, step=5)
        tax_rate = st.slider("出口退税率 (%)", min_value=0, max_value=17, value=13, step=1)
        incoterms = st.selectbox("贸易术语", ["FOB", "CIF", "CFR", "EXW", "DDP"])
        
        # 附加费用
        st.markdown("**附加费用 (CNY)**")
        col_s1, col_s2 = st.columns(2)
        with col_s1:
            handling_fee = st.number_input("操作费", value=100, step=10)
            inspection_fee = st.number_input("商检费", value=200, step=10)
        with col_s2:
            document_fee = st.number_input("文件费", value=300, step=10)
            insurance_rate = st.number_input("保险费率 (%)", value=0.3, step=0.1, format="%.1f")
    
    # 快速操作
    st.markdown("---")
    st.markdown("### 🚀 快速操作")
    
    col_btn1, col_btn2 = st.columns(2)
    with col_btn1:
        if st.button("📋 新建报价", use_container_width=True):
            st.session_state.current_product = "蓝宝石 (Sapphires)"
            st.session_state.customer_data_fetched = False
            st.session_state.product_data_fetched = False
            st.success("已创建新报价！")
    
    with col_btn2:
        if st.button("💾 保存模板", use_container_width=True):
            st.success("模板保存成功！")
    
    # 报价历史
    with st.expander("📜 报价历史", expanded=False):
        if st.session_state.quote_history:
            for i, quote in enumerate(st.session_state.quote_history[-5:]):  # 显示最近5条
                st.markdown(f"**{quote['date']}** - {quote['product']}")
                st.markdown(f"客户: {quote['customer']} | 金额: {quote['amount']}")
                st.markdown("---")
        else:
            st.info("暂无报价历史")
        
        if st.button("清空历史", use_container_width=True):
            st.session_state.quote_history = []
            st.rerun()
    
    # 帮助信息
    st.markdown("---")
    with st.expander("❓ 使用帮助", expanded=False):
        st.markdown("""
        **快速开始：**
        1. 点击"抓取客户信息"按钮从外部网站获取客户数据
        2. 点击"抓取商品信息"按钮从外部网站获取商品数据
        3. 或手动输入客户和商品信息
        4. 调整报价参数
        5. 点击开始计算
        
        **Power Automate集成：**
        - 支持从Power Automate Desktop抓取数据
        - 支持阿里巴巴国际站、Made-in-China等平台
        - 可配置自动刷新
        
        **快捷键：**
        - Ctrl+N: 新建报价
        - Ctrl+S: 保存模板
        - Ctrl+P: 打印报价单
        """)
    
    # 系统信息
    st.markdown("---")
    st.markdown(f"**当前时间:** {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    st.markdown(f"**版本:** v2.1.0")
    st.markdown(f"**用户:** 管理员")

# ==================== 主界面 ====================

# 头部
st.markdown("""
<div class="main-header">
    <h1 style="margin:0;">💰 AI价到 - 小微外贸智能报价助手</h1>
    <p style="margin:0.5rem 0 0 0;">{}</p>
</div>
""".format(company_name), unsafe_allow_html=True)

# 显示当前汇率信息和抓取状态
col_status1, col_status2, col_status3 = st.columns(3)
with col_status1:
    st.info(f"💰 当前报价货币: {st.session_state.selected_currency} | 汇率: 1 {st.session_state.selected_currency} = {exchange_rates[st.session_state.selected_currency]:.2f} CNY")
with col_status2:
    if st.session_state.customer_data_fetched:
        st.success("✅ 客户数据已从外部网站抓取")
    else:
        st.warning("⏳ 客户数据待抓取")
with col_status3:
    if st.session_state.product_data_fetched:
        st.success("✅ 商品数据已从外部网站抓取")
    else:
        st.warning("⏳ 商品数据待抓取")

# ==================== 客户信息和商品信息左右并列 ====================
col_left, col_right = st.columns(2, gap="large")

# 左侧：客户信息
with col_left:
    st.markdown("""
    <div class="section-header">
        📋 客户信息
    </div>
    """, unsafe_allow_html=True)
    
    # 抓取客户信息按钮
    col_fetch1, col_fetch2 = st.columns([3, 1])
    with col_fetch1:
        st.markdown(f"**数据来源:** {st.session_state.fetch_source}")
    with col_fetch2:
        if st.button("🔄 抓取客户信息", key="fetch_customer", use_container_width=True):
            with st.spinner("正在从外部网站抓取客户数据..."):
                result = fetch_customer_data_from_api(st.session_state.fetch_source)
                if result["success"]:
                    data = result["data"]
                    # 更新session state
                    st.session_state.customer_name = data["customer_name"]
                    st.session_state.customer_rep = data["customer_rep"]
                    st.session_state.customer_country = data["customer_country"]
                    st.session_state.customer_email = data["customer_email"]
                    st.session_state.customer_address = data["customer_address"]
                    st.session_state.payment_terms = data["payment_terms"]
                    st.session_state.customer_data_fetched = True
                    st.session_state.customer_source = data.get("source", "未知来源")
                    st.success(f"✅ 成功从{data.get('source', '外部网站')}抓取客户数据！")
                    time.sleep(1)
                    st.rerun()
                else:
                    st.error(f"❌ 抓取失败: {result.get('error', '未知错误')}")
    
    # 客户信息输入框
    customer = st.text_input("客户名称", 
                            value=st.session_state.get("customer_name", "Antonia Continental Commerce Ltd."), 
                            key="customer_name_input")
    rep = st.text_input("客户代表", 
                       value=st.session_state.get("customer_rep", "Alfredo Mariani"), 
                       key="customer_rep_input")
    
    # 国家选择
    country_index = 0
    if "customer_country" in st.session_state:
        countries = list(country_port_map.keys())
        if st.session_state.customer_country in countries:
            country_index = countries.index(st.session_state.customer_country)
    
    country = st.selectbox("目的国家", list(country_port_map.keys()), 
                          index=country_index, key="customer_country_input")
    port = country_port_map.get(country, "San Antonio")
    st.text_input("目的港口", value=port, disabled=True, key="customer_port_input")
    
    email = st.text_input("邮箱", 
                         value=st.session_state.get("customer_email", "16203962@yahoo.com"), 
                         key="customer_email_input")
    address = st.text_area("公司地址", 
                          value=st.session_state.get("customer_address", "4 Talcahuano Court, Talcahuano, Chile"), 
                          key="customer_address_input", height=100)
    
    payment_index = 0
    payment_options = ["T/T 30% deposit", "L/C at sight", "D/P", "D/A", "T/T 100% in advance"]
    if "payment_terms" in st.session_state:
        if st.session_state.payment_terms in payment_options:
            payment_index = payment_options.index(st.session_state.payment_terms)
    
    payment_terms = st.selectbox("付款方式", payment_options, 
                                index=payment_index, key="payment_terms_input")
    
    # 显示抓取来源信息
    if st.session_state.get("customer_data_fetched", False):
        st.info(f"📌 数据来源: {st.session_state.get('customer_source', '外部网站')} | 抓取时间: {datetime.now().strftime('%H:%M:%S')}")

# 右侧：商品信息
with col_right:
    st.markdown("""
    <div class="section-header">
        💎 商品信息
    </div>
    """, unsafe_allow_html=True)
    
    # 抓取商品信息按钮
    col_fetch3, col_fetch4 = st.columns([3, 1])
    with col_fetch3:
        st.markdown(f"**数据来源:** {st.session_state.fetch_source}")
    with col_fetch4:
        if st.button("🔄 抓取商品信息", key="fetch_product", use_container_width=True):
            with st.spinner("正在从外部网站抓取商品数据..."):
                result = fetch_product_data_from_api(st.session_state.fetch_source)
                if result["success"]:
                    data = result["data"]
                    # 更新session state
                    st.session_state.fetched_product_name = data["product_name"]
                    st.session_state.fetched_hs_code = data["hs_code"]
                    st.session_state.fetched_quantity = data["quantity"]
                    st.session_state.fetched_price = data["price_per_ct"]
                    st.session_state.fetched_volume = data["volume_per_pack"]
                    st.session_state.fetched_weight = data["weight_per_pack"]
                    st.session_state.fetched_description = data.get("description", "")
                    st.session_state.product_data_fetched = True
                    st.session_state.product_source = data.get("source", "未知来源")
                    
                    # 设置当前商品为自定义以便显示抓取的数据
                    st.session_state.current_product = "自定义商品"
                    
                    st.success(f"✅ 成功从{data.get('source', '外部网站')}抓取商品数据！")
                    time.sleep(1)
                    st.rerun()
                else:
                    st.error(f"❌ 抓取失败: {result.get('error', '未知错误')}")
    
    # 商品快速切换
    st.markdown("**快速选择商品：**")
    
    # 按类别显示商品
    categories = {}
    for product_name, product_data in product_presets.items():
        category = product_data.get("category", "其他")
        if category not in categories:
            categories[category] = []
        categories[category].append(product_name)
    
    # 创建商品选择按钮
    for category, products in categories.items():
        if category != "其他":  # 不显示"其他"类别在快速选择中
            st.markdown(f"**{category}**")
            # 每行最多显示4个按钮
            for i in range(0, len(products), 4):
                cols = st.columns(min(4, len(products) - i))
                for j, product_name in enumerate(products[i:i+4]):
                    with cols[j]:
                        if st.button(product_name.split()[0], key=f"btn_{product_name}_{i}_{j}", use_container_width=True):
                            st.session_state.current_product = product_name
                            st.rerun()
    
    st.markdown("---")
    
    # 根据选择的商品加载预设值
    selected_preset = product_presets[st.session_state.current_product]
    
    # 商品详细信息
    if st.session_state.current_product == "自定义商品":
        # 如果是从外部抓取的数据，使用抓取的值
        product_name = st.text_input("商品名称", 
                                    value=st.session_state.get("fetched_product_name", "请输入商品名称"), 
                                    key="product_name_custom")
        hs_code = st.text_input("HS编码", 
                               value=st.session_state.get("fetched_hs_code", ""), 
                               key="hs_code_custom")
    else:
        product_name = st.text_input("商品名称", st.session_state.current_product, disabled=True, key="product_name_preset")
        hs_code = st.text_input("HS编码", selected_preset["hs_code"], key="hs_code_preset")
    
    # 数量和单价
    col_q1, col_q2 = st.columns(2)
    with col_q1:
        if st.session_state.current_product == "自定义商品" and "fetched_quantity" in st.session_state:
            quantity = st.number_input("数量 (克拉)", 
                                      value=st.session_state.fetched_quantity, 
                                      step=100, key="quantity_custom")
        else:
            quantity = st.number_input("数量 (克拉)", value=5000, step=100, key="quantity_default")
    
    with col_q2:
        if st.session_state.current_product == "自定义商品":
            if "fetched_price" in st.session_state:
                price_per_ct = st.number_input(f"采购单价 (￥/克拉)", 
                                              value=st.session_state.fetched_price, 
                                              step=1.0, key="price_custom")
            else:
                price_per_ct = st.number_input(f"采购单价 (￥/克拉)", value=0.0, step=1.0, key="price_custom_default")
        else:
            price_per_ct = st.number_input(f"采购单价 (￥/克拉)", 
                                          value=selected_preset["price_per_ct"], 
                                          step=1.0, key="price_preset")
    
    # 包装信息
    st.markdown("**包装信息**")
    col_p1, col_p2 = st.columns(2)
    with col_p1:
        if st.session_state.current_product == "自定义商品":
            if "fetched_volume" in st.session_state:
                volume_per_pack = st.number_input("单箱体积 (CBM)", 
                                                 value=st.session_state.fetched_volume, 
                                                 format="%.3f", key="volume_custom")
            else:
                volume_per_pack = st.number_input("单箱体积 (CBM)", value=0.0, format="%.3f", key="volume_custom_default")
        else:
            volume_per_pack = st.number_input("单箱体积 (CBM)", 
                                             value=selected_preset["volume_per_pack"], 
                                             format="%.3f", key="volume_preset")
    
    with col_p2:
        if st.session_state.current_product == "自定义商品":
            if "fetched_weight" in st.session_state:
                weight_per_pack = st.number_input("单箱毛重 (KG)", 
                                                 value=st.session_state.fetched_weight, 
                                                 format="%.2f", key="weight_custom")
            else:
                weight_per_pack = st.number_input("单箱毛重 (KG)", value=0.0, format="%.2f", key="weight_custom_default")
        else:
            weight_per_pack = st.number_input("单箱毛重 (KG)", 
                                             value=selected_preset["weight_per_pack"], 
                                             format="%.2f", key="weight_preset")
    
    # 显示商品描述
    if st.session_state.current_product != "自定义商品":
        st.info(f"📝 {selected_preset['description']}")
    elif st.session_state.get("fetched_description", ""):
        st.info(f"📝 {st.session_state.fetched_description}")
    
    # 显示抓取来源信息
    if st.session_state.get("product_data_fetched", False):
        st.info(f"📌 数据来源: {st.session_state.get('product_source', '外部网站')} | 抓取时间: {datetime.now().strftime('%H:%M:%S')}")
    
    # 添加自定义商品按钮
    if st.session_state.current_product != "自定义商品":
        if st.button("➕ 添加新商品预设", use_container_width=True):
            st.session_state.current_product = "自定义商品"
            st.rerun()

# ==================== 计算结果区域 ====================
st.markdown("---")
st.markdown("### 📊 报价计算结果")

# 计算按钮和结果
col_btn1, col_btn2, col_btn3 = st.columns([1, 1, 2])
with col_btn1:
    if st.button("开始计算", type="primary", use_container_width=True):
        # 计算报价
        total_cost_cny = quantity * price_per_ct
        total_cost_target = total_cost_cny / exchange_rates[st.session_state.selected_currency]
        
        # 保存到历史
        quote_record = {
            "date": datetime.now().strftime("%Y-%m-%d %H:%M"),
            "product": product_name,
            "customer": customer,
            "amount": f"{total_cost_target:,.2f} {st.session_state.selected_currency}",
            "fetched": st.session_state.customer_data_fetched or st.session_state.product_data_fetched
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
    total_packages = math.ceil(quantity/100) if quantity > 0 else 0  # 假设每箱100克拉
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

# 显示抓取数据汇总
if st.session_state.customer_data_fetched or st.session_state.product_data_fetched:
    st.markdown("---")
    st.markdown("### 📋 抓取数据汇总")
    col_sum1, col_sum2 = st.columns(2)
    with col_sum1:
        if st.session_state.customer_data_fetched:
            st.success(f"✅ 客户数据已抓取: {st.session_state.get('customer_name', '')}")
    with col_sum2:
        if st.session_state.product_data_fetched:
            st.success(f"✅ 商品数据已抓取: {st.session_state.get('fetched_product_name', '')}")

st.markdown("---")
st.markdown(f"© 2026 {company_name} | 技术支持: AI价到团队 | Power Automate Desktop 集成")





