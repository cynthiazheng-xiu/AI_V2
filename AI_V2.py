import streamlit as st
import pandas as pd
import math
from datetime import datetime

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

# 初始化session state
if 'current_product' not in st.session_state:
    st.session_state.current_product = "蓝宝石 (Sapphires)"
if 'quote_history' not in st.session_state:
    st.session_state.quote_history = []
if 'dark_mode' not in st.session_state:
    st.session_state.dark_mode = False
if 'selected_currency' not in st.session_state:
    st.session_state.selected_currency = "USD"

# ==================== 侧边栏 ====================
with st.sidebar:
    st.markdown("### ⚙️ 系统设置")
    
    # 公司信息
    with st.expander("🏢 公司信息", expanded=True):
        company_name = st.text_input("公司名称", "ABC International Trading CO. Ltd")
        company_phone = st.text_input("联系电话", "+86 21 1234 5678")
        company_email = st.text_input("联系邮箱", "info@abctrading.com")
        company_website = st.text_input("网站", "www.abctrading.com")
    
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
        1. 选择或输入商品信息
        2. 填写客户资料
        3. 调整报价参数
        4. 点击开始计算
        
        **快捷键：**
        - Ctrl+N: 新建报价
        - Ctrl+S: 保存模板
        - Ctrl+P: 打印报价单
        """)
    
    # 系统信息
    st.markdown("---")
    st.markdown(f"**当前时间:** {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    st.markdown(f"**版本:** v2.0.0")
    st.markdown(f"**用户:** 管理员")

# ==================== 主界面 ====================

# 头部
st.markdown("""
<div class="main-header">
    <h1 style="margin:0;">💰 AI价到 - 小微外贸智能报价助手</h1>
    <p style="margin:0.5rem 0 0 0;">{}</p>
</div>
""".format(company_name), unsafe_allow_html=True)

# 显示当前汇率信息
st.info(f"💰 当前报价货币: {st.session_state.selected_currency} | 汇率: 1 {st.session_state.selected_currency} = {exchange_rates[st.session_state.selected_currency]:.2f} CNY")

# 客户信息
st.markdown("### 📋 客户信息")
col1, col2 = st.columns(2)

with col1:
    customer = st.text_input("客户名称", "Antonia Continental Commerce Ltd.")
    rep = st.text_input("客户代表", "Alfredo Mariani")
    country = st.selectbox("目的国家", list(country_port_map.keys()), index=0)
    port = country_port_map.get(country, "San Antonio")

with col2:
    st.text_input("目的港口", value=port, disabled=True)
    email = st.text_input("邮箱", "16203962@yahoo.com")
    address = st.text_area("公司地址", "4 Talcahuano Court, Talcahuano, Chile")
    payment_terms = st.selectbox("付款方式", ["T/T 30% deposit", "L/C at sight", "D/P", "D/A", "T/T 100% in advance"])

# 商品信息
st.markdown("### 💎 商品信息")

# 商品快速切换
st.markdown("**快速选择商品：**")
# 按类别显示商品
categories = {}
for product_name, product_data in product_presets.items():
    category = product_data.get("category", "其他")
    if category not in categories:
        categories[category] = []
    categories[category].append(product_name)

for category, products in categories.items():
    st.markdown(f"**{category}**")
    cols = st.columns(len(products))
    for i, product_name in enumerate(products):
        with cols[i]:
            if st.button(product_name.split()[0], key=f"btn_{product_name}", use_container_width=True):
                st.session_state.current_product = product_name
                st.rerun()

st.markdown("---")

# 根据选择的商品加载预设值
selected_preset = product_presets[st.session_state.current_product]

col3, col4 = st.columns(2)

with col3:
    if st.session_state.current_product == "自定义商品":
        product_name = st.text_input("商品名称", "请输入商品名称")
        hs_code = st.text_input("HS编码", "")
    else:
        product_name = st.text_input("商品名称", st.session_state.current_product, disabled=True)
        hs_code = st.text_input("HS编码", selected_preset["hs_code"])
    
    quantity = st.number_input("数量 (克拉)", value=5000, step=100)

with col4:
    if st.session_state.current_product == "自定义商品":
        price_per_ct = st.number_input(f"采购单价 (￥/克拉)", value=0.0, step=1.0)
        volume_per_pack = st.number_input("单箱体积 (CBM)", value=0.0, format="%.3f")
        weight_per_pack = st.number_input("单箱毛重 (KG)", value=0.0, format="%.2f")
    else:
        price_per_ct = st.number_input(f"采购单价 (￥/克拉)", value=selected_preset["price_per_ct"], step=1.0)
        volume_per_pack = st.number_input("单箱体积 (CBM)", value=selected_preset["volume_per_pack"], format="%.3f")
        weight_per_pack = st.number_input("单箱毛重 (KG)", value=selected_preset["weight_per_pack"], format="%.2f")
    
    # 显示商品描述
    if st.session_state.current_product != "自定义商品":
        st.info(f"📝 {selected_preset['description']}")

# 添加自定义商品按钮
if st.session_state.current_product != "自定义商品":
    if st.button("➕ 添加新商品预设", use_container_width=True):
        st.session_state.current_product = "自定义商品"
        st.rerun()

# 计算按钮和结果
col5, col6, col7 = st.columns([1, 1, 3])
with col5:
    if st.button("开始计算", type="primary", use_container_width=True):
        # 计算报价
        total_cost_cny = quantity * price_per_ct
        total_cost_target = total_cost_cny / exchange_rates[st.session_state.selected_currency]
        
        # 保存到历史
        quote_record = {
            "date": datetime.now().strftime("%Y-%m-%d %H:%M"),
            "product": product_name,
            "customer": customer,
            "amount": f"{total_cost_target:,.2f} {st.session_state.selected_currency}"
        }
        st.session_state.quote_history.append(quote_record)
        
        st.success("计算完成！")
        st.balloons()

with col6:
    if st.button("📧 发送报价", use_container_width=True):
        st.info("报价单已准备发送！")

st.markdown("---")

# 显示计算结果和当前选择
col_result1, col_result2, col_result3 = st.columns(3)

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

# 显示详细商品信息
st.markdown("### 📦 商品详情")
col_detail1, col_detail2, col_detail3 = st.columns(3)
with col_detail1:
    st.info(f"**HS编码:** {hs_code}")
with col_detail2:
    st.info(f"**总箱数:** {math.ceil(quantity/100):,} 箱")
with col_detail3:
    st.info(f"**总体积:** {quantity/100 * volume_per_pack:.2f} CBM")

st.markdown("---")
st.markdown(f"© 2026 {company_name} | 技术支持: AI价到团队")



