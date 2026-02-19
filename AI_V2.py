import streamlit as st
import pandas as pd
import math

# 页面配置
st.set_page_config(
    page_title="AI价到 - 小微外贸智能报价助手",
    page_icon="💰",
    layout="wide"
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
</style>
""", unsafe_allow_html=True)

# 头部
st.markdown("""
<div class="main-header">
    <h1 style="margin:0;">💰 AI价到 - 小微外贸智能报价助手</h1>
    <p style="margin:0.5rem 0 0 0;">ABC International Trading CO. Ltd</p>
</div>
""", unsafe_allow_html=True)

# 国家港口映射
country_port_map = {
    "Chile": "San Antonio", "USA": "Los Angeles", "Germany": "Hamburg",
    "Philippines": "Manila", "China": "Shanghai"
}

# 商品预设数据
product_presets = {
    "蓝宝石 (Sapphires)": {
        "hs_code": "7103910000",
        "price_per_ct": 50.0,
        "volume_per_pack": 0.04,
        "weight_per_pack": 0.7,
        "description": "天然蓝宝石，优质切割"
    },
    "红宝石 (Rubies)": {
        "hs_code": "7103910000",
        "price_per_ct": 80.0,
        "volume_per_pack": 0.035,
        "weight_per_pack": 0.6,
        "description": "缅甸红宝石，色泽鲜艳"
    },
    "祖母绿 (Emeralds)": {
        "hs_code": "7103910000",
        "price_per_ct": 120.0,
        "volume_per_pack": 0.045,
        "weight_per_pack": 0.8,
        "description": "哥伦比亚祖母绿，高净度"
    },
    "钻石 (Diamonds)": {
        "hs_code": "7102390000",
        "price_per_ct": 500.0,
        "volume_per_pack": 0.03,
        "weight_per_pack": 0.5,
        "description": "天然钻石，1克拉以上"
    },
    "水晶 (Crystals)": {
        "hs_code": "7103999000",
        "price_per_ct": 15.0,
        "volume_per_pack": 0.05,
        "weight_per_pack": 1.0,
        "description": "天然水晶，多种颜色"
    },
    "自定义商品": {
        "hs_code": "",
        "price_per_ct": 0.0,
        "volume_per_pack": 0.0,
        "weight_per_pack": 0.0,
        "description": "手动输入商品信息"
    }
}

# 初始化session state
if 'current_product' not in st.session_state:
    st.session_state.current_product = "蓝宝石 (Sapphires)"

# 主界面
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

# 商品信息
st.markdown("### 💎 商品信息")

# 商品快速切换
st.markdown("**快速选择商品：**")
product_cols = st.columns(6)
for i, (product_name, product_data) in enumerate(product_presets.items()):
    with product_cols[i % 6]:
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
        price_per_ct = st.number_input("采购单价 (￥/克拉)", value=0.0, step=1.0)
        volume_per_pack = st.number_input("单箱体积 (CBM)", value=0.0, format="%.3f")
        weight_per_pack = st.number_input("单箱毛重 (KG)", value=0.0, format="%.2f")
    else:
        price_per_ct = st.number_input("采购单价 (￥/克拉)", value=selected_preset["price_per_ct"], step=1.0)
        volume_per_pack = st.number_input("单箱体积 (CBM)", value=selected_preset["volume_per_pack"], format="%.3f")
        weight_per_pack = st.number_input("单箱毛重 (KG)", value=selected_preset["weight_per_pack"], format="%.2f")
    
    # 显示商品描述（如果不是自定义商品）
    if st.session_state.current_product != "自定义商品":
        st.info(f"📝 {selected_preset['description']}")

# 添加自定义商品按钮
if st.session_state.current_product != "自定义商品":
    if st.button("➕ 添加新商品预设", use_container_width=True):
        st.session_state.current_product = "自定义商品"
        st.rerun()

# 计算按钮
col5, col6 = st.columns([1, 5])
with col5:
    if st.button("开始计算", type="primary", use_container_width=True):
        st.success("计算完成！")
        st.balloons()

st.markdown("---")

# 显示当前选择的商品信息
st.markdown(f"**当前选择：** {st.session_state.current_product} | **数量：** {quantity} 克拉 | **总价：** ￥{quantity * price_per_ct:,.2f}")

st.markdown("---")
st.markdown("© 2026 ABC International Trading CO. Ltd")


