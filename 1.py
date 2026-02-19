import streamlit as st
import pandas as pd
import math

# 页面配置
st.set_page_config(
    page_title="AI价到 - 小微外贸智能报价助手",
    page_icon="??",
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
</style>
""", unsafe_allow_html=True)

# 头部
st.markdown("""
<div class="main-header">
    <h1 style="margin:0;">?? AI价到 - 小微外贸智能报价助手</h1>
    <p style="margin:0.5rem 0 0 0;">ABC International Trading CO. Ltd</p>
</div>
""", unsafe_allow_html=True)

# 国家港口映射
country_port_map = {
    "Chile": "San Antonio", "USA": "Los Angeles", "Germany": "Hamburg",
    "Philippines": "Manila", "China": "Shanghai"
}

# 中国港口
china_ports = ["Shanghai", "Ningbo", "Shenzhen", "Guangzhou"]

# 主界面
st.markdown("### ?? 客户信息")
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
st.markdown("### ?? 商品信息")
col3, col4 = st.columns(2)

with col3:
    product_name = st.text_input("商品名称", "蓝宝石 (Sapphires)")
    hs_code = st.text_input("HS编码", "7103910000")
    quantity = st.number_input("数量 (克拉)", value=5000, step=100)

with col4:
    price_per_ct = st.number_input("采购单价 (￥/克拉)", value=50.0, step=1.0)
    volume_per_pack = st.number_input("单箱体积 (CBM)", value=0.04, format="%.3f")
    weight_per_pack = st.number_input("单箱毛重 (KG)", value=0.7, format="%.2f")

# 计算按钮
if st.button("?? 开始计算", type="primary"):
    st.success("计算完成！")
    st.balloons()

st.markdown("---")
st.markdown("? 2026 ABC International Trading CO. Ltd")