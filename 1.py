import streamlit as st
import pandas as pd
import math
from datetime import datetime

# 页面配置
st.set_page_config(
    page_title="AI价到 - 小微外贸智能报价助手",
    page_icon="🤖",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 自定义CSS样式 - 更适合比赛的清爽专业风格
st.markdown("""
<style>
    /* 全局样式 */
    .main-header {
        background: linear-gradient(135deg, #0A174E 0%, #1D2B5E 100%);
        padding: 1.5rem 2rem;
        border-radius: 20px;
        color: white;
        margin-bottom: 2rem;
        box-shadow: 0 8px 16px rgba(0,0,0,0.1);
        border: 1px solid rgba(255,255,255,0.1);
    }
    .sub-header {
        color: #0A174E;
        font-size: 1.2rem;
        font-weight: 600;
        margin-bottom: 1rem;
        padding-left: 0.5rem;
        border-left: 4px solid #0A174E;
    }
    .card {
        background: white;
        border: 1px solid #E5E9F0;
        border-radius: 16px;
        padding: 1.5rem;
        box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05);
        margin-bottom: 1rem;
        transition: transform 0.2s, box-shadow 0.2s;
    }
    .card:hover {
        box-shadow: 0 10px 15px -3px rgba(0,0,0,0.1);
    }
    .card-title {
        color: #0A174E;
        font-size: 1.1rem;
        font-weight: 600;
        margin-bottom: 1rem;
        display: flex;
        align-items: center;
        gap: 0.5rem;
    }
    .card-title-icon {
        background: #0A174E;
        color: white;
        width: 28px;
        height: 28px;
        border-radius: 8px;
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 1rem;
    }
    .data-source-badge {
        background: #F0F4FF;
        color: #0A174E;
        padding: 0.2rem 1rem;
        border-radius: 50px;
        font-size: 0.75rem;
        font-weight: 600;
        display: inline-block;
        margin-left: 0.5rem;
        border: 1px solid #0A174E20;
    }
    .info-row {
        display: flex;
        justify-content: space-between;
        padding: 0.5rem 0;
        border-bottom: 1px dashed #E5E9F0;
    }
    .info-label {
        color: #64748B;
        font-size: 0.9rem;
    }
    .info-value {
        color: #0A174E;
        font-weight: 600;
        font-size: 0.95rem;
    }
    .profit-card {
        background: linear-gradient(135deg, #059669 0%, #10B981 100%);
        color: white;
        padding: 1.5rem;
        border-radius: 16px;
        text-align: center;
        box-shadow: 0 8px 16px rgba(5,150,105,0.2);
    }
    .metric-value {
        font-size: 2rem;
        font-weight: 700;
        margin: 0.5rem 0;
    }
    .metric-label {
        font-size: 0.9rem;
        opacity: 0.9;
    }
    .stDataFrame {
        border: 1px solid #E5E9F0;
        border-radius: 12px;
        overflow: hidden;
    }
    .button-primary {
        background: #0A174E;
        color: white;
        border: none;
        padding: 0.5rem 2rem;
        border-radius: 10px;
        font-weight: 600;
        transition: all 0.2s;
    }
    .button-primary:hover {
        background: #1D2B5E;
        box-shadow: 0 4px 12px rgba(10,23,78,0.3);
    }
    .rpa-highlight {
        background: #FEF9C3;
        border-left: 4px solid #FBBF24;
        padding: 1rem;
        border-radius: 8px;
        margin-bottom: 1rem;
        font-size: 0.9rem;
        color: #92400E;
    }
</style>
""", unsafe_allow_html=True)

# ==================== 数据准备 ====================

# 国家和港口映射表
country_port_map = {
    "China": "Shanghai", "USA": "Los Angeles", "Germany": "Hamburg",
    "Kazakhstan": "Aktau", "Australia": "Melbourne", "Japan": "Nagoya",
    "Russia": "St.Petersburg", "South Korea": "Busan", "South Africa": "Cape Town",
    "Brazil": "Rio De Janeiro", "Indonesia": "Jakarta", "Malaysia": "Penang",
    "Pakistan": "Karachi", "Philippines": "Manila", "Singapore": "Singapore",
    "Turkey": "Istanbul", "The United Arab Emirates": "Dubai", "Vietnam": "Hochiminh",
    "Belgium": "Antwerp", "Denmark": "Copenhagen", "U.K.": "London",
    "France": "Le Havre", "Ireland": "Dublin", "Italy": "Genoa",
    "Netherlands": "Rotterdam", "Greece": "Thessaloniki", "Portugal": "Lisbon",
    "Spain": "Barcelona", "Austria": "Vienna", "Finland": "Helsinki",
    "Hungary": "Budapest", "Malta": "Malta", "Norway": "Bergen",
    "Poland": "Gdansk", "Sweden": "Gothenburg", "Switzerland": "Zurich",
    "Ecuador": "Guayaquil", "Panama": "Colon", "El Salvador": "San Salvador",
    "Canada": "Toronto", "New Zealand": "Auckland", "Thailand": "Bangkok",
    "Kiribati": "Tarawa", "Chile": "San Antonio", "Cuba": "Havana",
    "Mexico": "Manzanillo", "Brunei": "Bandar Seri Begwan", "India": "Bombay",
    "Palau": "Koror", "Myanmar": "Yangon", "Peru": "Lima",
    "U.A.E.": "Dubai", "Korea": "Busan", "Costa Rica": "San Jose"
}

# 中国装运港选项
china_ports = ["Shanghai", "Ningbo", "Shenzhen", "Guangzhou", "Qingdao", "Tianjin", "Xiamen", "Dalian"]

# 集装箱数据
container_data = {
    "20HQ (28 CBM / 22吨)": {"volume": 28.0, "weight": 22000},
    "40HQ (67.7 CBM / 26吨)": {"volume": 67.7, "weight": 26000},
    "40FQ (69.7 CBM / 29吨)": {"volume": 69.7, "weight": 29000},
}

# ==================== 头部 ====================
st.markdown("""
<div class="main-header">
    <div style="display: flex; justify-content: space-between; align-items: center;">
        <div>
            <h1 style="margin:0; font-size:2.2rem; display: flex; align-items: center; gap: 10px;">
                🤖 AI价到
                <span style="font-size:1rem; background: rgba(255,255,255,0.2); padding: 0.2rem 1rem; border-radius: 50px;">小微外贸智能报价助手</span>
            </h1>
            <p style="margin:0.5rem 0 0 0; opacity:0.9;">ABC International Trading CO. Ltd · 让报价更智能</p>
        </div>
        <div style="text-align: right; background: rgba(255,255,255,0.1); padding: 0.5rem 1.5rem; border-radius: 50px;">
            <span style="font-size:0.8rem; opacity:0.8;">技能大赛 · RPA+AI混合智能</span>
        </div>
    </div>
</div>
""", unsafe_allow_html=True)

# ==================== RPA数据采集提示 ====================
st.markdown("""
<div class="rpa-highlight">
    <div style="display: flex; align-items: center; gap: 10px;">
        <span style="background: #FBBF24; width: 30px; height: 30px; border-radius: 8px; display: flex; align-items: center; justify-content: center; font-weight: 700;">🤖</span>
        <div>
            <strong>RPA自动采集就绪</strong> · 可从阿里巴巴国际站询盘界面和ERP系统自动抓取数据填充下方表单
        </div>
    </div>
</div>
""", unsafe_allow_html=True)

# ==================== 第一部分：数据来源展示 ====================
st.markdown("### 📊 数据来源")

source_col1, source_col2 = st.columns(2)

with source_col1:
    st.markdown("""
    <div class="card">
        <div class="card-title">
            <span class="card-title-icon">📋</span>
            阿里巴巴国际站 · 客户询盘信息
            <span class="data-source-badge">RPA可抓取</span>
        </div>
    """, unsafe_allow_html=True)
    
    # 模拟从RPA抓取的数据展示
    st.markdown("""
    <div class="info-row">
        <span class="info-label">客户代表</span>
        <span class="info-value">Alfredo Mariani</span>
    </div>
    <div class="info-row">
        <span class="info-label">国别</span>
        <span class="info-value">Chile</span>
    </div>
    <div class="info-row">
        <span class="info-label">公司名称</span>
        <span class="info-value">Antonia Continental Commerce Ltd.</span>
    </div>
    <div class="info-row">
        <span class="info-label">邮箱</span>
        <span class="info-value">16203962@yahoo.com</span>
    </div>
    <div class="info-row">
        <span class="info-label">公司地址</span>
        <span class="info-value">4 Talcahuano Court, Talcahuano, Chile</span>
    </div>
    """, unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)

with source_col2:
    st.markdown("""
    <div class="card">
        <div class="card-title">
            <span class="card-title-icon">📦</span>
            ERP系统 · 商品信息
            <span class="data-source-badge">RPA可抓取</span>
        </div>
    """, unsafe_allow_html=True)
    
    st.markdown("""
    <div class="info-row">
        <span class="info-label">商品编号</span>
        <span class="info-value">N003</span>
    </div>
    <div class="info-row">
        <span class="info-label">商品名称</span>
        <span class="info-value">蓝宝石 | Sapphires</span>
    </div>
    <div class="info-row">
        <span class="info-label">HS编码</span>
        <span class="info-value">7103910000</span>
    </div>
    <div class="info-row">
        <span class="info-label">规格型号</span>
        <span class="info-value">已加工，未镶嵌，天然，无等级，刚玉</span>
    </div>
    <div class="info-row">
        <span class="info-label">销售单位</span>
        <span class="info-value">克拉(CT)</span>
    </div>
    <div class="info-row">
        <span class="info-label">包装信息</span>
        <span class="info-value">纸箱 | 1000 CT/箱</span>
    </div>
    """, unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)

# ==================== 第二部分：可编辑表单 ====================
st.markdown("### ✏️ 报价参数录入")

# 创建三列布局
col_input1, col_input2, col_input3 = st.columns([1, 1, 1.2])

with col_input1:
    st.markdown("""
    <div class="card">
        <div class="card-title">
            <span class="card-title-icon">👤</span>
            客户信息
        </div>
    """, unsafe_allow_html=True)
    
    customer_name = st.text_input("客户名称", "Antonia Continental Commerce Ltd.", key="customer")
    rep_name = st.text_input("客户代表", "Alfredo Mariani", key="rep")
    country = st.selectbox(
        "目的国家", 
        options=list(country_port_map.keys()),
        index=list(country_port_map.keys()).index("Chile") if "Chile" in country_port_map else 0,
        key="country"
    )
    # 根据国家自动显示对应港口
    default_port = country_port_map.get(country, "Manila")
    port = st.text_input("目的港口", value=default_port, key="port", disabled=True)
    email = st.text_input("邮箱", "16203962@yahoo.com", key="email")
    address = st.text_area("公司地址", "4 Talcahuano Court, Talcahuano, Chile", key="address", height=70)
    st.markdown('</div>', unsafe_allow_html=True)

with col_input2:
    st.markdown("""
    <div class="card">
        <div class="card-title">
            <span class="card-title-icon">📦</span>
            商品信息
        </div>
    """, unsafe_allow_html=True)
    
    product_code = st.text_input("商品编号", "N003", key="product_code")
    product_name = st.text_input("商品名称", "蓝宝石 (Sapphires)", key="product_name")
    hs_code = st.text_input("HS编码", "7103910000", key="hs_code")
    
    col_unit1, col_unit2 = st.columns(2)
    with col_unit1:
        sale_unit = st.text_input("销售单位", "克拉(CT)", key="sale_unit")
    with col_unit2:
        pack_unit = st.text_input("包装单位", "纸箱(CARTON)", key="pack_unit")
    
    unit_convert = st.number_input("单位换算 (CT/箱)", value=1000, min_value=1, step=100, key="unit_convert")
    
    col_pack1, col_pack2 = st.columns(2)
    with col_pack1:
        gross_weight = st.number_input("毛重 (KG/箱)", value=0.70, min_value=0.01, format="%.2f", key="gross_weight")
    with col_pack2:
        net_weight = st.number_input("净重 (KG/箱)", value=0.20, min_value=0.01, format="%.2f", key="net_weight")
    
    col_vol, col_trans = st.columns(2)
    with col_vol:
        volume_per_pack = st.number_input("体积 (CBM/箱)", value=0.0400, min_value=0.0001, format="%.4f", key="volume_per_pack")
    with col_trans:
        transport_note = st.text_input("运输说明", "无", key="transport_note")
    
    st.markdown('</div>', unsafe_allow_html=True)

with col_input3:
    st.markdown("""
    <div class="card">
        <div class="card-title">
            <span class="card-title-icon">🚢</span>
            运输与报价参数
        </div>
    """, unsafe_allow_html=True)
    
    # 装运港选择（中国港口）
    col_shanghai1, col_shanghai2 = st.columns(2)
    with col_shanghai1:
        loading_port = st.selectbox("装运港 (中国)", options=china_ports, index=0, key="loading_port")
    with col_shanghai2:
        incoterm = st.selectbox("贸易术语", ["FOB", "CIF", "CIP", "EXW", "DAP"], index=0, key="incoterm")
    
    # 数量输入（基于克拉）
    col_qty1, col_qty2 = st.columns(2)
    with col_qty1:
        quantity_ct = st.number_input("数量 (克拉)", value=5000, min_value=1, step=100, key="quantity_ct")
    with col_qty2:
        # 自动计算箱数
        cartons = math.ceil(quantity_ct / unit_convert)
        st.number_input("箱数", value=cartons, disabled=True, key="cartons")
    
    # 单价
    purchase_price_per_ct = st.number_input("采购单价 (¥/克拉)", value=50.0, min_value=0.01, step=10.0, key="price_per_ct")
    purchase_total = purchase_price_per_ct * quantity_ct
    st.metric("采购总价 (¥)", f"¥{purchase_total:,.2f}")
    
    # 体积和重量自动计算
    total_volume = volume_per_pack * cartons
    total_weight = gross_weight * cartons
    
    col_calc1, col_calc2 = st.columns(2)
    with col_calc1:
        st.metric("总体积 (CBM)", f"{total_volume:.3f}")
    with col_calc2:
        st.metric("总毛重 (KG)", f"{total_weight:.2f}")
    
    st.markdown('</div>', unsafe_allow_html=True)

# ==================== 侧边栏参数 ====================
with st.sidebar:
    st.markdown("## ⚙️ 系统参数设置")
    
    with st.expander("💰 汇率设置", expanded=True):
        exchange_rate = st.number_input("美元汇率 (USD/CNY)", value=7.2, step=0.1)
    
    with st.expander("📦 集装箱参数", expanded=True):
        container_type = st.selectbox(
            "选择集装箱类型",
            options=list(container_data.keys()),
            index=1
        )
        container_volume = container_data[container_type]["volume"]
        container_weight = container_data[container_type]["weight"]
        st.info(f"体积: {container_volume} CBM, 限重: {container_weight/1000:.1f}吨")
    
    with st.expander("💰 费用参数", expanded=True):
        domestic_fee_base = st.number_input("国内运费基础 (¥)", value=3000, step=100)
        domestic_fee_per = st.number_input("每柜国内运费 (¥)", value=1500, step=100)
        freight_usd = st.number_input("海运费 (USD/柜)", value=1000, step=50)
    
    with st.expander("📊 税费参数", expanded=True):
        vat_rate = st.number_input("增值税率 (%)", value=13.0, step=0.5) / 100
        tariff_rate = st.number_input("关税率 (%)", value=5.0, step=0.5) / 100
        insurance_rate = st.number_input("保险费率 (%)", value=0.2, step=0.05) / 100
    
    with st.expander("📈 利润目标", expanded=True):
        profit_rate = st.number_input("目标利润率 (%)", value=20.0, step=1.0) / 100

# ==================== 计算按钮 ====================
col_btn1, col_btn2, col_btn3, col_btn4 = st.columns([1, 1, 1, 3])
with col_btn1:
    calculate = st.button("🚀 开始智能计算", type="primary", use_container_width=True)
with col_btn2:
    reset = st.button("🔄 重置表单", use_container_width=True)
with col_btn3:
    rpa_fill = st.button("🤖 RPA填充演示", use_container_width=True)

if rpa_fill:
    st.success("✅ RPA已自动填充客户和商品信息！")
    st.balloons()

# ==================== 计算结果 ====================
if calculate:
    # 集装箱计算
    containers_by_volume = math.ceil(total_volume / container_volume)
    containers_by_weight = math.ceil(total_weight / container_weight)
    containers = max(containers_by_volume, containers_by_weight)
    
    # 费用计算
    tax_rebate = purchase_total / (1 + vat_rate) * vat_rate
    domestic_fee = domestic_fee_base + domestic_fee_per * containers
    intl_freight = freight_usd * exchange_rate * containers
    insurance = (purchase_total + intl_freight) * 1.1 * insurance_rate
    tariff = purchase_total * tariff_rate
    
    # 成本计算
    total_cost = purchase_total - tax_rebate + domestic_fee + intl_freight + insurance + tariff
    
    # 利润计算
    target_profit = total_cost * profit_rate
    contract_amount = total_cost + target_profit
    
    # 单价计算（不同单位）
    unit_price_usd_per_ct = contract_amount / quantity_ct / exchange_rate
    unit_price_usd_per_1000ct = unit_price_usd_per_ct * 1000
    unit_price_usd_per_carton = unit_price_usd_per_ct * unit_convert
    
    # 显示结果
    st.markdown("### 🤖 智能报价结果")
    
    # 结果卡片
    res_col1, res_col2, res_col3 = st.columns(3)
    
    with res_col1:
        st.markdown("""
        <div class="card">
            <div class="card-title">📦 装箱方案</div>
        """, unsafe_allow_html=True)
        st.metric("集装箱类型", container_type)
        st.metric("需要集装箱", f"{containers} 个")
        st.metric("总体积", f"{total_volume:.3f} CBM")
        st.metric("总毛重", f"{total_weight:.2f} KG")
        st.markdown("</div>", unsafe_allow_html=True)
    
    with res_col2:
        st.markdown("""
        <div class="card">
            <div class="card-title">💰 报价建议</div>
        """, unsafe_allow_html=True)
        st.metric("单价 (USD/克拉)", f"USD {unit_price_usd_per_ct:.4f}")
        st.metric("单价 (USD/千克拉)", f"USD {unit_price_usd_per_1000ct:.2f}")
        st.metric("单价 (USD/箱)", f"USD {unit_price_usd_per_carton:.2f}")
        st.metric("合同总额", f"¥{contract_amount:,.2f}")
        st.markdown("</div>", unsafe_allow_html=True)
    
    with res_col3:
        st.markdown("""
        <div class="card">
            <div class="card-title">📊 成本利润</div>
        """, unsafe_allow_html=True)
        st.metric("采购总价", f"¥{purchase_total:,.2f}")
        st.metric("退税金额", f"¥{tax_rebate:,.2f}")
        st.metric("总成本", f"¥{total_cost:,.2f}")
        st.metric("预计利润", f"¥{target_profit:,.2f}")
        st.markdown("</div>", unsafe_allow_html=True)
    
    # 预算表
    st.markdown("### 📊 出口预算表")
    
    budget_data = {
        '项目': ['采购总价', '出口退税', '国内运费', '国际运费', '保险费', '出口关税', '总成本', '合同金额', '预计利润'],
        '金额': [
            f"¥{purchase_total:,.2f}",
            f"-¥{tax_rebate:,.2f}",
            f"¥{domestic_fee:,.2f}",
            f"¥{intl_freight:,.2f}",
            f"¥{insurance:,.2f}",
            f"¥{tariff:,.2f}",
            f"¥{total_cost:,.2f}",
            f"¥{contract_amount:,.2f}",
            f"¥{target_profit:,.2f}"
        ]
    }
    
    df = pd.DataFrame(budget_data)
    st.dataframe(df, use_container_width=True, hide_index=True)
    
    # 利润率卡片
    col_profit1, col_profit2, col_profit3 = st.columns(3)
    
    with col_profit1:
        st.markdown(f"""
        <div class="profit-card">
            <div class="metric-label">💰 预计利润</div>
            <div class="metric-value">¥{target_profit:,.2f}</div>
        </div>
        """, unsafe_allow_html=True)
    
    with col_profit2:
        st.markdown(f"""
        <div class="profit-card">
            <div class="metric-label">📈 利润率</div>
            <div class="metric-value">{profit_rate*100:.1f}%</div>
        </div>
        """, unsafe_allow_html=True)
    
    with col_profit3:
        roi = (target_profit / purchase_total) * 100
        st.markdown(f"""
        <div class="profit-card">
            <div class="metric-label">🔄 投资回报率</div>
            <div class="metric-value">{roi:.1f}%</div>
        </div>
        """, unsafe_allow_html=True)
    
    # 报价单生成
    col_gen1, col_gen2, col_gen3 = st.columns([1, 1, 1])
    with col_gen2:
        if st.button("📄 生成正式报价单", use_container_width=True):
            st.success("✅ 报价单已生成！")
            st.balloons()
            st.info("📧 报价单已发送 (演示版)")

# ==================== RPA说明 ====================
with st.expander("🤖 RPA自动化流程说明（点击展开）"):
    st.markdown("""
    ### Power Automate Desktop 自动化流程
    
    **1. 阿里巴巴国际站抓取**
    - 打开浏览器，进入阿里巴巴询盘页面
    - 提取客户名称、国家、邮箱、地址
    - 自动填入Web表单
    
    **2. ERP系统抓取**
    - 打开ERP系统，根据商品编号查询
    - 提取HS编码、规格、包装信息、重量体积
    - 自动填入Web表单
    
    **3. 自动计算**
    - 触发"开始智能计算"按钮
    - 抓取计算结果
    - 保存到Excel或生成报价单
    """)

# 页脚
st.markdown("---")
st.markdown(
    """
    <div style="text-align: center; color: #666; padding: 1rem;">
        <p>© 2026 ABC International Trading CO. Ltd · AI价到-小微外贸智能报价助手</p>
        <p style="font-size: 0.875rem;">技能大赛作品 · RPA + Streamlit 混合智能</p>
    </div>
    """, 
    unsafe_allow_html=True
)