import streamlit as st
import pandas as pd
import math
from datetime import datetime, timedelta, timezone
import os

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
    .section-header {
        background-color: #f0f2f6;
        padding: 0.5rem 1rem;
        border-radius: 8px;
        margin-bottom: 1rem;
        font-weight: bold;
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
    
    rates = DEFAULT_RATES.copy()  # 先用默认值
    rate_info = {
        "rates": rates,
        "publish_time": "未知",
        "fetch_time": "未知",
        "file_exists": False,
        "file_time": None
    }
    
    if os.path.exists(excel_path):
        try:
            # 读取Excel文件
            df = pd.read_excel(excel_path)
            
            # 检查文件修改时间
            mod_time = os.path.getmtime(excel_path)
            mod_time_beijing = datetime.fromtimestamp(mod_time) + timedelta(hours=8)
            rate_info["file_time"] = mod_time_beijing.strftime('%Y-%m-%d %H:%M:%S')
            rate_info["file_exists"] = True
            
            # 遍历每一行，提取汇率
            for index, row in df.iterrows():
                currency_code = str(row.iloc[0]).strip()  # A列：货币代码
                rate_value = row.iloc[2]  # C列：汇率
                
                if currency_code in rates:
                    try:
                        rates[currency_code] = float(rate_value)
                    except:
                        pass
            
            # 提取牌价时间和抓取时间（从第一行）
            if len(df) > 0:
                rate_info["publish_time"] = str(df.iloc[0, 3]) if pd.notna(df.iloc[0, 3]) else "未知"
                rate_info["fetch_time"] = str(df.iloc[0, 4]) if pd.notna(df.iloc[0, 4]) else "未知"
            
            rate_info["rates"] = rates
            
        except Exception as e:
            st.error(f"读取汇率文件时出错: {e}")
    
    return rate_info

# 初始化session state
if 'current_product' not in st.session_state:
    st.session_state.current_product = "蓝宝石 (Sapphires)"
if 'quote_history' not in st.session_state:
    st.session_state.quote_history = []
if 'selected_currency' not in st.session_state:
    st.session_state.selected_currency = "USD"
if 'customer_data_fetched' not in st.session_state:
    st.session_state.customer_data_fetched = False
if 'product_data_fetched' not in st.session_state:
    st.session_state.product_data_fetched = False

# 加载汇率数据
rate_info = load_rates_from_excel()
exchange_rates = rate_info["rates"]

# ==================== 侧边栏 ====================
with st.sidebar:
    st.markdown("### ⚙️ 系统设置")
    
    # 公司信息
    with st.expander("🏢 公司信息", expanded=True):
        company_name = st.text_input("公司名称", "ABC International Trading CO. Ltd")
        company_phone = st.text_input("联系电话", "+86 21 1234 5678")
        company_email = st.text_input("联系邮箱", "info@abctrading.com")
        company_website = st.text_input("网站", "www.abctrading.com")
    
    # 汇率设置（从Excel读取）
    with st.expander("💱 汇率设置（中国银行牌价）", expanded=True):
        # 显示数据源状态
        if rate_info["file_exists"]:
            st.success(f"✅ 已连接PAD数据源")
            st.info(f"📁 文件更新时间: {rate_info['file_time']}")
        else:
            st.warning("⚠️ 使用默认汇率数据\n请运行PAD流程生成汇率文件")
            st.info("📁 期望路径: C:\\ExchangeRates\\rates.xlsx")
        
        # 选择报价货币
        available_currencies = list(exchange_rates.keys())
        target_currency = st.selectbox("报价货币", available_currencies, 
                                      index=available_currencies.index(st.session_state.selected_currency) 
                                      if st.session_state.selected_currency in available_currencies else 0)
        st.session_state.selected_currency = target_currency
        
        # 显示当前汇率
        current_rate = exchange_rates[target_currency]
        
        # 显示汇率详细信息
        st.markdown("---")
        st.markdown("**当前汇率信息**")
        st.markdown(f"💰 **1 {target_currency}** = **{current_rate:.4f} CNY**")
        
        if rate_info["publish_time"] != "未知":
            st.markdown(f"📅 **中国银行牌价时间:** {rate_info['publish_time']}")
        if rate_info["fetch_time"] != "未知":
            st.markdown(f"🔄 **PAD抓取时间:** {rate_info['fetch_time']}")
        
        # 手动刷新按钮
        if st.button("🔄 手动刷新汇率", use_container_width=True):
            st.rerun()
        
        # 显示所有汇率
        with st.expander("查看所有汇率"):
            for currency, rate in exchange_rates.items():
                st.text(f"{currency}: {rate:.4f}")
    
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
            for i, quote in enumerate(st.session_state.quote_history[-5:]):
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
        **Power Automate Desktop 集成说明：**
        
        1. **运行PAD流程**：每天9:00和14:00自动运行
        2. **手动运行**：在PAD中点击"运行"
        3. **数据文件**：C:\\ExchangeRates\\rates.xlsx
        4. **汇率时效**：显示中国银行实时牌价
        
        **报价操作：**
        1. 填写客户信息
        2. 选择商品
        3. 调整利润率
        4. 点击"开始计算"
        """)
    
    # 系统信息
    st.markdown("---")
    st.markdown(f"**当前时间:** {format_beijing_time()} (北京时间)")
    st.markdown(f"**版本:** v2.2.0 (PAD集成版)")
    st.markdown(f"**数据源:** {'PAD实时抓取' if rate_info['file_exists'] else '默认数据'}")

# ==================== 主界面 ====================

# 头部
st.markdown("""
<div class="main-header">
    <h1 style="margin:0;">💰 AI价到 - 小微外贸智能报价助手</h1>
    <p style="margin:0.5rem 0 0 0;">{}</p>
</div>
""".format(company_name), unsafe_allow_html=True)

# 显示汇率状态
if rate_info["file_exists"]:
    st.success(f"✅ 当前使用中国银行实时汇率 | 数据时间: {rate_info['publish_time']} | PAD抓取: {rate_info['fetch_time']}")
else:
    st.warning("⚠️ 使用默认汇率数据，请运行Power Automate Desktop流程获取实时汇率")

# ==================== 客户信息和商品信息左右并列 ====================
col_left, col_right = st.columns(2, gap="large")

# 左侧：客户信息
with col_left:
    st.markdown("""
    <div class="section-header">
        📋 客户信息
    </div>
    """, unsafe_allow_html=True)
    
    customer = st.text_input("客户名称", "Antonia Continental Commerce Ltd.", key="customer_name")
    rep = st.text_input("客户代表", "Alfredo Mariani", key="customer_rep")
    country = st.selectbox("目的国家", list(country_port_map.keys()), index=0, key="customer_country")
    port = country_port_map.get(country, "San Antonio")
    st.text_input("目的港口", value=port, disabled=True, key="customer_port")
    email = st.text_input("邮箱", "16203962@yahoo.com", key="customer_email")
    address = st.text_area("公司地址", "4 Talcahuano Court, Talcahuano, Chile", key="customer_address", height=100)
    payment_terms = st.selectbox("付款方式", ["T/T 30% deposit", "L/C at sight", "D/P", "D/A", "T/T 100% in advance"], key="payment_terms")

# 右侧：商品信息
with col_right:
    st.markdown("""
    <div class="section-header">
        💎 商品信息
    </div>
    """, unsafe_allow_html=True)
    
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
        if category != "其他":
            st.markdown(f"**{category}**")
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
        product_name = st.text_input("商品名称", "请输入商品名称", key="product_name_custom")
        hs_code = st.text_input("HS编码", "", key="hs_code_custom")
    else:
        product_name = st.text_input("商品名称", st.session_state.current_product, disabled=True, key="product_name")
        hs_code = st.text_input("HS编码", selected_preset["hs_code"], key="hs_code")
    
    # 数量和单价
    col_q1, col_q2 = st.columns(2)
    with col_q1:
        quantity = st.number_input("数量 (克拉)", value=5000, step=100, key="quantity")
    with col_q2:
        if st.session_state.current_product == "自定义商品":
            price_per_ct = st.number_input(f"采购单价 (￥/克拉)", value=0.0, step=1.0, key="price_custom")
        else:
            price_per_ct = st.number_input(f"采购单价 (￥/克拉)", value=selected_preset["price_per_ct"], step=1.0, key="price")
    
    # 包装信息
    st.markdown("**包装信息**")
    col_p1, col_p2 = st.columns(2)
    with col_p1:
        if st.session_state.current_product == "自定义商品":
            volume_per_pack = st.number_input("单箱体积 (CBM)", value=0.0, format="%.3f", key="volume_custom")
        else:
            volume_per_pack = st.number_input("单箱体积 (CBM)", value=selected_preset["volume_per_pack"], format="%.3f", key="volume")
    with col_p2:
        if st.session_state.current_product == "自定义商品":
            weight_per_pack = st.number_input("单箱毛重 (KG)", value=0.0, format="%.2f", key="weight_custom")
        else:
            weight_per_pack = st.number_input("单箱毛重 (KG)", value=selected_preset["weight_per_pack"], format="%.2f", key="weight")
    
    # 显示商品描述
    if st.session_state.current_product != "自定义商品":
        st.info(f"📝 {selected_preset['description']}")
    
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
st.markdown(f"© 2026 {company_name} | 技术支持: AI价到团队 | 汇率数据: Power Automate Desktop 抓取自中国银行")







