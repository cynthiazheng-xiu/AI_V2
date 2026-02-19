# app.py - AI价到 - 小微外贸智能报价助手

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
    .empty-field {
        color: #999;
        font-style: italic;
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
    .term-desc {
        font-size: 0.9rem;
        color: #555;
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

# 2020版国际贸易术语完整列表
INCOTERMS_2020 = [
    {
        "code": "EXW",
        "name": "EXW (工厂交货)",
        "full_name": "Ex Works",
        "category": "任一地点",
        "description": "卖方在其所在地或其他指定地点将货物交给买方处置时即完成交货。卖方不负责装货，也不负责出口清关。卖方承担最小责任，买方负责所有运输、保险和进出口清关。",
        "responsibility_seller": "在指定地点提供货物",
        "responsibility_buyer": "所有运输、保险、出口/进口清关、装货",
        "risk_transfer": "卖方将货物交给买方处置时",
        "transport": "任何运输方式"
    },
    {
        "code": "FCA",
        "name": "FCA (货交承运人)",
        "full_name": "Free Carrier",
        "category": "主要运费未付",
        "description": "卖方在指定地点将货物交给买方指定的承运人即完成交货。卖方负责出口清关。如指定地点是卖方所在地，卖方负责装货；如在其他地点，卖方不负责卸货。",
        "responsibility_seller": "出口清关、将货物交给承运人",
        "responsibility_buyer": "主运输、保险、进口清关",
        "risk_transfer": "货物交给承运人时",
        "transport": "任何运输方式"
    },
    {
        "code": "FAS",
        "name": "FAS (船边交货)",
        "full_name": "Free Alongside Ship",
        "category": "主要运费未付",
        "description": "卖方在指定装运港将货物放在船边（例如码头上或驳船上）即完成交货。卖方负责出口清关。适用于海运或内河水运。",
        "responsibility_seller": "出口清关、将货物运至船边",
        "responsibility_buyer": "装船、主运输、保险、进口清关",
        "risk_transfer": "货物放在船边时",
        "transport": "海运和内河水运"
    },
    {
        "code": "FOB",
        "name": "FOB (船上交货)",
        "full_name": "Free On Board",
        "category": "主要运费未付",
        "description": "卖方在指定装运港将货物装到买方指定的船上即完成交货。卖方负责出口清关。风险和费用在货物装上船时转移。适用于海运或内河水运。",
        "responsibility_seller": "出口清关、将货物装上船",
        "responsibility_buyer": "主运输、保险、进口清关",
        "risk_transfer": "货物装上船时",
        "transport": "海运和内河水运"
    },
    {
        "code": "CFR",
        "name": "CFR (成本加运费)",
        "full_name": "Cost and Freight",
        "category": "主要运费已付",
        "description": "卖方支付将货物运至指定目的港的运费。货物在装运港装上船时风险转移给买方。卖方负责出口清关，但不负责保险。适用于海运或内河水运。",
        "responsibility_seller": "出口清关、将货物装上船、支付至目的港运费",
        "responsibility_buyer": "保险、进口清关、目的港卸货费",
        "risk_transfer": "货物装上船时",
        "transport": "海运和内河水运"
    },
    {
        "code": "CIF",
        "name": "CIF (成本、保险费加运费)",
        "full_name": "Cost, Insurance and Freight",
        "category": "主要运费已付",
        "description": "卖方支付将货物运至指定目的港的运费，并必须购买货物运输保险。货物在装运港装上船时风险转移给买方。卖方负责出口清关。适用于海运或内河水运。",
        "responsibility_seller": "出口清关、将货物装上船、支付至目的港运费和保险费",
        "responsibility_buyer": "进口清关、目的港卸货费",
        "risk_transfer": "货物装上船时",
        "transport": "海运和内河水运"
    },
    {
        "code": "CPT",
        "name": "CPT (运费付至)",
        "full_name": "Carriage Paid To",
        "category": "主要运费已付",
        "description": "卖方支付将货物运至指定目的地的运费。货物交给第一承运人时风险转移给买方。卖方负责出口清关，但不负责保险。适用于任何运输方式。",
        "responsibility_seller": "出口清关、将货物交给承运人、支付至目的地运费",
        "responsibility_buyer": "保险、进口清关、目的地卸货费",
        "risk_transfer": "货物交给第一承运人时",
        "transport": "任何运输方式"
    },
    {
        "code": "CIP",
        "name": "CIP (运费、保险费付至)",
        "full_name": "Carriage and Insurance Paid To",
        "category": "主要运费已付",
        "description": "卖方支付将货物运至指定目的地的运费，并必须购买货物运输保险（比CIF要求更高保额）。货物交给第一承运人时风险转移给买方。卖方负责出口清关。适用于任何运输方式。",
        "responsibility_seller": "出口清关、将货物交给承运人、支付至目的地运费和保险费",
        "responsibility_buyer": "进口清关、目的地卸货费",
        "risk_transfer": "货物交给第一承运人时",
        "transport": "任何运输方式"
    },
    {
        "code": "DAP",
        "name": "DAP (目的地交货)",
        "full_name": "Delivered At Place",
        "category": "到达",
        "description": "卖方将货物运至指定目的地，并将货物放在已到达的运输工具上（未卸货）交给买方处置即完成交货。卖方负责出口清关和运输，但不负责卸货和进口清关。",
        "responsibility_seller": "出口清关、运输至指定目的地",
        "responsibility_buyer": "卸货、进口清关",
        "risk_transfer": "货物在目的地交由买方处置时",
        "transport": "任何运输方式"
    },
    {
        "code": "DPU",
        "name": "DPU (卸货地交货)",
        "full_name": "Delivered At Place Unloaded",
        "category": "到达",
        "description": "卖方将货物运至指定目的地并卸货后交给买方处置即完成交货。这是Incoterms 2020中唯一要求卖方卸货的术语。卖方负责出口清关和运输。",
        "responsibility_seller": "出口清关、运输至指定目的地、卸货",
        "responsibility_buyer": "进口清关",
        "risk_transfer": "货物卸下并交由买方处置时",
        "transport": "任何运输方式"
    },
    {
        "code": "DDP",
        "name": "DDP (完税后交货)",
        "full_name": "Delivered Duty Paid",
        "category": "到达",
        "description": "卖方将货物运至指定目的地，并完成进口清关后交给买方处置即完成交货。卖方承担所有风险和费用，包括运输、保险、出口和进口关税。卖方承担最大责任。",
        "responsibility_seller": "所有运输、保险、出口/进口清关、关税支付",
        "responsibility_buyer": "极少责任，只需在目的地接收货物",
        "risk_transfer": "货物在目的地交由买方处置时",
        "transport": "任何运输方式"
    }
]

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

# 启动PAD流程的函数
def run_pad_flow(flow_name):
    """调用Power Automate Desktop运行指定流程"""
    try:
        # PAD的命令行调用方式（需要根据实际安装路径调整）
        pad_path = "C:\\Program Files (x86)\\Power Automate Desktop\\PAD.Console.exe"
        
        if os.path.exists(pad_path):
            # 直接调用PAD控制台
            result = subprocess.run([pad_path, flow_name], capture_output=True, text=True, timeout=10)
            return {"success": True, "message": f"已启动PAD流程: {flow_name}"}
        else:
            # 如果没有PAD，模拟成功（用于演示）
            st.info(f"模拟运行PAD流程: {flow_name}")
            return {"success": True, "message": f"模拟运行成功"}
            
    except Exception as e:
        return {"success": False, "message": str(e)}

# 初始化session state
if 'quote_history' not in st.session_state:
    st.session_state.quote_history = []
if 'selected_currency' not in st.session_state:
    st.session_state.selected_currency = "USD"
if 'customer_data' not in st.session_state:
    st.session_state.customer_data = {}  # 初始为空
if 'product_data' not in st.session_state:
    st.session_state.product_data = {}  # 初始为空
if 'customer_fetched' not in st.session_state:
    st.session_state.customer_fetched = False
if 'product_fetched' not in st.session_state:
    st.session_state.product_fetched = False

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
        
        # 2020版国际贸易术语选择
        term_options = [term["name"] for term in INCOTERMS_2020]
        selected_term = st.selectbox("贸易术语 (Incoterms 2020)", term_options, index=3)  # 默认FOB
        
        # 查找选中的术语详情
        selected_term_detail = next((term for term in INCOTERMS_2020 if term["name"] == selected_term), INCOTERMS_2020[0])
        
        # 显示术语简要说明
        st.info(f"📌 {selected_term_detail['description'][:100]}...")
        
        # 附加费用
        st.markdown("**附加费用 (CNY)**")
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
    
    # 显示抓取状态
    if st.session_state.customer_fetched:
        st.success("✅ 已从PAD抓取客户数据")
    else:
        st.info("⏳ 点击上方'抓取客户信息'按钮从阿里巴巴国际站获取客户数据")
    
    # 如果session中有抓取的数据，使用它；否则留空
    default_customer = st.session_state.customer_data if st.session_state.customer_data else {}
    
    customer = st.text_input("客户名称", 
                            value=default_customer.get("customer_name", ""), 
                            placeholder="例如: Antonia Continental Commerce Ltd.",
                            key="customer_name_input")
    
    rep = st.text_input("客户代表", 
                       value=default_customer.get("customer_rep", ""), 
                       placeholder="例如: Alfredo Mariani",
                       key="customer_rep_input")
    
    # 国家选择
    country_index = 0
    if default_customer.get("customer_country") and default_customer["customer_country"] in country_port_map:
        countries = list(country_port_map.keys())
        country_index = countries.index(default_customer["customer_country"])
    
    country = st.selectbox("目的国家", list(country_port_map.keys()), 
                          index=country_index, key="customer_country_input")
    port = country_port_map.get(country, "San Antonio")
    st.text_input("目的港口", value=port, disabled=True, key="customer_port_input")
    
    email = st.text_input("邮箱", 
                         value=default_customer.get("customer_email", ""), 
                         placeholder="例如: 16203962@yahoo.com",
                         key="customer_email_input")
    
    address = st.text_area("公司地址", 
                          value=default_customer.get("customer_address", ""), 
                          placeholder="例如: 4 Talcahuano Court, Talcahuano, Chile",
                          key="customer_address_input", height=100)
    
    payment_options = ["T/T 30% deposit", "L/C at sight", "D/P", "D/A", "T/T 100% in advance"]
    payment_index = 0
    if default_customer.get("payment_terms") in payment_options:
        payment_index = payment_options.index(default_customer["payment_terms"])
    
    payment_terms = st.selectbox("付款方式", payment_options, 
                                index=payment_index, key="payment_terms_input")
    
    # 显示抓取时间
    if default_customer.get("fetch_time"):
        st.info(f"📌 PAD抓取时间: {default_customer['fetch_time']}")

# 右侧：商品信息
with col_right:
    st.markdown("""
    <div class="section-header">
        💎 商品信息
    </div>
    """, unsafe_allow_html=True)
    
    # 显示抓取状态
    if st.session_state.product_fetched:
        st.success("✅ 已从PAD抓取商品数据")
    else:
        st.info("⏳ 点击上方'抓取商品信息'按钮从本公司ERP系统抓取商品数据，或手动输入")
    
    # 如果session中有抓取的商品数据，使用它
    default_product = st.session_state.product_data if st.session_state.product_data else {}
    
    # 商品编号
    product_code = st.text_input("商品编号", 
                                value=default_product.get("product_code", "N003"), 
                                placeholder="例如: N003",
                                key="product_code_input")
    
    # 货物类型
    goods_type = st.text_input("货物类型", 
                              value=default_product.get("goods_type", "宝石或半宝石"), 
                              placeholder="例如: 宝石或半宝石",
                              key="goods_type_input")
    
    # 商品名称
    product_name = st.text_input("商品名称", 
                                value=default_product.get("product_name", "蓝宝石"), 
                                placeholder="例如: 蓝宝石",
                                key="product_name_input")
    
    # 英文名称
    product_name_en = st.text_input("英文名称", 
                                   value=default_product.get("product_name_en", "Sapphires"), 
                                   placeholder="例如: Sapphires",
                                   key="product_name_en_input")
    
    # 规格型号（中文）
    specification_cn = st.text_input("规格型号（中文）", 
                                    value=default_product.get("specification_cn", "已加工，未镶嵌，天然，无等级，刚玉"), 
                                    placeholder="例如: 已加工，未镶嵌，天然，无等级，刚玉",
                                    key="specification_cn_input")
    
    # 规格型号（英文）
    specification_en = st.text_input("规格型号（英文）", 
                                    value=default_product.get("specification_en", "Processed,not inlaid,natural,no grade,corundum"), 
                                    placeholder="例如: Processed,not inlaid,natural,no grade,corundum",
                                    key="specification_en_input")
    
    # HS编码
    hs_code = st.text_input("HS编码", 
                           value=default_product.get("hs_code", "7103910000"), 
                           placeholder="例如: 7103910000",
                           key="hs_code_input")
    
    # 销售单位和数量
    col_q1, col_q2 = st.columns(2)
    with col_q1:
        sales_unit = st.text_input("销售单位", 
                                  value=default_product.get("sales_unit", "克拉（CT）"), 
                                  placeholder="例如: 克拉（CT）",
                                  key="sales_unit_input")
    
    with col_q2:
        quantity = st.number_input("数量 (克拉)", 
                                  value=default_product.get("quantity", 0), 
                                  step=100, 
                                  min_value=0,
                                  key="quantity_input")
    
    # 采购单价
    price_per_ct = st.number_input("采购单价 (￥/克拉)", 
                                  value=default_product.get("price_per_ct", 0.0), 
                                  step=1.0,
                                  min_value=0.0,
                                  key="price_input")
    
    # 包装信息
    st.markdown("**包装信息**")
    
    col_p1, col_p2 = st.columns(2)
    with col_p1:
        package_unit = st.text_input("包装单位", 
                                    value=default_product.get("package_unit", "纸箱（CARTON）"), 
                                    placeholder="例如: 纸箱（CARTON）",
                                    key="package_unit_input")
    
    with col_p2:
        unit_conversion = st.text_input("单位换算", 
                                       value=default_product.get("unit_conversion", "1000CT/CARTON"), 
                                       placeholder="例如: 1000CT/CARTON",
                                       key="unit_conversion_input")
    
    col_p3, col_p4 = st.columns(2)
    with col_p3:
        gross_weight = st.number_input("毛重 (KGS/纸箱)", 
                                      value=default_product.get("gross_weight", 0.70), 
                                      format="%.2f",
                                      min_value=0.0,
                                      key="gross_weight_input")
    
    with col_p4:
        net_weight = st.number_input("净重 (KGS/纸箱)", 
                                    value=default_product.get("net_weight", 0.20), 
                                    format="%.2f",
                                    min_value=0.0,
                                    key="net_weight_input")
    
    # 体积
    volume_per_pack = st.number_input("体积 (CBM/纸箱)", 
                                     value=default_product.get("volume_per_pack", 0.0400), 
                                     format="%.4f",
                                     min_value=0.0,
                                     key="volume_input")
    
    # 海关信息
    st.markdown("**海关信息**")
    
    col_h1, col_h2 = st.columns(2)
    with col_h1:
        legal_unit = st.text_input("法定单位", 
                                  value=default_product.get("legal_unit", "克拉（CT）"), 
                                  placeholder="例如: 克拉（CT）",
                                  key="legal_unit_input")
    
    with col_h2:
        customs_supervision = st.text_input("海关监管条件", 
                                           value=default_product.get("customs_supervision", "无"), 
                                           placeholder="例如: 无",
                                           key="customs_supervision_input")
    
    inspection_category = st.text_input("检验检疫类别", 
                                       value=default_product.get("inspection_category", "无"), 
                                       placeholder="例如: 无",
                                       key="inspection_category_input")
    
    # 运输说明
    transport_notes = st.text_input("运输说明", 
                                   value=default_product.get("transport_notes", "无"), 
                                   placeholder="例如: 无",
                                   key="transport_notes_input")
    
    # 商品描述（可选）
    description = st.text_area("商品描述", 
                              value=default_product.get("description", ""), 
                              placeholder="请输入补充商品描述",
                              key="description_input", height=80)
    
    # 显示抓取时间
    if default_product.get("fetch_time"):
        st.info(f"📌 PAD抓取时间: {default_product['fetch_time']}")

# ==================== 计算结果区域 ====================
st.markdown("---")
st.markdown("### 📊 报价计算结果")

col_btn1, col_btn2, col_btn3 = st.columns([1, 1, 2])
with col_btn1:
    if st.button("开始计算", type="primary", use_container_width=True):
        if quantity > 0 and price_per_ct > 0:
            total_cost_cny = quantity * price_per_ct
            total_cost_target = total_cost_cny / exchange_rates[st.session_state.selected_currency]
            
            quote_record = {
                "date": format_beijing_time(),
                "product": product_name if product_name else "未命名商品",
                "customer": customer if customer else "未填写客户",
                "amount": f"{total_cost_target:,.2f} {st.session_state.selected_currency}"
            }
            st.session_state.quote_history.append(quote_record)
            
            st.success("计算完成！")
            st.balloons()
        else:
            st.error("请填写商品数量和单价")

with col_btn2:
    if st.button("📧 发送报价", use_container_width=True):
        if customer and product_name and quantity > 0 and price_per_ct > 0:
            st.info("报价单已准备发送！")
        else:
            st.warning("请先填写完整的客户和商品信息")

# 显示计算结果
col_result1, col_result2, col_result3, col_result4 = st.columns(4)

with col_result1:
    total_cost_cny = quantity * price_per_ct
    st.metric("采购总成本 (CNY)", f"￥{total_cost_cny:,.2f}" if quantity > 0 else "￥0.00")

with col_result2:
    total_cost_target = total_cost_cny / exchange_rates[st.session_state.selected_currency] if quantity > 0 else 0
    st.metric(f"采购总成本 ({st.session_state.selected_currency})", 
              f"{total_cost_target:,.2f} {st.session_state.selected_currency}" if quantity > 0 else f"0.00 {st.session_state.selected_currency}")

with col_result3:
    suggested_price = total_cost_target * (1 + profit_margin/100) if quantity > 0 else 0
    st.metric(f"建议报价 ({st.session_state.selected_currency})", 
              f"{suggested_price:,.2f} {st.session_state.selected_currency}" if quantity > 0 else f"0.00 {st.session_state.selected_currency}",
              delta=f"{profit_margin}% 利润率")

with col_result4:
    # 根据单位换算计算总箱数
    if quantity > 0 and unit_conversion:
        try:
            # 解析单位换算，例如 "1000CT/CARTON"
            conversion_parts = unit_conversion.split('/')
            if len(conversion_parts) == 2:
                ct_per_carton = float(conversion_parts[0].replace('CT', '').strip())
                total_packages = math.ceil(quantity / ct_per_carton)
            else:
                total_packages = math.ceil(quantity / 1000)  # 默认1000CT/箱
        except:
            total_packages = math.ceil(quantity / 1000)  # 解析失败时使用默认值
    else:
        total_packages = 0
    st.metric("总箱数", f"{total_packages:,} 箱")

# 显示详细商品信息
if quantity > 0:
    st.markdown("### 📦 商品详情")
    col_detail1, col_detail2, col_detail3, col_detail4 = st.columns(4)
    with col_detail1:
        st.info(f"**HS编码:** {hs_code}" if hs_code else "**HS编码:** 未填写")
    with col_detail2:
        st.info(f"**总箱数:** {total_packages:,} 箱")
    with col_detail3:
        total_volume = total_packages * volume_per_pack if total_packages > 0 else 0
        st.info(f"**总体积:** {total_volume:.2f} CBM")
    with col_detail4:
        total_weight = total_packages * gross_weight if total_packages > 0 else 0
        st.info(f"**总毛重:** {total_weight:.2f} KG")

# 显示贸易术语详情
st.markdown("---")
st.markdown("### 📋 贸易术语详情 - Incoterms 2020")

col_term1, col_term2 = st.columns([1, 2])

with col_term1:
    st.markdown(f"""
    <div class="term-card">
        <div class="term-title">📌 当前选择</div>
        <div style="font-size: 1.2rem; font-weight: bold;">{selected_term_detail['name']}</div>
        <div style="font-style: italic;">{selected_term_detail['full_name']}</div>
    </div>
    """, unsafe_allow_html=True)
    
    st.markdown(f"""
    <div class="term-card">
        <div class="term-title">🚚 适用运输方式</div>
        <div>{selected_term_detail['transport']}</div>
    </div>
    """, unsafe_allow_html=True)

with col_term2:
    st.markdown(f"""
    <div class="term-card">
        <div class="term-title">📖 详细说明</div>
        <div>{selected_term_detail['description']}</div>
    </div>
    """, unsafe_allow_html=True)
    
    col_resp1, col_resp2 = st.columns(2)
    with col_resp1:
        st.markdown(f"""
        <div class="term-card">
            <div class="term-title">👤 卖方责任</div>
            <div>{selected_term_detail['responsibility_seller']}</div>
        </div>
        """, unsafe_allow_html=True)
    with col_resp2:
        st.markdown(f"""
        <div class="term-card">
            <div class="term-title">👥 买方责任</div>
            <div>{selected_term_detail['responsibility_buyer']}</div>
        </div>
        """, unsafe_allow_html=True)
    
    st.markdown(f"""
    <div class="term-card">
        <div class="term-title">⚠️ 风险转移点</div>
        <div>{selected_term_detail['risk_transfer']}</div>
    </div>
    """, unsafe_allow_html=True)

# 显示所有贸易术语快速参考
with st.expander("📚 查看所有 Incoterms 2020 术语"):
    for term in INCOTERMS_2020:
        st.markdown(f"""
        <div class="term-card">
            <div class="term-title">{term['name']}</div>
            <div class="term-desc">{term['description'][:150]}...</div>
            <div style="display: flex; gap: 1rem; margin-top: 0.3rem; font-size: 0.8rem;">
                <span>🚚 {term['transport']}</span>
                <span>👤 卖方: {term['responsibility_seller'][:30]}...</span>
            </div>
        </div>
        """, unsafe_allow_html=True)

st.markdown("---")

# 底部版权
col_footer1, col_footer2, col_footer3 = st.columns(3)
with col_footer1:
    st.markdown(f"© 2026 {company_name}")
with col_footer2:
    st.markdown("技术支持: AI价到团队")
with col_footer3:
    st.markdown("PAD数据源: 阿里巴巴国际站询价页 / 公司ERP / 中国银行")














