# app.py - AI价到 - 小微外贸智能报价助手 (表格版)

import streamlit as st
import pandas as pd
import math
import subprocess
import os
import time
import hashlib
import json
from datetime import datetime, timedelta, timezone
from pathlib import Path

# -------------------- 页面配置 - 必须放在最前面 --------------------
st.set_page_config(
    page_title="AI价到 - 小微外贸智能出口报价助手",
    page_icon="💰",
    layout="wide",
    initial_sidebar_state="expanded"
)

# -------------------- 常量定义 --------------------
BASE_DATA_PATH = "C:\\Basic Information"
EXCEL_FILE = os.path.join(BASE_DATA_PATH, "Data.xlsx")
CACHE_FILE = os.path.join(BASE_DATA_PATH, "cache.json")
CACHE_DURATION = 24 * 60 * 60  # 24小时缓存时间（秒）

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

# 集装箱参数表
CONTAINER_SPECS = {
    # 普通集装箱
    "20'GP": {"type": "普通", "code": "GP", "volume": 33, "weight": 25000, "tare": 2275, "display": "20' 普通"},
    "40'GP": {"type": "普通", "code": "GP", "volume": 67, "weight": 29000, "tare": 3760, "display": "40' 普通"},
    "40'HC": {"type": "普通", "code": "HC", "volume": 76, "weight": 29000, "tare": 3950, "display": "40' 高箱"},
    # 冷冻集装箱
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

# 2020版国际贸易术语完整列表
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
    .budget-table tr:hover {
        background-color: #f0f2f6;
    }
    .budget-total {
        background-color: #FFD700 !important;
        font-weight: bold;
    }
    .budget-total td {
        background-color: #FFD700;
    }
    .budget-highlight {
        background-color: #d4edda !important;
        font-weight: bold;
    }
    .budget-highlight td {
        background-color: #d4edda;
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
    .price-box {
        background: linear-gradient(135deg, #FFD700 0%, #FFA500 100%);
        padding: 1rem;
        border-radius: 10px;
        text-align: center;
        margin: 1rem 0;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
    }
    .price-box .label {
        color: #0A174E;
        font-size: 1rem;
        margin-bottom: 0.5rem;
    }
    .price-box .value {
        color: #0A174E;
        font-size: 2rem;
        font-weight: bold;
    }
    .refresh-button {
        background-color: #FFD700;
        color: #0A174E;
        border: none;
        padding: 0.25rem 1rem;
        border-radius: 5px;
        font-weight: bold;
        cursor: pointer;
    }
    .refresh-button:hover {
        background-color: #FFA500;
    }
    .file-status {
        background-color: #e9ecef;
        padding: 0.5rem;
        border-radius: 5px;
        margin-bottom: 0.5rem;
        font-size: 0.9rem;
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

# -------------------- 缓存管理函数 --------------------
def load_cache():
    """加载缓存数据"""
    if os.path.exists(CACHE_FILE):
        try:
            with open(CACHE_FILE, 'r', encoding='utf-8') as f:
                cache = json.load(f)
            return cache
        except:
            return None
    return None

def save_cache(data):
    """保存数据到缓存"""
    try:
        # 确保目录存在
        os.makedirs(BASE_DATA_PATH, exist_ok=True)
        
        cache_data = {
            "timestamp": time.time(),
            "data": data
        }
        with open(CACHE_FILE, 'w', encoding='utf-8') as f:
            json.dump(cache_data, f, ensure_ascii=False, indent=2)
        return True
    except Exception as e:
        st.error(f"保存缓存失败: {e}")
        return False

def is_cache_valid():
    """检查缓存是否在24小时内"""
    if not os.path.exists(CACHE_FILE):
        return False
    
    try:
        with open(CACHE_FILE, 'r', encoding='utf-8') as f:
            cache = json.load(f)
        
        cache_time = cache.get("timestamp", 0)
        current_time = time.time()
        
        return (current_time - cache_time) < CACHE_DURATION
    except:
        return False

# -------------------- 检查Excel文件 --------------------
def check_excel_file():
    """检查Excel文件是否存在并返回详细信息"""
    file_info = {
        "exists": False,
        "path": EXCEL_FILE,
        "size": 0,
        "modified": "",
        "sheets": []
    }
    
    if os.path.exists(EXCEL_FILE):
        file_info["exists"] = True
        file_info["size"] = os.path.getsize(EXCEL_FILE)
        mod_time = os.path.getmtime(EXCEL_FILE)
        file_info["modified"] = datetime.fromtimestamp(mod_time).strftime('%Y-%m-%d %H:%M:%S')
        
        try:
            excel_file = pd.ExcelFile(EXCEL_FILE)
            file_info["sheets"] = excel_file.sheet_names
        except:
            file_info["sheets"] = ["无法读取工作表"]
    
    return file_info

# -------------------- 从Excel加载所有数据 --------------------
def load_all_data_from_excel(force_refresh=False):
    """
    从Excel文件加载所有数据
    如果force_refresh=True，强制重新读取Excel，忽略缓存
    """
    
    # 检查缓存（如果不是强制刷新）
    if not force_refresh and is_cache_valid():
        cache = load_cache()
        if cache:
            return cache["data"]
    
    # 默认数据结构
    data = {
        "ports": {},
        "rates": DEFAULT_RATES.copy(),
        "hs_info": {},
        "customer": None,
        "product": None,
        "fetch_time": format_beijing_time(),
        "file_exists": False,
        "publish_time": "未知",
        "fetch_time_excel": "未知"
    }
    
    # 检查Excel文件是否存在
    if not os.path.exists(EXCEL_FILE):
        return data
    
    data["file_exists"] = True
    
    try:
        # 读取所有工作表
        excel_file = pd.ExcelFile(EXCEL_FILE)
        
        # ========== 1. 读取港口信息表 ==========
        if SHEET_PORTS in excel_file.sheet_names:
            df_ports = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_PORTS)
            if not df_ports.empty:
                ports_dict = {}
                for _, row in df_ports.iterrows():
                    if len(row) >= 2:
                        country = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
                        port = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ""
                        if country and port and country != 'nan' and port != 'nan':
                            ports_dict[country] = port
                data["ports"] = ports_dict
        
        # ========== 2. 读取汇率表 ==========
        if SHEET_RATES in excel_file.sheet_names:
            df_rates = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_RATES)
            if not df_rates.empty:
                rates_dict = DEFAULT_RATES.copy()
                for _, row in df_rates.iterrows():
                    if len(row) >= 3:
                        currency = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
                        try:
                            rate = float(row.iloc[2]) if pd.notna(row.iloc[2]) else 0
                            if currency and currency != 'nan' and rate > 0:
                                rates_dict[currency] = rate
                        except:
                            pass
                data["rates"] = rates_dict
                
                # 记录发布时间和抓取时间（如果有）
                if len(df_rates.columns) > 3:
                    data["publish_time"] = str(df_rates.iloc[0, 3]) if pd.notna(df_rates.iloc[0, 3]) else "未知"
                if len(df_rates.columns) > 4:
                    data["fetch_time_excel"] = str(df_rates.iloc[0, 4]) if pd.notna(df_rates.iloc[0, 4]) else format_beijing_time()
        
        # ========== 3. 读取HS信息表 ==========
        if SHEET_HS in excel_file.sheet_names:
            df_hs = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_HS)
            if not df_hs.empty:
                hs_dict = {}
                for _, row in df_hs.iterrows():
                    if len(row) >= 2:
                        hs_code = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
                        hs_desc = str(row.iloc[1]).strip() if len(row) > 1 and pd.notna(row.iloc[1]) else ""
                        if hs_code and hs_code != 'nan':
                            hs_dict[hs_code] = {
                                "description": hs_desc,
                                "tax_rate": float(row.iloc[2]) if len(row) > 2 and pd.notna(row.iloc[2]) else 0,
                                "supervision": str(row.iloc[3]) if len(row) > 3 and pd.notna(row.iloc[3]) else "无",
                                "inspection": str(row.iloc[4]) if len(row) > 4 and pd.notna(row.iloc[4]) else "无"
                            }
                data["hs_info"] = hs_dict
        
        # ========== 4. 读取客户信息表 ==========
        if SHEET_CUSTOMERS in excel_file.sheet_names:
            df_customers = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_CUSTOMERS)
            if not df_customers.empty:
                latest = df_customers.iloc[-1]
                data["customer"] = {
                    "customer_name": str(latest.iloc[0]) if len(df_customers.columns) > 0 and pd.notna(latest.iloc[0]) else "",
                    "customer_rep": str(latest.iloc[1]) if len(df_customers.columns) > 1 and pd.notna(latest.iloc[1]) else "",
                    "customer_country": str(latest.iloc[2]) if len(df_customers.columns) > 2 and pd.notna(latest.iloc[2]) else "",
                    "customer_email": str(latest.iloc[3]) if len(df_customers.columns) > 3 and pd.notna(latest.iloc[3]) else "",
                    "customer_address": str(latest.iloc[4]) if len(df_customers.columns) > 4 and pd.notna(latest.iloc[4]) else "",
                    "payment_terms": str(latest.iloc[5]) if len(df_customers.columns) > 5 and pd.notna(latest.iloc[5]) else "",
                    "fetch_time": str(latest.iloc[6]) if len(df_customers.columns) > 6 and pd.notna(latest.iloc[6]) else format_beijing_time()
                }
        
        # ========== 5. 读取商品信息表 ==========
        if SHEET_PRODUCTS in excel_file.sheet_names:
            df_products = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_PRODUCTS)
            if not df_products.empty:
                latest = df_products.iloc[-1]
                data["product"] = {
                    "product_code": str(latest.iloc[0]) if len(df_products.columns) > 0 and pd.notna(latest.iloc[0]) else "N003",
                    "goods_type": str(latest.iloc[1]) if len(df_products.columns) > 1 and pd.notna(latest.iloc[1]) else "宝石或半宝石",
                    "product_name": str(latest.iloc[2]) if len(df_products.columns) > 2 and pd.notna(latest.iloc[2]) else "蓝宝石",
                    "product_name_en": str(latest.iloc[3]) if len(df_products.columns) > 3 and pd.notna(latest.iloc[3]) else "Sapphires",
                    "specification_cn": str(latest.iloc[4]) if len(df_products.columns) > 4 and pd.notna(latest.iloc[4]) else "已加工，未镶嵌，天然，无等级，刚玉",
                    "specification_en": str(latest.iloc[5]) if len(df_products.columns) > 5 and pd.notna(latest.iloc[5]) else "Processed,not inlaid,natural,no grade,corundum",
                    "hs_code": str(latest.iloc[6]) if len(df_products.columns) > 6 and pd.notna(latest.iloc[6]) else "7103910000",
                    "sales_unit": str(latest.iloc[7]) if len(df_products.columns) > 7 and pd.notna(latest.iloc[7]) else "克拉（CT）",
                    "quantity": float(latest.iloc[8]) if len(df_products.columns) > 8 and pd.notna(latest.iloc[8]) else 0,
                    "price_per_ct": float(latest.iloc[9]) if len(df_products.columns) > 9 and pd.notna(latest.iloc[9]) else 0,
                    "package_unit": str(latest.iloc[10]) if len(df_products.columns) > 10 and pd.notna(latest.iloc[10]) else "纸箱（CARTON）",
                    "unit_conversion": str(latest.iloc[11]) if len(df_products.columns) > 11 and pd.notna(latest.iloc[11]) else "1000CT/CARTON",
                    "gross_weight": float(latest.iloc[12]) if len(df_products.columns) > 12 and pd.notna(latest.iloc[12]) else 0.70,
                    "net_weight": float(latest.iloc[13]) if len(df_products.columns) > 13 and pd.notna(latest.iloc[13]) else 0.20,
                    "volume_per_pack": float(latest.iloc[14]) if len(df_products.columns) > 14 and pd.notna(latest.iloc[14]) else 0.0400,
                    "legal_unit": str(latest.iloc[15]) if len(df_products.columns) > 15 and pd.notna(latest.iloc[15]) else "克拉（CT）",
                    "customs_supervision": str(latest.iloc[16]) if len(df_products.columns) > 16 and pd.notna(latest.iloc[16]) else "无",
                    "inspection_category": str(latest.iloc[17]) if len(df_products.columns) > 17 and pd.notna(latest.iloc[17]) else "无",
                    "transport_notes": str(latest.iloc[18]) if len(df_products.columns) > 18 and pd.notna(latest.iloc[18]) else "无",
                    "description": str(latest.iloc[19]) if len(df_products.columns) > 19 and pd.notna(latest.iloc[19]) else "",
                    "fetch_time": str(latest.iloc[20]) if len(df_products.columns) > 20 and pd.notna(latest.iloc[20]) else format_beijing_time()
                }
        
        # 保存到缓存
        save_cache(data)
        
    except Exception as e:
        st.error(f"读取Excel文件时出错: {e}")
    
    return data

# -------------------- 手动刷新数据函数 --------------------
def refresh_all_data():
    """手动刷新所有数据"""
    with st.spinner("正在从Excel重新加载数据..."):
        data = load_all_data_from_excel(force_refresh=True)
        
        # 更新session state
        if "ports" in data:
            st.session_state.ports = data["ports"]
        
        if "rates" in data:
            st.session_state.exchange_rates = data["rates"]
        
        if "customer" in data and data["customer"]:
            st.session_state.customer_data = data["customer"]
            st.session_state.customer_fetched = True
        
        if "product" in data and data["product"]:
            st.session_state.product_data = data["product"]
            st.session_state.product_fetched = True
        
        if "hs_info" in data:
            st.session_state.hs_info = data["hs_info"]
        
        st.session_state.last_refresh_time = format_beijing_time()
        st.session_state.data_source = "Excel" if data.get("file_exists") else "默认数据"
        st.session_state.publish_time = data.get("publish_time", "未知")
        st.session_state.fetch_time_excel = data.get("fetch_time_excel", "未知")
        
        st.success(f"数据刷新完成！")
        time.sleep(1)
        st.rerun()

# -------------------- 辅助函数：国家港口映射 --------------------
def get_country_port_map():
    """获取国家港口映射（优先使用Excel数据）"""
    if "ports" in st.session_state and st.session_state.ports:
        return st.session_state.ports
    
    # 默认映射（作为备用）
    return {
        "Chile": "San Antonio", "USA": "Los Angeles", "Germany": "Hamburg",
        "Philippines": "Manila", "China": "Shanghai", "Japan": "Tokyo",
        "UK": "Felixstowe", "France": "Le Havre", "Italy": "Genoa",
        "Australia": "Sydney", "Brazil": "Santos", "India": "Mumbai"
    }

# -------------------- 辅助函数：获取HS信息 --------------------
def get_hs_info(hs_code):
    """根据HS编码获取详细信息"""
    if "hs_info" in st.session_state and hs_code in st.session_state.hs_info:
        return st.session_state.hs_info[hs_code]
    return None

# -------------------- 运行PAD流程函数 --------------------
def run_pad_flow(flow_name):
    """调用Power Automate Desktop运行指定流程"""
    try:
        pad_path = "C:\\Program Files (x86)\\Power Automate Desktop\\PAD.Console.exe"
        if os.path.exists(pad_path):
            result = subprocess.run([pad_path, "/Run", flow_name], capture_output=True, text=True, timeout=10)
            return {"success": True, "message": f"已启动PAD流程: {flow_name}"}
        else:
            st.info(f"模拟运行PAD流程: {flow_name}")
            return {"success": True, "message": f"模拟运行成功"}
    except Exception as e:
        return {"success": False, "message": str(e)}

# -------------------- 计算运输方案函数 --------------------
def calculate_shipping_options(total_volume, total_weight, freight_rates, lcl_rate_cbm, lcl_rate_kg):
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
        calculation = f"{total_volume:.2f} CBM × ${lcl_rate_cbm:.2f}/CBM = ${lcl_volume_cost:,.2f}"
    else:
        cargo_type = "重货 (按重量计费)"
        calculation = f"{total_weight:.2f} KG ÷ 1000 × ${lcl_rate_kg:.2f}/吨 = ${lcl_weight_cost:,.2f}"
    
    options.append({
        "name": "LCL散货",
        "description": f"散货拼箱 - {cargo_type}",
        "type": "lcl",
        "containers": {},
        "cost_usd": lcl_cost,
        "volume_cost": lcl_volume_cost,
        "weight_cost": lcl_weight_cost,
        "calculation": calculation,
        "details": f"体积计费: ${lcl_volume_cost:,.2f} | 重量计费: ${lcl_weight_cost:,.2f} | 取大值: ${lcl_cost:,.2f}"
    })
    
    # ========== 方案2: 整箱运输 ==========
    # 获取所有集装箱类型
    container_list = []
    for container_name, spec in CONTAINER_SPECS.items():
        if container_name in freight_rates:
            container_list.append({
                "name": container_name,
                "display": spec["display"],
                "type": spec["type"],
                "volume": spec["volume"],
                "weight": spec["weight"],
                "freight": freight_rates[container_name]
            })
    
    # 如果没有设置任何集装箱运费，返回仅LCL方案
    if not container_list:
        return options, 0
    
    # 计算所需集装箱的最大数量（向上取整）- 取所有类型中的最大值
    max_containers = 0
    for container in container_list:
        by_volume = math.ceil(total_volume / container["volume"])
        by_weight = math.ceil(total_weight / container["weight"])
        max_for_type = max(by_volume, by_weight)
        max_containers = max(max_containers, max_for_type)
    
    # 限制搜索范围，避免组合爆炸
    max_containers = min(max_containers, 5)  # 最多搜索到5个柜子
    
    # 生成所有可能的组合
    fcl_options = []
    
    # 单一类型组合
    for container in container_list:
        for num in range(1, max_containers + 1):
            total_container_volume = num * container["volume"]
            total_container_weight = num * container["weight"]
            
            if total_container_volume >= total_volume and total_container_weight >= total_weight:
                container_cost = num * container["freight"]
                
                # 计算利用率
                volume_utilization = (total_volume / total_container_volume) * 100
                weight_utilization = (total_weight / total_container_weight) * 100
                
                # 构建方案名称
                scheme_name = f"{num}×{container['display']}"
                
                # 计算剩余空间
                remaining_volume = total_container_volume - total_volume
                remaining_weight = total_container_weight - total_weight
                
                # 计算原理
                calculation = f"{num} × ${container['freight']:,.2f} = ${container_cost:,.2f}"
                
                fcl_options.append({
                    "name": scheme_name,
                    "description": f"{scheme_name} - 体积利用率: {volume_utilization:.1f}%, 重量利用率: {weight_utilization:.1f}%",
                    "type": "fcl",
                    "container_type": container['name'],
                    "container_display": container['display'],
                    "container_count": num,
                    "cost_usd": container_cost,
                    "volume_utilization": volume_utilization,
                    "weight_utilization": weight_utilization,
                    "remaining_volume": remaining_volume,
                    "remaining_weight": remaining_weight,
                    "calculation": calculation,
                    "details": f"总运费: ${container_cost:,.2f} | 剩余体积: {remaining_volume:.1f}CBM | 剩余重量: {remaining_weight:.1f}KG"
                })
    
    # 按成本排序
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

# -------------------- 显示出口预算表函数 --------------------
def display_budget_table(budget, selected_term, exchange_rate, selected_currency):
    """按照附件3的格式显示出口预算表 - 3列表格"""
    
    # 计算6 = 1-2+3+4+5
    total = (budget.get('采购成本', 0) - budget.get('出口退税', 0) + 
             budget.get('国内费用合计', 0) + budget.get('银行费用合计', 0) + 
             budget.get('其他费用', 0))
    
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
            <td rowspan="9"><strong>3. 国内费用</strong></td>
            <td>出口国内运费</td>
            <td style="text-align: right">{budget.get('内陆运费', 6348.89):,.2f}</td>
        </tr>
        <tr>
            <td>国际运费</td>
            <td style="text-align: right">{budget.get('海运费', 13.85):,.2f}</td>
        </tr>
        <tr>
            <td>出口货代杂费</td>
            <td style="text-align: right">{budget.get('货代杂费', 1587.22):,.2f}</td>
        </tr>
        <tr>
            <td>出口商检费</td>
            <td style="text-align: right">{budget.get('商检费', 0):,.2f}</td>
        </tr>
        <tr>
            <td>检验检疫证书费</td>
            <td style="text-align: right">{budget.get('证书费', 0):,.2f}</td>
        </tr>
        <tr>
            <td>出口报关费</td>
            <td style="text-align: right">{budget.get('报关费', 41.04):,.2f}</td>
        </tr>
        <tr>
            <td>出口关税</td>
            <td style="text-align: right">{budget.get('出口关税', 0):,.2f}</td>
        </tr>
        <tr>
            <td>产地证书费</td>
            <td style="text-align: right">{budget.get('产地证费', 0):,.2f}</td>
        </tr>
        <tr>
            <td>保险费</td>
            <td style="text-align: right">{budget.get('保险费', 9534.81):,.2f}</td>
        </tr>
        <tr>
            <td></td>
            <td><strong>合计</strong></td>
            <td style="text-align: right"><strong>{budget.get('国内费用合计', 0):,.2f}</strong></td>
        </tr>
        <tr>
            <td rowspan="3"><strong>4. 银行费用</strong></td>
            <td>托收费用</td>
            <td style="text-align: right">{budget.get('托收费', 0):,.2f}</td>
        </tr>
        <tr>
            <td>信用证费用</td>
            <td style="text-align: right">{budget.get('信用证费', 969.40):,.2f}</td>
        </tr>
        <tr>
            <td>其他费用</td>
            <td style="text-align: right">{budget.get('其他银行费', 0):,.2f}</td>
        </tr>
        <tr>
            <td></td>
            <td><strong>合计</strong></td>
            <td style="text-align: right"><strong>{budget.get('银行费用合计', 969.40):,.2f}</strong></td>
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
            <td>{budget.get('对外报价', 0):,.2f} {selected_currency}</td>
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
if 'inland_freight' not in st.session_state:
    st.session_state.inland_freight = 6348.89  # 出口内陆运费
if 'forwarder_fee' not in st.session_state:
    st.session_state.forwarder_fee = 1587.22  # 出口货代杂费
if 'inspection_fee_detail' not in st.session_state:
    st.session_state.inspection_fee_detail = 0  # 出口商检费
if 'certificate_fee' not in st.session_state:
    st.session_state.certificate_fee = 0  # 检验检疫证书费
if 'customs_declare_fee' not in st.session_state:
    st.session_state.customs_declare_fee = 41.04  # 出口报关费
if 'export_tariff' not in st.session_state:
    st.session_state.export_tariff = 0  # 出口关税
if 'origin_cert_fee' not in st.session_state:
    st.session_state.origin_cert_fee = 0  # 产地证书费
if 'collection_fee' not in st.session_state:
    st.session_state.collection_fee = 0  # 托收费用
if 'lc_fee' not in st.session_state:
    st.session_state.lc_fee = 969.40  # 信用证费用
if 'other_bank_fee' not in st.session_state:
    st.session_state.other_bank_fee = 0  # 其他银行费用
if 'other_fee' not in st.session_state:
    st.session_state.other_fee = 0  # 其他费用
if 'freight_rates' not in st.session_state:
    # 初始化所有集装箱类型的运费
    st.session_state.freight_rates = {
        "20'GP": 1200.0,
        "40'GP": 1800.0,
        "40'HC": 2000.0,
        "20'RF": 2500.0,
        "40'RF": 3500.0,
        "40'RH": 3800.0
    }
if 'lcl_rate_cbm' not in st.session_state:
    st.session_state.lcl_rate_cbm = 50.0   # LCL(M) 按体积费率 (USD/CBM)
if 'lcl_rate_kg' not in st.session_state:
    st.session_state.lcl_rate_kg = 2000.0  # LCL(W) 按重量费率 (USD/吨)
if 'ports' not in st.session_state:
    st.session_state.ports = {}
if 'hs_info' not in st.session_state:
    st.session_state.hs_info = {}
if 'last_refresh_time' not in st.session_state:
    st.session_state.last_refresh_time = "从未刷新"
if 'data_source' not in st.session_state:
    st.session_state.data_source = "默认数据"
if 'publish_time' not in st.session_state:
    st.session_state.publish_time = "未知"
if 'fetch_time_excel' not in st.session_state:
    st.session_state.fetch_time_excel = "未知"
if 'budget' not in st.session_state:
    st.session_state.budget = None

# 初始化加载数据
if 'exchange_rates' not in st.session_state:
    # 首次加载，尝试从Excel读取
    data = load_all_data_from_excel()
    st.session_state.exchange_rates = data.get("rates", DEFAULT_RATES)
    st.session_state.ports = data.get("ports", {})
    st.session_state.hs_info = data.get("hs_info", {})
    
    if data.get("customer"):
        st.session_state.customer_data = data["customer"]
        st.session_state.customer_fetched = True
    
    if data.get("product"):
        st.session_state.product_data = data["product"]
        st.session_state.product_fetched = True
    
    if data.get("file_exists"):
        st.session_state.data_source = "Excel"
    else:
        st.session_state.data_source = "默认数据"
    
    st.session_state.last_refresh_time = format_beijing_time()
    st.session_state.publish_time = data.get("publish_time", "未知")
    st.session_state.fetch_time_excel = data.get("fetch_time_excel", "未知")

# ==================== 页面内容开始 ====================

# -------------------- 顶部公司信息及PAD按钮 --------------------
st.markdown("""
<div class="main-header">
    <h1>💰 AI价到 - 小微外贸智能出口报价助手</h1>
    <div class="subtitle">智能报价 · 精准计算 · 一键成交</div>
</div>
""", unsafe_allow_html=True)

# 检查Excel文件状态
file_info = check_excel_file()
if file_info["exists"]:
    st.success(f"✅ 找到Excel文件: C:\\Basic Information\\Data.xlsx")
    st.info(f"📁 文件大小: {file_info['size']} bytes | 修改时间: {file_info['modified']}")
    if file_info["sheets"]:
        st.info(f"📊 工作表: {', '.join(file_info['sheets'])}")
else:
    st.error(f"⚠️ Excel文件不存在: C:\\Basic Information\\Data.xlsx")
    st.info("请确认文件路径是否正确，或使用默认数据")

# 第一行：PAD抓取按钮和刷新按钮
col_pad1, col_pad2, col_pad3, col_pad4, col_refresh = st.columns([2,2,2,2,1])

with col_pad1:
    if st.button("🤖 抓取客户信息 (PAD)", use_container_width=True):
        with st.spinner("正在启动Power Automate Desktop抓取客户信息..."):
            result = run_pad_flow("FetchCustomerFromAlibaba")
            if result["success"]:
                st.success(result["message"])
                time.sleep(3)
                # 手动刷新数据
                refresh_all_data()
            else:
                st.error(f"启动失败: {result['message']}")

with col_pad2:
    if st.button("📦 抓取商品信息 (PAD)", use_container_width=True):
        with st.spinner("正在启动Power Automate Desktop抓取商品信息..."):
            result = run_pad_flow("FetchProductFromMarket")
            if result["success"]:
                st.success(result["message"])
                time.sleep(3)
                # 手动刷新数据
                refresh_all_data()
            else:
                st.error(f"启动失败: {result['message']}")

with col_pad3:
    if st.button("📊 刷新汇率 (PAD)", use_container_width=True):
        with st.spinner("正在启动Power Automate Desktop更新汇率..."):
            result = run_pad_flow("FetchBOERates")
            if result["success"]:
                st.success(result["message"])
                time.sleep(2)
                # 手动刷新数据
                refresh_all_data()
            else:
                st.error(f"启动失败: {result['message']}")

with col_pad4:
    st.markdown(f"<div style='text-align: center; padding: 0.5rem; background-color: #e9ecef; border-radius: 5px;'>🕒 {format_beijing_time()}</div>", unsafe_allow_html=True)

with col_refresh:
    if st.button("🔄 刷新数据", use_container_width=True, type="secondary"):
        refresh_all_data()

# 显示数据源状态
if st.session_state.data_source == "Excel":
    st.success(f"✅ 数据源: {st.session_state.data_source} | 最后刷新: {st.session_state.last_refresh_time}")
    if st.session_state.publish_time != "未知":
        st.info(f"📅 汇率牌价时间: {st.session_state.publish_time}")
else:
    st.warning(f"⚠️ 数据源: {st.session_state.data_source} (使用内置默认数据) | 最后刷新: {st.session_state.last_refresh_time}")

st.markdown("---")

# -------------------- 侧边栏：汇率、HS信息、物流信息 --------------------
with st.sidebar:
    st.markdown('<p class="sidebar-header">💱 汇率</p>', unsafe_allow_html=True)
    # 汇率状态
    if st.session_state.data_source == "Excel":
        st.markdown("✅ <span class='status-badge status-success'>Excel数据已连接</span>", unsafe_allow_html=True)
    else:
        st.markdown("⚠️ <span class='status-badge status-warning'>使用默认汇率</span>", unsafe_allow_html=True)
    
    # 货币选择
    available_currencies = list(st.session_state.exchange_rates.keys())
    if st.session_state.selected_currency not in available_currencies:
        st.session_state.selected_currency = "USD" if "USD" in available_currencies else available_currencies[0] if available_currencies else "USD"
    
    target_currency = st.selectbox("报价货币", available_currencies, 
                                  index=available_currencies.index(st.session_state.selected_currency) 
                                  if st.session_state.selected_currency in available_currencies else 0,
                                  key="sidebar_currency")
    st.session_state.selected_currency = target_currency
    current_rate = st.session_state.exchange_rates[target_currency]
    st.metric(f"1 {target_currency} = ", f"{current_rate:.4f} CNY")
    
    st.markdown("---")
    
    st.markdown('<p class="sidebar-header">📋 HS编码信息</p>', unsafe_allow_html=True)
    hs_code_display = st.session_state.product_data.get("hs_code", "未获取") if st.session_state.product_fetched else "未填写"
    
    # 如果有HS信息，显示详细信息
    hs_info = get_hs_info(hs_code_display)
    if hs_info:
        st.text_input("HS编码", value=hs_code_display, disabled=True, key="hs_code_sidebar")
        st.text_area("商品描述", value=hs_info.get("description", ""), disabled=True, height=60)
        st.metric("退税率", f"{hs_info.get('tax_rate', 0)}%")
        st.text_input("监管条件", value=hs_info.get("supervision", "无"), disabled=True)
        st.text_input("检验检疫", value=hs_info.get("inspection", "无"), disabled=True)
    else:
        st.text_input("HS编码", value=hs_code_display, disabled=True, key="hs_code_sidebar")
        st.info("无详细HS信息")
    
    st.markdown("---")
    
    st.markdown('<p class="sidebar-header">🚢 物流信息</p>', unsafe_allow_html=True)
    
    # 显示集装箱参数表
    st.markdown("""
    <div class="container-table">
        <table>
            <tr>
                <th>箱型</th>
                <th>代码</th>
                <th>体积(CBM)</th>
                <th>重量(KGS)</th>
                <th>自重(KGS)</th>
            </tr>
    """, unsafe_allow_html=True)
    
    for name, spec in CONTAINER_SPECS.items():
        st.markdown(f"""
            <tr>
                <td>{name}</td>
                <td>{spec['code']}</td>
                <td>{spec['volume']}</td>
                <td>{spec['weight']}</td>
                <td>{spec['tare']}</td>
            </tr>
        """, unsafe_allow_html=True)
    
    st.markdown("</table></div>", unsafe_allow_html=True)
    
    col_from, col_to = st.columns(2)
    with col_from:
        departure_port = st.text_input("起运港", value="Shanghai", key="departure_port")
    with col_to:
        # 使用从Excel加载的港口映射
        port_map = get_country_port_map()
        default_dest = port_map.get(st.session_state.customer_data.get("customer_country", ""), "")
        destination_port = st.text_input("目的港", value=default_dest, key="destination_port")
    
    st.markdown("**集装箱运费估算 (USD)**")
    
    # 创建运费估算表格
    # 使用可编辑的运费设置
    freight_rates = st.session_state.freight_rates.copy()
    
    for container in CONTAINER_TYPES:
        col_f1, col_f2 = st.columns([1, 1])
        with col_f1:
            st.markdown(f"**{container['display']}**")
        with col_f2:
            freight_rates[container['name']] = st.number_input(
                "", 
                value=float(freight_rates.get(container['name'], 1200.0)), 
                step=50.0, 
                key=f"freight_{container['name']}", 
                label_visibility="collapsed"
            )
    
    st.markdown("**LCL散货费率**")
    st.caption("LCL运费 = max(体积×LCL(M)单价, 重量×LCL(W)单价)")
    
    col_l1, col_l2 = st.columns([1, 1])
    with col_l1:
        st.markdown("**LCL(M) (USD/CBM)**")
    with col_l2:
        lcl_rate_cbm = st.number_input("", value=float(st.session_state.lcl_rate_cbm), step=5.0, key="lcl_rate_cbm_input", label_visibility="collapsed")
    
    col_l3, col_l4 = st.columns([1, 1])
    with col_l3:
        st.markdown("**LCL(W) (USD/吨)**")
    with col_l4:
        lcl_rate_kg = st.number_input("", value=float(st.session_state.lcl_rate_kg), step=100.0, key="lcl_rate_kg_input", label_visibility="collapsed")
    
    if st.button("更新运费设置", key="update_freight"):
        st.session_state.freight_rates = freight_rates
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
        st.success("✅ 已从Excel加载客户数据")
    else:
        st.info("⏳ 可点击上方抓取按钮获取或手动输入")
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
        # 使用从Excel加载的港口映射
        port_map = get_country_port_map()
        countries = list(port_map.keys())
        if not countries:
            countries = ["China", "USA", "Germany", "UK", "Japan"]
        
        country_index = 0
        if default_customer.get("customer_country") and default_customer["customer_country"] in countries:
            country_index = countries.index(default_customer["customer_country"])
        country = st.selectbox(
            "目的国家", 
            countries, 
            index=country_index, 
            key="customer_country_input"
        )
        port = port_map.get(country, "")
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
    st.success("✅ 已从Excel加载商品数据")
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
    selected_term = st.selectbox("贸易术语 (Incoterms 2020)", term_options, index=2, key="selected_term")  # 默认CIF
    selected_term_detail = next((term for term in INCOTERMS_2020 if term["name"] == selected_term), INCOTERMS_2020[0])
    st.caption(f"{selected_term_detail['description'][:100]}...")
    
    profit_margin = st.slider("利润率 (%)", min_value=5, max_value=100, value=20, step=5, key="profit_margin")
    
    # 如果有HS信息，使用HS的退税率
    hs_info = get_hs_info(hs_code)
    default_tax_rate = hs_info.get("tax_rate", 13) if hs_info else 13
    tax_rate = st.slider("出口退税率 (%)", min_value=0, max_value=17, value=int(default_tax_rate), step=1, key="tax_rate")

with col_pay:
    st.markdown("**付款方式**")
    payment_method = st.selectbox("付款方式", ["T/T", "L/C", "D/P", "D/A"], index=1, key="payment_method")  # 默认L/C
    payment_bank = st.text_input("付款银行", value="Bank of China", key="payment_bank")
    payment_terms_detail = st.text_input("付款条件", value="30% deposit, 70% against B/L", key="payment_terms_detail")

st.markdown("---")

# -------------------- 开始报价按钮和出口预算表 --------------------
col_btn1, col_btn2 = st.columns([1, 5])
with col_btn1:
    calculate_pressed = st.button("开始报价", type="primary", use_container_width=True)

# 获取当前汇率
current_exchange_rate = st.session_state.exchange_rates[st.session_state.selected_currency]

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
                # 解析类似 "1000CT/CARTON" 的格式
                ct_part = conversion_parts[0].replace('CT', '').strip()
                ct_per_carton = float(ct_part) if ct_part else 1000
                total_packages = math.ceil(quantity / ct_per_carton)
            else:
                total_packages = math.ceil(quantity / 1000)
        except:
            total_packages = math.ceil(quantity / 1000)
        
        total_volume = total_packages * volume_per_pack
        total_gross_weight = total_packages * gross_weight
        total_net_weight = total_packages * net_weight
        
        # 计算运输方案
        shipping_options, best_index = calculate_shipping_options(
            total_volume, 
            total_gross_weight,
            st.session_state.freight_rates,
            st.session_state.lcl_rate_cbm,
            st.session_state.lcl_rate_kg
        )
        
        best_option = shipping_options[best_index]
        freight_usd = best_option["cost_usd"]
        freight_cny = freight_usd * exchange_rate
        
        # 国内费用合计
        domestic_fee_total = (st.session_state.handling_fee + st.session_state.inspection_fee + 
                             st.session_state.document_fee + st.session_state.inland_freight +
                             st.session_state.forwarder_fee + st.session_state.inspection_fee_detail +
                             st.session_state.certificate_fee + st.session_state.customs_declare_fee +
                             st.session_state.export_tariff + st.session_state.origin_cert_fee + freight_cny)
        
        # 保险费
        insurance_fee = total_cost_cny * (st.session_state.insurance_rate / 100)
        
        # 银行费用合计
        bank_fee_total = st.session_state.collection_fee + st.session_state.lc_fee + st.session_state.other_bank_fee
        
        # 出口退税
        tax_refund = total_cost_cny * (tax_rate / 100)
        
        # 总成本 = 采购成本 - 退税 + 国内费用 + 银行费用 + 其他费用 + 保险费
        total_cost_with_freight = (total_cost_cny - tax_refund + domestic_fee_total + 
                                   bank_fee_total + st.session_state.other_fee + insurance_fee)
        
        # 对外报价 (基于利润率)
        quoted_price_target = total_cost_target * (1 + profit_margin/100)
        quoted_price_cny = quoted_price_target * exchange_rate
        
        # 预期盈亏额
        expected_profit = quoted_price_cny - total_cost_with_freight
        expected_profit_rate = (expected_profit / total_cost_with_freight) * 100 if total_cost_with_freight != 0 else 0
        
        budget = {
            "采购成本": total_cost_cny,
            "出口退税": tax_refund,
            "内陆运费": st.session_state.inland_freight,
            "海运费": freight_cny,
            "货代杂费": st.session_state.forwarder_fee,
            "商检费": st.session_state.inspection_fee_detail,
            "证书费": st.session_state.certificate_fee,
            "报关费": st.session_state.customs_declare_fee,
            "出口关税": st.session_state.export_tariff,
            "产地证费": st.session_state.origin_cert_fee,
            "保险费": insurance_fee,
            "国内费用合计": domestic_fee_total + insurance_fee,  # 包含保险费
            "托收费": st.session_state.collection_fee,
            "信用证费": st.session_state.lc_fee,
            "其他银行费": st.session_state.other_bank_fee,
            "银行费用合计": bank_fee_total,
            "其他费用": st.session_state.other_fee,
            "总成本": total_cost_with_freight,
            "对外报价": quoted_price_target,
            "对外报价CNY": quoted_price_cny,
            "预期盈亏额": expected_profit,
            "预期盈亏率": expected_profit_rate,
            "总箱数": total_packages,
            "总体积": total_volume,
            "总毛重": total_gross_weight,
            "总净重": total_net_weight,
            "shipping_options": shipping_options,
            "best_shipping_index": best_index,
            "selected_shipping": best_option["name"],
            "selected_shipping_desc": best_option.get("description", ""),
            "shipping_calculation": best_option.get("calculation", ""),
            "freight_usd": freight_usd
        }
        
        st.session_state.budget = budget
        st.success("计算完成！")
        st.balloons()
    else:
        st.warning("请填写商品数量和单价")

# 显示预算表
if st.session_state.budget:
    b = st.session_state.budget
    
    # 显示整批货物信息（放在预算表前面）
    st.markdown("#### 📦 整批货物信息")
    st.markdown(f"""
    <table class="cargo-info-table">
        <tr>
            <th>总箱数</th>
            <th>总体积 (CBM)</th>
            <th>总毛重 (KG)</th>
            <th>总净重 (KG)</th>
        </tr>
        <tr>
            <td>{b.get('总箱数', 0):,} 箱</td>
            <td>{b.get('总体积', 0):.2f}</td>
            <td>{b.get('总毛重', 0):.2f}</td>
            <td>{b.get('总净重', 0):.2f}</td>
        </tr>
    </table>
    """, unsafe_allow_html=True)
    
    # 显示最终运费方案信息
    st.markdown("#### 🚢 最终运费方案")
    st.markdown(f"""
    <div class="shipping-option best-option">
        <strong>✅ 选中方案: {b.get('selected_shipping', '未知')}</strong><br>
        {b.get('selected_shipping_desc', '')}<br>
        <strong>计算原理:</strong> {b.get('shipping_calculation', '')}<br>
        <strong>运费:</strong> ${b.get('freight_usd', 0):,.2f} | ￥{b.get('海运费', 0):,.2f}
    </div>
    """, unsafe_allow_html=True)
    
    # 显示运输方案对比（可折叠）
    with st.expander("查看所有运输方案对比"):
        st.markdown("#### 所有运输方案（按成本排序）")
        shipping_options = b.get("shipping_options", [])
        best_index = b.get("best_shipping_index", 0)
        
        for i, option in enumerate(shipping_options):
            if i == best_index:
                st.markdown(f"""
                <div class="shipping-option best-option">
                    <strong>✅ 最佳方案 #{i+1}: {option.get('name', '未知')}</strong><br>
                    {option.get('description', '')}<br>
                    <strong>计算原理:</strong> {option.get('calculation', '')}<br>
                    运费: ${option.get('cost_usd', 0):,.2f} | ￥{option.get('cost_usd', 0) * current_exchange_rate:,.2f}<br>
                    <small>{option.get('details', '')}</small>
                </div>
                """, unsafe_allow_html=True)
            else:
                st.markdown(f"""
                <div class="shipping-option">
                    <strong>方案 #{i+1}: {option.get('name', '未知')}</strong><br>
                    {option.get('description', '')}<br>
                    <strong>计算原理:</strong> {option.get('calculation', '')}<br>
                    运费: ${option.get('cost_usd', 0):,.2f} | ￥{option.get('cost_usd', 0) * current_exchange_rate:,.2f}
                </div>
                """, unsafe_allow_html=True)
    
    # 显示出口预算表（按照附件3的格式）
    st.markdown("### 📊 出口预算表")
    budget_html = display_budget_table(b, selected_term, current_exchange_rate, st.session_state.selected_currency)
    st.markdown(budget_html, unsafe_allow_html=True)

st.markdown("---"

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
    st.markdown(f"© 2026 {company_name if 'company_name' in locals() else 'ABC International Trading CO. Ltd'}")
with col_footer2:
    st.markdown("技术支持: AI价到团队")
with col_footer3:
    st.markdown("PAD数据源: 阿里巴巴国际站 | Excel数据源: C:\\Basic Information\\Data.xlsx")



























