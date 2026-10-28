import time
import xlwings as xw
import akshare as ak
import pandas as pd

# ================= 配置区域 =================
EXCEL_FILE_NAME = "油价表.xlsx"  # 替换为你实际打开的 Excel 文件名（不用写完整路径，如果都在一个文件夹）
REFRESH_INTERVAL = 60           # 刷新频率：60秒（国际油价更新快）
# ============================================

# 国内油价缓存（因为国内油价十几天才变一次，没必要每分钟去爬）
_cache_cn_df = None
_last_cn_update = 0
CN_CACHE_EXPIRE = 3600 * 12 # 12小时更新一次国内油价数据源

def fetch_intl_prices():
    """
    获取国际油价（彻底解决字段名和名称匹配问题）
    """
    wti_val = None
    brent_val = None

    def get_price_from_df(symbol_code):
        try:
            # 单独请求，避免在一个表里匹配名称
            df = ak.futures_foreign_commodity_realtime(symbol=symbol_code)
            if df is None or df.empty:
                return None
            
            # 获取第一行数据
            row = df.iloc[0]
            
            # 穷举可能出现的所有价格字段名（中英文兼容）
            for col in ["current_price", "最新价", "最新价格", "price"]:
                val = row.get(col)
                if val is not None and not pd.isna(val):
                    return float(val)
            return None
        except Exception as e:
            print(f"[{symbol_code}] 获取异常: {e}")
            return None

    # 分别获取 WTI (CL) 和 布伦特 (OIL)
    wti_val = get_price_from_df("CL")
    brent_val = get_price_from_df("OIL")

    # 如果新浪接口（上面这个）失败，回退到东财接口
    if wti_val is None or brent_val is None:
        try:
            df_em = ak.futures_global_spot_em()
            if df_em is not None and not df_em.empty:
                wti_row = df_em[df_em["名称"].str.contains("WTI", case=False, na=False)]
                brent_row = df_em[df_em["名称"].str.contains("布伦特", case=False, na=False)]
                
                if wti_val is None and not wti_row.empty:
                    wti_val = float(wti_row.iloc[0]["最新价"])
                if brent_val is None and not brent_row.empty:
                    brent_val = float(brent_row.iloc[0]["最新价"])
        except Exception as e:
            print(f"东财接口获取异常: {e}")

    # 确保返回的是保留两位小数的浮点数
    if wti_val is not None and brent_val is not None:
        return round(wti_val, 2), round(brent_val, 2)
    
    return None, None

def get_cn_oil_df():
    """获取国内油价明细（带缓存）"""
    global _cache_cn_df, _last_cn_update
    now = time.time()
    
    if _cache_cn_df is None or (now - _last_cn_update > CN_CACHE_EXPIRE):
        print("更新国内油价缓存...")
        try:
            hist_df = ak.energy_oil_hist()
            hist_df["调整日期"] = pd.to_datetime(hist_df["调整日期"], errors="coerce")
            hist_df = hist_df.dropna(subset=["调整日期"]).sort_values("调整日期", ascending=False)
            latest_date = hist_df.iloc[0]["调整日期"].strftime("%Y%m%d")
            
            detail_df = ak.energy_oil_detail(date=latest_date)
            if detail_df is not None and not detail_df.empty:
                # 自动纠偏 92 和 95 反了的问题
                col_92 = next((c for c in ["V_92", "V92"] if c in detail_df.columns), None)
                col_95 = next((c for c in ["V_95", "V95"] if c in detail_df.columns), None)
                if col_92 and col_95:
                    if pd.to_numeric(detail_df[col_92], errors="coerce").mean() > pd.to_numeric(detail_df[col_95], errors="coerce").mean():
                        detail_df = detail_df.rename(columns={col_92: col_95, col_95: col_92})
                
                _cache_cn_df = detail_df
                _last_cn_update = now
        except Exception as e:
            print(f"获取国内油价失败: {e}")
            
    return _cache_cn_df

def get_cn_price(detail_df, grade="92", region=None):
    """从明细中提取指定的价格"""
    if detail_df is None: return "获取失败"
    
    grade_str = str(grade).strip().replace("#", "")
    target_col = next((c for c in [f"V_{grade_str}", f"V{grade_str}"] if c in detail_df.columns), None)
    if target_col is None: return "找不到列"
    
    if region is None or str(region).strip() == "":
        valid_prices = pd.to_numeric(detail_df[target_col], errors="coerce").dropna()
        return round(float(valid_prices.mean()), 2) if not valid_prices.empty else "无数据"
    else:
        region_str = str(region).strip()
        region_col = next((c for c in ["省份", "地区", "城市"] if c in detail_df.columns), None)
        if region_col:
            matched = detail_df[detail_df[region_col].astype(str).str.contains(region_str, na=False)]
            if not matched.empty:
                price = pd.to_numeric(matched.iloc[0][target_col], errors="coerce")
                return round(float(price), 2) if not pd.isna(price) else "无数据"
        return "未找到地区"

import akshare as ak
import pandas as pd

import akshare as ak
import pandas as pd
import requests
import time

import requests

def fetch_exchange_rate():
    """
    直接获取新浪离岸人民币(USDCNH)实时行情
    返回数据与东财 USDCNH 完全一致
    """
    try:
        # 新浪财经离岸人民币实时接口
        url = "https://hq.sinajs.cn/list=fx_susdcnh"
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)",
            "Referer": "https://finance.sina.com.cn/"
        }
        resp = requests.get(url, headers=headers, timeout=5)
        
        if resp.status_code == 200:
            # 新浪返回的数据格式为：var hq_str_fx_susdcnh="时间,买价,卖价,最新价,...";
            text = resp.text
            if "=" in text and '"' in text:
                # 提取引号内的数据
                data_str = text.split('"')[1]
                if data_str:
                    parts = data_str.split(",")
                    if len(parts) > 3:
                        # 第4个字段（索引为3）通常是当前最新价
                        latest_price = float(parts[3])
                        return round(latest_price, 4)
    except Exception as e:
        print(f"新浪实时汇率接口异常: {e}")
    
    return None

import re

def fetch_usdcny_mid_and_buy():
    """
    获取美元兑人民币的央行中间价和中行汇买价
    返回：(中间价, 买入价)，单位：人民币元
    """
    try:
        df = ak.currency_boc_sina(symbol="美元")
        if df is not None and not df.empty:
            row = df.iloc[0]
            mid = row.get("央行中间价")
            buy = row.get("中行汇买价")
            # 接口单位是“人民币元/100外币”，除以100才是1美元对应的价格
            mid = round(float(mid) / 100, 4) if mid is not None and not pd.isna(mid) else None
            buy = round(float(buy) / 100, 4) if buy is not None and not pd.isna(buy) else None
            return mid, buy
    except Exception as e:
        print(f"中行牌价接口异常: {e}")
    return None, None

def fetch_dxy():
    """
    获取美元指数（DXY）实时行情
    数据源：新浪财经 DINIW
    """
    try:
        url = "https://hq.sinajs.cn/list=DINIW"
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)",
            "Referer": "https://finance.sina.com.cn/"
        }
        resp = requests.get(url, headers=headers, timeout=5)
        if resp.status_code == 200:
            # 返回格式：var hq_str_DINIW="时间,最新价,昨收,今开,成交量,买一,最高,最低,卖一,名称,日期";
            match = re.search(r'"([^"]*)"', resp.text)
            if match:
                parts = match.group(1).split(",")
                if len(parts) > 1:
                    return round(float(parts[1]), 4)  # 第2个字段是“最新价”
    except Exception as e:
        print(f"美元指数接口异常: {e}")
    return None

import re

def fetch_intl_gold_price():
    """
    获取实时国际金价（纽约黄金期货，单位：美元/盎司）
    数据源：新浪财经 hf_GC
    """
    try:
        url = "https://hq.sinajs.cn/list=hf_GC"
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)",
            "Referer": "https://finance.sina.com.cn/"
        }
        resp = requests.get(url, headers=headers, timeout=5)
        if resp.status_code == 200:
            match = re.search(r'"([^"]*)"', resp.text)
            if match:
                parts = match.group(1).split(",")
                if len(parts) > 0:
                    return round(float(parts[0]), 2)  # 第1个字段是当前价
    except Exception as e:
        print(f"国际金价接口异常: {e}")
    return None


def fetch_domestic_gold_price():
    """
    获取实时国内金价（上海黄金交易所 Au99.99，单位：元/克）
    数据源：AKShare spot_quotations_sge
    """
    try:
        df = ak.spot_quotations_sge(symbol="Au99.99")
        if df is not None and not df.empty:
            # 取最后一行（最新时间点）的“现价”
            latest_price = df.iloc[-1]["现价"]
            if latest_price is not None and not pd.isna(latest_price):
                return round(float(latest_price), 2)
    except Exception as e:
        print(f"国内金价接口异常: {e}")
    return None

def fetch_us_treasury_yields():
    """
    获取美国国债收益率（2年、5年、10年、30年）
    数据源：AKShare bond_gb_us_sina（新浪财经）
    返回：(y2, y5, y10, y30)，单位：%
    """
    result = {}
    symbols = {
        "2Y":  "美国2年期国债",
        "5Y":  "美国5年期国债",
        "10Y": "美国10年期国债",
        "30Y": "美国30年期国债",
    }
    try:
        for key, symbol in symbols.items():
            df = ak.bond_gb_us_sina(symbol=symbol)
            if df is not None and not df.empty:
                # 取最新一行，使用“收盘”列作为最新收益率
                # bond_gb_us_sina 返回列：date, open, high, low, close, volume[reference:1]
                latest_yield = df.iloc[-1]["close"]
                if latest_yield is not None and not pd.isna(latest_yield):
                    result[key] = round(float(latest_yield), 3)  # 收益率通常保留3位小数
                else:
                    result[key] = None
            else:
                result[key] = None
    except Exception as e:
        print(f"美债收益率接口异常: {e}")
        # 统一返回 None，避免部分期限缺失导致结构不一致
        return None, None, None, None

    return result.get("2Y"), result.get("5Y"), result.get("10Y"), result.get("30Y")

import requests
import re

def fetch_tencent_quotes(codes_list):
    """
    通过腾讯财经接口批量获取实时行情
    codes_list: 形如 ["sh600188", "sz000681", "sh000001"] 的列表
    返回: {代码: 价格} 的字典
    """
    result = {}
    if not codes_list:
        return result
    
    try:
        # 腾讯接口支持批量查询，用逗号分隔
        query_str = ",".join(codes_list)
        url = f"http://qt.gtimg.cn/q={query_str}"
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)",
            "Referer": "https://finance.qq.com/"
        }
        resp = requests.get(url, headers=headers, timeout=5)
        
        if resp.status_code == 200:
            # 返回文本按分号分割，每段对应一个代码
            for line in resp.text.split(";"):
                if "~" not in line:
                    continue
                # 提取代码和字段
                # 格式：v_sh600188="1~兖矿能源~600188~12.34~..."
                match = re.search(r'v_(\w+)="([^"]*)"', line)
                if not match:
                    continue
                code, data_str = match.groups()
                fields = data_str.split("~")
                if len(fields) > 3:
                    # 腾讯接口第 4 个字段（索引 3）是当前价
                    price = fields[3]
                    if price and price != "0.00":
                        result[code] = float(price)
    except Exception as e:
        print(f"[腾讯行情] 获取异常: {e}")
    
    return result

import re

def fetch_nasdaq_index():
    """获取纳斯达克综合指数实时点位，数据源：新浪财经"""
    try:
        url = "https://hq.sinajs.cn/list=gb_ixic"
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)",
            "Referer": "https://finance.sina.com.cn/"
        }
        resp = requests.get(url, headers=headers, timeout=5)
        if resp.status_code == 200:
            match = re.search(r'"([^"]*)"', resp.text)
            if match:
                parts = match.group(1).split(",")
                if len(parts) > 1:
                    # 新浪美股指数格式：名称,当前点位,涨跌额,涨跌幅,...
                    return round(float(parts[1]), 2)
    except Exception as e:
        print(f"[纳斯达克-新浪] 获取异常: {e}")
    return None



_cache_stocks = {"data": {}, "last_update": 0}
STOCK_CACHE_SECONDS = 300  # 个股/ETF 缓存 5 分钟

import akshare as ak
import pandas as pd
import time

# 缓存配置
_cache_stocks = {"data": {}, "last_update": 0}
STOCK_CACHE_SECONDS = 300  # 个股/ETF 缓存 5 分钟，避免频繁请求



_cache_stocks = {"data": {}, "last_update": 0}
STOCK_CACHE_SECONDS = 300  # 个股/ETF 缓存 5 分钟

def fetch_global_market_data():
    global _cache_stocks
    result = {}

    # ========== 1. 纳斯达克指数（新浪，保持不变） ==========
    nasdaq = fetch_nasdaq_index()
    if nasdaq is not None:
        result["B24"] = nasdaq

    # ========== 2. A股指数（腾讯） ==========
    index_map = {"sh000001": "B25", "sz399001": "B26", "sz399006": "B27"}
    index_quotes = fetch_tencent_quotes(list(index_map.keys()))
    for code, cell in index_map.items():
        if code in index_quotes:
            result[cell] = round(index_quotes[code], 2)

    # ========== 3. 个股 & ETF（腾讯 + 缓存） ==========
    now = time.time()
    if _cache_stocks["data"] and (now - _cache_stocks["last_update"] < STOCK_CACHE_SECONDS):
        result.update(_cache_stocks["data"])
    else:
        stock_data = {}
        
        # 你的代码列表，转换为腾讯格式（SH/SZ 前缀）
        stock_codes = {
            "sz300212": "B28", "sh600188": "B29", "sh601991": "B30",
            "sh603050": "B31", "sz000681": "B32", "sz000651": "B36", "sz002011": "B37",
        }
        etf_codes = {
            "sh518880": "B33", "sh501227": "B34", "sh513100": "B35",
        }
        
        # 合并所有代码，一次性请求
        all_codes = list({**stock_codes, **etf_codes}.keys())
        quotes = fetch_tencent_quotes(all_codes)
        
        for code, cell in {**stock_codes, **etf_codes}.items():
            if code in quotes:
                stock_data[cell] = round(quotes[code], 2)
        
        if stock_data:
            _cache_stocks["data"] = stock_data
            _cache_stocks["last_update"] = now
            result.update(stock_data)

    return result

def main():
    print("正在连接 Excel...")
    try:
        # 连接当前已经打开的 Excel 文件
        wb = xw.books[EXCEL_FILE_NAME] 
        sheet = wb.sheets.active
        print("Excel 连接成功，开始定时刷新！按 Ctrl+C 停止。")
    except Exception as e:
        print(f"连接 Excel 失败，请确认 '{EXCEL_FILE_NAME}' 已经打开！\n错误信息: {e}")
        return

    while True:
        # 1. 获取数据
        wti, brent = fetch_intl_prices()
        cn_detail = get_cn_oil_df()
        exchange_rate = fetch_exchange_rate()  # 新增：获取汇率，B11
        mid_price, buy_price = fetch_usdcny_mid_and_buy()  # B12 / B13
        dxy = fetch_dxy()                              # B14：美元指数
        intl_gold = fetch_intl_gold_price()          # B16：国际金价
        domestic_gold = fetch_domestic_gold_price()  # B17：国内金价
        y2, y5, y10, y30 = fetch_us_treasury_yields()  # 新增：美债收益率
        market_data = fetch_global_market_data()  # 新增
        
        # 2. 写入 Excel（假设你的数据放在 B2 到 B6 单元格）
        try:
            # 国际油价
            if wti: sheet["B2"].value = wti
            if brent: sheet["B3"].value = brent
            
            # 国内油价
            if cn_detail is not None:
                sheet["B4"].value = get_cn_price(cn_detail, "92")            # 全国92均价
                sheet["B5"].value = get_cn_price(cn_detail, "95")            # 全国95均价
                sheet["B6"].value = get_cn_price(cn_detail, "92", "北京")     # 北京92
                sheet["B7"].value = get_cn_price(cn_detail, "95", "广东")     # 广东95
                sheet["B8"].value = get_cn_price(cn_detail, "92", "天津")     # 天津92
                sheet["B9"].value = get_cn_price(cn_detail, "95", "天津")     # 天津95

            # 汇率类
            if exchange_rate is not None:
                sheet["B11"].value = exchange_rate
            if mid_price is not None:
                sheet["B12"].value = mid_price
            if buy_price is not None:
                sheet["B13"].value = buy_price
            if dxy is not None:
                sheet["B14"].value = dxy

            # 黄金类
            if intl_gold is not None:
                sheet["B16"].value = intl_gold
            if domestic_gold is not None:
                sheet["B17"].value = domestic_gold

            # 美债收益率
            if y2 is not None:
                sheet["B19"].value = y2
            if y5 is not None:
                sheet["B20"].value = y5
            if y10 is not None:
                sheet["B21"].value = y10
            if y30 is not None:
                sheet["B22"].value = y30

            # 批量写入指数、个股、ETF
            # 增加安全判断：确保 market_data 是字典再写入
            if isinstance(market_data, dict) and market_data:
                for cell, price in market_data.items():
                    sheet[cell].value = price

            sheet["D2"].value = time.strftime("%Y-%m-%d %H:%M:%S 已刷新")
            print(f"刷新成功: {time.strftime('%H:%M:%S')}")
        except Exception as e:
            print(f"写入 Excel 失败: {e}")

            # 写入更新时间
            sheet["D2"].value = time.strftime("%Y-%m-%d %H:%M:%S 已刷新")
            print(f"刷新成功: {time.strftime('%H:%M:%S')}")
        except Exception as e:
             print(f"写入 Excel 失败: {e}")

        # 3. 睡眠等待（在此期间 Excel 完全自由，绝不卡顿）
        time.sleep(REFRESH_INTERVAL)

if __name__ == "__main__":
    main()