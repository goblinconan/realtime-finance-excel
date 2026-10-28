import xlwings as xw
import akshare as ak
import pandas as pd
import time

# ================= 缓存机制 =================
_cache = {
    "cn_detail_df": None,
    "last_update": 0
}
CACHE_EXPIRE_SECONDS = 3600 * 12  # 缓存有效期：12小时（国内油价每天变一次足够）

@xw.func
def GET_OIL_PRICE(oil_type="WTI"):
    """获取国际油价（不缓存，因为波动快，但不设置为 volatile）"""
    symbol_map = {"WTI": "CL", "BRENT": "OIL"}
    symbol = symbol_map.get(oil_type.upper(), "CL")
    try:
        df = ak.futures_foreign_commodity_realtime(symbol=[symbol])
        price = df.iloc[0].get("current_price") or df.iloc[0].get("最新价")
        return round(float(price), 3)
    except Exception as e:
        return f"获取失败：{e}"

@xw.func
def GET_CN_GASOLINE_PRICE(grade="92", region=None):
    """获取国内油价（带缓存机制）"""
    global _cache
    grade_str = str(grade).strip().replace("#", "")
    if grade_str not in ("92", "95"):
        return "参数须为 92 或 95"

    try:
        # 1. 检查缓存是否有效
        now = time.time()
        if _cache["cn_detail_df"] is None or (now - _cache["last_update"] > CACHE_EXPIRE_SECONDS):
            # 缓存失效，重新联网获取
            hist_df = ak.energy_oil_hist()
            if hist_df is None or hist_df.empty:
                return "获取失败：历史数据为空"
            
            hist_df["调整日期"] = pd.to_datetime(hist_df["调整日期"], errors="coerce")
            hist_df = hist_df.dropna(subset=["调整日期"]).sort_values("调整日期", ascending=False)
            latest_date = hist_df.iloc[0]["调整日期"].strftime("%Y%m%d")

            detail_df = ak.energy_oil_detail(date=latest_date)
            if detail_df is None or detail_df.empty:
                return f"获取失败：{latest_date} 明细为空"

            # 处理数据源 92/95 反了的问题
            col_92 = next((c for c in ["V_92", "V92"] if c in detail_df.columns), None)
            col_95 = next((c for c in ["V_95", "V95"] if c in detail_df.columns), None)
            if col_92 and col_95:
                if pd.to_numeric(detail_df[col_92], errors="coerce").mean() > pd.to_numeric(detail_df[col_95], errors="coerce").mean():
                    detail_df = detail_df.rename(columns={col_92: col_95, col_95: col_92})

            # 更新缓存
            _cache["cn_detail_df"] = detail_df
            _cache["last_update"] = now

        # 2. 直接从缓存中读取
        detail_df = _cache["cn_detail_df"]

        # 3. 提取目标列并返回结果（此处省略了找列名和区域过滤的代码，保持原有逻辑）
        target_col = next((c for c in [f"V_{grade_str}", f"V{grade_str}"] if c in detail_df.columns), None)
        if target_col is None:
            return f"未找到 {grade_str}# 汽油价格列"

        if region is None or str(region).strip() == "" or pd.isna(region):
            valid_prices = pd.to_numeric(detail_df[target_col], errors="coerce").dropna()
            return round(float(valid_prices.mean()), 2)
        else:
            region_str = str(region).strip()
            region_col = next((c for c in ["省份", "地区", "城市"] if c in detail_df.columns), None)
            matched = detail_df[detail_df[region_col].astype(str).str.contains(region_str, na=False)]
            if matched.empty:
                return f"未找到地区：{region_str}"
            price = pd.to_numeric(matched.iloc[0][target_col], errors="coerce")
            return round(float(price), 2)

    except Exception as e:
        return f"获取失败：{e}"