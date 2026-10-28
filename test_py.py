import akshare as ak

try:
    # 1. 测试获取历史调价数据
    hist_df = ak.energy_oil_hist()
    if hist_df is None or hist_df.empty:
        print("❌ energy_oil_hist 返回为空")
    else:
        print("✅ energy_oil_hist 获取成功")
        print("最新日期列名:", hist_df.columns.tolist())
        print("最新一行数据:\n", hist_df.iloc[0])

        # 2. 提取最新日期（注意兼容不同格式）
        raw_date = hist_df.iloc[0]["调整日期"]
        date_str = str(raw_date).replace("-", "").replace("/", "")[:8]
        print(f"提取到的日期字符串: {date_str}")

        # 3. 测试获取明细数据
        detail_df = ak.energy_oil_detail(date=date_str)
        if detail_df is None or detail_df.empty:
            print("❌ energy_oil_detail 返回为空")
        else:
            print("✅ energy_oil_detail 获取成功")
            print("明细数据前几行:\n", detail_df.head())

except Exception as e:
    print(f"❌ 发生异常: {e}")