"""
IPv4 子网规划工具
根据 IP 地址需求，从 37.55.0.0/16 中自动划分子网，生成 Excel 规划表。

规则：
- 基础地址池：37.55.0.0/16
- 可用前缀：/30、/29、/28、/27、/26、/25、/24
- 网关 = 广播地址 - 1
- 可用 IP = 子网内排除网络地址、广播地址、网关后的所有 IP
- VLAN 从 100 开始顺延分配
"""

import ipaddress
import pandas as pd

# 各前缀对应的可用 IP 数 = 总地址数 - 网络地址(1) - 广播地址(1) - 网关(1)
PREFIX_CAPACITY = {
    30: 1,    # 4 - 3
    29: 5,    # 8 - 3
    28: 13,   # 16 - 3
    27: 29,   # 32 - 3
    26: 61,   # 64 - 3
    25: 125,  # 128 - 3
    24: 253,  # 256 - 3
}

def pick_prefix(need: int) -> int:
    """选择能容纳 need 个可用 IP 的最小前缀（最省地址）"""
    for prefix in [30, 29, 28, 27, 26, 25, 24]:
        if PREFIX_CAPACITY[prefix] >= need:
            return prefix
    raise ValueError(
        f"IP 需求 {need} 超出单个 /24 子网容量（最大 253），请拆分为多个子网"
    )

def build_plan(ip_needs, base="37.55.0.0/16", vlan_start=100):
    base_net = ipaddress.ip_network(base)
    pool_start = int(base_net.network_address)
    pool_end = int(base_net.broadcast_address)

    current = pool_start
    rows = []

    for idx, need in enumerate(ip_needs):
        if need <= 0:
            raise ValueError(f"第 {idx+1} 项 IP 需求 {need} 不合法，必须为正整数")
        prefix = pick_prefix(need)
        size = 2 ** (32 - prefix)

        # 对齐到子网边界（例如 /28 必须从 16 的倍数开始）
        if current % size != 0:
            current = ((current // size) + 1) * size

        network = ipaddress.ip_network((ipaddress.ip_address(current), prefix))

        if int(network.broadcast_address) > pool_end:
            raise ValueError(f"地址池 {base} 已耗尽，无法分配第 {idx+1} 个子网")

        broadcast = network.broadcast_address
        gateway = ipaddress.ip_address(int(broadcast) - 1)

        # 可用 IP：排除网络地址、广播地址、网关
        available = [ip for ip in network.hosts() if ip != gateway]

        if len(available) == 1:
            avail_str = str(available[0])
        else:
            avail_str = f"{available[0]} ~ {available[-1]}（共 {len(available)} 个）"

        rows.append({
            "IP地址需求": need,
            "建议网段": str(network),
            "网络地址": str(network.network_address),
            "广播地址": str(broadcast),
            "子网掩码": str(network.netmask),
            "网关": str(gateway),
            "可用IP地址": avail_str,
            "vlan": vlan_start + idx,
        })

        current = int(broadcast) + 1

    columns = ["IP地址需求", "建议网段", "网络地址", "广播地址",
               "子网掩码", "网关", "可用IP地址", "vlan"]
    return pd.DataFrame(rows, columns=columns)

def main():
    # 从输入 Excel 读取 IP 需求
    input_file = "input.xlsx"          # 输入文件路径
    input_df = pd.read_excel(input_file)

    # 校验列名
    if "IP地址需求" not in input_df.columns:
        raise ValueError("输入文件缺少列：IP地址需求")

    # 提取需求列表，去除空值
    ip_needs = pd.to_numeric(
        input_df["IP地址需求"], errors="coerce"
    ).dropna().astype(int).tolist()

    if not ip_needs:
        raise ValueError("输入文件中没有有效的 IP 地址需求")

    # 生成规划表
    df = build_plan(ip_needs, base="37.55.0.0/16", vlan_start=100)

    output = "ip_plan.xlsx"
    df.to_excel(output, index=False)
    print(f"已生成 Excel 文件：{output}\n")
    print(df.to_string(index=False))

if __name__ == "__main__":
    main()