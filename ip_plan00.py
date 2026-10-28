import ipaddress
import pandas as pd

def find_prefix(ip_needed):
    """根据 IP 需求找到合适的前缀（可用 IP >= ip_needed）"""
    # 从可用 IP 最少的开始（最节省）
    for prefix in [30, 29, 28, 27, 26, 25, 24]:
        total = 2 ** (32 - prefix)
        available = total - 3  # 减去网络地址、广播地址、网关
        if available >= ip_needed:
            return prefix
    return None  # 超出 /24 容量

def generate_excel(ip_needs, output_file='ip_plan.xlsx'):
    base_network = ipaddress.ip_network('37.55.0.0/16')
    current = int(base_network.network_address)
    vlan_start = 100
    
    rows = []
    for i, need in enumerate(ip_needs, 1):
        prefix = find_prefix(need)
        if prefix is None:
            continue
        
        # 对齐到子网边界
        subnet_size = 2 ** (32 - prefix)
        # 找到 >= current 且对齐的下一个地址
        current = ((current + subnet_size - 1) // subnet_size) * subnet_size
        
        net_addr = ipaddress.ip_address(current)
        network = ipaddress.ip_network(f"{net_addr}/{prefix}")
        broadcast = network.broadcast_address
        
        # 网关 = 广播地址 - 1
        gateway = ipaddress.ip_address(int(broadcast) - 1)
        
        # 可用 IP：排除网络地址、广播地址、网关
        available_ips = []
        for ip in network.hosts():
            if ip != gateway:
                available_ips.append(str(ip))
        
        # 只显示前几个，或者压缩表示
        if len(available_ips) <= 5:
            avail_str = ', '.join(available_ips)
        else:
            avail_str = f"{available_ips[0]} ~ {available_ips[-1]} (共 {len(available_ips)} 个)"
        
        rows.append({
            'IP地址需求': need,
            '建议网段': str(network),
            '网络地址': str(network.network_address),
            '广播地址': str(broadcast),
            '子网掩码': str(network.netmask),
            '网关': str(gateway),
            '可用IP地址': avail_str,
            'vlan': vlan_start + i - 1
        })
        
        current += subnet_size
    
    df = pd.DataFrame(rows)
    # 指定列顺序
    columns = ['IP地址需求', '建议网段', '网络地址', '广播地址', '子网掩码', '网关', '可用IP地址', 'vlan']
    df = df[columns]
    df.to_excel(output_file, index=False)
    print(f"已生成 {output_file}")
    return df

if __name__ == '__main__':
    ip_needs = [1, 5, 10, 20, 50, 100, 200]
    df = generate_excel(ip_needs)
    print(df)