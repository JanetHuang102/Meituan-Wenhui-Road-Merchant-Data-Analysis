import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from sklearn.cluster import DBSCAN
import re
import os
import requests
from io import BytesIO
from PIL import Image
import matplotlib.patches as mpatches

# 设置中文字体
plt.rcParams['font.sans-serif'] = ['SimHei', 'Microsoft YaHei', 'DejaVu Sans']
plt.rcParams['axes.unicode_minus'] = False

# 高德地图API Key
AMAP_API_KEY = "3b7e8ba36b7ba80177cfe3066f223918"


def read_excel_data(file_path):
    """读取Excel文件中的经纬度数据"""
    df = pd.read_excel(file_path)

    # 获取E列数据
    coord_data = df.iloc[:, 4]  # 索引4对应E列

    # 解析经纬度字符串
    lons, lats = [], []
    row_count = 0
    for idx, coord_str in enumerate(coord_data, start=2):
        if pd.notna(coord_str):
            # 使用正则表达式提取数字
            coords = re.findall(r'-?\d+\.\d+', str(coord_str))
            if len(coords) == 2:
                lons.append(float(coords[0]))
                lats.append(float(coords[1]))
                row_count += 1
            else:
                print(f"警告: 第{idx}行无法解析坐标字符串: {coord_str}")
        else:
            print(f"警告: 第{idx}行为空")

    print(f"成功解析 {row_count} 个有效坐标")
    return lons, lats


def cluster_points(lons, lats, eps=0.001, min_samples=2):
    """使用DBSCAN进行空间聚类"""
    coords = np.array(list(zip(lons, lats)))
    dbscan = DBSCAN(eps=eps, min_samples=min_samples)
    clusters = dbscan.fit_predict(coords)
    return clusters


def assign_cluster_names(lons, lats, clusters):
    """为聚类分配有意义的名称"""
    unique_clusters = np.unique(clusters)
    cluster_names = {}
    cluster_info = {}

    # 计算每个聚类的中心点
    for cluster_id in unique_clusters:
        if cluster_id == -1:
            cluster_names[cluster_id] = "零散分布"
            cluster_info[cluster_id] = {
                'center': (np.mean(lons), np.mean(lats)),
                'count': np.sum(clusters == cluster_id)
            }
            continue

        mask = clusters == cluster_id
        cluster_lons = np.array(lons)[mask]
        cluster_lats = np.array(lats)[mask]

        center_lon = np.mean(cluster_lons)
        center_lat = np.mean(cluster_lats)

        # 根据地理位置特征分配名称
        if center_lat > 31.06:  # 北部区域
            region = "北段"
        else:  # 南部区域
            region = "南段"

        if center_lon < 121.195:  # 西侧
            direction = "西侧"
        elif center_lon < 121.205:  # 中段
            direction = "中段"
        else:  # 东侧
            direction = "东侧"

        # 根据聚类大小添加规模描述
        count = np.sum(mask)
        if count > 20:
            size_desc = "密集区"
        elif count > 10:
            size_desc = "集中区"
        else:
            size_desc = "小聚集区"

        name = f"文汇路{region}{direction}{size_desc}"

        # 添加序号，避免重名
        base_name = name
        counter = 1
        while name in cluster_names.values():
            name = f"{base_name}({counter})"
            counter += 1

        cluster_names[cluster_id] = name
        cluster_info[cluster_id] = {
            'center': (center_lon, center_lat),
            'count': count,
            'lons': cluster_lons,
            'lats': cluster_lats,
            'min_lat': min(cluster_lats),
            'max_lat': max(cluster_lats)
        }

    return cluster_names, cluster_info


def get_amap_background(lons, lats, zoom=15, width=800, height=600):
    """使用高德地图API获取地图背景"""
    try:
        # 计算中心点
        center_lon = np.mean(lons)
        center_lat = np.mean(lats)

        print(f"地图中心点: 经度 {center_lon}, 纬度 {center_lat}")

        # 高德地图静态地图API
        url = "https://restapi.amap.com/v3/staticmap"
        params = {
            'location': f"{center_lon},{center_lat}",
            'zoom': zoom,
            'size': f"{width}*{height}",
            'key': AMAP_API_KEY,
            'scale': 2
        }

        print(f"请求高德地图API: {url}")
        response = requests.get(url, params=params, timeout=10)
        if response.status_code == 200:
            img = Image.open(BytesIO(response.content))
            print("成功获取高德地图背景")

            # 计算地图边界
            # 高德地图的经纬度范围计算
            # 使用近似公式计算范围
            lon_degree_per_pixel = 360 / (2 ** (zoom + 8))
            lat_degree_per_pixel = lon_degree_per_pixel * 0.8

            lon_range = width * lon_degree_per_pixel / 2
            lat_range = height * lat_degree_per_pixel / 2

            map_extent = [
                center_lon - lon_range,
                center_lon + lon_range,
                center_lat - lat_range,
                center_lat + lat_range
            ]

            print(f"地图范围: {map_extent}")
            return img, map_extent
        else:
            print(f"高德地图API请求失败，状态码: {response.status_code}")
            return None, None
    except Exception as e:
        print(f"获取高德地图背景时出错: {e}")
        return None, None


def create_waterdrop_marker(ax, lon, lat, color, size=0.00065, alpha=0.9):
    """创建水滴形状标记的辅助函数"""
    # 创建倒水滴形状（三角形底座）
    triangle = mpatches.RegularPolygon(
        (lon, lat),
        numVertices=3,
        radius=size,  # 大小
        orientation=np.pi,  # 倒三角形
        facecolor=color,
        alpha=alpha,
        edgecolor='white',
        linewidth=1
    )
    ax.add_patch(triangle)

    # 在上方偏移位置添加白色圆形
    circle = mpatches.Circle(
        (lon, lat + size * 0.6),  # 向上偏移
        radius=size * 0.4,  # 大小
        facecolor='white',
        alpha=alpha,
        edgecolor=color,  # 边框颜色与三角形相同
        linewidth=2  # 边框厚度
    )
    ax.add_patch(circle)

    return triangle, circle


def create_static_map(lons, lats, clusters, cluster_names, cluster_info, output_file='文汇路店铺分布图.png'):
    """创建静态地图（使用高德地图底图）"""
    # 创建图形
    fig, ax = plt.subplots(1, 1, figsize=(16, 12))

    # 打印坐标范围用于调试
    print(f"经度范围: {min(lons):.6f} - {max(lons):.6f}")
    print(f"纬度范围: {min(lats):.6f} - {max(lats):.6f}")

    # 尝试获取高德地图背景
    map_img, map_extent = get_amap_background(lons, lats, zoom=15, width=1600, height=1200)

    if map_img is not None and map_extent is not None:
        ax.imshow(map_img, extent=map_extent, alpha=0.9)
        ax.set_xlim(map_extent[0], map_extent[1])
        ax.set_ylim(map_extent[2], map_extent[3])
        background_used = "高德地图背景"
        print("使用高德地图背景")
    else:
        # 如果没有地图背景，使用常规坐标范围
        padding = 0.005
        x_min, x_max = min(lons) - padding, max(lons) + padding
        y_min, y_max = min(lats) - padding, max(lats) + padding
        ax.set_xlim(x_min, x_max)
        ax.set_ylim(y_min, y_max)
        background_used = "空白背景"
        print(f"使用空白背景，坐标范围: X({x_min:.6f}, {x_max:.6f}), Y({y_min:.6f}, {y_max:.6f})")

    # 按聚类分组绘制点 - 使用图片中的样式
    unique_clusters = np.unique(clusters)

    # 定义颜色方案：蓝色、紫色、黄色、绿色、红色（用于零散分布）
    colors = ['#4285F4','#9C27B0','#FBBC05', '#34A853','#EA4335']

    # 存储标记对象用于图例
    legend_elements = []

    for i, cluster_id in enumerate(unique_clusters):
        mask = clusters == cluster_id
        cluster_lons = np.array(lons)[mask]
        cluster_lats = np.array(lats)[mask]

        # 分配颜色
        if cluster_id == -1:  # 零散分布使用灰色
            color = 'gray'
        else:  # 其他聚类使用预设颜色
            color_idx = (cluster_id - 1) % len(colors)  # 确保颜色索引不越界
            color = colors[color_idx]

        label = f'{cluster_names[cluster_id]} ({np.sum(mask)}个{"店铺" if cluster_id != -1 else ""})'

        # 使用图片中的样式：红色倒水滴形状，中间有白色圆形
        # 对于每个点，绘制红色倒水滴标记
        for lon, lat in zip(cluster_lons, cluster_lats):
            # 创建倒水滴形状（红色三角形底座）- 放大5倍
            triangle = mpatches.RegularPolygon(
                (lon, lat),
                numVertices=3,
                radius=0.0006,  # 从0.00015放大到0.00075（5倍）
                orientation=np.pi,  # 倒三角形
                facecolor=color,
                alpha=0.9,
            )
            ax.add_patch(triangle)

            # 在中间添加白色圆形 - 放大5倍
            circle = mpatches.Circle(
                (lon, lat+0.0003),
                radius=0.00035,  # 从0.00005放大到0.00025（5倍）
                facecolor='white',
                edgecolor=color,
                alpha=0.9,
                linewidth=4
            )
            ax.add_patch(circle)

        # 创建自定义图例标记
        from matplotlib.lines import Line2D
        legend_marker = Line2D([0], [0],
                               marker='.',  # 使用倒三角形模拟水滴
                               color='white',
                               markerfacecolor=color,
                               markersize=15,
                               label=label)
        legend_elements.append(legend_marker)

        # 打印每个聚类的坐标范围
        print(f"聚类 {cluster_id} ({cluster_names[cluster_id]}): {len(cluster_lons)}个点")
        if len(cluster_lons) > 0:
            print(f"  经度范围: {min(cluster_lons):.6f} - {max(cluster_lons):.6f}")
            print(f"  纬度范围: {min(cluster_lats):.6f} - {max(cluster_lats):.6f}")

    # 设置图形属性
    ax.set_title('上海松江大学城文汇路店铺分布聚类图（高德地图底图）', fontsize=18, fontweight='bold', pad=20)
    ax.set_xlabel('经度', fontsize=12)
    ax.set_ylabel('纬度', fontsize=12)
    ax.grid(True, alpha=0.3)

    # 添加聚类中心标注
    for cluster_id in unique_clusters:
        if cluster_id == -1:
            continue

        center_lon, center_lat = cluster_info[cluster_id]['center']
        count = cluster_info[cluster_id]['count']
        name = cluster_names[cluster_id]

        # 根据聚类位置调整标注位置
        if "北段" in name:
            text_y = cluster_info[cluster_id]['min_lat'] - 0.0008
            va = 'top'
        else:
            text_y = cluster_info[cluster_id]['max_lat'] + 0.0008
            va = 'bottom'

        # 添加标注框
        ax.annotate(f'{name}\n({count}家)',
                    xy=(center_lon, center_lat),
                    xytext=(center_lon, text_y),
                    ha='center', va=va,
                    fontsize=10,
                    bbox=dict(boxstyle="round,pad=0.3",
                              facecolor='wheat',
                              edgecolor='orange',
                              alpha=0.9),
                    arrowprops=dict(arrowstyle="->",
                                    connectionstyle="arc3,rad=0.1",
                                    color='orange', alpha=0.7))

    # 添加图例
    ax.legend(handles=legend_elements, loc='upper left', bbox_to_anchor=(0, 1), fontsize=9,
              title="聚类区域", title_fontsize=11)

    # 添加统计信息文本框
    total_shops = len(lons)
    noise_shops = np.sum(clusters == -1)
    clustered_shops = total_shops - noise_shops
    cluster_count = len(unique_clusters) - (1 if -1 in unique_clusters else 0)

    stats_text = f"统计信息:\n总店铺数: {total_shops}\n聚类区域数: {cluster_count}\n"
    stats_text += f"已聚类店铺: {clustered_shops}\n零散店铺: {noise_shops}\n"
    stats_text += f"底图: {background_used}"

    ax.text(0.89, 0.02, stats_text, fontsize=10,
            transform=ax.transAxes,
            bbox=dict(boxstyle="round", facecolor='wheat', edgecolor='orange', alpha=0.9),
            verticalalignment='bottom', horizontalalignment='left')

    # 调整布局
    plt.tight_layout()

    # 保存图片
    plt.savefig(output_file, dpi=300, bbox_inches='tight', facecolor='white')
    print(f"地图已保存为: {output_file}")
    plt.show()

    return fig


def generate_cluster_report(cluster_names, cluster_info, output_file='聚类分析报告.txt'):
    """生成详细的聚类分析报告"""
    with open(output_file, 'w', encoding='utf-8') as f:
        f.write("上海松江大学城文汇路店铺聚类分析报告\n")
        f.write("=" * 50 + "\n\n")

        # 总体统计
        total_clusters = len([k for k in cluster_names.keys() if k != -1])
        total_shops = sum(info['count'] for info in cluster_info.values())

        f.write(f"总体统计:\n")
        f.write(f"- 聚类区域数量: {total_clusters}\n")
        f.write(f"- 总店铺数量: {total_shops}\n")
        f.write(f"- 零散店铺数量: {cluster_info.get(-1, {}).get('count', 0)}\n\n")

        f.write("各聚类区域详情:\n")
        f.write("-" * 30 + "\n")

        # 按店铺数量排序
        sorted_clusters = sorted(
            [(k, v) for k, v in cluster_info.items() if k != -1],
            key=lambda x: x[1]['count'], reverse=True
        )

        for i, (cluster_id, info) in enumerate(sorted_clusters, 1):
            f.write(f"\n{i}. {cluster_names[cluster_id]}:\n")
            f.write(f"   - 店铺数量: {info['count']}家\n")
            f.write(f"   - 中心位置: 经度 {info['center'][0]:.6f}, 纬度 {info['center'][1]:.6f}\n")
            if 'lons' in info:
                f.write(f"   - 经度范围: {min(info['lons']):.6f} - {max(info['lons']):.6f}\n")
                f.write(f"   - 纬度范围: {min(info['lats']):.6f} - {max(info['lats']):.6f}\n")

        # 零散店铺信息
        if -1 in cluster_info:
            f.write(f"\n零散分布区域:\n")
            f.write(f"- 店铺数量: {cluster_info[-1]['count']}家\n")
            f.write(f"- 说明: 这些店铺分布较为分散，未形成明显的聚集区域\n")

    print(f"聚类分析报告已保存为: {output_file}")


def main():
    """主函数"""
    excel_file = r"D:\Desktop\爬虫\店铺数据.xlsx"

    try:
        # 1. 读取数据
        print("正在读取Excel数据...")
        lons, lats = read_excel_data(excel_file)
        print(f"成功读取 {len(lons)} 个店铺坐标")

        if len(lons) == 0:
            print("未找到有效的经纬度数据，请检查Excel文件格式")
            return

        # 2. 空间聚类
        print("正在进行空间聚类...")
        clusters = cluster_points(lons, lats, eps=0.002, min_samples=3)

        # 3. 为聚类分配名称
        print("正在为聚类区域分配名称...")
        cluster_names, cluster_info = assign_cluster_names(lons, lats, clusters)

        # 统计聚类结果
        unique_clusters, counts = np.unique(clusters, return_counts=True)
        print("\n聚类结果统计:")
        for cluster_id, count in zip(unique_clusters, counts):
            if cluster_id == -1:
                print(f"{cluster_names[cluster_id]}: {count}个")
            else:
                print(f"{cluster_names[cluster_id]}: {count}个店铺")

        # 4. 创建地图
        print("正在生成地图（使用高德地图底图）...")
        create_static_map(lons, lats, clusters, cluster_names, cluster_info, '文汇路店铺分布图_高德地图.png')

        # 5. 生成分析报告
        generate_cluster_report(cluster_names, cluster_info)

        # 打开保存的图片
        if os.path.exists('文汇路店铺分布图_高德地图.png'):
            os.startfile('文汇路店铺分布图_高德地图.png')

    except Exception as e:
        print(f"处理过程中出现错误: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main()