import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import os


# 描述性分析：计算每个分组的指标统计量
def descriptive_analysis(group, category, indicator_start_col):
    # 提取指标列
    indicator_df = group.iloc[:, indicator_start_col:]
    # 计算统计量
    shop_count = group.shape[0]
    stats = indicator_df.describe()

    # 转置矩阵（以各指标为行数据，统计量为列数据）
    transposed_stats = stats.T.reset_index().rename(columns={"index": "指标名称"})
    # 增加店铺分类列
    transposed_stats["店铺分类"] = category

    # 以店铺分类为索引，此时一级列标签为统计量，二级列标签为指标名称
    desc_stats = transposed_stats.pivot(index="店铺分类",columns="指标名称").reset_index()
    # 交换列标签层级
    desc_stats.columns = desc_stats.columns.swaplevel(0, 1)
    # 按第一层排序，确保同一指标的统计量相邻
    desc_stats = desc_stats.sort_index(axis=1, level=0)
    # 调整列顺序：先店铺总数，再各指标数据
    desc_stats.insert(1, ("店铺总数", " "), shop_count)
    
    return desc_stats


# 使用熵权TOPSIS模型（欧式距离）计算店铺排名（只含指标数据）
def topsis(indicator_data, positive_indices, negative_indices):
    """
    TOPSIS函数：仅接收指标列数据
    :param indicator_data: 纯指标的矩阵（无文本）
    :param positive_indices: 正向指标在指标矩阵中的索引
    :param negative_indices: 负向指标在指标矩阵中的索引
    :return: 含贴近度的结果
    """
    # 获取样本数n和指标数m
    n, m = indicator_data.shape
    # 初始化矩阵
    data_std = np.zeros_like(indicator_data, dtype=float)

    # 标准化每个指标
    for j in range(m):
        col = indicator_data[:, j]
        max_val = np.max(col)
        min_val = np.min(col)
        # 如果指标无差异，则全部赋值1
        if max_val == min_val:
            data_std[:, j] = 1.0
            continue
        # 处理正向指标
        if j in positive_indices:
            data_std[:, j] = (col - min_val) / (max_val - min_val)
        # 处理负向指标
        elif j in negative_indices:
            data_std[:, j] = (max_val - col) / (max_val - min_val)

    # 计算权重
    # 避免后续log(0)，整体增加微小值
    data_std +=  1e-6
    # 将每个指标下的样本值转化为比重
    p = data_std / np.sum(data_std, axis=0)
    # 计算熵值（熵值越大，差异越小）
    e = -np.sum(p * np.log(p), axis=0) / np.log(n)
    # 计算差异系数
    g = 1 - e
    # 差异系数归一化/最终权重
    w = g / np.sum(g)

    # 加权矩阵
    weight_std_data = data_std * w

    # 确定正理想解（最优方案）和负理想解（最差方案）
    # 正理想解 = 各指标最大值
    z_positive = np.max(weight_std_data, axis=0)
    # 负理想解 = 各指标最小值
    z_negative = np.min(weight_std_data, axis=0)

    # 计算欧式距离
    d_positive = np.sqrt(np.sum((weight_std_data - z_positive) ** 2, axis=1))
    d_negative = np.sqrt(np.sum((weight_std_data - z_negative) ** 2, axis=1))
    #计算贴近度
    closeness = d_negative / (d_positive + d_negative)

    # 返回含贴近度的结果（后续需与非指标列整合）
    return pd.DataFrame({
        "到正理想解距离": d_positive.round(3),
        "到负理想解距离": d_negative.round(3),
        "贴近度": closeness.round(3)
        })


# 按组处理数据
def group_sort(group, category, indicator_start_col, positive_indices, negative_indices):
    # 提取指标矩阵
    non_indicator_cols = group.iloc[:, :indicator_start_col].copy()
    indicator_data = group.iloc[:, indicator_start_col:].values
    # 检查指标列是否为数值型
    if not np.issubdtype(indicator_data.dtype, np.number):
        raise ValueError(f"分类 {category} 的指标列包含非数值数据")

    # 调用TOPSIS函数
    result = topsis(indicator_data, positive_indices, negative_indices)

    # 对齐索引
    result.index = non_indicator_cols.index
    # 整合非指标列与计算结果
    group_result = pd.concat([non_indicator_cols, result], axis=1)
    # 计算排名
    group_result["排名"] = group_result["贴近度"].rank(method="min", ascending=False).astype(int)
    sort_group_result = group_result.sort_values("排名")

    return sort_group_result


# 可视化排名：横向柱状图
def group_barh(sort_group_result, save_folder, category):
    plt.rcParams["font.family"] = ["SimHei", "WenQuanYi Micro Hei", "Heiti TC"]
    plt.rcParams["axes.unicode_minus"] = False
    top_three = sort_group_result.head(3)[::-1]
    plt.figure(figsize=(8, 5))
    x = top_three["店铺名称"]
    y = top_three["贴近度"]
    plt.barh(x, y)
    title = f"{category}排名前3的店铺"
    plt.title(title)
    plt.xlabel("综合得分")
    plt.ylabel("店铺名称")
    plt.tight_layout()
    save_path_plot = os.path.join(save_folder, f"{title}.png")
    plt.savefig(save_path_plot, dpi=300, bbox_inches="tight")
    plt.show()
    plt.close()
    return print(f"已保存图片：{save_path_plot}")


# 主流程（需要分离文本列/非指标列）
def main(excel_path, sheet_name, category_col, indicator_start_col,
         positive_indices, negative_indices, save_folder,
         desc_save_path, sort_save_path):
    # 检查文件夹是否存在
    os.makedirs(save_folder, exist_ok=True)
    # 读取Excel数据
    try:
        df = pd.read_excel(excel_path, sheet_name=sheet_name)
        print(f"文件读取成功！共读取 {df.shape[0]} 行数据，{df.shape[1]} 列数据。")
        print(df.head())
    except Exception as e:
        raise ValueError(f"文件读取失败：{str(e)}")

    # 分组处理数据
    all_results = []
    all_stats = []
    for category, group in df.groupby(category_col, group_keys=True):
        print(f"\n处理分类：{category}")
        # 描述性分析
        group_stats =descriptive_analysis(group, category, indicator_start_col)
        all_stats.append(group_stats)
        # 排名结果
        sort_group_result = group_sort(group, category, indicator_start_col, positive_indices, negative_indices)
        all_results.append(sort_group_result)
        # 可视化情况
        group_barh(sort_group_result, save_folder, category)
    # 合并所有分类结果
    desc_final = pd.concat(all_stats, ignore_index=True)
    final_result = pd.concat(all_results, ignore_index=True)

    # 保存描述性分析结果
    try:
        desc_final.to_excel(desc_save_path, index=True)
        print(f"描述性分析结果已成功保存到：{desc_save_path}")
    except Exception as e:
        raise ValueError(f"描述性分析结果保存失败：{str(e)}")
    
    #保存排名结果
    try:
        final_result.to_excel(sort_save_path, index=False)
        print(f"结果已成功保存到：{sort_save_path}")
    except Exception as e:
        raise ValueError(f"结果保存失败：{str(e)}")
    return final_result


if __name__ == "__main__":
    # 读入Excel路径
    EXCEL_PATH = "./店铺数据.xlsx"
    # 工作表名称
    SHEET_NAME = "Sheet1"
    # 店铺分类索引
    CATEGORY_COL = "店铺分类"
    # 指标列起始位置索引
    INDICATOR_START_COL = 5
    # 正向指标位置索引
    POSITIVE_INDICES = [0, 2, 3, 4]
    # 负向指标位置索引
    NEGATIVE_INDICES = [1, 5]
    # 图片保存路径
    SAVE_FOLDER = "./店铺可视化"
    # 描述性分析保存路径
    DESC_SAVE_PATH = "./描述性分析结果.xlsx"
    # 排名结果保存路径
    SORT_SAVE_PATH = "./最终店铺排名.xlsx"
    # 执行主流程
    final_result = main(EXCEL_PATH, SHEET_NAME, CATEGORY_COL,
                        INDICATOR_START_COL, POSITIVE_INDICES, NEGATIVE_INDICES,
                        SAVE_FOLDER, DESC_SAVE_PATH, SORT_SAVE_PATH)

