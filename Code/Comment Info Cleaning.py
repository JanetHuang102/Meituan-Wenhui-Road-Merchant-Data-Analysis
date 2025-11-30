import pandas as pd
import re
import os
from snownlp import SnowNLP
from datetime import datetime
import warnings

warnings.filterwarnings('ignore')


def clean_text(text):
    """清洗评论内容，只保留中英文字符和中英文标点符号"""
    if pd.isna(text):
        return ""

    # 移除[]标记的表情符号
    text = re.sub(r'\[.*?\]', '', str(text))

    # 定义保留的字符：中文、英文、数字、空格和常见中英文标点
    # 中文标点：。，！？；："'（）《》【】·
    # 英文标点：.,!?;:'"(){}[]
    pattern = r'[^\u4e00-\u9fa5a-zA-Z0-9\s。，！？；："\'（）《》【】·\.\,\!\?\;\:\'\"\(\)\{\}\[\]］]'
    text = re.sub(pattern, '', str(text))

    # 移除多余的空格
    text = re.sub(r'\s+', ' ', text).strip()

    return text


def analyze_sentiment(text):
    """使用SnowNLP进行情感分析，返回0-1的得分"""
    if not text or text.strip() == "":
        return 0.5  # 空文本返回中性分数

    try:
        s = SnowNLP(text)
        return round(s.sentiments, 4)
    except:
        return 0.5  # 分析失败返回中性分数


def process_shop_data(df, shop_name):
    """处理单个店铺的数据"""
    # 确保评论时间是日期格式
    df['评论时间'] = pd.to_datetime(df['评论时间'], format='%Y/%m/%d', errors='coerce')

    # 清洗评论内容
    df['评论内容_清洗后'] = df['评论内容'].apply(clean_text)

    # 情感分析
    df['情感得分'] = df['评论内容_清洗后'].apply(analyze_sentiment)

    # 按用户名分组，计算回头客信息
    df = df.sort_values(['用户名', '评论时间'])  # 按用户和时间排序

    # 计算每个用户的评论次数和回头客信息
    user_stats = df.groupby('用户名').size().reset_index(name='总评论数')
    user_first_review = df.groupby('用户名')['评论时间'].min().reset_index(name='首次评论时间')

    # 合并用户统计信息
    user_info = pd.merge(user_stats, user_first_review, on='用户名')

    # 为每条评论添加序号
    df['评论序号'] = df.groupby('用户名').cumcount() + 1

    # 计算是否回头客和回头次数
    df['是否回头客'] = df['评论序号'] > 1
    df['回头次数'] = df['评论序号'] - 1

    # 选择需要的列并重新排序
    result_df = df[['用户名', '评分', '评论内容_清洗后', '评论时间', '情感得分', '是否回头客', '回头次数']]

    # 重命名列以匹配输出要求
    result_df = result_df.rename(columns={'评论内容_清洗后': '评论内容'})

    # 格式化评论时间为字符串（保持原始格式）
    result_df['评论时间'] = result_df['评论时间'].dt.strftime('%Y/%m/%d')

    return result_df


def main():
    # 文件夹路径
    input_folder = r"D:\Desktop\爬虫\美团\待清洗"
    output_folder = r"D:\Desktop\爬虫\美团\清洗后"

    # 确保输出文件夹存在
    os.makedirs(output_folder, exist_ok=True)

    # 获取所有Excel文件
    excel_files = [f for f in os.listdir(input_folder) if f.endswith(('.xlsx', '.xls'))]

    if not excel_files:
        print("在输入文件夹中未找到Excel文件")
        return

    print(f"找到 {len(excel_files)} 个Excel文件，开始处理...")

    # 处理每个文件
    for file_name in excel_files:
        try:
            print(f"正在处理文件: {file_name}")

            # 读取Excel文件
            file_path = os.path.join(input_folder, file_name)
            df = pd.read_excel(file_path)

            # 检查必要的列是否存在
            required_columns = ['用户名', '评分', '评论内容', '评论时间']
            if not all(col in df.columns for col in required_columns):
                print(f"文件 {file_name} 缺少必要的列，跳过处理")
                continue

            # 处理数据
            processed_df = process_shop_data(df, file_name)

            # 保存处理后的文件
            output_path = os.path.join(output_folder, f"{file_name}")
            processed_df.to_excel(output_path, index=False)

            # 打印店铺统计信息
            total_reviews = len(processed_df)
            returning_customers = processed_df['是否回头客'].sum()
            total_return_visits = processed_df['回头次数'].sum()
            avg_sentiment = processed_df['情感得分'].mean()
            avg_rating = processed_df['评分'].mean()

            print(f"店铺 {file_name} 统计:")
            print(f"  - 总评论数: {total_reviews}")
            print(f"  - 回头客数量: {returning_customers}")
            print(f"  - 总回头次数: {total_return_visits}")
            print(f"  - 平均情感得分: {avg_sentiment:.4f}")
            print(f"  - 平均评分: {avg_rating:.1f}")
            print(f"  - 文件已保存: {output_path}")
            print("-" * 50)

        except Exception as e:
            print(f"处理文件 {file_name} 时出错: {str(e)}")
            continue

    print("所有文件处理完成！")


if __name__ == "__main__":
    main()