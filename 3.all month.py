import pandas as pd
import thulac
from collections import Counter
from concurrent.futures import ProcessPoolExecutor

# 初始化thulac分词器
thu = thulac.thulac()

def process_month(month, group, text_column_name):
    # 合并该月所有文本为一个字符串
    all_text = ' '.join(group[text_column_name].astype(str))
    # 使用thulac进行分词和词性标注
    words_with_pos = thu.cut(all_text)
    # 筛选出名词并统计词频
    nouns = []
    for word, flag in words_with_pos:
        if flag.startswith('n'):
            nouns.append((word, flag))
    # 词频统计
    word_freq = Counter(nouns)
    # 准备输出数据
    output_data = []
    for (word, flag), freq in word_freq.items():
        output_data.append({
            '月份': month,
            '分词': word,
            '词频': freq,
            '词性': flag
        })
    return output_data

def main():
    # 读取Excel文件
    df = pd.read_excel(r'D:\scrapy\网页\懂车帝回答表.xlsx')
    # 假设日期列是第一列，根据实际情况修改
    date_column = df.iloc[:, 9]
    # 将日期列转换为日期时间类型
    date_column = pd.to_datetime(date_column)
    # 提取月份信息
    month_column = date_column.dt.to_period('M')
    # 提取文本列（假设是第五列，根据实际情况修改）
    text_column = df.iloc[:, 4]
    # 按月份分组
    grouped = df.groupby(month_column)
    # 创建一个空的列表用于存储所有月份的结果
    all_output_data = []
    text_column_name = text_column.name
    # 使用多进程并行处理不同月份的数据
    with ProcessPoolExecutor() as executor:
        results = []
        for month, group in grouped:
            result = executor.submit(process_month, month, group, text_column_name)
            results.append(result)
        for future in results:
            all_output_data.extend(future.result())
    # 创建DataFrame并输出到Excel
    output_df = pd.DataFrame(all_output_data)
    output_file = r'D:\scrapy\网页\回答-词频统计-名词-所有月份.xlsx'
    output_df.to_excel(output_file, index=False)
    print("所有月份的分词结果已成功保存到Excel文件！")

if __name__ == "__main__":
    main()