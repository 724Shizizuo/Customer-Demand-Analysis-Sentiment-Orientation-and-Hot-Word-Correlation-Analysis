import pandas as pd
from snownlp import SnowNLP
import re

# 读取Excel文件
try:
    excel_file = pd.ExcelFile('D:\scrapy\网页\懂车帝回答表.xlsx')
    df = excel_file.parse('Sheet1')
    print(f"数据加载成功，共{len(df)}条记录")
except Exception as e:
    print(f"文件读取错误: {e}")
    # 为演示创建示例数据
    df = pd.DataFrame({
        'Unnamed: 0': ['https://example.com/1', 'https://example.com/2'],
        'Unnamed: 1': ['用户A', '用户B'],
        'Unnamed: 2': ['/user/123', '/user/456'],
        'Unnamed: 3': ['宋PLUS DM车主·车龄2年', 'nan'],
        'Unnamed: 4': ['原车轮毂可以换更大的轮胎吗？', '这车续航太差了！'],
        'Unnamed: 5': ['nan', 'https://img.example.com/123.jpg'],
        'Unnamed: 6': ['回答', '2回答'],
        'Unnamed: 7': ['收藏', '收藏'],
        'Unnamed: 8': ['23分钟前', '5小时前'],
        '时间调整': ['2025-05-21 10:30', '2025-05-21 05:45']
    })
    print("使用示例数据进行演示")

# 数据清洗
def clean_text(text):
    if pd.isna(text):
        return ""
    # 移除"提问"前缀
    text = re.sub(r'^提问', '', text)
    # 移除表情符号
    emoji_pattern = re.compile("["
        u"\U0001F600-\U0001F64F"  # emoticons
        u"\U0001F300-\U0001F5FF"  # symbols & pictographs
        u"\U0001F680-\U0001F6FF"  # transport & map symbols
        u"\U0001F1E0-\U0001F1FF"  # flags (iOS)
                           "]+", flags=re.UNICODE)
    return emoji_pattern.sub(r'', text)

df['Unnamed: 4'] = df['Unnamed: 4'].apply(clean_text)

# 提取回答数（处理如"2回答"格式）
def extract_reply_count(text):
    if pd.isna(text):
        return 0
    match = re.search(r'\d+', str(text))
    return int(match.group(0)) if match else 0

df['回答数'] = df['Unnamed: 6'].apply(extract_reply_count)

# 情感分析
def get_sentiment(text):
    if not text.strip():
        return 0.5  # 空文本设为中性
    try:
        return SnowNLP(text).sentiments
    except Exception:
        return 0.5  # 异常情况设为中性

df['情感评分'] = df['Unnamed: 4'].apply(get_sentiment)
df['情感倾向'] = df['情感评分'].apply(lambda x: '正面' if x > 0.6 else ('负面' if x < 0.4 else '中性'))

# 构建最终表格
result_df = pd.DataFrame({
    '网址': df['Unnamed: 0'],
    '网名': df['Unnamed: 1'],
    'ID': df['Unnamed: 2'],
    '车主信息': df['Unnamed: 3'],
    '问题内容': df['Unnamed: 4'],
    '网页图片/附件': df['Unnamed: 5'],
    '回答数': df['回答数'],
    '收藏数': df['Unnamed: 7'],
    '提问时间': df['Unnamed: 8'],
    '时间调整': df['时间调整'],
    '情感评分': df['情感评分'],
    '情感倾向': df['情感倾向']
})

# 保存结果
try:
    result_df.to_excel('D:\scrapy\网页\情感分析结果\懂车帝回答表_情感分析.xlsx', index=False)
    print("结果已保存至 'D:\scrapy\网页\情感分析结果\懂车帝回答表_情感分析.xlsx'")
except Exception as e:
    print(f"保存文件时出错: {e}")
    # 保存到当前目录
    result_df.to_excel('D:\scrapy\网页\情感分析结果\懂车帝回答表_情感分析.xlsx', index=False)
    print("结果已保存至当前目录 'D:\scrapy\网页\情感分析结果\懂车帝回答表_情感分析.xlsx'")

# 显示情感分析结果摘要
sentiment_counts = result_df['情感倾向'].value_counts()
print("\n情感分析结果摘要:")
print(sentiment_counts)

# 显示前5条结果
print("\n前5条分析结果示例:")
print(result_df.head().to_string())