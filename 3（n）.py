import pandas as pd
import thulac
from collections import Counter

# 初始化thulac分词器
thu = thulac.thulac()

# 读取Excel文件
df = pd.read_excel(r'D:\scrapy\网页\懂车帝回答表.xlsx')

# 提取文本列（假设是第五列，根据实际情况修改）
text_column = df.iloc[:, 4]  # 如果你的文本在其他列，请修改索引

# 合并所有文本为一个字符串
all_text = ' '.join(text_column.astype(str))

# 分词并词性标注
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
        '分词': word,
        '词频': freq,
        '词性': flag
    })

# 创建DataFrame并输出到Excel
output_df = pd.DataFrame(output_data)
output_df.to_excel(r'D:\scrapy\网页\回答-词频统计-名词.xlsx', index=False)

print("分词结果已成功保存到Excel文件！")