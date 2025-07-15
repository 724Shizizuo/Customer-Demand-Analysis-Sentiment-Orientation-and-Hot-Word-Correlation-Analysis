import thulac
from collections import Counter
from pyecharts import options as opts
from pyecharts.charts import Bar
import pandas as pd
from multiprocessing import Pool

# 初始化THULAC分词器
thu = thulac.thulac(seg_only=False)  # 若只需要分词，可设置seg_only=True提高速度

# 定义分词和统计函数
def process_text(text):
    noun_counter = Counter()
    words_with_pos = thu.cut(text)
    for word, flag in words_with_pos:
        if flag.startswith('n'):  # 名词
            noun_counter[word] += 1
    return noun_counter

def main():
    # 读取Excel文件
    df = pd.read_excel(r'D:\scrapy\网页\懂车帝回答表.xlsx')
    # 提取文本列（假设是第五列，根据实际情况修改）
    text_column = df.iloc[:, 4]

    # 使用多进程处理文本
    with Pool() as p:
        results = p.map(process_text, text_column.astype(str))

    # 合并所有计数器
    final_noun_counter = Counter()
    for counter in results:
        final_noun_counter.update(counter)

    # 提取前15个最常见的词
    top_nouns = final_noun_counter.most_common(15)

    # 提取词和词频
    noun_words = [word for word, freq in top_nouns]
    noun_freqs = [freq for word, freq in top_nouns]

    # 创建柱状图
    bar = (
        Bar()
        .add_xaxis(noun_words)
        .add_yaxis("名词", noun_freqs)
        .set_global_opts(
            title_opts=opts.TitleOpts(title="不同词性的词频统计"),
            toolbox_opts=opts.ToolboxOpts(is_show=True),
            xaxis_opts=opts.AxisOpts(axislabel_opts={"rotate": 15})
        )
    )

    # 渲染图表
    bar.render("D:\scrapy\网页\回答-词频统计-名词（前15个）.html")

if __name__ == "__main__":
    main()