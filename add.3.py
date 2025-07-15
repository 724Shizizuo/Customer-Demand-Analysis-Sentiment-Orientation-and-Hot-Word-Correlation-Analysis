import pandas as pd
from pyecharts import options as opts
from pyecharts.charts import Bar, WordCloud
from pyecharts.globals import ThemeType
import jieba
from collections import Counter

# 读取文件
excel_file = pd.ExcelFile('D:\scrapy\网页\懂车帝回答表_情感分析.xlsx')

# 获取所有表名
sheet_names = excel_file.sheet_names
sheet_names

# 获取指定工作表中的数据
df = excel_file.parse('Sheet1')

# 查看数据的基本信息
print('数据基本信息：')
df.info()

# 查看数据集行数和列数
rows, columns = df.shape

if rows < 100 and columns < 20:
    # 短表数据（行数少于100且列数少于20）查看全量数据信息
    print('数据全部内容信息：')
    print(df.to_csv(sep='\t', na_rep='nan'))
else:
    # 长表数据查看数据前几行信息
    print('数据前几行内容信息：')
    print(df.head().to_csv(sep='\t', na_rep='nan'))

# 统计每月不同情感倾向的数量
monthly_sentiment_counts = df.groupby(['时间调整', '情感倾向']).size().unstack(fill_value=0)

# 计算占比
monthly_sentiment_percentage = monthly_sentiment_counts.div(monthly_sentiment_counts.sum(axis=1), axis=0) * 100

# 创建堆叠柱状图
bar = (
    Bar(init_opts=opts.InitOpts(theme=ThemeType.LIGHT, width="1800px", height="600px"))
    .add_xaxis(monthly_sentiment_percentage.index.tolist())
)

# 为每种情感倾向添加数据
for sentiment in monthly_sentiment_percentage.columns:
    bar.add_yaxis(sentiment, monthly_sentiment_percentage[sentiment].tolist(), stack="stack1")

# 设置全局选项
bar.set_global_opts(
    title_opts=opts.TitleOpts(title="每月情感倾向占比"),
    toolbox_opts=opts.ToolboxOpts(is_show=True),
    xaxis_opts=opts.AxisOpts(name="月份"),
    yaxis_opts=opts.AxisOpts(name="占比（%）"),
    legend_opts=opts.LegendOpts(pos_top="5%"),
)

# 筛选出负面问题的问题内容
negative_comments = df[df['情感倾向'] == '负面']['问题内容']

# 分词并统计词频
all_words = []
for comment in negative_comments:
    words = jieba.lcut(comment)
    all_words.extend(words)
word_freq = Counter(all_words)

# 生成词云图
wordcloud = (
    WordCloud(init_opts=opts.InitOpts(theme=ThemeType.LIGHT, width="1200px", height="600px"))
    .add("", list(word_freq.items()), word_size_range=[20, 100])
    .set_global_opts(
        title_opts=opts.TitleOpts(title="负面问题高频词云图"),
        toolbox_opts=opts.ToolboxOpts(is_show=True),
    )
)

# 将图表整合到一个网页中
from pyecharts.charts import Page
page = Page()
page.add(bar)
page.add(wordcloud)
page.render("D:\scrapy\网页\情感分析结果\情感占比-月统计+负面问题高频词云图.html")