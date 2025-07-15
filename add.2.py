import pandas as pd
import re
import os
from snownlp import SnowNLP
from pyecharts import options as opts
from pyecharts.charts import Bar
from pyecharts.charts import Pie

# 创建输出目录
output_dir = r'D:\scrapy\网页\情感分析结果'
os.makedirs(output_dir, exist_ok=True)

# 读取Excel文件
try:
    excel_file = pd.ExcelFile(r'D:\scrapy\网页\懂车帝回答表_情感分析.xlsx')
    df = excel_file.parse('Sheet1')
    print(f"数据加载成功，共{len(df)}条记录")
except Exception as e:
    print(f"文件读取错误: {e}")
    # 为演示创建示例数据
    df = pd.DataFrame({
        '问题内容': [
            '比亚迪这款车性能真不错，很满意！',
            '这个车的续航太差了，太失望了',
            '大家觉得宋PLUS DM - i这款车怎么样？',
            '服务态度很好，值得推荐',
            '充电速度快，外观也好看',
            '车内空间太小了，不太实用',
            '价格有点贵，性价比不高',
            '操作很流畅，驾驶体验好',
            '座椅不舒服，长时间开车很累',
            '动力很强劲，加速快'
        ]
    })
    print("使用示例数据进行演示")

# 数据清洗
def clean_text(text):
    if pd.isna(text):
        return ""
    # 移除特殊字符
    text = re.sub(r'[^\w\s]', '', text)
    return text.strip()

df['问题内容'] = df['问题内容'].apply(clean_text)

# 情感分析
def get_sentiment(text):
    if not text:
        return 0.5  # 空文本设为中性
    try:
        return SnowNLP(text).sentiments
    except Exception:
        return 0.5  # 异常情况设为中性

# 添加情感评分列
df['情感评分'] = df['问题内容'].apply(get_sentiment)

# 分类情感倾向
df['情感倾向'] = df['情感评分'].apply(
    lambda x: '正面' if x > 0.6 else ('负面' if x < 0.4 else '中性')
)

# 统计情感分布
sentiment_counts = df['情感倾向'].value_counts().reset_index()
sentiment_counts.columns = ['情感倾向', '数量']

# 提取情感倾向和数量
sentiment_types = sentiment_counts['情感倾向'].tolist()
sentiment_numbers = sentiment_counts['数量'].tolist()

# 创建柱状图
bar = (
    Bar()
    .add_xaxis(sentiment_types)
    .add_yaxis("情感数量", sentiment_numbers)
    .set_global_opts(
        title_opts=opts.TitleOpts(title="情感分析结果分布"),
        toolbox_opts=opts.ToolboxOpts(is_show=True),
        xaxis_opts=opts.AxisOpts(axislabel_opts={"rotate": 15})
    )
)

# 创建饼图
pie = (
    Pie()
    .add(
        series_name="情感分布",
        data_pair=[list(z) for z in zip(sentiment_types, sentiment_numbers)],
        radius=["30%", "75%"],
    )
    .set_global_opts(
        title_opts=opts.TitleOpts(title="情感分析结果分布"),
        toolbox_opts=opts.ToolboxOpts(is_show=True)
    )
    .set_series_opts(
        label_opts=opts.LabelOpts(formatter="{b}: {d}%")
    )
)

# 渲染图表
bar.render("D:\scrapy\网页\情感分析结果\情感分析结果柱状图.html")

# 渲染图表
pie.render("D:\scrapy\网页\情感分析结果\情感分析结果饼图.html")