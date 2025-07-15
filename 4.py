import os
import pandas as pd
from pyecharts import options as opts
from pyecharts.charts import Bar, Page

# 定义需要筛选的词性列表
pos = ['n', 'ng', 'nrfg', 'nrt', 'ns', 'nt', 'nz']

# 词频统计文件路径（合并后的文件）
file_path = r'D:\scrapy\网页\回答-词频统计-名词-所有月份.xlsx'

try:
    # 读取合并后的文件
    df = pd.read_excel(file_path, sheet_name='Sheet1')
    
    # 创建Page对象来组合所有图表
    page = Page()
    
    # 按月份分组（假设数据中有 '月份' 列）
    for month, group in df.groupby('月份'):
        # 筛选词性并取前15个
        filtered_df = group[group['词性'].isin(pos)].iloc[0:15, :]
        
        # 创建柱状图
        bar = (
            Bar()
            .add_xaxis(filtered_df['分词'].tolist())
            .add_yaxis("词频", filtered_df['词频'].tolist())
            .set_global_opts(
                title_opts=opts.TitleOpts(title=f"{month}关注点"),
                xaxis_opts=opts.AxisOpts(axislabel_opts={"rotate": 45}),
                toolbox_opts=opts.ToolboxOpts(is_show=True)
            )
        )
        
        # 将图表添加到Page对象
        page.add(bar)
        
        print(f"已处理 {month} 的数据")
        
    # 渲染所有图表到一个HTML文件
    output_file = os.path.join(r'D:\scrapy\网页', '所有月份关注点.html')
    page.render(output_file)
    
    print(f"所有月份的 HTML 页面已生成: {output_file}")
        
except Exception as e:
    print(f"处理文件时出错: {e}")


try:
    # 读取合并后的文件
    df = pd.read_excel(file_path, sheet_name='Sheet1')
    
    # 创建一个空的DataFrame用于存储所有月份的前15个高频词
    top_words_df = pd.DataFrame()
    
    # 按月份分组并提取每个月的前15个高频词
    for month, group in df.groupby('月份'):
        filtered_df = group[group['词性'].isin(pos)].sort_values('词频', ascending=False).head(15)
        filtered_df['月份'] = month
        top_words_df = pd.concat([top_words_df, filtered_df])
    
    # 获取所有唯一的分词（作为x轴）
    all_words = top_words_df['分词'].unique().tolist()
    
    # 创建柱状图
    bar = Bar()
    bar.add_xaxis(all_words)
    
    # 为每个月添加一个系列
    for month, group in top_words_df.groupby('月份'):
        # 创建该月的词频字典
        word_freq_dict = {word: freq for word, freq in zip(group['分词'], group['词频'])}
        # 生成该月的词频列表（按all_words的顺序）
        month_freq_list = [word_freq_dict.get(word, 0) for word in all_words]
        # 添加系列
        bar.add_yaxis(str(month), month_freq_list, label_opts=opts.LabelOpts(is_show=False))
    
    # 设置全局选项
    bar.set_global_opts(
        title_opts=opts.TitleOpts(title="各月高频词对比"),
        xaxis_opts=opts.AxisOpts(
            name="分词",
            axislabel_opts={"rotate": 90, "font_size": 9}
        ),
        yaxis_opts=opts.AxisOpts(name="词频"),
        legend_opts=opts.LegendOpts(pos_top="5%", pos_right="5%"),
        toolbox_opts=opts.ToolboxOpts(is_show=True),
        datazoom_opts=[opts.DataZoomOpts()],  # 添加数据缩放功能
    )
    
    # 渲染图表
    output_file = os.path.join(r'D:\scrapy\网页', '所有月份高频词对比.html')
    bar.render(output_file)
    
    print(f"所有月份高频词对比图表已生成: {output_file}")
        
except Exception as e:
    print(f"处理文件时出错: {e}")    