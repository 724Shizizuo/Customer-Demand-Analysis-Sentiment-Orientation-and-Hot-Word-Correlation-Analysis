import pandas as pd
from pyecharts import options as opts
from pyecharts.charts import Bar

try:
    # 读取 Excel 文件
    df = pd.read_excel(r'D:\scrapy\网页\懂车帝回答表.xlsx', sheet_name='Sheet1')
    
    if '时间调整' not in df.columns:
        print("错误：Excel 文件中未找到 '时间调整' 列。")
    else:
        # 转换为日期时间类型并提取月份
        df['时间调整'] = pd.to_datetime(df['时间调整'], errors='coerce')
        valid_dates = df['时间调整'].dropna()
        
        if valid_dates.empty:
            print("错误：没有有效的日期数据可用于分析。")
        else:
            # 提取月份（格式：YYYY-MM）
            df['月份'] = valid_dates.dt.strftime('%Y-%m')
            
            # 统计每月问题数并排序
            data_count = df['月份'].value_counts().sort_index()
            
            # 创建柱状图
            bar = (
                Bar()
                .add_xaxis(data_count.index.tolist())
                .add_yaxis("懂车帝问题数", data_count.values.tolist())
                .set_global_opts(
                    title_opts=opts.TitleOpts(title="懂车帝问题数月分布", subtitle="按月统计的问题数量"),
                    xaxis_opts=opts.AxisOpts(name="月份", axislabel_opts={"rotate": 45}),
                    yaxis_opts=opts.AxisOpts(name="问题数量"),
                    toolbox_opts=opts.ToolboxOpts(is_show=True),
                    legend_opts=opts.LegendOpts(pos_top="5%")
                )
            )
            
            # 渲染图表
            bar.render(r'D:\scrapy\网页\懂车帝-月统计.html')
            print("图表已成功生成：D:\\scrapy\\网页\\懂车帝-月统计.html")
            
except FileNotFoundError:
    print("错误：未找到 Excel 文件，请检查文件路径是否正确。")
except Exception as e:
    print(f"发生未知错误：{e}")    