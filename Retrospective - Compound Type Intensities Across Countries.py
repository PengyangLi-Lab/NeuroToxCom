import pandas as pd
import plotly.express as px

# 国家名称到 ISO-3 代码的映射
country_to_iso3 = {
    'China': 'CHN', 'USA': 'USA', 'United States': 'USA', 'United State': 'USA',
    'Japan': 'JPN', 'Germany': 'DEU', 'United Kingdom': 'GBR', 'France': 'FRA',
    'India': 'IND', 'Brazil': 'BRA', 'Canada': 'CAN', 'Australia': 'AUS',
    'Italy': 'ITA', 'Spain': 'ESP', 'Russia': 'RUS', 'South Korea': 'KOR',
    'Mexico': 'MEX', 'South Africa': 'ZAF', 'Austria': 'AUT', 'Finland': 'FIN',
    'Norway': 'NOR', 'Sweden': 'SWE', 'Thailand': 'THA', 'United Arab Emirates': 'ARE'
}

# 读取数据
input_file = r"C:\Users\wu'duo'duo\Desktop\Compound Type Intensities Across Countries.xlsx"
df = pd.read_excel(input_file)

# 过滤和映射
df = df[df['Sheet'] != '全球汇总']
df['ISO3'] = df['Sheet'].map(country_to_iso3)
df = df.dropna(subset=['ISO3'])

# 转换数值列
categories = [
    "Benzene and substituted derivatives",
    "Carboxylic acids and derivatives",
    "Steroids and steroid derivatives",
    "Organonitrogen compounds", "Indoles and derivatives",
    "Organooxygen compounds", "Quinolines and derivatives",
    "Pyridines and derivatives", "Phenols", "Fatty Acyls",
    "Naphthalenes", "Azoles"
]

for cat in categories:
    df[cat] = pd.to_numeric(df[cat], errors='coerce')

    fig = px.choropleth(
        df, locations='ISO3', color=cat,
        hover_name='Sheet', color_continuous_scale='Viridis',
        title='', labels={cat: ''}
    )

    # 透明背景设置，删除所有文字
    fig.update_layout(
        width=1000, height=600,
        paper_bgcolor='rgba(0,0,0,0)',
        plot_bgcolor='rgba(0,0,0,0)',
        geo=dict(
            bgcolor='rgba(0,0,0,0)',
            landcolor='rgba(200,200,200,0.5)',
            oceancolor='rgba(150,150,150,0.3)',
            showframe=False,
            showcountries=True,
            countrycolor='rgba(100,100,100,0.5)'
        ),
        title_font_color='black',
        font=dict(color='black', size=1),
        showlegend=False,
        coloraxis_showscale=False
    )

    # 隐藏标题和坐标轴文字
    fig.update_layout(
        title_text='',
        annotations=[],
        margin=dict(l=0, r=0, t=0, b=0)
    )

    # 保存为SVG
    safe_name = cat.replace('/', '_').replace(' ', '_').replace('(', '').replace(')', '')[:50]
    svg_path = rf"D:\MASST\地图_{safe_name}.svg"

    try:
        fig.write_image(svg_path, width=1000, height=600, scale=2)
        print(f"✓ SVG已生成: {svg_path}")
    except Exception as e:
        print(f"✗ SVG生成失败: {e}")
        print("  请运行: pip install kaleido")

print("完成！")