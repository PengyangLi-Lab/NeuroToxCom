import pandas as pd
import matplotlib.pyplot as plt
import numpy as np

# 设置全局字体为 Arial，字号为 6，不加粗
plt.rcParams['font.family'] = 'Arial'
plt.rcParams['font.size'] = 5
plt.rcParams['font.weight'] = 'normal'
plt.rcParams['axes.labelweight'] = 'normal'
plt.rcParams['axes.titleweight'] = 'normal'
plt.rcParams['legend.fontsize'] = 5

# 设置所有线条粗细为 0.3
plt.rcParams['lines.linewidth'] = 0.3
plt.rcParams['axes.linewidth'] = 0.3
plt.rcParams['xtick.major.width'] = 0.3
plt.rcParams['ytick.major.width'] = 0.3
plt.rcParams['grid.linewidth'] = 0.3
plt.rcParams['legend.handlelength'] = 1.0
plt.rcParams['legend.handleheight'] = 0.5

file_path = r"C:\Users\wu'duo'duo\Desktop\MASST\MASST - Data Sources Across Countries.xlsx"

# 读取数据
df = pd.read_excel(file_path, sheet_name=0)
df.set_index("Country", inplace=True)

# 只保留有数据的国家
df_plot = df[(df > 0).any(axis=1)]

# 计算每个国家的总数（用于后续添加标签）
df_plot["Total"] = df_plot.sum(axis=1)

# 按总数量排序
df_plot = df_plot.sort_values("Total", ascending=True)

# 保存总数到单独变量，然后从数据框中删除（避免在条形图中作为一列）
totals = df_plot["Total"].copy()
df_plot.drop("Total", axis=1, inplace=True)

# 绘图，画布大小 7cm x 7cm
fig_width = 7 / 2.54
fig_height = 7 / 2.54
fig, ax = plt.subplots(figsize=(fig_width, fig_height))
fig.patch.set_facecolor('none')
ax.set_facecolor('none')

# 设置坐标轴线宽，并去掉上边框和左边框
ax.spines['top'].set_visible(False)      # 去掉上边框
ax.spines['right'].set_visible(False)    # 保留右边框
ax.spines['bottom'].set_linewidth(0.3)   # 保留底边框
ax.spines['left'].set_linewidth(0.3) # 去掉左边框

bottom = np.zeros(len(df_plot))
colors = plt.cm.tab20.colors

# 条形图，添加透明度 alpha=0.7
for i, col in enumerate(df_plot.columns):
    ax.barh(df_plot.index, df_plot[col], left=bottom, label=col,
            color=colors[i % len(colors)], edgecolor='none', linewidth=0.3, alpha=0.7)
    bottom += df_plot[col].values

# 设置对数坐标轴
ax.set_xscale('log')
ax.set_xlabel("Count (log scale)", fontname='Arial', fontsize=6, fontweight='normal')
ax.grid(axis='x', linestyle='--', alpha=0.7, linewidth=0.3)

# 设置刻度标签字体
ax.tick_params(axis='both', labelsize=5, width=0.3, length=2)
# 关闭左边框的刻度标签（因为去掉了左边框，y轴刻度标签也可以去掉）
ax.tick_params(axis='y', left=False, labelleft=True)  # 保留y轴标签，去掉刻度线
for label in ax.get_xticklabels():
    label.set_fontname('Arial')
    label.set_fontweight('normal')
for label in ax.get_yticklabels():
    label.set_fontname('Arial')
    label.set_fontweight('normal')

# 在每个条形图的尽头添加总数标签
for i, (country, total) in enumerate(totals.items()):
    # 获取当前条形图的结束位置（总长度）
    # 由于使用了对数坐标，需要在对数空间中计算标签位置
    # 标签位置放在条形图右侧，即 x = total 的位置
    ax.text(total, i, f' {int(total)}',
            va='center', ha='left',
            fontsize=5, fontname='Arial', fontweight='normal')

# 图例放到右下角（loc='lower right'），放在条形图内部
ax.legend(loc='lower right', prop={'family':'Arial', 'size':5, 'weight':'normal'},
          frameon=False)

plt.tight_layout()

# 保存为矢量图 SVG（透明背景）
output_path = r"C:\Users\wu'duo'duo\Desktop\MASST\figure.svg"
plt.savefig(output_path, format='svg', transparent=True, bbox_inches='tight')
plt.show()

print(f"矢量图已保存至: {output_path}，画布尺寸：7cm × 7cm")