import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
from matplotlib.patches import Patch

# ============================================================
# 读取Excel数据
# ============================================================
file_path = r"C:\Users\wu'duo'duo\Desktop\toxpi.xlsx"

# 读取第一个sheet：ToxPi分数
df_scores = pd.read_excel(file_path, sheet_name='toxpi绘图数据')
print(f"共 {len(df_scores)} 个化学品")

# 读取第二个sheet：来源信息（两列：Name 和 Source）
df_sources = pd.read_excel(file_path, sheet_name='来源')

# 创建来源字典：Name -> Source
source_dict = dict(zip(df_sources.iloc[:, 0], df_sources.iloc[:, 1]))

# 为每个化合物添加来源信息，如果找不到则默认为'Literature search'
df_scores['Source'] = df_scores['Name'].map(source_dict).fillna('Literature search')

# 检查来源分布
print(f"\n来源分布：")
print(df_scores['Source'].value_counts())

# ============================================================
# 准备数据
# ============================================================
scores = df_scores['ToxPi Score'].values
names = df_scores['Name'].values
sources = df_scores['Source'].values

# ============================================================
# 定义颜色映射
# ============================================================
color_map = {
    'Literature search': '#2C3E50',  # #F15BB5
    'Industrial list': '#F15BB5',    # 蓝色
    'Both': '#E9C46A'                # 绿色
}

def get_color(source):
    if pd.isna(source):
        return '#95A5A6'
    if 'Literature' in str(source):
        return color_map['Literature search']
    elif 'Industrial' in str(source):
        return color_map['Industrial list']
    elif 'Both' in str(source):
        return color_map['Both']
    else:
        return '#95A5A6'

# ============================================================
# 按 ToxPi 分数降序排序
# ============================================================
sorted_idx = np.argsort(scores)[::-1]
sorted_scores = scores[sorted_idx]
sorted_names = [names[i] for i in sorted_idx]
sorted_sources = [sources[i] for i in sorted_idx]
ranks = np.arange(1, len(sorted_scores) + 1)

# 为每个点分配颜色
point_colors = [get_color(src) for src in sorted_sources]

# ============================================================
# 设置画布：9 × 5 cm，透明背景
# ============================================================
width_cm = 9
height_cm = 5
fig, ax = plt.subplots(figsize=(width_cm / 2.54, height_cm / 2.54))

# 设置透明背景
fig.patch.set_alpha(0)
ax.patch.set_alpha(0)

# 删除上边和右边的坐标轴
ax.spines['top'].set_visible(False)
ax.spines['right'].set_visible(False)

# 设置左边和下边坐标轴线宽
ax.spines['left'].set_linewidth(0.25)
ax.spines['bottom'].set_linewidth(0.25)

# 设置刻度线宽和字体
ax.tick_params(width=0.25, labelsize=5)
for label in ax.get_xticklabels():
    label.set_fontname('Arial')
for label in ax.get_yticklabels():
    label.set_fontname('Arial')

# 设置网格线
ax.grid(True, axis='y', alpha=0.3, linestyle='--', linewidth=0.25)

# 设置坐标轴范围
n_points = len(sorted_scores)
ax.set_xlim(0, n_points + 10)
ax.set_ylim(0, 1.05)



# ============================================================
# 绘制散点图
# ============================================================
for i in range(len(ranks)):
    ax.scatter(ranks[i], sorted_scores[i],
               s=2,  # 点的大小
               c=point_colors[i],
               alpha=0.8,
               edgecolors='none',
               zorder=2)

# ============================================================
# 添加分位数竖线（前5%、10%、20%）
# ============================================================

# 添加图例
#

# ============================================================
# 保存矢量图
# ============================================================
output_pdf = r"C:\Users\wu'duo'duo\Desktop\ToxPi_rank_all.pdf"
output_svg = r"C:\Users\wu'duo'duo\Desktop\ToxPi_rank_all.svg"

plt.savefig(output_pdf, format='pdf', dpi=300, bbox_inches='tight', transparent=True)
plt.savefig(output_svg, format='svg', bbox_inches='tight', transparent=True)

plt.show()

print(f"\n✅ 矢量图已保存：")
print(f"   PDF: {output_pdf}")
print(f"   SVG: {output_svg}")

# 打印统计信息
print(f"\n📊 化合物来源统计：")
lit_count = sum(1 for s in sources if 'Literature' in str(s))
ind_count = sum(1 for s in sources if 'Industrial' in str(s))
both_count = sum(1 for s in sources if 'Both' in str(s))
other_count = len(sources) - lit_count - ind_count - both_count

print(f"   文献报道 (Literature reported): {lit_count}")
print(f"   预测 (Predicted/Industrial list): {ind_count}")
print(f"   Both: {both_count}")
if other_count > 0:
    print(f"   其他: {other_count}")