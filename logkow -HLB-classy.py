import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
from matplotlib.patches import FancyBboxPatch
import matplotlib.colors as mcolors
from matplotlib.colors import LinearSegmentedColormap
from scipy import stats

# 设置字体为Arial，5号字体
plt.rcParams['font.family'] = 'Arial'
plt.rcParams['font.size'] = 5  # 5号字体
plt.rcParams['axes.linewidth'] = 0.3  # 细边框 - 从0.5改为0.3
plt.rcParams['svg.fonttype'] = 'none'  # 确保字体在SVG中可编辑

# 文件路径
file_path = r"C:\Users\wu'duo'duo\Desktop\LOGKOW画图.xlsx"

print("=" * 60)
print("数据详细分析报告")
print("=" * 60)

# 读取Excel文件
df = pd.read_excel(file_path)

# 转换log kow和HLB列为数值类型
df['log kow_numeric'] = pd.to_numeric(df['log kow'], errors='coerce')
df['HLB_numeric'] = pd.to_numeric(df['HLB'], errors='coerce')

# 清理后的数据
df_clean = df.dropna(subset=['log kow_numeric', 'HLB_numeric']).copy()

# 计算LOGHLB
df_clean['LOGHLB_numeric'] = np.log10(df_clean['HLB_numeric'])

print("1. 数据基本信息:")
print(f"   总行数: {df.shape[0]}")
print(f"   有效log kow数据行数: {df_clean.shape[0]}")
print(f"   无效log kow值数量: {df['log kow_numeric'].isna().sum()}")
print(f"   无效HLB值数量: {df['HLB_numeric'].isna().sum()}")

print("\n2. classy1 分类情况:")
classy1_counts = df_clean['classy1'].value_counts()
print(f"   唯一类别数: {len(classy1_counts)}")
print(f"   非空classy1数量: {df_clean['classy1'].notna().sum()}")

# 按classy1分组的统计
classy1_logkow_stats = df_clean.groupby('classy1')['log kow_numeric'].agg([
    'count', 'min', 'max', 'mean', 'median', 'std'
]).round(3)

classy1_hlb_stats = df_clean.groupby('classy1')['HLB_numeric'].agg([
    'count', 'min', 'max', 'mean', 'median', 'std'
]).round(3)

# 计算所有化合物的总体范围
all_compounds_logkow = {
    'min': df_clean['log kow_numeric'].min(),
    'max': df_clean['log kow_numeric'].max(),
    'count': len(df_clean),
    'mean': df_clean['log kow_numeric'].mean(),
    'median': df_clean['log kow_numeric'].median()
}

all_compounds_hlb = {
    'min': df_clean['HLB_numeric'].min(),
    'max': df_clean['HLB_numeric'].max(),
    'count': len(df_clean),
    'mean': df_clean['HLB_numeric'].mean(),
    'median': df_clean['HLB_numeric'].median()
}

print("\n3. HLB 统计信息:")
print(f"   范围: {df_clean['HLB_numeric'].min():.2f} 到 {df_clean['HLB_numeric'].max():.2f}")
print(f"   平均值: {df_clean['HLB_numeric'].mean():.2f}")
print(f"   中位数: {df_clean['HLB_numeric'].median():.2f}")

# 筛选至少有5个数据点的类别
valid_classes_logkow = classy1_logkow_stats[classy1_logkow_stats['count'] >= 5].sort_values('count', ascending=False)
valid_classes_hlb = classy1_hlb_stats[classy1_hlb_stats['count'] >= 5].sort_values('count', ascending=False)

print(f"\n4. 可用于绘图的类别:")
print(f"   logKow数据点≥5的类别数量: {len(valid_classes_logkow)}")
print(f"   HLB数据点≥5的类别数量: {len(valid_classes_hlb)}")

print("\n" + "=" * 60)
print("开始绘制矢量图...")
print("=" * 60)

# 创建18cm宽画布 (18cm ≈ 7.09英寸)
fig_width_cm = 24
fig_width_inch = fig_width_cm / 2.54
fig_height_inch = fig_width_inch * 0.48  # 增加高度以容纳All compound

fig = plt.figure(figsize=(fig_width_inch, fig_height_inch))

# 可调节参数：y轴标签距离图形的距离
y_label_offset = 0.321 # 控制y轴标签距离，数值越大距离越远，范围建议0-0.05

# 创建两个子图，调整位置让它们更靠近中间
left_ax = fig.add_axes([0.045, 0.15, 0.273, 0.65])  # [left, bottom, width, height]
right_ax = fig.add_axes([0.765 - y_label_offset, 0.15, 0.273, 0.65])  # 根据y_label_offset调整右图位置

# 设置横坐标范围
logkow_min, logkow_max = -10, 20  # 根据数据范围-7.75到18.08设置
hlb_min, hlb_max = 0.01, 10000  # HLB范围设置为0.01到10000

# 获取共同的类别（两个数据集中都有的类别）
common_classes = list(set(valid_classes_logkow.index) & set(valid_classes_hlb.index))

# 按log Kow范围区间跨度（最大值-最小值）排序（从大到小）
common_classes_sorted = sorted(common_classes,
                               key=lambda x: (valid_classes_logkow.loc[x, 'max'] - valid_classes_logkow.loc[x, 'min'])
                               if x in valid_classes_logkow.index else 0,
                               reverse=True)

# 如果共同类别太多，只显示前20个以避免过于拥挤
if len(common_classes_sorted) > 20:
    common_classes_sorted = common_classes_sorted[:20]
    print(f"注意: 只显示前20个类别以避免图表过于拥挤")

# 在类别列表最前添加"All compounds"
all_classes = ['All compounds'] + common_classes_sorted

# 图表1: logKow水平条形图 (左侧)
# 减小y坐标间距，让柱子更紧凑
y_spacing = 0.73  # 减小间距，让柱子更靠近
y_pos = np.arange(len(all_classes)) * y_spacing
bar_height = 0.55  # 保持柱子高度不变

# ========== 渐变参数 ==========
corner_radius = 0.15  # 控制圆角程度
gradient_segments = 50  # 渐变分段数，数值越大渐变越平滑
marker_size = 15  # 标记点大小

# 标记颜色定义
mean_color = '#FF6B6B'    # 平均值标记颜色 - 红色五角星
median_color = '#F7DC6F'  # 中位数标记颜色 - 棕色三角形


# =============================

def create_smooth_gradient(ax, y, width, height, left, data_values, base_color, alpha=1.0):
    """创建平滑渐变的圆角条形"""
    if len(data_values) <= 1:
        # 如果数据点太少，使用单一颜色
        box = FancyBboxPatch(
            (left, y - height / 2),
            width, height,
            boxstyle=f"round,pad=0,rounding_size={corner_radius}",
            facecolor=base_color,
            edgecolor='black',
            linewidth=0.0,
            alpha=alpha
        )
        ax.add_patch(box)
        return

    try:
        # 计算核密度估计
        kde = stats.gaussian_kde(data_values)

        # 创建渐变位置
        x_positions = np.linspace(left, left + width, gradient_segments)
        segment_width = width / gradient_segments

        # 计算每个位置的密度
        densities = kde(x_positions)

        # 归一化密度到 [0.2, 1.0] 范围
        if np.max(densities) > np.min(densities):
            normalized_densities = 0.2 + 0.8 * (densities - np.min(densities)) / (np.max(densities) - np.min(densities))
        else:
            normalized_densities = np.ones_like(densities) * 0.7

        # 创建自定义渐变色图
        base_rgb = mcolors.to_rgb(base_color)
        cmap = LinearSegmentedColormap.from_list(
            'density_cmap',
            [mcolors.to_rgba(base_color, alpha=0.3),  # 最浅
             mcolors.to_rgba(base_color, alpha=1.0)],  # 最深
            N=256
        )

        # 绘制每个渐变段
        for i, (x_pos, density_factor) in enumerate(zip(x_positions, normalized_densities)):
            segment_color = cmap(density_factor)
            segment_color = mcolors.to_rgba(segment_color, alpha=alpha)

            # 创建渐变段
            segment_box = FancyBboxPatch(
                (x_pos, y - height / 2),
                segment_width, height,
                boxstyle="square,pad=0",  # 内部段不使用圆角
                facecolor=segment_color,
                edgecolor='none',
                linewidth=0.0
            )
            ax.add_patch(segment_box)

        # 添加外边框（圆角矩形）
        border_box = FancyBboxPatch(
            (left, y - height / 2),
            width, height,
            boxstyle=f"round,pad=0,rounding_size={corner_radius}",
            facecolor='none',
            edgecolor='black',
            linewidth=0.3,
            alpha=0.8
        )
        ax.add_patch(border_box)

    except:
        # 如果计算失败，使用单一颜色
        box = FancyBboxPatch(
            (left, y - height / 2),
            width, height,
            boxstyle=f"round,pad=0,rounding_size={corner_radius}",
            facecolor=base_color,
            edgecolor='black',
            linewidth=0.0,
            alpha=alpha
        )
        ax.add_patch(box)


def create_center_gradient(ax, y, width, height, left, data_values, base_color, alpha=1.0):
    """创建中心向两端渐变的条形"""
    if len(data_values) <= 2:
        # 数据点太少，使用单一颜色
        box = FancyBboxPatch(
            (left, y - height / 2),
            width, height,
            boxstyle=f"round,pad=0,rounding_size={corner_radius}",
            facecolor=base_color,
            edgecolor='black',
            linewidth=0.0,
            alpha=alpha
        )
        ax.add_patch(box)
        return

    # 计算数据的统计信息
    mean_val = np.mean(data_values)
    std_val = np.std(data_values)

    # 创建从中心向两端渐变的颜色映射
    base_rgb = mcolors.to_rgb(base_color)

    # 创建渐变位置（从中心向两端）
    center_pos = (mean_val - left) / width  # 中心位置在条形中的相对位置

    # 创建渐变段
    for i in range(gradient_segments):
        x_segment = left + (i / gradient_segments) * width
        segment_width = width / gradient_segments

        # 计算当前段到中心的距离（归一化）
        segment_center = x_segment + segment_width / 2
        distance_to_center = abs(segment_center - (left + center_pos * width)) / (width / 2)
        distance_to_center = min(distance_to_center, 1.0)  # 限制在[0,1]范围内

        # 根据距离计算颜色强度：中心最深，两端最浅
        color_intensity = 1.0 - 0.7 * distance_to_center  # 0.3到1.0的范围
        segment_color = mcolors.to_rgba(base_color, alpha=color_intensity * alpha)

        # 创建渐变段
        segment_box = FancyBboxPatch(
            (x_segment, y - height / 2),
            segment_width, height,
            boxstyle="square,pad=0",
            facecolor=segment_color,
            edgecolor='none',
            linewidth=0.0
        )
        ax.add_patch(segment_box)

    # 添加外边框
    border_box = FancyBboxPatch(
        (left, y - height / 2),
        width, height,
        boxstyle=f"round,pad=0,rounding_size={corner_radius}",
        facecolor='none',
        edgecolor='black',
        linewidth=0.0,
        alpha=0.8
    )
    ax.add_patch(border_box)


# ========== 添加分区直线（统一红色）和标签 ==========
partition_lines = [-2, 5, 8]  # 分区直线位置

# 在logKow图上添加分区直线（统一红色虚线）- 放在最底层
for line_pos in partition_lines:
    left_ax.axvline(x=line_pos, color='red', linestyle='--',
                    linewidth=0.5, alpha=0.3, zorder=1)

# 设置坐标轴刻度和标签颜色
ticks = [-10, -5, 0, 5, 10, 15, 20]
tick_labels = ['-10', '-5', '0', '5', '10', '15', '20']
left_ax.set_xticks(ticks)
left_ax.set_xticklabels(tick_labels)

# 将5的刻度标签设置为红色
for i, label in enumerate(left_ax.get_xticklabels()):
    if ticks[i] == 5:
        label.set_color('red')
        label.set_weight('bold')
    else:
        label.set_color('black')

# 添加特殊的分区刻度标签（只添加-2和8）
special_ticks = [-2, 8]
for tick in special_ticks:
    left_ax.text(tick, -0.045, f'{tick}', transform=left_ax.get_xaxis_transform(),
                 ha='center', va='top', fontsize=5, color='red',
                 fontfamily='Arial', weight='bold')
# 绘制logKow条形图 - 使用中心渐变
for i, class_name in enumerate(all_classes):
    if class_name == 'All compounds':
        # 绘制All compounds
        bar_width = all_compounds_logkow['max'] - all_compounds_logkow['min']
        all_data = df_clean['log kow_numeric'].values
        create_center_gradient(left_ax, y_pos[i], bar_width, bar_height,
                               all_compounds_logkow['min'], all_data, '#C0C0C0', alpha=0.3)

        # 添加范围标注
        range_text = f"{all_compounds_logkow['min']:.1f}-{all_compounds_logkow['max']:.1f}"
        left_ax.text(all_compounds_logkow['min'] - 0.5, y_pos[i], range_text,
                     va='center', fontsize=5, fontfamily='Arial', ha='right', weight='normal')

        # 添加平均值和中位数标记
        mean_val = all_compounds_logkow['mean']
        median_val = all_compounds_logkow['median']

        # 平均值标记（五角星）- 红色
        left_ax.plot(mean_val, y_pos[i], '*', color=mean_color, markersize=marker_size / 4,
                     markeredgewidth=0.5, markeredgecolor=mean_color, zorder=10)

        # 中位数标记（三角形）- 棕色
        left_ax.plot(median_val, y_pos[i], '^', color=median_color, markersize=marker_size / 5,
                     markeredgewidth=0.5, markeredgecolor=median_color, zorder=10)

    elif class_name in valid_classes_logkow.index:
        row = valid_classes_logkow.loc[class_name]
        bar_width = row['max'] - row['min']
        # 获取该类别的所有数据点
        class_data = df_clean[df_clean['classy1'] == class_name]['log kow_numeric'].values

        # 计算该类别的中位数（如果不存在）
        if 'median' not in row:
            median_val = np.median(class_data)
        else:
            median_val = row['median']

        create_center_gradient(left_ax, y_pos[i], bar_width, bar_height,
                               row['min'], class_data, '#75B0A0', alpha=0.8)

        # 添加范围标注
        range_text = f"{row['min']:.1f}-{row['max']:.1f}"
        left_ax.text(row['min'] - 0.5, y_pos[i], range_text,
                     va='center', fontsize=5, fontfamily='Arial', ha='right')

        # 添加平均值和中位数标记
        mean_val = row['mean']

        # 平均值标记（五角星）- 红色
        left_ax.plot(mean_val, y_pos[i], '*', color=mean_color, markersize=marker_size / 4,
                     markeredgewidth=0.5, markeredgecolor=mean_color, zorder=10)

        # 中位数标记（三角形）- 棕色
        left_ax.plot(median_val, y_pos[i], '^', color=median_color, markersize=marker_size / 5,
                     markeredgewidth=0.5, markeredgecolor=median_color, zorder=10)

left_ax.set_yticks(y_pos)
left_ax.set_yticklabels([])  # 隐藏左图的y轴标签
left_ax.set_xlabel('log Kow', fontfamily='Arial', fontsize=5)
left_ax.set_xlim(logkow_min, logkow_max)
left_ax.grid(axis='x', linestyle='--', alpha=0.3, linewidth=0.2)
left_ax.tick_params(axis='x', labelsize=5, width=0.3)
left_ax.tick_params(axis='y', length=0)

# 设置y轴范围
left_ax.set_ylim(y_pos[0] - 0.5, y_pos[-1] + 0.5)

# 图表2: HLB水平条形图 (右侧) - 使用对数坐标
# 图表2: HLB水平条形图 (右侧) - 使用对数坐标
# 设置HLB为对数坐标
right_ax.set_xscale('log')

# 设置对数坐标的刻度 - 使用科学计数法e表示（上标格式）
hlb_ticks = [0.01, 0.1, 1, 10, 100, 1000, 10000, 100000, 1000000, 10000000]
hlb_tick_labels = [
    '$10^{-2}$',  # 0.01
    '$10^{-1}$',  # 0.1
    '$10^{0}$',   # 1
    '$10^{1}$',   # 10
    '$10^{2}$',   # 100
    '$10^{3}$',   # 1000
    '$10^{4}$',   # 10000
    '$10^{5}$',   # 100000
    '$10^{6}$',   # 1000000
    '$10^{7}$'    # 10000000
]
right_ax.set_xticks(hlb_ticks)
right_ax.set_xticklabels(hlb_tick_labels)
# 相应地调整HLB坐标轴范围
hlb_min, hlb_max = 0.01, 10000000  # HLB范围设置为0.01到10000000

for i, class_name in enumerate(all_classes):
    if class_name == 'All compounds':
        # 绘制All compounds
        bar_width = all_compounds_hlb['max'] - all_compounds_hlb['min']
        all_data = df_clean['HLB_numeric'].values
        create_center_gradient(right_ax, y_pos[i], bar_width, bar_height,
                               all_compounds_hlb['min'], all_data, '#C0C0C0', alpha=0.3)


        # 添加范围标注 - 严格保留两位有效数字
        def format_number_smart(value):
            """智能格式化数字，严格保留两位有效数字"""
            if value == 0:
                return "0"
            elif abs(value) < 100:
                # 100以下的数
                if abs(value) < 1:
                    return f"{value:.2f}".rstrip('0').rstrip('.')  # 0.01-0.99
                elif abs(value) < 10:
                    return f"{value:.1f}".rstrip('0').rstrip('.')  # 1.0-9.9
                else:
                    return f"{int(value)}"  # 10-99
            else:
                # 100以上的数，手动计算科学计数法
                exponent = int(np.floor(np.log10(abs(value))))
                coefficient = round(value / (10 ** exponent), 1)
                if coefficient == 10.0:
                    coefficient = 1.0
                    exponent += 1
                return f"{coefficient:.1f}e{exponent}"


        # 对于All compounds
        min_val = all_compounds_hlb['min']
        max_val = all_compounds_hlb['max']
        range_text = f"{format_number_smart(min_val)}-{format_number_smart(max_val)}"
        right_ax.text(all_compounds_hlb['max'] * 1.7, y_pos[i], range_text,
                      va='center', fontsize=5, fontfamily='Arial', weight='normal')

        # 添加平均值和中位数标记
        mean_val = all_compounds_hlb['mean']
        median_val = all_compounds_hlb['median']

        # 平均值标记（五角星）- 红色
        right_ax.plot(mean_val, y_pos[i], '*', color=mean_color, markersize=marker_size / 4,
                      markeredgewidth=0.5, markeredgecolor=mean_color, zorder=10)

        # 中位数标记（三角形）- 棕色
        right_ax.plot(median_val, y_pos[i], '^', color=median_color, markersize=marker_size / 5,
                      markeredgewidth=0.5, markeredgecolor=median_color, zorder=10)

    elif class_name in valid_classes_hlb.index:
        row = valid_classes_hlb.loc[class_name]
        bar_width = row['max'] - row['min']
        # 获取该类别的所有数据点
        class_data = df_clean[df_clean['classy1'] == class_name]['HLB_numeric'].values

        # 计算该类别的中位数（如果不存在）
        if 'median' not in row:
            median_val = np.median(class_data)
        else:
            median_val = row['median']

        create_center_gradient(right_ax, y_pos[i], bar_width, bar_height,
                               row['min'], class_data, '#A2C5DB', alpha=1.0)



        # 对于其他类别
        min_val = row['min']
        max_val = row['max']
        range_text = f"{format_number_smart(min_val)}-{format_number_smart(max_val)}"
        right_ax.text(row['max'] * 1.7, y_pos[i], range_text,
                      va='center', fontsize=5, fontfamily='Arial')
        # 添加平均值和中位数标记
        mean_val = row['mean']

        # 平均值标记（五角星）- 红色
        right_ax.plot(mean_val, y_pos[i], '*', color=mean_color, markersize=marker_size / 4,
                      markeredgewidth=0.5, markeredgecolor=mean_color, zorder=10)

        # 中位数标记（三角形）- 棕色
        right_ax.plot(median_val, y_pos[i], '^', color=median_color, markersize=marker_size / 5,
                      markeredgewidth=0.5, markeredgecolor=median_color, zorder=10)

right_ax.set_yticks(y_pos)
right_ax.set_yticklabels(all_classes, fontfamily='Arial', fontsize=5, ha='center')
# 设置x轴标签为HLB_B（B为下标）
right_ax.set_xlabel('HL$_B$/h', fontfamily='Arial', fontsize=5)
right_ax.set_xlim(hlb_min, hlb_max)
right_ax.grid(axis='x', linestyle='--', alpha=0.3, linewidth=0.2)
right_ax.tick_params(axis='x', labelsize=5, width=0.3)
right_ax.tick_params(axis='y', length=0, pad=10 + y_label_offset * 100)

# 设置y轴范围
right_ax.set_ylim(y_pos[0] - 0.5, y_pos[-1] + 0.5)

# 设置边框
for ax in [left_ax, right_ax]:
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['left'].set_visible(False)
    ax.spines['bottom'].set_visible(True)
    ax.spines['bottom'].set_linewidth(0.2)

# 添加中间的水平线
for i in y_pos:
    left_ax.axhline(i, color='gray', linewidth=0.08, alpha=0.2, zorder=0)
    right_ax.axhline(i, color='gray', linewidth=0.08, alpha=0.2, zorder=0)

# 在All compounds行下方添加分隔线
separator_pos = y_pos[0] + (y_pos[1] - y_pos[0]) / 2
left_ax.axhline(separator_pos, color='red', linewidth=0.3, alpha=0.7, linestyle='--')
right_ax.axhline(separator_pos, color='red', linewidth=0.3, alpha=0.7, linestyle='--')

# 添加图例说明 - 一行显示，放在All compounds正下方
legend_elements = [
    plt.Line2D([0], [0], marker='*', color=mean_color, markerfacecolor=mean_color,
               markersize=marker_size / 3, label='Mean', linestyle='None', markeredgewidth=0),
    plt.Line2D([0], [0], marker='^', color=median_color, markerfacecolor=median_color,
               markersize=marker_size / 4, label='Median', linestyle='None', markeredgewidth=0),
]

left_ax.legend(handles=legend_elements, loc='upper left',
               bbox_to_anchor=(1.066, -0),  # 放在All compounds正下方
               fontsize=5, framealpha=0.8, handlelength=0.5,
               frameon=False, ncol=2)  # ncol=2 表示两列并排显示为一行
# 保存为多种矢量图格式（透明背景）
output_files = [
    'compound_ranges_gradient.svg',
    'compound_ranges_gradient.pdf',
    'compound_ranges_gradient.eps'
]

for output_file in output_files:
    plt.savefig(output_file, dpi=300, bbox_inches='tight',
                facecolor='none', edgecolor='none',  # 将facecolor改为'none'
                transparent=True)  # 添加transparent=True
    print(f"已保存矢量图: {output_file}")

# 同时保存PNG格式作为参考（透明背景）
plt.savefig('compound_ranges_gradient.png', dpi=300, bbox_inches='tight',
            facecolor='none', edgecolor='none',
            transparent=True)
print(f"已保存位图参考: compound_ranges_gradient.png")

plt.show()

print("\n" + "=" * 60)
print("图表说明:")
print("=" * 60)
print("1. 分区直线:")
print(f"   • x = -2, 5, 8 (红色虚线，位于图层底部)")
print("\n2. 统计标记:")
print("   • ★ 五角星: 平均值 (Mean) - 红色")
print("   • ▲ 三角形: 中位数 (Median) - 棕色")
print("\n3. 标注格式:")
print("   • 仅显示范围: 最小值-最大值")
print("\n4. 坐标轴设置:")
print("   • 左侧: log Kow 线性坐标")
print("   • 右侧: HLB 对数坐标 (0.01 - 10000)")
print("\n5. 渐变效果:")
print("   • 中心区域颜色较深: 数据点集中")
print("   • 两端区域颜色较浅: 数据点稀疏")
print(f"   • 使用平滑渐变 ({gradient_segments}个渐变段)")

print(f"\n排序方式: 按log Kow范围区间跨度从大到小排列")
print(f"显示的类别数量: {len(all_classes)}")