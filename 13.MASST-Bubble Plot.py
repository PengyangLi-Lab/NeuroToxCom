import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
from matplotlib.lines import Line2D

# 读取数据
file_path = r"C:\Users\wu'duo'duo\Desktop\MASST\MASST-Bubble Plot.xlsx"
df = pd.read_excel(file_path, sheet_name="Sheet1")

# 定义来源列（缩写版：前三个字母加点）
source_cols_full = ["Human", "Animals", "Microorganisms", "Environmental samples",
                    "Plants", "Foods", "Chemical Standards", "Personal Beauty products", "N/A"]

source_cols_abbr = ["Human", "Ani.", "Mic.", "Env.", "Plants", "Food", "Che.", "PCP", "N/A"]

# 定义9种不同颜色（色盲友好配色）
colors = ["#1f77b4", "#ff7f0e", "#2ca02c", "#d62728", "#9467bd",
          "#8c564b", "#e377c2", "#7f7f7f", "#bcbd22"]

# 定义分类
categories = df["classy1"].unique()

# 设置字体和全局线宽
plt.rcParams['font.family'] = 'Arial'
plt.rcParams['font.size'] = 5
plt.rcParams['lines.linewidth'] = 0.25
plt.rcParams['axes.linewidth'] = 0.25
plt.rcParams['xtick.major.width'] = 0.25
plt.rcParams['ytick.major.width'] = 0.25

# -------------------- 1. 单独生成并保存图例图（包含颜色和大小，大小图例水平放置） --------------------
legend_fig, legend_ax = plt.subplots(figsize=(5, 2.5))  # 调整尺寸：更宽、更矮以容纳水平布局
legend_ax.axis('off')  # 隐藏坐标轴

# 第一部分：颜色图例（9个来源）
color_handles = []
for i, (abbr, color) in enumerate(zip(source_cols_abbr, colors)):
    handle = Line2D([0], [0], marker='o', color='w', markerfacecolor=color,
                    markersize=6, markeredgecolor='black', markeredgewidth=0.5,
                    label=abbr)
    color_handles.append(handle)

# 添加颜色图例（放在上方）
color_legend = legend_ax.legend(handles=color_handles, labels=source_cols_abbr,
                                loc='upper center', frameon=False,
                                ncol=5, fontsize=5, handletextpad=0.3,
                                columnspacing=0.8, title="Sources",
                                title_fontsize=6, bbox_to_anchor=(0.5, 1.0))

# 添加颜色图例到坐标轴
legend_ax.add_artist(color_legend)

# 第二部分：气泡大小图例（水平放置）
# 定义示例数值（根据实际数据范围调整）
example_values = [10, 50, 100]  # 原始数值示例
max_example = 100
scaled_sizes = [2 + (val / max_example) * 33 for val in example_values]

# 调整标记大小使其与气泡图一致
size_handles_corrected = []
for val, scaled_size in zip(example_values, scaled_sizes):
    # scatter的s是面积，markersize是半径，面积 = pi * (markersize/2)^2
    # 所以 markersize = 2 * sqrt(s / pi)
    markersize = 2 * np.sqrt(scaled_size / np.pi)
    handle = Line2D([0], [0], marker='o', color='w', markerfacecolor='#7f7f7f',
                    markersize=markersize, markeredgecolor='black', markeredgewidth=0.5,
                    label=f'{val}')
    size_handles_corrected.append(handle)

# 添加大小图例（水平放置，放在颜色图例下方）
size_legend = legend_ax.legend(handles=size_handles_corrected,
                               loc='lower center', frameon=False,
                               ncol=3, fontsize=5, handletextpad=0.5,
                               columnspacing=1.5, title="Values",
                               title_fontsize=6, bbox_to_anchor=(0.5, 0.0))

# 保存图例为单独的SVG
legend_save_path = r"C:\Users\wu'duo'duo\Desktop\MASST\MASST-Bubble Plot.svg"
legend_fig.savefig(legend_save_path, format='svg', dpi=300,
                   bbox_inches='tight', facecolor='none', edgecolor='none')
plt.close(legend_fig)

print(f"图例已保存（包含颜色和大小，大小图例水平放置）: {legend_save_path}")

# -------------------- 2. 原有绘图逻辑（完全不变） --------------------
# 为每个分类绘制独立气泡图
for cat in categories:
    cat_df = df[df["classy1"] == cat].copy()
    if cat_df.empty:
        continue

    # 化合物名称保持完整
    compounds = cat_df["Name"].tolist()
    n_compounds = len(compounds)

    # 每个化合物高度固定0.4cm，加上下边距各0.3cm，再加标题区域0.3cm
    content_height_cm = n_compounds * 0.4
    total_height_cm = content_height_cm + 0.6 + 0.3  # 上下边距0.6 + 标题区域0.3

    # 宽度自适应：根据最长化合物名称长度和横坐标数量计算
    max_name_len = max([len(name) for name in compounds]) if compounds else 10
    width_cm = max(8, 3 + max_name_len * 0.15)

    # 创建图形，转换英寸（1英寸=2.54cm）
    fig, ax = plt.subplots(figsize=(width_cm / 2.54, total_height_cm / 2.54))

    # 设置透明背景
    fig.patch.set_facecolor('none')
    ax.set_facecolor('none')

    # 准备数据矩阵
    data_matrix = []
    for _, row in cat_df.iterrows():
        data_matrix.append([row[col] for col in source_cols_full])
    data_matrix = np.array(data_matrix)

    # 为每个横坐标位置单独绘制气泡（不同颜色）
    for x_idx in range(len(source_cols_full)):
        # 找出该列非零的行
        y_positions = np.where(data_matrix[:, x_idx] > 0)[0]
        sizes = data_matrix[y_positions, x_idx]

        if len(y_positions) > 0:
            # 气泡大小缩放：最小2，最大35
            if len(sizes) > 0:
                sizes_max = sizes.max()
                scaled_sizes = 2 + (sizes / sizes_max) * 33 if sizes_max > 0 else sizes * 0 + 2
            else:
                scaled_sizes = []
            ax.scatter([x_idx] * len(y_positions), y_positions,
                       s=scaled_sizes, c=colors[x_idx], alpha=0.8,
                       edgecolors='black', linewidth=0.25)

    # 设置坐标轴（横坐标不倾斜 rotation=0）
    ax.set_xticks(range(len(source_cols_abbr)))
    ax.set_xticklabels(source_cols_abbr, rotation=0, ha='center', fontsize=5)
    ax.set_yticks(range(n_compounds))
    ax.set_yticklabels(compounds, fontsize=5)

    # 反转y轴使第一个化合物在顶部
    ax.invert_yaxis()

    # 设置坐标轴范围
    ax.set_xlim(-0.5, len(source_cols_abbr) - 0.5)
    ax.set_ylim(n_compounds - 0.5, -0.5)

    # 设置轴脊线宽
    for spine in ax.spines.values():
        spine.set_linewidth(0.25)

    # 设置刻度线宽
    ax.tick_params(width=0.25, length=2)

    # 添加网格线（线宽0.25）
    ax.grid(True, linestyle=':', alpha=0.5, linewidth=0.25)

    # 设置标题放在图的下方
    ax.set_xlabel(cat, fontsize=6, fontweight='bold', labelpad=5)

    # 紧凑布局，但保留边距
    plt.tight_layout(pad=0.5)

    # 保存为矢量图 SVG 格式（无背景）
    safe_cat = cat.replace("/", "_").replace(" ", "_").replace("and", "&")
    save_path = r"C:\Users\wu'duo'duo\Desktop\MASST\bubble_{}.svg".format(safe_cat)
    plt.savefig(save_path, format='svg', dpi=300, bbox_inches='tight', facecolor='none', edgecolor='none')
    plt.close()

    print("已保存: bubble_{}.svg (化合物数: {}, 内容高度: {:.1f}cm, 总高度: {:.1f}cm, 宽度: {:.1f}cm)".format(
        safe_cat, n_compounds, content_height_cm, total_height_cm, width_cm))

print("\n所有气泡图绘制完成！（SVG矢量图格式）")
print(f"图例已保存至: {legend_save_path}")
print("图例布局：上方是颜色图例（5列），下方是气泡大小图例（3列水平排列）")