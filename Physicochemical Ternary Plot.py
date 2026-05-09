# -*- coding: utf-8 -*-
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
from matplotlib.colors import LogNorm
from scipy.spatial import ConvexHull
from matplotlib.lines import Line2D
from matplotlib.patches import Polygon, Rectangle
import mplcursors

# 设置全局字体为 Arial，字号为 5号
plt.rcParams['font.family'] = 'Arial'
plt.rcParams['font.size'] = 5


# ------------------------- 1. 读取并预处理数据 -------------------------
def load_data(file_path):
    try:
        df = pd.read_excel(file_path, engine='openpyxl')
        print("实际列名:", df.columns.tolist())
        df.columns = df.columns.str.replace(' ', '_').str.replace('-', '_')
        print("标准化后列名:", df.columns.tolist())

        # 在load_data函数中，修改numeric_columns
        numeric_columns = ["logkoa", "Log_Kaw", "TE", "LogD", "BAF", "Half_life_in_water_(hours)",
                           "Half_life_in_soil_(hours)", "Half_life_in_sediment_(hours)"]
        for col in numeric_columns:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce')


        # 删除数值列中的缺失值
        df = df.dropna(subset=["logkoa", "Log_Kaw", "TE"])

        print(f"\n处理后数据信息:")
        print(f"数据行数: {len(df)}")
        print(f"logkoa 范围: {df['logkoa'].min():.2f} - {df['logkoa'].max():.2f}")
        print(f"Log_Kaw 范围: {df['Log_Kaw'].min():.2f} - {df['Log_Kaw'].max():.2f}")
        print(f"TE 范围: {df['TE'].min():.2e} - {df['TE'].max():.2f}")

        # 在数据信息输出部分添加沉积物半衰期
        pbt_pmt_columns = ["LogD", "BAF", "Half_life_in_water_(hours)", "Half_life_in_soil_(hours)",
                           "Half_life_in_sediment_(hours)"]
        for col in pbt_pmt_columns:
            if col in df.columns:
                print(f"{col} 范围: {df[col].min():.2f} - {df[col].max():.2f}")

        return df
    except Exception as e:
        print(f"读取文件失败: {e}")
        return None


# ------------------------- 识别PBT和PMT化合物 -------------------------
def identify_pbt_pmt_compounds(df):
    """识别PBT和PMT化合物"""
    # 创建空的DataFrame，保持相同的列结构
    pbt_compounds = pd.DataFrame()
    pmt_compounds = pd.DataFrame()
    both_compounds = pd.DataFrame()

    # PBT标准: BAF > 5000 且 (水半衰期 > 1440 或 土壤半衰期 > 4320 或 沉积物半衰期 > 4320)
    pbt_thresholds = {
        'BAF': 5000,
        'Half_life_in_water_(hours)': 1440,
        'Half_life_in_soil_(hours)': 4320,
        'Half_life_in_sediment_(hours)': 4320
    }

    # PMT标准: LogD < 3.5 且 (水半衰期 > 1440 或 土壤半衰期 > 4320 或 沉积物半衰期 > 4320)
    pmt_thresholds = {
        'LogD': 2.5,
        'Half_life_in_water_(hours)': 1440,
        'Half_life_in_soil_(hours)': 4320,
        'Half_life_in_sediment_(hours)': 4320
    }

    # 检查PBT化合物
    if all(col in df.columns for col in ['BAF', 'Half_life_in_water_(hours)']):
        # 检查土壤和沉积物半衰期列是否存在
        soil_col = 'Half_life_in_soil_(hours)' if 'Half_life_in_soil_(hours)' in df.columns else None
        sediment_col = 'Half_life_in_sediment_(hours)' if 'Half_life_in_sediment_(hours)' in df.columns else None

        # 构建PBT条件
        pbt_conditions = [
            (df['BAF'] > pbt_thresholds['BAF'])
        ]

        # 添加持久性条件（水、土壤或沉积物）
        persistence_conditions = [
            (df['Half_life_in_water_(hours)'] > pbt_thresholds['Half_life_in_water_(hours)'])
        ]

        if soil_col:
            persistence_conditions.append(df[soil_col] > pbt_thresholds['Half_life_in_soil_(hours)'])

        if sediment_col:
            persistence_conditions.append(df[sediment_col] > pbt_thresholds['Half_life_in_sediment_(hours)'])

        # 合并持久性条件（任一满足即可）
        persistence_mask = persistence_conditions[0]
        for condition in persistence_conditions[1:]:
            persistence_mask = persistence_mask | condition

        pbt_mask = pbt_conditions[0] & persistence_mask
        pbt_compounds = df[pbt_mask].copy()
        print(f"\n找到 {len(pbt_compounds)} 个PBT化合物")
        print("PBT标准: BAF > 5000 且 (水半衰期 > 1440 或 土壤半衰期 > 4320 或 沉积物半衰期 > 4320)")

    # 检查PMT化合物
    if all(col in df.columns for col in ['LogD', 'Half_life_in_water_(hours)']):
        # 检查土壤和沉积物半衰期列是否存在
        soil_col = 'Half_life_in_soil_(hours)' if 'Half_life_in_soil_(hours)' in df.columns else None
        sediment_col = 'Half_life_in_sediment_(hours)' if 'Half_life_in_sediment_(hours)' in df.columns else None

        # 构建PMT条件
        pmt_conditions = [
            (df['LogD'] < pmt_thresholds['LogD'])  # 修改为 < 3.5
        ]

        # 添加持久性条件（水、土壤或沉积物）
        persistence_conditions = [
            (df['Half_life_in_water_(hours)'] > pmt_thresholds['Half_life_in_water_(hours)'])
        ]

        if soil_col:
            persistence_conditions.append(df[soil_col] > pmt_thresholds['Half_life_in_soil_(hours)'])

        if sediment_col:
            persistence_conditions.append(df[sediment_col] > pmt_thresholds['Half_life_in_sediment_(hours)'])

        # 合并持久性条件（任一满足即可）
        persistence_mask = persistence_conditions[0]
        for condition in persistence_conditions[1:]:
            persistence_mask = persistence_mask | condition

        pmt_mask = pmt_conditions[0] & persistence_mask
        pmt_compounds = df[pmt_mask].copy()
        print(f"找到 {len(pmt_compounds)} 个PMT化合物")
        print("PMT标准: LogD < 3.5 且 (水半衰期 > 1440 或 土壤半衰期 > 4320 或 沉积物半衰期 > 4320)")

    # 识别同时属于PBT和PMT的化合物
    if not pbt_compounds.empty and not pmt_compounds.empty:
        both_mask = pbt_compounds.index.isin(pmt_compounds.index)
        both_compounds = pbt_compounds[both_mask].copy()
        print(f"找到 {len(both_compounds)} 个同时属于PBT和PMT的化合物")

    return pbt_compounds, pmt_compounds, both_compounds


def draw_square_bounding_box(ax, compounds, color, label, linewidth=1.5, linestyle='--'):
    """绘制化合物的正方形边界框"""
    if len(compounds) == 0:
        return

    # 计算边界框的坐标
    min_x = compounds['logkoa'].min()
    max_x = compounds['logkoa'].max()
    min_y = compounds['Log_Kaw'].min()
    max_y = compounds['Log_Kaw'].max()

    # 计算中心点
    center_x = (min_x + max_x) / 2
    center_y = (min_y + max_y) / 2

    # 计算正方形边长（取宽度和高度的最大值）
    side_length = max(max_x - min_x, max_y - min_y)

    # 根据点的数量调整框的大小
    num_points = len(compounds)
    if num_points <= 5:
        scale_factor = 1.3
    elif num_points <= 20:
        scale_factor = 1.15
    else:
        scale_factor = 1.05

    side_length *= scale_factor

    # 确保框不会太小
    min_box_size = 1.0
    side_length = max(side_length, min_box_size)

    # 计算新的边界
    new_min_x = center_x - side_length / 2
    new_max_x = center_x + side_length / 2
    new_min_y = center_y - side_length / 2
    new_max_y = center_y + side_length / 2

    # 绘制正方形框
    rect = Rectangle((new_min_x, new_min_y), side_length, side_length,
                     fill=False, edgecolor=color, linewidth=linewidth,
                     linestyle=linestyle, label=label, zorder=4)
    ax.add_patch(rect)

    # 在框的右上角添加标签
    ax.text(new_max_x, new_max_y, label, fontsize=4, color=color, weight='bold',
            ha='left', va='bottom', bbox=dict(boxstyle="round,pad=0.1", facecolor='white',
                                              alpha=0.8, linewidth=0.2), zorder=5)

    print(f"{label}框范围: logkoa({new_min_x:.2f}, {new_max_x:.2f}), Log_Kaw({new_min_y:.2f}, {new_max_y:.2f})")


file_path = r"C:\Users\wu'duo'duo\Desktop\Physicochemical Ternary Plot.xlsx"
df = load_data(file_path)
if df is None:
    exit()

print("\n数据预览：")
print(df.head())

# 识别PBT和PMT化合物
pbt_compounds, pmt_compounds, both_compounds = identify_pbt_pmt_compounds(df)

# ------------------------- 2. 绘制三相图 -------------------------
# 设置画布大小为 8cm x 8cm
fig_width_cm = 8
fig_height_cm = 8

# 创建图形，调整子图位置
fig = plt.figure(figsize=(fig_width_cm / 2.5, fig_height_cm / 2.5))

# 创建主图区域 - 调整位置为颜色条留出空间
ax = fig.add_axes([0.12, 0.25, 0.75, 0.6])  # [left, bottom, width, height]

# 定义分界线的点
A = (-5, -1)    # 水相区域左上点
B = (6, -1)     # 水相和土壤分界点
C = (6, 10)     # 空气和土壤分界点
D = (45, -24)   # 水相区域右下点
E = (-5, -25)   # 水相区域左下点

# 修改颜色映射：使用红蓝配色方案，使用对数归一化
cmap = plt.get_cmap('RdBu_r')

# 根据 TE 数据范围重新设置颜色条范围
vmin = 1e-13  # 10^-13
vmax = 1e4  # 10^4

print(f"\nTE 实际数据范围: {df['TE'].min():.2e} - {df['TE'].max():.2f}")
print(f"颜色条固定范围: {vmin:.2e} - {vmax:.2e}")

# 使用对数归一化
norm = LogNorm(vmin=vmin, vmax=vmax)

# 绘制所有数据点 - 使用圆形标记，调淡颜色
scatter = ax.scatter(
    df["logkoa"],
    df["Log_Kaw"],
    c=df['TE'],
    cmap=cmap,
    norm=norm,
    s=5,  # 减小点的大小
    edgecolor="black",
    linewidth=0.1,  # 更细的边框
    alpha=0.7,  # 调淡颜色透明度
    zorder=3,
    marker='o'  # 使用圆形标记
)

# ------------------------- 标记PBT、PMT和Both化合物 -------------------------
# 标记PBT化合物（红色边框，正方形）
if len(pbt_compounds) > 0:
    ax.scatter(
        pbt_compounds["logkoa"],
        pbt_compounds["Log_Kaw"],
        facecolors='none',
        edgecolors='red',
        s=10,  # 较大的点
        linewidth=0.2,
        marker='s',  # 圆形标记
        zorder=4
    )

# 标记PMT化合物（蓝色边框，圆形）
if len(pmt_compounds) > 0:
    ax.scatter(
        pmt_compounds["logkoa"],
        pmt_compounds["Log_Kaw"],
        facecolors='none',
        edgecolors='blue',
        s=10,  # 较大的点
        linewidth=0.2,
        marker='s',  # 圆形标记
        zorder=3
    )

# 标记Both化合物（紫色边框，圆形）
if len(both_compounds) > 0:
    ax.scatter(
        both_compounds["logkoa"],
        both_compounds["Log_Kaw"],
        facecolors='none',
        edgecolors='green',
        s=10,  # 更大的点
        linewidth=0.2,
        marker='s',  # 圆形标记
        zorder=4
    )

# 绘制四条分界线（全部改为黑色，线宽0.2）
# 水相上边界
ax.plot([A[0], B[0]], [A[1], B[1]], linestyle='--', color='black', linewidth=0.2, zorder=2)
# 空气-土壤分界线
ax.plot([B[0], C[0]], [B[1], C[1]], linestyle='--', color='black', linewidth=0.2, zorder=2)
# 水相-土壤斜边界
ax.plot([B[0], D[0]], [B[1], D[1]], linestyle='--', color='black', linewidth=0.2, zorder=2)
# 水相左边界
ax.plot([A[0], E[0]], [A[1], E[0]], linestyle='--', color='black', linewidth=0.2, zorder=2)


# ------------------------- 3. 添加背景色区分区域 -------------------------
# 水相区域 - 修改为四边形区域
water_vertices = np.array([
    [-5, -1], [6, -1], [45, -24], [-5, -25]
])
ax.fill(water_vertices[:, 0], water_vertices[:, 1], color='#E0F7FA', alpha=0.2, zorder=0)

# 空气相区域保持不变
ax.fill_between([-10, 6], -1, 10, color='lightblue', alpha=0.2, zorder=0)

# 土壤/沉积物相区域 - 调整为剩余的区域
soil_vertices = np.array([
    [6, -1], [6, 10], [45, 10], [45, -24], [6, -1]
])
ax.fill(soil_vertices[:, 0], soil_vertices[:, 1], color='#8B4513', alpha=0.1, zorder=0)

# ------------------------- 5. 添加图表标注和装饰 -------------------------
# 区域标注（使用彩色文字，字号5）
ax.text(0, 7, "Air", ha="center", fontsize=5, color="#87CEEB", weight="bold",
)
ax.text(2, -22, "Water", ha="center", fontsize=5, color="#ADD8E6", weight="bold",
        )
ax.text(35, 7, "Organics", ha="center", fontsize=5, color="#8B4513", weight="bold",
        )

# 坐标轴设置 - 调整标签位置
ax.set_xlim(-5, 45)
ax.set_ylim(-25, 10)
ax.set_xlabel("Log $K_{OA}$", fontsize=5, labelpad=0.321)  # 减少标签与坐标轴的距离
ax.set_ylabel("Log $K_{AW}$", fontsize=5, labelpad=0.321)  # 减少标签与坐标轴的距离
ax.set_xticks([-5,0, 5, 10, 15,20,25,30,35,40,45])
ax.set_yticks([-25, -20, -15, -10, -5, 0, 5, 10])
ax.grid(linestyle="--", alpha=0.2, linewidth=0.2)
ax.tick_params(axis='both', which='major', labelsize=5, width=0.2, pad=2)  # 减少刻度标签与坐标轴的距离

# 设置坐标轴线宽
for spine in ax.spines.values():
    spine.set_linewidth(0.2)

# ------------------------- 6. 修改颜色条为对数坐标 -------------------------
# 重新设计对数坐标刻度 - 从 10^-13 到 10^4
te_ticks = [1e-13, 1e-11, 1e-9, 1e-7, 1e-5, 1e-3, 1e-1, 1e1, 1e3]
te_tick_labels = [
    '$10^{-13}$', '$10^{-11}$', '$10^{-9}$', '$10^{-7}$', '$10^{-5}$',
    '$10^{-3}$', '$10^{-1}$', '$10^{1}$', '$10^{3}$'
]

print(f"颜色条刻度: {te_ticks}")
print(f"颜色条刻度标签: {te_tick_labels}")

# 创建颜色条并调整位置 - 下移颜色条
cbar = plt.colorbar(scatter, ax=ax, shrink=0.6, pad=0.03)  # 增加pad值使颜色条下移
cbar.set_label('TE', fontsize=5, labelpad=2)

# 设置颜色条刻度
cbar.set_ticks(te_ticks)
cbar.set_ticklabels(te_tick_labels)
cbar.ax.tick_params(labelsize=5, width=0.2, pad=0.321)  # 减少颜色条刻度标签的间距
cbar.outline.set_linewidth(0.2)

# ------------------------- 7. 添加PBT和PMT图例 -------------------------

# 添加PBT、PMT和Both图例 - 使用正方形的框作为图例标记
legend_elements = []
if len(pbt_compounds) > 0:
    legend_elements.append(Line2D([0], [0], marker='s', color='w', markerfacecolor='none',
                                  markeredgecolor='red', markersize=3, markeredgewidth=0.35, label='vPvBT'))
if len(pmt_compounds) > 0:
    legend_elements.append(Line2D([0], [0], marker='s', color='w', markerfacecolor='none',
                                  markeredgecolor='blue', markersize=3, markeredgewidth=0.35, label='vPvMT'))

if legend_elements:
    # 在颜色条上方添加图例
    ax.legend(handles=legend_elements, loc='upper left', bbox_to_anchor=(0.98, 1.0),
              fontsize=5, framealpha=0, ncol=1)

# 检查是否有数据点超出颜色条范围
out_of_range = df[(df['TE'] < vmin) | (df['TE'] > vmax)]
if not out_of_range.empty:
    print(f"\n警告: {len(out_of_range)} 个数据点超出颜色条范围:")
    print(out_of_range[['TE']].head())
else:
    print(f"\n所有数据点都在颜色条范围内")

# ------------------------- 8. 标签拖拽功能 -------------------------
label_objects = []


def on_pick(event):
    artist = event.artist
    if isinstance(artist, plt.Text):
        artist.set_animated(True)
        fig.canvas.draw()
        artist._dragging = True


def on_motion(event):
    for text in label_objects:
        if hasattr(text, '_dragging') and text._dragging:
            text.set_position((event.xdata, event.ydata))
            fig.canvas.draw_idle()


def on_release(event):
    for text in label_objects:
        if hasattr(text, '_dragging'):
            text._dragging = False


fig.canvas.mpl_connect('pick_event', on_pick)
fig.canvas.mpl_connect('motion_notify_event', on_motion)
fig.canvas.mpl_connect('button_release_event', on_release)

# 保存为矢量图 - 删除白色背景
plt.savefig('triple_phase_diagram_colored.pdf', dpi=300, bbox_inches='tight',
            pad_inches=0.02, transparent=True)
plt.savefig('triple_phase_diagram_colored.svg', dpi=300, bbox_inches='tight',
            pad_inches=0.02, transparent=True)
plt.show()