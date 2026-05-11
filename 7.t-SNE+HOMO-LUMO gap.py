import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
from matplotlib import rcParams
from matplotlib.lines import Line2D

# 设置全局字体为Arial，字号6
rcParams['font.family'] = 'Arial'
rcParams['font.size'] = 6
rcParams['svg.fonttype'] = 'none'  # 确保SVG中字体可编辑
rcParams['pdf.fonttype'] = 42
rcParams['ps.fonttype'] = 42
rcParams['axes.linewidth'] = 0.5  # 设置坐标轴线宽为0.5

# 定义颜色 - 使用你指定的颜色代码
COLOR_RED = '#A00F0A'    # 红色
COLOR_BLUE = '#173272'   # 蓝色

# 读取数据
file_path = r"C:\Users\wu'duo'duo\Desktop\7.homo-lumogap data.xlsx"
df = pd.read_excel(file_path)

print("数据列名:", df.columns.tolist())
print(f"原始总分子数: {len(df)}")

# ==================== 数据清洗函数 ====================
def clean_numeric(value):
    """清洗数值数据，将无效值转换为NaN"""
    if pd.isna(value):
        return np.nan
    if isinstance(value, str):
        if '****' in value or value.strip() == '':
            return np.nan
        try:
            return float(value)
        except:
            return np.nan
    return value

def clean_binary(value):
    """清洗二值数据（PMT/PBT）"""
    if pd.isna(value):
        return 0
    if isinstance(value, str):
        if '****' in value or value.strip() == '':
            return 0
        try:
            return int(float(value))
        except:
            return 0
    return 1 if value == 1 else 0

print("\n开始数据清洗...")

# 创建清洗后的DataFrame
df_clean = pd.DataFrame()
df_clean['SMILES'] = df['SMILES']
df_clean['log_kow'] = df['log kow'].apply(clean_numeric)
df_clean['bertz_ct'] = df['BertzCT'].apply(clean_numeric)
df_clean['homo_lumo_gap'] = df['HOMO-LUMO gap (eV)'].apply(clean_numeric)
df_clean['pmt'] = df['PMT'].apply(clean_binary)
df_clean['pbt'] = df['PBT'].apply(clean_binary)

# 删除包含NaN的行
df_clean = df_clean.dropna(subset=['log_kow', 'homo_lumo_gap', 'bertz_ct'])

print(f"清洗后有效分子数: {len(df_clean)}")
print(f"删除的无效数据行数: {len(df) - len(df_clean)}")

# 计算分类
df_clean['is_pmt'] = df_clean['pmt'] == 1
df_clean['is_pbt'] = df_clean['pbt'] == 1

print(f"PMT分子数: {df_clean['is_pmt'].sum()}")
print(f"PBT分子数: {df_clean['is_pbt'].sum()}")

# 如果没有有效数据，退出
if len(df_clean) == 0:
    print("错误：没有有效数据可绘图！")
    exit()

# ==================== 根据要求重新设计坐标轴 ====================
# X轴: logKow 范围 -10 到 20
# Y轴: HOMO-LUMO gap 范围 -2 到 18

x_min, x_max = -10, 20
y_min, y_max = -2, 18

print(f"\n设置的坐标轴范围:")
print(f"log Kow: {x_min} - {x_max}")
print(f"HOMO-LUMO gap: {y_min} - {y_max}")

# ==================== 根据bertz_ct范围0-3960重新定义点的大小范围 ====================
# BertzCT范围: 0-3960
min_b, max_b = 0, 3960
print(f"BertzCT范围: {min_b} - {max_b} (固定范围)")

# 设置点的大小范围：最小10，最大200
min_size, max_size = 10, 200

# 计算点的大小，使用固定范围0-3960
df_clean['marker_size'] = min_size + (max_size - min_size) * (df_clean['bertz_ct'] - min_b) / (max_b - min_b)

# 确保点大小在范围内
df_clean['marker_size'] = df_clean['marker_size'].clip(min_size, max_size)

print(f"点大小范围: {min_size} - {max_size}")

# 创建画布尺寸转换函数
def cm_to_inch(cm):
    return cm / 2.54

fig_width = cm_to_inch(8)
fig_height = cm_to_inch(7)

# 设置公共刻度
x_ticks = np.arange(-10, 21, 5)  # -10, -5, 0, 5, 10, 15, 20
y_ticks = np.arange(-2, 19, 4)   # -2, 2, 6, 10, 14, 18

# 创建自定义图例句柄（确保图例中的圆点大小相同）
legend_size = 30  # 图例中点的大小

# ==================== 图1: PBT分类 ====================
fig1, ax1 = plt.subplots(figsize=(fig_width, fig_height))

# 设置刻度线宽度
ax1.tick_params(width=0.5, length=3)

pbt_data = df_clean[df_clean['is_pbt']]
not_pbt_data = df_clean[~df_clean['is_pbt']]

# 绘制非PBT（蓝色 - #173272）
if len(not_pbt_data) > 0:
    ax1.scatter(
        not_pbt_data['log_kow'],
        not_pbt_data['homo_lumo_gap'],
        s=not_pbt_data['marker_size'],
        c=COLOR_BLUE,
        alpha=0.6,
        edgecolors='white',
        linewidth=0.3
    )

# 绘制PBT（红色 - #A00F0A）
if len(pbt_data) > 0:
    ax1.scatter(
        pbt_data['log_kow'],
        pbt_data['homo_lumo_gap'],
        s=pbt_data['marker_size'],
        c=COLOR_RED,
        alpha=0.6,
        edgecolors='white',
        linewidth=0.3
    )

# 创建自定义图例（确保圆点大小相同）
legend_elements = [
    Line2D([0], [0], marker='o', color='w', markerfacecolor=COLOR_BLUE,
           markersize=np.sqrt(legend_size), label='not PBT', alpha=0.6,
           markeredgecolor='white', markeredgewidth=0.3),
    Line2D([0], [0], marker='o', color='w', markerfacecolor=COLOR_RED,
           markersize=np.sqrt(legend_size), label='PBT', alpha=0.6,
           markeredgecolor='white', markeredgewidth=0.3)
]

# 设置坐标轴标签
ax1.set_xlabel('log Kow', fontsize=6)
ax1.set_ylabel('HOMO-LUMO gap [eV]', fontsize=6)

# 设置坐标轴范围
ax1.set_xlim(x_min, x_max)
ax1.set_ylim(y_min, y_max)

# 设置刻度
ax1.set_xticks(x_ticks)
ax1.set_yticks(y_ticks)

# 添加网格（线宽0.5）
ax1.grid(True, alpha=0.3, linestyle='--', linewidth=0.5)

# 添加自定义图例
ax1.legend(handles=legend_elements, fontsize=6, framealpha=0.9, loc='upper right')

# 设置刻度字体大小
ax1.tick_params(labelsize=6)

# 调整布局
plt.tight_layout()

# 保存为SVG矢量图（无背景）
plt.savefig('PBT_classification.svg', dpi=300, bbox_inches='tight',
            transparent=True, format='svg')
plt.savefig('PBT_classification.png', dpi=300, bbox_inches='tight',
            transparent=False)

print("\n图1已保存: PBT_classification.svg (无背景矢量图)")

# ==================== 图2: PMT分类 ====================
fig2, ax2 = plt.subplots(figsize=(fig_width, fig_height))

# 设置刻度线宽度
ax2.tick_params(width=0.5, length=3)

pmt_data = df_clean[df_clean['is_pmt']]
not_pmt_data = df_clean[~df_clean['is_pmt']]

# 绘制非PMT（蓝色 - #173272）
if len(not_pmt_data) > 0:
    ax2.scatter(
        not_pmt_data['log_kow'],
        not_pmt_data['homo_lumo_gap'],
        s=not_pmt_data['marker_size'],
        c=COLOR_BLUE,
        alpha=0.6,
        edgecolors='white',
        linewidth=0.3
    )

# 绘制PMT（红色 - #A00F0A）
if len(pmt_data) > 0:
    ax2.scatter(
        pmt_data['log_kow'],
        pmt_data['homo_lumo_gap'],
        s=pmt_data['marker_size'],
        c=COLOR_RED,
        alpha=0.6,
        edgecolors='white',
        linewidth=0.3
    )

# 创建自定义图例（确保圆点大小相同）
legend_elements_pmt = [
    Line2D([0], [0], marker='o', color='w', markerfacecolor=COLOR_BLUE,
           markersize=np.sqrt(legend_size), label='not PMT', alpha=0.6,
           markeredgecolor='white', markeredgewidth=0.3),
    Line2D([0], [0], marker='o', color='w', markerfacecolor=COLOR_RED,
           markersize=np.sqrt(legend_size), label='PMT', alpha=0.6,
           markeredgecolor='white', markeredgewidth=0.3)
]

# 设置坐标轴标签
ax2.set_xlabel('log Kow', fontsize=6)
ax2.set_ylabel('HOMO-LUMO gap [eV]', fontsize=6)

# 设置坐标轴范围
ax2.set_xlim(x_min, x_max)
ax2.set_ylim(y_min, y_max)

# 设置刻度
ax2.set_xticks(x_ticks)
ax2.set_yticks(y_ticks)

# 添加网格（线宽0.5）
ax2.grid(True, alpha=0.3, linestyle='--', linewidth=0.5)

# 添加自定义图例
ax2.legend(handles=legend_elements_pmt, fontsize=6, framealpha=0.9, loc='upper right')

# 设置刻度字体大小
ax2.tick_params(labelsize=6)

# 调整布局
plt.tight_layout()

# 保存为SVG矢量图（无背景）
plt.savefig('PMT_classification.svg', dpi=300, bbox_inches='tight',
            transparent=True, format='svg')
plt.savefig('PMT_classification.png', dpi=300, bbox_inches='tight',
            transparent=False)

print("图2已保存: PMT_classification.svg (无背景矢量图)")

# ==================== 统计信息 ====================
print("\n" + "="*50)
print("统计信息")
print("="*50)
print(f"总有效分子数: {len(df_clean)}")
print(f"PBT分子数: {len(pbt_data)} ({len(pbt_data)/len(df_clean)*100:.1f}%)")
print(f"PMT分子数: {len(pmt_data)} ({len(pmt_data)/len(df_clean)*100:.1f}%)")
print(f"\nBertzCT统计:")
print(f"  实际最小值: {df_clean['bertz_ct'].min():.2f}")
print(f"  实际最大值: {df_clean['bertz_ct'].max():.2f}")
print(f"  平均值: {df_clean['bertz_ct'].mean():.2f}")
print(f"  中位数: {df_clean['bertz_ct'].median():.2f}")
print(f"\nlog Kow统计:")
print(f"  实际最小值: {df_clean['log_kow'].min():.2f}")
print(f"  实际最大值: {df_clean['log_kow'].max():.2f}")
print(f"\nHOMO-LUMO gap统计:")
print(f"  实际最小值: {df_clean['homo_lumo_gap'].min():.4f}")
print(f"  实际最大值: {df_clean['homo_lumo_gap'].max():.4f}")

# 显示图形（可选）
plt.show()