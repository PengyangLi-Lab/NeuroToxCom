import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.patches import Patch
from matplotlib.lines import Line2D
from sklearn.decomposition import PCA
from sklearn.manifold import TSNE
from rdkit import Chem
from rdkit.Chem import AllChem
import warnings

warnings.filterwarnings('ignore')
import os

# 设置matplotlib全局字体为Arial，字号6
plt.rcParams['font.family'] = 'Arial'
plt.rcParams['font.size'] = 6
plt.rcParams['axes.labelsize'] = 6
plt.rcParams['axes.titlesize'] = 6
plt.rcParams['legend.fontsize'] = 6
plt.rcParams['xtick.labelsize'] = 6
plt.rcParams['ytick.labelsize'] = 6
plt.rcParams['axes.linewidth'] = 0.5  # 坐标轴线宽0.5磅

# 抑制RDKit的警告
from rdkit import RDLogger

RDLogger.DisableLog('rdApp.*')

# ==================== 设置文件路径 ====================
file_path = r"D:\结构 PBMT.xlsx"
desktop = os.path.join(os.path.expanduser("~"), "Desktop")

print("=" * 50)
print("第一步：检查文件路径")
print("=" * 50)
print(f"尝试打开文件: {file_path}")

if not os.path.exists(file_path):
    print(f"错误：文件不存在！")
    exit()

# ==================== 读取数据 ====================
print("=" * 50)
print("第二步：读取Excel数据")
print("=" * 50)

try:
    df = pd.read_excel(file_path, sheet_name='Sheet1')
    print(f"成功读取数据！")
    print(f"原始数据行数: {len(df)}")
    print(f"列名: {df.columns.tolist()}")
    print("\n数据预览:")
    print(df.head())
except Exception as e:
    print(f"读取文件失败: {e}")
    exit()

# ==================== 计算摩根指纹 ====================
print("\n" + "=" * 50)
print("第三步：计算摩根指纹 (radius=2, nBits=2048)")
print("=" * 50)


def compute_morgan_fingerprints(smiles_list, radius=2, nbits=2048):
    fingerprints = []
    valid_indices = []

    for i, smiles in enumerate(smiles_list):
        if pd.isna(smiles) or not isinstance(smiles, str):
            continue
        smiles_clean = smiles.split('|')[0].strip()
        mol = Chem.MolFromSmiles(smiles_clean)
        if mol is not None:
            try:
                fp = AllChem.GetMorganFingerprintAsBitVect(mol, radius, nBits=nbits)
                fingerprints.append(np.array(fp))
                valid_indices.append(i)
            except:
                continue

    return np.array(fingerprints), valid_indices


fingerprints, valid_indices = compute_morgan_fingerprints(df['SMILES'].values)
print(f"成功计算指纹的分子数: {len(fingerprints)}")
print(f"指纹矩阵形状: {fingerprints.shape}")

df_filtered = df.iloc[valid_indices].reset_index(drop=True)

# ==================== PCA降维 ====================
print("\n" + "=" * 50)
print("第四步：PCA降维 (2048 -> 50维)")
print("=" * 50)

pca = PCA(n_components=min(50, len(fingerprints)), random_state=42)
fingerprints_pca = pca.fit_transform(fingerprints)
print(f"PCA后数据形状: {fingerprints_pca.shape}")
print(f"PCA累积解释方差比: {pca.explained_variance_ratio_.sum():.3f}")

# ==================== t-SNE降维 ====================
print("\n" + "=" * 50)
print("第五步：t-SNE降维 (50 -> 2维)")
print("=" * 50)

perplexity_value = min(30, len(fingerprints_pca) - 1)
print(f"使用 perplexity = {perplexity_value}")

# 创建t-SNE对象
tsne = TSNE(
    n_components=2,
    perplexity=perplexity_value,
    random_state=42,
    metric='euclidean',
    init='random',
    learning_rate='auto'
)

fingerprints_tsne = tsne.fit_transform(fingerprints_pca)
print(f"t-SNE后数据形状: {fingerprints_tsne.shape}")

# ==================== 准备绘图数据 ====================
print("\n" + "=" * 50)
print("第六步：准备可视化数据")
print("=" * 50)


# 定义危害性质分类 - PBT和PBMT分开
def get_hazard_category(row):
    if row.get('PBMT') == 1:
        return 'PBMT'
    elif row.get('PBT') == 1:
        return 'PBT'
    elif row.get('PMT') == 1:
        return 'PMT'
    else:
        return 'Not PBT&PMT'


df_filtered['Hazard'] = df_filtered.apply(get_hazard_category, axis=1)

# 确保Superclass列有值
df_filtered['Superclass'] = df_filtered['Superclass'].fillna('Others')

# ==================== 添加Superclass缩写映射 ====================
superclass_abbrev = {
    'Benzenoids': 'Benzenoids',
    'Organoheterocyclic compounds': 'Heterocycles',
    'Organic acids and derivatives': 'Org. Acid',
    'Organic oxygen compounds': 'Org. O-Comp',
    'Organohalogen compounds': 'Org. Hal',
    'Lipids and lipid-like molecules': 'Lipid',
    'Organic nitrogen compounds': 'Org. N-Comp',
    'Hydrocarbons': 'HC',
    'Organosulfur compounds': 'Org. S-Comp',
    'Nucleosides, nucleotides, and analogues': 'Nucleoside/Tide',
    'Alkaloids and derivatives': 'Alkaloid',
    'Phenylpropanoids and polyketides': 'Phenylprop./Polyket.',
    'Organophosphorus compounds': 'Org. P-Comp',
    'Acetylides': 'Acetylides',
    'Others': 'Other'
}

# 创建缩写列
df_filtered['Superclass_abbrev'] = df_filtered['Superclass'].map(superclass_abbrev)

print("\n危害性质分布:")
print(df_filtered['Hazard'].value_counts())
print("\nSuperclass分布:")
print(df_filtered['Superclass'].value_counts())
print("\nSuperclass缩写分布:")
print(df_filtered['Superclass_abbrev'].value_counts())

# ==================== 定义颜色映射（使用指定颜色） ====================
color_mapping = {
    'PBT': '#B02042',    # 浅蓝色
    'PBMT': '#C39BE1',   # 浅紫色
    'PMT': '#DCF4FC',    # 非常浅的蓝色
    'Not PBT&PMT': '#808080'  # 灰色
}

# 定义形状映射
shape_mapping = {
    'Benzenoids': 'o',              # 圆形
    'Organoheterocyclic compounds': 's',        # 正方形
    'Organic acids and derivatives': '^',        # 上三角
    'Organic oxygen compounds': 'v',             # 下三角
    'Organohalogen compounds': 'D',              # 菱形
    'Lipids and lipid-like molecules': 'p',      # 五边形
    'Organic nitrogen compounds': 'h',           # 六边形
    'Hydrocarbons': '*',                         # 五角星
    'Organosulfur compounds': 'X',               # X形
    'Nucleosides, nucleotides, and analogues': 'd',  # 小菱形
    'Alkaloids and derivatives': '<',             # 左三角
    'Phenylpropanoids and polyketides': '>',      # 右三角
    'Organophosphorus compounds': 'H',            # 另一种六边形
    'Acetylides': 'P',                            # 加号
    'Others': '8'                                 # 八边形
}

# ==================== 创建第一张图：颜色填充 ====================
print("\n" + "=" * 50)
print("第七步：创建第一张图 - 颜色填充")
print("=" * 50)

# 创建图形 - 16x9 cm
fig1, ax1 = plt.subplots(figsize=(16 / 2.54, 9 / 2.54))

# 设置背景透明
fig1.patch.set_alpha(0)
ax1.patch.set_alpha(0)

# 绘制散点 - 填充版本
for superclass in df_filtered['Superclass'].unique():
    mask_super = df_filtered['Superclass'] == superclass
    shape = shape_mapping.get(superclass, '8')

    for hazard in df_filtered.loc[mask_super, 'Hazard'].unique():
        mask = mask_super & (df_filtered['Hazard'] == hazard)
        if mask.sum() == 0:
            continue

        color = color_mapping.get(hazard, 'gray')

        ax1.scatter(
            fingerprints_tsne[mask, 0],
            fingerprints_tsne[mask, 1],
            c=color,           # 填充颜色
            marker=shape,
            s=10,
            alpha=0.7,
            edgecolors='black',
            linewidth=0.3
        )

# 设置坐标轴范围 -100到100
ax1.set_xlim(-100, 100)
ax1.set_ylim(-100, 100)

# 设置坐标轴标签
ax1.set_xlabel('Dimension 1')
ax1.set_ylabel('Dimension 2')

# 移除标题
ax1.set_title('')

# 添加网格
ax1.grid(True, alpha=0.3, linestyle='--', linewidth=0.5)

# 创建图例
# 1. Hazard图例 - 右下角，无标题
hazard_legend_elements = [
    Patch(facecolor='#B02042', edgecolor='black', linewidth=0.5, label='PBT'),
    Patch(facecolor='#C39BE1', edgecolor='black', linewidth=0.5, label='PBMT'),
    Patch(facecolor='#DCF4FC', edgecolor='black', linewidth=0.5, label='PMT'),
    Patch(facecolor='#808080', edgecolor='black', linewidth=0.5, label='Not PBT&PMT'),
]

# 2. Superclass图例 - 左上角（使用缩写）- 调整markersize与图中点大小一致
superclass_legend_elements = []
for superclass, shape in shape_mapping.items():
    if superclass in df_filtered['Superclass'].unique():
        # 使用缩写作为标签
        abbrev = superclass_abbrev.get(superclass, superclass)

        # 对于填充形状（如圆形、正方形等）
        if shape in ['o', 's', '^', 'v', '<', '>', 'D', 'd', 'p', 'h', 'H', '*', '8']:
            superclass_legend_elements.append(
                Line2D([0], [0],
                       marker=shape,
                       color='black',  # 边框颜色
                       markerfacecolor='white',  # 填充白色
                       markersize=4,   # 调整到与图中点大小一致 (s=10对应markersize约5)
                       label=abbrev,
                       linestyle='None',
                       markeredgewidth=0.5,
                       markeredgecolor='black')  # 边框宽度和颜色
            )
        # 对于线条形状（如 +, X, 1, 2, 3, 4 等）
        else:
            superclass_legend_elements.append(
                Line2D([0], [0],
                       marker=shape,
                       color='black',  # 线条颜色
                       markersize=5,   # 调整到与图中点大小一致
                       label=abbrev,
                       linestyle='None',
                       markeredgewidth=0.5)
            )

# 创建图例 - 都无背景
legend1 = ax1.legend(handles=superclass_legend_elements,
                     loc='upper left',
                     frameon=False,
                     handletextpad=0.5,
                     markerscale=1.0)  # 确保图例标记不缩放

legend2 = ax1.legend(handles=hazard_legend_elements,
                     loc='lower right',
                     frameon=False,
                     title=None,
                     handletextpad=0.5,
                     markerscale=1.0)  # 确保图例标记不缩放

ax1.add_artist(legend1)

# 调整布局
plt.tight_layout()

# 保存第一张图
svg_path1 = os.path.join(desktop, 'tsne_visualization_filled.svg')
plt.savefig(svg_path1, dpi=300, bbox_inches='tight', format='svg', transparent=True)
print(f"填充版SVG已保存到: {svg_path1}")

pdf_path1 = os.path.join(desktop, 'tsne_visualization_filled.pdf')
plt.savefig(pdf_path1, bbox_inches='tight', format='pdf', transparent=True)
print(f"填充版PDF已保存到: {pdf_path1}")

png_path1 = os.path.join(desktop, 'tsne_visualization_filled.png')
plt.savefig(png_path1, dpi=300, bbox_inches='tight', transparent=True)
print(f"填充版PNG已保存到: {png_path1}")

plt.show()

# ==================== 保存结果 ====================
print("\n" + "=" * 50)
print("第八步：保存结果")
print("=" * 50)

results_df = pd.DataFrame({
    'SMILES': df_filtered['SMILES'],
    'Superclass': df_filtered['Superclass'],
    'Superclass_abbrev': df_filtered['Superclass_abbrev'],
    'Hazard': df_filtered['Hazard'],
    'tSNE1': fingerprints_tsne[:, 0],
    'tSNE2': fingerprints_tsne[:, 1]
})

csv_path = os.path.join(desktop, 'tsne_results.csv')
results_df.to_csv(csv_path, index=False, encoding='utf-8-sig')
print(f"结果已保存到: {csv_path}")

# 输出统计信息
print("\n" + "=" * 50)
print("最终统计信息")
print("=" * 50)
print(f"总分子数: {len(df_filtered)}")
print(f"\n危害类别分布:")
print(df_filtered['Hazard'].value_counts())
print(f"\nSuperclass分布:")
print(df_filtered['Superclass'].value_counts())
print(f"\nSuperclass缩写分布:")
print(df_filtered['Superclass_abbrev'].value_counts())

# ... 前面的代码保持不变 ...

print("\n" + "=" * 50)
print("分析完成！")
print("=" * 50)

# ==================== 添加结果分析 ====================
print("\n" + "=" * 50)
print("结果分析报告")
print("=" * 50)

# 1. 总体统计
print("\n1. 总体统计")
print("-" * 30)
total_molecules = len(df_filtered)
print(f"总分子数: {total_molecules}")

# 2. 危害性质分布分析
print("\n2. 危害性质分布分析")
print("-" * 30)
hazard_counts = df_filtered['Hazard'].value_counts()
hazard_percentages = df_filtered['Hazard'].value_counts(normalize=True) * 100

for hazard in ['PBMT', 'PBT', 'PMT', 'Not PBT&PMT']:
    if hazard in hazard_counts.index:
        count = hazard_counts[hazard]
        percentage = hazard_percentages[hazard]
        print(f"{hazard}: {count} 个分子 ({percentage:.1f}%)")
    else:
        print(f"{hazard}: 0 个分子 (0.0%)")

# 计算总危害分子数
hazardous_molecules = hazard_counts.get('PBMT', 0) + hazard_counts.get('PBT', 0) + hazard_counts.get('PMT', 0)
hazardous_percentage = (hazardous_molecules / total_molecules) * 100 if total_molecules > 0 else 0
print(f"\n总计危害分子 (PBT/PBMT/PMT): {hazardous_molecules} 个 ({hazardous_percentage:.1f}%)")

# 3. PBT/PMT/PBMT化合物详细分析
print("\n3. PBT/PMT/PBMT化合物详细分析")
print("-" * 30)

# 创建危害类别的DataFrame
hazard_details = []

for hazard_type in ['PBT', 'PMT', 'PBMT']:
    if hazard_type in hazard_counts.index and hazard_counts[hazard_type] > 0:
        mask = df_filtered['Hazard'] == hazard_type
        count = mask.sum()

        # 这些危害分子的Superclass分布
        superclass_in_hazard = df_filtered.loc[mask, 'Superclass'].value_counts()
        superclass_in_hazard_pct = df_filtered.loc[mask, 'Superclass'].value_counts(normalize=True) * 100

        print(f"\n{hazard_type} 化合物 (共 {count} 个):")
        print(f"  Superclass分布:")

        for i, (sc, sc_count) in enumerate(superclass_in_hazard.head(5).items()):
            sc_pct = superclass_in_hazard_pct[sc]
            abbrev = superclass_abbrev.get(sc, sc)
            print(f"    {i + 1}. {sc} ({abbrev}): {sc_count} 个 ({sc_pct:.1f}%)")

        if len(superclass_in_hazard) > 5:
            print(f"    ... 还有 {len(superclass_in_hazard) - 5} 个其他类别")

        # 计算这些危害分子的t-SNE分布范围
        x_coords = fingerprints_tsne[mask, 0]
        y_coords = fingerprints_tsne[mask, 1]
        x_range = x_coords.max() - x_coords.min()
        y_range = y_coords.max() - y_coords.min()

        print(f"  t-SNE分布范围: x轴范围 {x_range:.1f}, y轴范围 {y_range:.1f}")
        print(f"  中心位置: ({x_coords.mean():.1f}, {y_coords.mean():.1f})")

        # 保存详细信息
        hazard_details.append({
            '危害类型': hazard_type,
            '数量': count,
            '主要Superclass': superclass_in_hazard.index[0] if len(superclass_in_hazard) > 0 else '无',
            '主要Superclass占比': f"{superclass_in_hazard_pct.iloc[0]:.1f}%" if len(superclass_in_hazard) > 0 else '0%',
            't-SNE x中心': f"{x_coords.mean():.1f}",
            't-SNE y中心': f"{y_coords.mean():.1f}",
            '分布范围': f"{np.sqrt(x_range ** 2 + y_range ** 2):.1f}"
        })

# 4. Not PBT&PMT化合物详细分析
print("\n4. Not PBT&PMT化合物详细分析")
print("-" * 30)

if 'Not PBT&PMT' in hazard_counts.index and hazard_counts['Not PBT&PMT'] > 0:
    mask_safe = df_filtered['Hazard'] == 'Not PBT&PMT'
    safe_count = mask_safe.sum()
    safe_percentage = (safe_count / total_molecules) * 100

    print(f"Not PBT&PMT 化合物 (共 {safe_count} 个, 占总数的 {safe_percentage:.1f}%)")

    # 安全化合物的Superclass分布
    safe_superclass = df_filtered.loc[mask_safe, 'Superclass'].value_counts()
    safe_superclass_pct = df_filtered.loc[mask_safe, 'Superclass'].value_counts(normalize=True) * 100

    print("\n  Superclass分布 (Top 10):")
    for i, (sc, sc_count) in enumerate(safe_superclass.head(10).items()):
        sc_pct = safe_superclass_pct[sc]
        abbrev = superclass_abbrev.get(sc, sc)
        print(f"    {i + 1}. {sc} ({abbrev}): {sc_count} 个 ({sc_pct:.1f}%)")

    # 统计安全化合物的主要类别
    top_safe_class = safe_superclass.index[0] if len(safe_superclass) > 0 else '无'
    top_safe_count = safe_superclass.iloc[0] if len(safe_superclass) > 0 else 0
    top_safe_pct = safe_superclass_pct.iloc[0] if len(safe_superclass) > 0 else 0

    print(f"\n  最主要的安全化合物类别: {top_safe_class} ({superclass_abbrev.get(top_safe_class, top_safe_class)})")
    print(f"  共有 {top_safe_count} 个安全分子，占该类总数的 {top_safe_pct:.1f}%")

    # 安全化合物的t-SNE分布
    x_safe = fingerprints_tsne[mask_safe, 0]
    y_safe = fingerprints_tsne[mask_safe, 1]
    print(
        f"\n  t-SNE分布范围: x轴 [{x_safe.min():.1f}, {x_safe.max():.1f}], y轴 [{y_safe.min():.1f}, {y_safe.max():.1f}]")

    # 保存安全化合物信息
    safe_details = {
        '安全化合物总数': safe_count,
        '安全化合物占比': f"{safe_percentage:.1f}%",
        '主要安全类别': top_safe_class,
        '主要安全类别占比': f"{top_safe_pct:.1f}%"
    }
else:
    print("Not PBT&PMT 化合物: 0 个 (0.0%)")
    safe_details = {'安全化合物总数': 0, '安全化合物占比': '0%'}

# 5. Superclass分布分析
print("\n5. Superclass分布分析")
print("-" * 30)
superclass_counts = df_filtered['Superclass'].value_counts()
superclass_percentages = df_filtered['Superclass'].value_counts(normalize=True) * 100

print("Top 10 最常见的Superclass:")
for i, (superclass, count) in enumerate(superclass_counts.head(10).items()):
    percentage = superclass_percentages[superclass]
    abbrev = superclass_abbrev.get(superclass, superclass)
    print(f"{i + 1}. {superclass} ({abbrev}): {count} 个分子 ({percentage:.1f}%)")

# 6. 各类Superclass中的危害分子分布
print("\n6. 各类Superclass中的危害分子分布")
print("-" * 30)

hazard_superclass_summary = []

for superclass in df_filtered['Superclass'].unique():
    mask = df_filtered['Superclass'] == superclass
    total_in_class = mask.sum()

    pbmt_count = df_filtered[mask & (df_filtered['Hazard'] == 'PBMT')].shape[0]
    pbt_count = df_filtered[mask & (df_filtered['Hazard'] == 'PBT')].shape[0]
    pmt_count = df_filtered[mask & (df_filtered['Hazard'] == 'PMT')].shape[0]
    safe_count = df_filtered[mask & (df_filtered['Hazard'] == 'Not PBT&PMT')].shape[0]
    hazardous_in_class = pbmt_count + pbt_count + pmt_count
    hazardous_percentage = (hazardous_in_class / total_in_class) * 100 if total_in_class > 0 else 0
    safe_percentage = (safe_count / total_in_class) * 100 if total_in_class > 0 else 0

    hazard_superclass_summary.append({
        'Superclass': superclass,
        '缩写': superclass_abbrev.get(superclass, superclass),
        '总数': total_in_class,
        'PBMT': pbmt_count,
        'PBT': pbt_count,
        'PMT': pmt_count,
        '安全': safe_count,
        '危害总数': hazardous_in_class,
        '危害百分比': f"{hazardous_percentage:.1f}%",
        '危害百分比数值': hazardous_percentage,
        '安全百分比': f"{safe_percentage:.1f}%"
    })

# 转换为DataFrame并排序
hazard_superclass_df = pd.DataFrame(hazard_superclass_summary)
hazard_superclass_df = hazard_superclass_df.sort_values('危害总数', ascending=False)

print("各类别中危害分子数量排名 (Top 10):")
for i, row in hazard_superclass_df.head(10).iterrows():
    print(
        f"{i + 1}. {row['Superclass']} ({row['缩写']}): 总数 {row['总数']}，危害 {row['危害总数']} ({row['危害百分比']})")
    if row['PBMT'] > 0 or row['PBT'] > 0 or row['PMT'] > 0:
        print(f"   - PBMT: {row['PBMT']}, PBT: {row['PBT']}, PMT: {row['PMT']}, 安全: {row['安全']}")

# 7. 危害分子比例最高的Superclass
print("\n7. 危害分子比例最高的Superclass (总数≥5的类别)")
print("-" * 30)
high_risk_classes = hazard_superclass_df[hazard_superclass_df['总数'] >= 5].copy()
high_risk_classes = high_risk_classes.sort_values('危害百分比数值', ascending=False)

if len(high_risk_classes) > 0:
    for i, row in high_risk_classes.head(10).iterrows():
        print(f"{i + 1}. {row['Superclass']} ({row['缩写']}): {row['危害百分比']} ({row['危害总数']}/{row['总数']})")
        print(f"   - 组成: PBMT:{row['PBMT']}, PBT:{row['PBT']}, PMT:{row['PMT']}, 安全:{row['安全']}")
else:
    print("没有总数≥5的类别")

# 8. 安全分子比例最高的Superclass
print("\n8. 安全分子比例最高的Superclass (总数≥5的类别)")
print("-" * 30)
safe_classes = hazard_superclass_df[hazard_superclass_df['总数'] >= 5].copy()
safe_classes = safe_classes.sort_values('安全', ascending=False)

if len(safe_classes) > 0:
    for i, row in safe_classes.head(10).iterrows():
        safe_pct = (row['安全'] / row['总数']) * 100
        print(f"{i + 1}. {row['Superclass']} ({row['缩写']}): 安全 {row['安全']} 个 ({safe_pct:.1f}%)")
else:
    print("没有总数≥5的类别")

# 9. t-SNE可视化分析
print("\n9. t-SNE可视化分析")
print("-" * 30)

# 计算t-SNE坐标的统计信息
tsne_stats = {
    'x轴范围': f"[{fingerprints_tsne[:, 0].min():.1f}, {fingerprints_tsne[:, 0].max():.1f}]",
    'y轴范围': f"[{fingerprints_tsne[:, 1].min():.1f}, {fingerprints_tsne[:, 1].max():.1f}]",
    'x轴标准差': f"{fingerprints_tsne[:, 0].std():.1f}",
    'y轴标准差': f"{fingerprints_tsne[:, 1].std():.1f}"
}

print("t-SNE降维后的坐标统计:")
for key, value in tsne_stats.items():
    print(f"{key}: {value}")

# 10. 观察聚类情况
print("\n10. 聚类观察")
print("-" * 30)
print("基于t-SNE可视化可以观察到:")

# 检查不同危害类别的分布情况
hazard_clusters = {}
for hazard in ['PBMT', 'PBT', 'PMT', 'Not PBT&PMT']:
    mask = df_filtered['Hazard'] == hazard
    if mask.sum() > 0:
        x_coords = fingerprints_tsne[mask, 0]
        y_coords = fingerprints_tsne[mask, 1]
        x_std, y_std = x_coords.std(), y_coords.std()
        spread = np.sqrt(x_std ** 2 + y_std ** 2)  # 分布范围
        hazard_clusters[hazard] = spread

        print(f"\n{hazard} 类分子:")
        print(f"  - 数量: {mask.sum()}")
        print(f"  - 分布范围: {spread:.1f}")
        print(f"  - 中心位置: ({x_coords.mean():.1f}, {y_coords.mean():.1f})")

# 11. 结论总结
print("\n11. 结论总结")
print("-" * 30)

# 根据数据给出结论
if hazardous_molecules > 0:
    print(
        f"✓ 在{total_molecules}个分子中，发现{hazardous_molecules}个({hazardous_percentage:.1f}%)具有PBT/PBMT/PMT性质的分子。")

    # 找出最主要的危害类别
    hazard_counts_no_normal = hazard_counts.drop('Not PBT&PMT' if 'Not PBT&PMT' in hazard_counts.index else None,
                                                 errors='ignore')
    if len(hazard_counts_no_normal) > 0:
        main_hazard = hazard_counts_no_normal.idxmax()
        main_hazard_count = hazard_counts_no_normal.max()
        print(f"✓ 最主要的危害类型是{main_hazard}，共有{main_hazard_count}个分子。")

        # 显示该危害类型的主要Superclass
        main_hazard_mask = df_filtered['Hazard'] == main_hazard
        main_hazard_superclass = df_filtered.loc[main_hazard_mask, 'Superclass'].value_counts()
        if len(main_hazard_superclass) > 0:
            top_sc = main_hazard_superclass.index[0]
            top_sc_count = main_hazard_superclass.iloc[0]
            top_sc_abbrev = superclass_abbrev.get(top_sc, top_sc)
            print(f"   - 主要分布在 {top_sc} ({top_sc_abbrev})，共{top_sc_count}个")
    else:
        main_hazard = "无"
        print(f"✓ 未发现明显的危害类型。")

    # 找出最主要的危害Superclass
    if len(hazard_superclass_df) > 0:
        main_hazard_class = hazard_superclass_df.iloc[0]
        print(f"✓ 含有危害分子最多的Superclass是 '{main_hazard_class['Superclass']}' ({main_hazard_class['缩写']})，")
        print(f"  共有{main_hazard_class['危害总数']}个危害分子，占该类总数的{main_hazard_class['危害百分比']}。")
        print(
            f"  其中 PBMT:{main_hazard_class['PBMT']}, PBT:{main_hazard_class['PBT']}, PMT:{main_hazard_class['PMT']}")

    # 关于聚类的结论
    if len(hazard_clusters) >= 2:
        cluster_values = list(hazard_clusters.values())
        if max(cluster_values) / min(cluster_values) > 2:
            print(f"✓ t-SNE可视化显示不同危害类别的分子分布范围差异较大，可能存在一定的化学空间分离。")
        else:
            print(f"✓ t-SNE可视化显示不同危害类别的分子分布范围较为接近，在化学空间中有重叠。")

    # 关于特定类别风险的建议
    if len(high_risk_classes) > 0:
        high_risk_top = high_risk_classes.iloc[0]
        if high_risk_top['危害百分比数值'] > 50:
            print(
                f"\n⚠️ 重点关注：'{high_risk_top['Superclass']}' ({high_risk_top['缩写']})类分子的危害比例高达{high_risk_top['危害百分比']}，")
            print(f"   建议对此类化合物进行更详细的风险评估。")
            print(f"   该类别共有{high_risk_top['总数']}个分子，其中:")
            print(f"   - PBMT: {high_risk_top['PBMT']}个")
            print(f"   - PBT: {high_risk_top['PBT']}个")
            print(f"   - PMT: {high_risk_top['PMT']}个")
            print(f"   - 安全: {high_risk_top['安全']}个")

    # 关于安全分子的分析
    if 'safe_details' in locals() and safe_details['安全化合物总数'] > 0:
        print(
            f"\n✓ 安全分子 (Not PBT&PMT) 共有 {safe_details['安全化合物总数']} 个，占总数的 {safe_details['安全化合物占比']}")
        if '主要安全类别' in safe_details:
            print(
                f"  主要安全类别是 '{safe_details['主要安全类别']}'，占该类安全分子的 {safe_details['主要安全类别占比']}")
else:
    print("✓ 在分析的分子中未发现具有PBT/PBMT/PMT性质的分子。")
    print("✓ 所有分子均为 Not PBT&PMT 类别。")

# 12. 保存详细分析结果到CSV
print("\n12. 保存详细分析结果")
print("-" * 30)

# 保存危害类别按Superclass的详细分布
analysis_path = os.path.join(desktop, 'hazard_analysis_by_superclass.csv')
hazard_superclass_df.to_csv(analysis_path, index=False, encoding='utf-8-sig')
print(f"危害类别按Superclass的详细分析已保存到: {analysis_path}")

# 保存危害化合物详细分析
if len(hazard_details) > 0:
    hazard_details_df = pd.DataFrame(hazard_details)
    hazard_details_path = os.path.join(desktop, 'hazard_compounds_details.csv')
    hazard_details_df.to_csv(hazard_details_path, index=False, encoding='utf-8-sig')
    print(f"危害化合物详细分析已保存到: {hazard_details_path}")

# 创建简化的总结报告
main_hazard = "无"
main_hazard_class_name = "无"
high_risk_class_name = "无"
high_risk_percentage = "0.0%"
safe_class_name = "无"
safe_count_total = safe_details.get('安全化合物总数', 0) if 'safe_details' in locals() else 0

if hazardous_molecules > 0:
    if 'hazard_counts_no_normal' in locals() and len(hazard_counts_no_normal) > 0:
        main_hazard = hazard_counts_no_normal.idxmax()
    if len(hazard_superclass_df) > 0:
        main_hazard_class_name = hazard_superclass_df.iloc[0]['Superclass']
    if len(high_risk_classes) > 0:
        high_risk_class_name = high_risk_classes.iloc[0]['Superclass']
        high_risk_percentage = high_risk_classes.iloc[0]['危害百分比']
    if len(safe_classes) > 0:
        safe_class_name = safe_classes.iloc[0]['Superclass']

summary_report = f"""
========================================
          分析总结报告
========================================

分析时间: {pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')}
总分子数: {total_molecules}

【危害性质总体分布】
PBMT: {hazard_counts.get('PBMT', 0)} ({hazard_percentages.get('PBMT', 0):.1f}%)
PBT: {hazard_counts.get('PBT', 0)} ({hazard_percentages.get('PBT', 0):.1f}%)
PMT: {hazard_counts.get('PMT', 0)} ({hazard_percentages.get('PMT', 0):.1f}%)
非PBT/PMT: {hazard_counts.get('Not PBT&PMT', 0)} ({hazard_percentages.get('Not PBT&PMT', 0):.1f}%)

【PBT/PMT/PBMT化合物分析】
共发现 {hazardous_molecules} 个危害分子，占总数的 {hazardous_percentage:.1f}%。
最主要的危害类型: {main_hazard}
主要分布类别: {main_hazard_class_name}

【安全化合物分析】
安全分子总数: {safe_count_total} 个，占总数的 {(safe_count_total / total_molecules * 100):.1f}%
主要安全类别: {safe_class_name}

【高风险类别】
危害比例最高的类别: {high_risk_class_name} ({high_risk_percentage})

【可视化说明】
- 图中不同颜色代表不同的危害类别 (PBT:红色, PBMT:深红, PMT:蓝色, 安全:灰色)
- 不同形状代表不同的Superclass类别
- t-SNE降维展示了分子在化学空间中的分布

========================================
"""

summary_path = os.path.join(desktop, 'analysis_summary.txt')
with open(summary_path, 'w', encoding='utf-8') as f:
    f.write(summary_report)
print(f"分析总结报告已保存到: {summary_path}")

print("\n" + summary_report)
print("\n" + "=" * 50)
print("分析完成！")
print("=" * 50)