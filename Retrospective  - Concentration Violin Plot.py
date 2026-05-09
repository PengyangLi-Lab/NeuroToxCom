import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.lines import Line2D
import warnings

from scipy.stats import kruskal
import scikit_posthocs as sp

warnings.filterwarnings('ignore')

# ============================================================
# 全局绘图参数
# ============================================================

# 设置所有字体为Arial，字号6
plt.rcParams['font.family'] = 'Arial'
plt.rcParams['font.size'] = 6
plt.rcParams['axes.unicode_minus'] = False

# 设置线条宽度
plt.rcParams['lines.linewidth'] = 0.25
plt.rcParams['axes.linewidth'] = 0.25
plt.rcParams['xtick.major.width'] = 0.25
plt.rcParams['ytick.major.width'] = 0.25
plt.rcParams['patch.linewidth'] = 0.25

# 设置画布大小 6×7 cm
fig_width = 6 / 2.54
fig_height = 7 / 2.54

# 缩写映射
abbrev_map = {
    'Naphthalenes': 'Naphthalenes',
    'Benzene and substituted derivatives': 'Benz',
    'Carboxylic acids and derivatives': 'Carbox.',
    'Steroids and steroid derivatives': 'Ster.',
    'Organonitrogen compounds': 'N-organ.',
    'Indoles and derivatives': 'Indol.',
    'Organooxygen compounds': 'O-organ.',
    'Quinolines and derivatives': 'Quin.',
    'Pyridines and derivatives': 'Pyrid.',
    'Phenols': 'Phenols',
    'Fatty Acyls': 'Fat.',
    'Others': 'Others'
}


# ============================================================
# 1. 数据读取与预处理
# ============================================================

def load_and_preprocess_data(file_path):
    """
    读取数据，数量小于20的类别归为Others，并进行log10转换
    """

    # 如果是xlsx
    if file_path.endswith('.xlsx') or file_path.endswith('.xls'):
        df = pd.read_excel(file_path)
    # 如果是csv
    elif file_path.endswith('.csv'):
        df = pd.read_csv(file_path)
    else:
        raise ValueError("文件格式应为 .xlsx, .xls 或 .csv")

    print(f"原始数据形状: {df.shape}")

    # 检查必要列
    required_cols = ['Classification', 'intensity']
    for col in required_cols:
        if col not in df.columns:
            raise ValueError(f"缺少必要列: {col}")

    # 删除Classification或intensity为空的化合物
    df_clean = df.dropna(subset=['Classification', 'intensity']).copy()
    print(f"删除缺失值后: {len(df_clean)} 个化合物")

    # intensity转为数值
    df_clean['intensity'] = pd.to_numeric(df_clean['intensity'], errors='coerce')
    df_clean = df_clean.dropna(subset=['intensity']).copy()

    # 删除强度<=0的数据，因为log10不能处理非正数
    df_clean = df_clean[df_clean['intensity'] > 0].copy()
    print(f"删除非正强度后: {len(df_clean)} 个化合物")

    # 统计各类别数量
    class_counts = df_clean['Classification'].value_counts()

    # 找出数量>=20的类别
    major_classes = class_counts[class_counts >= 20].index.tolist()
    print(f"数量≥20的类别数: {len(major_classes)}")

    # 将数量<20的类别归为Others
    df_clean['Classification_Grouped'] = df_clean['Classification'].apply(
        lambda x: x if x in major_classes else 'Others'
    )

    # 对数转换
    df_clean['log_intensity'] = np.log10(df_clean['intensity'])

    # 添加缩写列
    df_clean['Class_Short'] = df_clean['Classification_Grouped'].apply(
        lambda x: abbrev_map.get(x, x[:10])
    )

    return df_clean, major_classes


# ============================================================
# 2. 统计分析：类别数量、平均值、中位数
# ============================================================

def generate_statistics(df_final, major_classes):
    """
    生成各类别的统计分析报告，按log10中位数排序
    """

    stats_data = []

    for cls in major_classes + ['Others']:
        cls_data = df_final[df_final['Classification_Grouped'] == cls]

        if len(cls_data) > 0:
            stats_data.append({
                '类别': cls,
                '缩写': abbrev_map.get(cls, cls[:10]),
                '化合物数量': len(cls_data),
                '原始强度_最小值': cls_data['intensity'].min(),
                '原始强度_最大值': cls_data['intensity'].max(),
                '原始强度_平均值': cls_data['intensity'].mean(),
                '原始强度_中位数': cls_data['intensity'].median(),
                '原始强度_标准差': cls_data['intensity'].std(),
                'log10强度_平均值': cls_data['log_intensity'].mean(),
                'log10强度_中位数': cls_data['log_intensity'].median(),
                'log10强度_标准差': cls_data['log_intensity'].std()
            })

    stats_df = pd.DataFrame(stats_data)

    # 按log10强度_中位数由高到低排序，Others始终放最后
    stats_df_no_others = stats_df[stats_df['类别'] != 'Others'].copy()
    stats_df_others = stats_df[stats_df['类别'] == 'Others'].copy()

    stats_df_no_others = stats_df_no_others.sort_values(
        'log10强度_中位数',
        ascending=False
    )

    stats_df = pd.concat(
        [stats_df_no_others, stats_df_others],
        ignore_index=True
    )

    return stats_df


def print_statistics_summary(stats_df):
    """
    打印统计摘要
    """

    print("\n" + "=" * 80)
    print("统计分析结果：按中位数由高到低排序")
    print("=" * 80)

    total_compounds = stats_df['化合物数量'].sum()

    print(f"总化合物数: {total_compounds}")
    print(f"类别数，含Others: {len(stats_df)}")

    print("\n各类别详细信息:")
    print("-" * 85)

    for _, row in stats_df.iterrows():
        print(
            f"{row['缩写']:12} : "
            f"n={int(row['化合物数量']):3}, "
            f"均值={row['原始强度_平均值']:.2e}, "
            f"中位数={row['原始强度_中位数']:.2e}, "
            f"log10中位数={row['log10强度_中位数']:.2f}"
        )


# ============================================================
# 3. 显著性分析：Kruskal-Wallis + Dunn's test
# ============================================================

def significance_stars(p):
    """
    根据p值返回显著性星号
    """

    if pd.isna(p):
        return 'NA'
    elif p < 0.001:
        return '***'
    elif p < 0.01:
        return '**'
    elif p < 0.05:
        return '*'
    else:
        return 'ns'


def format_p_value(p):
    """
    格式化p值
    """

    if pd.isna(p):
        return 'NA'
    elif p < 0.001:
        return 'p < 0.001'
    else:
        return f'p = {p:.3f}'


def perform_significance_analysis(df_final, stats_df):
    """
    对不同化合物分类的log10强度进行显著性分析

    方法：
    1. Kruskal-Wallis整体检验
    2. Dunn's test两两比较
    3. Benjamini-Hochberg/FDR校正
    """

    print("\n" + "=" * 80)
    print("显著性分析：Kruskal-Wallis + Dunn's test")
    print("=" * 80)

    # 按绘图顺序获取类别
    ordered_classes = stats_df['类别'].tolist()

    group_data = []
    valid_classes = []

    for cls in ordered_classes:
        values = df_final.loc[
            df_final['Classification_Grouped'] == cls,
            'log_intensity'
        ].dropna().values

        if len(values) > 0:
            group_data.append(values)
            valid_classes.append(cls)

    # ----------------------------
    # 3.1 Kruskal-Wallis整体检验
    # ----------------------------
    if len(group_data) >= 2:
        h_stat, kw_p = kruskal(*group_data)
    else:
        h_stat, kw_p = np.nan, np.nan

    kw_result = pd.DataFrame({
        'Test': ['Kruskal-Wallis'],
        'Variable': ['log10(intensity)'],
        'Group': ['Classification_Grouped'],
        'H_statistic': [h_stat],
        'p_value': [kw_p],
        'significance': [significance_stars(kw_p)]
    })

    print("\n整体差异检验 Kruskal-Wallis:")
    if not pd.isna(kw_p):
        print(f"H statistic = {h_stat:.4f}")
        print(f"p value     = {kw_p:.4e}")
        print(f"significance= {significance_stars(kw_p)}")
    else:
        print("有效分类数量不足，无法进行Kruskal-Wallis检验。")

    # ----------------------------
    # 3.2 Dunn's test两两比较
    # ----------------------------
    if len(valid_classes) >= 2:
        dunn_matrix = sp.posthoc_dunn(
            df_final,
            val_col='log_intensity',
            group_col='Classification_Grouped',
            p_adjust='fdr_bh'
        )

        # 按绘图顺序重新排列行列
        dunn_matrix = dunn_matrix.loc[valid_classes, valid_classes]

        pairwise_results = []

        for i in range(len(valid_classes)):
            for j in range(i + 1, len(valid_classes)):
                cls1 = valid_classes[i]
                cls2 = valid_classes[j]

                p_adj = dunn_matrix.loc[cls1, cls2]

                short1 = stats_df.loc[
                    stats_df['类别'] == cls1,
                    '缩写'
                ].values[0]

                short2 = stats_df.loc[
                    stats_df['类别'] == cls2,
                    '缩写'
                ].values[0]

                median1 = stats_df.loc[
                    stats_df['类别'] == cls1,
                    'log10强度_中位数'
                ].values[0]

                median2 = stats_df.loc[
                    stats_df['类别'] == cls2,
                    'log10强度_中位数'
                ].values[0]

                n1 = stats_df.loc[
                    stats_df['类别'] == cls1,
                    '化合物数量'
                ].values[0]

                n2 = stats_df.loc[
                    stats_df['类别'] == cls2,
                    '化合物数量'
                ].values[0]

                pairwise_results.append({
                    'Class_1': cls1,
                    'Class_1_short': short1,
                    'n_1': n1,
                    'Class_2': cls2,
                    'Class_2_short': short2,
                    'n_2': n2,
                    'log10_median_1': median1,
                    'log10_median_2': median2,
                    'median_difference_1_minus_2': median1 - median2,
                    'p_adjusted_BH': p_adj,
                    'significance': significance_stars(p_adj)
                })

        pairwise_df = pd.DataFrame(pairwise_results)

        pairwise_df = pairwise_df.sort_values(
            'p_adjusted_BH',
            ascending=True
        )

        significant_pairs = pairwise_df[
            pairwise_df['p_adjusted_BH'] < 0.05
        ].copy()

        print("\nDunn's test 两两比较完成，校正方法：BH/FDR")
        print(f"全部两两比较数量: {len(pairwise_df)}")
        print(f"显著差异数量 p.adj < 0.05: {len(significant_pairs)}")

        if len(significant_pairs) > 0:
            print("\n显��差异前10组:")
            print("-" * 80)

            for _, row in significant_pairs.head(10).iterrows():
                print(
                    f"{row['Class_1_short']} vs {row['Class_2_short']}: "
                    f"p.adj={row['p_adjusted_BH']:.3e}, "
                    f"{row['significance']}, "
                    f"median diff={row['median_difference_1_minus_2']:.2f}"
                )
        else:
            print("\n未发现校正后 p.adj < 0.05 的两两差异。")

    else:
        dunn_matrix = pd.DataFrame()
        pairwise_df = pd.DataFrame()
        significant_pairs = pd.DataFrame()

    return kw_result, dunn_matrix, pairwise_df, significant_pairs


# ============================================================
# 4. 绘制小提琴图
# ============================================================

def plot_violin_chart(
        df_final,
        stats_df,
        save_path_pdf,
        save_path_svg,
        kw_result=None
):
    """
    绘制小提琴图，按中位数由高到低排序
    """

    # 获取按中位数排序的类别顺序
    all_classes = stats_df['类别'].tolist()
    sorted_short = stats_df['缩写'].tolist()

    # 准备数据
    violin_data = [
        df_final[df_final['Classification_Grouped'] == cls]['log_intensity'].values
        for cls in all_classes
    ]

    # 获取均值和中位数
    medians = stats_df['log10强度_中位数'].tolist()

    # 创建图表
    fig, ax = plt.subplots(figsize=(fig_width, fig_height))
    fig.patch.set_facecolor('none')
    ax.patch.set_facecolor('none')

    # 绘制小提琴图
    parts = ax.violinplot(
        violin_data,
        positions=range(len(all_classes)),
        showmeans=True,
        showmedians=True,
        showextrema=True,
        widths=0.7
    )

    # 设置颜色
    colors = plt.cm.Spectral(np.linspace(0, 1, len(all_classes)))

    for i, pc in enumerate(parts['bodies']):
        pc.set_facecolor(colors[i])
        pc.set_edgecolor('black')
        pc.set_alpha(0.7)
        pc.set_linewidth(0.25)

    # 设置统计线
    parts['cmeans'].set_color('darkred')
    parts['cmeans'].set_linewidth(0.25)

    parts['cmedians'].set_color('darkblue')
    parts['cmedians'].set_linewidth(0.25)

    parts['cmaxes'].set_color('black')
    parts['cmaxes'].set_linewidth(0.25)

    parts['cmins'].set_color('black')
    parts['cmins'].set_linewidth(0.25)

    parts['cbars'].set_color('black')
    parts['cbars'].set_linewidth(0.25)

    # 设置x轴
    ax.set_xticks(range(len(all_classes)))
    ax.set_xticklabels(
        sorted_short,
        rotation=45,
        ha='right',
        fontsize=6
    )

    # 设置y轴
    y_min = np.floor(min([min(data) for data in violin_data if len(data) > 0]))
    y_max = np.ceil(max([max(data) for data in violin_data if len(data) > 0]))

    y_ticks = np.arange(y_min, y_max + 1, 1)

    ax.set_yticks(y_ticks)
    ax.set_yticklabels(
        [f'$10^{{{int(tick)}}}$' for tick in y_ticks],
        fontsize=6
    )

    ax.set_ylim(y_min - 0.5, y_max + 0.5)

    # 设置标签
    ax.set_xlabel('Compound Class', fontsize=6)
    ax.set_ylabel('Intensity', fontsize=6)

    # 添加整体显著性检验结果
    if kw_result is not None and len(kw_result) > 0:
        kw_p = kw_result.loc[0, 'p_value']
        kw_sig = kw_result.loc[0, 'significance']

        if not pd.isna(kw_p):
            if kw_p < 0.001:
                p_text = f"Kruskal-Wallis\np < 0.001 {kw_sig}"
            else:
                p_text = f"Kruskal-Wallis\np = {kw_p:.3f} {kw_sig}"

            ax.text(
                0.98,
                0.98,
                p_text,
                transform=ax.transAxes,
                ha='right',
                va='top',
                fontsize=5,
                bbox=dict(
                    boxstyle='round,pad=0.25',
                    facecolor='white',
                    edgecolor='black',
                    linewidth=0.25,
                    alpha=0.7
                )
            )

    # 添加网格
    ax.grid(
        axis='y',
        alpha=0.3,
        linestyle='--',
        linewidth=0.25
    )
    ax.set_axisbelow(True)

    # 设置边框
    for spine in ax.spines.values():
        spine.set_linewidth(0.25)

    # 添加图例
    legend_elements = [
        Line2D(
            [0],
            [0],
            color='darkred',
            linewidth=0.25,
            marker='_',
            markersize=5,
            label='Mean'
        ),
        Line2D(
            [0],
            [0],
            color='darkblue',
            linewidth=0.25,
            marker='_',
            markersize=5,
            label='Median'
        )
    ]

    ax.legend(
        handles=legend_elements,
        loc='lower right',
        fontsize=5,
        framealpha=0.8
    )

    plt.tight_layout()

    # 保存矢量图，透明背景
    plt.savefig(
        save_path_pdf,
        bbox_inches='tight',
        format='pdf',
        transparent=True
    )

    plt.savefig(
        save_path_svg,
        bbox_inches='tight',
        format='svg',
        transparent=True
    )

    plt.close()

    print(f"\n✅ 矢量图已保存:")
    print(f"   PDF: {save_path_pdf}")
    print(f"   SVG: {save_path_svg}")

    print(f"\n绘图顺序，按中位数由高到低:")
    for i, (cls, short, median) in enumerate(
            zip(all_classes, sorted_short, medians)
    ):
        print(f"   {i + 1}. {short}: log10中位数={median:.2f}")


# ============================================================
# 5. 主函数
# ============================================================

if __name__ == "__main__":

    # 文件路径
    file_path = r"C:\Users\wu'duo'duo\Desktop\Concentration Violin Plot.xlsx"

    try:
        # ----------------------------------------------------
        # 步骤1：数据读取与预处理
        # ----------------------------------------------------
        print("=" * 80)
        print("步骤1: 数据读取与预处理")
        print("=" * 80)

        df_final, major_classes = load_and_preprocess_data(file_path)

        # ----------------------------------------------------
        # 步骤2：统计分析
        # ----------------------------------------------------
        print("\n" + "=" * 80)
        print("步骤2: 统计分析")
        print("=" * 80)

        stats_df = generate_statistics(df_final, major_classes)
        print_statistics_summary(stats_df)

        # ----------------------------------------------------
        # 步骤3：显著性分析
        # ----------------------------------------------------
        print("\n" + "=" * 80)
        print("步骤3: 显著性分析")
        print("=" * 80)

        kw_result, dunn_matrix, pairwise_df, significant_pairs = perform_significance_analysis(
            df_final,
            stats_df
        )

        # ----------------------------------------------------
        # 步骤4：保存统计和显著性分析结果
        # ----------------------------------------------------
        output_excel = '类别统计与显著性分析结果.xlsx'

        with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
            stats_df.to_excel(
                writer,
                sheet_name='类别统计',
                index=False
            )

            kw_result.to_excel(
                writer,
                sheet_name='Kruskal_Wallis整体检验',
                index=False
            )

            if not dunn_matrix.empty:
                dunn_matrix.to_excel(
                    writer,
                    sheet_name='Dunn_BH校正p值矩阵'
                )

            if not pairwise_df.empty:
                pairwise_df.to_excel(
                    writer,
                    sheet_name='Dunn两两比较长表',
                    index=False
                )

            if not significant_pairs.empty:
                significant_pairs.to_excel(
                    writer,
                    sheet_name='显著差异分类对',
                    index=False
                )

        print(f"\n✅ 已保存: {output_excel}")

        # ----------------------------------------------------
        # 步骤5：绘制小提琴图
        # ----------------------------------------------------
        print("\n" + "=" * 80)
        print("步骤4: 绘制小提琴图")
        print("=" * 80)

        plot_violin_chart(
            df_final,
            stats_df,
            '小提琴图_6x7cm.pdf',
            '小提琴图_6x7cm.svg',
            kw_result=kw_result
        )

        print("\n✅ 分析完成！")

    except Exception as e:
        print(f"❌ 错误: {e}")

        import traceback
        traceback.print_exc()