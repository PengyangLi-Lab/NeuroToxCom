# -*- coding: utf-8 -*-
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
from scipy import stats
from sklearn.linear_model import LinearRegression
from sklearn.metrics import r2_score
import seaborn as sns

# 设置全局字体 - 使用标准的字号对应关系
plt.rcParams['font.family'] = 'Arial'
plt.rcParams['font.size'] = 10.5  # 5号字 ≈ 10.5磅
plt.rcParams['axes.titlesize'] = 6  # 5号字
plt.rcParams['axes.labelsize'] = 10.5  # 5号字
plt.rcParams['xtick.labelsize'] = 10.5  # 5号字
plt.rcParams['ytick.labelsize'] = 10.5  # 5号字
plt.rcParams['legend.fontsize'] = 10.5  # 5号字


# ------------------------- 1. 读取并预处理数据 -------------------------
def load_data(file_path):
    try:
        df = pd.read_excel(file_path, engine='openpyxl')
        print("实际列名:", df.columns.tolist())
        df.columns = df.columns.str.replace(' ', '_').str.replace('-', '_')
        print("标准化后列名:", df.columns.tolist())

        # 转换必要的数值列
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

        return df
    except Exception as e:
        print(f"读取文件失败: {e}")
        return None


# ------------------------- 识别PBT和PMT化合物 -------------------------
def identify_pbt_pmt_compounds(df):
    """识别PBT和PMT化合物"""
    pbt_compounds = pd.DataFrame()
    pmt_compounds = pd.DataFrame()
    both_compounds = pd.DataFrame()

    # PBT标准
    pbt_thresholds = {
        'BAF': 2000,
        'Half_life_in_water_(hours)': 960,
        'Half_life_in_soil_(hours)': 2880,
        'Half_life_in_sediment_(hours)': 2880
    }

    # PMT标准
    pmt_thresholds = {
        'LogD': 3.5,
        'Half_life_in_water_(hours)': 960,
        'Half_life_in_soil_(hours)': 2880,
        'Half_life_in_sediment_(hours)': 2880
    }

    # 检查PBT化合物
    if all(col in df.columns for col in ['BAF', 'Half_life_in_water_(hours)']):
        soil_col = 'Half_life_in_soil_(hours)' if 'Half_life_in_soil_(hours)' in df.columns else None
        sediment_col = 'Half_life_in_sediment_(hours)' if 'Half_life_in_sediment_(hours)' in df.columns else None

        persistence_conditions = [
            (df['Half_life_in_water_(hours)'] > pbt_thresholds['Half_life_in_water_(hours)'])
        ]

        if soil_col:
            persistence_conditions.append(df[soil_col] > pbt_thresholds['Half_life_in_soil_(hours)'])
        if sediment_col:
            persistence_conditions.append(df[sediment_col] > pbt_thresholds['Half_life_in_sediment_(hours)'])

        persistence_mask = persistence_conditions[0]
        for condition in persistence_conditions[1:]:
            persistence_mask = persistence_mask | condition

        pbt_mask = (df['BAF'] > pbt_thresholds['BAF']) & persistence_mask
        pbt_compounds = df[pbt_mask].copy()

    # 检查PMT化合物
    if all(col in df.columns for col in ['LogD', 'Half_life_in_water_(hours)']):
        soil_col = 'Half_life_in_soil_(hours)' if 'Half_life_in_soil_(hours)' in df.columns else None
        sediment_col = 'Half_life_in_sediment_(hours)' if 'Half_life_in_sediment_(hours)' in df.columns else None

        persistence_conditions = [
            (df['Half_life_in_water_(hours)'] > pmt_thresholds['Half_life_in_water_(hours)'])
        ]

        if soil_col:
            persistence_conditions.append(df[soil_col] > pmt_thresholds['Half_life_in_soil_(hours)'])
        if sediment_col:
            persistence_conditions.append(df[sediment_col] > pmt_thresholds['Half_life_in_sediment_(hours)'])

        persistence_mask = persistence_conditions[0]
        for condition in persistence_conditions[1:]:
            persistence_mask = persistence_mask | condition

        pmt_mask = (df['LogD'] < pmt_thresholds['LogD']) & persistence_mask
        pmt_compounds = df[pmt_mask].copy()

    # 识别同时属于PBT和PMT的化合物
    if not pbt_compounds.empty and not pmt_compounds.empty:
        both_mask = pbt_compounds.index.isin(pmt_compounds.index)
        both_compounds = pbt_compounds[both_mask].copy()

    return pbt_compounds, pmt_compounds, both_compounds


# ------------------------- 线性拟合函数 -------------------------
def perform_linear_regression(x, y, group_name):
    """执行线性回归并返回统计信息"""
    if len(x) < 2:
        return None, None, None, None, None, None, None

    # 移除NaN值
    mask = ~(np.isnan(x) | np.isnan(y))
    x_clean = x[mask]
    y_clean = y[mask]

    if len(x_clean) < 2:
        return None, None, None, None, None, None, None

    # 线性回归
    slope, intercept, r_value, p_value, std_err = stats.linregress(x_clean, y_clean)
    r_squared = r_value ** 2

    # 计算95%置信区间
    n = len(x_clean)
    t_value = stats.t.ppf(0.975, n - 2)  # 95%置信区间的t值
    slope_ci = (slope - t_value * std_err, slope + t_value * std_err)
    intercept_ci = (intercept - t_value * std_err, intercept + t_value * std_err)

    # 预测值
    x_pred = np.linspace(x_clean.min(), x_clean.max(), 100)
    y_pred = slope * x_pred + intercept

    print(f"\n{group_name} 线性回归结果:")
    print(f"样本数量: {n}")
    print(f"回归方程: y = {slope:.4f}x + {intercept:.4f}")
    print(f"R² = {r_squared:.4f}")
    print(f"p值 = {p_value:.4e}")
    print(f"斜率95%置信区间: [{slope_ci[0]:.4f}, {slope_ci[1]:.4f}]")
    print(f"截距95%置信区间: [{intercept_ci[0]:.4f}, {intercept_ci[1]:.4f}]")

    return slope, intercept, r_squared, p_value, x_pred, y_pred

# ------------------------- 绘制拟合图 -------------------------
def plot_linear_fit(ax, x_data, y_data, group_name, color):
    """绘制单个线性拟合图"""
    if len(x_data) < 2:
        ax.text(0.5, 0.5, f'Not enough data\nfor {group_name}',
                transform=ax.transAxes, ha='center', va='center', fontsize=10.5)
        return

    # 执行线性回归
    result = perform_linear_regression(x_data, y_data, group_name)
    if result[0] is None:
        ax.text(0.5, 0.5, f'Not enough data\nfor {group_name}',
                transform=ax.transAxes, ha='center', va='center', fontsize=10.5)
        return

    slope, intercept, r_squared, p_value, x_pred, y_pred = result

    # 计算R值（相关系数）
    r_value = np.sqrt(r_squared) if r_squared >= 0 else -np.sqrt(-r_squared)

    # 格式化p值
    if p_value == 0:
        p_text = '$\it{p}$ = 0'
    elif p_value < 0.001:
        p_text = '$\it{p}$ < 0.001'
    else:
        p_text = f'$\it{{p}}$ = {p_value:.3f}'

    # 使用同一颜色不同透明度
    point_color = color  # 点颜色（中等透明度）
    line_color = color  # 线颜色（最深）

    # 绘制散点（中等透明度）
    ax.scatter(x_data, y_data, alpha=0.4, color=point_color, s=15)

    # 绘制回归线（最深）
    ax.plot(x_pred, y_pred, color=line_color, linewidth=1.5)

    stats_info = [
        {'text': f'Log $K_{{AW}}$ = {slope:.4f}Log $K_{{OA}}$ + {intercept:.4f}', 'x': 0.05, 'y': 0.1},
        {'text': f'$\it{{r}}$ = {r_value:.4f}', 'x': 0.05, 'y': 0.175},
        {'text': p_text, 'x': 0.05, 'y': 0.25},
    ]

    for info in stats_info:
        ax.text(info['x'], info['y'], info['text'],
                transform=ax.transAxes, verticalalignment='top',
                bbox=dict(boxstyle='round,pad=0.2', facecolor='white', alpha=0.8, linewidth=0.0),
                fontsize=10.5)  # 5号字

    ax.set_xlabel('Log $K_{OA}$', fontsize=10.5, labelpad=2)
    ax.set_ylabel('Log $K_{AW}$', fontsize=10.5, labelpad=2)
    ax.set_title(f'{group_name}', fontsize=13.5, pad=3)  # 修改为13.5，去掉加粗
    ax.grid(True, alpha=0.3, linewidth=0.3)

    # 设置坐标轴刻度字体大小
    ax.tick_params(axis='both', which='major', labelsize=10.5)


# ------------------------- 主程序 -------------------------
def main():
    file_path = r"C:\Users\wu'duo'duo\Desktop\Fitting Plot.xlsx"
    df = load_data(file_path)
    if df is None:
        return

    print("\n数据预览：")
    print(df.head())

    # 识别PBT和PMT化合物
    pbt_compounds, pmt_compounds, both_compounds = identify_pbt_pmt_compounds(df)

    print(f"\n化合物统计:")
    print(f"总化合物数量: {len(df)}")
    print(f"PBT化合物数量: {len(pbt_compounds)}")
    print(f"PMT化合物数量: {len(pmt_compounds)}")
    print(f"Both化合物数量: {len(both_compounds)}")

    # 创建画布 - 18x4英寸，四个子图排成一行
    fig, axes = plt.subplots(1, 4, figsize=(18, 4))

    # 定义颜色（每组使用不同颜色）
    colors = ['#9467bd', 'red', '#1f77b4', '#2ca02c']  # 蓝色, 橙色, 绿色, 紫色

    # 绘制所有化合物的拟合图
    plot_linear_fit(axes[0], df['logkoa'], df['Log_Kaw'], 'All', colors[0])

    # 绘制PBT化合物的拟合图
    if len(pbt_compounds) > 0:
        plot_linear_fit(axes[1], pbt_compounds['logkoa'], pbt_compounds['Log_Kaw'], 'PBT', colors[1])
    else:
        axes[1].text(0.5, 0.5, 'No PBT compounds\nfound', transform=axes[1].transAxes,
                     ha='center', va='center', fontsize=10.5)

    # 绘制PMT化合物的拟合图
    if len(pmt_compounds) > 0:
        plot_linear_fit(axes[2], pmt_compounds['logkoa'], pmt_compounds['Log_Kaw'], 'PMT', colors[2])
    else:
        axes[2].text(0.5, 0.5, 'No PMT compounds\nfound', transform=axes[2].transAxes,
                     ha='center', va='center', fontsize=10.5)

    # 绘制Both化合物的拟合图
    if len(both_compounds) > 0:
        plot_linear_fit(axes[3], both_compounds['logkoa'], both_compounds['Log_Kaw'], 'Both', colors[3])
    else:
        axes[3].text(0.5, 0.5, 'No Both compounds\nfound', transform=axes[3].transAxes,
                     ha='center', va='center', fontsize=10.5)

    plt.tight_layout()

    # 保存多种格式的矢量图（透明背景）
    plt.savefig('linear_regression_analysis.pdf', bbox_inches='tight', dpi=300, transparent=True)  # PDF格式
    plt.savefig('linear_regression_analysis.svg', bbox_inches='tight', dpi=300, transparent=True)  # SVG格式
    plt.savefig('linear_regression_analysis.eps', bbox_inches='tight', dpi=300, transparent=True)  # EPS格式
    # 同时保存高分辨率位图（透明背景）
    plt.savefig('linear_regression_analysis.png', dpi=300, bbox_inches='tight', transparent=True)

    plt.show()

    # 导出详细结果到Excel
    try:
        with pd.ExcelWriter('regression_results.xlsx', engine='openpyxl') as writer:
            # 汇总统计
            summary_data = []
            groups = [
                ('All Compounds', df['logkoa'], df['Log_Kaw']),
                ('PBT Compounds', pbt_compounds['logkoa'] if len(pbt_compounds) > 0 else [],
                 pbt_compounds['Log_Kaw'] if len(pbt_compounds) > 0 else []),
                ('PMT Compounds', pmt_compounds['logkoa'] if len(pmt_compounds) > 0 else [],
                 pmt_compounds['Log_Kaw'] if len(pmt_compounds) > 0 else []),
                ('Both Compounds', both_compounds['logkoa'] if len(both_compounds) > 0 else [],
                 both_compounds['Log_Kaw'] if len(both_compounds) > 0 else [])
            ]

            for group_name, x, y in groups:
                if len(x) > 1:
                    result = perform_linear_regression(x, y, group_name)
                    if result[0] is not None:
                        slope, intercept, r_squared, p_value, _, _ = result
                        summary_data.append({
                            'Group': group_name,
                            'Sample Size': len(x),
                            'Slope': slope,
                            'Intercept': intercept,
                            'R-squared': r_squared,
                            'p-value': p_value,
                            'Equation': f'Log K_AW = {slope:.4f}Log K_OA + {intercept:.4f}'
                        })

            summary_df = pd.DataFrame(summary_data)
            summary_df.to_excel(writer, sheet_name='Regression_Summary', index=False)

            # 导出原始数据
            df_with_labels = df.copy()
            df_with_labels['PBT'] = df.index.isin(pbt_compounds.index) if len(pbt_compounds) > 0 else False
            df_with_labels['PMT'] = df.index.isin(pmt_compounds.index) if len(pmt_compounds) > 0 else False
            df_with_labels['Both'] = df.index.isin(both_compounds.index) if len(both_compounds) > 0 else False
            df_with_labels.to_excel(writer, sheet_name='All_Data', index=False)

        print(f"\n详细结果已导出到: regression_results.xlsx")

    except Exception as e:
        print(f"导出Excel文件失败: {e}")


if __name__ == "__main__":
    main()