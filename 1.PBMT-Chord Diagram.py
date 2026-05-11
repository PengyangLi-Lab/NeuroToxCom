import matplotlib.pyplot as plt
import numpy as np
from matplotlib.path import Path
from matplotlib.patches import PathPatch, Arc
import matplotlib.colors as mcolors
import pandas as pd
from collections import defaultdict
import os

# ================================
# 可调节参数设置
# ================================

# 连接线参数 - 只包含指定的连接
CONNECTION_PARAMS = {
    # P组连接
    ('P', 'B'): {
        'base_linewidth': 0.2,
        'alpha_range': (0.3, 0.3),
        'color_brightness': 0.0,
        'curve_strength': 0.5,
        'zorder': 3,
        'custom_color': None
    },
    ('P', 'vB'): {
        'base_linewidth': 0.2,
        'alpha_range': (0.3, 0.3),
        'color_brightness': 0.0,
        'curve_strength': 0.5,
        'zorder': 3,
        'custom_color': None
    },
    ('P', 'M'): {
        'base_linewidth': 0.2,
        'alpha_range': (0.3, 0.3),
        'color_brightness': 0.0,
        'curve_strength': 0.5,
        'zorder': 1,
        'custom_color': None
    },
    ('P', 'Vm'): {
        'base_linewidth': 0.2,
        'alpha_range': (0.3, 0.3),
        'color_brightness': 0.0,
        'curve_strength': 0.5,
        'zorder': 2,
        'custom_color': None
    },

    # Vp组连接
    ('Vp', 'B'): {
        'base_linewidth': 0.2,
        'alpha_range': (0.4, 0.9),
        'color_brightness': 0.2,
        'curve_strength': 0.5,
        'zorder': 2,
        'custom_color': None
    },
    ('Vp', 'vB'): {
        'base_linewidth': 0.2,
        'alpha_range': (0.3, 0.3),
        'color_brightness': 0.0,
        'curve_strength': 0.5,
        'zorder': 1,
        'custom_color': None
    },
    ('Vp', 'M'): {
        'base_linewidth': 0.2,
        'alpha_range': (0.3, 0.5),
        'color_brightness': 0.1,
        'curve_strength': 0.5,
        'zorder': 1,
        'custom_color': None
    },
    ('Vp', 'Vm'): {
        'base_linewidth': 0.2,
        'alpha_range': (0.4, 0.9),
        'color_brightness': 0.0,
        'curve_strength': 0.5,
        'zorder': 1,
        'custom_color': None
    },

    # B组连接
    ('B', 'M'): {
        'base_linewidth': 0.2,
        'alpha_range': (0.3, 0.3),
        'color_brightness': 0.0,
        'curve_strength': 0.5,
        'zorder': 2,
        'custom_color': None
    },
    ('B', 'Vm'): {
        'base_linewidth': 0.2,
        'alpha_range': (0.4, 0.9),
        'color_brightness': 0,
        'curve_strength': 0.5,
        'zorder': 2,
        'custom_color': None
    },

    # vB组连接
    ('vB', 'M'): {
        'base_linewidth': 0.2,
        'alpha_range': (0.4, 0.9),
        'color_brightness': 0,
        'curve_strength': 0.5,
        'zorder': 2,
        'custom_color': None
    },
    ('vB', 'Vm'): {
        'base_linewidth': 0.2,
        'alpha_range': (0.4, 0.9),
        'color_brightness': 0.0,
        'curve_strength': 0.5,
        'zorder': 2,
        'custom_color': None
    }
}

# 节点弧线参数
ARC_PARAMS = {
    'P': {
        'linewidth': 9,
        'alpha': 1
    },
    'other': {
        'linewidth': 9,
        'alpha': 1
    }
}

# 节点颜色设置
NODE_COLORS = {
    'P': '#A2C5DB',
    'Vp': '#8AB6D5',
    'B': '#EDC669',
    'vB': '#F4A666',
    'M': '#D2E4C5',
    'Vm': '#75B0A0'
}

# 图形参数 - 改为9x9cm (3.54x3.54英寸)
FIGURE_PARAMS = {
    'figsize': (3.54, 3.54),  # 9cm x 9cm
    'outer_radius': 1.0,
    'inner_radius': 0.5,
    'label_radius': 1.35,
    'gap_angle': 0.01
}


def parse_excel_data_correctly():
    """正确读取Excel数据"""
    try:
        file_path = os.path.join(os.path.expanduser("~"), "Desktop", "1.PBMT-Chord Diagram date.xlsx")
        df = pd.read_excel(file_path, sheet_name='Sheet2')

        # 确保列名正确
        required_columns = ['化合物', 'T', 'P', 'Vp', 'B', 'vB', 'M', 'Vm']
        for col in required_columns:
            if col not in df.columns:
                print(f"错误: 缺少列 {col}")
                return None, None, None

        # 创建节点数据
        nodes_data = {}
        for col in ['P', 'Vp', 'B', 'vB', 'M', 'Vm']:
            # 统计该列为1的数量
            count = df[col].sum()
            nodes_data[col] = count
            print(f"{col}: {count} compounds")

        # 创建化合物类别映射
        compound_categories = {}
        for _, row in df.iterrows():
            compound_id = int(row['化合物'])
            categories = []
            for col in ['P', 'Vp', 'B', 'vB', 'M', 'Vm']:
                if row[col] == 1:
                    categories.append(col)
            compound_categories[compound_id] = categories

        # 创建位置映射
        category_compound_positions = {}
        for col in ['P', 'Vp', 'B', 'vB', 'M', 'Vm']:
            # 获取该列为1的所有化合物
            compounds = df[df[col] == 1]['化合物'].tolist()
            compounds_sorted = sorted(compounds)
            category_compound_positions[col] = {}
            for idx, compound_id in enumerate(compounds_sorted):
                category_compound_positions[col][compound_id] = idx / len(compounds_sorted)

        return nodes_data, compound_categories, category_compound_positions, df

    except Exception as e:
        print(f"读取Excel文件时出错: {e}")
        return None, None, None, None


def export_connections_to_excel(compound_categories, original_df):
    """将连接关系信息导出到新的Excel文件 - 保持原始化合物顺序，用空格代替0"""
    try:
        # 指定的连接关系
        specified_connections = [
            ('P', 'B'), ('P', 'vB'), ('P', 'M'), ('P', 'Vm'),
            ('Vp', 'B'), ('Vp', 'vB'), ('Vp', 'M'), ('Vp', 'Vm'),
            ('B', 'M'), ('B', 'Vm'),
            ('vB', 'M'), ('vB', 'Vm')
        ]

        # 创建新的DataFrame，保持原始化合物的顺序
        connection_columns = [f"{source}-{target}" for source, target in specified_connections]

        # 创建空的DataFrame
        connections_df = pd.DataFrame()

        # 添加化合物ID列，保持原始顺序
        connections_df['化合物'] = original_df['化合物']

        # 为每个连接关系创建列
        for i, (source, target) in enumerate(specified_connections):
            col_name = f"{source}-{target}"
            # 初始化为空字符串（空格）
            connections_df[col_name] = ""

            # 为每个化合物设置连接关系
            for idx, row in original_df.iterrows():
                compound_id = int(row['化合物'])
                if (source in compound_categories[compound_id] and
                        target in compound_categories[compound_id]):
                    connections_df.at[idx, col_name] = 1

        # 保存到Excel文件
        output_path = os.path.join(os.path.expanduser("~"), "Desktop", "连接关系统计.xlsx")

        # 创建Excel writer对象，支持多个sheet
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            # 保存连接关系数据（保持原始顺序）
            connections_df.to_excel(writer, sheet_name='连接关系数据', index=False)

            # 创建连接统计汇总
            connection_stats = []
            for source, target in specified_connections:
                col_name = f"{source}-{target}"
                count = (connections_df[col_name] == 1).sum()  # 统计1的数量
                connection_stats.append({
                    '连接关系': f'{source}-{target}',
                    '共享化合物数量': count,
                    '源节点': source,
                    '目标节点': target
                })

            stats_df = pd.DataFrame(connection_stats)
            stats_df.to_excel(writer, sheet_name='连接统计汇总', index=False)

            # 创建节点统计
            node_stats = []
            for node in ['P', 'Vp', 'B', 'vB', 'M', 'Vm']:
                count = original_df[node].sum()
                node_stats.append({
                    '节点': node,
                    '化合物数量': count
                })

            node_df = pd.DataFrame(node_stats)
            node_df.to_excel(writer, sheet_name='节点统计', index=False)

        print(f"\n连接关系信息已导出到: {output_path}")
        print(f"连接关系数据: {len(connections_df)} 个化合物，{len(specified_connections)} 种连接关系")
        print(f"连接统计汇总: {len(stats_df)} 种连接关系")

        # 显示前几行数据作为示例
        print("\n前5行连接关系数据示例:")
        print(connections_df.head())

        return connections_df

    except Exception as e:
        print(f"导出Excel文件时出错: {e}")
        return None


def verify_specified_connections(nodes_data, compound_categories):
    """验证指定的连接数据"""
    print("\n=== 验证指定连接数据 ===")

    # 只检查指定的连接
    specified_connections = [
        ('P', 'B'), ('P', 'vB'), ('P', 'M'), ('P', 'Vm'),
        ('Vp', 'B'), ('Vp', 'vB'), ('Vp', 'M'), ('Vp', 'Vm'),
        ('B', 'M'), ('B', 'Vm'),
        ('vB', 'M'), ('vB', 'Vm')
    ]

    connection_stats = {}
    for source, target in specified_connections:
        count = 0
        for compound_id, categories in compound_categories.items():
            if source in categories and target in categories:
                count += 1
        connection_stats[(source, target)] = count
        if count > 0:
            print(f"{source}-{target}: {count} 个共享化合物")

    return connection_stats


def create_chord_diagram_specified(nodes_data, compound_categories, category_compound_positions):
    """创建只包含指定连接的弦图"""
    # 设置透明背景
    fig, ax = plt.subplots(figsize=FIGURE_PARAMS['figsize'])
    fig.patch.set_alpha(0.0)
    ax.set_facecolor('none')

    # 调整坐标轴范围以适应小画布
    ax.set_xlim(-1.2, 1.2)
    ax.set_ylim(-1.2, 1.2)
    ax.set_aspect('equal')
    ax.axis('off')

    # 计算节点角度
    node_names = list(nodes_data.keys())
    node_values = list(nodes_data.values())
    total_value = sum(node_values)

    if total_value == 0:
        ax.text(0, 0, "无数据可显示", ha='center', va='center', fontsize=16)
        return fig, ax

    start_angles, end_angles = calculate_node_angles(node_values, FIGURE_PARAMS['gap_angle'])

    # 只绘制指定的连接线
    total_connections = draw_specified_connections(ax, nodes_data, compound_categories, category_compound_positions,
                                                   start_angles, end_angles, node_names)

    # 绘制节点（不绘制标签）
    draw_nodes_no_labels(ax, nodes_data, start_angles, end_angles, node_names)

    print(f"\n总共绘制了 {total_connections} 条连接线")
    return fig, ax


def calculate_node_angles(node_values, gap_angle):
    """计算节点角度"""
    start_angles, end_angles = [], []
    current_angle = 0
    total_value = sum(node_values)

    for value in node_values:
        angle = 2 * np.pi * (value / total_value)
        start_angles.append(current_angle + gap_angle / 2)
        end_angles.append(current_angle + angle - gap_angle / 2)
        current_angle += angle

    return start_angles, end_angles


def draw_specified_connections(ax, nodes_data, compound_categories, category_compound_positions,
                               start_angles, end_angles, node_names):
    """只绘制指定的连接线"""
    print("\n=== 开始绘制指定连接线 ===")

    # 只绘制这些连接
    specified_connections = [
        ('P', 'B'), ('P', 'vB'), ('P', 'M'), ('P', 'Vm'),
        ('Vp', 'B'), ('Vp', 'vB'), ('Vp', 'M'), ('Vp', 'Vm'),
        ('B', 'M'), ('B', 'Vm'),
        ('vB', 'M'), ('vB', 'Vm')
    ]

    total_connections = 0

    # 为每个指定的节点对绘制连接线
    for source, target in specified_connections:
        connections_count = draw_single_connection_direct(ax, source, target, compound_categories,
                                                          category_compound_positions, start_angles,
                                                          end_angles, node_names)
        total_connections += connections_count

    return total_connections


def draw_single_connection_direct(ax, source, target, compound_categories, category_compound_positions,
                                  start_angles, end_angles, node_names):
    """直接绘制单个连接的每条线"""
    # 获取连接参数
    connection_key = (source, target)
    if connection_key in CONNECTION_PARAMS:
        params = CONNECTION_PARAMS[connection_key]
    else:
        print(f"警告: 未找到连接 {source}-{target} 的参数配置，跳过绘制")
        return 0

    # 收集共享化合物
    shared_compounds = []
    for compound_id, categories in compound_categories.items():
        if source in categories and target in categories:
            shared_compounds.append(compound_id)

    if not shared_compounds:
        return 0

    print(f"绘制 {source}-{target}: {len(shared_compounds)} 条连接线")

    # 获取节点角度
    source_idx = node_names.index(source)
    target_idx = node_names.index(target)
    source_start, source_end = start_angles[source_idx], end_angles[source_idx]
    target_start, target_end = start_angles[target_idx], end_angles[target_idx]

    # 获取颜色
    if params['custom_color']:
        base_color = mcolors.to_rgb(params['custom_color'])
    else:
        base_color = mcolors.to_rgb(NODE_COLORS[target])

    light_color = tuple(min(1.0, c + params['color_brightness']) for c in base_color)

    connections_drawn = 0

    # 为每个共享化合物绘制一条线
    for compound_id in shared_compounds:
        if (compound_id in category_compound_positions[source] and
                compound_id in category_compound_positions[target]):
            # 获取位置比例
            source_ratio = category_compound_positions[source][compound_id]
            target_ratio = category_compound_positions[target][compound_id]

            # 计算具体角度
            source_angle = source_start + (source_end - source_start) * source_ratio
            target_angle = target_start + (target_end - target_start) * target_ratio

            # 计算线宽和透明度
            line_width = params['base_linewidth']
            alpha_min, alpha_max = params['alpha_range']
            alpha = alpha_min  # 固定透明度

            # 计算弯曲程度
            curve_factor = params['curve_strength']
            control_radius = FIGURE_PARAMS['outer_radius'] + (
                    FIGURE_PARAMS['inner_radius'] - FIGURE_PARAMS['outer_radius']) * curve_factor

            # 创建贝塞尔曲线
            verts = [
                (FIGURE_PARAMS['outer_radius'] * np.cos(source_angle),
                 FIGURE_PARAMS['outer_radius'] * np.sin(source_angle)),
                (control_radius * np.cos(source_angle),
                 control_radius * np.sin(source_angle)),
                (control_radius * np.cos(target_angle),
                 control_radius * np.sin(target_angle)),
                (FIGURE_PARAMS['outer_radius'] * np.cos(target_angle),
                 FIGURE_PARAMS['outer_radius'] * np.sin(target_angle))
            ]

            path = Path(verts, [Path.MOVETO, Path.CURVE4, Path.CURVE4, Path.CURVE4])
            patch = PathPatch(path, facecolor='none', edgecolor=light_color,
                              alpha=alpha, linewidth=line_width, zorder=params['zorder'])
            ax.add_patch(patch)
            connections_drawn += 1

    return connections_drawn


def draw_nodes_no_labels(ax, nodes, start_angles, end_angles, node_names):
    """绘制节点（不绘制标签）"""
    print("\n绘制节点（无标签）...")
    for i, (name, value) in enumerate(nodes.items()):
        if value == 0:
            continue

        start_angle_rad = start_angles[i]
        end_angle_rad = end_angles[i]

        # 绘制外弧线
        start_angle = np.degrees(start_angle_rad)
        end_angle = np.degrees(end_angle_rad)

        arc_params = ARC_PARAMS['P'] if name == 'P' else ARC_PARAMS['other']
        arc = Arc((0, 0), 2 * FIGURE_PARAMS['outer_radius'], 2 * FIGURE_PARAMS['outer_radius'],
                  theta1=start_angle, theta2=end_angle,
                  color=NODE_COLORS[name], linewidth=arc_params['linewidth'],
                  alpha=arc_params['alpha'], zorder=3)
        ax.add_patch(arc)


def main():
    """主函数"""
    print("=== 开始处理化合物数据 ===")

    # 正确读取Excel数据
    nodes_data, compound_categories, category_compound_positions, original_df = parse_excel_data_correctly()

    if nodes_data is None:
        print("数据读取失败")
        return

    # 验证指定的连接数据
    connection_stats = verify_specified_connections(nodes_data, compound_categories)

    # 导出连接关系到Excel
    connections_df = export_connections_to_excel(compound_categories, original_df)

    # 特别检查PB和PM
    print(f"\n=== 重点检查 ===")
    print(f"P-B 连接: {connection_stats.get(('P', 'B'), 0)} 个化合物")
    print(f"P-M 连接: {connection_stats.get(('P', 'M'), 0)} 个化合物")

    # 创建弦图
    print("\n=== 开始创建弦图（9x9cm，无标签）===")
    fig, ax = create_chord_diagram_specified(nodes_data, compound_categories, category_compound_positions)

    plt.tight_layout()
    plt.show()

    # 保存图片
    try:
        fig.savefig('9cm_chord_diagram.svg', format='svg', dpi=600, bbox_inches='tight',
                    transparent=True, facecolor='none', edgecolor='none')
        fig.savefig('9cm_chord_diagram.pdf', format='pdf', bbox_inches='tight',
                    transparent=True, facecolor='none', edgecolor='none')
        fig.savefig('9cm_chord_diagram.png', format='png', dpi=600, bbox_inches='tight',
                    transparent=True, facecolor='none', edgecolor='none')
        print("9x9cm图片已保存")
    except Exception as e:
        print(f"保存图片时出错: {e}")

    print("完成!")


if __name__ == "__main__":
    main()