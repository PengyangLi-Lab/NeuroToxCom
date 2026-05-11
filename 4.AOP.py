import pandas as pd
import networkx as nx
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch
import warnings

# 禁用字体警告
warnings.filterwarnings("ignore", category=UserWarning, message="Glyph.*missing from font")


# 字符清理函数
def clean_text(text):
    if isinstance(text, str):
        # 替换制表符为空格
        text = text.replace('\t', ' ')
        # 移除其他可能导致问题的不可见字符
        text = ''.join(c for c in text if ord(c) > 31 or ord(c) == 9)
    return text


# 设置字体为 Times New Roman
font = {'family': 'Times New Roman', 'size': 15, 'weight': 'bold'}  # 字体加粗

# 1. 从两个Excel文件读取数据
try:
    # 读取AOP缩写数据
    abbrev_path = r"C:\Users\wu'duo'duo\Desktop\4.AOP data.xlsx"
    df_abbrev = pd.read_excel(abbrev_path, header=None, names=["FullName", "Abbreviation"])

    # 清理数据中的特殊字符
    df_abbrev["FullName"] = df_abbrev["FullName"].apply(clean_text)
    df_abbrev["Abbreviation"] = df_abbrev["Abbreviation"].apply(clean_text)

    abbrev_map = dict(zip(df_abbrev["FullName"], df_abbrev["Abbreviation"]))
    print("\n缩写表示例：\n", df_abbrev.head())

    # 读取化合物-AOP关联数据
    compound_assoc_path = r"C:\Users\wu'duo'duo\Desktop\4.AOP - Compound-Event Associations data.xlsx"
    df_compound_assoc = pd.read_excel(compound_assoc_path, header=None, names=["Compound", "AOP"])

    # 清理数据中的特殊字符
    df_compound_assoc["Compound"] = df_compound_assoc["Compound"].apply(clean_text)
    df_compound_assoc["AOP"] = df_compound_assoc["AOP"].apply(clean_text)

    print("化合物-AOP关联数据示例：\n", df_compound_assoc.head())

    # 读取事件关联数据
    assoc_path = r"C:\Users\wu'duo'duo\Desktop\4.AOP-Event Associations data.xlsx"
    df_assoc = pd.read_excel(assoc_path)

    print("关联数据示例：\n", df_assoc.head())

except Exception as e:
    print(f"文件读取错误：{e}")
    exit()

# 2. 数据预处理
# 应用缩写映射
df_assoc["event1_Abbrev"] = df_assoc["event 1"].apply(
    lambda x: abbrev_map.get(str(x).strip(), x) if pd.notnull(x) else x  # 保留未缩写的原始名称
)
df_assoc["event2_Abbrev"] = df_assoc["event 2"].apply(
    lambda x: abbrev_map.get(str(x).strip(), x) if pd.notnull(x) else x  # 保留未缩写的原始名称
)

# 筛选出仅包含缩写的事件关联
filtered_assoc = df_assoc[df_assoc["event1_Abbrev"].isin(df_abbrev["Abbreviation"])].dropna()
filtered_assoc = filtered_assoc[filtered_assoc["event2_Abbrev"].isin(df_abbrev["Abbreviation"])].dropna()

# 3. 创建网络图
G = nx.Graph()

# 添加AOP节点
aop_size = 1000  # AOP节点变大
aop_color = "#2E75B6"  # AOP节点颜色改为绿色（温和的绿色）

for event in filtered_assoc["event1_Abbrev"].unique():
    G.add_node(event, node_type="aop", size=aop_size, color=aop_color)

for event in filtered_assoc["event2_Abbrev"].unique():
    G.add_node(event, node_type="aop", size=aop_size, color=aop_color)

# 添加AOP边
for _, row in filtered_assoc.iterrows():
    G.add_edge(row["event1_Abbrev"], row["event2_Abbrev"])

# 添加化合物节点
compound_size = 1000  # 化合物节点变大
compound_color = "#7CB342"  # 化合物节点颜色改为蓝色（温和的蓝色）

for compound in df_compound_assoc["Compound"].unique():
    G.add_node(compound, node_type="compound", size=compound_size, color=compound_color)

# 添加化合物与AOP的边
for _, row in df_compound_assoc.iterrows():
    compound = row["Compound"]
    aop = abbrev_map.get(row["AOP"], row["AOP"])  # 使用缩写映射
    if aop in G.nodes():
        G.add_edge(compound, aop)

# 4. 可视化设置
plt.figure(figsize=(12, 12))

# 设置全局字体为 Times New Roman
plt.rcParams['font.family'] = 'Times New Roman'
plt.rcParams['font.weight'] = 'bold'

# 布局算法
pos = nx.kamada_kawai_layout(G) if len(G.nodes()) > 50 else nx.spring_layout(G, k=0.15, seed=42)

# 交互式绘图
plt.ion()  # 开启交互模式

# 绘制节点
node_colors = [data["color"] for _, data in G.nodes(data=True)]
node_sizes = [data["size"] for _, data in G.nodes(data=True)]

# 创建节点集合
nodes = nx.draw_networkx_nodes(
    G, pos,
    node_size=node_sizes,
    node_color=node_colors,
    alpha=0.9,
    linewidths=0.0,  # 移除节点边框
    edgecolors="none"  # 移除节点边框
)

# 创建边集合
edges = nx.draw_networkx_edges(
    G, pos,
    width=1.0,
    alpha=0.1,
    edge_color="#000000"
)

# 创建标签集合
labels = nx.draw_networkx_labels(
    G, pos,
    font_family='Times New Roman',
    font_size=font['size'],
    font_weight='bold',  # 字体加粗
    bbox=dict(
        facecolor="none",  # 移除白色背景
        alpha=0.0,
        edgecolor="none",
        boxstyle="round,pad=0.3"
    )
)

# 添加图例和标题
legend_elements = [
    plt.Line2D([0], [0], marker="o", color="w", label="AOP Events",
               markerfacecolor=aop_color, markersize=15),
    plt.Line2D([0], [0], marker="o", color="w", label="Compounds",
               markerfacecolor=compound_color, markersize=15)
]

plt.legend(
    handles=legend_elements,
    loc="upper right",
    prop={"family": 'Times New Roman', "size": 12, "weight": "bold"},
    framealpha=0.9
)

plt.title(
    "Adverse Outcome Pathways Network (Interactive Mode)",
    fontsize=16,
    fontdict={'family': 'Times New Roman', 'weight': 'bold'},
    pad=20
)

plt.axis("off")
plt.tight_layout()

# 交互功能：拖动节点
selected_node = None
offset = (0, 0)


def on_click(event):
    global selected_node, offset
    if event.inaxes != plt.gca():
        return

    # 检查是否点击了节点
    for node, (x, y) in pos.items():
        if (x - event.xdata) ** 2 + (y - event.ydata) ** 2 < 0.02:  # 阈值可调整
            selected_node = node
            offset = (x - event.xdata, y - event.ydata)
            break


def on_motion(event):
    global selected_node, pos
    if selected_node is None:
        return
    if event.inaxes != plt.gca():
        return

    # 更新节点位置
    x, y = event.xdata + offset[0], event.ydata + offset[1]
    pos[selected_node] = (x, y)

    # 更新图形
    nodes.set_offsets([pos[node] for node in G.nodes()])

    # 更新边
    edges.set_segments([(pos[u], pos[v]) for u, v in G.edges()])

    # 更新标签位置
    for node, text in labels.items():
        text.set_position(pos[node])

    plt.draw()


def on_release(event):
    global selected_node
    selected_node = None


# 连接事件处理函数
plt.gcf().canvas.mpl_connect('button_press_event', on_click)
plt.gcf().canvas.mpl_connect('motion_notify_event', on_motion)
plt.gcf().canvas.mpl_connect('button_release_event', on_release)


# 保存功能
def save_figure(event):
    if event.key == 's':
        plt.savefig("AOP_Network_with_Abbreviations_and_Compounds.png", dpi=300, bbox_inches="tight")
        print("已保存当前图形")


plt.gcf().canvas.mpl_connect('key_press_event', save_figure)

print("交互模式已启动：")
print("- 点击并拖动节点可以调整位置")
print("- 按 's' 键保存当前图形")

plt.show(block=True)  # 保持图形窗口打开