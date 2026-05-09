import pandas as pd
import networkx as nx
import matplotlib.pyplot as plt
import numpy as np

# 设置matplotlib参数
plt.rcParams['pdf.fonttype'] = 42
plt.rcParams['ps.fonttype'] = 42
plt.rcParams['svg.fonttype'] = 'none'
plt.rcParams['font.sans-serif'] = ['Arial']

# 读取数据
try:
    file_path = "C:/Users/wu'duo'duo/Desktop/Structure Alert.xlsx"
    df = pd.read_excel(file_path, sheet_name='Sheet1')
except:
    df = pd.read_excel("画图数据.xlsx", sheet_name='Sheet1')

# 创建图
G = nx.Graph()

# 添加节点和边
for _, row in df.iterrows():
    f1, f2, co_count, f1_count, co_rate = row
    if not G.has_node(f1):
        G.add_node(f1, size=f1_count)
    if not G.has_node(f2):
        if f2 in df['Fragment1'].values:
            size_val = df[df['Fragment1'] == f2]['Fragment1_Count'].iloc[0]
        else:
            frag1_counts = df[df['Fragment1'] == f2]['Fragment1_Count']
            size_val = frag1_counts.mean() if not frag1_counts.empty else 300
        G.add_node(f2, size=size_val)

    G.add_edge(f1, f2, weight=co_rate, co_count=co_count)

# 设置布局
pos = nx.spring_layout(G, seed=42, k=3.0, iterations=200, scale=2.0)

# 计算节点大小
node_sizes = [G.nodes[n].get('size', 250) for n in G.nodes()]
max_size = max(node_sizes) if node_sizes else 1
node_sizes = [s / max_size * 250 for s in node_sizes]

# 计算边粗细
edge_weights = []
for u, v in G.edges():
    weight = G[u][v].get('weight', 0.1)
    edge_weights.append(weight)

min_weight = min(edge_weights) if edge_weights else 0
max_weight = max(edge_weights) if edge_weights else 1

edge_widths = []
for w in edge_weights:
    normalized = (w - min_weight) / (max_weight - min_weight) if max_weight > min_weight else 0.5
    enhanced = normalized ** 1.5
    width = 0.15 + enhanced * 1.10
    edge_widths.append(width)

# 创建图形
fig, ax = plt.subplots(figsize=(10 / 2.54, 5 / 2.54))
plt.subplots_adjust(left=0.08, right=0.92, bottom=0.08, top=0.92)

ax.axis('off')
ax.set_facecolor('none')

# 水绿色节点
water_green = '#A0DCDC'
edge_color = '#666666'


# 绘制初始图形
def redraw_graph():
    ax.clear()
    ax.axis('off')
    ax.set_facecolor('none')

    # 绘制边
    nx.draw_networkx_edges(G, pos, width=edge_widths, alpha=0.9,
                           edge_color=edge_color, ax=ax)

    # 绘制节点
    nx.draw_networkx_nodes(G, pos, node_size=node_sizes,
                           node_color=water_green, edgecolors='white',
                           linewidths=0.8, ax=ax)

    # 绘制标签
    labels = {node: str(int(node)) for node in G.nodes()}
    text_items = nx.draw_networkx_labels(G, pos, labels=labels,
                                         font_size=6, font_family='Arial', ax=ax)

    for t in text_items.values():
        t.set_clip_on(False)
        t.set_zorder(10)

    fig.canvas.draw()


# 初始绘制
redraw_graph()


# ========== 真正可拖拽的实现 ==========
class DraggableGraph:
    def __init__(self, fig, ax, G, pos, node_sizes):
        self.fig = fig
        self.ax = ax
        self.G = G
        self.pos = pos
        self.node_sizes = node_sizes
        self.dragging = False
        self.dragged_node = None

        # 存储所有图形对象
        self.nodes_artist = None
        self.edges_artist = None
        self.labels_artist = None

        # 连接事件
        self.cid_press = fig.canvas.mpl_connect('button_press_event', self.on_press)
        self.cid_release = fig.canvas.mpl_connect('button_release_event', self.on_release)
        self.cid_motion = fig.canvas.mpl_connect('motion_notify_event', self.on_motion)
        self.cid_key = fig.canvas.mpl_connect('key_press_event', self.on_key)

        # 重新绘制
        self.redraw()

    def redraw(self):
        self.ax.clear()
        self.ax.axis('off')
        self.ax.set_facecolor('none')

        # 绘制边
        self.edges_artist = nx.draw_networkx_edges(self.G, self.pos, width=edge_widths,
                                                   alpha=0.9, edge_color=edge_color, ax=self.ax)

        # 绘制节点
        self.nodes_artist = nx.draw_networkx_nodes(self.G, self.pos, node_size=self.node_sizes,
                                                   node_color=water_green, edgecolors='white',
                                                   linewidths=0.8, ax=self.ax)

        # 绘制标签
        labels = {node: str(int(node)) for node in self.G.nodes()}
        self.labels_artist = nx.draw_networkx_labels(self.G, self.pos, labels=labels,
                                                     font_size=6, font_family='Arial', ax=self.ax)

        for t in self.labels_artist.values():
            t.set_clip_on(False)
            t.set_zorder(10)

        self.fig.canvas.draw()

    def find_nearest_node(self, x, y):
        """找到离点击位置最近的节点"""
        min_dist = float('inf')
        nearest_node = None

        for node, (nx_pos, ny_pos) in self.pos.items():
            # 计算距离
            dist = np.sqrt((x - nx_pos) ** 2 + (y - ny_pos) ** 2)

            # 考虑节点大小（将像素大小转换为坐标距离）
            node_index = list(self.G.nodes()).index(node)
            node_radius = np.sqrt(self.node_sizes[node_index] / np.pi) / 50  # 估算半径

            if dist < node_radius and dist < min_dist:
                min_dist = dist
                nearest_node = node

        return nearest_node

    def on_press(self, event):
        if event.inaxes != self.ax:
            return

        # 找到被点击的节点
        self.dragged_node = self.find_nearest_node(event.xdata, event.ydata)
        if self.dragged_node is not None:
            self.dragging = True
            print(f"开始拖拽节点 {self.dragged_node}")

    def on_motion(self, event):
        if not self.dragging or self.dragged_node is None:
            return

        if event.inaxes != self.ax:
            return

        # 更新节点位置
        self.pos[self.dragged_node] = (event.xdata, event.ydata)

        # 立即重绘（实时更新）
        self.redraw()

    def on_release(self, event):
        if self.dragging and self.dragged_node is not None:
            print(f"释放节点 {self.dragged_node}")
        self.dragging = False
        self.dragged_node = None

    def on_key(self, event):
        if event.key == 's' or event.key == 'S':
            # 保存为SVG矢量图
            plt.savefig('fragment_network_draggable.svg', format='svg',
                        transparent=True, bbox_inches='tight', pad_inches=0.05)
            print("已保存为 fragment_network_draggable.svg")

            # 保存为PDF
            plt.savefig('fragment_network_draggable.pdf', format='pdf',
                        transparent=True, bbox_inches='tight', pad_inches=0.05)
            print("已保存为 fragment_network_draggable.pdf")

            # 保存为PNG
            plt.savefig('fragment_network_draggable.png', format='png', dpi=600,
                        transparent=True, bbox_inches='tight', pad_inches=0.05)
            print("已保存为 fragment_network_draggable.png")

            # 保存节点位置
            with open('dragged_positions.txt', 'w') as f:
                for node, (x, y) in self.pos.items():
                    f.write(f"{node},{x:.4f},{y:.4f}\n")
            print("节点位置已保存到 dragged_positions.txt")


# 创建可拖拽图形
draggable_graph = DraggableGraph(fig, ax, G, pos, node_sizes)

plt.title("点击并拖拽节点 | 按 'S' 保存", fontsize=8, pad=10)
plt.show()