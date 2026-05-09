import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import os
from docx import Document
import shutil
import datetime

# ====================== 配置部分 ======================
# 颜色方案
colors = {
    "primary": "#3a7ca5",  # Steel blue
    "secondary": "#5c8bad",  # Lighter steel blue
    "accent": "#2f6690",  # Darker blue
    "light": "#f5f9ff",  # Very light blue
    "dark": "#16425b",  # Dark blue
    "success": "#4c9f70",  # Teal
    "danger": "#d64045",  # Coral red
    "warning": "#ff7d00",  # Orange
    "text": "#2f3640"  # Dark gray
}

# 字体
font_large = ("Segoe UI", 18, "bold")
font_medium = ("Segoe UI", 12)
font_small = ("Segoe UI", 10)

# 获取当前脚本所在目录
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")
os.makedirs(DATA_DIR, exist_ok=True)

# ====================== 数据部分 ======================
# 文件路径 (使用相对路径)
SIDEBAR_FILE_PATHS = {
    "Physicochemical": os.path.join(DATA_DIR, "Physcicochemical.xlsx"),
    "Compound Classes": os.path.join(DATA_DIR, "compound classifire.xlsx"),
    "AOP Analysis": os.path.join(DATA_DIR, "AOP Analysis.xlsx"),
    "Targets & Mechanisms": os.path.join(DATA_DIR, "Targets Mechanisms.xlsx"),
    "ADME Properties": os.path.join(DATA_DIR, "ADME Properties.xlsx"),
    "Structural Alerts": os.path.join(DATA_DIR, "Structural Alerts.xlsx"),
     "Suspect List Screening": os.path.join(DATA_DIR, "suspect_screening_tool.py"),
    "Spectral Library": DATA_DIR,
    "Suspect List Screening": "suspect_screening_tool"
}

SIDEBAR_HELP_PATHS = {
    "Physicochemical": os.path.join(DATA_DIR, "Physcicochemical.Help.docx"),
    "Compound Classes": os.path.join(DATA_DIR, "compound classifire.Help.docx"),
    "AOP Analysis": os.path.join(DATA_DIR, "AOP Analysis.Help.docx"),
    "Targets & Mechanisms": os.path.join(DATA_DIR, "Targets Mechanisms.Help.docx"),
    "ADME Properties": os.path.join(DATA_DIR, "ADME Properties.Help.docx"),
    "Structural Alerts": os.path.join(DATA_DIR, "Structural Alerts.Help.docx"),
    "Spectral Library": os.path.join(DATA_DIR, "Spectral Library.Help.docx"),
    "Suspect List Screening": os.path.join(DATA_DIR, "Suspect List Screening.Help.docx"),
    "main": os.path.join(DATA_DIR, "main.Help.docx")
}
MAIN_DATA_PATH = os.path.join(DATA_DIR, "all.xlsx")
# 光谱库数据
spectral_libraries = {
    "All Spectra": {
        "date": "2025/8/20 19:09",
        "size": "195,033 KB",
        "count": "81,585",
        "download_file": "All Spectra.msp"
    },
    "LC-MS Spectra": {
        "date": "2025/8/20 20:08",
        "size": "176,587 KB",
        "count": "70,135",
        "download_file": "LC-MS Spectra.msp"
    },
    "LC-ESI-QTOF": {
        "date": "2025/8/20 20:21",
        "size": "46,996 KB",
        "count": "22,758",
        "download_file": "LC-ESI-QTOF.msp"
    },
    "LC-ESI-IIFT": {
        "date": "2025/8/20 20:21",
        "size": "33,897 KB",
        "count": "17,004",
        "download_file": "LC-ESI-IIFT.msp"
    },
    "LC-ESI-QFT": {
        "date": "2025/8/20 20:21",
        "size": "32,437 KB",
        "count": "12,851",
        "download_file": "LC-ESI-QFT.msp"
    },
    "LC-MSMS Positive Mode": {
        "date": "2025/8/20 20:30",
        "size": "140,946 KB",
        "count": "53,571",
        "download_file": "LC-MSMS Positive Mode.msp"
    },
    "LC-MSMS Negative Mode": {
        "date": "2025/8/20 20:30",
        "size": "19,120 KB",
        "count": "35,641",
        "download_file": "LC-MSMS Negative Mode.msp"
    },
    "GC-MS Spectra": {
        "date": "2025/8/20 20:15",
        "size": "12,198 KB",
        "count": "17,111",
        "download_file": "GC-MS Spectra.msp"
    },
    "GC-EI Spectra": {
        "date": "2025/8/20 20:40",
        "size": "13,958 KB",
        "count": "9,759",
        "download_file": "GC-EI Spectra.msp"
    },
    "GC-CI Spectra": {
        "date": "2025/8/20 20:40",
        "size": "1,294 KB",
        "count": "1,442",
        "download_file": "GC-CI Spectra.msp"
    }
}

# 存储每个对话框的数据和文件路径的字典
dialog_data_cache = {}
# ====================== 功能函数部分 ======================
def load_compound_data():
    """加载主界面化合物数据"""
    try:
        if os.path.exists(MAIN_DATA_PATH):
            df = pd.read_excel(MAIN_DATA_PATH, engine='openpyxl')
            messagebox.showinfo("Success", f"Successfully loaded data: {MAIN_DATA_PATH}\nTotal records: {len(df)}")
            return df
        else:
            messagebox.showwarning("File Not Found", f"File not found: {MAIN_DATA_PATH}\nUsing sample data instead.")

            sample_columns = [
                "CAS RN", "Chemical name", "DTXSID", "InChiKey", "IUPAC name", "SMILES", "InChi",
                "Molecular formula", "Average mass", "Monoisotopic mass", "Log Kaw", "Log Koa",
                "Log Kow", "Half-life in air (hours)", "Half-life in soil (hours)",
                "Half-life in water (hours)", "log BCF", "T1/2", "Category", "Kingdom",
                "Superclass", "Class", "Subclass", "F(20%)", "F(30%)", "PPB", "CYP1A2-inh",
                "CYP1A2-sub", "CYP2C19-inh", "CYP2C19-sub", "CYP2C9-inh", "CYP2C9-sub",
                "CYP2D6-inh", "CYP2D6-sub", "CYP3A4-inh", "CYP3A4-sub", "CL"
            ]
            sample_data = [
                ["301-04-2", "Lead acetate", "DTXSID1020001", "XYZ123", "Lead(II) acetate",
                 "CC(=O)[O-].CC(=O)[O-].[Pb+2]", "InChI=1S/2C2H4O2.Pb/c2*1-2(3)4;/h2*1H3,(H,3,4);/q;;+2/p-2",
                 "C4H6O4Pb", "325.29", "325.0", "1.5", "4.21", "0.73", "240", "720", "240",
                 "2.1", "10 days", "Inorganic", "Inorganic compounds", "Metals", "Heavy metals",
                 "Lead compounds", "0.2", "0.3", "0.05", "Yes", "No", "No", "No", "No", "No",
                 "No", "No", "No", "No", "5.2"]
            ]
            return pd.DataFrame(sample_data, columns=sample_columns)
    except Exception as e:
        messagebox.showerror("Error", f"Failed to load compound data: {str(e)}")
        return None


def setup_treeview_columns(tree, columns):
    """设置树形视图的列"""
    tree['columns'] = columns
    for col in columns:
        tree.heading(col, text=col)
        if len(col) > 20:
            tree.column(col, width=200, anchor="w", minwidth=100)
        elif len(col) > 10:
            tree.column(col, width=150, anchor="w", minwidth=80)
        else:
            tree.column(col, width=100, anchor="w", minwidth=60)


def load_all_data_to_tree(tree, data, result_count):
    """将所有数据加载到树形视图"""
    for item in tree.get_children():
        tree.delete(item)

    for _, row in data.iterrows():
        values = []
        for col in data.columns:
            value = row[col]
            if pd.isna(value):
                values.append("N/A")
            else:
                values.append(str(value))
        tree.insert("", tk.END, values=values)
    result_count.config(text=f"Total: {len(data)} records")


def search_compounds(search_entry, tree, data, columns, result_count):
    """搜索化合物"""
    search_term = search_entry.get().strip()

    if not search_term:
        messagebox.showwarning("Input Error", "Please enter a search term (CAS No. or compound name)")
        return

    for item in tree.get_children():
        tree.delete(item)

    if data is not None and columns is not None:
        filtered_data = data[
            (data['CAS RN'].astype(str).str.contains(search_term, case=False, na=False)) |
            (data['Chemical name'].astype(str).str.contains(search_term, case=False, na=False))
            ]

        for _, row in filtered_data.iterrows():
            values = []
            for col in columns:
                if col in data.columns:
                    value = row[col]
                    if pd.isna(value):
                        values.append("N/A")
                    else:
                        values.append(str(value))
                else:
                    values.append("N/A")
            tree.insert("", tk.END, values=values)

        result_count.config(text=f"Found {len(filtered_data)} records")
    else:
        messagebox.showerror("Error", "Compound data not loaded properly")


def download_results(tree, columns):
    """下载搜索结果"""
    items = tree.get_children()
    data = []

    for item in items:
        values = tree.item(item, 'values')
        data.append(values)

    if not data:
        messagebox.showwarning("No Data", "No data to download")
        return

    df = pd.DataFrame(data, columns=columns)

    file_path = filedialog.asksaveasfilename(
        defaultextension=".csv",
        filetypes=[("CSV files", "*.csv"), ("Excel files", "*.xlsx"), ("All files", "*.*")],
        title="Save Search Results"
    )

    if file_path:
        try:
            if file_path.endswith('.xlsx'):
                df.to_excel(file_path, index=False, engine='openpyxl')
            else:
                df.to_csv(file_path, index=False)
            messagebox.showinfo("Success", "Results downloaded successfully!")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save file: {str(e)}")


def batch_search(compound_data, tree, result_count):
    """批量搜索"""
    file_path = filedialog.askopenfilename(
        filetypes=[("Text files", "*.txt")],
        title="选择包含CAS号的文件"
    )

    if not file_path:
        return

    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            cas_numbers = [line.strip() for line in f if line.strip()]

        if not cas_numbers:
            messagebox.showinfo("结果", "文件中未找到有效的CAS号")
            return

        if compound_data is None or compound_data.empty:
            messagebox.showwarning("警告", "化合物数据库未加载")
            return

        matches = compound_data[compound_data['CAS RN'].astype(str).str.strip().isin(cas_numbers)]

        for item in tree.get_children():
            tree.delete(item)

        if not matches.empty:
            for _, row in matches.iterrows():
                values = []
                for col in compound_data.columns:
                    value = row[col]
                    if pd.isna(value):
                        values.append("N/A")
                    else:
                        values.append(str(value))
                tree.insert("", tk.END, values=values)

            result_count.config(text=f"找到 {len(matches)}/{len(cas_numbers)} 个匹配项")
            messagebox.showinfo("成功", f"批量搜索完成！找到 {len(matches)} 个匹配化合物")
        else:
            result_count.config(text="找到 0 个匹配项")
            messagebox.showinfo("结果", "未找到匹配的化合物")
    except Exception as e:
        messagebox.showerror("错误", f"批量搜索失败: {str(e)}")


def clear_search(tree, compound_data, result_count, search_entry):
    """清除搜索结果，恢复到初始状态"""
    search_entry.delete(0, tk.END)

    for item in tree.get_children():
        tree.delete(item)

    if compound_data is not None:
        for _, row in compound_data.iterrows():
            values = []
            for col in compound_data.columns:
                value = row[col]
                if pd.isna(value):
                    values.append("N/A")
                else:
                    values.append(str(value))
            tree.insert("", tk.END, values=values)
        result_count.config(text=f"Total: {len(compound_data)} records")

    messagebox.showinfo("Cleared", "搜索已清除，恢复到初始状态")


def upload_compound_data(tree, compound_data, result_count):
    """上传化合物数据到主数据库"""
    if compound_data is None:
        messagebox.showerror("Error", "Compound data not loaded")
        return

    # 创建上传对话框
    upload_dialog = tk.Toplevel()
    upload_dialog.title("Upload Compound Information")
    upload_dialog.geometry("600x500")
    upload_dialog.configure(bg=colors["light"])

    # 创建滚动框架
    canvas = tk.Canvas(upload_dialog, bg=colors["light"])
    scrollbar = ttk.Scrollbar(upload_dialog, orient="vertical", command=canvas.yview)
    scrollable_frame = tk.Frame(canvas, bg=colors["light"])

    scrollable_frame.bind(
        "<Configure>",
        lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
    )

    canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)

    # 标题
    title_label = tk.Label(scrollable_frame,
                           text="Add New Compound Record",
                           font=("Segoe UI", 14, "bold"),
                           bg=colors["light"],
                           fg=colors["dark"])
    title_label.pack(pady=10)

    # 存储输入框的字典
    entries = {}

    # 为每一列创建输入框
    for i, col in enumerate(compound_data.columns):
        frame = tk.Frame(scrollable_frame, bg=colors["light"])
        frame.pack(fill="x", padx=20, pady=5)

        label = tk.Label(frame, text=f"{col}:", font=font_medium,
                         bg=colors["light"], fg=colors["dark"], width=20, anchor="w")
        label.pack(side="left")

        entry = tk.Entry(frame, font=font_medium, width=40)
        entry.pack(side="left", padx=10)
        entries[col] = entry

    def submit_new_compound():
        try:
            nonlocal compound_data
            # 收集输入的数据
            new_row = {}
            for col, entry in entries.items():
                value = entry.get().strip()
                new_row[col] = value if value else None

            # 添加到DataFrame
            compound_data = pd.concat([compound_data, pd.DataFrame([new_row])], ignore_index=True)

            # 保存到文件
            if os.path.exists(MAIN_DATA_PATH):
                compound_data.to_excel(MAIN_DATA_PATH, index=False, engine='openpyxl')

            # 更新树形视图
            for item in tree.get_children():
                tree.delete(item)

            for _, row in compound_data.iterrows():
                values = []
                for col in compound_data.columns:
                    value = row[col]
                    if pd.isna(value):
                        values.append("N/A")
                    else:
                        values.append(str(value))
                tree.insert("", tk.END, values=values)

            result_count.config(text=f"Total: {len(compound_data)} records")
            messagebox.showinfo("Success", "New compound record added successfully!")
            upload_dialog.destroy()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to add record: {str(e)}")

    # 按钮框架
    button_frame_dialog = tk.Frame(scrollable_frame, bg=colors["light"])
    button_frame_dialog.pack(fill="x", padx=20, pady=20)

    submit_btn = tk.Button(button_frame_dialog, text="Submit", command=submit_new_compound,
                           bg=colors["success"], fg="white", font=font_medium,
                           padx=20, pady=5)
    submit_btn.pack(side="left", padx=10)

    cancel_btn = tk.Button(button_frame_dialog, text="Cancel", command=upload_dialog.destroy,
                           bg=colors["danger"], fg="white", font=font_medium,
                           padx=20, pady=5)
    cancel_btn.pack(side="left", padx=10)

    canvas.pack(side="left", fill="both", expand=True, padx=10, pady=10)
    scrollbar.pack(side="right", fill="y")


def edit_selected_compound(tree, compound_data, result_count):
    """编辑选中的化合物"""
    if compound_data is None:
        messagebox.showerror("Error", "Compound data not loaded")
        return

    selected = tree.selection()
    if not selected:
        messagebox.showwarning("Warning", "Please select a row to edit")
        return

    # 获取选中的行数据
    item = selected[0]
    values = tree.item(item, 'values')

    # 创建编辑对话框
    edit_dialog = tk.Toplevel()
    edit_dialog.title("Edit Compound Information")
    edit_dialog.geometry("600x500")
    edit_dialog.configure(bg=colors["light"])

    # 创建滚动框架
    canvas = tk.Canvas(edit_dialog, bg=colors["light"])
    scrollbar = ttk.Scrollbar(edit_dialog, orient="vertical", command=canvas.yview)
    scrollable_frame = tk.Frame(canvas, bg=colors["light"])

    scrollable_frame.bind(
        "<Configure>",
        lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
    )

    canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)

    # 标题
    title_label = tk.Label(scrollable_frame,
                           text="Edit Compound Record",
                           font=("Segoe UI", 14, "bold"),
                           bg=colors["light"],
                           fg=colors["dark"])
    title_label.pack(pady=10)

    # 存储输入框的字典
    entries = {}

    # 为每一列创建输入框
    for i, col in enumerate(compound_data.columns):
        frame = tk.Frame(scrollable_frame, bg=colors["light"])
        frame.pack(fill="x", padx=20, pady=5)

        label = tk.Label(frame, text=f"{col}:", font=font_medium,
                         bg=colors["light"], fg=colors["dark"], width=20, anchor="w")
        label.pack(side="left")

        entry = tk.Entry(frame, font=font_medium, width=40)
        entry.pack(side="left", padx=10)

        # 填充现有数据
        if i < len(values) and values[i] != "N/A":
            entry.insert(0, values[i])
        entries[col] = entry

    def save_edit():
        try:
            nonlocal compound_data
            # 获取当前选中行的索引
            index = tree.index(item)

            # 收集编辑后的数据
            edited_row = {}
            for col, entry in entries.items():
                value = entry.get().strip()
                edited_row[col] = value if value else None

            # 更新DataFrame
            compound_data.loc[index] = edited_row

            # 保存到文件
            if os.path.exists(MAIN_DATA_PATH):
                compound_data.to_excel(MAIN_DATA_PATH, index=False, engine='openpyxl')

            # 更新树形视图
            tree.item(item, values=list(edited_row.values()))
            messagebox.showinfo("Success", "Record updated successfully!")
            edit_dialog.destroy()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to update record: {str(e)}")

    # 按钮框架
    button_frame_dialog = tk.Frame(scrollable_frame, bg=colors["light"])
    button_frame_dialog.pack(fill="x", padx=20, pady=20)

    save_btn = tk.Button(button_frame_dialog, text="Save", command=save_edit,
                         bg=colors["success"], fg="white", font=font_medium,
                         padx=20, pady=5)
    save_btn.pack(side="left", padx=10)

    cancel_btn = tk.Button(button_frame_dialog, text="Cancel", command=edit_dialog.destroy,
                           bg=colors["danger"], fg="white", font=font_medium,
                           padx=20, pady=5)
    cancel_btn.pack(side="left", padx=10)

    canvas.pack(side="left", fill="both", expand=True, padx=10, pady=10)
    scrollbar.pack(side="right", fill="y")


def delete_selected_compound(tree, compound_data, result_count):
    """删除选中的化合物"""
    if compound_data is None:
        messagebox.showerror("Error", "Compound data not loaded")
        return

    selected = tree.selection()
    if not selected:
        messagebox.showwarning("Warning", "Please select a row to delete")
        return

    if messagebox.askyesno("Confirm Delete", "Are you sure you want to delete this record?"):
        try:
            # 获取选中行的索引
            indices = [tree.index(item) for item in selected]

            # 从DataFrame中删除行（从后往前删，避免索引变化）
            for idx in sorted(indices, reverse=True):
                compound_data = compound_data.drop(compound_data.index[idx]).reset_index(drop=True)

            # 保存到文件
            if os.path.exists(MAIN_DATA_PATH):
                compound_data.to_excel(MAIN_DATA_PATH, index=False, engine='openpyxl')

            # 更新树形视图
            for item in selected:
                tree.delete(item)
            result_count.config(text=f"Total: {len(compound_data)} records")
            messagebox.showinfo("Success", "Record deleted successfully!")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to delete record: {str(e)}")


def load_sidebar_file(file_key):
    """加载侧边栏文件"""
    try:
        if file_key == "Suspect List Screening":
            run_suspect_list_screening()
            return
        elif file_key == "Spectral Library":
            show_spectral_library()
            return

        file_path = SIDEBAR_FILE_PATHS[file_key]
        if os.path.isdir(file_path):
            show_directory_contents(file_path, file_key)
        elif os.path.exists(file_path):
            if file_path.endswith(('.xlsx', '.xls')):
                # 读取数据
                df = pd.read_excel(file_path, engine='openpyxl')
                # 缓存数据和文件路径
                dialog_data_cache[file_key] = {
                    'df': df.copy(),
                    'file_path': file_path
                }
                show_data_in_window(df, file_key, file_path)
            else:
                show_file_content(file_path, file_key)
        else:
            # 如果文件不存在，创建新文件
            if file_key in ["Physicochemical", "Compound Classes", "AOP Analysis",
                            "Targets & Mechanisms", "ADME Properties", "Structural Alerts"]:
                # 创建默认的DataFrame
                default_columns = get_default_columns_for_tool(file_key)
                df = pd.DataFrame(columns=default_columns)
                # 保存到文件
                df.to_excel(file_path, index=False, engine='openpyxl')
                # 缓存数据和文件路径
                dialog_data_cache[file_key] = {
                    'df': df.copy(),
                    'file_path': file_path
                }
                show_data_in_window(df, file_key, file_path)
                messagebox.showinfo("Info", f"Created new file: {file_path}")
            else:
                messagebox.showwarning("File Not Found", f"File not found: {file_path}")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to load file: {str(e)}")


def get_default_columns_for_tool(tool_name):
    """为每个工具返回默认的列名"""
    default_columns = {
        "Physicochemical": ["CAS RN", "Chemical Name", "Molecular Formula", "Molecular Weight",
                            "Log P", "Log D", "pKa", "Solubility", "Melting Point", "Boiling Point"],
        "Compound Classes": ["CAS RN", "Chemical Name", "Class", "Subclass", "Category",
                             "Kingdom", "Superclass", "Description"],
        "AOP Analysis": ["Stressor", "KE ID", "KE Name", "MIE", "AO", "Organism",
                         "Evidence", "Reference"],
        "Targets & Mechanisms": ["CAS RN", "Chemical Name", "Target", "Mechanism",
                                 "Action Type", "Organism", "Reference"],
        "ADME Properties": ["CAS RN", "Chemical Name", "Absorption", "Distribution",
                            "Metabolism", "Excretion", "Bioavailability", "PPB", "CL"],
        "Structural Alerts": ["CAS RN", "Chemical Name", "Alert Name", "SMARTS",
                              "Toxicity Type", "Severity", "Reference"]
    }
    return default_columns.get(tool_name, ["CAS RN", "Chemical Name", "Data"])


def save_data_to_file(file_key, df):
    """保存数据到文件"""
    try:
        file_path = SIDEBAR_FILE_PATHS.get(file_key)
        if file_path and file_path.endswith(('.xlsx', '.xls')):
            df.to_excel(file_path, index=False, engine='openpyxl')
            # 更新缓存
            if file_key in dialog_data_cache:
                dialog_data_cache[file_key]['df'] = df.copy()
            return True
    except Exception as e:
        messagebox.showerror("Error", f"Failed to save data: {str(e)}")
        return False


def add_new_row_to_data(file_key, df, tree, result_count):
    """添加新行到数据中"""
    # 创建输入对话框
    dialog = tk.Toplevel()
    dialog.title(f"Add New Record - {file_key}")
    dialog.geometry("600x500")
    dialog.configure(bg=colors["light"])

    # 创建滚动框架
    canvas = tk.Canvas(dialog, bg=colors["light"])
    scrollbar = ttk.Scrollbar(dialog, orient="vertical", command=canvas.yview)
    scrollable_frame = tk.Frame(canvas, bg=colors["light"])

    scrollable_frame.bind(
        "<Configure>",
        lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
    )

    canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)

    # 标题
    title_label = tk.Label(scrollable_frame,
                           text=f"Add New {file_key} Record",
                           font=("Segoe UI", 14, "bold"),
                           bg=colors["light"],
                           fg=colors["dark"])
    title_label.pack(pady=10)

    # 存储输入框的字典
    entries = {}

    # 为每一列创建输入框
    for i, col in enumerate(df.columns):
        frame = tk.Frame(scrollable_frame, bg=colors["light"])
        frame.pack(fill="x", padx=20, pady=5)

        label = tk.Label(frame, text=f"{col}:", font=font_medium,
                         bg=colors["light"], fg=colors["dark"], width=20, anchor="w")
        label.pack(side="left")

        entry = tk.Entry(frame, font=font_medium, width=40)
        entry.pack(side="left", padx=10)
        entries[col] = entry

    def submit_new_row():
        try:
            # 收集输入的数据
            new_row = {}
            for col, entry in entries.items():
                value = entry.get().strip()
                new_row[col] = value if value else None

            # 添加到DataFrame
            new_df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)

            # 保存到文件
            if save_data_to_file(file_key, new_df):
                # 更新对话框中的数据
                dialog_data_cache[file_key]['df'] = new_df

                # 更新树形视图
                for item in tree.get_children():
                    tree.delete(item)

                for _, row in new_df.iterrows():
                    values = []
                    for col in new_df.columns:
                        value = row[col]
                        if pd.isna(value):
                            values.append("N/A")
                        else:
                            values.append(str(value))
                    tree.insert("", tk.END, values=values)

                result_count.config(text=f"Total: {len(new_df)} records")
                messagebox.showinfo("Success", "New record added successfully!")
                dialog.destroy()
            else:
                messagebox.showerror("Error", "Failed to save data")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to add record: {str(e)}")

    # 按钮框架
    button_frame_dialog = tk.Frame(scrollable_frame, bg=colors["light"])
    button_frame_dialog.pack(fill="x", padx=20, pady=20)

    submit_btn = tk.Button(button_frame_dialog, text="Submit", command=submit_new_row,
                           bg=colors["success"], fg="white", font=font_medium,
                           padx=20, pady=5)
    submit_btn.pack(side="left", padx=10)

    cancel_btn = tk.Button(button_frame_dialog, text="Cancel", command=dialog.destroy,
                           bg=colors["danger"], fg="white", font=font_medium,
                           padx=20, pady=5)
    cancel_btn.pack(side="left", padx=10)

    canvas.pack(side="left", fill="both", expand=True, padx=10, pady=10)
    scrollbar.pack(side="right", fill="y")


def edit_selected_row(file_key, df, tree, result_count):
    """编辑选中的行"""
    selected = tree.selection()
    if not selected:
        messagebox.showwarning("Warning", "Please select a row to edit")
        return

    # 获取选中的行数据
    item = selected[0]
    values = tree.item(item, 'values')

    # 创建编辑对话框
    dialog = tk.Toplevel()
    dialog.title(f"Edit Record - {file_key}")
    dialog.geometry("600x500")
    dialog.configure(bg=colors["light"])

    # 创建滚动框架
    canvas = tk.Canvas(dialog, bg=colors["light"])
    scrollbar = ttk.Scrollbar(dialog, orient="vertical", command=canvas.yview)
    scrollable_frame = tk.Frame(canvas, bg=colors["light"])

    scrollable_frame.bind(
        "<Configure>",
        lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
    )

    canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)

    # 标题
    title_label = tk.Label(scrollable_frame,
                           text=f"Edit {file_key} Record",
                           font=("Segoe UI", 14, "bold"),
                           bg=colors["light"],
                           fg=colors["dark"])
    title_label.pack(pady=10)

    # 存储输入框的字典
    entries = {}

    # 为每一列创建输入框，并填充现有数据
    for i, col in enumerate(df.columns):
        frame = tk.Frame(scrollable_frame, bg=colors["light"])
        frame.pack(fill="x", padx=20, pady=5)

        label = tk.Label(frame, text=f"{col}:", font=font_medium,
                         bg=colors["light"], fg=colors["dark"], width=20, anchor="w")
        label.pack(side="left")

        entry = tk.Entry(frame, font=font_medium, width=40)
        entry.pack(side="left", padx=10)

        # 填充现有数据
        if i < len(values) and values[i] != "N/A":
            entry.insert(0, values[i])
        entries[col] = entry

    def save_edit():
        try:
            # 获取当前选中行的索引
            index = tree.index(item)

            # 收集编辑后的数据
            edited_row = {}
            for col, entry in entries.items():
                value = entry.get().strip()
                edited_row[col] = value if value else None

            # 更新DataFrame
            df.loc[index] = edited_row

            # 保存到文件
            if save_data_to_file(file_key, df):
                # 更新树形视图
                tree.item(item, values=list(edited_row.values()))
                messagebox.showinfo("Success", "Record updated successfully!")
                dialog.destroy()
            else:
                messagebox.showerror("Error", "Failed to save data")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to update record: {str(e)}")

    # 按钮框架
    button_frame_dialog = tk.Frame(scrollable_frame, bg=colors["light"])
    button_frame_dialog.pack(fill="x", padx=20, pady=20)

    save_btn = tk.Button(button_frame_dialog, text="Save", command=save_edit,
                         bg=colors["success"], fg="white", font=font_medium,
                         padx=20, pady=5)
    save_btn.pack(side="left", padx=10)

    cancel_btn = tk.Button(button_frame_dialog, text="Cancel", command=dialog.destroy,
                           bg=colors["danger"], fg="white", font=font_medium,
                           padx=20, pady=5)
    cancel_btn.pack(side="left", padx=10)

    canvas.pack(side="left", fill="both", expand=True, padx=10, pady=10)
    scrollbar.pack(side="right", fill="y")


def delete_selected_row(file_key, df, tree, result_count):
    """删除选中的行"""
    selected = tree.selection()
    if not selected:
        messagebox.showwarning("Warning", "Please select a row to delete")
        return

    if messagebox.askyesno("Confirm Delete", "Are you sure you want to delete this record?"):
        try:
            # 获取选中行的索引
            indices = [tree.index(item) for item in selected]

            # 从DataFrame中删除行（从后往前删，避免索引变化）
            for idx in sorted(indices, reverse=True):
                df = df.drop(df.index[idx]).reset_index(drop=True)

            # 保存到文件
            if save_data_to_file(file_key, df):
                # 更新树形视图
                for item in selected:
                    tree.delete(item)
                result_count.config(text=f"Total: {len(df)} records")
                messagebox.showinfo("Success", "Record deleted successfully!")
            else:
                messagebox.showerror("Error", "Failed to save data")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to delete record: {str(e)}")


def show_data_in_window(df, title, file_path=None):
    """显示数据窗口，添加编辑功能"""
    window = tk.Toplevel()
    window.title(f"{title} - Data View")
    window.geometry("1000x800")
    window.configure(bg=colors["light"])

    # 获取文件key
    file_key = title

    # Search frame
    search_frame = tk.Frame(window, bg=colors["light"], padx=10, pady=10)
    search_frame.pack(fill="x")

    # 根据标题确定搜索字段
    if title == "AOP Analysis":
        search_label_text = "Search by Stressor:"
        search_placeholder = "Enter stressor name"
        search_column = "Stressor"
    else:
        search_label_text = "Search by CAS No:"
        search_placeholder = "Enter CAS number"
        search_column = "CAS RN"

    search_label = tk.Label(search_frame, text=search_label_text, font=font_medium,
                            bg=colors["light"], fg=colors["dark"])
    search_label.pack(side="left", padx=(0, 10))

    search_entry = tk.Entry(search_frame, width=30, font=font_medium)
    search_entry.insert(0, search_placeholder)
    search_entry.pack(side="left", padx=(0, 10))

    search_button = tk.Button(search_frame, text="🔍 Search",
                              command=lambda: perform_sidebar_search(tree_widget, df, search_entry.get().strip(),
                                                                     search_column, search_placeholder),
                              bg=colors["accent"], fg="white", font=font_medium)
    search_button.pack(side="left", padx=(0, 5))

    clear_button = tk.Button(search_frame, text="🔄 Clear",
                             command=lambda: load_all_sidebar_data(tree_widget, df),
                             bg=colors["warning"], fg="white", font=font_medium)
    clear_button.pack(side="left", padx=5)

    batch_search_button = tk.Button(search_frame, text="🧪 Batch Search",
                                    command=lambda: batch_search_sidebar(tree_widget, df, result_count, search_column),
                                    bg=colors["success"], fg="white", font=font_medium)
    batch_search_button.pack(side="left", padx=5)

    # 添加上传信息按钮
    upload_button = tk.Button(search_frame, text="📤 Upload Information",
                              command=lambda: add_new_row_to_data(file_key, df, tree_widget, result_count),
                              bg=colors["primary"], fg="white", font=font_medium)
    upload_button.pack(side="left", padx=5)

    # 添加编辑按钮
    edit_button = tk.Button(search_frame, text="✏️ Edit Selected",
                            command=lambda: edit_selected_row(file_key, df, tree_widget, result_count),
                            bg=colors["warning"], fg="white", font=font_medium)
    edit_button.pack(side="left", padx=5)

    # 添加删除按钮
    delete_button = tk.Button(search_frame, text="🗑️ Delete Selected",
                              command=lambda: delete_selected_row(file_key, df, tree_widget, result_count),
                              bg=colors["danger"], fg="white", font=font_medium)
    delete_button.pack(side="left", padx=5)

    result_count = tk.Label(search_frame, text=f"Total: {len(df)} records",
                            bg=colors["light"], fg=colors["dark"], font=font_medium)
    result_count.pack(side="right")

    # Main frame for treeview
    frame = tk.Frame(window, bg=colors["light"])
    frame.pack(fill="both", expand=True, padx=10, pady=10)

    # Create treeview with scrollbars
    tree_frame = tk.Frame(frame, bg=colors["light"])
    tree_frame.pack(fill="both", expand=True)

    h_scroll = ttk.Scrollbar(tree_frame, orient="horizontal")
    h_scroll.pack(side="bottom", fill="x")

    v_scroll = ttk.Scrollbar(tree_frame, orient="vertical")
    v_scroll.pack(side="right", fill="y")

    tree_widget = ttk.Treeview(tree_frame, columns=list(df.columns), show="headings", height=20,
                               xscrollcommand=h_scroll.set, yscrollcommand=v_scroll.set)

    h_scroll.config(command=tree_widget.xview)
    v_scroll.config(command=tree_widget.yview)

    for col in df.columns:
        tree_widget.heading(col, text=col)
        tree_widget.column(col, width=120, anchor="w", minwidth=80)

    tree_widget.pack(side="left", fill="both", expand=True)

    def load_all_sidebar_data(tree, data):
        for item in tree.get_children():
            tree.delete(item)
        for _, row in data.iterrows():
            values = []
            for col in data.columns:
                value = row[col]
                if pd.isna(value):
                    values.append("N/A")
                else:
                    values.append(str(value))
            tree.insert("", tk.END, values=values)
        result_count.config(text=f"Total: {len(data)} records")

    load_all_sidebar_data(tree_widget, df)

    def perform_sidebar_search(tree, data, search_term, search_column="CAS RN", placeholder=""):
        if not search_term or search_term == placeholder:
            messagebox.showwarning("Input Error", f"Please enter a {search_column.lower()} to search")
            return
        count = search_sidebar_data(tree, data, search_term, search_column)
        result_count.config(text=f"Found: {count} records")
        if count == 0:
            messagebox.showinfo("Search Result", f"No records found with the specified {search_column}")

    def search_sidebar_data(tree, data, search_term, search_column="CAS RN"):
        for item in tree.get_children():
            tree.delete(item)
        if data is not None:
            if search_column in data.columns:
                filtered_data = data[data[search_column].astype(str).str.contains(search_term, case=False, na=False)]
            else:
                mask = pd.Series(False, index=data.index)
                for col in data.columns:
                    mask = mask | data[col].astype(str).str.contains(search_term, case=False, na=False)
                filtered_data = data[mask]
            for _, row in filtered_data.iterrows():
                values = []
                for col in data.columns:
                    value = row[col]
                    if pd.isna(value):
                        values.append("N/A")
                    else:
                        values.append(str(value))
                tree.insert("", tk.END, values=values)
            return len(filtered_data)
        return 0

    def download_sidebar_data():
        items = tree_widget.get_children()
        if not items:
            messagebox.showwarning("No Data", "No data to download")
            return
        data = []
        for item in items:
            values = tree_widget.item(item, 'values')
            data.append(values)
        download_df = pd.DataFrame(data, columns=df.columns)
        file_path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv"), ("Excel files", "*.xlsx"), ("All files", "*.*")],
            title=f"Save {title} Data"
        )
        if file_path:
            try:
                if file_path.endswith('.xlsx'):
                    download_df.to_excel(file_path, index=False, engine='openpyxl')
                else:
                    download_df.to_csv(file_path, index=False)
                messagebox.showinfo("Success", "Data downloaded successfully!")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to save file: {str(e)}")

    download_btn = tk.Button(window, text="⬇️ Download Results",
                             command=download_sidebar_data,
                             bg=colors["success"], fg="white", font=font_medium,
                             padx=20, pady=5)
    download_btn.pack(pady=10)


def batch_search_sidebar(tree, data, result_count, search_column="CAS RN"):
    """批量搜索侧边栏数据"""
    file_path = filedialog.askopenfilename(
        filetypes=[("Text files", "*.txt")],
        title="选择包含搜索项的文件"
    )
    if not file_path:
        return
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            search_items = [line.strip() for line in f if line.strip()]

        if not search_items:
            messagebox.showinfo("结果", "文件中未找到有效的搜索项")
            return

        if data is None or data.empty:
            messagebox.showwarning("警告", "数据库未加载")
            return

        if search_column in data.columns:
            matches = data[data[search_column].astype(str).str.strip().isin(search_items)]
        else:
            mask = pd.Series(False, index=data.index)
            for col in data.columns:
                mask = mask | data[col].astype(str).str.strip().isin(search_items)
            matches = data[mask]

        for item in tree.get_children():
            tree.delete(item)

        if not matches.empty:
            for _, row in matches.iterrows():
                values = []
                for col in data.columns:
                    value = row[col]
                    if pd.isna(value):
                        values.append("N/A")
                    else:
                        values.append(str(value))
                tree.insert("", tk.END, values=values)

            result_count.config(text=f"找到 {len(matches)}/{len(search_items)} 个匹配项")
            messagebox.showinfo("成功", f"批量搜索完成！找到 {len(matches)} 个匹配项")
        else:
            result_count.config(text="找到 0 个匹配项")
            messagebox.showinfo("结果", "未找到匹配的项")
    except Exception as e:
        messagebox.showerror("错误", f"批量搜索失败: {str(e)}")


def download_file(file_path):
    """下载指定路径的文件到本地"""
    if not os.path.exists(file_path):
        messagebox.showerror("File Not Found", f"File not found: {file_path}")
        return
    save_path = filedialog.asksaveasfilename(
        defaultextension=os.path.splitext(file_path)[1],
        initialfile=os.path.basename(file_path),
        filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        title="Save Target Information"
    )
    if save_path:
        try:
            with open(file_path, 'rb') as src_file:
                with open(save_path, 'wb') as dest_file:
                    dest_file.write(src_file.read())
            messagebox.showinfo("Success", f"Target Information saved to:\n{save_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save file: {str(e)}")


def show_directory_contents(directory_path, title):
    """显示目录内容"""
    window = tk.Toplevel()
    window.title(f"{title} - Directory Contents")
    window.geometry("800x400")
    frame = tk.Frame(window)
    frame.pack(fill="both", expand=True, padx=10, pady=10)
    tree = ttk.Treeview(frame, columns=("Name", "Type", "Size"), show="headings", height=15)
    tree.heading("Name", text="Name")
    tree.heading("Type", text="Type")
    tree.heading("Size", text="Size")
    tree.column("Name", width=300)
    tree.column("Type", width=150)
    tree.column("Size", width=100)
    try:
        files = os.listdir(directory_path)
        for file in files:
            full_path = os.path.join(directory_path, file)
            if os.path.isfile(full_path):
                file_type = "File"
                size = os.path.getsize(full_path)
                size_str = f"{size / 1024:.1f} KB" if size < 1024 * 1024 else f"{size / (1024 * 1024):.1f} MB"
            else:
                file_type = "Directory"
                size_str = "-"
            tree.insert("", tk.END, values=(file, file_type, size_str))
    except Exception as e:
        tree.insert("", tk.END, values=(f"Error reading directory: {str(e)}", "Error", "Error"))
    v_scroll = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
    tree.configure(yscrollcommand=v_scroll.set)
    tree.pack(side="left", fill="both", expand=True)
    v_scroll.pack(side="right", fill="y")
    info_label = tk.Label(window, text=f"Directory: {directory_path}")
    info_label.pack(pady=5)


def show_file_content(file_path, title):
    """显示文件内容"""
    window = tk.Toplevel()
    window.title(f"{title} - File Content")
    window.geometry("800x600")
    text_widget = tk.Text(window, wrap="word", font=("Consolas", 10))
    text_widget.pack(fill="both", expand=True, padx=10, pady=10)
    try:
        with open(file_path, 'r', encoding='utf-8') as file:
            content = file.read()
            text_widget.insert("1.0", content)
    except Exception as e:
        text_widget.insert("1.0", f"Error reading file: {str(e)}")
    text_widget.config(state="disabled")


def upload_spectral_library():
    """上传自定义谱图库"""
    file_path = filedialog.askopenfilename(
        filetypes=[("Mass Spectra Files", "*.msp"), ("All files", "*.*")],
        title="Select Spectral Library File to Upload"
    )

    if not file_path:
        return

    try:
        # 获取文件名
        file_name = os.path.basename(file_path)

        # 目标路径
        dest_path = os.path.join(DATA_DIR, file_name)

        # 检查是否已存在
        if os.path.exists(dest_path):
            if not messagebox.askyesno("File Exists",
                                       f"File {file_name} already exists. Do you want to overwrite?"):
                return

        # 复制文件
        shutil.copy2(file_path, dest_path)

        # 添加到光谱库字典
        current_time = datetime.datetime.now().strftime("%Y/%m/%d %H:%M")

        # 获取文件大小
        size_bytes = os.path.getsize(dest_path)
        if size_bytes < 1024 * 1024:
            size_str = f"{size_bytes / 1024:.1f} KB"
        else:
            size_str = f"{size_bytes / (1024 * 1024):.1f} MB"

        # 添加到光谱库
        lib_name = os.path.splitext(file_name)[0]
        spectral_libraries[lib_name] = {
            "date": current_time,
            "size": size_str,
            "count": "Custom",
            "download_file": file_name
        }

        messagebox.showinfo("Success", f"Library '{lib_name}' uploaded successfully!")

        # 刷新光谱库显示
        refresh_spectral_library()

    except Exception as e:
        messagebox.showerror("Error", f"Failed to upload library: {str(e)}")


def refresh_spectral_library():
    """刷新光谱库显示"""
    # 查找现有的光谱库窗口并更新
    for widget in tk._default_root.winfo_children():
        if isinstance(widget, tk.Toplevel) and widget.title() == "Spectral Library":
            widget.destroy()
            show_spectral_library()
            break


def download_spectral_file(event):
    """下载谱图文件的正确实现"""
    global spectral_tree
    # 获取点击的单元格
    item = spectral_tree.identify_row(event.y)
    if not item:
        return
    # 获取库名称（第一列的值）
    values = spectral_tree.item(item)['values']
    if not values:
        return
    lib_name = values[0]
    lib_data = spectral_libraries.get(lib_name)
    if lib_data and "download_file" in lib_data:
        file_path = os.path.join(DATA_DIR, lib_data["download_file"])
        # 检查文件是否存在
        if not os.path.exists(file_path):
            messagebox.showerror("File Not Found",
                                 f"Library file not found:\n{file_path}\n"
                                 f"Please make sure the file exists in the data directory.")
            return
        # 选择保存位置
        save_path = filedialog.asksaveasfilename(
            defaultextension=".msp",
            initialfile=os.path.basename(file_path),
            filetypes=[("Mass Spectra Files", "*.msp"), ("All files", "*.*")],
            title=f"Save {lib_name} Library"
        )
        if save_path:
            try:
                # 复制文件
                shutil.copy2(file_path, save_path)
                messagebox.showinfo("Success",
                                    f"{lib_name} library downloaded successfully!\n"
                                    f"Saved to: {save_path}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to save file: {str(e)}")


def show_spectral_library():
    """显示谱图库窗口"""
    global spectral_tree

    spectral_window = tk.Toplevel()
    spectral_window.title("Spectral Library")
    spectral_window.geometry("1000x700")
    spectral_window.configure(bg=colors["light"])

    # 标题
    title_frame = tk.Frame(spectral_window, bg=colors["primary"], height=60)
    title_frame.pack(fill="x", pady=(0, 20))

    title_label = tk.Label(title_frame,
                           text="📊 Spectral Library",
                           font=("Segoe UI", 18, "bold"),
                           fg="white",
                           bg=colors["primary"])
    title_label.pack(pady=15)

    # 按钮框架
    button_frame = tk.Frame(spectral_window, bg=colors["light"])
    button_frame.pack(fill="x", padx=20, pady=(0, 10))

    # 上传按钮
    upload_btn = tk.Button(button_frame,
                           text="📤 Upload Custom Library",
                           command=upload_spectral_library,
                           bg=colors["success"],
                           fg="white",
                           font=font_medium,
                           padx=15,
                           pady=5)
    upload_btn.pack(side="left")

    # 说明文字
    info_label = tk.Label(spectral_window,
                          text="Click the download button (⬇) to download spectral libraries",
                          font=("Segoe UI", 10),
                          fg=colors["dark"],
                          bg=colors["light"])
    info_label.pack(pady=(0, 15))

    # 创建表格框架
    table_frame = tk.Frame(spectral_window)
    table_frame.pack(fill="both", expand=True, padx=20, pady=10)

    # 定义列
    columns = ("Library Name", "Date", "Size", "Spectra Count", "Download")

    # 创建滚动条
    v_scroll = ttk.Scrollbar(table_frame, orient="vertical")
    h_scroll = ttk.Scrollbar(table_frame, orient="horizontal")

    # 创建树形视图
    spectral_tree = ttk.Treeview(table_frame,
                                 columns=columns,
                                 show="headings",
                                 height=15,
                                 yscrollcommand=v_scroll.set,
                                 xscrollcommand=h_scroll.set)

    # 配置列
    column_widths = [250, 150, 120, 120, 100]
    for i, col in enumerate(columns):
        spectral_tree.heading(col, text=col)
        spectral_tree.column(col, width=column_widths[i], anchor="center")

    # 添加数据
    for lib_name, lib_data in spectral_libraries.items():
        spectral_tree.insert("", tk.END, values=(
            lib_name,
            lib_data["date"],
            lib_data["size"],
            lib_data["count"],
            "⬇ Download"
        ))

    # 布局
    spectral_tree.pack(side="left", fill="both", expand=True)
    v_scroll.pack(side="right", fill="y")
    h_scroll.pack(side="bottom", fill="x")

    # 配置滚动条
    v_scroll.config(command=spectral_tree.yview)
    h_scroll.config(command=spectral_tree.xview)

    # 绑定点击事件
    spectral_tree.bind("<ButtonRelease-1>", download_spectral_file)

    # 底部信息
    footer_label = tk.Label(spectral_window,
                            text=f"Data directory: {DATA_DIR}",
                            font=("Segoe UI", 9),
                            fg=colors["dark"],
                            bg=colors["light"])
    footer_label.pack(pady=(15, 5))


def get_word_content(file_path):
    """从Word文档中精确提取文本内容"""
    try:
        doc = Document(file_path)
        return "\n".join(para.text for para in doc.paragraphs)
    except Exception as e:
        return f"Error reading Word document: {str(e)}"


def open_help_document(file_key):
    """显示完全匹配图片样式的Word文档内容"""
    STYLE = {
        'colors': {
            'bg': "#FFFFFF",
            'text': "#000000",
            'header': "#3A7CA5"
        },
        'fonts': {
            'title': ("Times New Roman", 14, "bold"),
            'body': ("Times New Roman", 12),
            'ref': ("Times New Roman", 11, "italic")
        },
        'indent': 24
    }

    win = tk.Toplevel()
    win.title(f"Help Documentation: {file_key}")
    win.geometry("700x600")
    win.configure(bg=STYLE['colors']['bg'])

    main_frame = tk.Frame(win, bg=STYLE['colors']['bg'])
    main_frame.pack(fill="both", expand=True, padx=20, pady=20)

    tk.Label(main_frame,
             text=f"Help Documentation: {file_key}",
             fg=STYLE['colors']['text'],
             bg=STYLE['colors']['bg'],
             font=STYLE['fonts']['title'],
             anchor="w").pack(fill="x", pady=(0, 15))

    text_area = tk.Text(main_frame,
                        wrap="word",
                        font=STYLE['fonts']['body'],
                        bg=STYLE['colors']['bg'],
                        fg=STYLE['colors']['text'],
                        padx=20,
                        pady=10,
                        tabs=(f'{STYLE["indent"]}px', 'right'),
                        borderwidth=0,
                        highlightthickness=0)

    scrollbar = ttk.Scrollbar(main_frame, command=text_area.yview)
    scrollbar.pack(side="right", fill="y")
    text_area.config(yscrollcommand=scrollbar.set)
    text_area.pack(side="left", fill="both", expand=True)

    help_path = SIDEBAR_HELP_PATHS.get(file_key)
    if help_path and os.path.exists(help_path):
        content = get_word_content(help_path)
    else:
        content = f"Document not found: {help_path}"

    text_area.insert("1.0", content)

    text_area.tag_configure("indent", lmargin1=STYLE['indent'])
    text_area.tag_configure("ref", font=STYLE['fonts']['ref'])

    lines = content.split("\n")
    for i, line in enumerate(lines, start=1):
        if not line.strip(): continue
        if line.startswith(("US EPA", "Guoli Xiong", "Sushko, I", "Zhang L", "CompTox", "Djoumbou Feunang, Y")):
            text_area.tag_add("ref", f"{i}.0", f"{i}.end")

    text_area.config(state="disabled")


def run_suspect_list_screening():
    """Run Mass Spectrometry Data Matching Analysis Tool (LC-MS & GC-MS)"""
    try:
        # ========== 添加这行代码：设置 matchms 数据路径 ==========
        import os
        import sys
        import matchms
        matchms_data_path = os.path.join(os.path.dirname(matchms.__file__), 'data')
        if not os.path.exists(matchms_data_path):
            # 如果是打包环境，尝试使用临时目录
            if getattr(sys, 'frozen', False):
                base_path = sys._MEIPASS
                matchms_data_path = os.path.join(base_path, 'matchms', 'data')

        # ========== 修复版：Set OpenMS Environment Variables ==========
        # 检查是否是打包环境
        if getattr(sys, 'frozen', False):
            # 打包环境：使用临时目录中的路径
            base_path = sys._MEIPASS
            openms_data_path = os.path.join(base_path, 'pyopenms', 'share', 'OpenMS')

            # 如果上面的路径不存在，尝试其他可能的打包路径
            if not os.path.exists(openms_data_path):
                possible_packaged_paths = [
                    os.path.join(base_path, 'share', 'OpenMS'),
                    os.path.join(base_path, 'OpenMS'),
                    os.path.join(base_path, 'pyopenms', 'share', 'OpenMS'),
                ]
                for path in possible_packaged_paths:
                    if os.path.exists(path):
                        openms_data_path = path
                        break
        else:
            # 开发环境：使用虚拟环境中的路径
            # 首先尝试找到pyopenms的安装位置
            import pyopenms
            pyopenms_path = os.path.dirname(pyopenms.__file__)
            openms_data_path = os.path.join(pyopenms_path, 'share', 'OpenMS')

            # 如果找不到，尝试其他开发环境路径
            if not os.path.exists(openms_data_path):
                possible_paths = [
                    "D:/软件设计/.venv/Lib/site-packages/pyopenms/share/OpenMS",
                    os.path.join(sys.prefix, "Lib", "site-packages", "pyopenms", "share", "OpenMS"),
                    os.path.join(sys.prefix, "Library", "share", "OpenMS"),
                    os.path.join(os.path.dirname(__file__), ".venv", "Lib", "site-packages", "pyopenms", "share",
                                 "OpenMS"),
                ]
                for path in possible_paths:
                    if os.path.exists(path):
                        openms_data_path = path
                        break

        # 设置环境变量
        if os.path.exists(openms_data_path):
            os.environ['OPENMS_DATA_PATH'] = openms_data_path
            print(f"✓ Set OPENMS_DATA_PATH to: {openms_data_path}")
        else:
            print("⚠ Warning: Could not find OpenMS data path")
            # 创建临时目录作为备选
            import tempfile
            temp_dir = tempfile.mkdtemp(prefix='OpenMS_')
            os.environ['OPENMS_DATA_PATH'] = temp_dir
            print(f"✓ Created temporary OPENMS_DATA_PATH: {temp_dir}")

        # ========== Import Required Libraries ==========
        import tkinter as tk
        from tkinter import ttk, filedialog, messagebox
        import pandas as pd
        import threading
        import numpy as np
        from scipy.spatial.distance import cosine
        from scipy.signal import find_peaks, savgol_filter
        import re
        import glob
        import logging
        from matchms import Spectrum
        from openpyxl import Workbook
        from typing import List, Dict, Tuple
        import ast  # 用于安全解析字符串列表

        # 导入matplotlib用于镜像图绘制
        import matplotlib
        matplotlib.use('Agg')  # 使用Agg后端，不显示图形窗口
        import matplotlib.pyplot as plt

        # Import pyopenms
        try:
            import pyopenms as oms
            print("Successfully imported pyopenms")
        except Exception as e:
            error_msg = f"Failed to import pyopenms: {str(e)}\n\nPlease make sure OpenMS is properly installed."
            print(error_msg)
            messagebox.showerror("OpenMS Error", error_msg)
            return

        # Reduce log output
        logging.getLogger('matchms').setLevel(logging.ERROR)

        # Create analysis window
        analysis_window = tk.Toplevel()
        analysis_window.title("Mass Spectrometry Data Matching - LC-MS & GC-MS")
        analysis_window.geometry("1400x950")

        # ==================== 动态阈值计算函数 ====================
        def calculate_dynamic_thresholds(intensities):
            """根据数据自动计算阈值"""
            if len(intensities) == 0:
                return 1000, 500  # 默认值

            # 只考虑正强度值
            positive_intensities = intensities[intensities > 0]
            if len(positive_intensities) == 0:
                return 1000, 500

            # 使用百分位数计算阈值
            min_intensity = np.percentile(positive_intensities, 5)  # 5%分位数作为最小强度
            noise_threshold = np.percentile(positive_intensities, 10)  # 10%分位数作为噪声阈值

            # 确保阈值不小于合理的最小值
            min_intensity = max(100, min_intensity)
            noise_threshold = max(50, noise_threshold)

            return round(min_intensity, 0), round(noise_threshold, 0)

        def extract_all_intensities(exp):
            """从MSExperiment中提取所有强度值"""
            all_intensities = []
            try:
                for spectrum in exp:
                    if spectrum.getMSLevel() == 1:  # 只考虑MS1谱图
                        for peak in spectrum:
                            intensity = peak.getIntensity()
                            if intensity > 0:
                                all_intensities.append(intensity)
            except Exception as e:
                log_message(f"提取强度时出错: {str(e)}")
            return np.array(all_intensities)

        # ==================== 镜像图绘制函数 ====================
        def parse_peaks(peaks_str):
            """将峰字符串解析为字典{m/z: intensity}"""
            if isinstance(peaks_str, str):
                try:
                    # 尝试解析字符串格式的列表
                    peaks_list = ast.literal_eval(peaks_str)
                    if isinstance(peaks_list, list):
                        return {float(p[0]): float(p[1]) for p in peaks_list if len(p) >= 2}
                except:
                    # 如果解析失败，尝试其他格式
                    pass
            return {}

        def get_top_peaks(peaks_dict, n=5):
            """获取强度最高的n个峰"""
            sorted_peaks = sorted(peaks_dict.items(), key=lambda x: x[1], reverse=True)
            return dict(sorted_peaks[:n])

        def create_mirror_plot_from_excel_data(exp_peaks, lib_peaks, compound_name, precursor_mz,
                                               rt, score, save_path, method="LC-MS"):
            """从Excel数据创建镜像图"""
            try:
                # 设置全局样式
                plt.rcParams.update({
                    'font.family': 'Arial',
                    'font.size': 8,
                    'axes.titlesize': 8,
                    'axes.labelsize': 8,
                    'xtick.labelsize': 6,
                    'ytick.labelsize': 6,
                    'axes.edgecolor': 'black',
                    'axes.linewidth': 0.5,
                    'figure.facecolor': 'white',
                    'savefig.facecolor': 'white',
                })

                # 创建图形
                fig, ax = plt.subplots(figsize=(4, 3))

                # 归一化数据
                max_exp = max(exp_peaks.values()) if exp_peaks else 1
                exp_peaks_norm = {mz: intensity / max_exp * 100 for mz, intensity in exp_peaks.items()}

                max_lib = max(lib_peaks.values()) if lib_peaks else 1
                lib_peaks_norm = {mz: intensity / max_lib * 100 for mz, intensity in lib_peaks.items()}

                # 获取强度最高的峰用于标注
                top_exp_peaks = get_top_peaks(exp_peaks_norm, n=5)
                top_lib_peaks = get_top_peaks(lib_peaks_norm, n=5)

                # 绘制所有峰
                for mz, intensity in lib_peaks_norm.items():
                    ax.plot([mz, mz], [0, intensity], color='#ff7f0e', linewidth=0.8, alpha=0.7)

                for mz, intensity in exp_peaks_norm.items():
                    ax.plot([mz, mz], [0, -intensity], color='#1f77b4', linewidth=0.8, alpha=0.7)

                # 添加峰标注（只标注前5个强度最高的峰）- 改为横着显示
                for mz, intensity in top_lib_peaks.items():
                    ax.text(mz, intensity + 3, f"{mz:.1f}" if method == "GC-MS" else f"{mz:.2f}",
                            ha='center', va='bottom', fontsize=6,
                            color='black', rotation=0)  # rotation=0 表示横着显示

                for mz, intensity in top_exp_peaks.items():
                    ax.text(mz, -intensity - 3, f"{mz:.1f}" if method == "GC-MS" else f"{mz:.2f}",
                            ha='center', va='top', fontsize=6,
                            color='black', rotation=0)  # rotation=0 表示横着显示

                # 设置坐标轴
                ax.axhline(0, color='black', linewidth=0.5)

                # 添加标题（使用RT作为文件名的一部分）
                if method == "LC-MS":
                    title = f"{compound_name[:30]}{'...' if len(compound_name) > 30 else ''}\n"
                    title += f"Precursor: {precursor_mz:.4f} | RT: {rt:.2f}s\n"
                    title += f"Score: {score:.3f}"
                else:
                    title = f"{compound_name[:30]}{'...' if len(compound_name) > 30 else ''}\n"
                    title += f"RT: {rt:.2f}min | Score: {score:.3f}\n"

                ax.set_title(title, fontsize=7, pad=10)

                # 设置坐标轴范围
                all_mz = list(exp_peaks_norm.keys()) + list(lib_peaks_norm.keys())
                if all_mz:
                    if method == "GC-MS":
                        x_min = min(all_mz) - 5
                        x_max = max(all_mz) + 5
                    else:
                        x_min = min(all_mz) - 15
                        x_max = max(all_mz) + 15
                    ax.set_xlim(x_min, x_max)

                # 设置y轴范围
                y_max = max(lib_peaks_norm.values()) + 20 if lib_peaks_norm else 100
                y_min = -max(exp_peaks_norm.values()) - 20 if exp_peaks_norm else -100
                ax.set_ylim(y_min, y_max)

                # 设置坐标轴标签
                ax.set_xlabel('m/z', fontsize=7)
                ax.set_ylabel('Relative Intensity (%)', fontsize=7)

                # 添加图例说明 - 无边框
                ax.text(0.02, 0.98, f'Library ({len(lib_peaks)} peaks)',
                        transform=ax.transAxes, color='#ff7f0e', fontsize=6,
                        va='top')  # 移除bbox参数，无边框

                ax.text(0.02, 0.02, f'Experimental ({len(exp_peaks)} peaks)',
                        transform=ax.transAxes, color='#1f77b4', fontsize=6,
                        va='bottom')  # 移除bbox参数，无边框

                # 调整布局
                plt.tight_layout()

                # 保存为SVG
                fig.savefig(save_path, format='svg', dpi=300,
                            bbox_inches='tight', pad_inches=0.1,
                            facecolor='white')

                plt.close(fig)
                return True

            except Exception as e:
                log_message(f"创建镜像图失败: {str(e)}")
                return False

        def generate_mirror_plots_from_results(excel_path):
            """从结果Excel文件生成镜像图（只绘制得分大于阈值的匹配）"""
            try:
                # 读取Excel文件
                df = pd.read_excel(excel_path, sheet_name='All Results')

                if df.empty:
                    log_message("没有数据可生成镜像图")
                    return

                # 确定方法类型和阈值
                method = "LC-MS" if 'precursor_mz' in df.columns else "GC-MS"

                # 获取阈值
                if method == "LC-MS":
                    threshold = float(entries["DP_THRESHOLD"].get())
                    valid_df = df[df['cosine_score'] >= threshold]
                    log_message(f"LC-MS 阈值: {threshold}, 有效匹配数: {len(valid_df)}/{len(df)}")
                else:
                    threshold = float(gc_entries["GC_SIMILARITY_THRESHOLD"].get())
                    valid_df = df[df['similarity'] >= threshold]
                    log_message(f"GC-MS 阈值: {threshold}, 有效匹配数: {len(valid_df)}/{len(df)}")

                if valid_df.empty:
                    log_message("没有得分大于阈值的匹配，不生成镜像图")
                    return

                # 每个化合物只保留前5个最高得分的用于镜像图（避免生成太多）
                if method == "LC-MS" and 'Name' in valid_df.columns:
                    valid_df_sorted = valid_df.sort_values('cosine_score', ascending=False)
                    valid_df_filtered = valid_df_sorted.groupby('Name').head(5).reset_index(drop=True)
                    log_message(f"每个化合物保留前5个最高得分用于镜像图: {len(valid_df_filtered)}/{len(valid_df)}")
                    valid_df = valid_df_filtered
                elif method == "GC-MS" and 'compound' in valid_df.columns:
                    valid_df_sorted = valid_df.sort_values('similarity', ascending=False)
                    valid_df_filtered = valid_df_sorted.groupby('compound').head(5).reset_index(drop=True)
                    log_message(f"每个化合物保留前5个最高得分用于镜像图: {len(valid_df_filtered)}/{len(valid_df)}")
                    valid_df = valid_df_filtered

                # 获取原始文件名（用于镜像图文件夹命名）
                excel_filename = os.path.basename(excel_path)
                base_filename = os.path.splitext(excel_filename)[0]

                # 创建镜像图文件夹 - 添加文件名前缀以便区分
                save_dir = os.path.dirname(excel_path)
                mirror_dir = os.path.join(save_dir, f"{base_filename}_mirror_plots")
                os.makedirs(mirror_dir, exist_ok=True)

                log_message(f"镜像图保存文件夹: {mirror_dir}")

                generated_count = 0

                for idx, row in valid_df.iterrows():
                    try:
                        # 解析谱图数据
                        exp_peaks_str = row.get('exp_spectrum_peaks', '')
                        lib_peaks_str = row.get('lib_spectrum_peaks', '')

                        if not exp_peaks_str or not lib_peaks_str:
                            continue

                        exp_peaks = parse_peaks(exp_peaks_str)
                        lib_peaks = parse_peaks(lib_peaks_str)

                        if not exp_peaks or not lib_peaks:
                            continue

                        # 获取化合物信息
                        if method == "LC-MS":
                            compound_name = row.get('Name', 'Unknown')
                            precursor_mz = row.get('precursor_mz', 0)
                            rt = row.get('rt', 0)
                            score = row.get('cosine_score', 0)
                        else:
                            compound_name = row.get('compound', 'Unknown')
                            precursor_mz = 0
                            rt = row.get('retention_time', 0)
                            score = row.get('similarity', 0)

                        # 创建安全的文件名（使用RT、化合物名称和得分）
                        safe_rt = f"{float(rt):.2f}".replace('.', '_')
                        safe_name = re.sub(r'[\\/*?:"<>|]', "_", compound_name)[:30]
                        # 添加序号前缀，便于排序
                        svg_filename = f"{idx + 1:04d}_RT_{safe_rt}_{safe_name}_score_{score:.3f}.svg"
                        svg_path = os.path.join(mirror_dir, svg_filename)

                        # 创建镜像图
                        success = create_mirror_plot_from_excel_data(
                            exp_peaks, lib_peaks,
                            compound_name, precursor_mz, rt, score,
                            svg_path, method
                        )

                        if success:
                            generated_count += 1

                    except Exception as e:
                        log_message(f"处理第{idx + 1}行时出错: {str(e)}")
                        continue

                log_message(
                    f"✓ 镜像图生成完成: {generated_count}/{len(valid_df)} 个有效匹配已生成，保存至: {mirror_dir}")

            except Exception as e:
                log_message(f"生成镜像图失败: {str(e)}")

        def save_comprehensive_results(file_results, save_path, generate_mirror=False, method="LC-MS"):
            """保存全面的结果（每个化合物只保留得分前10的结果）"""
            if file_results is None or file_results.empty:
                return False

            try:
                save_dir = os.path.dirname(save_path)
                base_name = os.path.splitext(os.path.basename(save_path))[0]

                # 按化合物分组，每个化合物只保留得分前10的结果
                if method == "LC-MS":
                    # LC-MS按Name分组
                    if 'Name' in file_results.columns:
                        # 先按得分排序
                        file_results_sorted = file_results.sort_values('cosine_score', ascending=False)
                        # 按Name分组，取每个组的前10个
                        file_results_filtered = file_results_sorted.groupby('Name').head(10).reset_index(drop=True)
                        removed_count = len(file_results) - len(file_results_filtered)
                        log_message(
                            f"每个化合物保留前10个最高得分匹配: 原始 {len(file_results)} 条 → 保留 {len(file_results_filtered)} 条 (删除了 {removed_count} 条)")
                        file_results = file_results_filtered
                    else:
                        log_message("警告: 未找到Name列，无法按化合物分组")
                else:
                    # GC-MS按compound分组
                    if 'compound' in file_results.columns:
                        # 先按相似度排序
                        file_results_sorted = file_results.sort_values('similarity', ascending=False)
                        # 按compound分组，取每个组的前10个
                        file_results_filtered = file_results_sorted.groupby('compound').head(10).reset_index(drop=True)
                        removed_count = len(file_results) - len(file_results_filtered)
                        log_message(
                            f"每个化合物保留前10个最高得分匹配: 原始 {len(file_results)} 条 → 保留 {len(file_results_filtered)} 条 (删除了 {removed_count} 条)")
                        file_results = file_results_filtered
                    else:
                        log_message("警告: 未找到compound列，无法按化合物分组")

                # 1. 保存Excel结果文件
                with pd.ExcelWriter(save_path, engine='openpyxl') as writer:
                    # 所有结果表（已过滤，每个化合物只保留前10）
                    file_results.to_excel(writer, sheet_name='All Results', index=False)

                    # 有效结果表（只显示分数高于阈值的）
                    if method == "LC-MS" and 'is_valid_match' in file_results.columns:
                        valid_results = file_results[file_results['is_valid_match'] == True]
                        if not valid_results.empty:
                            valid_results.to_excel(writer, sheet_name='Valid Results', index=False)

                    # GC-MS的有效结果表
                    if method == "GC-MS" and 'similarity' in file_results.columns:
                        threshold = float(gc_entries["GC_SIMILARITY_THRESHOLD"].get())
                        valid_results = file_results[file_results['similarity'] >= threshold]
                        if not valid_results.empty:
                            valid_results.to_excel(writer, sheet_name='Valid Results', index=False)

                    # 统计信息表
                    stats_data = []
                    if not file_results.empty:
                        if method == "LC-MS":
                            threshold = float(entries["DP_THRESHOLD"].get())
                            stats_data = [
                                ['统计项', '数值'],
                                ['总匹配数（每个化合物前10）', len(file_results)],
                                [f'有效匹配(分数≥{threshold})',
                                 file_results[
                                     'is_valid_match'].sum() if 'is_valid_match' in file_results.columns else 0],
                                ['平均分数',
                                 file_results['cosine_score'].mean() if 'cosine_score' in file_results.columns else 0],
                                ['最高分数',
                                 file_results['cosine_score'].max() if 'cosine_score' in file_results.columns else 0],
                                ['平均前体偏差(ppm)', file_results[
                                    'precursor_error_ppm'].mean() if 'precursor_error_ppm' in file_results.columns else 0],
                                ['平均匹配碎片数', file_results[
                                    'matched_peaks_count'].mean() if 'matched_peaks_count' in file_results.columns else 0],
                                ['平均匹配偏差(mDa)', file_results[
                                    'matched_peaks_avg_error_mda'].mean() if 'matched_peaks_avg_error_mda' in file_results.columns else 0],
                                ['化合物数量',
                                 file_results['Name'].nunique() if 'Name' in file_results.columns else 'N/A'],
                            ]
                        else:  # GC-MS
                            threshold = float(gc_entries["GC_SIMILARITY_THRESHOLD"].get())
                            stats_data = [
                                ['统计项', '数值'],
                                ['总匹配数（每个化合物前10）', len(file_results)],
                                [f'有效匹配(分数≥{threshold})',
                                 len(file_results[file_results[
                                                      'similarity'] >= threshold]) if 'similarity' in file_results.columns else 0],
                                ['平均相似度',
                                 file_results['similarity'].mean() if 'similarity' in file_results.columns else 0],
                                ['最高相似度',
                                 file_results['similarity'].max() if 'similarity' in file_results.columns else 0],
                                ['平均匹配碎片数', file_results[
                                    'matched_peaks_count'].mean() if 'matched_peaks_count' in file_results.columns else 0],
                                ['化合物数量',
                                 file_results['compound'].nunique() if 'compound' in file_results.columns else 'N/A'],
                            ]
                    pd.DataFrame(stats_data).to_excel(writer, sheet_name='Statistics', index=False)

                log_message(f"✓ 结果文件已保存: {os.path.basename(save_path)}")

                # 2. 保存质量日志
                log_path = save_quality_log(save_dir)
                log_message(f"✓ 质量日志已保存: {os.path.basename(log_path)}")

                # 3. 如果选择生成镜像图，则生成（只绘制有效匹配）
                if generate_mirror:
                    log_message(f"开始生成镜像图（只绘制得分大于阈值的匹配）...")
                    generate_mirror_plots_from_results(save_path)

                return True

            except Exception as e:
                log_message(f"✗ 保存结果失败: {str(e)}")
                import traceback
                log_message(traceback.format_exc())
                return False

        # Configuration parameter defaults
        config = {
            # LC-MS parameters
            "MS1_TOL_DA": 0.005,
            "MS2_TOL_DA": 0.01,
            "MIN_INTENSITY": 5000,
            "DP_THRESHOLD": 0.6,
            "MIN_PEAKS": 3,  # 从2改为3，提高匹配可靠性
            "NORMALIZATION_METHOD": "max",
            "NOISE_THRESHOLD": 1000,
            "REMOVE_SINGLE_TRACES": "false",
            # GC-MS parameters
            "GC_PEAK_HEIGHT": 1000,
            "GC_PEAK_DISTANCE": 5,
            "GC_PEAK_PROMINENCE": 500,
            "GC_PEAK_WIDTH": 2,
            "GC_RT_TOLERANCE": 0.05,
            "GC_SIMILARITY_THRESHOLD": 0.7
        }

        # Store result data
        result_data = None
        processing = False
        batch_files = []
        current_file_index = 0
        current_method = tk.StringVar(value="LC-MS")
        generate_mirror_var = tk.BooleanVar(value=False)  # 是否生成镜像图
        use_dynamic_thresholds_var = tk.BooleanVar(value=True)  # 是否使用动态阈值
        quality_log = []  # 用于存储质量日志

        # ==================== LC-MS Functions ====================
        def normalize_intensity(spectrum, method="max", min_intensity=0):
            """Intensity normalization function"""
            if len(spectrum) == 0:
                return spectrum
            normalized_spectrum = spectrum.copy()
            intensities = normalized_spectrum[:, 1]
            if min_intensity > 0:
                intensities[intensities < min_intensity] = 0
            nonzero_mask = intensities > 0
            if not np.any(nonzero_mask):
                return np.array([], dtype=np.float64).reshape(0, 2)
            nonzero_intensities = intensities[nonzero_mask]
            if method == "max":
                max_val = np.max(nonzero_intensities)
                if max_val > 0:
                    nonzero_intensities = nonzero_intensities / max_val * 100
            elif method == "sqrt":
                nonzero_intensities = np.sqrt(nonzero_intensities)
                max_val = np.max(nonzero_intensities)
                if max_val > 0:
                    nonzero_intensities = nonzero_intensities / max_val * 100
            elif method == "sum":
                total = np.sum(nonzero_intensities)
                if total > 0:
                    nonzero_intensities = nonzero_intensities / total * 100
            intensities[nonzero_mask] = nonzero_intensities
            normalized_spectrum[:, 1] = intensities
            return normalized_spectrum[nonzero_mask]

        def peak_picking_ms2(df_peakpicking, input_path, ms_mode, noise_threshold=1000, remove_single_traces="false"):
            """改进的MS2谱图提取函数 - 包含质量日志记录"""
            df = df_peakpicking.copy()
            df = df.sort_values(by='RT', ascending=True)

            # 初始化列
            df['Precursor_mz'] = df.get('Precursor_mz', 0.0)
            df['MS2_rt'] = df.get('MS2_rt', 0.0)
            df['MS2_mz'] = df.get('MS2_mz', None)
            df['MS2_Intensity'] = df.get('MS2_Intensity', None)
            df['MS2_matched'] = False

            # 读取质谱文件
            run = oms.MSExperiment()
            oms.MzMLFile().load(input_path, run)
            match_count = 0
            total_ms2 = 0

            if ms_mode == 'DDA':
                for spectrum in run:
                    if spectrum.getMSLevel() == 2:
                        total_ms2 += 1
                        precursors = spectrum.getPrecursors()
                        if len(precursors) == 0:
                            continue

                        precursor_mz = precursors[0].getMZ()
                        isolation_window = precursors[0].getIsolationWindowLowerOffset()
                        if isolation_window <= 0:
                            isolation_window = 0.5

                        if spectrum.size() == 0:
                            continue

                        mzs2 = []
                        intensities2 = []
                        for peak in spectrum:
                            mzs2.append(peak.getMZ())
                            intensities2.append(peak.getIntensity())

                        mzs2 = np.array(mzs2)
                        intensities2 = np.array(intensities2)
                        rt2 = spectrum.getRT()

                        if intensities2.size == 0:
                            continue

                        # 计算噪声水平
                        threshold_ms2 = max(noise_threshold, intensities2.max() * 0.01)

                        filtered_indices_ms2 = intensities2 >= threshold_ms2
                        filtered_mzs2 = mzs2[filtered_indices_ms2]
                        filtered_intensities2 = intensities2[filtered_indices_ms2]

                        peak_count = len(filtered_mzs2)

                        mz_matches = df[
                            (abs(df['m/z'] - precursor_mz) <= isolation_window)
                        ]

                        if mz_matches.empty:
                            continue

                        rt_matches = mz_matches[
                            (abs(mz_matches['RT'] - rt2) <= 15)
                        ]

                        if rt_matches.empty:
                            continue

                        rt_diffs = abs(rt_matches['RT'] - rt2)
                        best_match_idx = rt_diffs.idxmin()

                        current_ms2_rt = df.at[best_match_idx, 'MS2_rt']
                        if (current_ms2_rt == 0.0 or
                                abs(df.at[best_match_idx, 'RT'] - rt2) < abs(
                                    df.at[best_match_idx, 'RT'] - current_ms2_rt)):
                            # 更新MS2数据
                            df.at[best_match_idx, 'Precursor_mz'] = round(precursor_mz, 5)
                            df.at[best_match_idx, 'MS2_rt'] = round(rt2, 5)
                            df.at[best_match_idx, 'MS2_mz'] = [round(mz, 5) for mz in filtered_mzs2]
                            df.at[best_match_idx, 'MS2_Intensity'] = [round(intensity, 5) for intensity in
                                                                      filtered_intensities2]
                            df.at[best_match_idx, 'MS2_matched'] = True
                            match_count += 1

            # 记录质量日志
            quality_log.append(f"MS2谱图质量统计:")
            quality_log.append(f"  总MS2谱图数: {total_ms2}")
            quality_log.append(f"  匹配特征数: {match_count}")

            return df

        def perform_feature_detection(input_path, noise_threshold=1000, remove_single_traces="false"):
            """改进的特征检测函数 - 包含质量日志记录"""
            print(f"Processing file {input_path}")

            # 加载MZML文件
            exp = oms.MSExperiment()
            oms.MzMLFile().load(input_path, exp)
            exp.sortSpectra(True)

            # 如果使用动态阈值，计算当前文件的动态阈值
            if use_dynamic_thresholds_var.get():
                all_intensities = extract_all_intensities(exp)
                min_intensity, noise_threshold = calculate_dynamic_thresholds(all_intensities)
                log_message(f"动态阈值计算: MIN_INTENSITY={min_intensity:.0f}, NOISE_THRESHOLD={noise_threshold:.0f}")
            else:
                min_intensity = float(entries["MIN_INTENSITY"].get())
                noise_threshold = float(entries["NOISE_THRESHOLD"].get())

            # Step 1: Mass Trace Detection
            mass_traces = []
            mtd = oms.MassTraceDetection()
            mtd_params = mtd.getDefaults()
            mtd_params.setValue("mass_error_ppm", float(10.0))
            mtd_params.setValue("noise_threshold_int", float(noise_threshold))
            mtd.setParameters(mtd_params)
            mtd.run(exp, mass_traces, 0)

            quality_log.append(f"检测到 {len(mass_traces)} 个质量踪迹")

            # Step 2: Elution Peak Detection
            mass_traces_split = []
            mass_traces_final = []
            epd = oms.ElutionPeakDetection()
            epd_params = epd.getDefaults()
            epd_params.setValue("width_filtering", "off")
            epd.setParameters(epd_params)
            epd.detectPeaks(mass_traces, mass_traces_split)

            if epd.getParameters().getValue("width_filtering") == "auto":
                epd.filterByPeakWidth(mass_traces_split, mass_traces_final)
            else:
                mass_traces_final = mass_traces_split

            # Step 3: Feature Finding
            fm = oms.FeatureMap()
            feat_chrom = []
            ffm = oms.FeatureFindingMetabo()
            ffm_params = ffm.getDefaults()
            ffm_params.setValue("isotope_filtering_model", "none")
            ffm_params.setValue("remove_single_traces", remove_single_traces)
            ffm_params.setValue("report_convex_hulls", "true")
            ffm.setParameters(ffm_params)
            ffm.run(mass_traces_final, fm, feat_chrom)

            quality_log.append(f"检测到 {fm.size()} 个特征")

            # 提取文件名
            source_file = os.path.basename(input_path).replace('.mzML', '').replace('.mzml', '')

            # 提取特征数据
            feature_data = []
            count = 1
            filtered_by_intensity = 0

            for feature in fm:
                rt_value = feature.getRT()
                if 3000 > rt_value > 60:
                    mz_value = feature.getMZ()
                    intensity_value = feature.getIntensity()

                    # 根据最小强度过滤
                    if intensity_value < min_intensity:
                        filtered_by_intensity += 1
                        continue

                    width = feature.getWidth()
                    rt_begin = rt_value - width
                    rt_end = rt_value + width

                    feature_data.append({
                        'FeatureID': count,
                        'm/z': round(mz_value, 5),
                        'RT': round(rt_value, 5),
                        'RT_begin': round(rt_begin, 5),
                        'RT_end': round(rt_end, 5),
                        'Intensity': round(intensity_value, 5),
                        'source_file': source_file,
                        'Precursor_mz': 0.0,
                        'MS2_rt': 0.0,
                        'MS2_mz': None,
                        'MS2_Intensity': None,
                        'MS2_matched': False
                    })
                    count += 1

            quality_log.append(f"特征质量统计:")
            quality_log.append(f"  总特征数: {len(feature_data)}")
            quality_log.append(f"  因强度过低被过滤的特征: {filtered_by_intensity}")

            df = pd.DataFrame(feature_data)
            return df

        # ==================== GC-MS Related Classes and Functions ====================
        class MSPParser:
            """MSP文件解析器 - 用于GC-MS"""

            def parse_msp_file(self, file_path: str) -> List[Dict]:
                """解析MSP文件，返回化合物列表"""
                compounds = []
                current_compound = {}
                spectrum = []
                skipped_compounds = 0
                valid_compounds = 0

                try:
                    with open(file_path, 'r', encoding='utf-8', errors='ignore') as file:
                        for line in file:
                            line = line.strip()

                            if not line:
                                if current_compound and spectrum:
                                    current_compound['spectrum'] = np.array(spectrum)
                                    compounds.append(current_compound)
                                    valid_compounds += 1
                                    current_compound = {}
                                    spectrum = []
                                continue

                            if ':' in line:
                                key, value = line.split(':', 1)
                                key, value = key.strip(), value.strip()
                                if key == 'Name':
                                    current_compound['name'] = value
                                elif key == 'Formula':
                                    current_compound['formula'] = value
                                elif key in ['CAS#', 'CAS', 'CASNO']:
                                    current_compound['cas_number'] = value
                                elif key == 'MW' or key == 'MolecularWeight':
                                    try:
                                        current_compound['mw'] = float(value)
                                    except:
                                        current_compound['mw'] = 0
                                elif key == 'Comments' or key == 'Comment':
                                    current_compound['comments'] = value
                            else:
                                parts = re.split(r'\s+', line)
                                if len(parts) >= 2:
                                    try:
                                        mz = float(parts[0])
                                        intensity = float(parts[1])
                                        if mz > 0 and intensity >= 0:
                                            spectrum.append([mz, intensity])
                                    except ValueError:
                                        continue

                    # 处理最后一个化合物
                    if current_compound and spectrum:
                        current_compound['spectrum'] = np.array(spectrum)
                        compounds.append(current_compound)
                        valid_compounds += 1

                    # 记录质量日志
                    quality_log.append(f"MSP库解析统计:")
                    quality_log.append(f"  总化合物数: {valid_compounds}")
                    quality_log.append(f"  跳过化合物: {skipped_compounds}")

                    # 统计谱图质量
                    spectra_with_peaks = [c for c in compounds if len(c.get('spectrum', [])) >= 10]
                    quality_log.append(f"  高质量谱图(≥10峰): {len(spectra_with_peaks)}")

                    return compounds

                except Exception as e:
                    log_message(f"Error parsing MSP file: {e}")
                    quality_log.append(f"MSP库解析错误: {str(e)}")
                    return []

        class GCMSProcessor:
            """GC-MS数据处理器"""

            def __init__(self):
                self.peak_params = {
                    'height': 1000, 'distance': 5, 'prominence': 500, 'width': 2
                }
                self.quality_log = []

            def extract_chromatogram(self, file_path: str) -> Tuple[np.ndarray, np.ndarray]:
                """提取TIC色谱图"""
                retention_times = []
                tic_intensities = []

                try:
                    exp = oms.MSExperiment()
                    oms.MzMLFile().load(file_path, exp)

                    ms1_count = 0
                    for spectrum in exp:
                        if spectrum.getMSLevel() == 1:
                            ms1_count += 1
                            retention_times.append(spectrum.getRT() / 60.0)  # 转换为分钟

                            # 计算TIC
                            tic = 0.0
                            for peak in spectrum:
                                tic += peak.getIntensity()
                            tic_intensities.append(tic)

                    quality_log.append(f"GC-MS: 读取到 {ms1_count} 个MS1谱图")
                    return np.array(retention_times), np.array(tic_intensities)

                except Exception as e:
                    log_message(f"Error reading mzML file: {e}")
                    quality_log.append(f"GC-MS读取错误: {str(e)}")
                    return np.array([]), np.array([])

            def preprocess_chromatogram(self, time: np.ndarray, intensity: np.ndarray) -> Tuple[np.ndarray, np.ndarray]:
                """预处理色谱数据"""
                if len(intensity) == 0:
                    return time, intensity

                # 记录原始数据质量
                quality_log.append(f"GC-MS: 原始色谱点数量: {len(intensity)}")
                quality_log.append(f"GC-MS: 最大强度: {np.max(intensity):.2f}")
                quality_log.append(f"GC-MS: 最小强度: {np.min(intensity):.2f}")

                # 去除零和低值
                mask = intensity > np.percentile(intensity, 5)
                time = time[mask]
                intensity = intensity[mask]

                quality_log.append(f"GC-MS: 过滤后色谱点数量: {len(intensity)}")

                if len(intensity) == 0:
                    return time, intensity

                # 应用平滑滤波
                if len(intensity) > 11:
                    window_length = min(11, len(intensity) - 1)
                    if window_length % 2 == 0:
                        window_length -= 1
                    try:
                        smoothed = savgol_filter(intensity, window_length=window_length, polyorder=3)
                        quality_log.append(f"GC-MS: 应用平滑滤波，窗口大小: {window_length}")
                    except:
                        smoothed = intensity
                else:
                    smoothed = intensity

                # 基线校正
                baseline = np.percentile(smoothed, 10)
                corrected = smoothed - baseline
                corrected[corrected < 0] = 0

                quality_log.append(f"GC-MS: 基线水平: {baseline:.2f}")

                return time, corrected

            def detect_peaks(self, time: np.ndarray, intensity: np.ndarray,
                             height_threshold=1000, distance=5, prominence=500, width=2) -> List[Dict]:
                """检测色谱峰"""
                if len(intensity) == 0:
                    return []

                max_intensity = np.max(intensity)
                min_height = max(max_intensity * 0.001, height_threshold)

                quality_log.append(f"GC-MS: 峰检测参数 - 最小高度: {min_height:.2f}, 最小距离: {distance}")

                try:
                    peaks, properties = find_peaks(
                        intensity,
                        height=min_height,
                        distance=distance,
                        prominence=min_height,
                        width=width
                    )

                    peak_info = []
                    for i, peak_idx in enumerate(peaks):
                        # 评估峰质量
                        peak_height = properties['peak_heights'][i] if 'peak_heights' in properties else intensity[
                            peak_idx]
                        peak_prominence = properties['prominences'][i] if 'prominences' in properties else 0

                        quality_issues = []
                        if peak_height < height_threshold * 2:
                            quality_issues.append("峰高偏低")
                        if peak_prominence < prominence:
                            quality_issues.append("峰突出度不足")

                        peak_info.append({
                            'peak_id': i + 1,
                            'retention_time': time[peak_idx],
                            'intensity': intensity[peak_idx],
                            'peak_index': peak_idx,
                            'peak_height': peak_height,
                            'peak_prominence': peak_prominence,
                            'peak_quality': '; '.join(quality_issues) if quality_issues else 'Good'
                        })

                    quality_log.append(f"GC-MS: 检测到 {len(peak_info)} 个色谱峰")
                    quality_log.append(f"GC-MS: 高质量峰: {sum(1 for p in peak_info if p['peak_quality'] == 'Good')}")

                    return peak_info

                except Exception as e:
                    log_message(f"Error in peak detection: {e}")
                    quality_log.append(f"GC-MS峰检测错误: {str(e)}")
                    return []

            def extract_spectrum_at_rt(self, file_path: str, retention_time: float,
                                       rt_tolerance: float = 0.1) -> np.ndarray:
                """在指定保留时间提取质谱"""
                try:
                    exp = oms.MSExperiment()
                    oms.MzMLFile().load(file_path, exp)

                    target_spectrum = None
                    min_diff = float('inf')

                    for spectrum in exp:
                        if spectrum.getMSLevel() == 1:
                            rt_min = spectrum.getRT() / 60.0
                            diff = abs(rt_min - retention_time)
                            if diff < min_diff and diff <= rt_tolerance:
                                min_diff = diff
                                target_spectrum = spectrum

                    if target_spectrum is not None and target_spectrum.size() > 0:
                        mz_array = []
                        intensity_array = []

                        for peak in target_spectrum:
                            mz = peak.getMZ()
                            intensity_val = peak.getIntensity()
                            if intensity_val > 0:
                                mz_array.append(mz)
                                intensity_array.append(intensity_val)

                        if len(mz_array) > 0:
                            spectrum_data = np.column_stack((mz_array, intensity_array))
                            return spectrum_data

                    return np.array([])

                except Exception as e:
                    log_message(f"Error extracting mass spectrum: {e}")
                    return np.array([])

        class AdvancedSpectrumMatcher:
            """高级谱图匹配器 - 用于GC-MS"""

            def __init__(self, reference_library: List[Dict]):
                self.reference_library = reference_library

            def preprocess_spectrum(self, spectrum: np.ndarray, min_intensity: float = 0.5,
                                    top_peaks: int = 50) -> np.ndarray:
                """预处理谱图用于匹配"""
                if len(spectrum) == 0:
                    return spectrum

                # 去除零强度峰
                spectrum = spectrum[spectrum[:, 1] > 0]
                if len(spectrum) == 0:
                    return spectrum

                # 强度归一化
                max_intensity = np.max(spectrum[:, 1])
                if max_intensity > 0:
                    spectrum[:, 1] = (spectrum[:, 1] / max_intensity) * 100

                # 过滤低强度峰
                spectrum = spectrum[spectrum[:, 1] >= min_intensity]

                # 按强度排序并选择前N个峰
                if len(spectrum) > top_peaks:
                    intensity_order = spectrum[:, 1].argsort()[::-1]
                    spectrum = spectrum[intensity_order[:top_peaks]]

                return spectrum[spectrum[:, 0].argsort()]

            def improved_cosine_similarity(self, spectrum1: np.ndarray, spectrum2: np.ndarray,
                                           tolerance: float = 0.5) -> Tuple[float, int, List, List, List]:
                """计算改进的余弦相似度，返回匹配信息和偏差"""
                if len(spectrum1) == 0 or len(spectrum2) == 0:
                    return 0.0, 0, [], [], []

                # 创建整数m/z索引以简化匹配
                all_mz = set()
                for mz in spectrum1[:, 0]:
                    all_mz.add(int(round(mz)))
                for mz in spectrum2[:, 0]:
                    all_mz.add(int(round(mz)))

                all_mz = sorted(all_mz)
                vec1, vec2 = np.zeros(len(all_mz)), np.zeros(len(all_mz))
                matched_peaks = []  # 存储匹配的峰对 (mz1, mz2, intensity)
                matched_mz_list = []  # 存储匹配的m/z值列表
                matched_errors = []  # 存储匹配碎片的偏差(mDa)

                for i, mz_ref in enumerate(all_mz):
                    # 在谱图1中查找匹配峰
                    matches1 = spectrum1[
                        (spectrum1[:, 0] >= mz_ref - tolerance) & (spectrum1[:, 0] <= mz_ref + tolerance)]
                    if len(matches1) > 0:
                        vec1[i] = np.max(matches1[:, 1])
                        mz1 = matches1[0, 0]
                    else:
                        mz1 = None

                    # 在谱图2中查找匹配峰
                    matches2 = spectrum2[
                        (spectrum2[:, 0] >= mz_ref - tolerance) & (spectrum2[:, 0] <= mz_ref + tolerance)]
                    if len(matches2) > 0:
                        vec2[i] = np.max(matches2[:, 1])
                        mz2 = matches2[0, 0]
                    else:
                        mz2 = None

                    # 如果两个谱图都有峰，记录为匹配
                    if mz1 is not None and mz2 is not None:
                        error_mda = (mz1 - mz2) * 1000  # 转换为mDa
                        matched_peaks.append((mz1, mz2, vec1[i], error_mda))
                        matched_mz_list.append(mz1)
                        matched_errors.append(error_mda)

                dot_product = np.dot(vec1, vec2)
                norm1, norm2 = np.linalg.norm(vec1), np.linalg.norm(vec2)

                if norm1 == 0 or norm2 == 0:
                    return 0.0, 0, [], [], []

                similarity = dot_product / (norm1 * norm2)
                similarity = max(0.0, min(1.0, similarity))

                matched_count = len(matched_peaks)

                return similarity, matched_count, matched_peaks, matched_mz_list, matched_errors

            def match_spectrum(self, query_spectrum: np.ndarray,
                               similarity_threshold: float = 0.7,
                               top_n: int = 1) -> List[Dict]:
                """匹配查询谱图与参考库"""
                if len(query_spectrum) == 0:
                    return []

                query_processed = self.preprocess_spectrum(query_spectrum)
                if len(query_processed) == 0:
                    return []

                results = []
                for ref_compound in self.reference_library:
                    ref_spectrum = ref_compound.get('spectrum', np.array([]))
                    if len(ref_spectrum) == 0:
                        continue

                    ref_processed = self.preprocess_spectrum(ref_spectrum)
                    if len(ref_processed) == 0:
                        continue

                    similarity, matched_count, matched_peaks, matched_mz_list, matched_errors = self.improved_cosine_similarity(
                        query_processed, ref_processed
                    )

                    if similarity >= similarity_threshold:
                        # 格式化匹配的m/z列表和偏差
                        matched_info = []
                        for mz, error in zip(matched_mz_list, matched_errors):
                            matched_info.append(f"{mz:.2f}({error:.1f}mDa)")
                        matched_mz_with_error = "; ".join(matched_info)

                        results.append({
                            'compound': ref_compound.get('name', 'Unknown'),
                            'formula': ref_compound.get('formula', ''),
                            'cas_number': ref_compound.get('cas_number', ''),
                            'similarity': similarity,
                            'matched_peaks_count': matched_count,
                            'total_lib_peaks': len(ref_processed),
                            'matched_peaks_mz': matched_mz_with_error,  # 包含偏差的匹配m/z列表
                            'query_spectrum': query_processed.tolist(),  # 实验谱图数据
                            'reference_spectrum': ref_processed.tolist(),  # 库谱图数据
                            'query_peaks_count': len(query_processed),
                            'ref_peaks_count': len(ref_processed)
                        })

                # 按相似度排序
                results.sort(key=lambda x: x['similarity'], reverse=True)
                return results[:top_n]

        # ==================== LC-MS DDA Processor ====================
        class CosineSimilarityCalculator:
            def __init__(self, config_params):
                self.config = config_params
                self.normalization_method = config_params["NORMALIZATION_METHOD"]

            def calculate_similarity(self, query_peaks: np.ndarray, library_peaks: np.ndarray) -> float:
                if len(query_peaks) == 0 or len(library_peaks) == 0:
                    return 0.0

                query_peaks = self._ensure_float_array(query_peaks)
                library_peaks = self._ensure_float_array(library_peaks)

                if len(query_peaks) == 0 or len(library_peaks) == 0:
                    return 0.0

                lib_peaks_list = [[mz, intensity] for mz, intensity in library_peaks]
                return compute_cosine_similarity(
                    query_peaks,
                    lib_peaks_list,
                    Precision_ref=self.config["MS2_TOL_DA"],
                    Minimum_peak_ref=self.config["MIN_PEAKS"],
                    normalization_method=self.normalization_method
                )

            def calculate_similarity_with_details(self, query_peaks, lib_peaks):
                """计算相似度并返回详细的匹配信息（包含偏差）"""
                query_peaks = np.asarray(query_peaks, dtype=np.float64)
                lib_peaks = np.asarray(lib_peaks, dtype=np.float64)

                query_peaks = normalize_intensity(query_peaks, method=self.normalization_method)
                lib_peaks = normalize_intensity(lib_peaks, method=self.normalization_method)

                if len(query_peaks) == 0 or len(lib_peaks) == 0:
                    return 0.0, None

                matched_peaks = []
                matched_mz_list = []
                matched_errors = []  # 存储匹配碎片的偏差信息(mDa)
                used_query_idx = set()

                aligned_query, aligned_lib = [], []

                for lib_mz, lib_intensity in lib_peaks:
                    diffs = np.abs(query_peaks[:, 0] - lib_mz)
                    valid_idxs = np.where(diffs <= self.config["MS2_TOL_DA"])[0]

                    if len(valid_idxs) > 0:
                        best_idx = valid_idxs[np.argmin(diffs[valid_idxs])]
                        query_intensity = query_peaks[best_idx][1]
                        used_query_idx.add(best_idx)

                        # 计算偏差(mDa)
                        error_mda = (query_peaks[best_idx][0] - lib_mz) * 1000
                        matched_peaks.append((query_peaks[best_idx][0], lib_mz, query_intensity, error_mda))
                        matched_mz_list.append(query_peaks[best_idx][0])
                        matched_errors.append(error_mda)
                        query_intensity_for_aligned = query_intensity
                    else:
                        query_intensity_for_aligned = 0.0

                    aligned_lib.append(lib_intensity)
                    aligned_query.append(query_intensity_for_aligned)

                matched_count = len(matched_peaks)  # 实际匹配到的峰数量

                for i, (mz, intensity) in enumerate(query_peaks):
                    if i not in used_query_idx:
                        aligned_lib.append(0.0)
                        aligned_query.append(intensity)

                query_vector = np.array(aligned_query, dtype=np.float64)
                lib_vector = np.array(aligned_lib, dtype=np.float64)

                if np.linalg.norm(query_vector) == 0 or np.linalg.norm(lib_vector) == 0:
                    return 0.0, None

                similarity = 1 - cosine(query_vector, lib_vector)
                similarity = similarity if not np.isnan(similarity) else 0.0

                # 计算平均偏差
                avg_error_mda = np.mean(matched_errors) if matched_errors else 0

                match_details = {
                    'matched_count': matched_count,
                    'matched_peaks': matched_peaks,
                    'matched_mz_list': matched_mz_list,
                    'matched_errors': matched_errors,
                    'avg_error_mda': avg_error_mda
                }

                return similarity, match_details

            def _ensure_float_array(self, data):
                if not isinstance(data, np.ndarray):
                    data = np.array(data, dtype=np.float64)
                return data[:, :2] if data.shape[1] >= 2 else np.zeros((0, 2), dtype=np.float64)

        class DDAProcessor:
            def __init__(self, config_params):
                self.config = config_params
                self.similarity_calculator = CosineSimilarityCalculator(config_params)

            def load_dda_file(self, mzml_path):
                """加载DDA文件"""
                try:
                    feature_df = perform_feature_detection(
                        mzml_path,
                        noise_threshold=self.config.get("NOISE_THRESHOLD", 1000),
                        remove_single_traces=self.config.get("REMOVE_SINGLE_TRACES", "false")
                    )

                    feature_df_with_ms2 = peak_picking_ms2(
                        feature_df,
                        mzml_path,
                        'DDA',
                        noise_threshold=self.config.get("NOISE_THRESHOLD", 1000),
                        remove_single_traces=self.config.get("REMOVE_SINGLE_TRACES", "false")
                    )

                    matched_features = []
                    for idx, row in feature_df_with_ms2.iterrows():
                        if (row.get('MS2_matched', False) and
                                row['MS2_mz'] is not None and
                                len(row['MS2_mz']) >= self.config.get("MIN_PEAKS", 3)):

                            peaks_array = np.column_stack([row['MS2_mz'], row['MS2_Intensity']])
                            peaks_array = peaks_array.astype(np.float64)

                            if len(peaks_array) > 0:
                                intensity_threshold = max(100, peaks_array[:, 1].max() * 0.01)
                                peaks_array = peaks_array[peaks_array[:, 1] >= intensity_threshold]

                            if len(peaks_array) >= self.config.get("MIN_PEAKS", 3):
                                matched_features.append({
                                    "mz": float(row['m/z']),
                                    "rt": float(row['RT']),
                                    "intensity": float(row['Intensity']),
                                    "charge": 1,
                                    "ms2_spectrum": {
                                        'peaks': peaks_array,
                                        'precursor_mz': float(row['Precursor_mz'])
                                    },
                                    "precursor_mz": float(row['Precursor_mz'])
                                })

                    log_message(f"Found {len(matched_features)} valid features with MS2 spectra")
                    return pd.DataFrame(matched_features), None

                except Exception as e:
                    log_message(f"Error processing DDA file: {e}")
                    import traceback
                    log_message(traceback.format_exc())
                    return pd.DataFrame(), None

            def annotate_features(self, feature_df, database):
                """特征注释 - 包含详细的匹配信息和质量日志"""
                annotations = []
                if feature_df.empty:
                    log_message("No features to annotate")
                    return pd.DataFrame()

                min_peaks_required = self.config["MIN_PEAKS"]
                log_message(f"Annotating {len(feature_df)} features using database with {len(database)} spectra")
                log_message(f"Minimum matched peaks required: {min_peaks_required}")

                # 预处理数据库
                valid_db_spectra = []
                for spec in database:
                    try:
                        precursor_mz = float(spec.get("precursor_mz", 0))
                        if precursor_mz <= 0:
                            continue

                        if hasattr(spec, 'peaks'):
                            peaks = np.array(spec.peaks, dtype=np.float64)
                        elif hasattr(spec, 'mz') and hasattr(spec, 'intensities'):
                            peaks = np.column_stack([spec.mz, spec.intensities])
                        else:
                            continue

                        if len(peaks) >= min_peaks_required:
                            valid_db_spectra.append({
                                'spec': spec,
                                'precursor_mz': precursor_mz,
                                'peaks': peaks,
                                'metadata': spec.metadata if hasattr(spec, 'metadata') else {}
                            })
                    except Exception as e:
                        log_message(f"Error processing library spectrum: {e}")
                        continue

                log_message(f"Valid database spectra after preprocessing: {len(valid_db_spectra)}")

                # 匹配统计
                total_matches = 0
                low_score_matches = 0
                matches_below_min_peaks = 0
                matches_with_1_peak = 0

                for idx, row in feature_df.iterrows():
                    if row["ms2_spectrum"] is None:
                        continue

                    query_peaks = row["ms2_spectrum"]['peaks']
                    query_precursor = row["precursor_mz"]

                    if len(query_peaks) < min_peaks_required:
                        continue

                    best_match = None
                    best_score = 0.0
                    best_match_details = None
                    best_lib_peaks = None

                    for db_item in valid_db_spectra:
                        try:
                            if abs(db_item['precursor_mz'] - query_precursor) > self.config["MS1_TOL_DA"]:
                                continue

                            lib_peaks = db_item['peaks']
                            if len(lib_peaks) < min_peaks_required:
                                continue

                            cosine_score, matched_peaks_info = self.similarity_calculator.calculate_similarity_with_details(
                                query_peaks, lib_peaks
                            )

                            if cosine_score > best_score:
                                best_score = cosine_score
                                best_match = db_item
                                best_match_details = matched_peaks_info
                                best_lib_peaks = lib_peaks

                        except Exception as e:
                            log_message(f"Error calculating similarity: {e}")
                            continue

                    if best_match is not None and best_match_details is not None:
                        # 检查匹配的峰数量是否达到要求
                        matched_count = best_match_details['matched_count']

                        # 关键修复：如果匹配峰数量小于最小要求，跳过这个匹配
                        if matched_count < min_peaks_required:
                            matches_below_min_peaks += 1
                            if matched_count == 1:
                                matches_with_1_peak += 1
                            continue  # 不保存这个匹配

                        lib_spec = best_match['spec']
                        lib_metadata = best_match['metadata']

                        # 计算一级偏差
                        precursor_error_da = abs(query_precursor - best_match['precursor_mz'])
                        precursor_error_ppm = (precursor_error_da / best_match['precursor_mz']) * 1e6 if best_match[
                                                                                                             'precursor_mz'] > 0 else 0

                        # 格式化匹配的二级碎片信息（包含偏差）
                        matched_info = []
                        for i, mz in enumerate(best_match_details['matched_mz_list']):
                            error_mda = best_match_details['matched_errors'][i]
                            matched_info.append(f"{mz:.4f}({error_mda:.1f}mDa)")
                        matched_peaks_with_error = "; ".join(matched_info)

                        # 计算平均偏差
                        avg_error_mda = best_match_details.get('avg_error_mda', 0)

                        annotation = {
                            # 基本特征信息
                            "mz": float(row["mz"]),
                            "rt": float(row["rt"]),
                            "intensity": float(row["intensity"]),
                            "precursor_mz": float(query_precursor),
                            "charge": int(row["charge"]),

                            # 匹配结果
                            "cosine_score": float(best_score),
                            "is_valid_match": best_score >= self.config["DP_THRESHOLD"],  # 根据阈值判断是否有效

                            # 一级信息偏差
                            "precursor_error_da": round(precursor_error_da, 5),
                            "precursor_error_ppm": round(precursor_error_ppm, 2),
                            "library_precursor_mz": round(best_match['precursor_mz'], 5),

                            # 二级碎片信息（删除百分比，添加偏差信息）
                            "exp_peaks_count": len(query_peaks),
                            "lib_peaks_count": len(best_lib_peaks) if best_lib_peaks is not None else 0,
                            "matched_peaks_count": matched_count,
                            "matched_peaks_avg_error_mda": round(avg_error_mda, 2),
                            "matched_peaks_mz": matched_peaks_with_error,  # 包含偏差的匹配m/z值

                            # 谱图数据（用于镜像图）- 保存为字符串，方便Excel存储
                            "exp_spectrum_peaks": str(query_peaks.tolist()),  # 转换为字符串存储
                            "lib_spectrum_peaks": str(best_lib_peaks.tolist()) if best_lib_peaks is not None else "",

                            # 化合物信息
                            "Name": str(lib_metadata.get('Name', lib_metadata.get('compound_name', 'N/A'))),
                            "InChIKey": str(lib_metadata.get('InChIKey', lib_metadata.get('inchikey', 'N/A'))),
                            "Formula": str(lib_metadata.get('Formula', lib_metadata.get('formula', 'N/A'))),
                            "MolecularWeight": str(
                                lib_metadata.get('MolecularWeight', lib_metadata.get('molecularweight', 'N/A'))),
                            "ExactMass": str(lib_metadata.get('ExactMass', lib_metadata.get('exactmass', 'N/A'))),
                            "LibrarySource": str(
                                lib_metadata.get('LibrarySource', lib_metadata.get('instrument_type', 'N/A'))),
                            "Comments": str(lib_metadata.get('Comments', lib_metadata.get('comment', 'N/A'))),

                            # 样品信息
                            "sample": row.get('source_file', 'Unknown')
                        }
                        annotations.append(annotation)
                        total_matches += 1
                        if best_score < self.config["DP_THRESHOLD"]:
                            low_score_matches += 1

                # 记录匹配质量日志
                quality_log.append(f"LC-MS匹配质量统计:")
                quality_log.append(f"  总匹配数: {total_matches}")
                quality_log.append(f"  匹配峰数量低于要求({min_peaks_required})的匹配: {matches_below_min_peaks}")
                quality_log.append(f"  其中匹配峰数量为1的匹配: {matches_with_1_peak}")
                quality_log.append(f"  低分数匹配: {low_score_matches}")

                result_df = pd.DataFrame(annotations)
                if not result_df.empty:
                    result_df = result_df.sort_values('cosine_score', ascending=False)
                    log_message(f"Annotation completed: {len(result_df)} matches found")

                    # 显示匹配峰数量的详细统计
                    matched_counts = result_df['matched_peaks_count'].value_counts().sort_index()
                    log_message(f"匹配峰数量分布: {dict(matched_counts)}")

                    # 验证没有匹配峰数量小于最小要求的
                    min_in_result = result_df['matched_peaks_count'].min()
                    if min_in_result < min_peaks_required:
                        log_message(
                            f"警告: 结果中存在匹配峰数量为 {min_in_result} 的匹配，低于最小要求 {min_peaks_required}")
                else:
                    log_message("No annotations found")

                return result_df

        # ==================== 通用MSP加载函数 ====================
        def custom_load_msp(filename):
            """通用MSP文件解析器（用于LC-MS）"""
            spectra = []
            current_metadata = {}
            peaks = []
            peak_count = 0
            skipped_spectra = 0
            valid_spectra = 0

            def get_field(*field_names):
                for name in field_names:
                    lower_name = name.lower()
                    for key in current_metadata:
                        if key.lower() == lower_name:
                            val = current_metadata[key]
                            return str(val) if not isinstance(val, (int, float)) else val
                    for key in current_metadata:
                        key_lower = key.lower()
                        for field_name in field_names:
                            if field_name.lower() in key_lower:
                                val = current_metadata[key]
                                return str(val) if not isinstance(val, (int, float)) else val
                return ""

            try:
                with open(filename, 'r', encoding='utf-8', errors='ignore') as file:
                    lines = file.readlines()
            except UnicodeDecodeError:
                with open(filename, 'r', encoding='gbk', errors='ignore') as file:
                    lines = file.readlines()

            i = 0
            while i < len(lines):
                line = lines[i].strip()

                if not line:
                    i += 1
                    continue

                if line.lower().startswith('name:') or (
                        ':' not in line and not re.match(r'^\d+\.?\d*\s+\d+\.?\d*', line)):
                    if current_metadata and peaks:
                        structured_metadata = build_structured_metadata(current_metadata, peak_count)
                        spectrum = create_spectrum_from_data(structured_metadata, peaks)
                        if spectrum:
                            spectra.append(spectrum)
                            valid_spectra += 1
                        else:
                            skipped_spectra += 1

                    current_metadata = {}
                    peaks = []
                    peak_count = 0

                    if line.lower().startswith('name:'):
                        current_metadata['name'] = line[5:].strip()
                    else:
                        current_metadata['name'] = line.strip()
                elif ':' in line:
                    key, value = line.split(':', 1)
                    key, value = key.strip(), value.strip()
                    current_metadata[key] = value
                elif re.match(r'^\d+\.?\d*\s+\d+\.?\d*', line):
                    parts = re.split(r'\s+', line)
                    if len(parts) >= 2:
                        try:
                            mz = float(parts[0])
                            intensity = float(parts[1])
                            if mz > 0 and intensity >= 0:
                                peaks.append((mz, intensity))
                                peak_count += 1
                        except ValueError:
                            pass
                i += 1

            if current_metadata and peaks:
                structured_metadata = build_structured_metadata(current_metadata, peak_count)
                spectrum = create_spectrum_from_data(structured_metadata, peaks)
                if spectrum:
                    spectra.append(spectrum)
                    valid_spectra += 1
                else:
                    skipped_spectra += 1

            quality_log.append(f"MSP库加载统计:")
            quality_log.append(f"  总谱图数: {len(spectra) + skipped_spectra}")
            quality_log.append(f"  有效谱图: {valid_spectra}")
            quality_log.append(f"  跳过谱图: {skipped_spectra}")

            return spectra

        def build_structured_metadata(metadata, peak_count):
            """构建结构化的元数据"""

            def get_field(*field_names):
                for name in field_names:
                    lower_name = name.lower()
                    for key in metadata:
                        if key.lower() == lower_name:
                            val = metadata[key]
                            if name.lower() in ['precursormz', 'precursorm/z', 'parentmass', 'exactmass', 'mw',
                                                'molecularweight', 'retentiontime', 'rt']:
                                try:
                                    num_str = re.search(r'[-+]?\d*\.\d+|\d+', str(val))
                                    if num_str:
                                        return float(num_str.group())
                                except:
                                    return 0.0
                            return str(val) if not isinstance(val, (int, float)) else val
                    for key in metadata:
                        key_lower = key.lower()
                        for field_name in field_names:
                            if field_name.lower() in key_lower:
                                val = metadata[key]
                                if name.lower() in ['precursormz', 'precursorm/z', 'parentmass', 'exactmass', 'mw',
                                                    'molecularweight', 'retentiontime', 'rt']:
                                    try:
                                        num_str = re.search(r'[-+]?\d*\.\d+|\d+', str(val))
                                        if num_str:
                                            return float(num_str.group())
                                    except:
                                        return 0.0
                                return str(val) if not isinstance(val, (int, float)) else val
                return ""

            instrument_type = get_field('Instrument', 'InstrumentType', 'Instrument_type', 'Instrumentation')
            if not instrument_type:
                comments = metadata.get('comments', '') or metadata.get('Comments', '') or metadata.get('note',
                                                                                                        '') or metadata.get(
                    'Note', '')
                comments_lower = str(comments).lower()
                if 'qtof' in comments_lower:
                    instrument_type = 'QTOF'
                elif 'orbitrap' in comments_lower:
                    instrument_type = 'Orbitrap'
                elif 'qqq' in comments_lower or 'triple quad' in comments_lower:
                    instrument_type = 'Triple Quadrupole'
                elif 'gc-ms' in comments_lower or 'gc/ms' in comments_lower:
                    instrument_type = 'GC-MS'
                elif 'lc-ms' in comments_lower or 'lc/ms' in comments_lower:
                    instrument_type = 'LC-MS'

            ion_mode = get_field('Ion_mode', 'IonMode', 'Ionization', 'Precursor_type', 'Polarity')
            if not ion_mode:
                name_str = str(metadata.get('name', '')).lower()
                comments_str = str(metadata.get('comments', '')).lower()
                if 'positive' in name_str or 'positive' in comments_str or '[m+h]' in name_str or '[m+h]+' in name_str:
                    ion_mode = 'Positive'
                elif 'negative' in name_str or 'negative' in comments_str or '[m-h]' in name_str or '[m-h]-' in name_str:
                    ion_mode = 'Negative'

            structured_metadata = {
                'Name': get_field('Name', 'Compound', 'COMPOUND', 'Title', 'Synon'),
                'InChIKey': get_field('InChIKey', 'INCHIKEY', 'InChI Key'),
                'InChI': get_field('InChI', 'INCHI'),
                'SMILES': get_field('SMILES', 'Smiles'),
                'SpectrumType': get_field('Spectrum_type', 'SpectrumType', 'Type', 'Ionization'),
                'InstrumentType': instrument_type,
                'IonMode': ion_mode,
                'Formula': get_field('Formula', 'FORMULA', 'MolecularFormula'),
                'MolecularWeight': get_field('MW', 'MolecularWeight', 'Mass', 'MOLECULAR_WEIGHT'),
                'ExactMass': get_field('ExactMass', 'Exact_Mass', 'PrecursorMZ', 'ParentMass'),
                'RetentionTime': get_field('RetentionTime', 'RT', 'Retention_Time'),
                'PrecursorMZ': get_field('PrecursorMZ', 'Precursor', 'ParentMass', 'PRECURSORMZ', 'PrecursorM/Z'),
                'PeakCount': str(peak_count),
                'LibrarySource': get_field('Library', 'LIBRARY', 'Source', 'Database'),
                'Comments': get_field('Comments', 'Note', 'Remarks'),
            }

            return structured_metadata

        def create_spectrum_from_data(metadata, peaks):
            """从数据创建Spectrum对象"""
            try:
                precursor_mz = 0.0
                precursor_str = metadata.get('PrecursorMZ', '')

                if precursor_str and str(precursor_str) != '0.0':
                    try:
                        precursor_mz = float(precursor_str)
                    except (ValueError, TypeError):
                        exact_mass = metadata.get('ExactMass', 0.0)
                        if exact_mass and exact_mass != 0.0:
                            precursor_mz = float(exact_mass)

                if precursor_mz == 0.0:
                    name_str = str(metadata.get('Name', ''))
                    mass_match = re.search(r'\[M[^\]]+\]\s*([\d.]+)', name_str)
                    if mass_match:
                        try:
                            precursor_mz = float(mass_match.group(1))
                        except (ValueError, TypeError):
                            pass

                peaks_array = []
                for mz, intensity in peaks:
                    try:
                        mz_float = float(mz) if mz else 0.0
                        intensity_float = float(intensity) if intensity else 0.0
                        if mz_float > 0 and intensity_float >= 0:
                            peaks_array.append([mz_float, intensity_float])
                    except (ValueError, TypeError):
                        continue

                if not peaks_array:
                    return None

                peaks_array = np.array(peaks_array, dtype=np.float64)
                peaks_array = peaks_array[peaks_array[:, 0].argsort()]

                spectrum = Spectrum(
                    mz=peaks_array[:, 0],
                    intensities=peaks_array[:, 1],
                    metadata=metadata
                )
                spectrum.set('precursor_mz', precursor_mz)

                return spectrum

            except Exception as e:
                log_message(f"Error creating spectrum: {e}")
                return None

        # ==================== 保存结果函数 ====================
        def save_quality_log(file_path):
            """保存质量日志到文件"""
            log_path = os.path.join(os.path.dirname(file_path), "quality_log.txt")
            with open(log_path, 'w', encoding='utf-8') as f:
                f.write("=" * 60 + "\n")
                f.write("质量日志报告\n")
                f.write("=" * 60 + "\n\n")
                for log_entry in quality_log:
                    f.write(log_entry + "\n")
            return log_path

        # ==================== GUI Functions ====================
        def log_message(message):
            """添加消息到进度区域"""
            progress_text.insert(tk.END, message + "\n")
            progress_text.see(tk.END)
            analysis_window.update()
            quality_log.append(message)  # 同时添加到质量日志

        def browse_file(entry_widget, filetypes=None):
            """浏览单个文件"""
            try:
                if filetypes is None:
                    filetypes = [("mzML files", "*.mzML"), ("All files", "*.*")]

                file_path = filedialog.askopenfilename(filetypes=filetypes)
                if file_path:
                    entry_widget.delete(0, tk.END)
                    if isinstance(entry_widget, ttk.Combobox):
                        entry_widget.set(file_path)
                    else:
                        entry_widget.insert(0, file_path)
            except Exception as e:
                messagebox.showerror("Error", f"File selection error: {str(e)}")

        def browse_folder():
            """浏览文件夹进行批处理"""
            nonlocal batch_files
            folder_path = filedialog.askdirectory(title="Select folder containing mzML files")
            if folder_path:
                mzml_files = []
                for ext in ["*.mzML", "*.mzml"]:
                    files = glob.glob(os.path.join(folder_path, ext))
                    for file in files:
                        normalized_path = os.path.normcase(os.path.normpath(file))
                        if normalized_path not in [os.path.normcase(os.path.normpath(f)) for f in mzml_files]:
                            mzml_files.append(file)

                if not mzml_files:
                    messagebox.showwarning("Warning", "No mzML files found in folder")
                    return

                batch_files = sorted(mzml_files)
                batch_count_label.config(text=f"Selected {len(batch_files)} files")
                log_message(f"Selected {len(batch_files)} mzML files for batch processing")

                if batch_files:
                    exp_entry.delete(0, tk.END)
                    exp_entry.insert(0, batch_files[0])

        def update_parameter_display():
            """更新参数显示"""
            if current_method.get() == "LC-MS":
                config_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
                gc_config_frame.pack_forget()
                update_result_tree_columns("LC-MS")
            else:
                config_frame.pack_forget()
                gc_config_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
                update_result_tree_columns("GC-MS")

        def update_result_tree_columns(method):
            """更新结果表格列"""
            for col in result_tree.get_children():
                result_tree.delete(col)

            if method == "LC-MS":
                columns = ("Precursor m/z", "RT", "Name", "Score", "Matched Peaks", "Error(ppm)", "Sample")
                col_widths = [100, 70, 200, 70, 90, 70, 120]
            else:
                columns = ("Peak ID", "RT (min)", "Compound", "Formula", "Similarity", "Matched Peaks", "Sample")
                col_widths = [70, 80, 200, 100, 80, 80, 120]

            for col in result_tree['columns']:
                result_tree.heading(col, text="")

            result_tree['columns'] = columns

            for col, width in zip(columns, col_widths):
                result_tree.heading(col, text=col)
                result_tree.column(col, width=width, anchor="center")

        def run_analysis():
            """运行分析"""
            nonlocal processing, result_data, current_file_index, batch_files, quality_log

            if processing:
                messagebox.showwarning("Warning", "Analysis is already running")
                return

            # 清空质量日志
            quality_log = []

            # 获取参数
            lib_path = lib_combobox.get()

            if not lib_path:
                messagebox.showerror("Error", "Please select spectral library")
                return

            # 确定要处理的文件
            if batch_files:
                files_to_process = batch_files
            else:
                exp_path = exp_entry.get()
                if not exp_path:
                    messagebox.showerror("Error", "Please select experimental data")
                    return
                files_to_process = [exp_path]

            try:
                if current_method.get() == "LC-MS":
                    config_params = {
                        "MS1_TOL_DA": float(entries["MS1_TOL_DA"].get()),
                        "MS2_TOL_DA": float(entries["MS2_TOL_DA"].get()),
                        "MIN_INTENSITY": float(entries["MIN_INTENSITY"].get()),
                        "DP_THRESHOLD": float(entries["DP_THRESHOLD"].get()),
                        "MIN_PEAKS": int(entries["MIN_PEAKS"].get()),
                        "NORMALIZATION_METHOD": entries["NORMALIZATION_METHOD"].get(),
                        "NOISE_THRESHOLD": float(entries["NOISE_THRESHOLD"].get()),
                        "REMOVE_SINGLE_TRACES": entries["REMOVE_SINGLE_TRACES"].get()
                    }
                else:
                    config_params = {
                        "GC_PEAK_HEIGHT": float(gc_entries["GC_PEAK_HEIGHT"].get()),
                        "GC_PEAK_DISTANCE": float(gc_entries["GC_PEAK_DISTANCE"].get()),
                        "GC_PEAK_PROMINENCE": float(gc_entries["GC_PEAK_PROMINENCE"].get()),
                        "GC_PEAK_WIDTH": float(gc_entries["GC_PEAK_WIDTH"].get()),
                        "GC_RT_TOLERANCE": float(gc_entries["GC_RT_TOLERANCE"].get()),
                        "GC_SIMILARITY_THRESHOLD": float(gc_entries["GC_SIMILARITY_THRESHOLD"].get())
                    }
            except ValueError as e:
                messagebox.showerror("Error", f"Invalid parameter value: {str(e)}")
                return

            # 清空结果
            for item in result_tree.get_children():
                result_tree.delete(item)
            progress_text.delete(1.0, tk.END)

            processing = True
            run_btn.config(text="Processing...", state="disabled")
            current_file_index = 0

            def process_lcms_file(file_path, all_results):
                """处理单个LC-MS文件"""
                log_message(f"\n=== Processing LC-MS file: {os.path.basename(file_path)} ===")

                # 加载数据库
                database = custom_load_msp(lib_path)
                valid_spectra = [s for s in database if s is not None]
                log_message(f"Successfully loaded {len(valid_spectra)} valid spectra")

                # 处理mzML文件
                processor = DDAProcessor(config_params)
                feature_df, run = processor.load_dda_file(file_path)

                if feature_df.empty:
                    log_message("Cannot extract features from mzML file")
                    return all_results

                log_message(f"Matched {len(feature_df)} features with MS2 spectra")

                file_results = None
                if len(feature_df) > 0:
                    annotations = processor.annotate_features(feature_df, database)

                    if not annotations.empty:
                        # 添加文件名信息
                        annotations["sample"] = os.path.basename(file_path)
                        file_results = annotations.copy()

                        # 合并结果
                        if all_results is None:
                            all_results = annotations
                        else:
                            all_results = pd.concat([all_results, annotations], ignore_index=True)

                        # 显示当前文件结果
                        file_valid_matches = annotations[
                            'is_valid_match'].sum() if 'is_valid_match' in annotations.columns else 0
                        log_message(f"Found {len(annotations)} matches, {file_valid_matches} valid matches")

                        # 自动保存结果（根据generate_mirror_var决定是否生成镜像图）
                        save_dir = os.path.dirname(file_path)
                        safe_filename = re.sub(r'[\\/*?:"<>|]', "_", os.path.splitext(os.path.basename(file_path))[0])
                        excel_filename = f"{safe_filename}_LC-MS_results.xlsx"
                        excel_path = os.path.join(save_dir, excel_filename)

                        save_comprehensive_results(file_results, excel_path, generate_mirror_var.get(), "LC-MS")
                    else:
                        log_message("No matches found")
                else:
                    log_message("No features with MS2 spectra found")

                return all_results

            def process_gcms_file(file_path, all_results):
                """处理单个GC-MS文件"""
                log_message(f"\n=== Processing GC-MS file: {os.path.basename(file_path)} ===")

                # 加载MSP库
                msp_parser = MSPParser()
                reference_library = msp_parser.parse_msp_file(lib_path)

                if not reference_library:
                    log_message("Warning: Spectral library is empty or parsing failed")
                    return all_results

                log_message(f"Successfully loaded {len(reference_library)} reference compounds")

                # 处理GC-MS数据
                gcms_processor = GCMSProcessor()
                rt, tic = gcms_processor.extract_chromatogram(file_path)

                if len(rt) == 0:
                    log_message("Cannot read mzML file data")
                    return all_results

                rt_processed, tic_processed = gcms_processor.preprocess_chromatogram(rt, tic)

                peaks = gcms_processor.detect_peaks(
                    rt_processed, tic_processed,
                    height_threshold=config_params["GC_PEAK_HEIGHT"],
                    distance=config_params["GC_PEAK_DISTANCE"],
                    prominence=config_params["GC_PEAK_PROMINENCE"],
                    width=config_params["GC_PEAK_WIDTH"]
                )

                if not peaks:
                    log_message("No chromatographic peaks detected")
                    return all_results

                log_message(f"Detected {len(peaks)} chromatographic peaks for analysis")

                # 谱图匹配
                matcher = AdvancedSpectrumMatcher(reference_library)
                file_results = []
                total_peaks = len(peaks)

                log_message(f"Starting spectrum matching for {total_peaks} peaks...")

                for i, peak in enumerate(peaks):
                    if i % 10 == 0:
                        log_message(f"  Progress: {i}/{total_peaks} ({i / total_peaks * 100:.1f}%)")

                    # 提取质谱
                    spectrum = gcms_processor.extract_spectrum_at_rt(
                        file_path, peak['retention_time'],
                        rt_tolerance=config_params["GC_RT_TOLERANCE"]
                    )

                    if len(spectrum) > 0:
                        # 谱图匹配
                        matches = matcher.match_spectrum(
                            spectrum,
                            similarity_threshold=config_params["GC_SIMILARITY_THRESHOLD"],
                            top_n=1
                        )

                        if matches:
                            best_match = matches[0]
                            result = {
                                'peak_id': peak['peak_id'],
                                'retention_time': round(peak['retention_time'], 3),
                                'intensity': int(peak['intensity']),
                                'peak_quality': peak['peak_quality'],
                                'compound': best_match['compound'],
                                'formula': best_match.get('formula', ''),
                                'cas_number': best_match.get('cas_number', ''),
                                'similarity': round(best_match['similarity'], 4),
                                'matched_peaks_count': best_match['matched_peaks_count'],
                                'total_lib_peaks': best_match['total_lib_peaks'],
                                'matched_peaks_mz': best_match.get('matched_peaks_mz', ''),
                                'sample': os.path.basename(file_path),
                                'query_spectrum': str(best_match.get('query_spectrum', [])),
                                'reference_spectrum': str(best_match.get('reference_spectrum', []))
                            }
                            file_results.append(result)

                            log_message(f"  ✅ Peak {peak['peak_id']} (RT: {peak['retention_time']:.2f}min): "
                                        f"{best_match['compound']} - Similarity: {best_match['similarity']:.3f}")

                # 转换为DataFrame
                if file_results:
                    file_df = pd.DataFrame(file_results)

                    # 记录质量统计
                    quality_log.append(f"GC-MS匹配质量统计:")
                    quality_log.append(f"  总匹配数: {len(file_df)}")
                    quality_log.append(f"  平均相似度: {file_df['similarity'].mean():.3f}")
                    quality_log.append(f"  最高相似度: {file_df['similarity'].max():.3f}")

                    if all_results is None:
                        all_results = file_df
                    else:
                        all_results = pd.concat([all_results, file_df], ignore_index=True)

                    log_message(f"GC-MS analysis completed: {len(file_df)} high-confidence matches found")

                    # 自动保存结果（根据generate_mirror_var决定是否生成镜像图）
                    save_dir = os.path.dirname(file_path)
                    safe_filename = re.sub(r'[\\/*?:"<>|]', "_", os.path.splitext(os.path.basename(file_path))[0])
                    excel_filename = f"{safe_filename}_GC-MS_results.xlsx"
                    excel_path = os.path.join(save_dir, excel_filename)

                    save_comprehensive_results(file_df, excel_path, generate_mirror_var.get(), "GC-MS")

                else:
                    log_message("GC-MS analysis completed: No high-confidence matches found")

                return all_results

            def analysis_thread():
                nonlocal result_data, current_file_index
                all_results = None

                try:
                    # 依次处理每个文件
                    for i, file_path in enumerate(files_to_process):
                        current_file_index = i
                        log_message(f"\nProcessing progress: {i + 1}/{len(files_to_process)}")

                        if current_method.get() == "LC-MS":
                            all_results = process_lcms_file(file_path, all_results)
                        else:
                            all_results = process_gcms_file(file_path, all_results)

                    # 所有文件处理完成
                    if all_results is not None and not all_results.empty:
                        if current_method.get() == "LC-MS":
                            all_results = all_results.sort_values('precursor_mz', ascending=True)
                        else:
                            all_results = all_results.sort_values('retention_time', ascending=True)

                        result_data = all_results.copy()

                        # 在表格中显示所有结果
                        for _, row in all_results.iterrows():
                            if current_method.get() == "LC-MS":
                                result_tree.insert("", tk.END, values=(
                                    f"{row['precursor_mz']:.4f}" if 'precursor_mz' in row else "N/A",
                                    f"{row['rt']:.2f}" if 'rt' in row else "N/A",
                                    row.get("Name", "N/A")[:30],
                                    f"{row['cosine_score']:.3f}" if 'cosine_score' in row else "N/A",
                                    row.get("matched_peaks_count", "N/A"),
                                    f"{row['precursor_error_ppm']:.1f}" if 'precursor_error_ppm' in row else "N/A",
                                    row.get("sample", "N/A")
                                ))
                            else:
                                result_tree.insert("", tk.END, values=(
                                    row.get('peak_id', "N/A"),
                                    f"{row['retention_time']:.2f}" if 'retention_time' in row else "N/A",
                                    row.get("compound", "N/A")[:30],
                                    row.get("formula", "N/A"),
                                    f"{row['similarity']:.3f}" if 'similarity' in row else "N/A",
                                    row.get("matched_peaks_count", "N/A"),
                                    row.get("sample", "N/A")
                                ))

                        total_matches = len(all_results)
                        log_message(f"\n=== Batch processing completed ===")
                        log_message(f"Total files processed: {len(files_to_process)}")
                        log_message(f"Total matches: {total_matches}")

                        # 保存批处理汇总结果
                        if len(files_to_process) > 1:
                            save_dir = os.path.dirname(files_to_process[0])
                            summary_path = os.path.join(save_dir, f"batch_summary_{current_method.get()}_results.xlsx")
                            save_comprehensive_results(all_results, summary_path, generate_mirror_var.get(),
                                                       current_method.get())
                    else:
                        log_message("\nNo matches found in all files")

                except Exception as e:
                    log_message(f"Error during analysis: {str(e)}")
                    import traceback
                    log_message(traceback.format_exc())
                finally:
                    processing = False
                    analysis_window.after(0, lambda: run_btn.config(text="▶ START ANALYSIS", state="normal"))

            # 在新线程中运行分析
            threading.Thread(target=analysis_thread, daemon=True).start()

        def download_results():
            """下载汇总结果"""
            nonlocal result_data
            if result_data is None or result_data.empty:
                messagebox.showinfo("Information", "No results to download")
                return

            try:
                save_path = filedialog.asksaveasfilename(
                    defaultextension=".xlsx",
                    filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
                    title="Save consolidated results as"
                )

                if not save_path:
                    return

                # 保存汇总结果（根据generate_mirror_var决定是否生成镜像图）
                save_comprehensive_results(result_data, save_path, generate_mirror_var.get(), current_method.get())

                messagebox.showinfo("Success", f"Consolidated results saved as: {os.path.basename(save_path)}")

            except Exception as e:
                messagebox.showerror("Error", f"Save failed: {str(e)}")

        # ==================== GUI Interface Creation ====================
        # 创建主框架
        main_frame = ttk.Frame(analysis_window)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=15)

        # 创建顶部水平框架
        top_frame = ttk.Frame(main_frame)
        top_frame.pack(fill=tk.X, pady=(0, 10))

        # 左侧：方法选择和文件选择
        left_frame = ttk.Frame(top_frame)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # 方法选择区域
        method_frame = ttk.LabelFrame(left_frame, text="Selection of Analytical Methods")
        method_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Label(method_frame, text="Analytical Method:").grid(row=0, column=0, padx=10, pady=8, sticky=tk.W)
        method_combo = ttk.Combobox(method_frame, textvariable=current_method,
                                    values=["LC-MS", "GC-MS"], state="readonly", width=15)
        method_combo.grid(row=0, column=1, padx=10, pady=8, sticky=tk.W)
        method_combo.bind('<<ComboboxSelected>>', lambda e: update_parameter_display())

        # 文件选择区域
        file_frame = ttk.LabelFrame(left_frame, text="File Selection")
        file_frame.pack(fill=tk.X, pady=(0, 0))

        # 实验数据选择
        ttk.Label(file_frame, text="Experimental Data:").grid(row=0, column=0, padx=10, pady=8, sticky=tk.W)
        exp_entry = ttk.Entry(file_frame, width=50)
        exp_entry.grid(row=0, column=1, padx=10, pady=8, sticky=tk.W + tk.E)
        ttk.Button(file_frame, text="Browse Files", command=lambda: browse_file(exp_entry)).grid(row=0, column=2,
                                                                                                 padx=10, pady=8)

        # 批处理选择
        ttk.Button(file_frame, text="Batch Select Folder", command=browse_folder).grid(row=1, column=0, padx=10, pady=8)
        batch_count_label = ttk.Label(file_frame, text="No files selected", foreground="blue")
        batch_count_label.grid(row=1, column=1, padx=10, pady=8, sticky=tk.W)

        # 谱库选择
        ttk.Label(file_frame, text="Spectral Library File:").grid(row=2, column=0, padx=10, pady=8, sticky=tk.W)
        lib_combobox = ttk.Combobox(file_frame, width=50)
        lib_combobox.grid(row=2, column=1, padx=10, pady=8, sticky=tk.W + tk.E)
        ttk.Button(file_frame, text="Browse Library",
                   command=lambda: browse_file(lib_combobox, [("MSP files", "*.msp"), ("All files", "*.*")])).grid(
            row=2, column=2, padx=10, pady=8)

        file_frame.columnconfigure(1, weight=1)

        # 右侧：参数配置区域
        right_frame = ttk.Frame(top_frame)
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, padx=(20, 0))

        # LC-MS参数配置区域
        config_frame = ttk.LabelFrame(right_frame, text="LC-MS Analysis Parameter Configuration")

        entries = {}
        lc_param_grid = tk.Frame(config_frame)
        lc_param_grid.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        lc_params = [
            ("MS1_TOL_DA", "MS1 Tolerance (Da):", "0.005"),  # 优化默认值
            ("MS2_TOL_DA", "MS2 Tolerance (Da):", "0.01"),  # 优化默认值
            ("MIN_INTENSITY", "Min Intensity:", "1000"),  # 降低默认值
            ("DP_THRESHOLD", "Score Threshold:", "0.6"),
            ("MIN_PEAKS", "Min Matched Peaks:", "3"),  # 增加到3
            ("NORMALIZATION_METHOD", "Normalization Method:", "max"),
            ("NOISE_THRESHOLD", "Noise Threshold:", "500"),  # 降低默认值
            ("REMOVE_SINGLE_TRACES", "Remove Single Traces:", "false")
        ]

        for i, (param, label, default) in enumerate(lc_params):
            ttk.Label(lc_param_grid, text=label, width=20).grid(row=i, column=0, padx=5, pady=3, sticky=tk.W)
            if param == "NORMALIZATION_METHOD":
                var = tk.StringVar(value=default)
                entry = ttk.Combobox(lc_param_grid, textvariable=var, values=["max", "sqrt", "sum"], width=15)
                entries[param] = var
            elif param == "REMOVE_SINGLE_TRACES":
                var = tk.StringVar(value=default)
                entry = ttk.Combobox(lc_param_grid, textvariable=var, values=["true", "false"], width=15)
                entries[param] = var
            else:
                var = tk.StringVar(value=default)
                entry = ttk.Entry(lc_param_grid, textvariable=var, width=15)
                entries[param] = var
            entry.grid(row=i, column=1, padx=5, pady=3, sticky=tk.W)

        # GC-MS参数配置区域
        gc_config_frame = ttk.LabelFrame(right_frame, text="GC-MS Analysis Parameter Configuration")

        gc_entries = {}
        gc_param_grid = tk.Frame(gc_config_frame)
        gc_param_grid.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        gc_params = [
            ("GC_PEAK_HEIGHT", "Peak Height Threshold:", "1000"),
            ("GC_PEAK_DISTANCE", "Peak Distance:", "5"),
            ("GC_PEAK_PROMINENCE", "Peak Prominence:", "500"),
            ("GC_PEAK_WIDTH", "Peak Width:", "2"),
            ("GC_RT_TOLERANCE", "RT Tolerance (min):", "0.05"),
            ("GC_SIMILARITY_THRESHOLD", "Similarity Threshold:", "0.7")
        ]

        for i, (param, label, default) in enumerate(gc_params):
            ttk.Label(gc_param_grid, text=label, width=20).grid(row=i, column=0, padx=5, pady=3, sticky=tk.W)
            var = tk.StringVar(value=default)
            entry = ttk.Entry(gc_param_grid, textvariable=var, width=15)
            gc_entries[param] = var
            entry.grid(row=i, column=1, padx=5, pady=3, sticky=tk.W)

        # 操作按钮区域 - 将镜像图设置移到按钮旁边
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(10, 10))

        run_btn = ttk.Button(button_frame, text="▶ START ANALYSIS", command=run_analysis, width=25)
        run_btn.pack(side=tk.LEFT, padx=(0, 20))

        download_btn = ttk.Button(button_frame, text="📥 DOWNLOAD RESULTS", command=download_results, width=25)
        download_btn.pack(side=tk.LEFT, padx=(0, 20))

        # 镜像图设置区域 - 放在按钮旁边
        mirror_frame = ttk.Frame(button_frame)
        mirror_frame.pack(side=tk.LEFT, padx=(10, 0))

        mirror_check = ttk.Checkbutton(mirror_frame, text="Generate Mirror Spectra (SVG)",
                                       variable=generate_mirror_var)
        mirror_check.pack(side=tk.LEFT, padx=5)

        # 动态阈值设置区域
        dynamic_threshold_frame = ttk.Frame(button_frame)
        dynamic_threshold_frame.pack(side=tk.LEFT, padx=(10, 0))

        dynamic_threshold_check = ttk.Checkbutton(dynamic_threshold_frame, text="Use Dynamic Thresholds",
                                                  variable=use_dynamic_thresholds_var)
        dynamic_threshold_check.pack(side=tk.LEFT, padx=5)

        # 结果展示区域
        result_frame = ttk.LabelFrame(main_frame, text="Analysis Results")
        result_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        # 初始显示LC-MS列
        columns = ("Precursor m/z", "RT", "Name", "Score", "Matched Peaks", "Error(ppm)", "Sample")
        result_tree = ttk.Treeview(result_frame, columns=columns, show="headings", height=12)

        col_widths = [100, 70, 200, 70, 90, 70, 120]
        for col, width in zip(columns, col_widths):
            result_tree.heading(col, text=col)
            result_tree.column(col, width=width, anchor="center")

        tree_scroll_y = ttk.Scrollbar(result_frame, orient=tk.VERTICAL, command=result_tree.yview)
        tree_scroll_x = ttk.Scrollbar(result_frame, orient=tk.HORIZONTAL, command=result_tree.xview)
        result_tree.configure(yscrollcommand=tree_scroll_y.set, xscrollcommand=tree_scroll_x.set)

        result_tree.grid(row=0, column=0, sticky="nsew")
        tree_scroll_y.grid(row=0, column=1, sticky="ns")
        tree_scroll_x.grid(row=1, column=0, sticky="ew")

        result_frame.columnconfigure(0, weight=1)
        result_frame.rowconfigure(0, weight=1)

        # 进度显示区域
        progress_frame = ttk.LabelFrame(main_frame, text="Progress")
        progress_frame.pack(fill=tk.X, pady=(0, 0))

        progress_text = tk.Text(progress_frame, height=6, wrap=tk.WORD, font=("Consolas", 10))
        progress_scroll = ttk.Scrollbar(progress_frame, orient=tk.VERTICAL, command=progress_text.yview)
        progress_text.configure(yscrollcommand=progress_scroll.set)

        progress_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)
        progress_scroll.pack(side=tk.RIGHT, fill=tk.Y, padx=(0, 5), pady=5)

        # 初始更新参数显示
        update_parameter_display()

        # 设置窗口关闭事件
        def on_closing():
            if processing:
                if messagebox.askokcancel("Quit", "Analysis is running. Are you sure you want to quit?"):
                    analysis_window.destroy()
            else:
                analysis_window.destroy()

        analysis_window.protocol("WM_DELETE_WINDOW", on_closing)
        analysis_window.mainloop()

    except Exception as e:
        messagebox.showerror("Error", f"Failed to start application: {str(e)}")
        import traceback
        traceback.print_exc()
def main():
    """主函数"""
    # Create main window
    root = tk.Tk()
    root.title("NeuroToxCom Seeker V1.0")
    root.geometry("1400x800")
    root.configure(bg=colors["light"])

    # Global variables
    compound_data = None
    tree_columns = None

    # Top navigation bar
    top_bar = tk.Frame(root, bg=colors["primary"], height=70, bd=0)
    top_bar.pack(fill="x", side="top")

    db_name = tk.Label(top_bar,
                       text="🧪 NeuroToxCom Seeker V1.0",
                       fg="white",
                       bg=colors["primary"],
                       font=font_large,
                       padx=20)
    db_name.pack(side="left")

    nav_buttons = tk.Frame(top_bar, bg=colors["primary"])
    nav_buttons.pack(side="right", padx=20)

    def create_nav_button(parent, text, icon=None, command=None):
        btn = tk.Button(parent,
                        text=f" {text}" if not icon else f"{icon} {text}",
                        fg="white",
                        bg=colors["primary"],
                        activeforeground="white",
                        activebackground=colors["secondary"],
                        font=font_medium,
                        bd=0,
                        padx=15,
                        pady=5,
                        command=command)
        btn.pack(side="left", padx=5)
        return btn

    help_doc = create_nav_button(nav_buttons, "Help", "❓", command=lambda: open_help_document("main"))

    # Main container
    main_container = tk.Frame(root, bg=colors["light"])
    main_container.pack(fill="both", expand=True, padx=0, pady=0)

    # Left sidebar
    sidebar = tk.Frame(main_container,
                       bg=colors["secondary"],
                       width=250,
                       padx=10,
                       pady=20)
    sidebar.pack(fill="y", side="left")

    sidebar_title = tk.Label(sidebar,
                             text="Navigation",
                             fg="white",
                             bg=colors["secondary"],
                             font=("Segoe UI", 14, "bold"),
                             pady=10)
    sidebar_title.pack(fill="x")

    # Sidebar items
    sidebar_items = [
        ("🔬 Physicochemical", "Physicochemical"),
        ("🧫 Compound Classes", "Compound Classes"),
        ("⚠️ AOP Analysis", "AOP Analysis"),
        ("🎯 Targets & Mechanisms", "Targets & Mechanisms"),
        ("💊 ADME Properties", "ADME Properties"),
        ("❗ Structural Alerts", "Structural Alerts"),
        ("🔍 Suspect List Screening", "Suspect List Screening"),
        ("📊 Spectral Library", "Spectral Library")
    ]

    for text, file_key in sidebar_items:
        item_frame = tk.Frame(sidebar, bg=colors["secondary"])
        item_frame.pack(fill="x", pady=2)

        btn = tk.Button(item_frame,
                        text=text,
                        anchor="w",
                        fg="white",
                        bg=colors["secondary"],
                        activeforeground="white",
                        activebackground=colors["primary"],
                        font=font_medium,
                        bd=0,
                        padx=15,
                        pady=12,
                        relief="flat",
                        command=lambda fk=file_key: load_sidebar_file(fk))
        btn.pack(side="left", fill="x", expand=True)

        help_btn = tk.Button(item_frame,
                             text="?",
                             fg="white",
                             bg=colors["accent"],
                             activeforeground="white",
                             activebackground=colors["dark"],
                             font=("Segoe UI", 10, "bold"),
                             bd=0,
                             padx=5,
                             pady=0,
                             width=2,
                             relief="flat",
                             command=lambda fk=file_key: open_help_document(fk))
        help_btn.pack(side="right", padx=(0, 5))

    # Main content area
    main_area = tk.Frame(main_container, bg=colors["light"], padx=25, pady=25)
    main_area.pack(fill="both", expand=True)

    # Search card
    search_card = tk.Frame(main_area,
                           bg="white",
                           padx=25,
                           pady=20,
                           highlightbackground="#d1e0ed",
                           highlightthickness=1)
    search_card.pack(fill="x")

    search_title = tk.Label(search_card,
                            text="Compound Search",
                            fg=colors["dark"],
                            bg="white",
                            font=("Segoe UI", 16, "bold"),
                            anchor="w")
    search_title.pack(fill="x", pady=(0, 15))

    # 使用网格布局，但增加列数以容纳所有按钮
    search_form = tk.Frame(search_card, bg="white")
    search_form.pack(fill="x")

    # 第一行：标签和搜索框
    search_label = tk.Label(search_form,
                            text="Enter CAS No. or Compound Name:",
                            fg=colors["text"],
                            bg="white",
                            font=font_medium)
    search_label.grid(row=0, column=0, padx=(0, 10), sticky="w")

    # 增加搜索框宽度，但仍留出空间给按钮
    search_entry = tk.Entry(search_form,
                            width=35,
                            font=font_medium,
                            bd=1,
                            relief="solid",
                            highlightcolor=colors["accent"],
                            highlightthickness=1)
    search_entry.grid(row=0, column=1, columnspan=2, padx=10, sticky="ew")

    # 配置列的权重，使搜索框可以扩展
    search_form.grid_columnconfigure(1, weight=1)

    # 第二行：所有按钮放在一行
    button_row = 1
    button_frame = tk.Frame(search_form, bg="white")
    button_frame.grid(row=button_row, column=0, columnspan=6, pady=10, sticky="ew")

    # 搜索按钮
    search_button = tk.Button(button_frame,
                              text="🔍 Search",
                              fg="white",
                              bg=colors["accent"],
                              font=font_medium,
                              padx=10,
                              pady=5,
                              command=lambda: search_compounds(search_entry, tree, compound_data, tree_columns,
                                                               result_count))
    search_button.pack(side="left", padx=2)

    # 清空按钮
    clear_button = tk.Button(button_frame,
                             text="🗑️ Clear",
                             fg="white",
                             bg=colors["danger"],
                             font=font_medium,
                             padx=10,
                             pady=5,
                             command=lambda: clear_search(tree, compound_data, result_count, search_entry))
    clear_button.pack(side="left", padx=2)

    # 批量搜索按钮
    batch_search_btn = tk.Button(button_frame,
                                 text="🧪 Batch Search",
                                 fg="white",
                                 bg=colors["warning"],
                                 font=font_medium,
                                 padx=10,
                                 pady=5,
                                 command=lambda: batch_search(compound_data, tree, result_count))
    batch_search_btn.pack(side="left", padx=2)

    # 上传信息按钮
    upload_button = tk.Button(button_frame,
                              text="📤 Upload Information",
                              fg="white",
                              bg=colors["primary"],
                              font=font_medium,
                              padx=10,
                              pady=5,
                              command=lambda: upload_compound_data(tree, compound_data, result_count))
    upload_button.pack(side="left", padx=2)

    # 编辑按钮
    edit_button = tk.Button(button_frame,
                            text="✏️ Edit Selected",
                            fg="white",
                            bg=colors["warning"],
                            font=font_medium,
                            padx=10,
                            pady=5,
                            command=lambda: edit_selected_compound(tree, compound_data, result_count))
    edit_button.pack(side="left", padx=2)

    # 删除按钮
    delete_button = tk.Button(button_frame,
                              text="🗑️ Delete Selected",
                              fg="white",
                              bg=colors["danger"],
                              font=font_medium,
                              padx=10,
                              pady=5,
                              command=lambda: delete_selected_compound(tree, compound_data, result_count))
    delete_button.pack(side="left", padx=2)

    # Results area
    result_frame = tk.Frame(main_area, bg="white", padx=0, pady=0)
    result_frame.pack(fill="both", expand=True, pady=(20, 0))

    result_header = tk.Frame(result_frame, bg=colors["primary"], padx=20, pady=10)
    result_header.pack(fill="x")

    result_label = tk.Label(result_header,
                            text="Search Results",
                            fg="white",
                            bg=colors["primary"],
                            font=("Segoe UI", 14, "bold"))
    result_label.pack(side="left")

    result_actions = tk.Frame(result_header, bg=colors["primary"])
    result_actions.pack(side="right")

    result_count = tk.Label(result_actions,
                            text="0 records",
                            fg="white",
                            bg=colors["primary"],
                            font=font_medium)
    result_count.pack(side="left", padx=(0, 15))

    download_btn = tk.Button(result_actions,
                             text="⬇️ Download",
                             fg="white",
                             bg=colors["success"],
                             font=font_medium,
                             padx=10,
                             pady=2,
                             bd=0,
                             command=lambda: download_results(tree, tree_columns))
    download_btn.pack(side="left")

    # Create tree container with scrollbars
    tree_container = tk.Frame(result_frame)
    tree_container.pack(fill="both", expand=True)

    # Create horizontal scrollbar
    hscroll = ttk.Scrollbar(tree_container, orient="horizontal")
    hscroll.pack(side="bottom", fill="x")

    # Create vertical scrollbar
    vscroll = ttk.Scrollbar(tree_container, orient="vertical")
    vscroll.pack(side="right", fill="y")

    # Create the treeview with initial empty columns
    tree = ttk.Treeview(tree_container, columns=[], show="headings", height=10,
                        xscrollcommand=hscroll.set, yscrollcommand=vscroll.set)

    # Configure scrollbars
    hscroll.config(command=tree.xview)
    vscroll.config(command=tree.yview)

    # Style the table
    style = ttk.Style()
    style.configure("Treeview",
                    font=font_small,
                    rowheight=25,
                    background="white",
                    fieldbackground="white")
    style.configure("Treeview.Heading",
                    font=("Segoe UI", 9, "bold"),
                    background="#e1ebf5",
                    foreground=colors["dark"])

    tree.pack(fill="both", expand=True)

    # Status bar
    bottom_bar = tk.Frame(root,
                          bg=colors["primary"],
                          height=50,
                          padx=20)
    bottom_bar.pack(fill="x", side="bottom")

    status_left = tk.Frame(bottom_bar, bg=colors["primary"])
    status_left.pack(side="left")

    status_right = tk.Frame(bottom_bar, bg=colors["primary"])
    status_right.pack(side="right")

    def create_status_label(parent, text):
        return tk.Label(parent,
                        text=text,
                        fg="white",
                        bg=colors["primary"],
                        font=font_small)

    data_source = create_status_label(status_left, "Data sources: Comptox, ChemNTP, ToxAlerts")
    data_source.pack(side="left", padx=10)

    contact_info = create_status_label(status_right, "Contact: 2798063448@qq.com")
    contact_info.pack(side="right", padx=10)

    version_info = create_status_label(status_right, "Version: v1.0")
    version_info.pack(side="right", padx=10)

    # Load compound data and setup treeview
    compound_data = load_compound_data()
    if compound_data is not None:
        tree_columns = list(compound_data.columns)
        setup_treeview_columns(tree, tree_columns)
        load_all_data_to_tree(tree, compound_data, result_count)

    root.mainloop()


if __name__ == "__main__":
    main()