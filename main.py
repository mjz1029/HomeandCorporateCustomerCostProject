import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
from tkcalendar import DateEntry
import pandas as pd
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
import os
import json
from datetime import datetime

# ===================== 配置与常量 =====================
CONFIG_FILE = "config.json"
BUDGET_DATA_FILE = "budget_data.json"
EXCEL_SHEETS = ["施工项目（Sheet1）", "材料项目（Sheet2）"]
MAX_IMG_WIDTH = Inches(4)
MAX_IMG_HEIGHT = Inches(3)


class HomeAndEnterpriseTool:
    def __init__(self, root):
        self.root = root
        self.root.title("家集客项目预算与文档生成系统 v2.1")

        # ========== UI适配优化：调整窗口大小以适配老旧屏幕 ==========
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()

        # 默认大小设置为屏幕的 80% 或固定值，适配 1366x768 分辨率
        default_width = 1200
        default_height = 760
        if screen_width < 1280:
            default_width = 1000
            default_height = 700

        # 居中显示
        x_cordinate = int((screen_width / 2) - (default_width / 2))
        y_cordinate = int((screen_height / 2) - (default_height / 2))

        self.root.geometry(f"{default_width}x{default_height}+{x_cordinate}+{y_cordinate}")
        self.root.minsize(960, 600)

        # 核心数据存储
        self.budget_data = []
        self.total_amount = 0.0
        self.base_info = {}
        self.word_app_template = None
        self.word_review_template = None
        self.image_paths = []

        self.status_var = tk.StringVar(value="✅ 系统初始化完成")

        # 加载数据
        self.load_config()
        self.load_budget_data()
        if not self.budget_data:
            self.load_budget_excel()

        # 初始化GUI
        self.setup_style()
        self.setup_ui()

    # ===================== 样式配置（美化版） =====================
    def setup_style(self):
        self.style = ttk.Style(self.root)
        self.style.theme_use("clam")

        # 定义颜色和字体
        primary_color = "#0078D7"  # 商务蓝
        bg_color = "#F0F2F5"  # 浅灰背景
        font_main = ("Microsoft YaHei UI", 9)
        font_bold = ("Microsoft YaHei UI", 9, "bold")

        self.root.configure(bg=bg_color)

        # LabelFrame 样式
        self.style.configure("Custom.TLabelframe",
                             background=bg_color,
                             relief="flat",
                             borderwidth=1)
        self.style.configure("Custom.TLabelframe.Label",
                             font=("Microsoft YaHei UI", 10, "bold"),
                             foreground=primary_color,
                             background=bg_color)

        # Frame 样式
        self.style.configure("TFrame", background=bg_color)
        self.style.configure("TLabelframe", background=bg_color)

        # Label 样式
        self.style.configure("TLabel", background=bg_color, font=font_main, foreground="#333")

        # Button 样式
        self.style.configure("Accent.TButton",
                             font=font_main,
                             background=primary_color,
                             foreground="white",
                             borderwidth=0,
                             focuscolor="none")
        self.style.map("Accent.TButton",
                       background=[('active', '#005A9E'), ('pressed', '#004578')])

        self.style.configure("Generate.TButton",
                             font=("Microsoft YaHei UI", 12, "bold"),
                             background="#28a745",  # 绿色
                             foreground="white",
                             padding=10)
        self.style.map("Generate.TButton",
                       background=[('active', '#218838')])

        # Treeview (表格) 样式
        self.style.configure("Treeview",
                             font=("Microsoft YaHei UI", 9),
                             rowheight=28,
                             background="white",
                             fieldbackground="white",
                             borderwidth=0)
        self.style.configure("Treeview.Heading",
                             font=font_bold,
                             background="#E1E4E8",
                             foreground="#333",
                             relief="flat")
        self.style.map("Treeview", background=[("selected", primary_color)])

    # ===================== GUI界面布局（紧凑优化版） =====================
    def setup_ui(self):
        # 主容器
        main_container = ttk.Frame(self.root, padding="10 10 10 10")
        main_container.pack(fill=tk.BOTH, expand=True)

        # --- 1. 顶部区域 ---
        top_frame = ttk.LabelFrame(main_container, text="🛠️ 项目与基础信息配置", style="Custom.TLabelframe")
        top_frame.pack(fill=tk.X, pady=(0, 10))

        # 第一行：项目核心信息
        input_frame_1 = ttk.Frame(top_frame)
        input_frame_1.pack(fill=tk.X, padx=10, pady=5)

        ttk.Label(input_frame_1, text="项目名称：").pack(side=tk.LEFT)
        self.project_name_var = tk.StringVar(value="广电项目光猫安装、开通")
        ttk.Entry(input_frame_1, textvariable=self.project_name_var, width=35).pack(side=tk.LEFT, padx=(0, 15))

        ttk.Label(input_frame_1, text="项目日期：").pack(side=tk.LEFT)
        self.date_entry = DateEntry(input_frame_1, width=12, background="#0078D7", foreground="white",
                                    date_pattern="yyyy年MM月dd日")
        self.date_entry.pack(side=tk.LEFT, padx=(0, 15))

        ttk.Label(input_frame_1, text="实施周期：").pack(side=tk.LEFT)
        self.cycle_var = tk.StringVar(value="15天")
        ttk.Entry(input_frame_1, textvariable=self.cycle_var, width=8).pack(side=tk.LEFT)

        ttk.Separator(top_frame, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=10, pady=5)

        # 第二行：基础信息
        info_frame = ttk.Frame(top_frame)
        info_frame.pack(fill=tk.X, padx=10, pady=(0, 8))

        fields = [
            ("申请单位", "申请单位", 0, 0), ("申请人", "申请人", 0, 2),
            ("联系电话", "联系电话", 0, 4), ("实施单位", "实施单位", 0, 6),
            ("项目经理", "项目经理", 1, 0), ("经理电话", "项目经理联系电话", 1, 2),
            ("负责人", "项目负责人", 1, 4)
        ]

        for label, key, r, c in fields:
            ttk.Label(info_frame, text=f"{label}：").grid(row=r, column=c, sticky=tk.W, padx=(0, 5), pady=2)
            entry = ttk.Entry(info_frame, width=15)
            entry.grid(row=r, column=c + 1, sticky=tk.W, padx=(0, 15), pady=2)
            entry.insert(0, self.base_info.get(key, ""))
            entry.bind("<FocusOut>", lambda e, k=key, ent=entry: self.update_base_info(k, ent.get()))

        ttk.Button(info_frame, text="💾 保存默认信息", command=self.save_config, style="Accent.TButton").grid(row=1,
                                                                                                             column=6,
                                                                                                             columnspan=2,
                                                                                                             sticky=tk.EW,
                                                                                                             padx=5)

        # --- 2. 中间区域：预算编辑 ---
        budget_frame = ttk.LabelFrame(main_container, text="💰 预算明细编辑 (工程量为0不导出)",
                                      style="Custom.TLabelframe")
        budget_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        # 工具栏
        tool_bar = ttk.Frame(budget_frame)
        tool_bar.pack(fill=tk.X, padx=5, pady=5)

        for txt, cmd in [("➕ 施工项", self.add_construction_project),
                         ("➕ 材料项", self.add_material_project),
                         ("✏️ 修改", self.edit_project_info),
                         ("🗑️ 删除", self.delete_selected_project)]:
            ttk.Button(tool_bar, text=txt, command=cmd, style="Accent.TButton", width=10).pack(side=tk.LEFT, padx=3)

        ttk.Button(tool_bar, text="📤 导出Excel", command=self.export_budget_to_excel).pack(side=tk.RIGHT, padx=5)

        # 标签页 (Tab)
        notebook = ttk.Notebook(budget_frame)
        notebook.pack(fill=tk.BOTH, expand=True, padx=5, pady=0)

        # ========== 修复点：正确添加 Tab ==========
        self.construction_tree = self.create_scrolled_tree(notebook, "施工项目")
        # .master 是 frame，.master.master 是 notebook。我们只需添加 frame。
        notebook.add(self.construction_tree.master, text="  🚧 施工项目  ")

        self.material_tree = self.create_scrolled_tree(notebook, "材料项目")
        notebook.add(self.material_tree.master, text="  🔩 材料项目  ")

        # 总金额条
        total_bar = ttk.Frame(budget_frame, style="TFrame")
        total_bar.pack(fill=tk.X, padx=10, pady=5)
        self.total_var = tk.StringVar(value="当前总金额：0.00元")
        lbl_total = ttk.Label(total_bar, textvariable=self.total_var, font=("Microsoft YaHei UI", 11, "bold"),
                              foreground="#D32F2F")
        lbl_total.pack(side=tk.RIGHT)
        ttk.Label(total_bar, text="双击表格行可快速修改工程量", foreground="#888", font=("Microsoft YaHei UI", 8)).pack(
            side=tk.LEFT)

        # --- 3. 底部区域：模板与生成 ---
        bottom_frame = ttk.LabelFrame(main_container, text="📄 文档生成配置", style="Custom.TLabelframe")
        bottom_frame.pack(fill=tk.X, pady=(0, 0))

        # 模板选择
        tpl_frame = ttk.Frame(bottom_frame)
        tpl_frame.pack(fill=tk.X, padx=10, pady=5)

        ttk.Label(tpl_frame, text="申请表模板：").grid(row=0, column=0, sticky=tk.W)
        self.app_template_var = tk.StringVar(value="未选择")
        ttk.Entry(tpl_frame, textvariable=self.app_template_var, state="readonly", width=25).grid(row=0, column=1,
                                                                                                  padx=5)
        ttk.Button(tpl_frame, text="📂", width=3, command=lambda: self.select_template("app")).grid(row=0, column=2,
                                                                                                   padx=(0, 15))

        ttk.Label(tpl_frame, text="会审单模板：").grid(row=0, column=3, sticky=tk.W)
        self.review_template_var = tk.StringVar(value="未选择")
        ttk.Entry(tpl_frame, textvariable=self.review_template_var, state="readonly", width=25).grid(row=0, column=4,
                                                                                                     padx=5)
        ttk.Button(tpl_frame, text="📂", width=3, command=lambda: self.select_template("review")).grid(row=0, column=5,
                                                                                                      padx=(0, 15))

        # 图片上传
        ttk.Label(tpl_frame, text="现场图片：").grid(row=0, column=6, sticky=tk.W)
        self.image_count_var = tk.StringVar(value="0张")
        ttk.Label(tpl_frame, textvariable=self.image_count_var,
                  foreground=self.style.lookup("Accent.TButton", "background")).grid(row=0, column=7, padx=5)
        ttk.Button(tpl_frame, text="⬆ 上传", width=6, command=self.upload_images, style="Accent.TButton").grid(row=0,
                                                                                                               column=8,
                                                                                                               padx=2)
        ttk.Button(tpl_frame, text="♻ 清空", width=6, command=self.clear_images).grid(row=0, column=9, padx=2)

        # 底部大按钮与状态栏
        action_frame = ttk.Frame(main_container)
        action_frame.pack(fill=tk.X, pady=10)

        self.generate_btn = ttk.Button(action_frame, text="🚀 一键生成申请表 + 会审单", command=self.generate_documents,
                                       style="Generate.TButton")
        self.generate_btn.pack(side=tk.RIGHT, padx=10)

        status_label = ttk.Label(action_frame, textvariable=self.status_var, foreground="#0078D7",
                                 font=("Microsoft YaHei UI", 9))
        status_label.pack(side=tk.LEFT, padx=10)

        self.refresh_treeviews()

    # ===================== 辅助UI构建函数 =====================
    def create_scrolled_tree(self, parent, category):
        """创建一个带滚动条的Treeview容器"""
        # ========== 修复点：移除 frame.pack() ==========
        frame = ttk.Frame(parent)
        # frame.pack(fill=tk.BOTH, expand=True) <--- 已删除此行

        vscroll = ttk.Scrollbar(frame, orient=tk.VERTICAL)
        hscroll = ttk.Scrollbar(frame, orient=tk.HORIZONTAL)

        columns = ["id", "name", "unit_price", "quantity", "total"]
        tree = ttk.Treeview(frame, columns=columns, show="headings",
                            yscrollcommand=vscroll.set, xscrollcommand=hscroll.set,
                            selectmode="browse")

        vscroll.config(command=tree.yview)
        hscroll.config(command=tree.xview)

        vscroll.pack(side=tk.RIGHT, fill=tk.Y)
        hscroll.pack(side=tk.BOTTOM, fill=tk.X)
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        tree.heading("id", text="序号")
        tree.heading("name", text="项目名称")
        tree.heading("unit_price", text="单价 (元)")
        tree.heading("quantity", text="工程量")
        tree.heading("total", text="合计 (元)")

        tree.column("id", width=50, anchor="center")
        tree.column("name", width=400, anchor="w")
        tree.column("unit_price", width=100, anchor="e")
        tree.column("quantity", width=100, anchor="center")
        tree.column("total", width=100, anchor="e")

        tree.tag_configure("oddrow", background="white")
        tree.tag_configure("evenrow", background="#F8F9FA")

        tree.bind("<Double-1>", self.edit_quantity)
        return tree

    # ===================== 逻辑功能保持不变 =====================
    def save_budget_data(self):
        try:
            with open(BUDGET_DATA_FILE, "w", encoding="utf-8") as f:
                json.dump(self.budget_data, f, ensure_ascii=False, indent=2)
            self.status_var.set("✅ 预算数据已保存到本地")
        except Exception as e:
            messagebox.showerror("数据保存失败", f"错误原因：{str(e)}")

    def load_budget_data(self):
        if os.path.exists(BUDGET_DATA_FILE):
            try:
                with open(BUDGET_DATA_FILE, "r", encoding="utf-8") as f:
                    self.budget_data = json.load(f)
                for item in self.budget_data:
                    item["quantity"] = 0.0
                    item["total"] = 0.0
                for idx, item in enumerate(self.budget_data):
                    item["id"] = idx + 1
            except Exception as e:
                messagebox.showwarning("本地数据加载失败", f"将重新导入Excel：{str(e)}")
                self.budget_data = []
        else:
            self.budget_data = []

    def load_config(self):
        default_info = {
            "申请单位": "奇台县分公司", "申请人": "樊斌", "联系电话": "13909949883",
            "实施单位": "中移建设", "项目经理": "吴斌", "项目经理联系电话": "18899661100",
            "项目负责人": "樊斌"
        }
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                    self.base_info = json.load(f)
                for key, val in default_info.items():
                    if key not in self.base_info:
                        self.base_info[key] = val
            except Exception as e:
                self.base_info = default_info
        else:
            self.base_info = default_info
            self.save_config()

    def save_config(self):
        try:
            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump(self.base_info, f, ensure_ascii=False, indent=2)
            self.status_var.set("✅ 基础信息已保存")
        except Exception as e:
            messagebox.showerror("配置保存失败", str(e))

    def load_budget_excel(self):
        file_path = filedialog.askopenfilename(
            title="选择家集客预算表",
            filetypes=[("Excel文件", "*.xlsx"), ("所有文件", "*.*")]
        )
        if not file_path:
            messagebox.showwarning("提示", "未选择预算表，应用将无法正常使用！")
            return

        try:
            sheet1 = pd.read_excel(file_path, sheet_name=0)
            if sheet1.empty: raise ValueError("Sheet1为空")
            sheet1_data = self.parse_sheet1(sheet1)

            sheet2 = pd.read_excel(file_path, sheet_name=1)
            if sheet2.empty: raise ValueError("Sheet2为空")
            sheet2_data = self.parse_sheet2(sheet2)

            self.budget_data = sheet1_data + sheet2_data
            for idx, item in enumerate(self.budget_data):
                item["id"] = idx + 1
            self.save_budget_data()
            messagebox.showinfo("加载成功", f"共加载{len(self.budget_data)}个项目")
        except Exception as e:
            messagebox.showerror("预算表加载失败", f"错误原因：{str(e)}")

    def parse_sheet1(self, df):
        parsed = []
        df.columns = df.columns.str.strip()
        required_cols = ["类别", "折扣后（含税）37%/元"]
        missing_cols = [col for col in required_cols if col not in df.columns]
        if missing_cols: raise ValueError(f"Sheet1缺少必要列：{', '.join(missing_cols)}")

        for _, row in df.iterrows():
            project_name = str(row["类别"]).strip()
            if not project_name or project_name == "nan": continue
            unit_price = float(pd.to_numeric(row["折扣后（含税）37%/元"], errors="coerce")) if pd.notna(
                row["折扣后（含税）37%/元"]) else 0.0
            is_length_unit = "元/公里" in project_name
            parsed.append({
                "id": len(parsed) + 1, "category": "施工项目", "name": project_name,
                "unit": "公里" if is_length_unit else "个/户/处等",
                "unit_price": unit_price, "quantity": 0.0, "total": 0.0, "is_length": is_length_unit
            })
        if not parsed: raise ValueError("Sheet1无有效数据")
        return parsed

    def parse_sheet2(self, df):
        parsed = []
        df.columns = df.columns.str.strip()
        required_cols = ["材料", "含税"]
        missing_cols = [col for col in required_cols if col not in df.columns]
        if missing_cols: raise ValueError(f"Sheet2缺少必要列：{', '.join(missing_cols)}")

        for _, row in df.iterrows():
            project_name = str(row["材料"]).strip()
            if not project_name or project_name == "nan": continue
            unit_price = float(pd.to_numeric(row["含税"], errors="coerce")) if pd.notna(row["含税"]) else 0.0
            parsed.append({
                "id": len(parsed) + 1, "category": "材料项目", "name": project_name,
                "unit": "个", "unit_price": unit_price, "quantity": 0.0, "total": 0.0, "is_length": False
            })
        if not parsed: raise ValueError("Sheet2无有效数据")
        return parsed

    def refresh_treeviews(self):
        for item in self.construction_tree.get_children():
            self.construction_tree.delete(item)
        for item in self.material_tree.get_children():
            self.material_tree.delete(item)

        if not self.budget_data:
            self.total_var.set(f"当前总金额：0.00元")
            return

        self.total_amount = 0.0

        count_c = 0
        count_m = 0

        for item in self.budget_data:
            total = float(item["quantity"]) * float(item["unit_price"])
            item["total"] = total
            self.total_amount += total

            quantity_str = f"{float(item['quantity']):.2f}"
            unit_price_str = f"{float(item['unit_price']):.2f}"
            total_str = f"{total:.2f}"

            values = [item["id"], item["name"], unit_price_str, quantity_str, total_str]

            if item["category"] == "施工项目":
                tag = "evenrow" if count_c % 2 == 0 else "oddrow"
                self.construction_tree.insert("", tk.END, values=values, tags=(tag,))
                count_c += 1
            else:
                tag = "evenrow" if count_m % 2 == 0 else "oddrow"
                self.material_tree.insert("", tk.END, values=values, tags=(tag,))
                count_m += 1

        self.total_var.set(f"当前总金额：{self.total_amount:.2f}元")

    def add_construction_project(self):
        name = simpledialog.askstring("新增施工项目", "请输入项目名称：")
        if not name: return
        unit_price = simpledialog.askfloat("新增施工项目", "请输入单价（元）：", initialvalue=0.0)
        if unit_price is None: return
        unit = simpledialog.askstring("新增施工项目", "请输入单位（如：公里、个）：", initialvalue="公里")
        if not unit: unit = "公里"
        is_length = simpledialog.askyesno("新增施工项目", "是否为长度类项目（单位：公里）？")
        new_quantity = simpledialog.askfloat("新增施工项目", "请输入工程量：", initialvalue=0.0)
        quantity = new_quantity if new_quantity is not None else 0.0

        new_id = len(self.budget_data) + 1
        self.budget_data.append({
            "id": new_id, "category": "施工项目", "name": name.strip(),
            "unit": unit.strip(), "unit_price": unit_price, "quantity": quantity,
            "total": unit_price * quantity, "is_length": is_length
        })
        self.save_budget_data()
        self.refresh_treeviews()
        self.status_var.set(f"✅ 新增施工项目：{name}")

    def add_material_project(self):
        name = simpledialog.askstring("新增材料项目", "请输入材料名称：")
        if not name: return
        unit_price = simpledialog.askfloat("新增材料项目", "请输入单价（元）：", initialvalue=0.0)
        if unit_price is None: return
        unit = simpledialog.askstring("新增材料项目", "请输入单位（如：个）：", initialvalue="个")
        if not unit: unit = "个"
        new_quantity = simpledialog.askfloat("新增材料项目", "请输入工程量：", initialvalue=0.0)
        quantity = new_quantity if new_quantity is not None else 0.0

        new_id = len(self.budget_data) + 1
        self.budget_data.append({
            "id": new_id, "category": "材料项目", "name": name.strip(),
            "unit": unit.strip(), "unit_price": unit_price, "quantity": quantity,
            "total": unit_price * quantity, "is_length": False
        })
        self.save_budget_data()
        self.refresh_treeviews()
        self.status_var.set(f"✅ 新增材料项目：{name}")

    def delete_selected_project(self):
        selected_item = None
        current_tree = None
        if self.construction_tree.focus():
            selected_item = self.construction_tree.focus()
            current_tree = self.construction_tree
        elif self.material_tree.focus():
            selected_item = self.material_tree.focus()
            current_tree = self.material_tree

        if not selected_item:
            messagebox.showwarning("提示", "请先选中要删除的项目！")
            return

        item_values = current_tree.item(selected_item)["values"]
        project_id = int(item_values[0])
        self.budget_data = [item for item in self.budget_data if item["id"] != project_id]
        for idx, item in enumerate(self.budget_data): item["id"] = idx + 1
        self.save_budget_data()
        self.refresh_treeviews()
        self.status_var.set(f"✅ 删除项目ID：{project_id}")

    def edit_project_info(self):
        selected_item = None
        current_tree = None
        if self.construction_tree.focus():
            selected_item = self.construction_tree.focus()
            current_tree = self.construction_tree
        elif self.material_tree.focus():
            selected_item = self.material_tree.focus()
            current_tree = self.material_tree

        if not selected_item:
            messagebox.showwarning("提示", "请先选中要修改的项目！")
            return

        item_values = current_tree.item(selected_item)["values"]
        project_id = int(item_values[0])
        target_item = next((item for item in self.budget_data if item["id"] == project_id), None)
        if not target_item: return

        new_name = simpledialog.askstring("修改", "项目名称：", initialvalue=target_item["name"])
        if not new_name: return
        new_unit_price = simpledialog.askfloat("修改", "单价（元）：", initialvalue=target_item["unit_price"])
        if new_unit_price is None: return
        new_quantity = simpledialog.askfloat("修改", "工程量：", initialvalue=target_item["quantity"])
        if new_quantity is None: return

        if target_item["category"] == "施工项目":
            new_unit = simpledialog.askstring("修改", "单位：", initialvalue=target_item["unit"])
            target_item["unit"] = new_unit.strip() if new_unit else target_item["unit"]
            new_is_length = simpledialog.askyesno("修改", "是否为长度类项目？", initialvalue=target_item["is_length"])
            target_item["is_length"] = new_is_length
        else:
            new_unit = simpledialog.askstring("修改", "单位：", initialvalue=target_item["unit"])
            target_item["unit"] = new_unit.strip() if new_unit else target_item["unit"]

        target_item["name"] = new_name.strip()
        target_item["unit_price"] = new_unit_price
        target_item["quantity"] = new_quantity
        target_item["total"] = new_unit_price * new_quantity

        self.save_budget_data()
        self.refresh_treeviews()
        self.status_var.set(f"✅ 修改项目ID：{project_id}")

    def edit_quantity(self, event):
        tree = event.widget
        focus_item = tree.focus()
        if not focus_item: return
        item_values = tree.item(focus_item)["values"]
        try:
            project_id = int(item_values[0])
            current_quantity = float(item_values[3])
        except:
            return

        new_quantity = simpledialog.askfloat("修改工程量", f"项目：{item_values[1]}\n请输入新工程量：",
                                             initialvalue=current_quantity)
        if new_quantity is None or new_quantity < 0: return

        for item in self.budget_data:
            if item.get("id") == project_id:
                item["quantity"] = float(new_quantity)
                item["total"] = float(new_quantity) * float(item["unit_price"])
                break

        self.save_budget_data()
        self.refresh_treeviews()
        self.status_var.set(f"✅ 更新工程量：{new_quantity:.2f}")

    def export_budget_to_excel(self):
        export_data = [item for item in self.budget_data if item["quantity"] > 0]
        if not export_data:
            messagebox.showwarning("提示", "无工程量>0的项目可导出！")
            return

        df = pd.DataFrame({
            "序号": [item["id"] for item in export_data],
            "类别": [item["category"] for item in export_data],
            "项目名称": [item["name"] for item in export_data],
            "单位": [item["unit"] for item in export_data],
            "单价（元）": [item["unit_price"] for item in export_data],
            "工程量": [item["quantity"] for item in export_data],
            "合计（元）": [item["total"] for item in export_data]
        })

        save_path = filedialog.asksaveasfilename(
            title="导出", defaultextension=".xlsx",
            filetypes=[("Excel文件", "*.xlsx")],
            initialfile=f"预算清单_{datetime.now().strftime('%Y%m%d')}.xlsx"
        )
        if save_path:
            try:
                df.to_excel(save_path, index=False)
                messagebox.showinfo("成功", f"导出{len(export_data)}条数据！")
            except Exception as e:
                messagebox.showerror("失败", str(e))

    def select_template(self, template_type):
        path = filedialog.askopenfilename(
            title=f"选择{'申请表' if template_type == 'app' else '会审单'}模板",
            filetypes=[("Word文件", "*.docx"), ("所有文件", "*.*")]
        )
        if not path: return
        if template_type == "app":
            self.word_app_template = path
            self.app_template_var.set(os.path.basename(path))
        else:
            self.word_review_template = path
            self.review_template_var.set(os.path.basename(path))

    def upload_images(self):
        paths = filedialog.askopenfilenames(
            title="选择支撑图片",
            filetypes=[("图片", "*.jpg;*.jpeg;*.png;*.bmp")]
        )
        if paths:
            remaining = 12 - len(self.image_paths)
            if len(paths) > remaining:
                paths = paths[:remaining]
            self.image_paths.extend(paths)
            self.image_count_var.set(f"{len(self.image_paths)}张")

    def clear_images(self):
        self.image_paths.clear()
        self.image_count_var.set("0张")

    def update_base_info(self, key, value):
        self.base_info[key] = value.strip()

    def generate_work_list(self):
        work_list = []
        for item in self.budget_data:
            if item["quantity"] <= 0: continue
            quantity = float(item["quantity"])
            if item["is_length"]:
                item_str = f"{quantity:.2f}公里 {item['name']}"
            else:
                item_str = f"{quantity:.2f}{item['unit']} {item['name']}"
            work_list.append(item_str)
        return "，".join(work_list) if work_list else "无有效项目"

    def insert_images_to_cell(self, cell, image_paths):
        if not image_paths: return
        cell.text = ""
        for img_path in image_paths:
            try:
                para = cell.add_paragraph()
                run = para.add_run()
                img = run.add_picture(img_path, width=MAX_IMG_WIDTH, height=MAX_IMG_HEIGHT)
                para.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            except:
                pass

    def find_cell_by_text(self, table, keyword_list):
        for row_idx, row in enumerate(table.rows):
            for col_idx, cell in enumerate(row.cells):
                for keyword in keyword_list:
                    if keyword in cell.text.strip():
                        return (row_idx, col_idx, cell)
        return (None, None, None)

    def generate_documents(self):
        if not self.word_app_template or not self.word_review_template:
            messagebox.showwarning("提示", "请先选择模板！")
            return
        if self.total_amount <= 0:
            messagebox.showwarning("提示", "无有效项目！")
            return

        project_name = self.project_name_var.get().strip()
        project_date = self.date_entry.get()
        cycle = self.cycle_var.get().strip()

        try:
            work_list = self.generate_work_list()
            self.fill_application_form(project_name, project_date, cycle, work_list)
            self.fill_review_form(project_name, project_date, cycle, work_list)
            messagebox.showinfo("成功", "文档生成完成！")
            self.status_var.set(f"✅ 生成成功！金额：{self.total_amount:.2f}元")
        except Exception as e:
            messagebox.showerror("失败", str(e))

    def fill_application_form(self, project_name, project_date, cycle, work_list):
        doc = Document(self.word_app_template)
        target_table = doc.tables[0]

        fill_items = [
            (0, 1, self.base_info["申请单位"], WD_PARAGRAPH_ALIGNMENT.LEFT),
            (0, 3, project_date, WD_PARAGRAPH_ALIGNMENT.CENTER),
            (0, 4, project_date, WD_PARAGRAPH_ALIGNMENT.CENTER),
            (0, 6, self.base_info["申请人"], WD_PARAGRAPH_ALIGNMENT.LEFT),
            (1, 6, self.base_info["联系电话"], WD_PARAGRAPH_ALIGNMENT.LEFT),
            (2, 1, cycle, WD_PARAGRAPH_ALIGNMENT.LEFT),
            (2, 3, f"{self.total_amount:.2f}元", WD_PARAGRAPH_ALIGNMENT.CENTER),
            (2, 4, f"{self.total_amount:.2f}元", WD_PARAGRAPH_ALIGNMENT.CENTER),
        ]
        for r, c, t, a in fill_items:
            try:
                cell = target_table.cell(r, c)
                cell.text = t
                for p in cell.paragraphs: p.alignment = a
            except:
                pass

        name_row, name_col, _ = self.find_cell_by_text(target_table, ["维修项目名称", "项目名称"])
        if name_row is None: name_row, name_col = 1, 1
        name_fill_col = min(name_col + 1, len(target_table.columns) - 1)
        cell = target_table.cell(name_row, name_fill_col)
        cell.text = project_name
        for p in cell.paragraphs:
            p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            for r in p.runs: r.font.size = Pt(10)

        list_row, list_col, _ = self.find_cell_by_text(target_table, ["工作量及材料清单", "工作量", "清单"])
        if list_row is None: list_row, list_col = max(0, len(target_table.rows) - 3), 0
        list_fill_col = min(list_col + 1, len(target_table.columns) - 1)
        cell = target_table.cell(list_row, list_fill_col)
        cell.text = work_list
        for p in cell.paragraphs:
            p.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
            for r in p.runs: r.font.size = Pt(9)

        row, col, _ = self.find_cell_by_text(target_table, ["其他需求支撑文件"])
        if row is not None:
            target_col = col + 1 if (col + 1) < len(target_table.columns) else len(target_table.columns) - 1
            self.insert_images_to_cell(target_table.cell(row, target_col), self.image_paths)
        else:
            self.insert_images_to_cell(
                target_table.cell(max(0, len(target_table.rows) - 2), len(target_table.columns) - 1), self.image_paths)

        save_path = filedialog.asksaveasfilename(
            title="保存申请表", defaultextension=".docx", filetypes=[("Word文件", "*.docx")],
            initialfile=f"{project_name}_申请表.docx"
        )
        if save_path: doc.save(save_path)

    def fill_review_form(self, project_name, project_date, cycle, work_list):
        doc = Document(self.word_review_template)
        target_table = doc.tables[0]

        fill_items = [
            (1, 1, f"{self.total_amount:.2f}元", WD_PARAGRAPH_ALIGNMENT.CENTER),
            (1, 5, project_date, WD_PARAGRAPH_ALIGNMENT.CENTER),
            (1, 9, cycle, WD_PARAGRAPH_ALIGNMENT.CENTER),
            (2, 1, self.base_info["项目负责人"], WD_PARAGRAPH_ALIGNMENT.CENTER),
            (2, 5, self.base_info["联系电话"], WD_PARAGRAPH_ALIGNMENT.CENTER),
            (3, 1, self.base_info["实施单位"], WD_PARAGRAPH_ALIGNMENT.CENTER),
            (3, 5, self.base_info["项目经理"], WD_PARAGRAPH_ALIGNMENT.CENTER),
            (3, 9, self.base_info["项目经理联系电话"], WD_PARAGRAPH_ALIGNMENT.CENTER),
        ]
        for r, c, t, a in fill_items:
            try:
                target_table.cell(r, c).text = t
            except:
                pass

        name_row, name_col, _ = self.find_cell_by_text(target_table, ["维修项目名称", "项目名称"])
        if name_row is None: name_row, name_col = 0, 1
        name_fill_col = min(name_col + 1, len(target_table.columns) - 1)
        cell = target_table.cell(name_row, name_fill_col)
        cell.text = project_name
        for p in cell.paragraphs:
            p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            for r in p.runs: r.font.size = Pt(10)

        list_row, list_col, _ = self.find_cell_by_text(target_table, ["主要工作量及材料清单", "工作量", "清单"])
        if list_row is None: list_row, list_col = max(0, len(target_table.rows) - 2), 0
        list_fill_col = min(list_col + 1, len(target_table.columns) - 1)
        cell = target_table.cell(list_row, list_fill_col)
        cell.text = f"工作量：{work_list}"
        for p in cell.paragraphs:
            p.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
            for r in p.runs: r.font.size = Pt(9)

        plan_row, plan_col, _ = self.find_cell_by_text(target_table, ["施工方实施计划"])
        if plan_row is None: plan_row, plan_col = list_row + 1, list_col
        plan_fill_col = min(plan_col + 1, len(target_table.columns) - 1)
        cell = target_table.cell(plan_row, plan_fill_col)
        cell.text = f"我方计划安排1辆车2人在{cycle}完成施工。"
        for p in cell.paragraphs:
            p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            for r in p.runs: r.font.size = Pt(9)

        save_path = filedialog.asksaveasfilename(
            title="保存会审单", defaultextension=".docx", filetypes=[("Word文件", "*.docx")],
            initialfile=f"{project_name}_会审单.docx"
        )
        if save_path: doc.save(save_path)


if __name__ == "__main__":
    root = tk.Tk()
    app = HomeAndEnterpriseTool(root)
    root.mainloop()