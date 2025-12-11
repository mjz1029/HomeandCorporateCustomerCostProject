import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
from tkcalendar import DateEntry
import pandas as pd
from docx import Document
from docx.shared import Pt, Inches  # 新增：用于图片大小调整
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from PIL import Image  # 用于获取图片尺寸（可选）
import os
import json
from datetime import datetime

# ===================== 配置与常量 =====================
CONFIG_FILE = "config.json"  # 存储可复用基础信息
EXCEL_SHEETS = ["施工项目（Sheet1）", "材料项目（Sheet2）"]  # 两个Sheet的显示名称
MAX_IMG_WIDTH = Inches(4)  # 图片最大宽度（英寸），可调整
MAX_IMG_HEIGHT = Inches(3)  # 图片最大高度（英寸），可调整


class HomeAndEnterpriseTool:
    def __init__(self, root):
        self.root = root
        self.root.title("家集客项目预算与文档生成系统")
        self.root.geometry("1400x900")  # 扩大窗口尺寸
        self.root.minsize(1300, 850)  # 设置最小窗口尺寸，避免缩小后内容溢出

        # 核心数据存储
        self.budget_data = []  # 整合后的预算项目（含两个Sheet）
        self.total_amount = 0.0  # 总金额
        self.base_info = {}  # 可复用基础信息（申请单位、申请人等）
        self.word_app_template = None  # 申请表模板路径
        self.word_review_template = None  # 会审单模板路径
        self.image_paths = []  # 支撑图片

        # 加载配置与预算数据
        self.load_config()
        self.load_budget_excel()

        # 初始化GUI
        self.setup_style()
        self.setup_ui()

    # ===================== 基础配置加载/保存 =====================
    def load_config(self):
        """加载可复用基础信息（config.json）"""
        default_info = {
            "申请单位": "奇台县分公司",
            "申请人": "樊斌",
            "联系电话": "13909949883",
            "实施单位": "中移建设",
            "项目经理": "吴斌",
            "项目经理联系电话": "18899661100",
            "项目负责人": "樊斌"
        }
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                    self.base_info = json.load(f)
                # 补充缺失的默认字段
                for key, val in default_info.items():
                    if key not in self.base_info:
                        self.base_info[key] = val
            except Exception as e:
                messagebox.showwarning("配置加载失败", f"使用默认基础信息：{str(e)}")
                self.base_info = default_info
        else:
            self.base_info = default_info
            self.save_config()

    def save_config(self):
        """保存基础信息到配置文件"""
        try:
            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump(self.base_info, f, ensure_ascii=False, indent=2)
        except Exception as e:
            messagebox.showerror("配置保存失败", str(e))

    # ===================== 预算表加载（修复索引越界：添加数据校验）=====================
    def load_budget_excel(self):
        """加载Excel的两个Sheet，整合为统一项目列表"""
        file_path = filedialog.askopenfilename(
            title="选择家集客预算表",
            filetypes=[("Excel文件", "*.xlsx"), ("所有文件", "*.*")]
        )
        if not file_path:
            messagebox.showwarning("提示", "未选择预算表，应用将无法正常使用！")
            return

        try:
            # 加载Sheet1（施工项目）
            sheet1 = pd.read_excel(file_path, sheet_name=0)
            if sheet1.empty:
                raise ValueError("Sheet1（施工项目）为空")
            sheet1_data = self.parse_sheet1(sheet1)

            # 加载Sheet2（材料项目）
            sheet2 = pd.read_excel(file_path, sheet_name=1)
            if sheet2.empty:
                raise ValueError("Sheet2（材料项目）为空")
            sheet2_data = self.parse_sheet2(sheet2)

            # 整合两个Sheet的数据（添加类别标识）
            self.budget_data = sheet1_data + sheet2_data
            # 重新生成连续ID，避免ID索引混乱
            for idx, item in enumerate(self.budget_data):
                item["id"] = idx + 1
            messagebox.showinfo("加载成功",
                                f"共加载{len(self.budget_data)}个项目（施工{len(sheet1_data)}个+材料{len(sheet2_data)}个）")
        except Exception as e:
            messagebox.showerror("预算表加载失败", f"错误原因：{str(e)}")

    def parse_sheet1(self, df):
        """解析Sheet1（施工项目）- 确保工程量初始值为0.0，添加列数据校验"""
        parsed = []
        # 清理列名
        df.columns = df.columns.str.strip()
        required_cols = ["类别", "折扣后（含税）37%/元", "数量"]
        # 校验列是否存在
        missing_cols = [col for col in required_cols if col not in df.columns]
        if missing_cols:
            raise ValueError(f"Sheet1缺少必要列：{', '.join(missing_cols)}")

        # 遍历数据行（跳过空行）
        for _, row in df.iterrows():
            project_name = str(row["类别"]).strip()
            if not project_name or project_name == "nan":
                continue
            # 单价（折扣后含税）- 确保为数值类型
            unit_price = float(pd.to_numeric(row["折扣后（含税）37%/元"], errors="coerce")) if pd.notna(
                row["折扣后（含税）37%/元"]) else 0.0
            # 工程量（默认0.0，避免NaN）
            quantity = float(pd.to_numeric(row["数量"], errors="coerce")) if pd.notna(row["数量"]) else 0.0
            # 判断是否为长度类项目（单位：元/公里）
            is_length_unit = "元/公里" in project_name
            parsed.append({
                "id": len(parsed) + 1,  # 临时ID，后续会重新生成
                "category": "施工项目",
                "name": project_name,
                "unit": "公里" if is_length_unit else "个/户/处等",
                "unit_price": unit_price,
                "quantity": quantity,
                "total": quantity * unit_price,
                "is_length": is_length_unit
            })
        if not parsed:
            raise ValueError("Sheet1（施工项目）无有效数据")
        return parsed

    def parse_sheet2(self, df):
        """解析Sheet2（材料项目）- 确保工程量初始值为0.0，添加列数据校验"""
        parsed = []
        df.columns = df.columns.str.strip()
        required_cols = ["材料", "含税", "数量"]
        # 校验列是否存在
        missing_cols = [col for col in required_cols if col not in df.columns]
        if missing_cols:
            raise ValueError(f"Sheet2缺少必要列：{', '.join(missing_cols)}")

        # 遍历数据行（跳过空行）
        for _, row in df.iterrows():
            project_name = str(row["材料"]).strip()
            if not project_name or project_name == "nan":
                continue
            # 单价（含税）- 确保为数值类型
            unit_price = float(pd.to_numeric(row["含税"], errors="coerce")) if pd.notna(row["含税"]) else 0.0
            # 工程量（默认0.0，避免NaN）
            quantity = float(pd.to_numeric(row["数量"], errors="coerce")) if pd.notna(row["数量"]) else 0.0
            parsed.append({
                "id": len(parsed) + 1,  # 临时ID，后续会重新生成
                "category": "材料项目",
                "name": project_name,
                "unit": "个",  # 材料默认单位为“个”
                "unit_price": unit_price,
                "quantity": quantity,
                "total": quantity * unit_price,
                "is_length": False
            })
        if not parsed:
            raise ValueError("Sheet2（材料项目）无有效数据")
        return parsed

    # ===================== 样式配置 =====================
    def setup_style(self):
        """GUI样式配置"""
        self.style = ttk.Style(self.root)
        self.style.theme_use("clam")
        # 框架样式
        self.style.configure("Custom.TLabelframe", font=("Arial", 10), foreground="#333")
        self.style.configure("Custom.TLabelframe.Label", font=("Arial", 10, "bold"))
        # 按钮样式
        self.style.configure("Accent.TButton", font=("Arial", 10), background="#4A90E2", foreground="white", padding=4)
        self.style.configure("Generate.TButton", font=("Arial", 11, "bold"), background="#2196F3", foreground="white",
                             padding=6)
        # 表格样式
        self.style.configure("Treeview.Heading", font=("Arial", 9, "bold"), background="#E0E0E0")
        self.style.configure("Treeview", font=("Arial", 8), rowheight=22)
        self.style.map("Treeview", background=[("selected", "#81C784")])

    # ===================== GUI界面布局（核心修复：滚动条+布局适配）=====================
    def setup_ui(self):
        # 1. 基础信息设置区（可复用）
        base_frame = ttk.LabelFrame(self.root, text="📝 基础信息设置（设置后自动复用）", style="Custom.TLabelframe")
        base_frame.pack(fill=tk.X, padx=15, pady=8)

        # 基础信息表单（2列布局）
        info_grid = [
            ("申请单位", "申请单位"), ("申请人", "申请人"),
            ("联系电话", "联系电话"), ("实施单位", "实施单位"),
            ("项目经理", "项目经理"), ("项目经理联系电话", "项目经理联系电话"),
            ("项目负责人", "项目负责人")
        ]
        for i, (label, key) in enumerate(info_grid):
            row = i // 2
            col = i % 2
            ttk.Label(base_frame, text=f"{label}：", font=("Arial", 9)).grid(row=row, column=col * 3, padx=5, pady=5,
                                                                            sticky=tk.W)
            entry = ttk.Entry(base_frame, width=30, font=("Arial", 9))
            entry.grid(row=row, column=col * 3 + 1, padx=5, pady=5)
            entry.insert(0, self.base_info.get(key, ""))
            # 绑定变量存储
            entry.bind("<FocusOut>", lambda e, k=key, ent=entry: self.update_base_info(k, ent.get()))

        # 保存基础信息按钮
        ttk.Button(base_frame, text="💾 保存基础信息", command=self.save_config, style="Accent.TButton").grid(row=4,
                                                                                                             column=0,
                                                                                                             columnspan=6,
                                                                                                             pady=8)

        # 2. 预算表编辑区（标签页：施工项目+材料项目）- 修复滚动条
        budget_frame = ttk.LabelFrame(self.root, text="💰 预算项目编辑（仅工程量>0计入统计）", style="Custom.TLabelframe")
        budget_frame.pack(fill=tk.BOTH, padx=15, pady=5, expand=True)

        # 标签页
        notebook = ttk.Notebook(budget_frame)
        notebook.pack(fill=tk.BOTH, padx=5, pady=5, expand=True)

        # -------------------------- 施工项目标签页（修复滚动条）--------------------------
        construction_tab = ttk.Frame(notebook)
        # 垂直+横向滚动条容器
        construction_canvas = tk.Canvas(construction_tab)
        construction_vscroll = ttk.Scrollbar(construction_tab, orient=tk.VERTICAL, command=construction_canvas.yview)
        construction_hscroll = ttk.Scrollbar(construction_tab, orient=tk.HORIZONTAL, command=construction_canvas.xview)
        construction_scrollable_frame = ttk.Frame(construction_canvas)

        # 绑定滚动事件
        construction_scrollable_frame.bind(
            "<Configure>",
            lambda e: construction_canvas.configure(scrollregion=construction_canvas.bbox("all"))
        )
        construction_canvas.create_window((0, 0), window=construction_scrollable_frame, anchor="nw")
        construction_canvas.configure(yscrollcommand=construction_vscroll.set, xscrollcommand=construction_hscroll.set)

        # 创建施工项目表格
        self.construction_tree = self.create_treeview(construction_scrollable_frame, "施工项目")
        self.construction_tree.pack(fill=tk.BOTH, padx=5, pady=5, expand=True)

        # 布局滚动条
        construction_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        construction_vscroll.pack(side=tk.RIGHT, fill=tk.Y)
        construction_hscroll.pack(side=tk.BOTTOM, fill=tk.X)
        notebook.add(construction_tab, text="施工项目")

        # -------------------------- 材料项目标签页（修复滚动条）--------------------------
        material_tab = ttk.Frame(notebook)
        material_canvas = tk.Canvas(material_tab)
        material_vscroll = ttk.Scrollbar(material_tab, orient=tk.VERTICAL, command=material_canvas.yview)
        material_hscroll = ttk.Scrollbar(material_tab, orient=tk.HORIZONTAL, command=material_canvas.xview)
        material_scrollable_frame = ttk.Frame(material_canvas)

        material_scrollable_frame.bind(
            "<Configure>",
            lambda e: material_canvas.configure(scrollregion=material_canvas.bbox("all"))
        )
        material_canvas.create_window((0, 0), window=material_scrollable_frame, anchor="nw")
        material_canvas.configure(yscrollcommand=material_vscroll.set, xscrollcommand=material_hscroll.set)

        self.material_tree = self.create_treeview(material_scrollable_frame, "材料项目")
        self.material_tree.pack(fill=tk.BOTH, padx=5, pady=5, expand=True)

        material_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        material_vscroll.pack(side=tk.RIGHT, fill=tk.Y)
        material_hscroll.pack(side=tk.BOTTOM, fill=tk.X)
        notebook.add(material_tab, text="材料项目")

        # 总金额显示
        self.total_var = tk.StringVar(value=f"当前总金额：0.00元")
        ttk.Label(budget_frame, textvariable=self.total_var, font=("Arial", 10, "bold"), foreground="#E64A19").pack(
            pady=5)

        # 3. 项目核心信息区（优化布局，避免溢出）
        project_frame = ttk.LabelFrame(self.root, text="📋 项目核心信息", style="Custom.TLabelframe")
        project_frame.pack(fill=tk.X, padx=15, pady=5)

        # 项目名称（加宽输入框）
        ttk.Label(project_frame, text="项目名称：", font=("Arial", 9)).grid(row=0, column=0, padx=5, pady=8, sticky=tk.W)
        self.project_name_var = tk.StringVar(value="广电项目光猫安装、开通")
        ttk.Entry(project_frame, textvariable=self.project_name_var, width=50, font=("Arial", 9)).grid(row=0, column=1,
                                                                                                       padx=5, pady=8)

        # 项目日期
        ttk.Label(project_frame, text="项目日期：", font=("Arial", 9)).grid(row=0, column=2, padx=15, pady=8,
                                                                           sticky=tk.W)
        self.date_entry = DateEntry(project_frame, width=20, background="#4A90E2", foreground="white",
                                    date_pattern="yyyy年MM月dd日", font=("Arial", 9))
        self.date_entry.grid(row=0, column=3, padx=5, pady=8)

        # 计划实施周期
        ttk.Label(project_frame, text="计划实施周期：", font=("Arial", 9)).grid(row=0, column=4, padx=15, pady=8,
                                                                               sticky=tk.W)
        self.cycle_var = tk.StringVar(value="15天")
        ttk.Entry(project_frame, textvariable=self.cycle_var, width=15, font=("Arial", 9)).grid(row=0, column=5, padx=5,
                                                                                                pady=8)

        # 4. 模板与支撑文件区（优化布局）
        template_frame = ttk.LabelFrame(self.root, text="📄 模板与支撑文件", style="Custom.TLabelframe")
        template_frame.pack(fill=tk.X, padx=15, pady=5)

        # 申请表模板（调整列宽适配）
        ttk.Label(template_frame, text="申请表模板：", font=("Arial", 9)).grid(row=0, column=0, padx=5, pady=6,
                                                                              sticky=tk.W)
        self.app_template_var = tk.StringVar(value="未选择")
        ttk.Entry(template_frame, textvariable=self.app_template_var, state="readonly", width=45,
                  font=("Arial", 9)).grid(row=0, column=1, padx=5, pady=6)
        ttk.Button(template_frame, text="浏览", command=lambda: self.select_template("app"),
                   style="Accent.TButton").grid(row=0, column=2, padx=5, pady=6)

        # 会审单模板
        ttk.Label(template_frame, text="会审单模板：", font=("Arial", 9)).grid(row=1, column=0, padx=5, pady=6,
                                                                              sticky=tk.W)
        self.review_template_var = tk.StringVar(value="未选择")
        ttk.Entry(template_frame, textvariable=self.review_template_var, state="readonly", width=45,
                  font=("Arial", 9)).grid(row=1, column=1, padx=5, pady=6)
        ttk.Button(template_frame, text="浏览", command=lambda: self.select_template("review"),
                   style="Accent.TButton").grid(row=1, column=2, padx=5, pady=6)

        # 支撑图片上传（调整位置，避免拥挤）
        ttk.Label(template_frame, text="支撑图片（最多12张）：", font=("Arial", 9)).grid(row=0, column=3, padx=15, pady=6,
                                                                                      sticky=tk.W)
        self.image_count_var = tk.StringVar(value="0张")
        ttk.Label(template_frame, textvariable=self.image_count_var, font=("Arial", 9)).grid(row=0, column=4, padx=5,
                                                                                             pady=6)
        ttk.Button(template_frame, text="上传", command=self.upload_images, style="Accent.TButton").grid(row=0,
                                                                                                         column=5,
                                                                                                         padx=5, pady=6)
        ttk.Button(template_frame, text="清空", command=self.clear_images, style="Accent.TButton").grid(row=0, column=6,
                                                                                                        padx=5, pady=6)

        # 5. 生成按钮
        self.generate_btn = ttk.Button(
            self.root, text="🚀 生成申请表+会审单", command=self.generate_documents, style="Generate.TButton"
        )
        self.generate_btn.pack(pady=15)

        # 状态提示
        self.status_var = tk.StringVar(value="✅ 基础信息已加载，可编辑预算项目工程量（双击表格修改）")
        status_label = ttk.Label(self.root, textvariable=self.status_var, font=("Arial", 9), foreground="#2196F3")
        status_label.pack(pady=5)

        # 刷新表格数据
        self.refresh_treeviews()

    # ===================== 表格创建与刷新（修复NaN显示）=====================
    def create_treeview(self, parent, category):
        """创建预算项目表格（施工/材料）"""
        tree = ttk.Treeview(
            parent,
            columns=["id", "name", "unit_price", "quantity", "total"],
            show="headings",
            selectmode="browse"
        )
        # 表头配置
        tree.heading("id", text="序号")
        tree.heading("name", text="项目名称")
        tree.heading("unit_price", text="单价（元）")
        tree.heading("quantity", text="工程量")
        tree.heading("total", text="合计（元）")
        # 列宽配置（优化列宽，避免内容溢出）
        tree.column("id", width=60)
        tree.column("name", width=450)  # 加宽项目名称列
        tree.column("unit_price", width=100)
        tree.column("quantity", width=100)
        tree.column("total", width=100)
        # 双击修改工程量
        tree.bind("<Double-1>", self.edit_quantity)
        return tree

    def refresh_treeviews(self):
        """刷新两个标签页的表格数据 - 修复NaN显示"""
        # 清空表格
        for item in self.construction_tree.get_children():
            self.construction_tree.delete(item)
        for item in self.material_tree.get_children():
            self.material_tree.delete(item)

        # 填充数据（添加空数据校验）
        if not self.budget_data:
            self.total_var.set(f"当前总金额：0.00元")
            return

        construction_idx = 1
        material_idx = 1
        self.total_amount = 0.0

        for item in self.budget_data:
            # 确保合计金额为数值类型
            total = float(item["quantity"]) * float(item["unit_price"])
            item["total"] = total
            self.total_amount += total

            # 格式化数值，避免NaN和科学计数法
            quantity_str = f"{float(item['quantity']):.2f}" if item["quantity"] is not None else "0.00"
            unit_price_str = f"{float(item['unit_price']):.2f}" if item["unit_price"] is not None else "0.00"
            total_str = f"{total:.2f}"

            # 按类别填充到对应表格
            values = [
                item["id"],
                item["name"],
                unit_price_str,
                quantity_str,
                total_str
            ]
            if item["category"] == "施工项目":
                self.construction_tree.insert("", tk.END, values=values, tags=("construction",))
                construction_idx += 1
            else:
                self.material_tree.insert("", tk.END, values=values, tags=("material",))
                material_idx += 1

        # 更新总金额显示
        self.total_var.set(f"当前总金额：{self.total_amount:.2f}元")

    def edit_quantity(self, event):
        """双击修改工程量 - 确保输入为数值，避免NaN"""
        tree = event.widget
        focus_item = tree.focus()
        if not focus_item:
            return
        # 获取选中行数据（添加索引校验）
        item_values = tree.item(focus_item)["values"]
        if len(item_values) < 4:  # 确保有足够的列数据
            messagebox.showwarning("提示", "选中行数据不完整！")
            return
        try:
            project_id = int(item_values[0])
            current_quantity = float(item_values[3]) if item_values[3] not in ["nan", ""] else 0.0
        except (ValueError, IndexError):
            messagebox.showwarning("提示", "工程量数据异常！")
            return

        # 弹窗输入新工程量（限制为数值）
        new_quantity = simpledialog.askfloat(
            "修改工程量",
            f"项目：{item_values[1]}\n当前工程量：{current_quantity:.2f}\n请输入新工程量（数字）：",
            initialvalue=current_quantity
        )
        if new_quantity is None:  # 用户取消输入
            return
        if new_quantity < 0:
            messagebox.showwarning("警告", "工程量不能为负数！")
            return

        # 更新预算数据（确保为float类型，添加索引校验）
        for item in self.budget_data:
            if item.get("id") == project_id:
                item["quantity"] = float(new_quantity)
                break

        # 刷新表格
        self.refresh_treeviews()
        self.status_var.set(f"✅ 已更新项目工程量：{item_values[1]} → {new_quantity:.2f}")

    # ===================== 模板选择与图片上传 =====================
    def select_template(self, template_type):
        """选择申请表/会审单模板"""
        path = filedialog.askopenfilename(
            title=f"选择{'申请表' if template_type == 'app' else '会审单'}模板",
            filetypes=[("Word文件", "*.docx"), ("所有文件", "*.*")]
        )
        if not path:
            return
        if template_type == "app":
            self.word_app_template = path
            self.app_template_var.set(os.path.basename(path))
        else:
            self.word_review_template = path
            self.review_template_var.set(os.path.basename(path))
        self.status_var.set(f"✅ 已选择{'申请表' if template_type == 'app' else '会审单'}模板：{os.path.basename(path)}")

    def upload_images(self):
        """上传支撑图片"""
        paths = filedialog.askopenfilenames(
            title="选择支撑图片",
            filetypes=[("图片文件", "*.jpg;*.jpeg;*.png;*.bmp"), ("所有文件", "*.*")]
        )
        if paths:
            remaining = 12 - len(self.image_paths)
            if len(paths) > remaining:
                messagebox.showwarning("提示", f"最多上传12张图片，本次仅上传{remaining}张！")
                paths = paths[:remaining]
            self.image_paths.extend(paths)
            self.image_count_var.set(f"{len(self.image_paths)}张")
            self.status_var.set(f"✅ 新增上传{len(paths)}张图片，累计{len(self.image_paths)}张")

    def clear_images(self):
        """清空支撑图片"""
        self.image_paths.clear()
        self.image_count_var.set("0张")
        self.status_var.set("✅ 已清空所有支撑图片")

    # ===================== 基础信息更新 =====================
    def update_base_info(self, key, value):
        """更新基础信息（失去焦点时触发）"""
        self.base_info[key] = value.strip()
        self.status_var.set(f"✅ 已更新{key}：{self.base_info[key]}（需点击保存按钮生效）")

    # ===================== 工作量清单生成 =====================
    def generate_work_list(self):
        """生成工作量及材料清单（匹配示例格式）"""
        work_list = []

        # 收集有效项目（工程量>0）
        for item in self.budget_data:
            if item["quantity"] <= 0:
                continue
            quantity = float(item["quantity"])
            if item["is_length"]:
                # 长度类项目：X公里 项目名称（简化显示，匹配示例）
                item_str = f"{quantity:.2f}公里 {item['name'].split('，')[0] if '，' in item['name'] else item['name']}"
            else:
                # 其他项目：X单位 项目名称
                item_str = f"{quantity:.2f}{item['unit']} {item['name'].split('，')[0] if '，' in item['name'] else item['name']}"

            work_list.append(item_str)

        # 匹配示例格式：用逗号连接，无分类前缀
        return "，".join(work_list) if work_list else "无有效项目"

    # ===================== 新增：图片插入辅助方法 =====================
    def insert_images_to_cell(self, cell, image_paths):
        """将图片插入到指定单元格中，自动调整大小"""
        if not image_paths:
            return
        # 清空单元格原有内容（可选）
        cell.text = ""
        # 遍历图片路径，依次插入
        for img_path in image_paths:
            try:
                # 添加段落，插入图片
                para = cell.add_paragraph()
                run = para.add_run()
                # 插入图片并调整大小
                img = run.add_picture(img_path, width=MAX_IMG_WIDTH, height=MAX_IMG_HEIGHT)
                # 居中对齐（可选）
                para.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            except Exception as e:
                messagebox.warning("图片插入失败", f"图片{os.path.basename(img_path)}插入失败：{str(e)}")

    def find_cell_by_text(self, table, keyword):
        """在表格中查找包含指定关键词的单元格，返回(行索引, 列索引, 单元格)"""
        for row_idx, row in enumerate(table.rows):
            for col_idx, cell in enumerate(row.cells):
                if keyword in cell.text.strip():
                    return (row_idx, col_idx, cell)
        return (None, None, None)

    # ===================== Word文档生成（核心修改：插入图片到“其他需求支撑文件”后一格）=====================
    def generate_documents(self):
        """生成申请表和会审单"""
        # 前置校验
        if not self.word_app_template or not self.word_review_template:
            messagebox.showwarning("提示", "请先选择申请表和会审单模板！")
            return
        if self.total_amount <= 0:
            messagebox.showwarning("提示", "无有效项目（请设置工程量>0的项目）！")
            return
        project_name = self.project_name_var.get().strip()
        if not project_name:
            messagebox.showwarning("提示", "请输入项目名称！")
            return
        project_date = self.date_entry.get()
        cycle = self.cycle_var.get().strip()
        if not cycle:
            messagebox.showwarning("提示", "请输入计划实施周期！")
            return

        try:
            # 生成工作量清单（匹配示例格式）
            work_list = self.generate_work_list()

            # 填充申请表
            self.fill_application_form(project_name, project_date, cycle, work_list)

            # 填充会审单
            self.fill_review_form(project_name, project_date, cycle, work_list)

            messagebox.showinfo("生成成功",
                                f"✅ 两个文档已生成完成！\n总金额：{self.total_amount:.2f}元\n已插入{len(self.image_paths)}张支撑图片")
            self.status_var.set(f"🎉 生成成功！总金额：{self.total_amount:.2f}元，插入{len(self.image_paths)}张图片")
        except IndexError as e:
            messagebox.showerror("生成失败",
                                 f"错误原因：模板表格行列索引越界（你的模板表格行列数与代码不匹配）\n详细错误：{str(e)}")
            self.status_var.set(f"❌ 生成失败：模板表格索引越界")
        except Exception as e:
            messagebox.showerror("生成失败", f"错误原因：{str(e)}")
            self.status_var.set(f"❌ 生成失败：{str(e)}")

    def fill_application_form(self, project_name, project_date, cycle, work_list):
        """填充申请表（核心修改：插入图片到“其他需求支撑文件”后一格）"""
        doc = Document(self.word_app_template)
        # 校验表格是否存在
        if not doc.tables:
            raise ValueError("申请表模板中未找到表格！")
        target_table = doc.tables[0]  # 申请表核心表格

        # 定义需要填充的内容（行, 列, 内容, 对齐方式）
        fill_items = [
            (0, 1, self.base_info["申请单位"], WD_PARAGRAPH_ALIGNMENT.LEFT),
            (0, 3, project_date, WD_PARAGRAPH_ALIGNMENT.CENTER),
            (0, 4, project_date, WD_PARAGRAPH_ALIGNMENT.CENTER),
            (0, 6, self.base_info["申请人"], WD_PARAGRAPH_ALIGNMENT.LEFT),
            (1, 1, project_name, WD_PARAGRAPH_ALIGNMENT.CENTER),
            (1, 2, project_name, WD_PARAGRAPH_ALIGNMENT.CENTER),
            (1, 3, project_name, WD_PARAGRAPH_ALIGNMENT.CENTER),
            (1, 4, project_name, WD_PARAGRAPH_ALIGNMENT.CENTER),
            (1, 6, self.base_info["联系电话"], WD_PARAGRAPH_ALIGNMENT.LEFT),
            (2, 1, cycle, WD_PARAGRAPH_ALIGNMENT.LEFT),
            (2, 3, f"{self.total_amount:.2f}元", WD_PARAGRAPH_ALIGNMENT.CENTER),
            (2, 4, f"{self.total_amount:.2f}元", WD_PARAGRAPH_ALIGNMENT.CENTER),
        ]

        # 填充基础信息（添加索引校验）
        for row_idx, col_idx, text, align in fill_items:
            try:
                cell = target_table.cell(row_idx, col_idx)
                cell.text = text
                for para in cell.paragraphs:
                    para.alignment = align
            except IndexError:
                raise IndexError(f"申请表表格缺少行{row_idx}列{col_idx}的单元格")

        # 填充主要工作量及材料清单（动态查找行，避免固定索引）
        list_row_idx = None
        for idx, row in enumerate(target_table.rows):
            cell_text = "".join([cell.text for cell in row.cells]).strip()
            if "工作量" in cell_text or "清单" in cell_text:
                list_row_idx = idx
                break
        if list_row_idx is None:
            list_row_idx = max(0, len(target_table.rows) - 3)
        # 填充清单内容（跨列填充）
        for col_idx in range(1, min(6, len(target_table.columns))):
            try:
                cell = target_table.cell(list_row_idx, col_idx)
                cell.text = work_list
                for para in cell.paragraphs:
                    para.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
                    para.runs[0].font.size = Pt(9)
            except IndexError:
                pass

        # ===================== 核心修改：插入图片到“其他需求支撑文件”后一格 =====================
        # 1. 查找包含“其他需求支撑文件”的单元格
        keyword = "其他需求支撑文件"
        row_idx, col_idx, _ = self.find_cell_by_text(target_table, keyword)
        if row_idx is not None and col_idx is not None:
            # 2. 确定后一格的列索引（col_idx + 1），若超出列数则取最后一列
            target_col_idx = col_idx + 1 if (col_idx + 1) < len(target_table.columns) else len(target_table.columns) - 1
            try:
                # 3. 获取后一格的单元格，插入图片
                target_cell = target_table.cell(row_idx, target_col_idx)
                self.insert_images_to_cell(target_cell, self.image_paths)
            except IndexError:
                # 若后一格不存在，则插入到当前行的最后一列
                target_cell = target_table.cell(row_idx, len(target_table.columns) - 1)
                self.insert_images_to_cell(target_cell, self.image_paths)
        else:
            # 若未找到关键词，默认插入到表格倒数第2行的最后一列（可选）
            target_cell = target_table.cell(max(0, len(target_table.rows) - 2), len(target_table.columns) - 1)
            self.insert_images_to_cell(target_cell, self.image_paths)
            messagebox.showwarning("提示", "申请表模板中未找到“其他需求支撑文件”单元格，图片已插入到表格默认位置")

        # 保存申请表
        save_path = filedialog.asksaveasfilename(
            title="保存申请表",
            defaultextension=".docx",
            filetypes=[("Word文件", "*.docx")],
            initialfile=f"{project_name}_申请表_{project_date.replace('年', '').replace('月', '').replace('日', '')}.docx"
        )
        if save_path:
            doc.save(save_path)

    def fill_review_form(self, project_name, project_date, cycle, work_list):
        """填充会审单（核心修改：插入图片到“其他需求支撑文件”后一格）"""
        doc = Document(self.word_review_template)
        # 校验表格是否存在
        if not doc.tables:
            raise ValueError("会审单模板中未找到表格！")
        target_table = doc.tables[0]  # 会审单核心表格

        # 定义需要填充的内容（行, 列, 内容, 对齐方式）
        fill_items = [
            (0, 1, project_name, WD_PARAGRAPH_ALIGNMENT.CENTER),
            (1, 1, f"{self.total_amount:.2f}元", WD_PARAGRAPH_ALIGNMENT.CENTER),
            (1, 5, project_date, WD_PARAGRAPH_ALIGNMENT.CENTER),
            (1, 9, cycle, WD_PARAGRAPH_ALIGNMENT.CENTER),
            (2, 1, self.base_info["项目负责人"], WD_PARAGRAPH_ALIGNMENT.CENTER),
            (2, 5, self.base_info["联系电话"], WD_PARAGRAPH_ALIGNMENT.CENTER),
            (3, 1, self.base_info["实施单位"], WD_PARAGRAPH_ALIGNMENT.CENTER),
            (3, 5, self.base_info["项目经理"], WD_PARAGRAPH_ALIGNMENT.CENTER),
            (3, 9, self.base_info["项目经理联系电话"], WD_PARAGRAPH_ALIGNMENT.CENTER),
        ]

        # 填充基础信息（添加索引校验，扩大列范围）
        for row_idx, col_idx, text, align in fill_items:
            try:
                # 跨列填充（比如项目名称填充到0行所有列）
                if row_idx == 0:  # 项目名称行：填充所有列
                    for c in range(1, min(len(target_table.columns), 11)):
                        cell = target_table.cell(row_idx, c)
                        cell.text = text
                        for para in cell.paragraphs:
                            para.alignment = align
                elif row_idx == 1 and col_idx == 1:  # 预算金额行：填充1-3列
                    for c in range(1, 4):
                        cell = target_table.cell(row_idx, c)
                        cell.text = text
                        for para in cell.paragraphs:
                            para.alignment = align
                elif row_idx == 1 and col_idx == 5:  # 会审日期行：填充5-7列
                    for c in range(5, 8):
                        cell = target_table.cell(row_idx, c)
                        cell.text = text
                        for para in cell.paragraphs:
                            para.alignment = align
                elif row_idx == 1 and col_idx == 9:  # 实施周期行：填充9-10列
                    for c in range(9, min(11, len(target_table.columns))):
                        cell = target_table.cell(row_idx, c)
                        cell.text = text
                        for para in cell.paragraphs:
                            para.alignment = align
                else:
                    cell = target_table.cell(row_idx, col_idx)
                    cell.text = text
                    for para in cell.paragraphs:
                        para.alignment = align
            except IndexError:
                raise IndexError(f"会审单表格缺少行{row_idx}列{col_idx}的单元格")

        # 填充工作量及材料清单（动态查找行，避免固定索引）
        list_row_idx = None
        for idx, row in enumerate(target_table.rows):
            cell_text = "".join([cell.text for cell in row.cells]).strip()
            if "工作量" in cell_text or "清单" in cell_text:
                list_row_idx = idx
                break
        if list_row_idx is None:
            list_row_idx = max(0, len(target_table.rows) - 2)
        # 填充清单内容（跨列填充，添加前缀）
        work_list_with_prefix = f"工作量：{work_list}"
        for col_idx in range(1, min(len(target_table.columns), 11)):
            try:
                cell = target_table.cell(list_row_idx, col_idx)
                cell.text = work_list_with_prefix
                for para in cell.paragraphs:
                    para.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
                    para.runs[0].font.size = Pt(9)
            except IndexError:
                pass

        # 填充施工方实施计划（清单行下一行）
        plan_row_idx = list_row_idx + 1
        plan_text = f"我方计划安排1辆车2人在{cycle}完成施工。"
        try:
            for col_idx in range(1, min(len(target_table.columns), 11)):
                cell = target_table.cell(plan_row_idx, col_idx)
                cell.text = plan_text
                for para in cell.paragraphs:
                    para.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        except IndexError:
            pass

        # 保存会审单
        save_path = filedialog.asksaveasfilename(
            title="保存会审单",
            defaultextension=".docx",
            filetypes=[("Word文件", "*.docx")],
            initialfile=f"{project_name}_会审单_{project_date.replace('年', '').replace('月', '').replace('日', '')}.docx"
        )
        if save_path:
            doc.save(save_path)


if __name__ == "__main__":
    root = tk.Tk()
    app = HomeAndEnterpriseTool(root)
    root.mainloop()