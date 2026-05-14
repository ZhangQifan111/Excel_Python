"""
Excel数据可视化工具 v2.9.2 - 用户体验优化版本
功能：
1. 选择Excel文件（带进度条显示）
2. 动态显示列名勾选框
3. 生成折线图（横轴固定为第一列时间数据）
4. 多线程处理避免界面卡顿
5. 滚轮缩放功能（仅图表区域）
6. 常用分组管理（新建、保存、删除、重命名分组，支持数据持久化）
7. 故障数据分析模式（解析故障码、统计分析）
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure
import matplotlib.dates as mdates
from datetime import datetime
import threading
import time
import queue
import json
import os

# 导入故障码解析模块
try:
    from fault_code_parser import (
        parse_fault_value, 
        parse_bit_list,
        parse_register_bits,
        get_fault_info,
        get_register_name,
        get_fault_category,
        get_supported_registers,
        load_fault_codes_from_excel,
        load_fault_codes_from_point_table,
        USER_FAULT_CODES,
        REGISTER_NAMES,
        MAIN_FAULT_CODES,
        ALL_FAULT_CODES
    )
    FAULT_PARSER_AVAILABLE = True
    print("[INFO] 故障码解析模块加载成功")
except ImportError as e:
    print(f"[ERROR] 故障码解析模块导入失败: {e}")
    FAULT_PARSER_AVAILABLE = False


class CollapsibleFrame(ttk.Frame):
    """可折叠面板组件
    
    功能：
    - 点击标题栏可以展开/折叠内容
    - 折叠时只显示标题栏，节省空间
    - 支持默认展开/折叠状态
    """
    
    def __init__(self, parent, title="", collapsed=False, **kwargs):
        """
        Args:
            parent: 父容器
            title: 面板标题
            collapsed: 初始是否折叠
        """
        super().__init__(parent, **kwargs)
        
        self.is_collapsed = collapsed
        self.title = title
        
        # 标题栏框架（使用 tk.Frame 支持 background 属性）
        self.header_frame = tk.Frame(self, bg="#f0f0f0", cursor="hand2")
        self.header_frame.pack(fill=tk.X)
        
        # 折叠图标和标题（使用 tk.Label 支持 background 属性）
        self.toggle_btn = tk.Label(
            self.header_frame, 
            text="▼ " if not collapsed else "▶ ", 
            cursor="hand2",
            font=("Arial", 10, "bold"),
            bg="#f0f0f0"
        )
        self.toggle_btn.pack(side=tk.LEFT)
        
        self.title_label = tk.Label(
            self.header_frame, 
            text=title, 
            font=("Arial", 10, "bold"),
            cursor="hand2",
            bg="#f0f0f0"
        )
        self.title_label.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        # 内容框架
        self.content_frame = ttk.Frame(self)
        if not collapsed:
            self.content_frame.pack(fill=tk.BOTH, expand=True, pady=(5, 0))
        
        # 绑定点击事件
        self.toggle_btn.bind("<Button-1>", self.toggle)
        self.title_label.bind("<Button-1>", self.toggle)
        self.header_frame.bind("<Button-1>", self.toggle)
        
        # 鼠标悬停效果
        self.header_frame.bind("<Enter>", self._on_enter)
        self.header_frame.bind("<Leave>", self._on_leave)
        
    def _on_enter(self, event):
        """鼠标悬停效果"""
        self.header_frame.configure(bg="#e0e0e0")
        self.toggle_btn.configure(bg="#e0e0e0")
        self.title_label.configure(bg="#e0e0e0")
        
    def _on_leave(self, event):
        """鼠标离开恢复"""
        self.header_frame.configure(bg="#f0f0f0")
        self.toggle_btn.configure(bg="#f0f0f0")
        self.title_label.configure(bg="#f0f0f0")
    
    def toggle(self, event=None):
        """切换折叠状态"""
        self.is_collapsed = not self.is_collapsed
        
        if self.is_collapsed:
            # 折叠：隐藏内容
            self.content_frame.pack_forget()
            self.toggle_btn.configure(text="▶ ")
        else:
            # 展开：显示内容
            self.content_frame.pack(fill=tk.BOTH, expand=True, pady=(5, 0))
            self.toggle_btn.configure(text="▼ ")
    
    def expand(self):
        """展开面板"""
        if self.is_collapsed:
            self.toggle()
    
    def collapse(self):
        """折叠面板"""
        if not self.is_collapsed:
            self.toggle()
    
    def get_content_frame(self):
        """获取内容框架，用于添加子组件"""
        return self.content_frame


class ExcelDataVisualizer:
    def load_custom_columns(self):
        """加载自定义计算列配置"""
        try:
            if os.path.exists(self.custom_columns_file):
                with open(self.custom_columns_file, 'r', encoding='utf-8') as f:
                    self.custom_columns = json.load(f)
        except Exception as e:
            print(f"加载自定义计算列配置失败: {e}")
            self.custom_columns = {}
    
    def save_custom_columns(self):
        """保存自定义计算列配置"""
        try:
            with open(self.custom_columns_file, 'w', encoding='utf-8') as f:
                json.dump(self.custom_columns, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"保存自定义计算列配置失败: {e}")
    
    def _apply_custom_columns(self):
        """重新应用自定义计算列到当前数据"""
        if self.df is None:
            return
        
        print(f"[DEBUG] _apply_custom_columns: 开始重新应用 {len(self.custom_columns)} 个自定义列")
        
        for col_name, col_config in self.custom_columns.items():
            # 如果列已存在，跳过
            if col_name in self.df.columns:
                continue
            
            formula = col_config.get('formula', '')
            if not formula:
                continue
            
            # 获取可用列
            available_columns = [c for c in self.df.columns if c != self.time_column]
            
            # 提取公式中的列名
            formula_columns = []
            sorted_columns = sorted(available_columns, key=len, reverse=True)
            remaining_formula = formula
            
            for col in sorted_columns:
                if col in remaining_formula:
                    formula_columns.append(col)
                    remaining_formula = remaining_formula.replace(col, '')
            
            # 检查是否有缺失的列
            missing_columns = [col for col in formula_columns if col not in self.df.columns]
            if missing_columns:
                print(f"[DEBUG] 自定义列 {col_name} 缺少列: {missing_columns}")
                continue
            
            # 尝试计算
            try:
                calc_expr = formula
                for col in formula_columns:
                    calc_expr = calc_expr.replace(col, f"self.df['{col}']")
                self.df[col_name] = eval(calc_expr)
                print(f"[DEBUG] 自定义列 {col_name} 已应用")
            except Exception as e:
                print(f"[DEBUG] 自定义列 {col_name} 计算失败: {e}")
        
        # 刷新列列表
        if hasattr(self, 'refresh_column_list'):
            self.refresh_column_list()
    




    def load_groups(self):
        """加载常用分组数据"""
        try:
            if os.path.exists(self.groups_file):
                with open(self.groups_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    self.groups = data.get('groups', {})
                    self.group_columns = data.get('group_columns', {})
                print(f"成功加载 {len(self.groups)} 个分组")
                self.update_group_combobox()
            else:
                self.groups = {}
                self.group_columns = {}
        except Exception as e:
            print(f"加载分组数据失败: {e}")
            self.groups = {}
            self.group_columns = {}
    
    def __init__(self, root):
        self.root = root
        self.root.title("Excel数据可视化工具 v2.9.0")
        self.root.geometry("1700x1000")  # 窗口初始尺寸
        
        # 【修复】设置窗口最小尺寸，避免在小屏幕电脑上显示异常
        self.root.minsize(1400, 800)  # 增大最小尺寸确保所有内容可见
        
        # 【修复】让窗口居中显示
        self.root.update_idletasks()  # 更新窗口信息
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        window_width = min(1700, screen_width - 50)  # 不超过屏幕宽度，留50像素边距
        window_height = min(1000, screen_height - 50)  # 不超过屏幕高度
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        self.root.geometry(f"{window_width}x{window_height}+{x}+{y}")
        
        # 数据存储
        self.df = None
        self.time_column = None
        self.column_vars = {} # 存储每列的勾选状态
        self.selected_columns = []
        
        # 常用分组相关数据存储
        self.groups = {}  # 存储所有常用分组，格式: {组名: {列名: 勾选状态}}
        self.current_group = None  # 当前选中的分组
        self.group_columns = {}  # 分组与列的映射关系
        self.groups_file = "common_groups.json"  # 分组数据存储文件
        self.selected_columns = []
        
        # 自定义计算列相关数据存储
        self.custom_columns = {}  # 存储所有自定义计算列，格式: {列名: {formula: "", description: ""}}
        self.custom_columns_file = "custom_columns.json"  # 自定义计算列数据存储文件
        
        # 故障配置相关数据存储
        self.fault_config_done = False  # 是否已完成故障配置（首次切换时弹出配置对话框）
        self.fault_source = "builtin"  # 故障码来源: "builtin"(内置) 或 "external"(外部文件)
        self.fault_external_file = ""  # 外部故障码文件路径
        self.selected_fault_registers = set()  # 选中的故障寄存器（已废弃，改用 selected_fault_columns）
        self.selected_fault_columns = []  # 选中的故障寄存器列名
        self.detected_registers = []  # 检测到的故障寄存器列表 [(reg_num, col_name, type), ...]
        self.register_types = {}  # 列名 -> 类型映射
        self.fault_register_vars = {}  # 故障寄存器复选框变量
        self.fault_register_checkboxes = {}  # 故障寄存器复选框控件
        
        # 数据诊断相关数据存储
        self.diagnosis_results = None  # 存储诊断结果
        self.diagnosis_strategies = {}  # 存储自定义诊断策略
        
        # 【新增】线条可见性状态和线条对象存储
        self.line_visibility = {}  # {列名: True/False} 线条是否可见
        self.line_objects = {}  # {列名: (line_object, marker_object)} 存储线条对象
        self.legend_items = {}  # {列名: legend_frame} 存储图例项Frame
        
        # 线程通信队列
        self.queue = queue.Queue()
        
        # 设置中文显示
        plt.rcParams['font.sans-serif'] = ['SimHei', 'Microsoft YaHei', 'Arial Unicode MS']
        plt.rcParams['axes.unicode_minus'] = False
        
        # 创建菜单栏
        self._create_menu_bar()
        
        # 创建界面
        self.create_widgets()
        
        # 加载常用分组数据
        self.load_groups()
        
        # 加载自定义计算列数据
        self.load_custom_columns()
        
        # 检查队列消息
        self.root.after(100, self.process_queue)
        
    def _init_sash_position(self, paned_window, position, retry_count=0):
        """初始化分割条位置（带延迟和重试机制）
        
        Args:
            paned_window: PanedWindow组件
            position: 分割条位置（像素）
            retry_count: 重试次数
        """
        max_retries = 10
        initial_delay = 100  # 初始延迟100ms
        
        def try_set():
            try:
                current_pos = paned_window.sashpos(0)
                # 如果当前位置接近0或接近窗口右边界，说明需要设置
                if current_pos < 50 or current_pos > 1000:
                    paned_window.sashpos(0, position)
                    print(f"[DEBUG] 设置分割条位置: {position}")
                else:
                    print(f"[DEBUG] 分割条位置已设置: {current_pos}")
                    return  # 位置已正确，无需继续
            except Exception as e:
                print(f"[DEBUG] 设置分割条位置失败: {e}")
            
            # 如果还没达到最大重试次数，继续重试
            if retry_count < max_retries - 1:
                self.root.after(100, lambda: self._init_sash_position(paned_window, position, retry_count + 1))
        
        # 首次延迟执行
        self.root.after(initial_delay + retry_count * 50, try_set)
    
    def _init_chart_paned_sash(self):
        """初始化图表-图例分割条位置"""
        def try_set(retry=0):
            try:
                if hasattr(self, 'chart_paned') and self.chart_paned:
                    # 获取总宽度，设置图例面板占约 15%
                    total_width = self.chart_paned.winfo_width()
                    if total_width > 100:
                        # 图表占 85%，图例占 15%
                        chart_width = int(total_width * 0.82)
                        try:
                            self.chart_paned.sashpos(0, chart_width)
                            print(f"[DEBUG] 图表-图例分割条位置: {chart_width} (总宽度: {total_width})")
                        except Exception as e:
                            print(f"[DEBUG] sashpos 失败: {e}")
                        return
            except Exception as e:
                print(f"[DEBUG] 设置图表-图例分割条失败: {e}")
            
            # 重试最多5次
            if retry < 5:
                self.root.after(200, lambda: try_set(retry + 1))
        
        # 延迟执行，等待窗口渲染
        self.root.after(300, try_set)
    
    def _on_tab_changed(self, event):
        """Tab 切换事件处理"""
        try:
            current_tab = self.main_notebook.index(self.main_notebook.select())
            # current_tab: 0 = 单文件分析, 1 = 批量分析
            
            if current_tab == 1 and hasattr(self, 'batch_paned') and not getattr(self, 'batch_sash_initialized', False):
                # 切换到批量分析 Tab 时，重新设置分割条位置
                # 左侧面板约350像素，右侧区域更大用于显示异常列表
                self._init_sash_position(self.batch_paned, 350, 0)
                self.batch_sash_initialized = True
                print("[DEBUG] Tab 切换到批量分析，重新设置分割条位置")
        except Exception as e:
            print(f"[DEBUG] Tab 切换处理失败: {e}")
    
    def _on_analysis_mode_change(self):
        """分析模式切换事件处理"""
        mode = self.analysis_mode.get()
        print(f"[DEBUG] 分析模式切换: {mode}")
        
        if mode == "normal":
            # 普通数据分析模式
            self.progress_label.config(text="就绪", foreground="blue")
            # 清空当前数据，重新选择文件
            self._reset_for_mode_change()
            
        elif mode == "fault":
            # 故障数据分析模式
            self.progress_label.config(text="故障分析模式", foreground="orange")
            # 清空当前数据，重新选择文件
            self._reset_for_mode_change()
    
    def _reset_for_mode_change(self):
        """模式切换时重置界面状态"""
        # 清空数据
        self.df = None
        self.time_column = None
        self.selected_columns = []
        self.column_vars = {}
        
        # 重置文件选择
        self.file_path_label.config(text="未选择文件")
        
        # 清空列选择
        if hasattr(self, 'checkbox_frame'):
            for widget in self.checkbox_frame.winfo_children():
                widget.destroy()
        
        # 清空图表
        if hasattr(self, 'canvas'):
            self.ax.clear()
            self.ax.set_xlabel("时间")
            self.ax.set_ylabel("数值")
            self.ax.set_title("数据趋势图")
            self.canvas.draw()
        
        # 清空图例
        if hasattr(self, 'legend_frame'):
            for widget in self.legend_frame.winfo_children():
                widget.destroy()
        
        # 重置诊断结果
        self.diagnosis_results = None

    def set_sash_position(self, paned_window, position):
        """设置分割条的初始位置
        
        Args:
            paned_window: PanedWindow组件
            position: 分割条位置（像素）
        """
        try:
            # 尝试使用sashpos方法设置分割条位置
            paned_window.sashpos(0, position)
        except AttributeError:
            # 如果Tkinter版本不支持sashpos，使用替代方法
            self._try_set_sash_retry(paned_window, position, 0)
            
    def _try_set_sash_retry(self, paned_window, position, retry_count=0):
        """尝试设置分割条位置（辅助方法，带重试）
        
        Args:
            paned_window: PanedWindow组件
            position: 分割条位置（像素）
            retry_count: 重试次数
        """
        try:
            paned_window.sashpos(0, position)
        except:
            if retry_count < 5: # 最多重试5次
                self.root.after(50, lambda: self._try_set_sash_retry(paned_window, position, retry_count + 1))

    def insert_to_formula(self, entry, text):
        """插入文本到公式输入框"""
        current_text = entry.get()
        cursor_pos = entry.index(tk.INSERT)
        new_text = current_text[:cursor_pos] + text + current_text[cursor_pos:]
        entry.delete(0, tk.END)
        entry.insert(0, new_text)
        entry.focus()
    
    def validate_formula(self, name, formula, available_columns):
        """验证公式是否正确"""
        if not name.strip():
            messagebox.showerror("错误", "请输入列名称!")
            return False

        if not formula.strip():
            messagebox.showerror("错误", "请输入计算公式!")
            return False

        # 检查列名是否存在（使用更智能的匹配方式）
        # 提取公式中的所有列名（按最长匹配优先）
        used_columns = []
        remaining_formula = formula

        # 按列名长度降序排序，优先匹配长列名
        sorted_columns = sorted(available_columns, key=len, reverse=True)

        # 查找所有在公式中出现的列名
        found_columns = []
        for col in sorted_columns:
            if col in remaining_formula:
                found_columns.append(col)
                remaining_formula = remaining_formula.replace(col, '')

        # 检查是否还有未识别的非数字、非运算符的内容
        # 移除运算符、括号、空格、数字
        import re
        cleaned_remaining = re.sub(r'[+\-*/() ]', '', remaining_formula)

        if cleaned_remaining:
            # 还有未识别的内容，可能是错误的列名
            messagebox.showerror("错误", f"未找到数据列: {cleaned_remaining}\n\n可用列名:\n" + "\n".join(available_columns[:10]))
            return False

        # 尝试计算验证
        try:
            # 创建测试数据
            test_data = {}
            for col in available_columns:
                test_data[col] = 10  # 使用固定值测试

            # 替换列名为测试值（按最长匹配优先）
            test_formula = formula
            for col in sorted_columns:
                test_formula = test_formula.replace(col, str(test_data[col]))

            # 计算结果
            result = eval(test_formula, {"__builtins__": None}, {})

            # 显示找到的列名
            if found_columns:
                columns_info = "\n使用的列:\n" + "\n".join(f"  • {col}" for col in found_columns)
            else:
                columns_info = ""

            messagebox.showinfo("验证成功", f"公式验证通过!\n测试结果: {result}{columns_info}")
            return True
        except Exception as e:
            messagebox.showerror("错误", f"公式验证失败: {e}")
            return False
    
    def refresh_column_list(self):
        """刷新数据列选择列表（使用虚拟列表优化）"""
        if self.df is None:
            print("[DEBUG] refresh_column_list: df is None, 跳过刷新")
            return

        print(f"[DEBUG] refresh_column_list: 开始刷新，当前df有 {len(self.df.columns)} 列")
        print(f"[DEBUG] refresh_column_list: 列名: {list(self.df.columns)}")

        # 【修复】只清空 scrollable_frame 中的复选框，不清空搜索框和按钮
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()

        # 清空复选框状态字典
        self.column_vars.clear()

        # 获取所有数据列（排除时间列）
        all_columns = [col for col in self.df.columns if col != self.time_column]
        print(f"[DEBUG] refresh_column_list: 排除时间列后剩余 {len(all_columns)} 列")

        # 【性能优化】先创建所有变量（很快）
        for col in all_columns:
            self.column_vars[col] = tk.IntVar(value=0)

        # 【性能优化】使用虚拟列表方式 - 只创建可见区域的复选框
        total_columns = len(all_columns)
        
        # 检查滚动区域的高度，估算可见行数
        try:
            visible_height = self.scrollable_frame.winfo_height()
            if visible_height < 100:
                visible_height = 400  # 默认高度
        except:
            visible_height = 400
        
        # 每个复选框约20像素高，估算可见数量
        visible_count = min(int(visible_height / 20) + 10, total_columns)
        
        print(f"[PERF] refresh_column_list: 可见区域高度: {visible_height}px, 预估可见: {visible_count} 个")
        
        # 只创建前N个复选框
        for i, col in enumerate(all_columns[:visible_count]):
            var = self.column_vars[col]
            cb = ttk.Checkbutton(
                self.scrollable_frame,
                text=col,
                variable=var
            )
            cb.pack(anchor=tk.W, padx=5, pady=2)
        
        # 存储待创建的列，滚动时动态创建
        self._pending_columns = all_columns[visible_count:]
        self._created_column_count = visible_count

        # 【修复】如果当前有选中的分组，根据分组数据恢复勾选状态
        if self.current_group and self.current_group in self.groups:
            group_data = self.groups[self.current_group]
            print(f"[DEBUG] refresh_column_list: 当前分组 '{self.current_group}' 有 {len(group_data)} 列")
            for col, checked in group_data.items():
                if col in self.column_vars:
                    self.column_vars[col].set(checked)
                    print(f"[DEBUG] refresh_column_list: 恢复勾选状态 - {col}: {checked}")
                else:
                    print(f"[DEBUG] refresh_column_list: 警告 - 列 '{col}' 不在 column_vars 中")

        print(f"[INFO] 已刷新数据列列表，共 {len(all_columns)} 列，已勾选 {sum(1 for v in self.column_vars.values() if v.get() == 1)} 列")
    

    def create_custom_column(self, name, formula, available_columns, editor_window):
        """创建自定义计算列"""
        print(f"\n[DEBUG] ========== 开始创建自定义列 ==========")
        print(f"[DEBUG] 列名称: {name}")
        print(f"[DEBUG] 公式: {formula}")
        print(f"[DEBUG] 可用列数量: {len(available_columns)}")
        
        if not name.strip():
            messagebox.showerror("错误", "请输入列名称!")
            return
        
        if not formula.strip():
            messagebox.showerror("错误", "请输入计算公式!")
            return
        
        # 生成唯一ID
        col_id = f"(CALC{len(self.custom_columns) + 1:03d})"
        full_name = f"{name}{col_id}"
        print(f"[DEBUG] 完整列名: {full_name}")
        
        # 计算新列数据
        try:
            print(f"[DEBUG] 当前 df 列数: {len(self.df.columns)}")
            print(f"[DEBUG] df 列名: {list(self.df.columns)[:10]}...")  # 只显示前10个
            
            # 构建计算表达式
            calc_expr = formula
            print(f"[DEBUG] 原始公式: {formula}")
            
            for col in available_columns:
                if col in formula:
                    print(f"[DEBUG] 找到列引用: {col}")
                    calc_expr = calc_expr.replace(col, f"self.df['{col}']")
            
            print(f"[DEBUG] 计算表达式: {calc_expr}")
            
            # 执行计算
            print(f"[DEBUG] 开始执行计算...")
            self.df[full_name] = eval(calc_expr)
            print(f"[DEBUG] 计算完成，结果列名: {full_name}")
            print(f"[DEBUG] 计算后 df 列数: {len(self.df.columns)}")
            print(f"[DEBUG] 新列数据预览 (前5行): {self.df[full_name].head().tolist()}")
            
            # 检查是否有空值
            if self.df[full_name].isnull().any():
                null_count = self.df[full_name].isnull().sum()
                print(f"[DEBUG] 检测到 {null_count} 个空值")
                response = messagebox.askyesno("警告", 
                    f"计算结果中有 {null_count} 个空值!\n是否继续?")
                if not response:
                    del self.df[full_name]
                    print(f"[DEBUG] 用户取消，删除列 {full_name}")
                    return
            
            # 保存计算列信息
            self.custom_columns[full_name] = {
                'name': name,
                'formula': formula,
                'description': f"自定义计算列: {formula}"
            }
            self.save_custom_columns()
            print(f"[DEBUG] 已保存计算列配置到文件")

            # 更新界面
            print(f"[DEBUG] 准备刷新界面...")
            try:
                if hasattr(self, 'refresh_column_list'):
                    print(f"[DEBUG] 调用 refresh_column_list")
                    self.refresh_column_list()
                    print(f"[DEBUG] refresh_column_list 调用完成")
                    print(f"[DEBUG] 当前 column_vars 数量: {len(self.column_vars)}")
                    print(f"[DEBUG] 新列是否在 column_vars 中: {full_name in self.column_vars}")
                else:
                    print("[ERROR] refresh_column_list 方法不存在！")
                    print("自定义计算列已创建成功，但数据列列表未刷新")
                    print("请重新加载文件以查看新列")
            except Exception as e:
                print(f"[ERROR] 刷新数据列列表时出错: {e}")
                import traceback
                traceback.print_exc()
                print("自定义计算列已创建成功，但界面更新失败")

            # 询问是否添加到当前分组
            if self.current_group:
                response = messagebox.askyesno("提示", 
                    f'计算列 "{full_name}" 已创建成功!\n是否添加到当前分组?')
                if response:
                    if self.current_group not in self.groups:
                        self.groups[self.current_group] = {}
                    self.groups[self.current_group][full_name] = True
                    self.save_groups()
                    
                    # 【修复】实际勾选这个列的复选框
                    if full_name in self.column_vars:
                        self.column_vars[full_name].set(1)
                    
                    self.update_group_display(self.current_group)
            
            messagebox.showinfo("成功", f'自定义计算列 "{full_name}" 已创建成功!')
            editor_window.destroy()
            
        except Exception as e:
            messagebox.showerror("错误", f"创建计算列失败: {e}")


    def open_custom_column_editor(self):
        """打开自定义列编辑器窗口"""
        if self.df is None:
            messagebox.showwarning("警告", "请先加载Excel文件!")
            return
        
        # 获取所有可用的数据列（排除时间列）
        available_columns = [col for col in self.df.columns if col != self.time_column]
        
        # 创建编辑器窗口
        editor_window = tk.Toplevel(self.root)
        editor_window.title("自定义计算列")
        editor_window.geometry("600x550")
        editor_window.resizable(False, False)
        
        # 【修复】设置窗口层级，防止被主窗口遮挡
        editor_window.transient(self.root)  # 设置为主窗口的临时窗口
        editor_window.grab_set()  # 设置为模态窗口，强制用户处理完此窗口
        editor_window.focus_force()  # 强制获取焦点
        
        # 列名称输入
        name_frame = ttk.Frame(editor_window)
        name_frame.pack(fill=tk.X, padx=10, pady=10)
        ttk.Label(name_frame, text="列名称:").pack(side=tk.LEFT)
        name_var = tk.StringVar()
        name_entry = ttk.Entry(name_frame, textvariable=name_var, width=40)
        name_entry.pack(side=tk.LEFT, padx=5)
        
        # 计算公式输入
        formula_frame = ttk.Frame(editor_window)
        formula_frame.pack(fill=tk.X, padx=10, pady=5)
        ttk.Label(formula_frame, text="计算公式:").pack(anchor=tk.W)
        formula_var = tk.StringVar()
        formula_entry = ttk.Entry(formula_frame, textvariable=formula_var, width=50)
        formula_entry.pack(fill=tk.X, pady=5)
        
        # 公式预览
        preview_frame = ttk.Frame(editor_window)
        preview_frame.pack(fill=tk.X, padx=10, pady=5)
        ttk.Label(preview_frame, text="公式预览:", font=('Arial', 9, 'bold')).pack(anchor=tk.W)
        preview_label = tk.Label(preview_frame, text="", bg="white", relief="solid", width=60, height=3, justify=tk.LEFT, anchor=tk.W)
        preview_label.pack(fill=tk.X, pady=5)
        
        # 更新预览的函数
        def update_preview(*args):
            preview_label.config(text=formula_var.get())
        formula_var.trace('w', update_preview)
        
        # 数据列选择区域
        columns_frame = ttk.LabelFrame(editor_window, text="可用数据列", padding=10)
        columns_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # 创建搜索框和下拉框的组合
        search_col_frame = ttk.Frame(columns_frame)
        search_col_frame.pack(side=tk.LEFT, padx=5)
        
        # 搜索输入框
        ttk.Label(search_col_frame, text="搜索:").pack(anchor=tk.W)
        search_col_var = tk.StringVar()
        search_col_entry = ttk.Entry(search_col_frame, textvariable=search_col_var, width=30)
        search_col_entry.pack(fill=tk.X, pady=2)
        
        # 数据列下拉框（改为可搜索模式）
        col_var = tk.StringVar()
        col_combobox = ttk.Combobox(search_col_frame, textvariable=col_var, width=30)
        col_combobox['values'] = available_columns
        if available_columns:
            col_combobox.current(0)
        col_combobox.pack(fill=tk.X, pady=2)
        
        # 实时过滤功能
        def filter_columns(*args):
            search_text = search_col_var.get().lower()
            if search_text:
                filtered = [col for col in available_columns if search_text in col.lower()]
            else:
                filtered = available_columns
            col_combobox['values'] = filtered
            # 不再自动选中第一个，让用户自己选择
            # 如果当前值不在过滤结果中，清空显示
            current_value = col_var.get()
            if current_value and current_value not in filtered:
                col_var.set('')  # 清空，但不自动选中第一个
        
        search_col_var.trace('w', filter_columns)
        
        # 支持回车键直接添加第一个匹配项
        def on_enter(event):
            search_text = search_col_var.get().lower()
            if search_text:
                filtered = [col for col in available_columns if search_text in col.lower()]
                if filtered:
                    self.insert_to_formula(formula_entry, filtered[0])
                    search_col_var.set('')  # 清空搜索框
                    col_var.set('')
        
        search_col_entry.bind('<Return>', on_enter)
        
        # 支持下拉键打开下拉列表
        def on_down(event):
            col_combobox.event_generate('<Down>')
        search_col_entry.bind('<Down>', on_down)
        
        # 运算符按钮
        operators_frame = ttk.Frame(columns_frame)
        operators_frame.pack(side=tk.LEFT, padx=10)
        
        operators = ['+', '-', '*', '/', '(', ')']
        for op in operators:
            btn = ttk.Button(operators_frame, text=op, width=4,
                           command=lambda o=op: self.insert_to_formula(formula_entry, o))
            btn.pack(side=tk.LEFT, padx=2)
        
        # 添加列名按钮
        ttk.Button(columns_frame, text="添加列名", 
                  command=lambda: self.insert_to_formula(formula_entry, col_var.get()),
                  width=10).pack(side=tk.LEFT, padx=5)
        
        # 常量输入区域
        constant_frame = ttk.LabelFrame(editor_window, text="常量输入", padding=10)
        constant_frame.pack(fill=tk.X, padx=10, pady=5)
        
        const_var = tk.StringVar()
        const_entry = ttk.Entry(constant_frame, textvariable=const_var, width=20)
        const_entry.pack(side=tk.LEFT, padx=5)
        ttk.Button(constant_frame, text="添加常量",
                  command=lambda: self.insert_to_formula(formula_entry, const_var.get()),
                  width=10).pack(side=tk.LEFT, padx=5)
        
        # 按钮区域
        button_frame = ttk.Frame(editor_window)
        button_frame.pack(fill=tk.X, padx=10, pady=10)
        
        ttk.Button(button_frame, text="验证公式",
                  command=lambda: self.validate_formula(name_var.get(), formula_var.get(), available_columns),
                  width=15).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(button_frame, text="确定",
                  command=lambda: self.create_custom_column(name_var.get(), formula_var.get(), available_columns, editor_window),
                  width=10).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(button_frame, text="取消",
                  command=editor_window.destroy,
                  width=10).pack(side=tk.LEFT, padx=5)
    



    def update_group_display(self, group_name):
        """更新分组列显示区域 - 只显示已勾选的列"""
        self.clear_group_display()
        
        if group_name not in self.groups:
            return
        
        group_data = self.groups[group_name]
        if not group_data:
            ttk.Label(
                self.group_scrollable_frame,
                text="该分组暂无列数据",
                foreground="gray"
            ).pack(anchor=tk.W, padx=5, pady=5)
            return
        
        # 只显示已勾选的列（checked == 1）
        has_checked = False
        for col, checked in group_data.items():
            if checked == 1:  # 只显示勾选的列
                has_checked = True
                label = ttk.Label(
                    self.group_scrollable_frame,
                    text=f"[√] {col}",
                    foreground="green",
                    font=("Arial", 9)
                )
                label.pack(anchor=tk.W, padx=5, pady=2)
                self.group_labels[col] = label
        
        # 如果没有勾选任何列，显示提示
        if not has_checked:
            ttk.Label(
                self.group_scrollable_frame,
                text="该分组暂无勾选的列",
                foreground="gray"
            ).pack(anchor=tk.W, padx=5, pady=5)
    
    def _create_menu_bar(self):
        """创建菜单栏"""
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        
        # 文件菜单
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="文件", menu=file_menu)
        file_menu.add_command(label="选择Excel文件", command=self.select_file)
        file_menu.add_separator()
        file_menu.add_command(label="退出", command=self.root.quit)
        
        # 配置管理菜单
        config_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="配置管理", menu=config_menu)
        config_menu.add_command(label="导出配置...", command=self._export_config)
        config_menu.add_command(label="导入配置...", command=self._import_config)
        config_menu.add_separator()
        config_menu.add_command(label="重置所有配置", command=self._reset_all_configs)
        
        # 帮助菜单
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="帮助", menu=help_menu)
        help_menu.add_command(label="使用说明", command=self._show_help)
        help_menu.add_command(label="关于", command=self._show_about)
    
    def _show_help(self):
        """显示使用说明窗口"""
        help_window = tk.Toplevel(self.root)
        help_window.title("使用说明 - Excel数据可视化工具")
        help_window.geometry("900x700")
        help_window.transient(self.root)
        
        # 创建主框架
        main_frame = ttk.Frame(help_window)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 创建左侧导航
        nav_frame = ttk.Frame(main_frame, width=180)
        nav_frame.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 10))
        nav_frame.pack_propagate(False)
        
        # 创建右侧内容区域
        content_frame = ttk.Frame(main_frame)
        content_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # 创建Text组件显示帮助内容
        help_text = tk.Text(content_frame, wrap=tk.WORD, font=("Microsoft YaHei", 10), 
                           padx=10, pady=10)
        help_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # 添加滚动条
        scrollbar = ttk.Scrollbar(content_frame, orient=tk.VERTICAL, command=help_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        help_text.config(yscrollcommand=scrollbar.set)
        
        # 定义帮助内容
        help_content = """
═══════════════════════════════════════════════════════════════════
                    Excel数据可视化工具 - 使用说明
                         版本：v2.9.2
═══════════════════════════════════════════════════════════════════

【目录】
  一、软件概述
  二、单文件分析模式
  三、批量分析模式
  四、常用分组功能
  五、数据诊断功能
  六、图表交互功能
  七、快捷键与技巧
  八、常见问题解答

═══════════════════════════════════════════════════════════════════
一、软件概述
═══════════════════════════════════════════════════════════════════

【功能简介】
本工具用于快速分析Excel数据文件，生成可视化折线图，并支持批量检测数据异常。

【主要功能】
• 单文件分析：选择单个Excel文件，勾选数据列生成折线图
• 批量分析：批量扫描文件夹，自动检测异常数据
• 常用分组：保存常用的数据列组合，一键加载
• 数据诊断：检测波动剧烈、超出范围、同增同减等异常
• 图表交互：缩放、拖动、显示数据点详情

【界面布局】
┌─────────────────────────────────────────────────────────┐
│  [单文件分析]  [批量分析]                    ← 标签页    │
├───────────────┬─────────────────────────────────────────┤
│               │                                         │
│   左侧面板    │              右侧图表区域               │
│   (控制选项)  │              (折线图显示)               │
│               │                                         │
│               │         ┌─────────────┐                │
│               │         │  图例面板   │ ← 可交互图例   │
│               │         └─────────────┘                │
└───────────────┴─────────────────────────────────────────┘

═══════════════════════════════════════════════════════════════════
二、单文件分析模式
═══════════════════════════════════════════════════════════════════

【步骤1：选择文件】
1. 点击"选择Excel文件"按钮
2. 在弹出的文件对话框中选择.xlsx或.xls文件
3. 文件路径会显示在按钮上方

【步骤2：选择数据列】
1. 文件加载后，左侧"数据列选择"区域会显示所有列名
2. 勾选需要显示的数据列
3. 可使用搜索框快速筛选列名
4. 点击"全选"/"全不选"批量操作

【步骤3：生成图表】
1. 点击"生成折线图"按钮
2. 图表显示在右侧区域
3. 横轴为第一列（通常为时间列）

【折叠面板功能】
• 点击各区域标题栏可展开/折叠
• ▼ 表示展开状态
• ▶ 表示折叠状态
• 折叠后可节省界面空间

【图表设置】
在图表区域上方可设置：
• 图表标题：自定义图表名称
• X轴标签：横轴说明文字
• Y轴标签：纵轴说明文字

═══════════════════════════════════════════════════════════════════
三、批量分析模式
═══════════════════════════════════════════════════════════════════

【步骤1：选择文件夹】
1. 切换到"批量分析"标签页
2. 点击"选择文件夹"按钮
3. 选择包含Excel文件的文件夹

【步骤2：扫描文件】
1. 点击"扫描文件"按钮
2. 程序会自动扫描文件夹中的Excel文件
3. 识别文件类型（电池/PCS）
4. 显示扫描进度和统计信息

【步骤3：选择数据类型】
• BATTERY：电池相关数据
• PCS：PCS相关数据

【步骤4：设置异常条件】
点击"+ 添加条件"可添加检测条件：

【条件类型说明】

┌──────────────┬─────────────────────────────────────┐
│   条件类型   │              参数说明               │
├──────────────┼─────────────────────────────────────┤
│ 波动剧烈检测 │  列选择 + 阈值百分比                │
│              │  检测相邻时刻变化率超过阈值的点     │
├──────────────┼─────────────────────────────────────┤
│ 同增同减检测 │  多列选择（2-5列）                 │
│              │  检测多列数据是否同时增减           │
├──────────────┼─────────────────────────────────────┤
│ 数据范围异常 │  列选择 + 最小值 + 最大值           │
│              │  检测数据是否超出正常范围           │
├──────────────┼─────────────────────────────────────┤
│ 多条件组合   │  列选择 + 多个条件 + 逻辑关系       │
│              │  支持AND/OR组合多个判断条件         │
│              │  例：值!=1 AND 值!=2 AND 值!=3     │
├──────────────┼─────────────────────────────────────┤
│ 自定义       │  自定义计算公式                     │
│              │  支持算术运算符和逻辑运算符         │
│              │  算术：+ - * / ( )                  │
│              │  逻辑：AND OR NOT                   │
└──────────────┴─────────────────────────────────────┘

【多条件组合使用示例】
• 检测值不能是1、2、3：选择列 → 添加条件"!=1" → 添加条件"!=2" → 添加条件"!=3" → 逻辑选择AND
• 检测值大于100或小于-100：选择列 → 添加条件">100" → 添加条件"<-100" → 逻辑选择OR

【自定义公式使用示例】
• 简单比较：[列A] > 100
• 算术运算：[列A] + [列B] * 2 > 500
• 逻辑组合：([列A] > 100) AND ([列B] < 50)
• 逻辑或：([列A] > 100) OR ([列A] < -100)
• 逻辑非：NOT ([列A] = 0)

【条件配置保存】
1. 设置好条件后，点击"保存配置"
2. 输入配置名称保存
3. 下次可通过下拉框一键加载

【步骤5：开始分析】
1. 勾选要启用的条件
2. 点击"开始批量分析"
3. 查看分析结果和异常文件列表

【结果导出】
• 点击"导出结果"可保存分析报告
• 支持Excel格式导出

═══════════════════════════════════════════════════════════════════
四、常用分组功能
═══════════════════════════════════════════════════════════════════

【功能说明】
常用分组用于保存常用的数据列组合，避免每次手动勾选。

【新建分组】
1. 点击"新建"按钮
2. 输入分组名称
3. 在弹出的窗口中勾选数据列
4. 点击"确定"保存

【编辑分组】
1. 从下拉框选择要编辑的分组
2. 点击"编辑"按钮
3. 修改勾选的数据列
4. 点击"确定"保存修改

【删除分组】
1. 从下拉框选择要删除的分组
2. 点击"删除"按钮
3. 确认删除

【重命名分组】
1. 从下拉框选择要重命名的分组
2. 点击"重命名"按钮
3. 输入新名称

【自定义列】
• 点击"自定义列"可添加计算列
• 支持自定义公式计算
• 如：A列 + B列，生成新的计算列

═══════════════════════════════════════════════════════════════════
五、数据诊断功能
═══════════════════════════════════════════════════════════════════

【功能说明】
数据诊断用于分析当前文件的数据质量，发现潜在问题。

【诊断策略】

┌──────────────┬─────────────────────────────────────┐
│   策略名称   │              检测内容               │
├──────────────┼─────────────────────────────────────┤
│ 波动剧烈检测 │ 检测数据波动超过阈值的异常点        │
├──────────────┼─────────────────────────────────────┤
│ 同增同减检测 │ 检测多列数据同步变化异常            │
├──────────────┼─────────────────────────────────────┤
│ 数据范围异常 │ 检测数据是否超出正常范围            │
└──────────────┴─────────────────────────────────────┘

【使用步骤】
1. 选择诊断策略
2. 设置相关参数（阈值、检测列等）
3. 点击"运行诊断"
4. 点击"查看报告"查看详细结果

═══════════════════════════════════════════════════════════════════
六、图表交互功能
═══════════════════════════════════════════════════════════════════

【滚轮缩放】
• 在图表区域滚动鼠标滚轮可缩放图表
• 向上滚动：放大
• 向下滚动：缩小

【拖动平移】
• 按住鼠标左键拖动可平移图表
• 查看不同时间段的数据

【数据点详情】
• 点击图表上的数据点
• 显示该点的详细信息：
  - 时间
  - 列名
  - 数值

【工具栏功能】
图表上方工具栏提供：
┌────────┬──────────────────────────────────┐
│  按钮  │             功能                 │
├────────┼──────────────────────────────────┤
│  🏠    │ 重置视图，恢复初始状态           │
│  ← →   │ 前进/后退视图历史                │
│  📷    │ 保存图片                         │
│  ⚙     │ 图表设置                         │
└────────┴──────────────────────────────────┘

【图例面板】
右侧图例面板显示所有数据列：
• 点击列名：切换该列显示/隐藏
• 点击👁按钮：切换显示/隐藏
• 点击✕按钮：从图表中删除该列
• 隐藏状态：文字和颜色标记变灰

═══════════════════════════════════════════════════════════════════
七、快捷键与技巧
═══════════════════════════════════════════════════════════════════

【常用快捷键】
┌──────────┬────────────────────────────────┐
│  快捷键  │            功能                │
├──────────┼────────────────────────────────┤
│  Ctrl+O  │ 打开文件                       │
│  Ctrl+S  │ 保存图片                       │
│  Ctrl+Q  │ 退出程序                       │
│  F1      │ 显示帮助                       │
│  F5      │ 刷新图表                       │
└──────────┴────────────────────────────────┘

【使用技巧】
1. 搜索列名时，输入数字可匹配括号内的编码
   例：输入"20570"可快速找到"直流PV总功率(20570)"

2. 折叠不常用的面板可节省界面空间

3. 使用常用分组功能快速加载预设的数据列组合

4. 批量分析前先保存条件配置，方便复用

5. 图表缩放后点击"🏠"按钮可恢复初始视图

═══════════════════════════════════════════════════════════════════
八、常见问题解答
═══════════════════════════════════════════════════════════════════

【Q1：为什么图表中文显示为方框？】
A：系统缺少中文字体。可安装"SimHei"或"Microsoft YaHei"字体。

【Q2：为什么扫描不到Excel文件？】
A：请确保：
   1. 文件夹中确实有.xlsx或.xls文件
   2. 文件未被其他程序占用
   3. 文件格式正确，非损坏文件

【Q3：为什么批量分析没有结果？】
A：请检查：
   1. 是否已勾选启用条件
   2. 条件参数设置是否合理
   3. 数据类型选择是否正确

【Q4：如何导出高清图表？】
A：点击工具栏中的"📷"按钮，选择保存路径和格式。

【Q5：数据量很大时程序卡顿怎么办？】
A：
   1. 减少勾选的数据列数量
   2. 使用搜索功能筛选关键列
   3. 分批次进行分析

【Q6：配置保存在哪里？】
A：配置文件保存在程序同目录下：
   • common_groups.json - 常用分组配置
   • custom_columns.json - 自定义列配置
   • batch_conditions_config.json - 批量分析条件配置

【Q7：如何在不同电脑间共享配置？】
A：使用菜单栏「配置管理」功能：
   1. 导出配置：将所有配置导出为单个 JSON 文件
   2. 导入配置：从导出文件导入配置（支持合并或覆盖）
   3. 可选择性地导入特定配置项

═══════════════════════════════════════════════════════════════════
                        技术支持与反馈
═══════════════════════════════════════════════════════════════════

如有问题或建议，请联系开发团队。

版本历史：
• v2.9.2 - 用户体验优化：重新配置时记忆上次选择状态
• v2.9.1 - 故障数据解析模块重大修复：重新配置、类型名称、路径查找
• v2.9.0 - 故障配置对话框、寄存器选择优化
• v2.8.0 - 显示所有bit位(含未定义)、四宫格布局、饼图优化
• v2.7.9 - 四宫格布局、饼图优化、点表格式支持、滚动条
• v2.7.2 - 修复故障分析方法，使用正确的parse_fault_value函数
• v2.5.1 - 故障数据分析完善，故障码解析模块
• v2.5.0 - 分析模式切换（普通数据/故障数据）
• v2.4.0 - 波动剧烈检测逻辑修正
• v2.3.0 - 批量分析导出优化：按文件名分sheet
• v2.2.0 - 同增同减异常整体一致性检查
• v2.1.0 - 新增帮助系统，优化批量分析条件配置
• v2.0.0 - 新增可折叠面板功能
• v1.0.0 - 初始版本，可交互图例面板

═══════════════════════════════════════════════════════════════════
"""
        
        # 插入帮助内容
        help_text.insert(tk.END, help_content)
        help_text.config(state=tk.DISABLED)  # 设为只读
        
        # 创建导航按钮
        ttk.Label(nav_frame, text="快速导航", font=("Microsoft YaHei", 11, "bold")).pack(pady=10)
        
        sections = [
            ("一、软件概述", "一、软件概述"),
            ("二、单文件分析", "二、单文件分析模式"),
            ("三、批量分析", "三、批量分析模式"),
            ("四、常用分组", "四、常用分组功能"),
            ("五、数据诊断", "五、数据诊断功能"),
            ("六、图表交互", "六、图表交互功能"),
            ("七、快捷键", "七、快捷键与技巧"),
            ("八、常见问题", "八、常见问题解答"),
        ]
        
        for title, search_text in sections:
            btn = ttk.Button(nav_frame, text=title, width=16,
                           command=lambda s=search_text, t=help_text: self._scroll_to_section(t, s))
            btn.pack(pady=3, padx=5)
        
        # 添加关闭按钮
        ttk.Separator(nav_frame, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=20)
        ttk.Button(nav_frame, text="关闭", command=help_window.destroy, width=16).pack(pady=5)
    
    def _scroll_to_section(self, text_widget, section_title):
        """滚动到指定章节"""
        content = text_widget.get("1.0", tk.END)
        
        # 目录中的章节格式：  一、软件概述 (有缩进)
        # 正文中的章节格式：一、软件概述 (无缩进)
        # 所以搜索无缩进的格式即可跳过目录
        
        # 搜索正文中的章节标题（无缩进，前后有换行）
        search_text = f"\n{section_title}\n"
        
        idx = content.find(search_text)
        if idx != -1:
            # 找到了，计算行号并滚动到章节标题行
            line_num = content[:idx].count('\n') + 2  # +2 跳过换行
            text_widget.see(f"{line_num}.0")
            return
        
        # 备用方案：直接搜索章节标题
        idx = content.find(section_title)
        if idx != -1:
            line_num = content[:idx].count('\n') + 1
            text_widget.see(f"{line_num}.0")
    
    def _show_about(self):
        """显示关于窗口"""
        about_window = tk.Toplevel(self.root)
        about_window.title("关于")
        about_window.geometry("400x300")
        about_window.transient(self.root)
        about_window.resizable(False, False)
        
        # 居中显示
        about_window.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width() - 400) // 2
        y = self.root.winfo_y() + (self.root.winfo_height() - 300) // 2
        about_window.geometry(f"+{x}+{y}")
        
        main_frame = ttk.Frame(about_window, padding=20)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 软件名称
        ttk.Label(main_frame, text="Excel数据可视化工具", 
                 font=("Microsoft YaHei", 16, "bold")).pack(pady=10)
        
        # 版本信息
        ttk.Label(main_frame, text="版本：v2.9.2", 
                 font=("Microsoft YaHei", 10)).pack()
        
        ttk.Separator(main_frame, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=15)
        
        # 功能描述
        desc = """功能特点：
• 单文件/批量数据分析
• 可交互图例面板
• 可折叠功能区
• 数据诊断与异常检测
• 常用分组管理
• 图表缩放与导出"""
        
        ttk.Label(main_frame, text=desc, font=("Microsoft YaHei", 9), 
                 justify=tk.LEFT).pack(pady=10)
        
        # 版权信息
        ttk.Label(main_frame, text="© 2025 Excel Data Visualizer", 
                 font=("Microsoft YaHei", 8), foreground="gray").pack(side=tk.BOTTOM)
    
    def _export_config(self):
        """导出配置到外部文件"""
        from tkinter import filedialog
        import datetime
        
        # 选择保存路径
        file_path = filedialog.asksaveasfilename(
            title="导出配置",
            defaultextension=".json",
            filetypes=[("JSON配置文件", "*.json"), ("所有文件", "*.*")],
            initialfile=f"excel_visualizer_config_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
        )
        
        if not file_path:
            return
        
        # 合并所有配置
        config = {
            "version": "2.1.1",
            "export_time": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "common_groups": {},
            "custom_columns": {},
            "batch_conditions": {}
        }
        
        # 读取常用分组配置
        if os.path.exists(self.groups_file):
            try:
                with open(self.groups_file, 'r', encoding='utf-8') as f:
                    config["common_groups"] = json.load(f)
            except Exception as e:
                print(f"读取常用分组配置失败: {e}")
        
        # 读取自定义列配置
        if os.path.exists(self.custom_columns_file):
            try:
                with open(self.custom_columns_file, 'r', encoding='utf-8') as f:
                    config["custom_columns"] = json.load(f)
            except Exception as e:
                print(f"读取自定义列配置失败: {e}")
        
        # 读取批量分析条件配置
        if os.path.exists(self.batch_conditions_config_file):
            try:
                with open(self.batch_conditions_config_file, 'r', encoding='utf-8') as f:
                    config["batch_conditions"] = json.load(f)
            except Exception as e:
                print(f"读取批量分析条件配置失败: {e}")
        
        # 保存到指定路径
        try:
            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
            messagebox.showinfo("导出成功", f"配置已导出到：\n{file_path}\n\n包含：\n• 常用分组配置\n• 自定义列配置\n• 批量分析条件配置")
        except Exception as e:
            messagebox.showerror("导出失败", f"保存配置文件失败：\n{e}")
    
    def _import_config(self):
        """从外部文件导入配置"""
        from tkinter import filedialog
        
        # 选择配置文件
        file_path = filedialog.askopenfilename(
            title="导入配置",
            filetypes=[("JSON配置文件", "*.json"), ("所有文件", "*.*")]
        )
        
        if not file_path:
            return
        
        # 读取配置文件
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                config = json.load(f)
        except Exception as e:
            messagebox.showerror("导入失败", f"读取配置文件失败：\n{e}")
            return
        
        # 验证配置格式
        if not isinstance(config, dict):
            messagebox.showerror("导入失败", "配置文件格式错误！")
            return
        
        # 创建导入选项对话框
        import_window = tk.Toplevel(self.root)
        import_window.title("选择要导入的配置")
        import_window.geometry("480x380")
        import_window.transient(self.root)
        import_window.resizable(False, False)
        
        # 居中显示
        import_window.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width() - 480) // 2
        y = self.root.winfo_y() + (self.root.winfo_height() - 380) // 2
        import_window.geometry(f"+{x}+{y}")
        
        main_frame = ttk.Frame(import_window, padding=20)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 显示配置文件信息
        if "export_time" in config:
            ttk.Label(main_frame, text=f"配置导出时间: {config.get('export_time', '未知')}",
                     font=("Microsoft YaHei", 9)).pack(anchor=tk.W)
        if "version" in config:
            ttk.Label(main_frame, text=f"配置版本: {config.get('version', '未知')}",
                     font=("Microsoft YaHei", 9)).pack(anchor=tk.W, pady=(0, 10))
        
        ttk.Separator(main_frame, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=10)
        
        # 选择要导入的配置项
        ttk.Label(main_frame, text="请选择要导入的配置项：",
                 font=("Microsoft YaHei", 10, "bold")).pack(anchor=tk.W, pady=(0, 10))
        
        import_vars = {}
        
        # 常用分组
        if "common_groups" in config and config["common_groups"]:
            var = tk.BooleanVar(value=True)
            import_vars["common_groups"] = var
            cb = ttk.Checkbutton(main_frame, text=f"常用分组配置 ({len(config['common_groups'])} 个分组)",
                                variable=var)
            cb.pack(anchor=tk.W, pady=2)
        
        # 自定义列
        if "custom_columns" in config and config["custom_columns"]:
            var = tk.BooleanVar(value=True)
            import_vars["custom_columns"] = var
            cb = ttk.Checkbutton(main_frame, text=f"自定义列配置 ({len(config['custom_columns'])} 个列)",
                                variable=var)
            cb.pack(anchor=tk.W, pady=2)
        
        # 批量分析条件
        if "batch_conditions" in config and config["batch_conditions"]:
            var = tk.BooleanVar(value=True)
            import_vars["batch_conditions"] = var
            cb = ttk.Checkbutton(main_frame, text=f"批量分析条件配置 ({len(config['batch_conditions'])} 个配置)",
                                variable=var)
            cb.pack(anchor=tk.W, pady=2)
        
        if not import_vars:
            ttk.Label(main_frame, text="配置文件中没有可导入的配置项！",
                     font=("Microsoft YaHei", 9), foreground="red").pack(pady=20)
            ttk.Button(main_frame, text="关闭", command=import_window.destroy).pack(pady=10)
            return
        
        ttk.Separator(main_frame, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=10)
        
        # 导入模式选择
        mode_frame = ttk.Frame(main_frame)
        mode_frame.pack(fill=tk.X, pady=5)
        ttk.Label(mode_frame, text="导入模式：", font=("Microsoft YaHei", 9)).pack(side=tk.LEFT)
        mode_var = tk.StringVar(value="merge")
        ttk.Radiobutton(mode_frame, text="合并（保留现有配置）", variable=mode_var, 
                       value="merge").pack(side=tk.LEFT, padx=10)
        ttk.Radiobutton(mode_frame, text="覆盖（替换现有配置）", variable=mode_var,
                       value="overwrite").pack(side=tk.LEFT)
        
        def do_import():
            imported = []
            
            # 导入常用分组
            if "common_groups" in import_vars and import_vars["common_groups"].get():
                try:
                    imported_groups = config["common_groups"]
                    # 兼容两种格式：{groups: {...}} 或直接 {...}
                    if "groups" in imported_groups:
                        groups_data = imported_groups["groups"]
                        group_columns_data = imported_groups.get("group_columns", {})
                    else:
                        groups_data = imported_groups
                        group_columns_data = {}
                    
                    if mode_var.get() == "overwrite":
                        self.groups = groups_data.copy()
                        self.group_columns = group_columns_data.copy()
                    else:
                        self.groups.update(groups_data)
                        self.group_columns.update(group_columns_data)
                    
                    # 保存时保持原有格式
                    save_data = {
                        'groups': self.groups,
                        'group_columns': self.group_columns
                    }
                    with open(self.groups_file, 'w', encoding='utf-8') as f:
                        json.dump(save_data, f, ensure_ascii=False, indent=2)
                    imported.append("常用分组配置")
                    self.update_group_combobox()
                except Exception as e:
                    print(f"导入常用分组配置失败: {e}")
            
            # 导入自定义列
            if "custom_columns" in import_vars and import_vars["custom_columns"].get():
                try:
                    if mode_var.get() == "overwrite":
                        self.custom_columns = config["custom_columns"].copy()
                    else:
                        self.custom_columns.update(config["custom_columns"])
                    
                    with open(self.custom_columns_file, 'w', encoding='utf-8') as f:
                        json.dump(self.custom_columns, f, ensure_ascii=False, indent=2)
                    imported.append("自定义列配置")
                    # 重新加载文件以应用自定义列
                    if self.df is not None:
                        self._apply_custom_columns()
                except Exception as e:
                    print(f"导入自定义列配置失败: {e}")
            
            # 导入批量分析条件
            if "batch_conditions" in import_vars and import_vars["batch_conditions"].get():
                try:
                    batch_config = config["batch_conditions"]
                    if mode_var.get() == "overwrite":
                        with open(self.batch_conditions_config_file, 'w', encoding='utf-8') as f:
                            json.dump(batch_config, f, ensure_ascii=False, indent=2)
                    else:
                        existing = {}
                        if os.path.exists(self.batch_conditions_config_file):
                            try:
                                with open(self.batch_conditions_config_file, 'r', encoding='utf-8') as f:
                                    existing = json.load(f)
                            except:
                                pass
                        existing.update(batch_config)
                        with open(self.batch_conditions_config_file, 'w', encoding='utf-8') as f:
                            json.dump(existing, f, ensure_ascii=False, indent=2)
                    imported.append("批量分析条件配置")
                    self._load_batch_config_list()
                except Exception as e:
                    print(f"导入批量分析条件配置失败: {e}")
            
            import_window.destroy()
            
            if imported:
                messagebox.showinfo("导入成功", f"已成功导入以下配置：\n• " + "\n• ".join(imported))
            else:
                messagebox.showwarning("导入结果", "没有配置被导入。")
        
        # 按钮
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=15)
        ttk.Button(btn_frame, text="导入", command=do_import).pack(side=tk.LEFT, expand=True, padx=5)
        ttk.Button(btn_frame, text="取消", command=import_window.destroy).pack(side=tk.LEFT, expand=True, padx=5)
    
    def _reset_all_configs(self):
        """重置所有配置"""
        if not messagebox.askyesno("确认重置", 
                "确定要重置所有配置吗？\n\n这将删除：\n• 所有常用分组\n• 所有自定义列\n• 所有批量分析条件\n\n此操作不可撤销！"):
            return
        
        try:
            # 清空常用分组
            self.groups = {}
            self.group_columns = {}
            if os.path.exists(self.groups_file):
                with open(self.groups_file, 'w', encoding='utf-8') as f:
                    json.dump({'groups': {}, 'group_columns': {}}, f)
            self.update_group_combobox()
            
            # 清空自定义列
            self.custom_columns = {}
            if os.path.exists(self.custom_columns_file):
                with open(self.custom_columns_file, 'w', encoding='utf-8') as f:
                    json.dump({}, f)
            
            # 清空批量分析条件
            if os.path.exists(self.batch_conditions_config_file):
                with open(self.batch_conditions_config_file, 'w', encoding='utf-8') as f:
                    json.dump({}, f)
            self._load_batch_config_list()
            
            messagebox.showinfo("重置成功", "所有配置已重置！")
        except Exception as e:
            messagebox.showerror("重置失败", f"重置配置时发生错误：\n{e}")
    
    def create_widgets(self):
        """创建UI组件"""
        # 数据模式变量（运行数据 / 故障数据）
        self.data_mode = tk.StringVar(value="runtime")
        
        # ========== 顶层Notebook：运行数据解析 / 故障数据解析 ==========
        self.top_notebook = ttk.Notebook(self.root)
        self.top_notebook.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)
        
        # 绑定顶层Tab切换事件
        self.top_notebook.bind('<<NotebookTabChanged>>', self._on_top_tab_changed)
        
        # ========== Tab 1: 运行数据解析 ==========
        self.runtime_tab = ttk.Frame(self.top_notebook)
        self.top_notebook.add(self.runtime_tab, text="📊 运行数据解析")
        
        # 运行数据解析的内部Notebook
        self.runtime_notebook = ttk.Notebook(self.runtime_tab)
        self.runtime_notebook.pack(fill=tk.BOTH, expand=True)
        self.runtime_notebook.bind('<<NotebookTabChanged>>', self._on_tab_changed)
        
        # 运行数据 - 单文件分析
        self.runtime_single_tab = ttk.Frame(self.runtime_notebook)
        self.runtime_notebook.add(self.runtime_single_tab, text="单文件分析")
        
        # 运行数据 - 批量分析
        self.runtime_batch_tab = ttk.Frame(self.runtime_notebook)
        self.runtime_notebook.add(self.runtime_batch_tab, text="批量分析")
        
        # ========== Tab 2: 故障数据解析 ==========
        self.fault_tab = ttk.Frame(self.top_notebook)
        self.top_notebook.add(self.fault_tab, text="🔧 故障数据解析")
        
        # 故障数据解析的内部Notebook
        self.fault_notebook = ttk.Notebook(self.fault_tab)
        self.fault_notebook.pack(fill=tk.BOTH, expand=True)
        
        # 故障数据 - 单文件分析
        self.fault_single_tab = ttk.Frame(self.fault_notebook)
        self.fault_notebook.add(self.fault_single_tab, text="单文件分析")
        
        # 故障数据 - 批量分析
        self.fault_batch_tab = ttk.Frame(self.fault_notebook)
        self.fault_notebook.add(self.fault_batch_tab, text="批量分析")
        
        # 创建单文件分析界面（运行数据）
        self._create_single_file_widgets(self.runtime_single_tab, "runtime")
        
        # 创建批量分析界面（运行数据）
        self._create_batch_analysis_widgets(self.runtime_batch_tab, "runtime")
        
        # 创建单文件分析界面（故障数据）
        self._create_fault_single_file_widgets(self.fault_single_tab)
        
        # 创建批量分析界面（故障数据）- 暂时显示开发中提示
        self._create_fault_batch_placeholder(self.fault_batch_tab)
    
    def _show_fault_config_dialog(self):
        """显示故障配置对话框（选择文件后弹出）
        
        基于检测到的寄存器列，让用户选择要分析的列
        """
        print("[DEBUG] _show_fault_config_dialog 被调用")
        
        # 检查是否有检测结果
        if not hasattr(self, 'detected_registers') or not self.detected_registers:
            print("[DEBUG] 没有检测结果，显示提示")
            messagebox.showinfo("提示", "请先选择包含故障数据的Excel文件！")
            return
        
        # 确保 register_types 存在
        if not hasattr(self, 'register_types'):
            self.register_types = {}
        
        print(f"[DEBUG] 检测到 {len(self.detected_registers)} 个寄存器列")
        print(f"[DEBUG] detected_registers: {self.detected_registers[:3]}...")
        
        # 创建对话框窗口
        dialog = tk.Toplevel(self.root)
        dialog.title("故障分析配置")
        dialog.geometry("600x650")
        dialog.transient(self.root)  # 设置为父窗口的临时窗口
        dialog.grab_set()  # 模态对话框
        
        # 居中显示
        dialog.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width() - 600) // 2
        y = self.root.winfo_y() + (self.root.winfo_height() - 650) // 2
        dialog.geometry(f"+{x}+{y}")
        
        # 主框架
        main_frame = ttk.Frame(dialog, padding=15)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 标题
        ttk.Label(main_frame, text="🔧 选择要分析的故障寄存器", 
                  font=('Microsoft YaHei', 14, 'bold')).pack(anchor=tk.W, pady=(0, 10))
        
        # 提示信息
        ttk.Label(main_frame, 
                  text=f"已检测到 {len(self.detected_registers)} 个故障寄存器列，请勾选要分析的列：", 
                  font=('Microsoft YaHei', 10), foreground='blue').pack(anchor=tk.W, pady=(0, 10))
        
        # ========== 故障码来源选择 ==========
        source_frame = ttk.LabelFrame(main_frame, text="故障码定义来源", padding=10)
        source_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.dialog_source_var = tk.StringVar(value=self.fault_source)
        
        source_options = [
            ("builtin", "使用内置故障码表（推荐）", "629 个故障码定义，覆盖 43 个寄存器"),
            ("external", "使用外部故障码文件", "自定义故障码定义"),
        ]
        
        for value, label, desc in source_options:
            frame = ttk.Frame(source_frame)
            frame.pack(fill=tk.X, pady=2)
            
            rb = ttk.Radiobutton(frame, text=label, variable=self.dialog_source_var, 
                                 value=value, command=self._on_source_changed)
            rb.pack(anchor=tk.W)
            
            ttk.Label(frame, text=f"    {desc}", foreground='gray', 
                      font=('Microsoft YaHei', 9)).pack(anchor=tk.W)
        
        # 外部文件选择区域
        self.external_file_frame = ttk.Frame(source_frame)
        self.external_file_label = ttk.Label(self.external_file_frame, text="未选择文件", foreground='gray')
        self.external_file_label.pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Button(self.external_file_frame, text="选择文件...", 
                   command=self._select_external_fault_file).pack(side=tk.RIGHT, padx=(5, 0))
        
        # ========== 寄存器选择区域 ==========
        register_frame = ttk.LabelFrame(main_frame, text="故障寄存器选择", padding=10)
        register_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        # 快捷操作按钮
        quick_frame = ttk.Frame(register_frame)
        quick_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Button(quick_frame, text="全选", width=8,
                   command=lambda: self._select_all_detected_registers(True)).pack(side=tk.LEFT, padx=2)
        ttk.Button(quick_frame, text="全不选", width=8,
                   command=lambda: self._select_all_detected_registers(False)).pack(side=tk.LEFT, padx=2)
        ttk.Button(quick_frame, text="只选用户展示", 
                   command=self._select_detected_user_registers).pack(side=tk.LEFT, padx=2)
        ttk.Button(quick_frame, text="只选主故障码", 
                   command=self._select_detected_main_registers).pack(side=tk.LEFT, padx=2)
        
        # 寄存器选择区域（带滚动）
        register_container = ttk.Frame(register_frame)
        register_container.pack(fill=tk.BOTH, expand=True)
        
        canvas = tk.Canvas(register_container, highlightthickness=0)
        scrollbar = ttk.Scrollbar(register_container, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw", width=530)
        canvas.configure(yscrollcommand=scrollbar.set)
        
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # 绑定鼠标滚轮
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        canvas.bind("<MouseWheel>", _on_mousewheel)
        scrollable_frame.bind("<MouseWheel>", _on_mousewheel)
        
        # 按类型分组显示检测到的寄存器
        self.dialog_register_vars = {}
        
        # 获取上次选择的列（用于恢复勾选状态）
        previous_selections = set(getattr(self, 'selected_fault_columns', []))
        
        user_regs = [(r, c, t) for r, c, t in self.detected_registers if t == '用户展示']
        main_regs = [(r, c, t) for r, c, t in self.detected_registers if t == '主故障码']
        
        # 用户展示故障码
        if user_regs:
            user_frame = ttk.LabelFrame(scrollable_frame, text=f"用户展示故障码 ({len(user_regs)}个)", padding=5)
            user_frame.pack(fill=tk.X, pady=5, padx=5)
            
            for reg, col, _ in user_regs:
                # 如果有上次选择记录，使用上次的选择状态；否则默认全选
                if previous_selections:
                    default_checked = col in previous_selections
                else:
                    default_checked = True
                var = tk.BooleanVar(value=default_checked)
                self.dialog_register_vars[col] = var
                cb = ttk.Checkbutton(user_frame, text=col, variable=var)
                cb.pack(anchor=tk.W, pady=1)
        
        # 主故障码
        if main_regs:
            main_frame_reg = ttk.LabelFrame(scrollable_frame, text=f"主故障码 ({len(main_regs)}个)", padding=5)
            main_frame_reg.pack(fill=tk.X, pady=5, padx=5)
            
            for reg, col, _ in main_regs:
                # 如果有上次选择记录，使用上次的选择状态；否则默认全选
                if previous_selections:
                    default_checked = col in previous_selections
                else:
                    default_checked = True
                var = tk.BooleanVar(value=default_checked)
                self.dialog_register_vars[col] = var
                cb = ttk.Checkbutton(main_frame_reg, text=col, variable=var)
                cb.pack(anchor=tk.W, pady=1)
        
        # ========== 底部按钮 ==========
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=(10, 0))
        
        ttk.Button(btn_frame, text="取消", command=dialog.destroy, width=10).pack(side=tk.RIGHT, padx=5)
        ttk.Button(btn_frame, text="确定并开始分析", 
                   command=lambda: self._apply_fault_config(dialog), 
                   width=15).pack(side=tk.RIGHT, padx=5)
        
        # 初始化显示状态
        self._on_source_changed()
    
    def _select_all_detected_registers(self, select=True):
        """全选/全不选检测到的寄存器"""
        for var in self.dialog_register_vars.values():
            var.set(select)
    
    def _select_detected_user_registers(self):
        """只选择检测到的用户展示寄存器"""
        self._select_all_detected_registers(False)
        for col, var in self.dialog_register_vars.items():
            if col in self.register_types and self.register_types[col] == '用户展示':
                var.set(True)
    
    def _select_detected_main_registers(self):
        """只选择检测到的主故障码寄存器"""
        self._select_all_detected_registers(False)
        for col, var in self.dialog_register_vars.items():
            if col in self.register_types and self.register_types[col] == '主故障码':
                var.set(True)
    
    def _on_source_changed(self):
        """故障码来源改变时的回调"""
        if self.dialog_source_var.get() == "external":
            self.external_file_frame.pack(fill=tk.X, pady=(10, 0))
        else:
            self.external_file_frame.pack_forget()
    
    def _select_external_fault_file(self):
        """选择外部故障码文件"""
        file_path = filedialog.askopenfilename(
            title="选择故障码定义文件",
            filetypes=[
                ("Excel文件", "*.xlsx *.xls"),
                ("所有文件", "*.*")
            ]
        )
        if file_path:
            self.fault_external_file = file_path
            self.external_file_label.config(text=os.path.basename(file_path), foreground='blue')
    
    def _apply_fault_config(self, dialog):
        """应用故障配置并关闭对话框"""
        try:
            # 保存故障码来源
            self.fault_source = self.dialog_source_var.get()
            
            # 保存选中的列（现在是列名）
            self.selected_fault_columns = []
            for col, var in self.dialog_register_vars.items():
                if var.get():
                    self.selected_fault_columns.append(col)
            
            # 更新故障寄存器选择区域的复选框
            self._update_fault_register_checkboxes()
            
            # 更新配置状态提示
            if hasattr(self, 'fault_config_hint'):
                source_text = "内置故障码表" if self.fault_source == "builtin" else f"外部文件: {os.path.basename(self.fault_external_file) if self.fault_external_file else '未选择'}"
                self.fault_config_hint.config(
                    text=f"✅ 配置完成 - 来源: {source_text}",
                    foreground='green'
                )
            
            if hasattr(self, 'fault_register_count_label'):
                self.fault_register_count_label.config(
                    text=f"已选择 {len(self.selected_fault_columns)} 个寄存器列进行监控"
                )
            
            # 显示重新配置按钮
            if hasattr(self, 'fault_reconfig_btn'):
                self.fault_reconfig_btn.pack(anchor=tk.W, pady=5)
            
            # 标记配置完成
            self.fault_config_done = True
            
            # 加载故障码定义
            if self.fault_source == "external" and self.fault_external_file:
                print(f"[INFO] 使用外部故障码文件: {self.fault_external_file}")
                from fault_code_parser import load_fault_codes_from_point_table
                load_fault_codes_from_point_table(self.fault_external_file)
            else:
                print("[INFO] 使用内置故障码定义")
            
            print(f"[INFO] 已选择 {len(self.selected_fault_columns)} 个寄存器列")
            
            # 更新进度提示
            self.fault_progress_bar['value'] = 100
            
            # 确保 register_types 存在
            if not hasattr(self, 'register_types'):
                self.register_types = {}
            
            type_summary = []
            user_count = len([c for c in self.selected_fault_columns if c in self.register_types and self.register_types[c] == '用户展示'])
            main_count = len([c for c in self.selected_fault_columns if c in self.register_types and self.register_types[c] == '主故障码'])
            if user_count:
                type_summary.append(f"用户展示{user_count}个")
            if main_count:
                type_summary.append(f"主故障码{main_count}个")
            
            self.fault_progress_label.config(
                text=f"[√] 配置完成！({', '.join(type_summary)})\n    请点击下方【开始故障分析】按钮", 
                foreground="green"
            )
            
            # 更新右侧图表区域的提示
            self.fault_fig.clear()
            self.fault_ax = self.fault_fig.add_subplot(111)
            ready_text = (
                f'✅ 准备就绪\n\n'
                f'已选择 {len(self.selected_fault_columns)} 个故障寄存器列\n\n'
                f'点击左侧「开始故障分析」按钮\n'
                f'开始分析故障数据'
            )
            self.fault_ax.text(0.5, 0.5, ready_text, 
                               ha='center', va='center', fontsize=14, color='green',
                               linespacing=1.5)
            self.fault_ax.set_xlim(0, 1)
            self.fault_ax.set_ylim(0, 1)
            self.fault_ax.axis('off')
            self.fault_canvas.draw()
            
            # 关闭对话框
            dialog.destroy()
            
        except Exception as e:
            print(f"[ERROR] _apply_fault_config 异常: {e}")
            import traceback
            traceback.print_exc()
            # 关闭对话框
            dialog.destroy()
            # 显示错误
            self.fault_progress_label.config(text=f"配置失败: {str(e)[:50]}", foreground="red")
    
    def _update_fault_register_checkboxes(self):
        """更新故障寄存器选择区域的复选框（基于列名）"""
        # 清除现有复选框
        for widget in self.fault_register_frame.winfo_children():
            widget.destroy()
        
        self.fault_register_vars.clear()
        self.fault_register_checkboxes.clear()
        
        if not hasattr(self, 'selected_fault_columns') or not self.selected_fault_columns:
            ttk.Label(self.fault_register_frame, text="未选择任何寄存器列", 
                      foreground='red').pack(anchor=tk.W)
            return
        
        # 确保 register_types 存在
        if not hasattr(self, 'register_types'):
            self.register_types = {}
        
        # 按类别显示
        user_cols = [c for c in self.selected_fault_columns if c in self.register_types and self.register_types[c] == '用户展示']
        main_cols = [c for c in self.selected_fault_columns if c in self.register_types and self.register_types[c] == '主故障码']
        
        ttk.Label(self.fault_register_frame, text=f"检测到 {len(self.selected_fault_columns)} 个故障寄存器列：", 
                  font=('Arial', 9, 'bold')).pack(anchor=tk.W, pady=5)
        
        if user_cols:
            ttk.Label(self.fault_register_frame, text="【用户展示故障码】", 
                      font=('Arial', 9), foreground='blue').pack(anchor=tk.W, padx=5)
            for col in user_cols:
                var = tk.BooleanVar(value=True)
                self.fault_register_vars[col] = var
                cb = ttk.Checkbutton(self.fault_register_frame, text=col, variable=var)
                cb.pack(anchor=tk.W, padx=15)
        
        if main_cols:
            ttk.Label(self.fault_register_frame, text="【主故障码】", 
                      font=('Arial', 9), foreground='green').pack(anchor=tk.W, padx=5, pady=(10, 0))
            for col in main_cols:
                var = tk.BooleanVar(value=True)
                self.fault_register_vars[col] = var
                cb = ttk.Checkbutton(self.fault_register_frame, text=col, variable=var)
                cb.pack(anchor=tk.W, padx=15)
    
    def _on_top_tab_changed(self, event=None):
        """顶层Tab切换事件处理"""
        current_tab = self.top_notebook.index(self.top_notebook.select())
        if current_tab == 0:
            self.data_mode.set("runtime")
            print("[DEBUG] 切换到运行数据解析模式")
        else:
            self.data_mode.set("fault")
            print("[DEBUG] 切换到故障数据解析模式")
            # 注意：配置对话框在选择文件后弹出，而不是切换Tab时
    
    def _create_fault_single_file_widgets(self, parent_frame):
        """创建故障数据单文件分析的界面组件
        
        Args:
            parent_frame: 父容器
        """
        # 主框架：左侧控制面板，右侧图表区域
        main_paned = ttk.PanedWindow(parent_frame, orient=tk.HORIZONTAL)
        main_paned.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)

        # ========== 左侧控制面板（带滚动功能）==========
        left_container = ttk.Frame(main_paned)
        main_paned.add(left_container, weight=0)

        left_inner = ttk.Frame(left_container)
        left_inner.pack(side="left", fill="both", expand=True)
        
        # 创建Canvas和Scrollbar
        fault_left_canvas = tk.Canvas(left_inner, bg='white', highlightthickness=0, width=405)
        fault_left_scrollbar = ttk.Scrollbar(left_inner, orient="vertical", command=fault_left_canvas.yview)
        fault_left_scrollable_frame = ttk.Frame(fault_left_canvas)
        
        # 保存引用
        self.fault_left_canvas = fault_left_canvas
        self.fault_left_scrollbar = fault_left_scrollbar
        self.fault_left_scrollable_frame = fault_left_scrollable_frame

        # 绑定Configure事件
        fault_left_scrollable_frame.bind(
            "<Configure>",
            lambda e: fault_left_canvas.configure(scrollregion=fault_left_canvas.bbox("all"))
        )

        # 创建Canvas窗口
        fault_left_canvas.create_window((0, 0), window=fault_left_scrollable_frame, anchor="nw", width=400)
        fault_left_canvas.configure(yscrollcommand=fault_left_scrollbar.set)

        # 添加鼠标滚轮支持
        def _on_mousewheel(event):
            fault_left_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        fault_left_canvas.bind("<MouseWheel>", _on_mousewheel)
        fault_left_scrollable_frame.bind("<MouseWheel>", _on_mousewheel)

        fault_left_canvas.pack(side="left", fill="y", expand=False)
        fault_left_scrollbar.pack(side="left", fill="y")

        left_frame = fault_left_scrollable_frame
        
        # ========== 文件选择区域 ==========
        file_collapsible = CollapsibleFrame(left_frame, title="文件选择", collapsed=False)
        file_collapsible.pack(fill=tk.X, padx=2, pady=2)
        file_group = file_collapsible.get_content_frame()
        
        fault_file_path_label = ttk.Label(file_group, text="未选择文件", wraplength=400)
        fault_file_path_label.pack(fill=tk.X, pady=2)
        
        file_btn_frame = ttk.Frame(file_group)
        file_btn_frame.pack(fill=tk.X)
        
        ttk.Button(file_btn_frame, text="选择Excel文件", 
                   command=lambda: self.select_fault_file(fault_file_path_label)).pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        # 存储引用
        self.fault_file_path_label = fault_file_path_label
        
        # ========== 故障寄存器选择区域 ==========
        register_collapsible = CollapsibleFrame(left_frame, title="故障寄存器选择", collapsed=False)
        register_collapsible.pack(fill=tk.X, padx=2, pady=2)
        register_group = register_collapsible.get_content_frame()
        
        # 配置状态提示
        self.fault_config_hint = ttk.Label(register_group, 
            text="💡 请先选择Excel文件，然后配置要分析的寄存器", 
            font=('Arial', 9), foreground='blue')
        self.fault_config_hint.pack(anchor=tk.W, pady=5)
        
        # 已选择的寄存器数量提示
        self.fault_register_count_label = ttk.Label(register_group, 
            text="", font=('Arial', 9), foreground='gray')
        self.fault_register_count_label.pack(anchor=tk.W, pady=2)
        
        # 故障寄存器复选框容器
        self.fault_register_frame = ttk.Frame(register_group)
        self.fault_register_frame.pack(fill=tk.X, pady=5)
        
        # 重新配置按钮（初始隐藏，配置完成后显示）
        self.fault_reconfig_btn = ttk.Button(register_group, text="⚙️ 重新配置故障范围", 
                   command=self._show_fault_config_dialog)
        # 初始不显示
        
        # ========== 处理进度区域 ==========
        progress_collapsible = CollapsibleFrame(left_frame, title="处理进度", collapsed=False)
        progress_collapsible.pack(fill=tk.X, padx=2, pady=2)
        progress_frame = progress_collapsible.get_content_frame()
        
        self.fault_progress_label = ttk.Label(progress_frame, text="就绪", foreground="blue")
        self.fault_progress_label.pack(anchor=tk.W)
        
        self.fault_progress_bar = ttk.Progressbar(progress_frame, mode='determinate', length=380)
        self.fault_progress_bar.pack(fill=tk.X, pady=(5, 0))
        
        # ========== 分析按钮区域 ==========
        btn_frame = ttk.Frame(left_frame)
        btn_frame.pack(fill=tk.X, padx=2, pady=10)
        
        # 使用样式让按钮更醒目
        style = ttk.Style()
        style.configure('FaultAnalysis.TButton', font=('Arial', 11, 'bold'))
        
        self.fault_analysis_btn = ttk.Button(
            btn_frame, 
            text="📊 开始故障分析", 
            command=self.start_fault_analysis,
            style='FaultAnalysis.TButton'
        )
        self.fault_analysis_btn.pack(fill=tk.X, pady=5, ipady=8)
        
        # 操作提示
        ttk.Label(btn_frame, text="💡 点击上方按钮开始分析故障数据", 
                  font=('Arial', 9), foreground='gray').pack(anchor=tk.W, pady=2)
        
        # ========== 右侧图表区域（可滚动） ==========
        right_frame = ttk.Frame(main_paned)
        main_paned.add(right_frame, weight=1)
        
        # 创建可滚动的图表区域
        self.fault_chart_frame = ttk.Frame(right_frame)
        self.fault_chart_frame.pack(fill=tk.BOTH, expand=True)
        
        # 创建图表
        self.fault_fig = Figure(figsize=(10, 7), dpi=100)
        self.fault_ax = self.fault_fig.add_subplot(111)
        
        self.fault_canvas = FigureCanvasTkAgg(self.fault_fig, self.fault_chart_frame)
        self.fault_canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        
        # 初始提示 - 更详细的操作指引
        hint_text = (
            '🔧 故障数据分析\n\n'
            '操作步骤：\n'
            '1. 点击左侧「选择Excel文件」选择故障数据文件\n'
            '2. 系统自动检测故障寄存器列\n'
            '3. 点击「开始故障分析」按钮\n\n'
            '支持的寄存器：\n'
            '• 用户展示故障码：20712-20717\n'
            '• 主故障码：20650-20689（偏移量0-39）\n\n'
            '支持故障码：E100-E609'
        )
        self.fault_ax.text(0.5, 0.5, hint_text, 
                           ha='center', va='center', fontsize=12, color='gray',
                           linespacing=1.5)
        self.fault_ax.set_xlim(0, 1)
        self.fault_ax.set_ylim(0, 1)
        self.fault_ax.axis('off')
        self.fault_canvas.draw()
    
    def _create_fault_batch_placeholder(self, parent_frame):
        """创建故障数据批量分析的占位界面
        
        Args:
            parent_frame: 父容器
        """
        # 显示开发中提示
        placeholder_frame = ttk.Frame(parent_frame)
        placeholder_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        ttk.Label(placeholder_frame, text="🔧 故障数据批量分析", 
                  font=('Arial', 16, 'bold')).pack(pady=20)
        
        ttk.Label(placeholder_frame, text="功能开发中，敬请期待...", 
                  font=('Arial', 12), foreground='gray').pack(pady=10)
        
        ttk.Label(placeholder_frame, text="\n计划功能：\n• 批量扫描多个故障数据文件\n• 汇总故障统计\n• 生成故障报告", 
                  font=('Arial', 10), foreground='gray', justify=tk.LEFT).pack(pady=10)
    
    def _create_single_file_widgets(self, parent_frame, mode="runtime"):
        """创建单文件分析的界面组件
        
        Args:
            parent_frame: 父容器
            mode: 数据模式，"runtime" 或 "fault"
        """
        # 主框架：左侧控制面板，右侧图表区域（带可拖动分割条）
        main_paned = ttk.PanedWindow(parent_frame, orient=tk.HORIZONTAL)
        main_paned.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)

        # ========== 左侧控制面板（带滚动功能）==========
        # 创建外层容器，用于放置Canvas和Scrollbar
        left_container = ttk.Frame(main_paned)
        # 【修复】设置初始宽度，避免在某些电脑上初始宽度为0导致左侧面板不可见
        main_paned.add(left_container, weight=0)

        # 创建内部容器，让滚动条紧挨着内容区域
        left_inner = ttk.Frame(left_container)
        left_inner.pack(side="left", fill="both", expand=True)
        
        # 创建Canvas和Scrollbar，用于左侧面板的滚动
        self.left_canvas = tk.Canvas(left_inner, bg='white', highlightthickness=0, width=405)
        self.left_scrollbar = ttk.Scrollbar(left_inner, orient="vertical", command=self.left_canvas.yview)
        self.left_scrollable_frame = ttk.Frame(self.left_canvas)

        # 绑定Configure事件，自动调整Canvas的滚动区域
        self.left_scrollable_frame.bind(
            "<Configure>",
            lambda e: self.left_canvas.configure(scrollregion=self.left_canvas.bbox("all"))
        )

        # 创建Canvas上的窗口，设置固定宽度
        self.left_canvas_window = self.left_canvas.create_window((0, 0), window=self.left_scrollable_frame, anchor="nw", width=400)
        self.left_canvas.configure(yscrollcommand=self.left_scrollbar.set)

        # 添加鼠标滚轮滚动支持
        def _on_left_mousewheel(event):
            # Windows 和 Linux
            self.left_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        self.left_canvas.bind("<MouseWheel>", _on_left_mousewheel)
        self.left_scrollable_frame.bind("<MouseWheel>", _on_left_mousewheel)

        # Linux 鼠标滚轮支持
        def _on_left_button4(event):
            self.left_canvas.yview_scroll(-1, "units")
        def _on_left_button5(event):
            self.left_canvas.yview_scroll(1, "units")
        self.left_canvas.bind("<Button-4>", _on_left_button4)
        self.left_canvas.bind("<Button-5>", _on_left_button5)
        self.left_scrollable_frame.bind("<Button-4>", _on_left_button4)
        self.left_scrollable_frame.bind("<Button-5>", _on_left_button5)

        # 布局Canvas和Scrollbar（滚动条紧挨着Canvas右侧）
        self.left_canvas.pack(side="left", fill="y", expand=False)
        self.left_scrollbar.pack(side="left", fill="y")

        # 在left_scrollable_frame上创建所有左侧组件
        left_frame = self.left_scrollable_frame
        
        # 文件选择区域（默认展开）
        self.file_collapsible = CollapsibleFrame(left_frame, title="文件选择", collapsed=False)
        self.file_collapsible.pack(fill=tk.X, padx=2, pady=2)
        file_group = self.file_collapsible.get_content_frame()
        
        self.file_path_label = ttk.Label(file_group, text="未选择文件", wraplength=400)
        self.file_path_label.pack(fill=tk.X, pady=2)
        
        file_btn_frame = ttk.Frame(file_group)
        file_btn_frame.pack(fill=tk.X)
        
        ttk.Button(file_btn_frame, text="选择Excel文件", command=self.select_file).pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        # ========== 处理进度区域 ==========
        self.progress_collapsible = CollapsibleFrame(left_frame, title="处理进度", collapsed=False)
        self.progress_collapsible.pack(fill=tk.X, padx=2, pady=2)
        progress_frame = self.progress_collapsible.get_content_frame()
        self.progress_frame = progress_frame  # 保持兼容性
        
        self.progress_label = ttk.Label(progress_frame, text="就绪", foreground="blue")
        self.progress_label.pack(anchor=tk.W)
        
        self.progress_bar = ttk.Progressbar(progress_frame, mode='determinate', length=380)
        self.progress_bar.pack(fill=tk.X, pady=(5, 0))
        
        self.progress_detail = ttk.Label(progress_frame, text="", font=("Arial", 8))
        self.progress_detail.pack(anchor=tk.W, pady=(2, 0))
        
        # 耗时显示标签
        self.time_label = ttk.Label(progress_frame, text="耗时: 0.0秒", font=("Arial", 9), foreground="gray")
        self.time_label.pack(anchor=tk.W, pady=(2, 0))
        
        # ========== 常用分组窗口（默认折叠）==========
        self.groups_collapsible = CollapsibleFrame(left_frame, title="常用分组", collapsed=True)
        self.groups_collapsible.pack(fill=tk.X, padx=5, pady=5)
        groups_group = self.groups_collapsible.get_content_frame()
        
        # 分组选择下拉框
        group_select_frame = ttk.Frame(groups_group)
        group_select_frame.pack(fill=tk.X, pady=(0, 5))
        
        ttk.Label(group_select_frame, text="选择分组:").pack(side=tk.LEFT)
        self.group_var = tk.StringVar()
        self.group_combobox = ttk.Combobox(group_select_frame, textvariable=self.group_var, state='readonly', width=20)
        self.group_combobox.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        self.group_combobox.bind('<<ComboboxSelected>>', self.on_group_selected)
        
        # 操作按钮区域
        group_btn_frame = ttk.Frame(groups_group)
        group_btn_frame.pack(fill=tk.X, pady=5)
        
        ttk.Button(group_btn_frame, text="新建", command=self.create_new_group, width=8).pack(side=tk.LEFT, padx=2)
        ttk.Button(group_btn_frame, text="编辑", command=self.edit_current_group, width=8).pack(side=tk.LEFT, padx=2)
        ttk.Button(group_btn_frame, text="删除", command=self.delete_current_group, width=8).pack(side=tk.LEFT, padx=2)
        ttk.Button(group_btn_frame, text="重命名", command=self.rename_current_group, width=8).pack(side=tk.LEFT, padx=2)
        ttk.Button(group_btn_frame, text="自定义列", command=self.open_custom_column_editor, width=10).pack(side=tk.LEFT, padx=2)
        
        # 当前分组列显示区域
        ttk.Label(groups_group, text="当前分组包含的列:", font=('Arial', 9, 'bold')).pack(anchor=tk.W, pady=(5, 2))
        
        # 创建分组列列表（可滚动）
        group_list_canvas_frame = ttk.Frame(groups_group)
        group_list_canvas_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 5))
        
        self.group_canvas = tk.Canvas(group_list_canvas_frame, bg='white', height=100)
        group_list_scrollbar = ttk.Scrollbar(group_list_canvas_frame, orient="vertical", command=self.group_canvas.yview)
        self.group_scrollable_frame = ttk.Frame(self.group_canvas)
        
        self.group_scrollable_frame.bind(
            "<Configure>",
            lambda e: self.group_canvas.configure(scrollregion=self.group_canvas.bbox("all"))
        )
        
        self.group_canvas.create_window((0, 0), window=self.group_scrollable_frame, anchor="nw")
        self.group_canvas.configure(yscrollcommand=group_list_scrollbar.set)
        
        self.group_canvas.pack(side="left", fill="both", expand=True)
        group_list_scrollbar.pack(side="right", fill="y")
        
        # 绑定滚轮事件
        self.group_canvas.bind("<MouseWheel>", self.on_group_canvas_scroll)
        self.group_canvas.bind("<Button-4>", self.on_group_canvas_scroll)
        self.group_canvas.bind("<Button-5>", self.on_group_canvas_scroll)
        self.group_scrollable_frame.bind("<MouseWheel>", self.on_group_canvas_scroll)
        self.group_scrollable_frame.bind("<Button-4>", self.on_group_canvas_scroll)
        self.group_scrollable_frame.bind("<Button-5>", self.on_group_canvas_scroll)
        
        # 存储分组列的标签
        self.group_labels = {}
        
        # ========== 数据诊断模块（默认折叠）==========
        self.diagnosis_collapsible = CollapsibleFrame(left_frame, title="数据诊断", collapsed=True)
        self.diagnosis_collapsible.pack(fill=tk.X, padx=5, pady=5)
        diagnosis_group = self.diagnosis_collapsible.get_content_frame()
        
        # 预设策略选择
        strategy_select_frame = ttk.Frame(diagnosis_group)
        strategy_select_frame.pack(fill=tk.X, pady=2)
        ttk.Label(strategy_select_frame, text="预设策略:").pack(side=tk.LEFT)
        self.diagnosis_strategy_var = tk.StringVar()
        strategy_combobox = ttk.Combobox(strategy_select_frame, textvariable=self.diagnosis_strategy_var, 
                                        state='readonly', width=25)
        strategy_combobox['values'] = [
            '波动剧烈检测',
            '同增同减检测',
            '数据范围异常',
            '自定义策略...'
        ]
        strategy_combobox.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        strategy_combobox.current(0)
        strategy_combobox.bind('<<ComboboxSelected>>', self.on_strategy_selected)
        
        # 策略参数配置区域（初始隐藏，选择策略后显示）
        self.strategy_config_frame = ttk.Frame(diagnosis_group)
        
        # 波动检测参数
        self.volatility_config = ttk.Frame(self.strategy_config_frame)
        ttk.Label(self.volatility_config, text="波动阈值(%):").pack(side=tk.LEFT)
        self.volatility_threshold_var = tk.StringVar(value="30")
        ttk.Entry(self.volatility_config, textvariable=self.volatility_threshold_var, width=8).pack(side=tk.LEFT, padx=5)
        ttk.Label(self.volatility_config, text="检测列:").pack(side=tk.LEFT, padx=(10, 0))
        self.volatility_column_var = tk.StringVar()
        ttk.Combobox(self.volatility_config, textvariable=self.volatility_column_var, 
                    state='readonly', width=15).pack(side=tk.LEFT, padx=5)
        
        # 同增同减参数
        self.correlation_config = ttk.Frame(self.strategy_config_frame)
        ttk.Label(self.correlation_config, text="列1:").pack(side=tk.LEFT)
        self.correlation_col1_var = tk.StringVar()
        ttk.Combobox(self.correlation_config, textvariable=self.correlation_col1_var, 
                    state='readonly', width=15).pack(side=tk.LEFT, padx=5)
        ttk.Label(self.correlation_config, text="列2:").pack(side=tk.LEFT, padx=(10, 0))
        self.correlation_col2_var = tk.StringVar()
        ttk.Combobox(self.correlation_config, textvariable=self.correlation_col2_var, 
                    state='readonly', width=15).pack(side=tk.LEFT, padx=5)
        
        # 数据范围参数（分成两行显示，避免窗口过窄时看不到最大值）
        self.range_config = ttk.Frame(self.strategy_config_frame)
        
        # 第一行：检测列选择
        row1 = ttk.Frame(self.range_config)
        row1.pack(fill=tk.X, pady=2)
        ttk.Label(row1, text="检测列:").pack(side=tk.LEFT)
        self.range_column_var = tk.StringVar()
        ttk.Combobox(row1, textvariable=self.range_column_var, 
                    state='readonly', width=15).pack(side=tk.LEFT, padx=5)
        
        # 第二行：范围设置
        row2 = ttk.Frame(self.range_config)
        row2.pack(fill=tk.X, pady=2)
        ttk.Label(row2, text="最小值:").pack(side=tk.LEFT)
        self.range_min_var = tk.StringVar(value="0")
        ttk.Entry(row2, textvariable=self.range_min_var, width=10).pack(side=tk.LEFT, padx=5)
        ttk.Label(row2, text="最大值:").pack(side=tk.LEFT, padx=(15, 0))
        self.range_max_var = tk.StringVar(value="10000")
        ttk.Entry(row2, textvariable=self.range_max_var, width=10).pack(side=tk.LEFT, padx=5)
        
        # 显示默认配置（波动检测）
        self.strategy_config_frame.pack(fill=tk.X, pady=5)
        self.volatility_config.pack(fill=tk.X)
        self.current_config_frame = self.volatility_config
        
        # 操作按钮
        diagnosis_btn_frame = ttk.Frame(diagnosis_group)
        diagnosis_btn_frame.pack(fill=tk.X, pady=5)
        ttk.Button(diagnosis_btn_frame, text="运行诊断", command=self.run_diagnosis, width=12).pack(side=tk.LEFT, padx=2)
        ttk.Button(diagnosis_btn_frame, text="查看报告", command=self.view_diagnosis_report, width=12).pack(side=tk.LEFT, padx=2)
        ttk.Button(diagnosis_btn_frame, text="清除结果", command=self.clear_diagnosis, width=12).pack(side=tk.LEFT, padx=2)
        
        # 诊断结果状态显示
        self.diagnosis_status_label = ttk.Label(diagnosis_group, text="状态: 未运行", foreground="gray")
        self.diagnosis_status_label.pack(anchor=tk.W, pady=2)
        
        # 列选择区域（默认展开）
        self.column_collapsible = CollapsibleFrame(left_frame, title="数据列选择", collapsed=False)
        self.column_collapsible.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.column_group = self.column_collapsible.get_content_frame()
        
        # 添加搜索框
        search_frame = ttk.Frame(self.column_group)
        search_frame.pack(fill=tk.X, pady=(0, 10))
        ttk.Label(search_frame, text="搜索列名:").pack(side=tk.LEFT)
        self.search_var = tk.StringVar()
        self.search_entry = ttk.Entry(search_frame, textvariable=self.search_var)
        self.search_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        self.search_entry.bind('<KeyRelease>', self.filter_columns)
        
        # 添加全选/全不选按钮
        select_frame = ttk.Frame(self.column_group)
        select_frame.pack(fill=tk.X, pady=(0, 10))
        ttk.Button(select_frame, text="全选", command=self.select_all).pack(side=tk.LEFT, padx=2)
        ttk.Button(select_frame, text="全不选", command=self.deselect_all).pack(side=tk.LEFT, padx=2)
        
        # 创建可滚动的列名列表
        canvas_frame = ttk.Frame(self.column_group)
        canvas_frame.pack(fill=tk.BOTH, expand=True)

        self.canvas = tk.Canvas(canvas_frame, bg='white', height=200)  # 设置最小高度确保滚动条可用
        self.column_scrollbar = ttk.Scrollbar(canvas_frame, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = ttk.Frame(self.canvas)

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )

        self.canvas_window = self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.column_scrollbar.set)

        self.canvas.pack(side="left", fill="both", expand=True)
        self.column_scrollbar.pack(side="left", fill="y")  # 滚动条紧挨着Canvas左侧

        # ========== 已勾选数据显示区域 ==========
        
        # 左侧区域的滚轮事件：用于上下滚动列表
        # 绑定到Canvas和scrollable_frame，确保无论鼠标在哪里都能响应
        self.canvas.bind("<MouseWheel>", self.on_canvas_scroll)
        self.canvas.bind("<Button-4>", self.on_canvas_scroll)
        self.canvas.bind("<Button-5>", self.on_canvas_scroll)
        
        self.scrollable_frame.bind("<MouseWheel>", self.on_canvas_scroll)
        self.scrollable_frame.bind("<Button-4>", self.on_canvas_scroll)
        self.scrollable_frame.bind("<Button-5>", self.on_canvas_scroll)
        
        # 生成图表按钮
        self.generate_btn = ttk.Button(left_frame, text="生成折线图", command=self.generate_plot, state=tk.DISABLED)
        self.generate_btn.pack(fill=tk.X, padx=5, pady=10)
        
        # 图表选项
        options_group = ttk.LabelFrame(left_frame, text="图表选项", padding=10)
        options_group.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(options_group, text="图表标题:").pack(anchor=tk.W)
        self.title_var = tk.StringVar(value="数据趋势图")
        ttk.Entry(options_group, textvariable=self.title_var).pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(options_group, text="X轴标签:").pack(anchor=tk.W)
        self.xlabel_var = tk.StringVar(value="时间")
        ttk.Entry(options_group, textvariable=self.xlabel_var).pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(options_group, text="Y轴标签:").pack(anchor=tk.W)
        self.ylabel_var = tk.StringVar(value="数值")
        ttk.Entry(options_group, textvariable=self.ylabel_var).pack(fill=tk.X)
        
        # 保存图片按钮
        self.save_btn = ttk.Button(left_frame, text="保存图片", command=self.save_plot, state=tk.DISABLED)
        self.save_btn.pack(fill=tk.X, padx=5, pady=5)
        
        # ========== 右侧图表区域 ==========
        right_frame = ttk.Frame(main_paned)
        main_paned.add(right_frame, weight=3)
        
        # 设置分割条初始位置（延迟执行，确保窗口已显示）
        # 【修复】增加延迟时间和重试次数，确保在不同电脑上都能正确设置分割条位置
        self._init_sash_position(main_paned, 400, 0)
        
        # 【新增】使用 PanedWindow 分隔图表和图例面板
        chart_paned = ttk.PanedWindow(right_frame, orient=tk.HORIZONTAL)
        chart_paned.pack(fill=tk.BOTH, expand=True)
        
        # 图表区域
        chart_frame = ttk.Frame(chart_paned)
        chart_paned.add(chart_frame, weight=4)
        
        # 创建matplotlib图表
        self.figure = Figure(figsize=(10, 6), dpi=100)
        self.ax = self.figure.add_subplot(111)
        self.ax.set_title("请选择Excel文件并勾选要显示的数据列")
        self.ax.set_xlabel("时间")
        self.ax.set_ylabel("数值")
        
        self.canvas_plot = FigureCanvasTkAgg(self.figure, chart_frame)
        self.canvas_plot_widget = self.canvas_plot.get_tk_widget()
        self.canvas_plot_widget.pack(fill=tk.BOTH, expand=True)
        
        # 添加工具栏
        toolbar_frame = ttk.Frame(chart_frame)
        toolbar_frame.pack(fill=tk.X)
        from matplotlib.backends.backend_tkagg import NavigationToolbar2Tk
        toolbar = NavigationToolbar2Tk(self.canvas_plot, toolbar_frame)
        toolbar.update()
        
        # 【新增】图例面板
        legend_frame = ttk.LabelFrame(chart_paned, text="图例", padding=5)
        chart_paned.add(legend_frame, weight=1)
        
        # 图例面板内部结构
        legend_container = ttk.Frame(legend_frame)
        legend_container.pack(fill=tk.BOTH, expand=True)
        
        # 创建Canvas用于滚动
        self.legend_canvas = tk.Canvas(legend_container, highlightthickness=0)
        legend_scrollbar = ttk.Scrollbar(legend_container, orient="vertical", command=self.legend_canvas.yview)
        self.legend_scrollable_frame = ttk.Frame(self.legend_canvas)
        
        self.legend_scrollable_frame.bind(
            "<Configure>",
            lambda e: self.legend_canvas.configure(scrollregion=self.legend_canvas.bbox("all"))
        )
        
        self.legend_canvas.create_window((0, 0), window=self.legend_scrollable_frame, anchor="nw")
        self.legend_canvas.configure(yscrollcommand=legend_scrollbar.set)
        
        self.legend_canvas.pack(side="left", fill="both", expand=True)
        legend_scrollbar.pack(side="right", fill="y")
        
        # 绑定滚轮事件到图例面板
        self.legend_canvas.bind("<MouseWheel>", self._on_legend_scroll)
        self.legend_canvas.bind("<Button-4>", self._on_legend_scroll)
        self.legend_canvas.bind("<Button-5>", self._on_legend_scroll)
        self.legend_scrollable_frame.bind("<MouseWheel>", self._on_legend_scroll)
        
        # 保存图例面板引用
        self.legend_panel = legend_frame
        self.chart_paned = chart_paned
        
        # 【修复】设置图例面板分割条初始位置，确保图例面板可见
        self._init_chart_paned_sash()
        
        # 绑定 Home 按钮的复位功能 - 恢复到生成折线图时的初始状态
        if hasattr(toolbar, 'home'):
            toolbar.home = self.reset_to_original_view
        
        # 绑定滚轮缩放事件 - 只在右侧图表区域生效
        self.canvas_plot_widget.bind("<MouseWheel>", self.on_scroll)
        self.canvas_plot_widget.bind("<Button-4>", self.on_scroll)
        self.canvas_plot_widget.bind("<Button-5>", self.on_scroll)
        
        # 绑定鼠标点击事件 - 显示数据点信息
        self.canvas_plot.mpl_connect('button_press_event', self.on_plot_click)
        
        # 绑定鼠标拖动事件 - 左键拖动平移图表
        self.canvas_plot.mpl_connect('button_press_event', self.on_mouse_press)
        self.canvas_plot.mpl_connect('button_release_event', self.on_mouse_release)
        self.canvas_plot.mpl_connect('motion_notify_event', self.on_mouse_drag)
        
        # 添加焦点设置，确保右侧图表区域能优先接收滚轮事件
        self.canvas_plot_widget.bind("<Enter>", lambda e: self.canvas_plot_widget.focus_set())
        
        # Tooltip显示相关变量
        self.tooltip = None
        self.annotation = None
        self.last_selected_columns = []
        self.plot_data_cache = {} # 缓存绘制的数据，用于点击查询
        
        # 缩放状态变量
        self.x_zoom_factor = 1.0
        self.y_zoom_factor = 1.0
        self.zoom_center = None
        
        # 原始坐标轴范围
        self.original_xlim = None
        self.original_ylim = None
        
        # 鼠标拖动平移相关变量
        self.drag_start = None  # 拖动起始位置
        self.drag_xlim = None  # 拖动时的X轴范围
    
    def _create_batch_analysis_widgets(self, parent_frame, mode="runtime"):
        """创建批量分析的界面组件
        
        Args:
            parent_frame: 父容器
            mode: 数据模式，"runtime" 或 "fault"
        """
        # 主框架：左侧配置面板，右侧结果展示区域
        self.batch_paned = ttk.PanedWindow(parent_frame, orient=tk.HORIZONTAL)
        self.batch_paned.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)
        
        # ========== 左侧配置面板 ==========
        left_frame = ttk.Frame(self.batch_paned)
        
        # ---------- 文件夹选择区域 ----------
        folder_group = ttk.LabelFrame(left_frame, text="文件夹选择", padding=10)
        folder_group.pack(fill=tk.X, padx=5, pady=5)
        
        self.batch_folder_label = ttk.Label(folder_group, text="未选择文件夹", wraplength=350)
        self.batch_folder_label.pack(fill=tk.X, pady=2)
        
        folder_btn_frame = ttk.Frame(folder_group)
        folder_btn_frame.pack(fill=tk.X, pady=5)
        ttk.Button(folder_btn_frame, text="选择文件夹", command=self.select_batch_folder).pack(side=tk.LEFT, padx=2)
        ttk.Button(folder_btn_frame, text="扫描文件", command=self.scan_excel_files).pack(side=tk.LEFT, padx=2)
        
        # ---------- 进度显示区域（移到文件夹选择下面）----------
        progress_group = ttk.LabelFrame(left_frame, text="分析进度", padding=10)
        progress_group.pack(fill=tk.X, padx=5, pady=5)
        
        self.batch_progress_label = ttk.Label(progress_group, text="就绪")
        self.batch_progress_label.pack(anchor=tk.W)
        
        self.batch_progress_bar = ttk.Progressbar(progress_group, mode='determinate')
        self.batch_progress_bar.pack(fill=tk.X, pady=5)
        
        # 批量分析耗时显示标签
        self.batch_time_label = ttk.Label(progress_group, text="耗时: 0.0秒", font=("Arial", 9), foreground="gray")
        self.batch_time_label.pack(anchor=tk.W, pady=(2, 0))
        
        # ---------- 数据类型选择区域 ----------
        type_select_frame = ttk.LabelFrame(left_frame, text="数据类型", padding=5)
        type_select_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(type_select_frame, text="选择要分析的数据类型:").pack(side=tk.LEFT)
        self.data_type_var = tk.StringVar(value="BATTERY")
        type_combo = ttk.Combobox(type_select_frame, textvariable=self.data_type_var, 
                                   state='readonly', width=12)
        type_combo['values'] = ['BATTERY', 'PCS']
        type_combo.pack(side=tk.LEFT, padx=5)
        type_combo.bind('<<ComboboxSelected>>', self._on_data_type_change)
        
        # 类型统计标签
        self.type_stats_label = ttk.Label(type_select_frame, text="(电池: 0, PCS: 0)")
        self.type_stats_label.pack(side=tk.LEFT, padx=5)
        
        # ---------- 文件列表区域 ----------
        file_list_group = ttk.LabelFrame(left_frame, text="文件列表", padding=10)
        file_list_group.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 文件列表（带滚动条）
        list_frame = ttk.Frame(file_list_group)
        list_frame.pack(fill=tk.BOTH, expand=True)
        
        self.file_listbox = tk.Listbox(list_frame, selectmode=tk.EXTENDED, height=6)
        file_scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.file_listbox.yview)
        self.file_listbox.configure(yscrollcommand=file_scrollbar.set)
        
        self.file_listbox.pack(side="left", fill="both", expand=True)
        file_scrollbar.pack(side="right", fill="y")
        
        # 文件统计标签
        self.file_count_label = ttk.Label(file_list_group, text="共 0 个文件")
        self.file_count_label.pack(anchor=tk.W, pady=2)
        
        # ---------- 关键值追踪区域 ----------
        key_values_group = ttk.LabelFrame(left_frame, text="关键值追踪（可选）", padding=5)
        key_values_group.pack(fill=tk.X, padx=5, pady=5)
        
        # 说明文字
        ttk.Label(key_values_group, text="选择要追踪的关键列，分析后显示在异常文件列表中:", 
                  font=('Arial', 8), foreground='gray').pack(anchor=tk.W)
        
        # 存储关键值追踪的变量
        self.key_value_vars = []  # 存储选择的关键值列变量
        self.key_value_combos = []  # 存储下拉框引用
        self.key_value_frames = []  # 存储行框架
        
        # 关键值行容器
        self.key_values_container = ttk.Frame(key_values_group)
        self.key_values_container.pack(fill=tk.X, pady=2)
        
        # 操作按钮
        kv_btn_frame = ttk.Frame(key_values_group)
        kv_btn_frame.pack(fill=tk.X, pady=2)
        ttk.Button(kv_btn_frame, text="+ 添加关键值", command=self._add_key_value_row, width=12).pack(side=tk.LEFT, padx=2)
        ttk.Button(kv_btn_frame, text="清空", command=self._clear_key_values, width=8).pack(side=tk.LEFT, padx=2)
        
        # 默认添加3个关键值行
        self._add_key_value_row()  # 第1行
        self._add_key_value_row()  # 第2行
        self._add_key_value_row()  # 第3行
        
        # ---------- 异常条件设置区域 ----------
        condition_group = ttk.LabelFrame(left_frame, text="异常条件设置", padding=5)
        condition_group.pack(fill=tk.X, padx=5, pady=5)
        
        # 存储异常条件的变量
        self.condition_frames = []
        self.condition_enabled = []
        self.condition_types = []
        self.condition_columns = []  # 每个条件的列列表（最多5列，用于同增同减检测）
        self.condition_operators = []  # 每个条件的运算符列表
        self.condition_compare_type = []  # 比较类型 (> < = 等)
        self.condition_threshold = []  # 阈值
        self.condition_min_max = []  # 超出范围的最小最大值
        self.condition_column_combos = []  # 列下拉框列表（每个条件最多3个）
        self.operator_combos = []  # 运算符下拉框列表
        self.params_frames = []
        self.condition_count = 0  # 当前条件数量
        self.condition_parent = condition_group  # 保存父容器引用
        
        # 配置保存文件
        self.batch_conditions_config_file = "batch_conditions_config.json"
        
        # 条件操作按钮区域
        condition_btn_frame = ttk.Frame(condition_group)
        condition_btn_frame.pack(fill=tk.X, pady=2)
        
        ttk.Button(condition_btn_frame, text="+ 添加条件", command=self._add_condition_row, width=12).pack(side=tk.LEFT, padx=2)
        ttk.Button(condition_btn_frame, text="保存配置", command=self._save_batch_conditions_config, width=10).pack(side=tk.LEFT, padx=2)
        
        # 配置选择下拉框
        ttk.Label(condition_btn_frame, text="配置:").pack(side=tk.LEFT, padx=(10, 2))
        self.saved_config_var = tk.StringVar()
        self.saved_config_combo = ttk.Combobox(condition_btn_frame, textvariable=self.saved_config_var, 
                                                state='readonly', width=15)
        self.saved_config_combo.pack(side=tk.LEFT, padx=2)
        self.saved_config_combo.bind('<<ComboboxSelected>>', self._load_batch_conditions_config)
        
        ttk.Button(condition_btn_frame, text="删除配置", command=self._delete_batch_conditions_config, width=10).pack(side=tk.LEFT, padx=2)
        
        # 条件列表容器（可滚动）
        condition_container = ttk.Frame(condition_group)
        condition_container.pack(fill=tk.BOTH, expand=True, pady=5)
        
        self.condition_list_frame = ttk.Frame(condition_container)
        self.condition_list_frame.pack(fill=tk.BOTH, expand=True)
        
        # 存储列名（扫描后更新）
        self.battery_columns = []
        self.pcs_columns = []
        self.current_columns = []  # 当前选中类型对应的列名
        
        # 加载已保存的配置列表
        self._load_batch_config_list()
        
        # 默认创建一个条件行
        self._create_condition_row(self.condition_list_frame, 0, enabled=False)
        
        # ---------- 操作按钮区域 ----------
        action_frame = ttk.Frame(left_frame)
        action_frame.pack(fill=tk.X, padx=5, pady=10)
        
        # 第一行：开始和暂停按钮
        btn_frame1 = ttk.Frame(action_frame)
        btn_frame1.pack(fill=tk.X, pady=2)
        
        self.start_btn = ttk.Button(btn_frame1, text="开始批量分析", command=self.run_batch_analysis)
        self.start_btn.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 2))
        
        self.pause_btn = ttk.Button(btn_frame1, text="暂停", command=self.toggle_pause_batch, state=tk.DISABLED)
        self.pause_btn.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(2, 2))
        
        self.stop_btn = ttk.Button(btn_frame1, text="停止", command=self.stop_batch_analysis, state=tk.DISABLED)
        self.stop_btn.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(2, 0))
        
        ttk.Button(action_frame, text="导出结果", command=self.export_batch_results).pack(fill=tk.X, pady=2)
        ttk.Button(action_frame, text="清空结果", command=self.clear_batch_results).pack(fill=tk.X, pady=2)
        
        # 添加左侧面板到 PanedWindow（必须先创建内容再 add）
        self.batch_paned.add(left_frame, weight=0)
        
        # ========== 右侧结果展示区域 ==========
        right_frame = ttk.Frame(self.batch_paned)
        
        # 结果统计区域
        stats_frame = ttk.LabelFrame(right_frame, text="分析统计", padding=10)
        stats_frame.pack(fill=tk.X, padx=5, pady=5)
        
        self.batch_stats_label = ttk.Label(stats_frame, text="尚未开始分析", font=('Arial', 10))
        self.batch_stats_label.pack(anchor=tk.W)
        
        # 使用 PanedWindow 分隔异常文件列表和异常时间点列表
        self.result_paned = ttk.PanedWindow(right_frame, orient=tk.VERTICAL)
        self.result_paned.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 结果列表区域
        result_frame = ttk.LabelFrame(self.result_paned, text="异常文件列表（双击跳转到单文件分析）", padding=5)
        self.result_paned.add(result_frame, weight=1)
        
        # 创建内部容器用于放置 Treeview 和滚动条
        tree_container = ttk.Frame(result_frame)
        tree_container.pack(fill=tk.BOTH, expand=True)
        
        # 创建 Treeview 显示结果（添加关键值列）
        columns = ('文件名', '路径', '设备编码', '异常类型', '异常详情', '关键值1', '关键值2', '关键值3')
        self.result_tree = ttk.Treeview(tree_container, columns=columns, show='headings')
        
        self.result_tree.heading('文件名', text='文件名')
        self.result_tree.heading('路径', text='相对路径')
        self.result_tree.heading('设备编码', text='设备编码')
        self.result_tree.heading('异常类型', text='异常类型')
        self.result_tree.heading('异常详情', text='异常详情')
        self.result_tree.heading('关键值1', text='关键值1')
        self.result_tree.heading('关键值2', text='关键值2')
        self.result_tree.heading('关键值3', text='关键值3')
        
        # 设置列宽，stretch=tk.NO 让列宽固定，水平滚动条才能正常工作
        self.result_tree.column('文件名', width=180, minwidth=120, stretch=tk.NO)
        self.result_tree.column('路径', width=120, minwidth=80, stretch=tk.NO)
        self.result_tree.column('设备编码', width=120, minwidth=80, stretch=tk.NO)
        self.result_tree.column('异常类型', width=100, minwidth=80, stretch=tk.NO)
        self.result_tree.column('异常详情', width=300, minwidth=150, stretch=tk.NO)
        self.result_tree.column('关键值1', width=150, minwidth=100, stretch=tk.NO)
        self.result_tree.column('关键值2', width=150, minwidth=100, stretch=tk.NO)
        self.result_tree.column('关键值3', width=150, minwidth=100, stretch=tk.NO)
        
        # 垂直滚动条
        result_vscrollbar = ttk.Scrollbar(tree_container, orient="vertical", command=self.result_tree.yview)
        self.result_tree.configure(yscrollcommand=result_vscrollbar.set)
        
        # 水平滚动条
        result_hscrollbar = ttk.Scrollbar(tree_container, orient="horizontal", command=self.result_tree.xview)
        self.result_tree.configure(xscrollcommand=result_hscrollbar.set)
        
        # 使用 grid 布局放置 Treeview 和滚动条
        self.result_tree.grid(row=0, column=0, sticky="nsew")
        result_vscrollbar.grid(row=0, column=1, sticky="ns")
        result_hscrollbar.grid(row=1, column=0, sticky="ew")
        
        # 配置 grid 权重，让 Treeview 可以扩展
        tree_container.grid_rowconfigure(0, weight=1)
        tree_container.grid_columnconfigure(0, weight=1)
        
        # 绑定双击事件，跳转到单文件分析
        self.result_tree.bind('<Double-Button-1>', self._on_result_double_click)
        
        # === 异常时间面板（可折叠） ===
        # 创建一个带折叠功能的面板
        self.anomaly_time_frame = ttk.LabelFrame(self.result_paned, text="异常时间点列表（点击文件名展开/折叠）", padding=5)
        self.result_paned.add(self.anomaly_time_frame, weight=2)  # 异常时间点列表占更大比例
        
        # 创建内部容器
        anomaly_tree_container = ttk.Frame(self.anomaly_time_frame)
        anomaly_tree_container.pack(fill=tk.BOTH, expand=True)
        
        # 创建 Treeview 显示异常时间（支持层级折叠）
        # 使用 Treeview 的层级结构，文件名为父节点，异常时间为子节点
        self.anomaly_time_tree = ttk.Treeview(anomaly_tree_container, columns=('异常时间', '异常类型', '异常列', '异常详情'), show='tree headings')
        
        # 设置列
        self.anomaly_time_tree.heading('#0', text='文件名')
        self.anomaly_time_tree.heading('异常时间', text='异常时间')
        self.anomaly_time_tree.heading('异常类型', text='异常类型')
        self.anomaly_time_tree.heading('异常列', text='异常列')
        self.anomaly_time_tree.heading('异常详情', text='异常详情')
        
        # 设置列宽
        # #0 列（文件名）使用 stretch=tk.YES 自动填充剩余空间，避免右侧空白
        # 其他列固定宽度，确保内容对齐
        self.anomaly_time_tree.column('#0', width=200, minwidth=120, stretch=tk.NO)
        self.anomaly_time_tree.column('异常时间', width=160, minwidth=100, stretch=tk.NO)
        self.anomaly_time_tree.column('异常类型', width=100, minwidth=80, stretch=tk.NO)
        self.anomaly_time_tree.column('异常列', width=180, minwidth=80, stretch=tk.NO)
        # 异常详情列：自动扩展填充剩余空间，minwidth 确保最小可读宽度
        self.anomaly_time_tree.column('异常详情', width=400, minwidth=300, stretch=tk.YES)
        
        # 滚动条
        anomaly_vscrollbar = ttk.Scrollbar(anomaly_tree_container, orient="vertical", command=self.anomaly_time_tree.yview)
        self.anomaly_time_tree.configure(yscrollcommand=anomaly_vscrollbar.set)
        
        anomaly_hscrollbar = ttk.Scrollbar(anomaly_tree_container, orient="horizontal", command=self.anomaly_time_tree.xview)
        self.anomaly_time_tree.configure(xscrollcommand=anomaly_hscrollbar.set)
        
        # 使用 grid 布局
        self.anomaly_time_tree.grid(row=0, column=0, sticky="nsew")
        anomaly_vscrollbar.grid(row=0, column=1, sticky="ns")
        anomaly_hscrollbar.grid(row=1, column=0, sticky="ew")
        
        # 配置 grid 权重
        anomaly_tree_container.grid_rowconfigure(0, weight=1)
        anomaly_tree_container.grid_columnconfigure(0, weight=1)
        
        # 添加右侧面板到 PanedWindow
        self.batch_paned.add(right_frame, weight=1)
        
        # 设置 result_paned 的初始分割条位置（异常时间点列表占更多空间）
        def init_result_paned_sash(retry=0):
            try:
                if hasattr(self, 'result_paned') and self.result_paned:
                    total_height = self.result_paned.winfo_height()
                    if total_height > 100:
                        # 异常文件列表占 30%，异常时间点列表占 70%
                        file_list_height = int(total_height * 0.25)
                        self.result_paned.sashpos(0, file_list_height)
                        print(f"[DEBUG] result_paned 分割条位置: {file_list_height} (总高度: {total_height})")
                        return
            except Exception as e:
                print(f"[DEBUG] 设置 result_paned 分割条失败: {e}")
            
            if retry < 5:
                self.root.after(200, lambda: init_result_paned_sash(retry + 1))
        
        self.root.after(300, init_result_paned_sash)
        
        # 存储批量分析结果
        self.batch_results = []
        self.batch_folder_path = None
        
        # 暂停控制
        self.batch_paused = False
        self.batch_running = False
        self.pause_event = None  # threading.Event 对象
        
        # 使用重试机制设置分割条位置（复用单文件分析的方法）
        # 左侧面板约350像素，右侧区域更大用于显示异常列表
        self._init_sash_position(self.batch_paned, 350, 0)
        # 标记需要重新设置分割条位置（在 Tab 切换时）
        self.batch_sash_initialized = False
    
    def _create_condition_row(self, parent, index, enabled=False):
        """创建单个异常条件设置行
        
        支持多种异常类型：
        - 波动剧烈检测：列选择 + 阈值
        - 数据范围异常：列选择 + 最小值 + 最大值
        - 同增同减检测：多列选择
        - 自定义：多列运算
        
        Args:
            parent: 父容器
            index: 条件索引
            enabled: 是否默认启用
        """
        # 主框架
        row_frame = ttk.Frame(parent)
        row_frame.pack(fill=tk.X, pady=2)
        self.condition_frames.append(row_frame)
        
        # 启用复选框
        enabled_var = tk.IntVar(value=1 if enabled else 0)
        self.condition_enabled.append(enabled_var)
        ttk.Checkbutton(row_frame, text=f"条件{index + 1}", variable=enabled_var).pack(side=tk.LEFT)
        
        # === 异常类型下拉框 ===
        type_var = tk.StringVar(value="波动剧烈检测")
        self.condition_types.append(type_var)
        type_combo = ttk.Combobox(row_frame, textvariable=type_var, state='readonly', width=14)
        type_combo['values'] = ['波动剧烈检测', '同增同减检测', '数据范围异常', '多条件组合', '自定义']
        type_combo.pack(side=tk.LEFT, padx=2)
        type_combo.bind('<<ComboboxSelected>>', lambda e: self._on_condition_type_change(index))
        
        # === 参数区域（根据类型动态变化）===
        params_frame = ttk.Frame(row_frame)
        params_frame.pack(side=tk.LEFT, padx=2)
        self.params_frames.append(params_frame)
        
        # 初始化所有可能需要的变量（支持最多5列，用于同增同减检测）
        col1_var = tk.StringVar()
        col2_var = tk.StringVar()
        col3_var = tk.StringVar()
        col4_var = tk.StringVar()
        col5_var = tk.StringVar()
        op1_var = tk.StringVar(value="+")
        op2_var = tk.StringVar(value="+")
        compare_var = tk.StringVar(value=">")
        threshold_var = tk.StringVar(value="30")
        min_var = tk.StringVar(value="0")
        max_var = tk.StringVar(value="10000")
        
        # 存储所有变量
        self.condition_columns.append([col1_var, col2_var, col3_var, col4_var, col5_var])
        self.condition_operators.append([op1_var, op2_var])
        self.condition_compare_type.append(compare_var)
        self.condition_threshold.append(threshold_var)
        self.condition_min_max.append([min_var, max_var])
        
        # 存储下拉框引用（延迟初始化，支持最多5列）
        self.condition_column_combos.append([None, None, None, None, None])
        self.operator_combos.append([None, None])
        
        # 存储自定义公式数据
        if not hasattr(self, 'custom_formulas'):
            self.custom_formulas = []
        while len(self.custom_formulas) <= index:
            self.custom_formulas.append(None)
        
        # 删除按钮
        delete_btn = ttk.Button(row_frame, text="✕", width=3, 
                                command=lambda: self._remove_condition_row(index))
        delete_btn.pack(side=tk.RIGHT, padx=2)
        
        # 更新条件计数
        self.condition_count = len(self.condition_frames)
        
        # 创建默认的波动剧烈参数界面
        self._create_params_for_type(index, "波动剧烈")
        
        # 更新所有条件的编号
        self._update_condition_numbers()
    
    def _auto_complete_column(self, event, combo, var):
        """列名自动补全
        
        支持多种匹配方式：
        1. 前缀匹配：输入"今日"匹配"今日并网量"
        2. 包含匹配：输入"20570"匹配"直流PV总功率(20570)"
        3. 括号内数字匹配：输入"20570"优先匹配"(20570)"
        
        注意：不会自动锁定到某个值，而是弹出下拉列表让用户选择
        """
        # 获取用户当前输入
        current = combo.get()
        
        # 如果用户按了方向键或删除键，不处理
        if event and event.keysym in ('Down', 'Up', 'Left', 'Right', 'BackSpace', 'Delete'):
            if event.keysym == 'Down':
                # 按下键时弹出下拉列表
                combo.event_generate('<Down>')
            return
        
        columns = self.current_columns if hasattr(self, 'current_columns') and self.current_columns else []
        
        if not columns:
            return
        
        if not current:
            # 输入为空时，显示所有列
            combo['values'] = columns
            return
        
        # 收集所有匹配结果（按优先级排序）
        all_matches = []
        seen = set()
        
        # 1. 括号内数字匹配优先（如输入"20570"匹配"(20570)"）
        if current.isdigit():
            bracket_pattern = f"({current})"
            for col in columns:
                if bracket_pattern in col and col not in seen:
                    all_matches.append(col)
                    seen.add(col)
        
        # 2. 前缀匹配
        for col in columns:
            if col.lower().startswith(current.lower()) and col not in seen:
                all_matches.append(col)
                seen.add(col)
        
        # 3. 包含匹配
        for col in columns:
            if current.lower() in col.lower() and col not in seen:
                all_matches.append(col)
                seen.add(col)
        
        # 更新下拉框的可选值
        if all_matches:
            combo['values'] = all_matches
            # 关键：恢复用户输入，不让 Combobox 自动选中第一个
            combo.delete(0, tk.END)
            combo.insert(0, current)
            # 将光标移到末尾
            combo.icursor(tk.END)
            
            # 如果只有一个匹配项，可以显示在列表中但不自动选中
            # 弹出下拉列表让用户选择
            if len(all_matches) > 1:
                try:
                    combo.event_generate('<Down>')
                except:
                    pass
        else:
            # 没有匹配项，保持用户输入
            combo['values'] = []
    
    def _create_search_combo(self, parent, var, columns, width=20):
        """创建带搜索功能的下拉框
        
        使用 Entry + 下拉按钮 + 下拉 Frame 的方式
        避免 ttk.Combobox 的自动补全行为
        下拉列表在主窗口内显示，不会有遮挡问题
        """
        # 主容器
        container = ttk.Frame(parent)
        
        # 输入行（输入框 + 下拉按钮）
        input_frame = ttk.Frame(container)
        input_frame.pack(side=tk.TOP, fill=tk.X)
        
        # 输入框
        entry = ttk.Entry(input_frame, width=width)
        entry.pack(side=tk.LEFT)
        
        # 下拉按钮
        dropdown_btn = ttk.Button(input_frame, text="▼", width=2)
        dropdown_btn.pack(side=tk.LEFT)
        
        # 下拉列表 Frame（初始隐藏）
        dropdown_frame = ttk.Frame(container)
        # 不 pack，需要时再 pack
        
        # Listbox 和滚动条
        listbox = tk.Listbox(dropdown_frame, width=width + 2, height=1)  # 初始高度为1
        scrollbar = ttk.Scrollbar(dropdown_frame, orient=tk.VERTICAL, command=listbox.yview)
        listbox.config(yscrollcommand=scrollbar.set)
        
        listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 如果已有值，显示
        if var.get():
            entry.insert(0, var.get())
        
        is_dropdown_visible = [False]  # 使用列表跟踪状态
        
        def show_dropdown(matched):
            """显示下拉列表"""
            if not matched:
                return
            
            # 更新列表内容
            listbox.delete(0, tk.END)
            for col in matched:
                listbox.insert(tk.END, col)
            
            # 选中第一个
            if matched:
                listbox.selection_set(0)
            
            # 设置高度
            listbox.config(height=min(10, len(matched)))
            
            # 显示下拉框
            if not is_dropdown_visible[0]:
                dropdown_frame.pack(side=tk.TOP, fill=tk.X, after=input_frame)
                is_dropdown_visible[0] = True
        
        def hide_dropdown():
            """隐藏下拉列表"""
            if is_dropdown_visible[0]:
                dropdown_frame.pack_forget()
                is_dropdown_visible[0] = False
        
        def toggle_dropdown():
            """切换下拉列表显示/隐藏"""
            if is_dropdown_visible[0]:
                hide_dropdown()
            else:
                # 【优化】点击下拉按钮时显示所有选项，不根据当前输入过滤
                # 这样用户可以直接选择其他选项，无需先清除输入框
                cols = getattr(container, '_columns', columns)
                if cols:
                    show_dropdown(cols)
        
        def on_select(event=None):
            selection = listbox.curselection()
            if selection:
                selected = listbox.get(selection[0])
                entry.delete(0, tk.END)
                entry.insert(0, selected)
                var.set(selected)
                hide_dropdown()
        
        def on_listbox_click(event):
            """处理鼠标点击事件，在鼠标释放时获取选中的项"""
            # 获取鼠标点击位置对应的索引
            index = listbox.nearest(event.y)
            if 0 <= index < listbox.size():
                selected = listbox.get(index)
                entry.delete(0, tk.END)
                entry.insert(0, selected)
                var.set(selected)
                hide_dropdown()
        
        def on_key_release(event):
            # 获取当前列名列表（支持动态更新）
            cols = getattr(container, '_columns', columns)
            
            # 处理特殊键
            if event.keysym == 'Down':
                if not is_dropdown_visible[0]:
                    current_text = entry.get().lower()
                    matched = [col for col in cols if current_text in col.lower()] if current_text else cols
                    show_dropdown(matched)
                # 移动焦点到列表
                if is_dropdown_visible[0]:
                    listbox.focus_set()
                    if listbox.size() > 0:
                        listbox.selection_set(0)
                return
            
            if event.keysym == 'Escape':
                hide_dropdown()
                return
            
            if event.keysym == 'Return':
                if is_dropdown_visible[0]:
                    on_select()
                return
            
            # 输入时自动弹出下拉列表
            current_text = entry.get().lower()
            if current_text:
                matched = [col for col in cols if current_text in col.lower()]
            else:
                matched = cols
            
            if matched:
                show_dropdown(matched)
            else:
                hide_dropdown()
            
            # 同步到 var
            var.set(entry.get())
        
        def on_entry_change(*args):
            # 检查 entry 是否还存在（避免组件已销毁时访问）
            try:
                if not entry.winfo_exists():
                    return
                current = var.get()
                if current != entry.get():
                    entry.delete(0, tk.END)
                    entry.insert(0, current)
            except tk.TclError:
                # 组件已被销毁，忽略此回调
                pass
        
        # 绑定事件
        entry.bind('<KeyRelease>', on_key_release)
        entry.bind('<Down>', lambda e: on_key_release(e))
        dropdown_btn.config(command=toggle_dropdown)
        
        # 【修复】使用 <ButtonRelease-1> 事件处理鼠标点击选择，避免选中项错误的问题
        listbox.bind('<ButtonRelease-1>', on_listbox_click)
        listbox.bind('<Return>', on_select)
        listbox.bind('<Escape>', lambda e: hide_dropdown())
        
        # 监听 var 变化
        var.trace_add('write', on_entry_change)
        
        # 保存引用
        container.entry = entry
        container.var = var
        container.toggle_dropdown = toggle_dropdown
        container.hide_dropdown = hide_dropdown
        container.show_dropdown = show_dropdown
        container.listbox = listbox  # 保存 listbox 引用
        
        # 保存原始列名列表，用于更新
        container._columns = columns
        
        def update_columns(new_columns):
            """更新下拉框的选项列表"""
            container._columns = new_columns
            # 如果当前有输入，过滤后更新
            current_text = entry.get().lower()
            if current_text:
                matched = [col for col in new_columns if current_text in col.lower()]
            else:
                matched = new_columns
            # 如果下拉框可见，更新显示
            if is_dropdown_visible[0] and matched:
                listbox.delete(0, tk.END)
                for col in matched:
                    listbox.insert(tk.END, col)
                listbox.config(height=min(10, len(matched)))
        
        container.update_columns = update_columns
        
        return container, entry
    
    def _create_params_for_type(self, index, condition_type):
        """根据条件类型创建参数界面"""
        params_frame = self.params_frames[index]
        col1_var, col2_var, col3_var, col4_var, col5_var = self.condition_columns[index]
        op1_var, op2_var = self.condition_operators[index]
        compare_var = self.condition_compare_type[index]
        threshold_var = self.condition_threshold[index]
        min_var, max_var = self.condition_min_max[index]
        
        # 清空现有控件
        for widget in params_frame.winfo_children():
            widget.destroy()
        
        # 兼容旧名称：转换为新名称
        old_to_new = {
            '波动剧烈': '波动剧烈检测',
            '超出范围': '数据范围异常',
            '同增同减异常': '同增同减检测'
        }
        condition_type = old_to_new.get(condition_type, condition_type)
        
        # 获取当前列名（过滤掉时间列，因为时间列是第一列，用于排序和索引，不需要在条件中选择）
        columns = self.current_columns if hasattr(self, 'current_columns') and self.current_columns else []
        columns = self._filter_time_columns(columns)  # 过滤掉时间列
        
        if condition_type == '波动剧烈检测':
            # 列选择 + 阈值（使用搜索组件）
            combo_frame, col1_entry = self._create_search_combo(params_frame, col1_var, columns, width=20)
            combo_frame.pack(side=tk.LEFT, padx=1)
            # 保存入口引用
            if not hasattr(self, 'search_combo_entries'):
                self.search_combo_entries = {}
            self.search_combo_entries[(index, 0)] = col1_entry
            self.condition_column_combos[index][0] = col1_entry
            
            ttk.Label(params_frame, text="阈值(%):").pack(side=tk.LEFT)
            ttk.Entry(params_frame, textvariable=threshold_var, width=6).pack(side=tk.LEFT)
            threshold_var.set("30")
            
        elif condition_type == '数据范围异常':
            # 列选择 + 最小值 + 最大值（使用搜索组件）
            combo_frame, col1_entry = self._create_search_combo(params_frame, col1_var, columns, width=20)
            combo_frame.pack(side=tk.LEFT, padx=1)
            if not hasattr(self, 'search_combo_entries'):
                self.search_combo_entries = {}
            self.search_combo_entries[(index, 0)] = col1_entry
            self.condition_column_combos[index][0] = col1_entry
            
            ttk.Label(params_frame, text="最小:").pack(side=tk.LEFT)
            ttk.Entry(params_frame, textvariable=min_var, width=6).pack(side=tk.LEFT)
            ttk.Label(params_frame, text="最大:").pack(side=tk.LEFT)
            ttk.Entry(params_frame, textvariable=max_var, width=6).pack(side=tk.LEFT)
            
        elif condition_type == '同增同减检测':
            # 支持最多5列一起判断同增同减
            # 创建一个内部 Frame 来容纳所有列选择
            cols_frame = ttk.Frame(params_frame)
            cols_frame.pack(side=tk.LEFT, fill=tk.X)
            
            # 存储当前显示的列数
            if not hasattr(self, 'sync_cols_count'):
                self.sync_cols_count = [2, 2, 2]  # 默认每个条件显示2列
            if len(self.sync_cols_count) <= index:
                self.sync_cols_count.append(2)
            
            # 列变量
            col_vars = [col1_var, col2_var, col3_var, col4_var, col5_var]
            
            # 创建列选择下拉框的函数（使用搜索组件）
            def create_col_combo(col_idx, col_var):
                frame = ttk.Frame(cols_frame)
                frame.pack(side=tk.LEFT, padx=2)
                
                ttk.Label(frame, text=f"列{col_idx + 1}:").pack(side=tk.LEFT)
                combo_frame, combo_entry = self._create_search_combo(frame, col_var, columns, width=18)
                combo_frame.pack(side=tk.LEFT)
                if not hasattr(self, 'search_combo_entries'):
                    self.search_combo_entries = {}
                self.search_combo_entries[(index, col_idx)] = combo_entry
                self.condition_column_combos[index][col_idx] = combo_entry
                return frame
            
            # 创建前两列（必须）
            create_col_combo(0, col_vars[0])
            create_col_combo(1, col_vars[1])
            
            # 存储额外的列框架引用
            if not hasattr(self, 'extra_col_frames'):
                self.extra_col_frames = [[], [], []]
            if len(self.extra_col_frames) <= index:
                self.extra_col_frames.append([])
            self.extra_col_frames[index] = []
            
            # 添加/删除列的按钮
            btn_frame = ttk.Frame(cols_frame)
            btn_frame.pack(side=tk.LEFT, padx=5)
            
            def add_column():
                if self.sync_cols_count[index] < 5:
                    col_idx = self.sync_cols_count[index]
                    frame = create_col_combo(col_idx, col_vars[col_idx])
                    self.extra_col_frames[index].append(frame)
                    self.sync_cols_count[index] += 1
                    # 更新按钮状态
                    if self.sync_cols_count[index] >= 5:
                        add_btn.config(state='disabled')
                    remove_btn.config(state='normal')
            
            def remove_column():
                if self.sync_cols_count[index] > 2:
                    # 销毁最后一个额外的列框架
                    if self.extra_col_frames[index]:
                        self.extra_col_frames[index].pop().destroy()
                    self.sync_cols_count[index] -= 1
                    # 清空对应的变量
                    col_vars[self.sync_cols_count[index]].set('')
                    # 更新按钮状态
                    if self.sync_cols_count[index] <= 2:
                        remove_btn.config(state='disabled')
                    add_btn.config(state='normal')
            
            add_btn = ttk.Button(btn_frame, text="+", width=2, command=add_column)
            add_btn.pack(side=tk.LEFT, padx=1)
            remove_btn = ttk.Button(btn_frame, text="-", width=2, command=remove_column, state='disabled')
            remove_btn.pack(side=tk.LEFT, padx=1)
            
        elif condition_type == '多条件组合':
            # 多条件组合检测：选择一列 + 多个条件 + 逻辑关系（AND/OR）
            # 格式：列名 + (条件1 AND 条件2 OR 条件3...)
            
            # 存储多条件组合的数据结构
            if not hasattr(self, 'multi_condition_data'):
                self.multi_condition_data = {}
            
            # 初始化该条件的数据
            if index not in self.multi_condition_data:
                self.multi_condition_data[index] = {
                    'column': col1_var,
                    'logic': tk.StringVar(value='AND'),  # AND 或 OR
                    'conditions': []  # 每个条件是 {'op': '!=', 'value': '1'}
                }
            
            mc_data = self.multi_condition_data[index]
            
            # 第一行：列选择 + 逻辑关系
            row1 = ttk.Frame(params_frame)
            row1.pack(fill=tk.X, pady=2)
            
            ttk.Label(row1, text="检测列:").pack(side=tk.LEFT)
            combo_frame, col1_entry = self._create_search_combo(row1, col1_var, columns, width=20)
            combo_frame.pack(side=tk.LEFT, padx=2)
            if not hasattr(self, 'search_combo_entries'):
                self.search_combo_entries = {}
            self.search_combo_entries[(index, 0)] = col1_entry
            self.condition_column_combos[index][0] = col1_entry
            
            ttk.Label(row1, text="逻辑:").pack(side=tk.LEFT, padx=(10, 0))
            logic_combo = ttk.Combobox(row1, textvariable=mc_data['logic'], state='readonly', width=6)
            logic_combo['values'] = ['AND', 'OR']
            logic_combo.pack(side=tk.LEFT, padx=2)
            
            # 存储条件框架引用
            if not hasattr(self, 'multi_cond_frames'):
                self.multi_cond_frames = {}
            self.multi_cond_frames[index] = []
            
            # 第二行：条件列表（使用 Frame 容器）
            cond_container = ttk.Frame(params_frame)
            cond_container.pack(fill=tk.X, pady=2)
            self.multi_cond_containers = getattr(self, 'multi_cond_containers', {})
            self.multi_cond_containers[index] = cond_container
            
            # 添加条件按钮
            btn_row = ttk.Frame(params_frame)
            btn_row.pack(fill=tk.X, pady=2)
            
            def add_condition_row():
                """添加一个条件行"""
                cond_idx = len(mc_data['conditions'])
                if cond_idx >= 10:  # 最多10个条件
                    return
                
                # 创建条件变量
                cond_op = tk.StringVar(value='!=')
                cond_val = tk.StringVar()
                mc_data['conditions'].append({'op': cond_op, 'value': cond_val})
                
                # 创建条件行
                cond_row = ttk.Frame(cond_container)
                cond_row.pack(fill=tk.X, pady=1)
                self.multi_cond_frames[index].append(cond_row)
                
                if cond_idx > 0:
                    ttk.Label(cond_row, text=mc_data['logic'].get(), foreground='gray').pack(side=tk.LEFT, padx=2)
                
                ttk.Label(cond_row, text="条件:").pack(side=tk.LEFT)
                op_combo = ttk.Combobox(cond_row, textvariable=cond_op, state='readonly', width=4)
                op_combo['values'] = ['=', '!=', '>', '<', '>=', '<=']
                op_combo.pack(side=tk.LEFT, padx=2)
                
                ttk.Entry(cond_row, textvariable=cond_val, width=10).pack(side=tk.LEFT, padx=2)
                
                # 删除按钮
                ttk.Button(cond_row, text="✕", width=2,
                          command=lambda: remove_condition_row(cond_idx)).pack(side=tk.LEFT, padx=2)
            
            def remove_condition_row(cond_idx):
                """删除指定条件行"""
                if len(mc_data['conditions']) > 1:
                    # 删除对应的数据和UI
                    mc_data['conditions'].pop(cond_idx)
                    if cond_idx < len(self.multi_cond_frames[index]):
                        self.multi_cond_frames[index][cond_idx].destroy()
                        self.multi_cond_frames[index].pop(cond_idx)
                    # 重建所有条件行（因为索引会变化）
                    for frame in self.multi_cond_frames[index]:
                        frame.destroy()
                    self.multi_cond_frames[index] = []
                    mc_data['conditions'].clear()
                    # 重新添加初始条件
                    add_condition_row()
            
            ttk.Button(btn_row, text="+ 添加条件", command=add_condition_row).pack(side=tk.LEFT, padx=2)
            
            # 默认添加一个条件
            add_condition_row()
            
        elif condition_type == '自定义':
            # 显示当前公式预览 + 编辑按钮
            formula_text = "点击设置公式..."
            if self.custom_formulas[index]:
                formula_data = self.custom_formulas[index]
                formula_text = self._format_formula_preview(formula_data)
            
            self.custom_formula_labels = getattr(self, 'custom_formula_labels', [None, None, None])
            if len(self.custom_formula_labels) <= index:
                self.custom_formula_labels.append(None)
            
            label = ttk.Label(params_frame, text=formula_text, foreground='blue')
            label.pack(side=tk.LEFT, padx=5)
            self.custom_formula_labels[index] = label
            
            btn = ttk.Button(params_frame, text="设置公式", command=lambda i=index: self._open_custom_formula_dialog(i))
            btn.pack(side=tk.LEFT, padx=5)
    
    def _format_formula_preview(self, formula_data):
        """格式化公式预览文本（支持新旧格式）"""
        compare_op = formula_data.get('compare', '>')
        threshold = formula_data.get('threshold', '0')
        
        # 优先使用新格式
        elements = formula_data.get('elements', [])
        if elements:
            parts = []
            for elem in elements:
                if elem['type'] == 'column':
                    parts.append(f"[{elem['value'][:10]}]")
                elif elem['type'] == 'operator':
                    parts.append(f" {elem['value']} ")
                elif elem['type'] == 'constant':
                    parts.append(str(elem['value']))
                elif elem['type'] == 'parenthesis':
                    parts.append(elem['value'])
            text = ''.join(parts) + f" {compare_op} {threshold}"
        else:
            # 使用旧格式
            items = formula_data.get('items', [])
            text_parts = []
            for item in items:
                op = item.get('operator', '+')
                col = item['column']
                text_parts.append(f"{op} {col}")
            text = " ".join(text_parts) + f" {compare_op} {threshold}"
        
        return text[:35] + "..." if len(text) > 35 else text
    
    def _open_custom_formula_dialog(self, index):
        """打开自定义公式设置弹窗"""
        dialog = tk.Toplevel(self.root)
        dialog.title("自定义计算公式")
        dialog.geometry("650x550")
        dialog.transient(self.root)
        dialog.grab_set()
        
        # 获取当前列名
        columns = self.current_columns if hasattr(self, 'current_columns') and self.current_columns else []
        
        # 公式元素列表：每个元素是 {'type': 'column'/'operator'/'constant', 'value': '...'}
        formula_elements = []
        compare_op = tk.StringVar(value='>')
        threshold = tk.StringVar(value='0')
        
        # 主框架
        main_frame = ttk.Frame(dialog, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # ========== 计算公式区域 ==========
        formula_frame = ttk.LabelFrame(main_frame, text="计算公式", padding="5")
        formula_frame.pack(fill=tk.X, pady=5)
        
        formula_text = tk.Text(formula_frame, height=2, wrap=tk.WORD, font=('Consolas', 11))
        formula_text.pack(fill=tk.X, padx=5, pady=5)
        formula_text.config(state=tk.DISABLED)
        
        # ========== 公式预览区域 ==========
        preview_frame = ttk.LabelFrame(main_frame, text="公式预览", padding="5")
        preview_frame.pack(fill=tk.X, pady=5)
        
        preview_label = ttk.Label(preview_frame, text="等待输入...", foreground='blue', font=('Consolas', 10))
        preview_label.pack(anchor=tk.W, padx=5, pady=5)
        
        def update_formula_display():
            """更新公式显示和预览"""
            try:
                # 检查组件是否存在
                if not formula_text.winfo_exists() or not preview_label.winfo_exists():
                    return
                
                # 更新公式文本
                formula_text.config(state=tk.NORMAL)
                formula_text.delete(1.0, tk.END)
                
                display_parts = []
                for elem in formula_elements:
                    if elem['type'] == 'column':
                        display_parts.append(f"[{elem['value']}]")
                    elif elem['type'] == 'operator':
                        # 逻辑运算符用不同颜色显示
                        if elem['value'] in ['AND', 'OR', 'NOT']:
                            display_parts.append(f" {elem['value']} ")
                        else:
                            display_parts.append(f" {elem['value']} ")
                    elif elem['type'] == 'constant':
                        display_parts.append(str(elem['value']))
                    elif elem['type'] == 'parenthesis':
                        display_parts.append(elem['value'])
                
                formula_text.insert(tk.END, ''.join(display_parts))
                formula_text.config(state=tk.DISABLED)
                
                # 更新预览
                preview_parts = []
                for elem in formula_elements:
                    if elem['type'] == 'column':
                        preview_parts.append(f"[{elem['value']}]")
                    elif elem['type'] == 'operator':
                        preview_parts.append(f" {elem['value']} ")
                    elif elem['type'] == 'constant':
                        preview_parts.append(str(elem['value']))
                    elif elem['type'] == 'parenthesis':
                        preview_parts.append(elem['value'])
                
                preview_expr = ''.join(preview_parts)
                preview_text = f"{preview_expr} {compare_op.get()} {threshold.get()}" if preview_expr else "等待输入..."
                preview_label.config(text=preview_text)
            except tk.TclError:
                # 组件已被销毁，忽略此回调
                pass
        
        # ========== 可用数据列区域 ==========
        columns_frame = ttk.LabelFrame(main_frame, text="可用数据列", padding="5")
        columns_frame.pack(fill=tk.X, pady=5)
        
        # 搜索框
        search_frame = ttk.Frame(columns_frame)
        search_frame.pack(fill=tk.X, pady=2)
        
        ttk.Label(search_frame, text="搜索:").pack(side=tk.LEFT)
        search_var = tk.StringVar()
        search_entry = ttk.Entry(search_frame, textvariable=search_var, width=25)
        search_entry.pack(side=tk.LEFT, padx=5)
        
        # 列选择下拉框
        col_select_frame = ttk.Frame(columns_frame)
        col_select_frame.pack(fill=tk.X, pady=2)
        
        col_var = tk.StringVar()
        col_combo = ttk.Combobox(col_select_frame, textvariable=col_var, width=35)
        col_combo['values'] = columns
        col_combo.pack(side=tk.LEFT, padx=5)
        
        # 运算符按钮组
        btn_frame = ttk.Frame(columns_frame)
        btn_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(btn_frame, text="运算符:").pack(side=tk.LEFT, padx=2)
        
        # 算术运算符按钮
        operators = ['+', '-', '*', '/', '(', ')']
        for op in operators:
            btn = ttk.Button(btn_frame, text=op, width=3,
                            command=lambda o=op: add_operator(o))
            btn.pack(side=tk.LEFT, padx=2)
        
        # 逻辑运算符按钮（新增）
        ttk.Label(btn_frame, text="逻辑:").pack(side=tk.LEFT, padx=(10, 2))
        logic_operators = ['AND', 'OR', 'NOT']
        for op in logic_operators:
            btn = ttk.Button(btn_frame, text=op, width=4,
                            command=lambda o=op: add_operator(o))
            btn.pack(side=tk.LEFT, padx=2)
        
        # 添加列按钮
        ttk.Button(btn_frame, text="添加列", width=8, 
                   command=lambda: add_column()).pack(side=tk.LEFT, padx=10)
        
        # 删除最后一个元素按钮
        ttk.Button(btn_frame, text="退格", width=6,
                   command=lambda: remove_last()).pack(side=tk.LEFT, padx=5)
        
        # 清空按钮
        ttk.Button(btn_frame, text="清空", width=6,
                   command=lambda: clear_formula()).pack(side=tk.LEFT, padx=5)
        
        def add_column():
            """添加选中的列到公式"""
            col = col_var.get()
            if col:
                formula_elements.append({'type': 'column', 'value': col})
                update_formula_display()
        
        def add_operator(op):
            """添加运算符到公式"""
            formula_elements.append({'type': 'operator', 'value': op})
            update_formula_display()
        
        def remove_last():
            """删除最后一个元素"""
            if formula_elements:
                formula_elements.pop()
                update_formula_display()
        
        def clear_formula():
            """清空公式"""
            formula_elements.clear()
            update_formula_display()
        
        # 搜索过滤
        def filter_columns(*args):
            try:
                # 检查组件是否存在
                if not col_combo.winfo_exists():
                    return
                keyword = search_var.get().lower()
                if keyword:
                    filtered = [col for col in columns if keyword in col.lower()]
                else:
                    filtered = columns
                col_combo['values'] = filtered
                # 不再自动选中，让用户自己选择
            except tk.TclError:
                # 组件已被销毁，忽略此回调
                pass
        
        search_var.trace_add('write', filter_columns)
        col_combo.bind('<KeyRelease>', lambda e: filter_columns())
        
        # ========== 常量输入区域 ==========
        const_frame = ttk.LabelFrame(main_frame, text="常量输入", padding="5")
        const_frame.pack(fill=tk.X, pady=5)
        
        const_input_frame = ttk.Frame(const_frame)
        const_input_frame.pack(fill=tk.X)
        
        ttk.Label(const_input_frame, text="数值:").pack(side=tk.LEFT)
        const_var = tk.StringVar()
        const_entry = ttk.Entry(const_input_frame, textvariable=const_var, width=15)
        const_entry.pack(side=tk.LEFT, padx=5)
        
        ttk.Button(const_input_frame, text="添加常量", width=10,
                   command=lambda: add_constant()).pack(side=tk.LEFT, padx=10)
        
        def add_constant():
            """添加常量到公式"""
            val = const_var.get()
            if val:
                try:
                    # 验证是否为有效数字
                    float(val)
                    formula_elements.append({'type': 'constant', 'value': val})
                    const_var.set('')
                    update_formula_display()
                except ValueError:
                    messagebox.showwarning("警告", "请输入有效的数字！")
        
        # ========== 比较条件区域 ==========
        compare_frame = ttk.LabelFrame(main_frame, text="比较条件", padding="5")
        compare_frame.pack(fill=tk.X, pady=5)
        
        compare_input_frame = ttk.Frame(compare_frame)
        compare_input_frame.pack(fill=tk.X)
        
        ttk.Label(compare_input_frame, text="计算结果").pack(side=tk.LEFT)
        compare_combo = ttk.Combobox(compare_input_frame, textvariable=compare_op, state='readonly', width=5)
        compare_combo['values'] = ['>', '<', '=', '>=', '<=', '!=']
        compare_combo.pack(side=tk.LEFT, padx=5)
        
        ttk.Label(compare_input_frame, text="阈值:").pack(side=tk.LEFT)
        threshold_entry = ttk.Entry(compare_input_frame, textvariable=threshold, width=12)
        threshold_entry.pack(side=tk.LEFT, padx=5)
        
        # 绑定阈值变化更新预览
        threshold.trace_add('write', lambda *args: update_formula_display())
        compare_op.trace_add('write', lambda *args: update_formula_display())
        
        # ========== 底部按钮区域 ==========
        bottom_frame = ttk.Frame(main_frame)
        bottom_frame.pack(fill=tk.X, pady=10)
        
        def validate_formula():
            """验证公式是否合法"""
            if not formula_elements:
                messagebox.showwarning("警告", "公式为空！")
                return False
            
            # 检查列名是否存在
            for elem in formula_elements:
                if elem['type'] == 'column' and elem['value'] not in columns:
                    messagebox.showwarning("警告", f"列 '{elem['value']}' 不存在！")
                    return False
            
            # 检查括号是否匹配
            paren_count = 0
            for elem in formula_elements:
                if elem['type'] == 'parenthesis':
                    if elem['value'] == '(':
                        paren_count += 1
                    else:
                        paren_count -= 1
                    if paren_count < 0:
                        messagebox.showwarning("警告", "括号不匹配！")
                        return False
            
            if paren_count != 0:
                messagebox.showwarning("警告", "括号不匹配！")
                return False
            
            # 检查阈值是否有效
            try:
                float(threshold.get())
            except ValueError:
                messagebox.showwarning("警告", "请输入有效的阈值！")
                return False
            
            messagebox.showinfo("成功", "公式验证通过！")
            return True
        
        def on_confirm():
            """确认按钮"""
            if not formula_elements:
                messagebox.showwarning("警告", "请构建计算公式！")
                return
            
            # 验证列名
            for elem in formula_elements:
                if elem['type'] == 'column' and elem['value'] not in columns:
                    messagebox.showwarning("警告", f"列 '{elem['value']}' 不存在！")
                    return
            
            # 保存公式数据
            self.custom_formulas[index] = {
                'elements': formula_elements.copy(),
                'compare': compare_op.get(),
                'threshold': threshold.get()
            }
            
            # 更新显示
            if hasattr(self, 'custom_formula_labels') and self.custom_formula_labels[index]:
                self.custom_formula_labels[index].config(
                    text=self._format_formula_preview(self.custom_formulas[index]))
            
            dialog.destroy()
        
        def on_cancel():
            dialog.destroy()
        
        ttk.Button(bottom_frame, text="验证公式", command=validate_formula, width=10).pack(side=tk.LEFT, padx=10)
        ttk.Button(bottom_frame, text="确定", command=on_confirm, width=10).pack(side=tk.LEFT, padx=10)
        ttk.Button(bottom_frame, text="取消", command=on_cancel, width=10).pack(side=tk.LEFT, padx=10)
        
        # 如果已有公式，加载它
        if hasattr(self, 'custom_formulas') and self.custom_formulas[index]:
            data = self.custom_formulas[index]
            # 兼容旧格式
            if 'elements' in data:
                formula_elements.extend(data['elements'])
            elif 'items' in data:
                # 转换旧格式到新格式
                for item in data['items']:
                    op = item.get('operator', '+')
                    col = item['column']
                    formula_elements.append({'type': 'operator', 'value': op})
                    formula_elements.append({'type': 'column', 'value': col})
            compare_op.set(data.get('compare', '>'))
            threshold.set(data.get('threshold', '0'))
            update_formula_display()
    
    def _add_key_value_row(self):
        """添加关键值追踪行"""
        # 限制最多5个关键值
        if len(self.key_value_vars) >= 5:
            messagebox.showinfo("提示", "最多支持追踪5个关键值")
            return
        
        # 创建行框架
        row_frame = ttk.Frame(self.key_values_container)
        row_frame.pack(fill=tk.X, pady=1)
        self.key_value_frames.append(row_frame)
        
        # 创建下拉框变量
        var = tk.StringVar()
        self.key_value_vars.append(var)
        
        # 获取可用的列名（已过滤时间列）
        columns = []
        if hasattr(self, 'batch_available_columns') and self.batch_available_columns:
            columns = self.batch_available_columns
        elif hasattr(self, 'current_columns') and self.current_columns:
            columns = self._filter_time_columns(self.current_columns)
        
        # 【修复】强制再次过滤时间列，确保没有遗漏
        if columns:
            columns = self._filter_time_columns(columns)
        
        # 创建带搜索功能的下拉框
        combo_container, entry = self._create_search_combo(row_frame, var, columns, width=32)
        combo_container.pack(side=tk.LEFT, padx=2)
        self.key_value_combos.append(combo_container)  # 保存容器引用
        
        # 添加删除按钮
        ttk.Button(row_frame, text="×", width=3, 
                   command=lambda: self._remove_key_value_row(len(self.key_value_vars) - 1)).pack(side=tk.LEFT, padx=2)
    
    def _remove_key_value_row(self, index):
        """删除关键值追踪行"""
        if 0 <= index < len(self.key_value_frames):
            # 销毁框架
            self.key_value_frames[index].destroy()
            # 从列表中移除
            self.key_value_frames.pop(index)
            self.key_value_vars.pop(index)
            self.key_value_combos.pop(index)
            
            # 重新绑定删除按钮的命令（索引需要更新）
            for i, frame in enumerate(self.key_value_frames):
                for widget in frame.winfo_children():
                    if isinstance(widget, ttk.Button):
                        widget.config(command=lambda idx=i: self._remove_key_value_row(idx))
                        break
    
    def _clear_key_values(self):
        """清空所有关键值追踪设置"""
        # 销毁所有行框架
        for frame in self.key_value_frames:
            frame.destroy()
        # 清空列表
        self.key_value_frames.clear()
        self.key_value_vars.clear()
        self.key_value_combos.clear()
    
    def _get_selected_key_columns(self):
        """获取用户选择的关键值列名列表"""
        result = []
        for var in self.key_value_vars:
            col_name = var.get().strip()
            # 过滤掉时间列，避免出现在关键值中
            # 时间列的特征：包含"时间"或"time"关键词
            if col_name and not any(keyword in col_name.lower() for keyword in ['时间', 'time', 'timestamp']):
                result.append(col_name)
        return result
    
    def _add_condition_row(self):
        """添加新的条件行"""
        # 获取下一个索引
        index = len(self.condition_frames)
        self._create_condition_row(self.condition_list_frame, index, enabled=True)
        print(f"[INFO] 添加条件 {index + 1}")
    
    def _remove_condition_row(self, index):
        """删除指定索引的条件行
        
        Args:
            index: 要删除的条件索引
        """
        if len(self.condition_frames) <= 1:
            messagebox.showwarning("提示", "至少需要保留一个条件！")
            return
        
        # 删除 UI 组件
        self.condition_frames[index].destroy()
        
        # 删除存储的数据
        del self.condition_frames[index]
        del self.condition_enabled[index]
        del self.condition_types[index]
        del self.condition_columns[index]
        del self.condition_operators[index]
        del self.condition_compare_type[index]
        del self.condition_threshold[index]
        del self.condition_min_max[index]
        del self.condition_column_combos[index]
        del self.operator_combos[index]
        del self.params_frames[index]
        if index < len(self.custom_formulas):
            del self.custom_formulas[index]
        
        # 更新条件计数
        self.condition_count = len(self.condition_frames)
        
        # 更新所有条件的编号和事件绑定
        self._rebind_condition_events()
        self._update_condition_numbers()
        
        print(f"[INFO] 删除条件 {index + 1}")
    
    def _update_condition_numbers(self):
        """更新所有条件的编号显示"""
        for i, frame in enumerate(self.condition_frames):
            # 查找并更新复选框文本
            for widget in frame.winfo_children():
                if isinstance(widget, ttk.Checkbutton):
                    widget.configure(text=f"条件{i + 1}")
                    break
    
    def _rebind_condition_events(self):
        """重新绑定条件行的事件（删除后索引变化）"""
        for i in range(len(self.condition_frames)):
            # 更新类型下拉框绑定
            for widget in self.condition_frames[i].winfo_children():
                if isinstance(widget, ttk.Combobox) and widget.cget('width') == 10:
                    # 重新绑定事件
                    widget.bind('<<ComboboxSelected>>', lambda e, idx=i: self._on_condition_type_change(idx))
                elif isinstance(widget, ttk.Button) and widget.cget('text') == "✕":
                    # 更新删除按钮命令
                    widget.configure(command=lambda idx=i: self._remove_condition_row(idx))
    
    def _save_batch_conditions_config(self):
        """保存当前条件配置"""
        # 弹出对话框输入配置名称
        dialog = tk.Toplevel(self.root)
        dialog.title("保存配置")
        dialog.geometry("300x120")
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.focus_force()
        
        ttk.Label(dialog, text="配置名称:").pack(pady=10)
        name_var = tk.StringVar()
        name_entry = ttk.Entry(dialog, textvariable=name_var, width=30)
        name_entry.pack(pady=5)
        name_entry.focus_set()
        
        def do_save():
            name = name_var.get().strip()
            if not name:
                messagebox.showwarning("提示", "请输入配置名称！")
                return
            
            # 收集当前条件配置
            config = {
                'name': name,
                'conditions': []
            }
            
            for i in range(len(self.condition_frames)):
                condition = {
                    'enabled': self.condition_enabled[i].get() == 1,
                    'type': self.condition_types[i].get(),
                    'columns': [col.get() for col in self.condition_columns[i]],
                    'operators': [op.get() for op in self.condition_operators[i]],
                    'compare': self.condition_compare_type[i].get(),
                    'threshold': self.condition_threshold[i].get(),
                    'min_max': [self.condition_min_max[i][0].get(), self.condition_min_max[i][1].get()]
                }
                
                # 保存多条件组合数据
                if hasattr(self, 'multi_condition_data') and i in self.multi_condition_data:
                    mc_data = self.multi_condition_data[i]
                    condition['multi_condition'] = {
                        'column': mc_data['column'].get(),
                        'logic': mc_data['logic'].get(),
                        'conditions': [{'op': c['op'].get(), 'value': c['value'].get()} for c in mc_data['conditions']]
                    }
                
                # 保存自定义公式数据
                if hasattr(self, 'custom_formulas') and i < len(self.custom_formulas) and self.custom_formulas[i]:
                    condition['custom_formula'] = self.custom_formulas[i]
                
                config['conditions'].append(condition)
            
            # 加载现有配置文件
            configs = {}
            if os.path.exists(self.batch_conditions_config_file):
                try:
                    with open(self.batch_conditions_config_file, 'r', encoding='utf-8') as f:
                        configs = json.load(f)
                except:
                    configs = {}
            
            # 保存配置
            configs[name] = config
            
            try:
                with open(self.batch_conditions_config_file, 'w', encoding='utf-8') as f:
                    json.dump(configs, f, ensure_ascii=False, indent=2)
                messagebox.showinfo("成功", f"配置 '{name}' 已保存！")
                self._load_batch_config_list()
                self.saved_config_var.set(name)
                dialog.destroy()
            except Exception as e:
                messagebox.showerror("错误", f"保存配置失败: {e}")
        
        ttk.Button(dialog, text="保存", command=do_save).pack(pady=10)
        dialog.bind('<Return>', lambda e: do_save())
    
    def _load_batch_config_list(self):
        """加载已保存的配置列表"""
        configs = []
        if os.path.exists(self.batch_conditions_config_file):
            try:
                with open(self.batch_conditions_config_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    configs = list(data.keys())
            except:
                pass
        
        if hasattr(self, 'saved_config_combo'):
            self.saved_config_combo['values'] = configs
    
    def _load_batch_conditions_config(self, event=None):
        """加载选中的配置"""
        name = self.saved_config_var.get()
        if not name:
            return
        
        if not os.path.exists(self.batch_conditions_config_file):
            messagebox.showwarning("提示", "配置文件不存在！")
            return
        
        try:
            with open(self.batch_conditions_config_file, 'r', encoding='utf-8') as f:
                configs = json.load(f)
            
            if name not in configs:
                messagebox.showwarning("提示", f"配置 '{name}' 不存在！")
                return
            
            config = configs[name]
            conditions = config.get('conditions', [])
            
            # 清空现有条件
            self._clear_all_conditions()
            
            # 旧名称到新名称的映射
            old_to_new = {
                '波动剧烈': '波动剧烈检测',
                '超出范围': '数据范围异常',
                '同增同减异常': '同增同减检测'
            }
            
            # 创建新条件
            for i, cond in enumerate(conditions):
                # 转换旧名称
                cond_type = cond.get('type', '波动剧烈检测')
                cond_type = old_to_new.get(cond_type, cond_type)
                
                self._create_condition_row(self.condition_list_frame, i, enabled=cond.get('enabled', False))
                
                # 设置条件值
                self.condition_types[i].set(cond_type)
                self._create_params_for_type(i, cond_type)
                
                # 设置列值
                cols = cond.get('columns', [])
                for j, col_val in enumerate(cols):
                    if j < len(self.condition_columns[i]) and col_val:
                        self.condition_columns[i][j].set(col_val)
                
                # 设置其他参数
                ops = cond.get('operators', [])
                for j, op_val in enumerate(ops):
                    if j < len(self.condition_operators[i]):
                        self.condition_operators[i][j].set(op_val)
                
                self.condition_compare_type[i].set(cond.get('compare', '>'))
                self.condition_threshold[i].set(cond.get('threshold', '30'))
                
                min_max = cond.get('min_max', ['0', '10000'])
                self.condition_min_max[i][0].set(min_max[0])
                self.condition_min_max[i][1].set(min_max[1])
                
                # 加载多条件组合数据
                if cond_type == '多条件组合' and 'multi_condition' in cond:
                    mc_data = cond['multi_condition']
                    if hasattr(self, 'multi_condition_data') and i in self.multi_condition_data:
                        self.multi_condition_data[i]['column'].set(mc_data.get('column', ''))
                        self.multi_condition_data[i]['logic'].set(mc_data.get('logic', 'AND'))
                        # 清除默认条件并加载保存的条件
                        self.multi_condition_data[i]['conditions'] = []
                        for c in mc_data.get('conditions', []):
                            self.multi_condition_data[i]['conditions'].append({
                                'op': tk.StringVar(value=c.get('op', '!=')),
                                'value': tk.StringVar(value=c.get('value', ''))
                            })
                
                # 加载自定义公式数据
                if cond_type == '自定义' and 'custom_formula' in cond:
                    if hasattr(self, 'custom_formulas') and i < len(self.custom_formulas):
                        self.custom_formulas[i] = cond['custom_formula']
            
            messagebox.showinfo("成功", f"配置 '{name}' 已加载！")
            
        except Exception as e:
            messagebox.showerror("错误", f"加载配置失败: {e}")
    
    def _clear_all_conditions(self):
        """清空所有条件"""
        # 销毁所有条件框架
        for frame in self.condition_frames:
            frame.destroy()
        
        # 清空列表
        self.condition_frames.clear()
        self.condition_enabled.clear()
        self.condition_types.clear()
        self.condition_columns.clear()
        self.condition_operators.clear()
        self.condition_compare_type.clear()
        self.condition_threshold.clear()
        self.condition_min_max.clear()
        self.condition_column_combos.clear()
        self.operator_combos.clear()
        self.params_frames.clear()
        self.custom_formulas.clear()
        self.condition_count = 0
    
    def _delete_batch_conditions_config(self):
        """删除选中的配置"""
        name = self.saved_config_var.get()
        if not name:
            messagebox.showwarning("提示", "请选择要删除的配置！")
            return
        
        if not os.path.exists(self.batch_conditions_config_file):
            return
        
        try:
            with open(self.batch_conditions_config_file, 'r', encoding='utf-8') as f:
                configs = json.load(f)
            
            if name in configs:
                del configs[name]
                
                with open(self.batch_conditions_config_file, 'w', encoding='utf-8') as f:
                    json.dump(configs, f, ensure_ascii=False, indent=2)
                
                messagebox.showinfo("成功", f"配置 '{name}' 已删除！")
                self._load_batch_config_list()
                self.saved_config_var.set('')
        except Exception as e:
            messagebox.showerror("错误", f"删除配置失败: {e}")
    
    def _on_condition_type_change(self, index):
        """条件类型改变时的回调"""
        condition_type = self.condition_types[index].get()
        self._create_params_for_type(index, condition_type)
    
    def _on_condition_toggle(self, index):
        """条件启用/禁用时的回调"""
        pass
    
    def _on_data_type_change(self, event=None):
        """数据类型改变时的回调"""
        data_type = self.data_type_var.get()
        
        # 更新当前列名
        if data_type == 'BATTERY':
            self.current_columns = self.battery_columns if hasattr(self, 'battery_columns') else []
        else:
            self.current_columns = self.pcs_columns if hasattr(self, 'pcs_columns') else []
        
        # 更新关键值追踪下拉框（过滤掉时间列）
        if hasattr(self, 'key_value_combos') and hasattr(self, 'current_columns') and self.current_columns:
            filtered_columns = self._filter_time_columns(self.current_columns)
            # 【修复】同时更新 batch_available_columns，确保添加新行时使用正确的列
            self.batch_available_columns = filtered_columns
            for combo in self.key_value_combos:
                if hasattr(combo, 'update_columns'):
                    combo.update_columns(filtered_columns)
        
        # 更新所有条件的列下拉框
        condition_count = len(self.condition_types) if hasattr(self, 'condition_types') else 0
        for i in range(condition_count):
            # 重新创建参数界面以更新列名
            condition_type = self.condition_types[i].get()
            self._create_params_for_type(i, condition_type)
        
        # 更新文件列表显示
        self._filter_file_list()
    
    def _filter_time_columns(self, columns):
        """过滤掉时间列，用于关键值追踪下拉框
        
        时间列通常包含：时间、time、date 等关键词
        """
        if not columns:
            return []
        
        time_keywords = ['时间', 'time', 'date', 'timestamp', 'datetime']
        filtered = []
        for col in columns:
            col_lower = col.lower()
            # 检查是否包含时间关键词
            is_time_col = False
            for keyword in time_keywords:
                if keyword in col_lower:
                    is_time_col = True
                    break
            if not is_time_col:
                filtered.append(col)
        return filtered
    
    def _filter_file_list(self):
        """根据选择的数据类型过滤文件列表"""
        if not hasattr(self, 'valid_excel_files'):
            return
        
        data_type = self.data_type_var.get()
        
        # 清空列表
        self.file_listbox.delete(0, tk.END)
        
        # 只显示选中类型的文件
        count = 0
        for file_info in self.valid_excel_files:
            if file_info['file_type'] == data_type:
                display_name = f"{file_info['filename']}"
                self.file_listbox.insert(tk.END, display_name)
                count += 1
        
        # 更新统计
        self.file_count_label.config(text=f"共 {count} 个{data_type}文件")
        
    def process_queue(self):
        """处理线程发送的消息"""
        try:
            while True:
                msg = self.queue.get_nowait()
                if msg['type'] == 'progress':
                    # 更新进度条值
                    self.progress_bar['value'] = msg['value']
                    if 'detail' in msg:
                        self.progress_detail.config(text=msg['detail'])
                    # 【修复】只有当动画未激活时才更新进度文字
                    if not getattr(self, '_loading_active', False):
                        self.progress_label.config(text=f"{msg['text']} ({int(msg['value'])}%)")
                    self.root.update_idletasks()
                elif msg['type'] == 'loading_start':
                    # 启动加载动画
                    self.progress_bar['value'] = msg.get('value', 15)
                    if 'detail' in msg:
                        self.progress_detail.config(text=msg['detail'])
                    self._start_loading_animation(msg.get('text', '正在加载'))
                elif msg['type'] == 'loading_stop':
                    # 停止加载动画
                    self._stop_loading_animation()
                elif msg['type'] == 'error':
                    self._stop_loading_animation()
                    self._stop_simulated_progress()
                    messagebox.showerror("错误", msg['text'])
                    self.reset_progress()
                elif msg['type'] == 'success':
                    self._stop_loading_animation()
                    self._stop_simulated_progress()
                    # 不重置耗时，让用户看到总耗时
                    self.reset_progress(reset_time=False)
                elif msg['type'] == 'finish':
                    self._stop_loading_animation()
                    self._stop_simulated_progress()
                    self.finish_file_reading(
                        msg['file_path'],
                        msg['df'],
                        msg['time_column'],
                        msg.get('mode', 'normal')
                    )
                elif msg['type'] == 'start_simulated_progress':
                    # 启动模拟进度动画
                    self._start_simulated_progress(
                        msg.get('text', '正在处理'),
                        msg.get('start_value', 0),
                        msg.get('end_value', 100),
                        msg.get('duration', 30),
                        msg.get('show_time', False)
                    )
                elif msg['type'] == 'stop_simulated_progress':
                    self._stop_simulated_progress()
        except queue.Empty:
            pass
        finally:
            self.root.after(50, self.process_queue)
    
    def _start_loading_animation(self, base_text):
        """启动加载动画"""
        self._loading_text = base_text
        self._loading_dots = 0
        self._loading_active = True
        self._update_loading_animation()
    
    def _update_loading_animation(self):
        """更新加载动画（动态点点点效果）"""
        if not getattr(self, '_loading_active', False):
            return
        
        self._loading_dots = (self._loading_dots + 1) % 4
        dots = '.' * self._loading_dots
        progress_value = self.progress_bar['value']
        self.progress_label.config(text=f"{self._loading_text}{dots}")
        
        # 继续动画
        self._loading_timer = self.root.after(300, self._update_loading_animation)
    
    def _stop_loading_animation(self):
        """停止加载动画"""
        self._loading_active = False
        if hasattr(self, '_loading_timer'):
            try:
                self.root.after_cancel(self._loading_timer)
            except:
                pass
    
    def _update_progress_ui(self, text, value, detail=None, elapsed_time=None):
        """统一更新进度的方法，避免冲突"""
        self.progress_bar['value'] = value
        if detail:
            self.progress_detail.config(text=detail)
        # 只有当动画未激活时才更新进度文字
        if not getattr(self, '_loading_active', False) and not getattr(self, '_simulated_progress_active', False):
            self.progress_label.config(text=f"{text} ({int(value)}%)")
        # 更新耗时显示
        if elapsed_time is not None:
            self.time_label.config(text=f"耗时: {elapsed_time:.1f}秒", foreground="gray")
        self.root.update_idletasks()
    
    def _start_simulated_progress(self, text, start_value, end_value, duration, show_time=False):
        """启动模拟进度动画（在长时间操作期间平滑递增进度）"""
        self._simulated_text = text
        self._simulated_start = start_value
        self._simulated_end = end_value
        self._simulated_duration = duration  # 预估持续时间（秒）
        self._simulated_start_time = None
        self._simulated_progress_active = True
        self._simulated_show_time = show_time
        
        # 初始化进度
        self.progress_bar['value'] = start_value
        self.progress_detail.config(text='请稍候，正在处理...')
        
        # 启动动画
        self._update_simulated_progress()
    
    def _update_simulated_progress(self):
        """更新模拟进度"""
        import time
        
        if not getattr(self, '_simulated_progress_active', False):
            return
        
        # 初始化开始时间
        if self._simulated_start_time is None:
            self._simulated_start_time = time.time()
        
        # 计算当前进度 - 使用非线性曲线，前期快后期慢，更符合用户预期
        elapsed = time.time() - self._simulated_start_time
        # 使用平方根曲线，让进度看起来更均匀
        progress_ratio = min((elapsed / self._simulated_duration) ** 0.7, 0.95)
        
        current_value = self._simulated_start + (self._simulated_end - self._simulated_start) * progress_ratio
        
        # 更新UI
        self.progress_bar['value'] = current_value
        dots = '.' * (int(elapsed * 2) % 4)  # 动态点点点
        self.progress_label.config(text=f"{self._simulated_text}{dots} ({int(current_value)}%)")
        
        # 更新耗时显示
        if getattr(self, '_simulated_show_time', False):
            self.time_label.config(text=f"耗时: {elapsed:.1f}秒", foreground="gray")
        
        # 继续动画（每200ms更新一次，减少UI压力）
        self._simulated_timer = self.root.after(200, self._update_simulated_progress)
    
    def _stop_simulated_progress(self):
        """停止模拟进度"""
        self._simulated_progress_active = False
        if hasattr(self, '_simulated_timer'):
            try:
                self.root.after_cancel(self._simulated_timer)
            except:
                pass
            
    def reset_progress(self, reset_time=True):
        """重置进度条
        
        Args:
            reset_time: 是否重置耗时显示，默认True。文件读取完成后传入False以保留总耗时
        """
        self.progress_bar['value'] = 0
        self.progress_label.config(text="就绪", foreground="blue")
        self.progress_detail.config(text="")
        if reset_time:
            self.time_label.config(text="耗时: 0.0秒", foreground="gray")
        self.root.update_idletasks()
        
    def select_file(self):
        """选择Excel文件"""
        import time
        
        file_path = filedialog.askopenfilename(
            title="选择数据文件",
            filetypes=[("Excel和CSV文件", "*.xlsx *.xls *.csv"), ("所有文件", "*.*")]
        )
        
        if file_path:
            # 重置耗时显示（新操作开始）
            self.time_label.config(text="耗时: 0.0秒", foreground="gray")
            
            # 禁用按钮防止重复点击
            self.generate_btn.config(state=tk.DISABLED)
            
            # 显示初始进度
            self.queue.put({
                'type': 'progress',
                'text': '正在初始化...',
                'value': 5,
                'detail': '准备读取Excel文件'
            })
            
            # 启动计时器
            self._load_start_time = time.time()
            
            # 使用线程避免界面卡顿
            def read_excel_thread():
                try:
                    import os
                    import time
                    start_time = time.time()
                    
                    # 获取文件大小，给用户预估
                    file_size = os.path.getsize(file_path)
                    file_size_mb = file_size / (1024 * 1024)
                    
                    # 【优化】启动模拟进度动画（在读取期间平滑递增）
                    # 进度分配：读取文件占主要部分（5%-60%）
                    file_type_name = "CSV文件" if os.path.splitext(file_path)[1].lower() == '.csv' else "Excel文件"
                    self.queue.put({
                        'type': 'start_simulated_progress',
                        'text': f'正在读取{file_type_name}...',
                        'start_value': 5,
                        'end_value': 60,  # 读取完成后到60%
                        'duration': 25,  # 预估25秒完成读取
                        'show_time': True  # 显示耗时
                    })
                    
                    # 【性能优化】根据文件类型选择读取方式
                    df = None
                    engine_used = None
                    
                    # 检查文件扩展名
                    file_ext = os.path.splitext(file_path)[1].lower()
                    
                    if file_ext == '.csv':
                        # CSV 文件：使用 pd.read_csv
                        try:
                            # 尝试自动检测编码
                            for encoding in ['utf-8', 'gbk', 'gb2312', 'latin1']:
                                try:
                                    df = pd.read_csv(file_path, encoding=encoding)
                                    engine_used = f'csv ({encoding})'
                                    print(f"[PERF] 使用 pd.read_csv 读取成功，编码: {encoding}")
                                    break
                                except UnicodeDecodeError:
                                    continue
                            if df is None:
                                raise Exception("无法自动检测CSV编码")
                        except Exception as e:
                            raise Exception(f"读取CSV文件失败: {str(e)}")
                    else:
                        # Excel 文件：按优先级尝试不同的读取引擎
                        # 1. 尝试 calamine（Rust实现，最快）
                        try:
                            df = pd.read_excel(file_path, engine='calamine')
                            engine_used = 'calamine'
                            print(f"[PERF] 使用 calamine 引擎读取成功")
                        except Exception as e:
                            print(f"[PERF] calamine 不可用: {e}")
                            
                            # 2. 尝试 openpyxl（支持xlsx）
                            try:
                                df = pd.read_excel(file_path, engine='openpyxl')
                                engine_used = 'openpyxl'
                                print(f"[PERF] 使用 openpyxl 引擎读取")
                            except Exception as e2:
                                print(f"[PERF] openpyxl 不可用: {e2}")
                                
                                # 3. 使用默认引擎
                                df = pd.read_excel(file_path)
                                engine_used = 'default'
                    
                    read_time = time.time() - start_time
                    file_type = "CSV" if file_ext == '.csv' else "Excel"
                    print(f"[PERF] {file_type}读取耗时: {read_time:.2f}秒, 引擎: {engine_used}, 大小: {file_size_mb:.1f}MB")
                    
                    # 停止模拟进度
                    self.queue.put({'type': 'stop_simulated_progress'})
                    
                    # 处理数据阶段：60% - 65%
                    self.queue.put({
                        'type': 'progress',
                        'text': f'正在处理数据...',
                        'value': 65,
                        'detail': f'已读取 {len(df)} 行 × {len(df.columns)} 列'
                    })
                    
                    # 检查数据格式
                    if len(df.columns) < 1:
                        self.queue.put({
                            'type': 'error',
                            'text': '文件没有列数据！'
                        })
                        return
                    
                    # 使用第一列作为时间轴
                    time_column = df.columns[0]
                    
                    # 尝试将第一列转换为时间格式
                    try:
                        df[time_column] = pd.to_datetime(df[time_column])
                    except:
                        pass # 保持原样
                    
                    self.queue.put({
                        'type': 'progress',
                        'text': '正在准备界面...',
                        'value': 70,
                        'detail': '生成数据列选项'
                    })
                    
                    # 发送完成信号（根据分析模式）
                    mode = self.analysis_mode.get() if hasattr(self, 'analysis_mode') else 'normal'
                    self.queue.put({
                        'type': 'finish',
                        'file_path': file_path,
                        'df': df,
                        'time_column': time_column,
                        'mode': mode
                    })
                    
                except Exception as e:
                    self.queue.put({
                        'type': 'error',
                        'text': f'读取文件失败：{str(e)}'
                    })
            
            # 启动线程
            thread = threading.Thread(target=read_excel_thread)
            thread.daemon = True
            thread.start()
            
    def finish_file_reading(self, file_path, df, time_column, mode='normal'):
        """完成文件读取后的UI更新"""
        try:
            import time
            start_time = time.time()
            
            print(f"\n[PERF] ========== finish_file_reading 开始 ==========")
            print(f"[PERF] 文件路径: {file_path}")
            print(f"[PERF] df 列数: {len(df.columns)}")
            print(f"[PERF] 时间列: {time_column}")
            print(f"[PERF] 分析模式: {mode}")
            
            # 更新文件路径显示
            self.file_path_label.config(text=f"已选择: {file_path}")
            
            # 保存数据
            self.file_path = file_path  # 添加这一行
            self.df = df
            self.time_column = time_column
            
            # 根据分析模式处理
            if mode == 'fault':
                # 故障数据分析模式
                self._process_fault_data(df, time_column)
            else:
                # 普通数据分析模式
                self._process_normal_data(df, time_column)
            
            print(f"[PERF] finish_file_reading 总耗时: {time.time()-start_time:.2f}秒")
            
        except Exception as e:
            messagebox.showerror("错误", f"处理数据失败：{str(e)}")
            self.reset_progress()
    
    def _process_normal_data(self, df, time_column):
        """处理普通数据分析模式"""
        try:
            import time
            t1 = time.time()
            
            # 清空之前的复选框
            for widget in self.scrollable_frame.winfo_children():
                widget.destroy()
            self.column_vars.clear()
            print(f"[PERF] 清空控件耗时: {time.time()-t1:.2f}秒")
            
            # 为每一列（除了第一列时间列）创建复选框
            total_columns = len(df.columns) - 1
            print(f"[PERF] 准备创建 {total_columns} 个复选框")
            
            # 【性能优化】先创建所有变量
            t2 = time.time()
            columns_to_create = list(df.columns[1:])
            for col in columns_to_create:
                self.column_vars[col] = tk.IntVar(value=0)
            print(f"[PERF] 创建IntVar变量耗时: {time.time()-t2:.2f}秒")
            
            # 【关键优化】使用虚拟列表方式 - 只创建可见区域的复选框
            t3 = time.time()
            self._create_visible_checkboxes(columns_to_create)
            print(f"[PERF] 创建复选框耗时: {time.time()-t3:.2f}秒")
            
        except Exception as e:
            messagebox.showerror("错误", f"处理普通数据失败：{str(e)}")
            self.reset_progress()
    
    def _process_fault_data(self, df, time_column):
        """处理故障数据分析模式"""
        try:
            # 清空之前的复选框
            for widget in self.scrollable_frame.winfo_children():
                widget.destroy()
            self.column_vars.clear()
            
            # 显示故障分析信息
            info_label = ttk.Label(
                self.scrollable_frame,
                text="故障数据分析模式\n\n系统将自动分析寄存器20712-20717中的故障码数据。",
                font=('微软雅黑', 10),
                justify=tk.LEFT
            )
            info_label.pack(anchor=tk.W, padx=5, pady=10)
            
            # 提供故障寄存器列选择
            fault_registers = ['20712', '20713', '20714', '20715', '20716', '20717']
            register_label = ttk.Label(
                self.scrollable_frame,
                text="故障寄存器列（自动检测）：",
                font=('微软雅黑', 9, 'bold')
            )
            register_label.pack(anchor=tk.W, padx=5, pady=5)
            
            # 检测DataFrame中是否包含故障寄存器列
            detected_registers = []
            for reg in fault_registers:
                matching_cols = [col for col in df.columns if reg in str(col)]
                if matching_cols:
                    detected_registers.extend(matching_cols)
            
            if detected_registers:
                for reg_col in detected_registers:
                    var = tk.IntVar(value=1)
                    self.column_vars[reg_col] = var
                    cb = ttk.Checkbutton(
                        self.scrollable_frame,
                        text=f"[√] {reg_col}",
                        variable=var
                    )
                    cb.pack(anchor=tk.W, padx=20, pady=2)
                
                info_text = f"\n已检测到 {len(detected_registers)} 个故障寄存器列"
            else:
                info_text = "\n⚠ 未检测到标准故障寄存器列（20712-20717）\n请在下方手动选择数据列"
                
                # 显示所有可用列供手动选择
                all_cols_label = ttk.Label(
                    self.scrollable_frame,
                    text="\n所有数据列：",
                    font=('微软雅黑', 9, 'bold')
                )
                all_cols_label.pack(anchor=tk.W, padx=5, pady=5)
                
                for col in df.columns[1:]:  # 跳过时间列
                    var = tk.IntVar(value=0)
                    self.column_vars[col] = var
                    cb = ttk.Checkbutton(
                        self.scrollable_frame,
                        text=col,
                        variable=var
                    )
                    cb.pack(anchor=tk.W, padx=20, pady=1)
            
            info_label2 = ttk.Label(
                self.scrollable_frame,
                text=info_text,
                font=('微软雅黑', 9)
            )
            info_label2.pack(anchor=tk.W, padx=5, pady=5)
            
            # 更新进度
            self._update_progress_ui("故障数据加载完成", 100, f"共 {len(df)} 行数据")
            
        except Exception as e:
            messagebox.showerror("错误", f"处理故障数据失败：{str(e)}")
            self.reset_progress()
    
    def select_fault_file(self, file_path_label):
        """选择故障数据Excel文件
        
        Args:
            file_path_label: 显示文件路径的标签控件
        """
        file_path = filedialog.askopenfilename(
            title="选择故障数据文件",
            filetypes=[("Excel和CSV文件", "*.xlsx *.xls *.csv"), ("所有文件", "*.*")]
        )
        
        if file_path:
            file_path_label.config(text=f"已选择: {file_path}")
            self.fault_file_path = file_path
            
            # 更新进度
            self.fault_progress_label.config(text="正在读取文件...", foreground="blue")
            self.fault_progress_bar['value'] = 10
            
            # 使用线程读取文件
            def read_file_thread():
                try:
                    import os
                    file_ext = os.path.splitext(file_path)[1].lower()
                    
                    if file_ext == '.csv':
                        # CSV 文件：尝试自动检测编码
                        df = None
                        for encoding in ['utf-8', 'gbk', 'gb2312', 'latin1']:
                            try:
                                df = pd.read_csv(file_path, encoding=encoding)
                                break
                            except UnicodeDecodeError:
                                continue
                        if df is None:
                            raise Exception("无法自动检测CSV编码")
                    else:
                        df = pd.read_excel(file_path)
                    
                    # 保存数据
                    self.fault_df = df
                    self.fault_time_column = df.columns[0] if len(df.columns) > 0 else None
                    
                    # 更新UI（需要在主线程执行）
                    self.root.after(0, lambda: self._update_fault_file_ui(df))
                    
                except Exception as e:
                    # 使用默认参数捕获异常值，避免闭包问题
                    self.root.after(0, lambda err=str(e): self._show_fault_error(f"读取文件失败：{err}"))
            
            thread = threading.Thread(target=read_file_thread)
            thread.daemon = True
            thread.start()
    
    def _update_fault_file_ui(self, df):
        """更新故障文件选择后的UI
        
        新流程：检测寄存器列后，弹出配置对话框让用户选择要分析的寄存器
        """
        print("[DEBUG] _update_fault_file_ui 被调用")
        
        # 清空之前的寄存器复选框
        for widget in self.fault_register_frame.winfo_children():
            widget.destroy()
        self.fault_register_vars.clear()
        
        # 检测故障寄存器列（支持多种格式）
        # 1. 用户展示故障码：20712-20717
        # 2. 主故障码：20650-20689（偏移量0-39）
        user_registers = ['20712', '20713', '20714', '20715', '20716', '20717']
        main_registers = [str(20650 + i) for i in range(40)]  # 20650-20689
        
        detected_registers = []
        register_types = {}  # 记录寄存器类型
        
        # 检测用户展示故障码
        for reg in user_registers:
            matching_cols = [col for col in df.columns if reg in str(col)]
            for col in matching_cols:
                detected_registers.append((reg, col, '用户展示'))  # 修复：类型名统一为'用户展示'
                register_types[col] = '用户展示'
        
        # 检测主故障码
        for reg in main_registers:
            matching_cols = [col for col in df.columns if reg in str(col)]
            for col in matching_cols:
                detected_registers.append((reg, col, '主故障码'))
                register_types[col] = '主故障码'
        
        # 去重
        seen = set()
        unique_registers = []
        for item in detected_registers:
            if item[1] not in seen:
                seen.add(item[1])
                unique_registers.append(item)
        detected_registers = unique_registers
        
        print(f"[DEBUG] 检测到 {len(detected_registers)} 个故障寄存器列")
        
        # 保存检测结果到实例变量
        self.detected_registers = detected_registers
        self.register_types = register_types
        
        if detected_registers:
            print("[DEBUG] 准备弹出配置对话框...")
            # 更新进度提示
            self.fault_progress_label.config(
                text=f"检测到 {len(detected_registers)} 个故障寄存器列，请配置分析范围...", 
                foreground="blue"
            )
            self.fault_progress_bar['value'] = 50
            
            # 弹出配置对话框，让用户选择要分析的寄存器
            self.root.after(100, self._show_fault_config_dialog)
            print("[DEBUG] 已安排弹出配置对话框 (after 100ms)")
            
        else:
            ttk.Label(self.fault_register_frame, text="⚠ 未检测到标准故障寄存器列", 
                      font=('Arial', 9), foreground='orange').pack(anchor=tk.W, pady=5)
            ttk.Label(self.fault_register_frame, text="请确保Excel包含故障数据列", 
                      font=('Arial', 8), foreground='gray').pack(anchor=tk.W)
            
            self.fault_progress_bar['value'] = 100
            self.fault_progress_label.config(text="⚠ 未检测到故障寄存器，请检查文件格式", foreground="orange")
            
            # 更新右侧图表区域的提示
            self.fault_fig.clear()
            self.fault_ax = self.fault_fig.add_subplot(111)
            warning_text = (
                '⚠ 未检测到故障寄存器\n\n'
                '请确保Excel文件包含以下列：\n'
                '• 电网故障（用户展示）(20712)\n'
                '• 离网侧故障（用户展示）(20713)\n'
                '• PV侧故障（用户展示）(20714)\n'
                '• 电池故障（用户展示）(20715)\n'
                '• 油机故障（用户展示）(20716)\n'
                '• 系统故障（用户展示）(20717)\n'
                '• 或主故障码列 (20650-20689)'
            )
            self.fault_ax.text(0.5, 0.5, warning_text, 
                               ha='center', va='center', fontsize=12, color='orange',
                               linespacing=1.5)
            self.fault_ax.set_xlim(0, 1)
            self.fault_ax.set_ylim(0, 1)
            self.fault_ax.axis('off')
            self.fault_canvas.draw()
    
    def _on_fault_select(self, event):
        """处理故障列表选择事件"""
        try:
            selection = self.fault_listbox.curselection()
            if selection:
                index = selection[0]
                # 获取故障码（从列表项中提取）
                item_text = self.fault_listbox.get(index)
                # 格式: "E100: 电网未接 (193次)"
                fault_code = item_text.split(':')[0].strip()
                
                # 高亮显示选中的故障
                if hasattr(self, 'fault_stats') and self.fault_stats is not None:
                    # 在图表中高亮显示
                    self._highlight_fault_in_chart(fault_code)
        except Exception as e:
            print(f"[DEBUG] 选择故障时出错: {e}")
    
    def _highlight_fault_in_chart(self, fault_code):
        """在图表中高亮显示选中的故障"""
        try:
            if fault_code in self.fault_stats.index:
                print(f"[DEBUG] 高亮显示故障: {fault_code}")
        except Exception as e:
            print(f"[DEBUG] 高亮显示故障时出错: {e}")
    
    def _show_fault_error(self, error_msg):
        """显示故障分析错误"""
        messagebox.showerror("错误", error_msg)
        self.fault_progress_label.config(text="读取失败", foreground="red")
    
    def start_fault_analysis(self):
        """开始故障数据分析"""
        if not hasattr(self, 'fault_df') or self.fault_df is None:
            messagebox.showwarning("警告", "请先选择故障数据文件！")
            return
        
        # 检查是否已完成故障配置
        if not self.fault_config_done:
            messagebox.showinfo("提示", "请先选择Excel文件完成故障配置！")
            return
        
        # 获取选中的列名（直接从复选框获取）
        selected_registers = [col for col, var in self.fault_register_vars.items() if var.get()]
        if not selected_registers:
            messagebox.showwarning("警告", "请至少选择一个故障寄存器！")
            return
        
        print(f"[DEBUG] ========== start_fault_analysis ==========")
        print(f"[DEBUG] 选中的故障寄存器列: {selected_registers}")
        print(f"[DEBUG] fault_df 列数: {len(self.fault_df.columns)}")
        print(f"[DEBUG] fault_df 列名: {list(self.fault_df.columns)}")
        
        # 验证列名
        missing_cols = [col for col in selected_registers if col not in self.fault_df.columns]
        if missing_cols:
            print(f"[ERROR] 以下列名在 fault_df 中不存在: {missing_cols}")
            # 尝试模糊匹配
            for missing in missing_cols:
                matches = [c for c in self.fault_df.columns if missing.split('(')[0].strip() in c]
                print(f"[DEBUG] '{missing}' 可能的匹配: {matches}")
        
        self.fault_progress_label.config(text="正在分析故障数据...", foreground="blue")
        self.fault_progress_bar['value'] = 50
        
        # 使用线程进行分析
        def analyze_thread():
            try:
                fault_events = self._analyze_fault_data(selected_registers)
                self.root.after(0, lambda: self._display_fault_results(fault_events))
            except Exception as e:
                # 使用默认参数捕获异常值，避免闭包问题
                self.root.after(0, lambda err=str(e): self._show_fault_error(f"分析失败：{err}"))
        
        thread = threading.Thread(target=analyze_thread)
        thread.daemon = True
        thread.start()
    
    def _analyze_fault_data(self, selected_registers):
        """分析故障数据
        
        Args:
            selected_registers: 选中的寄存器列列表
            
        Returns:
            故障事件列表
        """
        print("\n[DEBUG] ========== 开始故障数据分析 ==========")
        print(f"[DEBUG] FAULT_PARSER_AVAILABLE = {FAULT_PARSER_AVAILABLE}")
        print(f"[DEBUG] 故障码来源设置: {getattr(self, 'fault_source', 'builtin')}")
        
        fault_events = []
        
        # 检查故障码定义是否已加载
        if FAULT_PARSER_AVAILABLE:
            total_bits = sum(len(bits) for bits in MAIN_FAULT_CODES.values())
            print(f"[DEBUG] 故障码定义状态: {len(MAIN_FAULT_CODES)} 个寄存器, {total_bits} 个故障码")
            
            # 只有在使用外部文件且完全没有加载时才尝试加载
            # 如果用户选择内置配置，使用模块初始化时已加载的定义
            if total_bits == 0 and hasattr(self, 'fault_source') and self.fault_source == 'external':
                print("[DEBUG] 故障码定义未加载，尝试从外部文件加载...")
                from fault_code_parser import load_fault_codes_from_point_table
                load_fault_codes_from_point_table(self.fault_external_file if hasattr(self, 'fault_external_file') else None)
                total_bits = sum(len(bits) for bits in MAIN_FAULT_CODES.values())
            elif total_bits == 0:
                print("[DEBUG] 使用内置故障码定义（用户展示故障码 20712-20717）")
                # 内置定义已经在 USER_FAULT_CODES 中，无需额外加载
            
            print(f"[DEBUG] ALL_FAULT_CODES 寄存器数: {len(ALL_FAULT_CODES)}")
            print(f"[DEBUG] MAIN_FAULT_CODES 寄存器数: {len(MAIN_FAULT_CODES)}")
            print(f"[DEBUG] 总故障码定义数: {total_bits}")
            
            # 检查几个关键寄存器是否有定义
            check_regs = [20650, 20652, 20659, 20663, 20664, 20712, 20715]
            for reg in check_regs:
                if reg in ALL_FAULT_CODES:
                    print(f"[DEBUG]   寄存器 {reg}: 有 {len(ALL_FAULT_CODES[reg])} 个bit定义")
                else:
                    print(f"[DEBUG]   寄存器 {reg}: 无定义")
        else:
            print("[DEBUG] ⚠️ 故障码解析模块不可用!")
            return fault_events
        
        # 构建寄存器检测列表
        user_registers = ['20712', '20713', '20714', '20715', '20716', '20717']
        main_registers = [str(20650 + i) for i in range(40)]  # 20650-20689
        all_registers = user_registers + main_registers
        
        print(f"[DEBUG] 选中的寄存器列: {selected_registers}")
        print(f"[DEBUG] 数据行数: {len(self.fault_df)}")
        
        # 打印 DataFrame 实际的列名，用于调试
        print(f"[DEBUG] DataFrame 列名前10个: {list(self.fault_df.columns[:10])}")
        
        # 验证列名是否存在于 DataFrame 中
        missing_cols = [col for col in selected_registers if col not in self.fault_df.columns]
        if missing_cols:
            print(f"[ERROR] 以下列名在 DataFrame 中不存在: {missing_cols}")
            # 尝试模糊匹配
            for missing_col in missing_cols:
                # 查找相似的列名
                similar_cols = [c for c in self.fault_df.columns if missing_col.split('(')[0].strip() in str(c)]
                if similar_cols:
                    print(f"[DEBUG] '{missing_col}' 的相似列: {similar_cols}")
        
        # 统计非空值
        for col in selected_registers:
            if col not in self.fault_df.columns:
                print(f"[ERROR] 列 '{col}' 不存在于 DataFrame 中，跳过")
                continue
            non_null_count = self.fault_df[col].notna().sum()
            print(f"[DEBUG] 列 '{col}' 非空值数量: {non_null_count}")
            if non_null_count > 0:
                # 显示前3个非空值示例
                sample_values = self.fault_df[self.fault_df[col].notna()][col].head(3).tolist()
                print(f"[DEBUG]   示例值: {sample_values}")
        
        # 记录解析统计
        parse_success_count = 0
        parse_no_fault_count = 0  # 无故障（值为空或bit未定义）
        parse_fail_count = 0      # 真正的解析失败
        
        for idx, row in self.fault_df.iterrows():
            timestamp = row[self.fault_time_column] if self.fault_time_column else idx
            
            for col in selected_registers:
                # 从列名中提取寄存器编号
                register_num = None
                for reg in all_registers:
                    if reg in str(col):
                        register_num = int(reg)
                        break
                
                if register_num is None:
                    continue
                
                # 获取寄存器值（先检查列是否存在）
                if col not in row.index:
                    continue
                value = row[col]
                if pd.isna(value):
                    continue
                
                # 解析bit位
                bits = parse_bit_list(value)
                
                # 如果没有bit位（值为空列表或0），跳过
                if not bits:
                    parse_no_fault_count += 1
                    continue
                
                # 解析故障码（对于未定义的bit，也要显示）
                faults = parse_fault_value(register_num, value)
                defined_bits = set(f.bit for f in faults) if faults else set()
                
                # 对于未定义的bit，生成默认故障信息
                for bit in bits:
                    if bit in defined_bits:
                        # 已定义的bit，使用解析结果
                        continue
                    
                    # 未定义的bit，生成默认信息
                    register_name = get_register_name(register_num)
                    fault_code = f'E{register_num}_{bit}'
                    description = f'未定义故障(bit{bit})'
                    category = '未知分类'
                    severity = '告警'
                    
                    # 添加到故障事件
                    fault_events.append({
                        'timestamp': timestamp,
                        'register': register_num,
                        'fault_code': fault_code,
                        'description': description,
                        'category': category,
                        'severity': severity,
                        'register_name': register_name
                    })
                    
                    parse_fail_count += 1
                    if parse_fail_count == 1:
                        print(f"[DEBUG] ⚠️ 发现未定义bit位: 寄存器={register_num}, bit{bit}")
                
                # 添加已定义的故障
                if faults:
                    parse_success_count += 1
                    if parse_success_count == 1:
                        print(f"[DEBUG] ✅ 第一次成功解析: 寄存器={register_num}, 值='{value}'")
                        for f in faults:
                            print(f"[DEBUG]    -> {f.code}: {f.name_cn}")
                    
                    for fault in faults:
                        fault_events.append({
                            'timestamp': timestamp,
                            'register': register_num,
                            'fault_code': fault.code,
                            'description': fault.name_cn,
                            'category': get_fault_category(fault.code),
                            'severity': fault.fault_type,
                            'register_name': fault.register_name
                        })
        
        print(f"[DEBUG] 解析统计: 成功={parse_success_count}, 无故障={parse_no_fault_count}, 未定义bit={parse_fail_count}")
        print(f"[DEBUG] ========== 分析完成，共检测到 {len(fault_events)} 个故障事件 ==========\n")
        return fault_events
    
    def _display_fault_results(self, fault_events):
        """显示故障分析结果"""
        if not fault_events:
            messagebox.showinfo("分析完成", "未检测到任何故障！")
            self.fault_progress_label.config(text="分析完成，无故障", foreground="green")
            return
        
        # 转换为DataFrame
        fault_df = pd.DataFrame(fault_events)
        
        # 更新进度
        self.fault_progress_label.config(text=f"检测到 {len(fault_events)} 个故障事件", foreground="green")
        self.fault_progress_bar['value'] = 100
        
        # 更新右侧故障列表
        self._update_fault_listbox(fault_df)
        
        # 绘制故障分析图表
        self._plot_fault_results(fault_df)
    
    def _update_fault_listbox(self, fault_df):
        """更新故障统计（内部方法）"""
        # 生成故障统计
        fault_stats = fault_df.groupby('fault_code').agg({
            'description': 'first',
            'category': 'first',
            'severity': 'first',
            'timestamp': 'count'
        }).rename(columns={'timestamp': 'count'}).sort_values('count', ascending=False)
        
        # 保存故障统计供后续使用
        self.fault_stats = fault_stats
        return fault_stats
    
    def _plot_fault_results(self, fault_df):
        """绘制故障分析结果图表 - 四宫格布局"""
        # 清空当前图表框架
        for widget in self.fault_chart_frame.winfo_children():
            widget.destroy()
        
        # 生成故障统计
        fault_stats = self._update_fault_listbox(fault_df)
        
        # 计算所有故障数量
        total_faults = len(fault_stats)
        
        # 每页显示的故障数量
        self.faults_per_page = 10
        self.current_fault_start = 0
        self.fault_stats = fault_stats
        self.fault_events_df = fault_df  # 故障事件结果，不要覆盖原始数据 self.fault_df
        
        # ===== 四宫格布局 =====
        # 创建主框架
        main_frame = ttk.Frame(self.fault_chart_frame)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 配置网格权重
        main_frame.columnconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(0, weight=1)
        main_frame.rowconfigure(1, weight=1)
        
        # ===== 左上：故障频率统计（带滚动条）=====
        top_left_frame = ttk.Frame(main_frame)
        top_left_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 5), pady=(0, 5))
        
        # 左侧：图表区域
        chart_container = ttk.Frame(top_left_frame)
        chart_container.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # 右侧：滚动条区域
        scrollbar_frame = ttk.Frame(top_left_frame, width=30)
        scrollbar_frame.pack(side=tk.RIGHT, fill=tk.Y, padx=(5, 0))
        scrollbar_frame.pack_propagate(False)
        
        # 创建滚动条
        if total_faults > self.faults_per_page:
            max_scroll = total_faults - self.faults_per_page
            
            up_btn = ttk.Button(scrollbar_frame, text="▲", width=3)
            up_btn.pack(side=tk.TOP, pady=2)
            
            self.fault_scroll = ttk.Scale(
                scrollbar_frame, 
                from_=0, 
                to=max_scroll,
                orient=tk.VERTICAL,
                command=self._on_fault_scroll
            )
            self.fault_scroll.set(0)
            self.fault_scroll.pack(side=tk.TOP, fill=tk.Y, expand=True, pady=2)
            
            down_btn = ttk.Button(scrollbar_frame, text="▼", width=3)
            down_btn.pack(side=tk.BOTTOM, pady=2)
            
            up_btn.config(command=lambda: self._scroll_fault_up())
            down_btn.config(command=lambda: self._scroll_fault_down())
            
            def _on_mousewheel(event):
                if event.delta > 0:
                    self._scroll_fault_up()
                else:
                    self._scroll_fault_down()
            chart_container.bind_all("<MouseWheel>", _on_mousewheel)
        else:
            self.fault_scroll = None
        
        # 创建故障频率图
        self.fault_freq_fig = Figure(figsize=(6, 5), dpi=100)
        self.fault_freq_ax = self.fault_freq_fig.add_subplot(111)
        
        self.fault_freq_canvas = FigureCanvasTkAgg(self.fault_freq_fig, chart_container)
        self.fault_freq_canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        
        # ===== 右上：故障类别分布 =====
        top_right_frame = ttk.Frame(main_frame)
        top_right_frame.grid(row=0, column=1, sticky="nsew", padx=(5, 0), pady=(0, 5))
        
        self.category_fig = Figure(figsize=(5, 5), dpi=100)
        self.category_ax = self.category_fig.add_subplot(111)
        
        category_counts = fault_df['category'].value_counts()
        total_count = category_counts.sum()
        
        # 过滤掉占比小于1%的类别，合并为"其他"
        threshold = 0.01  # 1%阈值
        main_categories = category_counts[category_counts / total_count >= threshold]
        other_count = category_counts[category_counts / total_count < threshold].sum()
        
        if other_count > 0:
            main_categories = pd.concat([main_categories, pd.Series([other_count], index=['其他'])])
        
        # 使用饼图，autopct只显示大于3%的百分比
        def make_autopct(values):
            def my_autopct(pct):
                return f'{pct:.1f}%' if pct > 3 else ''
            return my_autopct
        
        category_colors = ['#FF6B6B', '#FFE66D', '#4ECDC4', '#95E1D3', '#F38181', '#DDA0DD']
        wedges, texts, autotexts = self.category_ax.pie(
            main_categories, 
            autopct=make_autopct(main_categories),
            colors=category_colors[:len(main_categories)],
            startangle=90
        )
        
        # 使用图例代替直接标签
        self.category_ax.legend(
            wedges, 
            [f'{label}: {count}次' for label, count in zip(main_categories.index, main_categories)],
            loc='center left',
            bbox_to_anchor=(1, 0.5),
            fontsize=9
        )
        self.category_ax.set_title('故障类别分布', fontweight='bold', fontsize=11)
        self.category_fig.tight_layout()
        
        self.category_canvas = FigureCanvasTkAgg(self.category_fig, top_right_frame)
        self.category_canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        self.category_canvas.draw()
        
        # ===== 左下：故障时间分布 =====
        bottom_left_frame = ttk.Frame(main_frame)
        bottom_left_frame.grid(row=1, column=0, sticky="nsew", padx=(0, 5), pady=(5, 0))
        
        self.time_fig = Figure(figsize=(6, 4), dpi=100)
        self.time_ax = self.time_fig.add_subplot(111)
        
        fault_df_sorted = fault_df.sort_values('timestamp')
        try:
            fault_df_sorted['hour'] = pd.to_datetime(fault_df_sorted['timestamp']).dt.hour
            hourly_counts = fault_df_sorted.groupby('hour').size()
            self.time_ax.bar(hourly_counts.index, hourly_counts.values, color='#4ECDC4', alpha=0.7)
            self.time_ax.set_xlabel('小时')
            self.time_ax.set_ylabel('故障次数')
            self.time_ax.set_title('故障时间分布', fontweight='bold', fontsize=11)
        except:
            self.time_ax.text(0.5, 0.5, '时间数据格式不支持', ha='center', va='center')
        self.time_fig.tight_layout(pad=1.5)
        
        self.time_canvas = FigureCanvasTkAgg(self.time_fig, bottom_left_frame)
        self.time_canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        self.time_canvas.draw()
        
        # ===== 右下：故障严重程度分布 =====
        bottom_right_frame = ttk.Frame(main_frame)
        bottom_right_frame.grid(row=1, column=1, sticky="nsew", padx=(5, 0), pady=(5, 0))
        
        self.severity_fig = Figure(figsize=(5, 4), dpi=100)
        self.severity_ax = self.severity_fig.add_subplot(111)
        
        severity_counts = fault_df['severity'].value_counts()
        severity_colors = {'故障': '#FF6B6B', '告警': '#FFE66D', '提示': '#4ECDC4'}
        colors = [severity_colors.get(s, '#95E1D3') for s in severity_counts.index]
        
        # 使用图例代替直接标签
        wedges, texts, autotexts = self.severity_ax.pie(
            severity_counts, 
            autopct=lambda pct: f'{pct:.1f}%' if pct > 3 else '',
            colors=colors,
            startangle=90
        )
        
        self.severity_ax.legend(
            wedges,
            [f'{label}: {count}次' for label, count in zip(severity_counts.index, severity_counts)],
            loc='center left',
            bbox_to_anchor=(1, 0.5),
            fontsize=9
        )
        self.severity_ax.set_title('故障严重程度分布', fontweight='bold', fontsize=11)
        self.severity_fig.tight_layout()
        
        self.severity_canvas = FigureCanvasTkAgg(self.severity_fig, bottom_right_frame)
        self.severity_canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        self.severity_canvas.draw()
        
        # 绘制初始故障频率图
        self._draw_fault_freq_chart()
    
    def _draw_fault_freq_chart(self):
        """绘制故障频率图表（当前可见范围）"""
        # 清空图表
        self.fault_freq_ax.clear()
        
        fault_stats = self.fault_stats
        total_faults = len(fault_stats)
        
        # 获取当前显示范围的故障
        start = self.current_fault_start
        end = min(start + self.faults_per_page, total_faults)
        visible_faults = fault_stats.iloc[start:end]
        
        # 绘制柱状图
        colors = ['#FF6B6B' if s == '故障' else '#FFE66D' if s == '告警' else '#4ECDC4' 
                  for s in visible_faults['severity']]
        
        bars = self.fault_freq_ax.barh(range(len(visible_faults)), visible_faults['count'], color=colors)
        self.fault_freq_ax.set_yticks(range(len(visible_faults)))
        
        # 显示故障码和中文名称（带全局编号）
        self.fault_freq_ax.set_yticklabels(
            [f"{start+i+1}. {code}: {desc}" 
             for i, (code, desc) in enumerate(zip(visible_faults.index, visible_faults['description']))],
            fontsize=9
        )
        
        self.fault_freq_ax.set_xlabel('出现次数', fontsize=10)
        self.fault_freq_ax.set_title(
            f'故障频率统计 (共{total_faults}种故障，当前显示 {start+1}-{end})', 
            fontweight='bold', fontsize=11
        )
        self.fault_freq_ax.invert_yaxis()
        
        # 添加数值标签
        for bar, count in zip(bars, visible_faults['count']):
            self.fault_freq_ax.text(
                bar.get_width() + 0.5, 
                bar.get_y() + bar.get_height()/2,
                f'{count}', va='center', fontsize=8
            )
        
        self.fault_freq_fig.tight_layout()
        self.fault_freq_canvas.draw()
    
    def _on_fault_scroll(self, value):
        """滚动条移动时的回调"""
        new_start = int(float(value))
        if new_start != self.current_fault_start:
            self.current_fault_start = new_start
            self._draw_fault_freq_chart()
    
    def _scroll_fault_up(self):
        """向上滚动（显示前面的故障）"""
        if self.current_fault_start > 0:
            self.current_fault_start = max(0, self.current_fault_start - 1)
            if self.fault_scroll:
                self.fault_scroll.set(self.current_fault_start)
            self._draw_fault_freq_chart()
    
    def _scroll_fault_down(self):
        """向下滚动（显示后面的故障）"""
        max_start = len(self.fault_stats) - self.faults_per_page
        if self.current_fault_start < max_start:
            self.current_fault_start = min(max_start, self.current_fault_start + 1)
            if self.fault_scroll:
                self.fault_scroll.set(self.current_fault_start)
            self._draw_fault_freq_chart()
    
    def _generate_fault_analysis(self):
        """生成故障数据分析图表和报告"""
        try:
            print(f"\n[DEBUG] ========== 开始故障数据分析 ==========")
            
            if not FAULT_PARSER_AVAILABLE:
                messagebox.showerror("错误", "故障码解析模块未正确加载！")
                return
            
            # 确保故障码定义已加载
            load_fault_codes_from_excel()
            
            if self.df is None:
                messagebox.showwarning("警告", "请先选择Excel文件！")
                return
            
            # 获取选中的列（故障寄存器列）
            selected_columns = [col for col, var in self.column_vars.items() if var.get() == 1]
            if not selected_columns:
                messagebox.showwarning("警告", "请至少选择一个故障寄存器列！")
                return
            
            print(f"[DEBUG] 选中的故障寄存器列: {selected_columns}")
            
            # 构建寄存器检测列表
            user_registers = ['20712', '20713', '20714', '20715', '20716', '20717']
            main_registers = [str(20650 + i) for i in range(40)]  # 20650-20689
            all_registers = user_registers + main_registers
            
            # 解析所有行的故障数据
            fault_events = []
            for idx, row in self.df.iterrows():
                timestamp = row[self.time_column] if self.time_column and self.time_column in row.index else idx
                
                for col in selected_columns:
                    # 从列名中提取寄存器编号
                    register_num = None
                    for reg in all_registers:
                        if reg in str(col):
                            register_num = int(reg)
                            break
                    
                    if register_num is None:
                        continue
                    
                    # 获取寄存器值
                    value = row[col]
                    if pd.isna(value):
                        continue
                    
                    # 解析故障码（支持字符串列表格式如 '[6,15]' 或数字格式）
                    faults = parse_fault_value(register_num, value)
                    for fault in faults:
                        fault_events.append({
                            'timestamp': timestamp,
                            'register': register_num,
                            'fault_code': fault.code,
                            'description': fault.name_cn,
                            'category': get_fault_category(fault.code),
                            'severity': fault.fault_type,
                            'register_name': fault.register_name
                        })
            
            print(f"[DEBUG] 检测到 {len(fault_events)} 个故障事件")
            
            if not fault_events:
                messagebox.showinfo("提示", "未检测到任何故障！")
                return
            
            # 转换为DataFrame
            fault_df = pd.DataFrame(fault_events)
            
            # 保存故障事件结果（不要覆盖原始数据）
            self.fault_events_df = fault_df
            
            # 生成故障统计
            fault_stats = fault_df.groupby('fault_code').agg({
                'description': 'first',
                'category': 'first',
                'severity': 'first',
                'timestamp': 'count'
            }).rename(columns={'timestamp': 'count'}).sort_values('count', ascending=False)
            
            # 生成可视化图表
            self._plot_fault_analysis(fault_df, fault_stats)
            
            # 生成故障报告
            self._generate_fault_report(fault_df, fault_stats)
            
        except ImportError as e:
            messagebox.showerror("错误", f"无法导入故障码解析模块：{str(e)}\n请确保 fault_code_parser.py 存在。")
        except Exception as e:
            messagebox.showerror("错误", f"故障数据分析失败：{str(e)}")
            import traceback
            traceback.print_exc()
    
    def _plot_fault_analysis(self, fault_df, fault_stats):
        """绘制故障分析图表"""
        try:
            # 清空图表
            self.ax.clear()
            self._clear_legend_panel()
            
            # 创建2x2子图
            fig = self.ax.figure
            fig.clear()
            
            # 子图1：故障频率柱状图（前10）
            ax1 = fig.add_subplot(2, 2, 1)
            top_faults = fault_stats.head(10)
            colors = ['#FF6B6B' if s == '严重' else '#FFE66D' if s == '警告' else '#4ECDC4' 
                      for s in top_faults['severity']]
            bars = ax1.barh(range(len(top_faults)), top_faults['count'], color=colors)
            ax1.set_yticks(range(len(top_faults)))
            # 显示故障码和中文名称
            ax1.set_yticklabels([f"{idx+1}. {code}: {desc}" for idx, (code, desc) in 
                                enumerate(zip(top_faults.index, top_faults['description']))])
            ax1.set_xlabel('出现次数')
            ax1.set_title('故障频率 Top 10', fontweight='bold')
            ax1.invert_yaxis()
            
            # 添加数值标签
            for i, (bar, count) in enumerate(zip(bars, top_faults['count'])):
                ax1.text(bar.get_width() + 0.5, bar.get_y() + bar.get_height()/2,
                        f'{count}', va='center', fontsize=8)
            
            # 子图2：故障类别分布饼图
            ax2 = fig.add_subplot(2, 2, 2)
            category_counts = fault_df['category'].value_counts()
            ax2.pie(category_counts, labels=category_counts.index, autopct='%1.1f%%',
                   colors=['#FF6B6B', '#FFE66D', '#4ECDC4', '#95E1D3', '#F38181'])
            ax2.set_title('故障类别分布', fontweight='bold')
            
            # 子图3：故障时间线
            ax3 = fig.add_subplot(2, 2, 3)
            fault_df_sorted = fault_df.sort_values('timestamp')
            fault_df_sorted['hour'] = pd.to_datetime(fault_df_sorted['timestamp']).dt.hour
            hourly_counts = fault_df_sorted.groupby('hour').size()
            ax3.bar(hourly_counts.index, hourly_counts.values, color='#4ECDC4', alpha=0.7)
            ax3.set_xlabel('小时')
            ax3.set_ylabel('故障次数')
            ax3.set_title('故障时间分布', fontweight='bold')
            
            # 子图4：严重程度分布
            ax4 = fig.add_subplot(2, 2, 4)
            severity_counts = fault_df['severity'].value_counts()
            severity_colors = {'严重': '#FF6B6B', '警告': '#FFE66D', '提示': '#4ECDC4'}
            colors = [severity_colors.get(s, '#95E1D3') for s in severity_counts.index]
            ax4.pie(severity_counts, labels=severity_counts.index, autopct='%1.1f%%',
                   colors=colors)
            ax4.set_title('故障严重程度分布', fontweight='bold')
            
            fig.tight_layout()
            
            # 刷新画布
            self.canvas.draw()
            
        except Exception as e:
            print(f"[ERROR] 绘制故障图表失败: {str(e)}")
            import traceback
            traceback.print_exc()
    
    def _generate_fault_report(self, fault_df, fault_stats):
        """生成故障分析报告"""
        try:
            report_lines = [
                "=" * 60,
                "故障数据分析报告",
                "=" * 60,
                "",
                f"分析时间: {pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')}",
                f"数据范围: {fault_df['timestamp'].min()} 至 {fault_df['timestamp'].max()}",
                f"总故障事件数: {len(fault_df)}",
                f"故障类型数: {len(fault_stats)}",
                "",
                "=" * 60,
                "故障频率统计",
                "=" * 60,
                ""
            ]
            
            for idx, (fault_code, row) in enumerate(fault_stats.iterrows(), 1):
                report_lines.extend([
                    f"{idx}. {fault_code}",
                    f"   描述: {row['description']}",
                    f"   类别: {row['category']}",
                    f"   严重程度: {row['severity']}",
                    f"   出现次数: {row['count']}",
                    ""
                ])
            
            report_lines.extend([
                "=" * 60,
                "严重故障详情",
                "=" * 60,
                ""
            ])
            
            critical_faults = fault_df[fault_df['severity'] == '严重']
            if len(critical_faults) > 0:
                for _, row in critical_faults.iterrows():
                    report_lines.extend([
                        f"时间: {row['timestamp']}",
                        f"故障码: {row['fault_code']}",
                        f"描述: {row['description']}",
                        ""
                    ])
            else:
                report_lines.append("无严重故障")
            
            # 将报告存储到实例变量
            self.fault_report = "\n".join(report_lines)
            
            # 在诊断结果标签页显示报告
            self._display_fault_report()
            
            # 切换到诊断结果标签页
            self.notebook.select(self.diagnosis_tab)
            
            messagebox.showinfo("分析完成", 
                f"故障分析完成！\n检测到 {len(fault_df)} 个故障事件\n"
                f"涉及 {len(fault_stats)} 种故障类型\n\n"
                f"详细报告已显示在\"诊断结果\"标签页")
            
        except Exception as e:
            print(f"[ERROR] 生成故障报告失败: {str(e)}")
            import traceback
            traceback.print_exc()
    
    def _display_fault_report(self):
        """在诊断结果标签页显示故障报告"""
        try:
            # 清空诊断结果文本框
            self.diagnosis_text.config(state=tk.NORMAL)
            self.diagnosis_text.delete(1.0, tk.END)
            
            # 插入故障报告
            if hasattr(self, 'fault_report') and self.fault_report:
                self.diagnosis_text.insert(tk.END, self.fault_report)
            else:
                self.diagnosis_text.insert(tk.END, "无故障分析结果")
            
            self.diagnosis_text.config(state=tk.DISABLED)
            
        except Exception as e:
            print(f"[ERROR] 显示故障报告失败: {str(e)}")
    
    def _create_visible_checkboxes(self, columns_to_create):
        """创建可见区域的复选框（虚拟列表优化）"""
        total_columns = len(columns_to_create)
        
        # 检查滚动区域的高度，估算可见行数
        try:
            visible_height = self.scrollable_frame.winfo_height()
            if visible_height < 100:
                visible_height = 400  # 默认高度
        except:
            visible_height = 400
        
        # 每个复选框约20像素高，估算可见数量
        visible_count = min(int(visible_height / 20) + 10, total_columns)
        
        print(f"[PERF] 可见区域高度: {visible_height}px, 预估可见: {visible_count} 个")
        
        # 只创建前N个复选框
        for i, col in enumerate(columns_to_create[:visible_count]):
            var = self.column_vars[col]
            cb = ttk.Checkbutton(
                self.scrollable_frame,
                text=col,
                variable=var
            )
            cb.pack(anchor=tk.W, padx=5, pady=2)
        
        # 存储待创建的列，滚动时动态创建
        self._pending_columns = columns_to_create[visible_count:]
        self._created_column_count = visible_count
        
        # 绑定滚动事件，动态创建更多复选框
        if hasattr(self, 'canvas'):
            # 使用多种事件确保能捕获滚动
            self.canvas.bind('<MouseWheel>', self._on_column_scroll)
            self.canvas.bind('<ButtonRelease>', self._on_column_scroll)
            # 绑定滚动条变化
            self.canvas.bind('<Configure>', self._on_column_scroll)
        
        # 更新进度：创建控件阶段 70% -> 90%
        self._update_progress_ui("正在创建控件", 75, f'已加载 {total_columns} 个数据列')
        
        # 继续后续处理
        self._finish_column_creation(self.file_path, self.df, self.time_column)
    
    def _on_column_scroll(self, event=None):
        """滚动时动态创建更多复选框"""
        if not hasattr(self, '_pending_columns') or not self._pending_columns:
            return
        
        # 检查滚动位置
        try:
            canvas = self.canvas
            scroll_pos = canvas.yview()[1]  # 0-1, 1表示滚动到底部
            
            # 当滚动到50%时，创建更多（更早触发）
            if scroll_pos > 0.5:
                # 一次创建50个
                batch = self._pending_columns[:50]
                self._pending_columns = self._pending_columns[50:]
                
                if batch:
                    for col in batch:
                        if col in self.column_vars:
                            var = self.column_vars[col]
                            cb = ttk.Checkbutton(
                                self.scrollable_frame,
                                text=col,
                                variable=var
                            )
                            cb.pack(anchor=tk.W, padx=5, pady=2)
                            self._created_column_count += 1
                    
                    # 更新滚动区域
                    self.scrollable_frame.update_idletasks()
                    canvas.configure(scrollregion=canvas.bbox("all"))
        except Exception as e:
            print(f"[DEBUG] 滚动创建复选框异常: {e}")
    
    def _finish_column_creation(self, file_path, df, time_column):
        """复选框创建完成后的后续处理"""
        try:
            # 【新增】先重新应用自定义计算列
            print(f"\n[DEBUG] ========== 重新应用自定义计算列 ==========")
            print(f"[DEBUG] 当前有 {len(self.custom_columns)} 个自定义列配置")
            
            if self.custom_columns:
                reapplied_count = 0
                failed_columns = []
                
                for col_name, col_config in self.custom_columns.items():
                    print(f"[DEBUG] 尝试重新应用自定义列: {col_name}")
                    print(f"[DEBUG] 配置: {col_config}")
                    
                    formula = col_config.get('formula', '')
                    if not formula:
                        print(f"[DEBUG] 跳过: 公式为空")
                        continue
                    
                    # 检查公式中需要的原始列是否在新文件中存在
                    available_columns = [c for c in self.df.columns if c != self.time_column]
                    missing_columns = []
                    formula_columns = []
                    
                    # 提取公式中的列名（按最长匹配优先）
                    sorted_columns = sorted(available_columns, key=len, reverse=True)
                    remaining_formula = formula
                    
                    for col in sorted_columns:
                        if col in remaining_formula:
                            formula_columns.append(col)
                            remaining_formula = remaining_formula.replace(col, '')
                    
                    # 检查是否有缺失的列
                    for col in formula_columns:
                        if col not in self.df.columns:
                            missing_columns.append(col)
                    
                    if missing_columns:
                        print(f"[DEBUG] 失败: 缺少列 {missing_columns}")
                        failed_columns.append(f"{col_name} (缺少: {', '.join(missing_columns)})")
                        continue
                    
                    # 尝试重新计算
                    try:
                        # 构建计算表达式
                        calc_expr = formula
                        for col in formula_columns:
                            calc_expr = calc_expr.replace(col, f"self.df['{col}']")
                        
                        print(f"[DEBUG] 计算表达式: {calc_expr}")
                        
                        # 执行计算
                        self.df[col_name] = eval(calc_expr)
                        print(f"[DEBUG] 成功: 自定义列 {col_name} 已重新计算")
                        
                        # 为新列创建复选框
                        var = tk.IntVar(value=0)
                        self.column_vars[col_name] = var
                        cb = ttk.Checkbutton(
                            self.scrollable_frame,
                            text=col_name,
                            variable=var
                        )
                        cb.pack(anchor=tk.W, padx=5, pady=2)
                        
                        reapplied_count += 1
                        
                    except Exception as e:
                        print(f"[DEBUG] 失败: 计算错误 - {e}")
                        failed_columns.append(f"{col_name} (计算错误)")
                
                print(f"[DEBUG] 重新应用完成: 成功 {reapplied_count} 个，失败 {len(failed_columns)} 个")
                
                # 如果有失败的自定义列，显示警告
                if failed_columns:
                    failed_msg = "\n".join(failed_columns[:5])  # 最多显示5个
                    if len(failed_columns) > 5:
                        failed_msg += f"\n... 还有 {len(failed_columns) - 5} 个"
                    
                    messagebox.showwarning(
                        "自定义列警告",
                        f"以下自定义列无法重新应用，因为新文件缺少所需的数据列：\n\n{failed_msg}\n\n请检查数据文件或重新创建自定义列。"
                    )
            
            # 自动勾选：如果当前选择了常用分组，且新文件中有同名列，则自动勾选
            current_group_name = self.group_var.get()
            print(f"[DEBUG] 当前分组: {current_group_name}")
            
            if current_group_name and current_group_name in self.groups:
                group_data = self.groups[current_group_name]
                print(f"[DEBUG] 分组数据: {group_data}")
                matched_count = 0
                
                # 遍历分组中的列，如果在新文件中存在同名列，则自动勾选
                for col_name, checked in group_data.items():
                    print(f"[DEBUG] 检查列 '{col_name}'，勾选状态: {checked}")
                    if checked == 1 and col_name in self.column_vars:
                        self.column_vars[col_name].set(1)
                        matched_count += 1
                        print(f"[DEBUG] 已勾选: {col_name}")
                    elif checked == 1 and col_name not in self.column_vars:
                        print(f"[DEBUG] 警告: 列 '{col_name}' 不在 column_vars 中")
                
                # 启用生成按钮和禁用保存按钮
                self.generate_btn.config(state=tk.NORMAL)
                self.save_btn.config(state=tk.DISABLED)
                
                print(f"[DEBUG] 自动勾选完成，匹配 {matched_count} 列")
                
                # 如果有匹配的列，显示提示信息
                if matched_count > 0:
                    self.queue.put({
                        'type': 'success',
                        'text': f"成功读取Excel文件！\n总行数: {len(df)}\n总列数: {len(df.columns)}\n时间列: {self.time_column}\n\n[√] 已自动勾选分组 '{current_group_name}' 中的 {matched_count} 个数据列"
                    })
                else:
                    # 如果分组中没有可匹配的列，显示提示
                    self.queue.put({
                        'type': 'success',
                        'text': f"成功读取Excel文件！\n总行数: {len(df)}\n总列数: {len(df.columns)}\n时间列: {self.time_column}\n\n⚠ 分组 '{current_group_name}' 中的列在新文件中未找到匹配，请手动勾选"
                    })
            else:
                # 启用生成按钮和禁用保存按钮
                self.generate_btn.config(state=tk.NORMAL)
                self.save_btn.config(state=tk.DISABLED)
                
                # 没有选择分组时，正常显示成功消息
                self.queue.put({
                    'type': 'success',
                    'text': f"成功读取Excel文件！\n总行数: {len(df)}\n总列数: {len(df.columns)}\n时间列: {self.time_column}"
                })
            
            # 更新诊断下拉框选项
            self.update_diagnosis_comboboxes()
            
            # 计算总耗时
            import time
            total_time = time.time() - getattr(self, '_load_start_time', time.time())
            
            # 更新进度到100%，显示总耗时
            self._update_progress_ui("准备完成", 100, f'已完成，共 {len(df.columns)} 列数据', total_time)
            
            # 3秒后隐藏进度条（保留耗时显示）
            self.root.after(3000, lambda: self.reset_progress(reset_time=False))
            
        except Exception as e:
            messagebox.showerror("错误", f"处理数据失败：{str(e)}")
            self.reset_progress()
            
    def auto_select_common_columns(self):
        """自动识别并预选中常见的功率列"""
        # 常见的关键词模式
        priority_keywords = [
            ['逆变器', '有功', '功率', '总和'], # 逆变器有功功率总和
            ['直流', 'PV', '功率'], # 直流PV总功率
            ['电网', '功率'], # 电网总功率
            ['交流光伏', '有功', '功率', '总和'], # 交流光伏逆变器有功功率总和
            ['电池', '功率'], # 电池总功率
            ['负载', '功率'] # 负载总功率
        ]
        
        selected_count = 0
        selected_columns = []
        
        # 按优先级查找列
        for keywords in priority_keywords:
            found = False
            for col in self.df.columns[1:]: # 跳过时间列
                if found:
                    break
                # 检查列名是否包含所有关键词
                if all(keyword in str(col) for keyword in keywords):
                    if self.column_vars.get(col):
                        self.column_vars[col].set(1)
                        selected_count += 1
                        selected_columns.append(col)
                        found = True
                        break
        
        # 如果没有找到任何常见列，尝试单关键词匹配
        if selected_count == 0:
            single_keywords = ['逆变器', '电网', '电池', 'PV', '负载', '功率']
            for col in self.df.columns[1:]:
                if selected_count >= 5: # 最多预选5个
                    break
                for keyword in single_keywords:
                    if keyword in str(col) and self.column_vars.get(col):
                        # 避免重复选中
                        if self.column_vars[col].get() == 0:
                            self.column_vars[col].set(1)
                            selected_count += 1
                            selected_columns.append(col)
                            break
        
        # 更新图表标题，反映自动选择的列
        if selected_columns:
            # 生成简短的标题
            short_names = []
            for col in selected_columns:
                # 提取关键信息作为简短名称
                if '逆变器' in col and '功率' in col:
                    short_names.append('逆变器功率')
                elif '直流' in col and 'PV' in col:
                    short_names.append('直流PV功率')
                elif '电网' in col and '功率' in col:
                    short_names.append('电网功率')
                elif '电池' in col and '功率' in col:
                    short_names.append('电池功率')
                elif '负载' in col and '功率' in col:
                    short_names.append('负载功率')
                else:
                    # 使用列名的前15个字符
                    short_names.append(col[:15])
            
            self.title_var.set(f"{'-'.join(short_names[:3])} 趋势图" + ("..." if len(short_names) > 3 else ""))
            
    def filter_columns(self, event=None):
        """根据搜索框过滤列名"""
        search_text = self.search_var.get().lower()
        
        for widget in self.scrollable_frame.winfo_children():
            if isinstance(widget, ttk.Checkbutton):
                col_text = widget.cget("text").lower()
                if search_text in col_text:
                    widget.pack(anchor=tk.W, padx=5, pady=2)
                else:
                    widget.pack_forget()
        
        # 搜索后自动滚动到列表顶部
        self.canvas.yview_moveto(0)
        
    def select_all(self):
        """全选所有可见的列"""
        for widget in self.scrollable_frame.winfo_children():
            if isinstance(widget, ttk.Checkbutton) and widget.winfo_ismapped():
                col_name = widget.cget("text")
                self.column_vars[col_name].set(1)
                # 更新已勾选数据显示

    def deselect_all(self):
        """全不选所有列"""
        for col_name in self.column_vars:
            self.column_vars[col_name].set(0)
            # 更新已勾选数据显示
        
    def generate_plot(self):
        """生成折线图"""
        print(f"\n[DEBUG] ========== 开始生成折线图 ==========")
        
        # 检查分析模式
        if hasattr(self, 'analysis_mode') and self.analysis_mode.get() == 'fault':
            # 故障数据分析模式
            self._generate_fault_analysis()
            return
        
        if self.df is None:
            print("[ERROR] df is None")
            messagebox.showwarning("警告", "请先选择Excel文件！")
            return
        
        print(f"[DEBUG] df 列数: {len(self.df.columns)}")
        print(f"[DEBUG] column_vars 数量: {len(self.column_vars)}")
        print(f"[DEBUG] column_vars 中的列: {list(self.column_vars.keys())[:10]}...")
        
        # 获取选中的列
        selected_columns = [col for col, var in self.column_vars.items() if var.get() == 1]
        print(f"[DEBUG] 勾选的列数量: {len(selected_columns)}")
        print(f"[DEBUG] 勾选的列: {selected_columns}")
        
        # 检查并过滤不存在的列（切换文件时可能发生）
        valid_columns = [col for col in selected_columns if col in self.df.columns]
        print(f"[DEBUG] 有效列数量: {len(valid_columns)}")
        print(f"[DEBUG] 有效列: {valid_columns}")
        
        # 如果存在无效列，自动取消勾选并更新显示
        if len(valid_columns) < len(selected_columns):
            invalid_columns = [col for col in selected_columns if col not in self.df.columns]
            print(f"[DEBUG] 无效列: {invalid_columns}")
            # 自动取消无效列的勾选状态
            for col in invalid_columns:
                if col in self.column_vars:
                    self.column_vars[col].set(0)
            # 更新"已勾选数据"显示
            
            # 如果所有选中的列都无效
            if not valid_columns:
                print("[ERROR] 所有选中的列都无效")
                messagebox.showwarning("警告", "当前选中的列在新文件中不存在！请重新勾选数据列。")
                return
        
        # 使用过滤后的有效列
        selected_columns = valid_columns
        
        if not selected_columns:
            print("[ERROR] 没有选中任何列")
            messagebox.showwarning("警告", "请至少选择一列数据！")
            return
        
        if len(selected_columns) > 20:
            result = messagebox.askyesno("确认", f"您选择了 {len(selected_columns)} 列数据，图表可能会比较拥挤，是否继续？")
            if not result:
                return
        
        try:
            # 清空图表
            self.ax.clear()

            # 清除旧的标注（重要：每次生成新图表时都要清除）
            if hasattr(self, 'annotation') and self.annotation is not None:
                try:
                    self.annotation.remove()
                except:
                    pass
                self.annotation = None
            
            # 【新增】清空旧的图例面板和线条对象
            self._clear_legend_panel()
            
            # 【修复】处理中文时间格式 (如 "7时51分58秒" -> datetime)
            if not pd.api.types.is_datetime64_any_dtype(self.df[self.time_column]):
                import re
                def parse_chinese_time(t):
                    if pd.isna(t):
                        return pd.NaT
                    t = str(t)
                    # 匹配 "7时51分58秒" 格式
                    match = re.match(r'(\d{1,2})时(\d{1,2})分(\d{1,2})秒', t)
                    if match:
                        h, m, s = match.groups()
                        try:
                            return pd.to_datetime(f"2024-01-01 {int(h):02d}:{int(m):02d}:{int(s):02d}")
                        except:
                            return pd.NaT
                    # 如果不是中文格式，尝试直接转换
                    try:
                        return pd.to_datetime(t)
                    except:
                        return pd.NaT
                
                self.df[self.time_column] = self.df[self.time_column].apply(parse_chinese_time)
                # 删除转换失败的行
                self.df = self.df.dropna(subset=[self.time_column])
            
            # 动态标记点密度：基于总时间跨度（分钟）
            total_minutes = (self.df[self.time_column].iloc[-1] - self.df[self.time_column].iloc[0]).total_seconds() / 60
            if total_minutes <= 0:
                total_minutes = 1
            # 目标标记点数量约15个，且保证不少于1
            marker_interval = max(1, int(total_minutes / 15))

            # 使用matplotlib的颜色循环生成颜色，避免依赖numpy
            color_cycle = plt.rcParams['axes.prop_cycle'].by_key()['color']

            # 准备用于缓存的完整绘制数据（确保与折线完全一致）
            cached_time_data = self.df[self.time_column].tolist()
            cached_column_data = {col: self.df[col].tolist() for col in selected_columns}

            # 【修复】跟踪非数值列，用于提示用户（但不跳过）
            non_numeric_columns = []
            plot_idx = 0  # 实际绘制的列索引

            for col in selected_columns:
                # 【修复】检查列是否有有效数据
                col_series = self.df[col]
                valid_count = col_series.notna().sum()
                
                if valid_count == 0:
                    # 列为空或全为NaN，跳过
                    print(f"[WARNING] 列 '{col}' 无有效数据，已跳过绘制")
                    continue
                
                # 检查是否是数值类型
                is_numeric = pd.api.types.is_numeric_dtype(col_series)
                if not is_numeric:
                    # 尝试转换为数值类型
                    try:
                        numeric_data = pd.to_numeric(col_series, errors='coerce')
                        numeric_valid = numeric_data.notna().sum()
                        if numeric_valid > 0:
                            # 可以转换为数值，使用转换后的数据
                            self.df[col] = numeric_data
                            print(f"[DEBUG] 列 '{col}' 已转换为数值类型，有效数据={numeric_valid}")
                        else:
                            # 无法转换为数值，将其作为分类数据处理
                            # 将唯一值映射为数字
                            unique_values = col_series.dropna().unique()
                            value_map = {v: i for i, v in enumerate(unique_values)}
                            self.df[f'{col}_numeric'] = col_series.map(value_map)
                            self.df[col] = self.df[f'{col}_numeric']  # 用数值映射替换原列
                            non_numeric_columns.append((col, unique_values, value_map))
                            print(f"[DEBUG] 列 '{col}' 作为分类数据处理，唯一值数量={len(unique_values)}")
                    except Exception as e:
                        print(f"[DEBUG] 列 '{col}' 处理异常: {e}")
                
                # 从颜色循环中获取颜色（使用实际绘制索引）
                color = color_cycle[plot_idx % len(color_cycle)]
                plot_idx += 1
                
                # 【新增】将颜色转换为hex格式用于图例
                if isinstance(color, (tuple, list)):
                    hex_color = '#{:02x}{:02x}{:02x}'.format(
                        int(color[0] * 255), 
                        int(color[1] * 255), 
                        int(color[2] * 255)
                    )
                else:
                    # 如果是颜色名称，尝试获取RGB值
                    try:
                        import matplotlib.colors as mcolors
                        rgb = mcolors.to_rgb(color)
                        hex_color = '#{:02x}{:02x}{:02x}'.format(
                            int(rgb[0] * 255), 
                            int(rgb[1] * 255), 
                            int(rgb[2] * 255)
                        )
                    except:
                        hex_color = '#1f77b4'  # 默认蓝色

                # 仅对时间戳按分钟等距采样，不做重采样聚合
                data_to_plot = self.df.copy()
                # 将时间戳解析为datetime（如果尚未解析）
                if not pd.api.types.is_datetime64_any_dtype(data_to_plot[self.time_column]):
                    data_to_plot[self.time_column] = pd.to_datetime(data_to_plot[self.time_column])
                data_to_plot['time_min'] = data_to_plot[self.time_column].dt.floor('min')
                # 按分钟去重，保留每分钟首个出现值
                data_to_plot = data_to_plot.drop_duplicates(subset='time_min', keep='first')
                # 按分钟等距采样：从起点开始，每隔marker_interval分钟取一个点
                start_time = data_to_plot['time_min'].min()
                sample_times = pd.date_range(start=start_time, periods=15, freq=f'{marker_interval}min')
                sampled = data_to_plot[data_to_plot['time_min'].isin(sample_times)].copy()
                # 若采样不足3个，回退到原始数据并直接按索引等距取点
                if len(sampled) < 3:
                    sampled = self.df.copy()
                    sampled['time_min'] = sampled[self.time_column].dt.floor('min')
                    sampled = sampled.drop_duplicates(subset='time_min', keep='first')
                    idx_step = max(1, len(sampled) // 15)
                    sampled = sampled.iloc[::idx_step]
                # 绘制原始全部点（仅线），并叠加等距标记点
                line_obj = self.ax.plot(
                    self.df[self.time_column],
                    self.df[col],
                    linewidth=1.5,
                    color=color,
                    alpha=0.7,
                    label=col,
                    zorder=2
                )[0]  # 【修改】保存线条对象
                marker_obj = self.ax.plot(
                    sampled[self.time_column],
                    sampled[col],
                    marker='o',
                    markersize=4,
                    markeredgecolor='white',
                    markeredgewidth=0.8,
                    color=color,
                    linestyle='None',
                    zorder=3
                )[0]  # 【修改】保存标记点对象
                
                # 【新增】保存线条对象和可见性状态
                self.line_objects[col] = (line_obj, marker_obj)
                self.line_visibility[col] = True
                
                # 【新增】创建图例项
                self._create_legend_item(col, hex_color, visible=True)
            
            # 【修复】如果有非数值列被转换，提示用户
            if non_numeric_columns:
                info_msg = "以下数据列已转换为分类数值显示：\n"
                for col, unique_values, value_map in non_numeric_columns:
                    # 只显示前几个唯一值，避免提示过长
                    sample_values = list(unique_values[:5])
                    if len(unique_values) > 5:
                        sample_values.append(f"... 等{len(unique_values)}个值")
                    info_msg += f"\n• {col}: {', '.join(str(v) for v in sample_values)}"
                print(f"[INFO] {info_msg}")
                
                # 在状态栏显示提示（不弹出对话框，避免打断用户）
                if hasattr(self, 'status_var'):
                    self.status_var.set(f"提示: {len(non_numeric_columns)} 个非数值列已转换为分类显示")
            
            # 【修复】如果没有绘制任何列，显示错误提示
            if plot_idx == 0:
                self.ax.clear()
                self.ax.text(0.5, 0.5, "没有可绘制的数据\n所选列均无有效数据", 
                           ha='center', va='center', fontsize=14, color='gray',
                           transform=self.ax.transAxes)
                self.canvas_plot.draw()
                if hasattr(self, 'status_var'):
                    self.status_var.set("错误: 所选列均无有效数据，无法绘制")
                return
            
            # 设置图表标题和标签
            self.ax.set_title(self.title_var.get(), fontsize=14, fontweight='bold')
            self.ax.set_xlabel(self.xlabel_var.get(), fontsize=12)
            self.ax.set_ylabel(self.ylabel_var.get(), fontsize=12)
            
            # 时间轴对齐与格式化
            # 强制确保时间列为 datetime 类型
            if not pd.api.types.is_datetime64_any_dtype(self.df[self.time_column]):
                self.df[self.time_column] = pd.to_datetime(self.df[self.time_column], errors='coerce')
                # 删除转换失败的行
                self.df = self.df.dropna(subset=[self.time_column])
            
            # 确保 self.df[self.time_column] 是 datetime 类型
            if pd.api.types.is_datetime64_any_dtype(self.df[self.time_column]):
                # 根据数据点数量智能选择时间刻度策略
                num_points = len(self.df)
                time_range = (self.df[self.time_column].iloc[-1] - self.df[self.time_column].iloc[0]).total_seconds() / 60
                
                if num_points <= 20:
                    # 数据点较少时，显示所有时间点
                    self.ax.set_xticks(self.df[self.time_column])
                    if time_range <= 60:
                        self.ax.xaxis.set_major_formatter(mdates.DateFormatter('%H:%M:%S'))
                    else:
                        self.ax.xaxis.set_major_formatter(mdates.DateFormatter('%H:%M'))
                else:
                    # 数据点较多时，根据时间范围智能选择刻度间隔
                    if time_range >= 1200:  # 20小时以上（接近24小时），使用1小时间隔
                        locator = mdates.HourLocator(interval=1)
                    elif time_range >= 480:  # 8小时以上，使用2小时间隔
                        locator = mdates.HourLocator(interval=2)
                    elif time_range >= 240:  # 4小时以上，使用30分钟间隔
                        locator = mdates.MinuteLocator(byminute=[0, 30])
                    else:  # 短时间范围，使用自动刻度
                        locator = mdates.AutoDateLocator()
                    
                    self.ax.xaxis.set_major_locator(locator)
                    if time_range <= 60:
                        self.ax.xaxis.set_major_formatter(mdates.DateFormatter('%H:%M:%S'))
                    else:
                        self.ax.xaxis.set_major_formatter(mdates.DateFormatter('%H:%M'))
                
                self.ax.tick_params(axis='x', rotation=45)
                self.figure.autofmt_xdate()
            
            # 严格按数据范围设置刻度范围
            self.ax.set_xlim(self.df[self.time_column].iloc[0], self.df[self.time_column].iloc[-1])
            self.ax.margins(x=0)
            
            # 【修复】自动调整 Y 轴范围，确保所有数据可见
            # 获取所有选中列的数据范围
            all_values = []
            for col in selected_columns:
                col_data = self.df[col].dropna()
                if len(col_data) > 0:
                    # 尝试将数据转换为数值类型
                    try:
                        # 先尝试直接转换为数值
                        numeric_data = pd.to_numeric(col_data, errors='coerce')
                        valid_data = numeric_data.dropna()
                        if len(valid_data) > 0:
                            all_values.extend(valid_data.tolist())
                    except Exception as e:
                        print(f"[WARNING] 列 '{col}' 数据无法转换为数值: {e}")
                        continue
            
            if all_values:
                import numpy as np
                y_min = min(all_values)
                y_max = max(all_values)
                y_range = y_max - y_min
                
                # 添加 5% 的边距，确保数据不会贴边
                y_margin = y_range * 0.05 if y_range > 0 else 1
                self.ax.set_ylim(y_min - y_margin, y_max + y_margin)
                print(f"[DEBUG] Y轴范围: {y_min - y_margin:.2f} 到 {y_max + y_margin:.2f}")
            
            # 添加网格
            self.ax.grid(True, alpha=0.3, linestyle='--')
            
            # 添加图例
            # 【修改】不再使用 matplotlib 的内置图例，使用自定义图例面板
            # if len(selected_columns) <= 15:
            #     self.ax.legend(bbox_to_anchor=(1.02, 1), loc='upper left', fontsize=8)
            # else:
            #     self.ax.legend_.remove()
            #     self.show_legend_popup(selected_columns)
            
            # 调整布局（为图例面板留出空间）
            self.figure.tight_layout()
            
            # 缓存绘制的数据，用于点击查询
            # 【修复】只缓存非空值，确保索引对应正确
            cache_data = {}
            for col in selected_columns:
                valid_mask = self.df[col].notna()
                cache_data[col] = {
                    'values': self.df.loc[valid_mask, col].tolist(),
                    'times': self.df.loc[valid_mask, self.time_column].tolist()
                }
            
            self.plot_data_cache = {
                'columns': selected_columns,
                'time_data': self.df[self.time_column].tolist(),
                'column_data': {col: self.df[col].tolist() for col in selected_columns},
                'valid_data': cache_data  # 新增：非空数据
            }
            self.last_selected_columns = selected_columns
            
            # 刷新画布
            self.canvas_plot.draw()
            
            # 启用保存按钮
            self.save_btn.config(state=tk.NORMAL)
            
            # 保存原始坐标轴范围用于缩放限制
            self.original_xlim = self.ax.get_xlim()
            self.original_ylim = self.ax.get_ylim()
            
            # 重置缩放状态
            self.x_zoom_factor = 1.0
            self.y_zoom_factor = 1.0
            self.zoom_center = None
            
        except Exception as e:
            import traceback
            messagebox.showerror("错误", f"生成图表失败：{str(e)}\n\n{traceback.format_exc()}")
            
    def on_scroll(self, event):
        """处理鼠标滚轮缩放事件 - 简化版"""
        try:
            # 检查是否有数据
            if self.df is None:
                return
            
            # 获取当前坐标轴范围
            xlim = self.ax.get_xlim()
            ylim = self.ax.get_ylim()
            
            # 缩放倍率
            zoom_step = 1.1
            
            # 确定缩放因子
            if event.delta > 0 or event.num == 4:
                # 向上滚动，放大
                scale = 1 / zoom_step
            elif event.delta < 0 or event.num == 5:
                # 向下滚动，缩小
                scale = zoom_step
            else:
                return
            
            # 获取鼠标在数据坐标系中的位置
            mouse_x = event.xdata
            mouse_y = event.ydata
            
            # 如果无法获取鼠标位置，使用中心缩放
            if mouse_x is None:
                mouse_x = (xlim[0] + xlim[1]) / 2
            if mouse_y is None:
                mouse_y = (ylim[0] + ylim[1]) / 2
            
            # 计算新的坐标轴范围 - 围绕鼠标位置缩放
            # X轴缩放 - 保持鼠标位置不变
            x_range = xlim[1] - xlim[0]
            new_x_range = x_range * scale
            x_ratio = (mouse_x - xlim[0]) / x_range if x_range != 0 else 0.5
            new_xlim = [
                mouse_x - x_ratio * new_x_range,
                mouse_x + (1 - x_ratio) * new_x_range
            ]
            
            # Y轴缩放 - 保持鼠标位置不变
            y_range = ylim[1] - ylim[0]
            new_y_range = y_range * scale
            y_ratio = (mouse_y - ylim[0]) / y_range if y_range != 0 else 0.5
            new_ylim = [
                mouse_y - y_ratio * new_y_range,
                mouse_y + (1 - y_ratio) * new_y_range
            ]
            
            # 应用缩放限制
            if self.original_xlim is not None:
                # 限制X轴缩放范围
                if pd.api.types.is_datetime64_any_dtype(self.df[self.time_column]):
                    try:
                        original_range = (self.original_xlim[1] - self.original_xlim[0]).total_seconds()
                        current_range = (new_xlim[1] - new_xlim[0]).total_seconds()
                        
                        min_range = original_range / 100  # 允许放大到原始范围的 1/100，更精细的缩放
                        max_range = original_range * 10  # 允许放大到原始范围的 10 倍
                        
                        if current_range < min_range or current_range > max_range:
                            return # 超出缩放限制，不执行
                    except:
                        pass # 如果时间计算失败，允许缩放
            
            # 应用新的坐标轴范围
            self.ax.set_xlim(new_xlim)
            self.ax.set_ylim(new_ylim)
            
            # 动态更新时间精度 - 必须在设置xlim之后立即执行
            self.update_time_precision(new_xlim)
            
            # 动态调整X轴刻度，保持尽可能多的刻度点
            if pd.api.types.is_datetime64_any_dtype(self.df[self.time_column]):
                try:
                    current_range = (new_xlim[1] - new_xlim[0]).total_seconds()
                    original_range = (self.original_xlim[1] - self.original_xlim[0]).total_seconds()
                    
                    # 根据当前缩放范围，智能选择刻度间隔，保持尽可能多的刻度点
                    # 目标：显示15-25个刻度点，确保放大后能看到更细粒度的时间
                    
                    # 计算理想刻度数量：范围越小，刻度越少但精度越高
                    # 优化：使用反向映射，确保放大时有足够的刻度显示
                    if current_range >= original_range * 0.5:
                        target_ticks = 20  # 大范围：20个点
                    elif current_range >= original_range * 0.3:
                        target_ticks = 18  # 中等范围：18个点
                    elif current_range >= original_range * 0.2:
                        target_ticks = 15  # 较小范围：15个点
                    elif current_range >= original_range * 0.1:
                        target_ticks = 12  # 小范围：12个点
                    elif current_range >= original_range * 0.05:
                        target_ticks = 10  # 很小范围：10个点
                    elif current_range >= original_range * 0.01:
                        target_ticks = 8   # 极小范围：8个点
                    else:
                        target_ticks = 6   # 超小范围：6个点
                    
                    # 计算刻度间隔
                    target_interval = current_range / target_ticks
                    
                    # 智能选择最接近的标准间隔（根据时间范围选择合适的精度）
                    if target_interval >= 7200:  # 2小时以上
                        interval_hours = max(1, round(target_interval / 3600))
                        # 使用小时级刻度
                        locator = mdates.HourLocator(interval=interval_hours)
                        formatter = mdates.DateFormatter('%H:%M')
                    elif target_interval >= 300:  # 5分钟到2小时之间
                        interval_minutes = max(1, round(target_interval / 60))
                        # 使用分钟级刻度（更精细的间隔）
                        if interval_minutes <= 1:
                            interval_minutes = 1
                        elif interval_minutes <= 2:
                            interval_minutes = 2
                        elif interval_minutes <= 5:
                            interval_minutes = 5
                        elif interval_minutes <= 10:
                            interval_minutes = 10
                        elif interval_minutes <= 15:
                            interval_minutes = 15
                        elif interval_minutes <= 30:
                            interval_minutes = 30
                        else:
                            interval_minutes = 60
                        locator = mdates.MinuteLocator(byminute=range(0, 60, interval_minutes))
                        formatter = mdates.DateFormatter('%H:%M')
                    else:  # 小于5分钟，使用秒级刻度
                        interval_seconds = max(1, int(round(target_interval)))
                        # 使用秒级刻度（更精细的间隔）
                        if interval_seconds <= 1:
                            interval_seconds = 1
                        elif interval_seconds <= 2:
                            interval_seconds = 2
                        elif interval_seconds <= 5:
                            interval_seconds = 5
                        elif interval_seconds <= 10:
                            interval_seconds = 10
                        elif interval_seconds <= 15:
                            interval_seconds = 15
                        elif interval_seconds <= 30:
                            interval_seconds = 30
                        else:
                            interval_seconds = 60
                        locator = mdates.SecondLocator(interval=interval_seconds)
                        formatter = mdates.DateFormatter('%H:%M:%S')
                    
                    # 强制刷新X轴刻度
                    self.ax.xaxis.set_major_locator(locator)
                    self.ax.xaxis.set_major_formatter(formatter)
                    
                    # 自动调整刻度标签格式，防止重叠
                    self.ax.figure.autofmt_xdate(rotation=45, ha='right')
                    
                    # 动态调整Y轴刻度，根据显示的数据范围动态调整
                    try:
                        # 获取当前X轴范围内的数据
                        mask = (self.df[self.time_column] >= new_xlim[0]) & (self.df[self.time_column] <= new_xlim[1])
                        visible_data = self.df[mask]
                        
                        if len(visible_data) > 0:
                            # 计算可见数据的Y轴范围
                            for col in visible_data.columns:
                                if col != self.time_column and pd.api.types.is_numeric_dtype(visible_data[col]):
                                    y_data = visible_data[col].dropna()
                                    if len(y_data) > 0:
                                        y_min = y_data.min()
                                        y_max = y_data.max()
                                        y_range = y_max - y_min
                                        
                                        # 根据Y轴范围智能选择刻度间隔
                                        if y_range > 0:
                                            # 目标显示8-12个Y轴刻度
                                            target_y_ticks = 10
                                            y_interval = y_range / target_y_ticks
                                            
                                            # 选择合适的精度
                                            if y_range >= 1000:
                                                y_step = 100
                                            elif y_range >= 500:
                                                y_step = 50
                                            elif y_range >= 100:
                                                y_step = 10
                                            elif y_range >= 50:
                                                y_step = 5
                                            elif y_range >= 10:
                                                y_step = 2
                                            else:
                                                y_step = 1
                                            
                                            # 设置Y轴刻度
                                            import numpy as np
                                            y_ticks = np.arange(y_min - y_step, y_max + y_step + 0.1, y_step)
                                            self.ax.set_yticks(y_ticks)
                    except Exception as e:
                        pass  # Y轴调整失败，保持原样
                        
                except Exception as e:
                    pass  # 如果计算失败，保持原样
            
            # 重新绘制图表
            self.canvas_plot.draw_idle()
            
        except Exception as e:
            # 静默处理错误，避免显示大量错误信息
            pass
            
    def on_canvas_scroll(self, event):
        """处理左侧Canvas的滚轮事件 - 用于上下滚动列表"""
        try:
            # Windows系统: event.delta
            # Linux系统: event.num
            if event.delta:
                # Windows: 向上滚动为正数，向下为负数
                scroll_units = -1 * int(event.delta / 120)
            else:
                # Linux
                if event.num == 4:
                    scroll_units = -1 # 向上滚动
                elif event.num == 5:
                    scroll_units = 1 # 向下滚动
                else:
                    return
            
            # 执行滚动
            self.canvas.yview_scroll(scroll_units, "units")
            
        except Exception as e:
            # 静默处理错误
            pass
    
    def _on_legend_scroll(self, event):
        """处理图例面板的滚轮事件"""
        try:
            if event.delta:
                scroll_units = -1 * int(event.delta / 120)
            else:
                if event.num == 4:
                    scroll_units = -1
                elif event.num == 5:
                    scroll_units = 1
                else:
                    return
            
            self.legend_canvas.yview_scroll(scroll_units, "units")
        except:
            pass
    
    def _create_legend_item(self, col, color, visible=True):
        """创建单个图例项
        
        Args:
            col: 列名
            color: 颜色（hex格式）
            visible: 是否可见
        """
        # 创建图例项Frame
        item_frame = ttk.Frame(self.legend_scrollable_frame)
        item_frame.pack(fill=tk.X, padx=2, pady=2)
        
        # 颜色标记
        color_canvas = tk.Canvas(item_frame, width=20, height=20, highlightthickness=0)
        color_canvas.pack(side=tk.LEFT, padx=(2, 5))
        color_canvas.create_rectangle(2, 2, 18, 18, fill=color, outline="gray")
        
        # 列名标签（可点击切换显示/隐藏）
        # 截断过长的名称
        display_name = col if len(col) <= 20 else col[:17] + "..."
        
        name_label = ttk.Label(item_frame, text=display_name, cursor="hand2")
        name_label.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        # 存储完整列名
        name_label.full_name = col
        
        # 隐藏按钮（眼睛图标）
        visibility_btn = ttk.Button(item_frame, text="👁", width=3, 
                                    command=lambda: self._toggle_line_visibility(col))
        visibility_btn.pack(side=tk.LEFT, padx=2)
        
        # 删除按钮（叉号）
        delete_btn = ttk.Button(item_frame, text="✕", width=3,
                               command=lambda: self._remove_line(col))
        delete_btn.pack(side=tk.LEFT, padx=2)
        
        # 绑定点击事件（切换显示/隐藏）
        def toggle_visibility(event):
            self._toggle_line_visibility(col)
        
        name_label.bind("<Button-1>", toggle_visibility)
        color_canvas.bind("<Button-1>", toggle_visibility)
        
        # 存储引用
        self.legend_items[col] = {
            'frame': item_frame,
            'color_canvas': color_canvas,
            'name_label': name_label,
            'visibility_btn': visibility_btn,
            'delete_btn': delete_btn,
            'color': color
        }
        
        # 根据可见性设置初始状态
        if not visible:
            name_label.configure(foreground='gray')
            visibility_btn.configure(text="👁‍🗨")
        
        return item_frame
    
    def _toggle_line_visibility(self, col):
        """切换线条的显示/隐藏
        
        Args:
            col: 列名
        """
        if col not in self.line_objects:
            return
        
        # 切换状态
        current_visibility = self.line_visibility.get(col, True)
        new_visibility = not current_visibility
        self.line_visibility[col] = new_visibility
        
        # 更新线条可见性
        line_obj, marker_obj = self.line_objects[col]
        if line_obj:
            line_obj.set_visible(new_visibility)
        if marker_obj:
            marker_obj.set_visible(new_visibility)
        
        # 更新图例项外观
        if col in self.legend_items:
            item = self.legend_items[col]
            if new_visibility:
                item['name_label'].configure(foreground='black')
                item['visibility_btn'].configure(text="👁")
                # 恢复颜色
                item['color_canvas'].delete("all")
                item['color_canvas'].create_rectangle(2, 2, 18, 18, 
                                                      fill=item['color'], outline="gray")
            else:
                item['name_label'].configure(foreground='gray')
                item['visibility_btn'].configure(text="👁‍🗨")
                # 变灰
                item['color_canvas'].delete("all")
                item['color_canvas'].create_rectangle(2, 2, 18, 18, 
                                                      fill="lightgray", outline="gray")
        
        # 刷新图表
        self.canvas_plot.draw_idle()
        
        print(f"[INFO] 切换线条 '{col}' 可见性: {new_visibility}")
    
    def _remove_line(self, col):
        """删除线条（从图表中移除）
        
        Args:
            col: 列名
        """
        if col not in self.line_objects:
            return
        
        # 从图表中移除线条
        line_obj, marker_obj = self.line_objects[col]
        if line_obj:
            try:
                line_obj.remove()
            except:
                pass
        if marker_obj:
            try:
                marker_obj.remove()
            except:
                pass
        
        # 删除存储的对象
        del self.line_objects[col]
        del self.line_visibility[col]
        
        # 从图例面板移除
        if col in self.legend_items:
            item = self.legend_items[col]
            item['frame'].destroy()
            del self.legend_items[col]
        
        # 从缓存中移除
        if col in self.plot_data_cache.get('columns', []):
            self.plot_data_cache['columns'].remove(col)
        if col in self.plot_data_cache.get('column_data', {}):
            del self.plot_data_cache['column_data'][col]
        if col in self.plot_data_cache.get('valid_data', {}):
            del self.plot_data_cache['valid_data'][col]
        
        # 取消勾选
        if col in self.column_vars:
            self.column_vars[col].set(0)
        
        # 刷新图表
        self.canvas_plot.draw_idle()
        
        print(f"[INFO] 已删除线条 '{col}'")
    
    def _clear_legend_panel(self):
        """清空图例面板"""
        for col, item in list(self.legend_items.items()):
            try:
                item['frame'].destroy()
            except:
                pass
        self.legend_items.clear()
        self.line_objects.clear()
        self.line_visibility.clear()

    def update_time_precision(self, new_xlim):
        """根据新的X轴范围动态更新时间精度"""
        try:
            if self.df is None or self.time_column is None:
                return
                
            if not pd.api.types.is_datetime64_any_dtype(self.df[self.time_column]):
                return
                
            # 检查 new_xlim 是否为时间类型，如果是数值则转换
            try:
                # 尝试直接计算时间范围（秒）
                current_range = (new_xlim[1] - new_xlim[0]).total_seconds()
                original_range = (self.original_xlim[1] - self.original_xlim[0]).total_seconds()
            except (AttributeError, TypeError):
                # 如果不是时间类型，说明是 matplotlib 的数值时间戳，需要转换
                import matplotlib.dates as mdates
                try:
                    # 将 matplotlib 数值时间戳转换为 datetime
                    new_xlim_datetime = [mdates.num2date(x) for x in new_xlim]
                    current_range = (new_xlim_datetime[1] - new_xlim_datetime[0]).total_seconds()
                    original_range_datetime = [mdates.num2date(x) for x in self.original_xlim]
                    original_range = (original_range_datetime[1] - original_range_datetime[0]).total_seconds()
                    print(f"[DEBUG update_time_precision] 数值时间戳已转换: {current_range:.1f}秒")
                except Exception as e:
                    # 如果转换失败，直接返回
                    print(f"[DEBUG update_time_precision] 时间戳转换失败，跳过更新: {e}")
                    return
            
            # 计算缩放比例
            if original_range > 0:
                zoom_ratio = current_range / original_range
            else:
                zoom_ratio = 1.0
                
            print(f"[DEBUG update_time_precision] 当前范围: {current_range:.1f}秒, 原始范围: {original_range:.1f}秒, 缩放比例: {zoom_ratio:.3f}")
            
            # 根据时间范围直接选择合适的刻度间隔
            if current_range >= 7200:  # 2小时以上：使用小时级刻度
                interval_hours = max(1, int(current_range / 3600 / 15))  # 显示约15个刻度
                if interval_hours == 0:
                    interval_hours = 1
                locator = mdates.HourLocator(interval=interval_hours)
                formatter = mdates.DateFormatter('%H:%M')
                print(f"[DEBUG update_time_precision] 使用小时级刻度，间隔: {interval_hours}小时")
                
            elif current_range >= 600:  # 10分钟到2小时：使用分钟级刻度
                # 计算合适的分钟间隔，目标显示约10-15个刻度
                target_ticks = 12
                interval_minutes = max(1, int(current_range / 60 / target_ticks))
                # 标准化间隔到常用值：1, 2, 5, 10, 15, 30分钟
                if interval_minutes <= 1:
                    interval_minutes = 1
                elif interval_minutes <= 2:
                    interval_minutes = 2
                elif interval_minutes <= 5:
                    interval_minutes = 5
                elif interval_minutes <= 10:
                    interval_minutes = 10
                elif interval_minutes <= 15:
                    interval_minutes = 15
                elif interval_minutes <= 30:
                    interval_minutes = 30
                else:
                    interval_minutes = 60
                locator = mdates.MinuteLocator(byminute=range(0, 60, interval_minutes))
                formatter = mdates.DateFormatter('%H:%M')
                print(f"[DEBUG update_time_precision] 使用分钟级刻度，间隔: {interval_minutes}分钟")
                
            elif current_range >= 120:  # 2分钟到10分钟：使用分钟级精细刻度
                # 使用1-5分钟的间隔
                target_ticks = 10
                interval_minutes = max(1, int(current_range / 60 / target_ticks))
                if interval_minutes <= 1:
                    interval_minutes = 1
                elif interval_minutes <= 2:
                    interval_minutes = 2
                else:
                    interval_minutes = 5
                locator = mdates.MinuteLocator(byminute=range(0, 60, interval_minutes))
                formatter = mdates.DateFormatter('%H:%M')
                print(f"[DEBUG update_time_precision] 使用分钟级精细刻度，间隔: {interval_minutes}分钟")
                
            elif current_range >= 30:  # 30秒到2分钟：使用秒级刻度
                # 使用10-30秒的间隔
                target_ticks = 8
                interval_seconds = max(10, int(current_range / target_ticks))
                # 标准化间隔到常用值：10, 15, 30秒
                if interval_seconds <= 10:
                    interval_seconds = 10
                elif interval_seconds <= 15:
                    interval_seconds = 15
                elif interval_seconds <= 30:
                    interval_seconds = 30
                else:
                    interval_seconds = 60
                locator = mdates.SecondLocator(interval=interval_seconds)
                formatter = mdates.DateFormatter('%H:%M:%S')
                print(f"[DEBUG update_time_precision] 使用秒级刻度，间隔: {interval_seconds}秒")
                
            else:  # 小于30秒：使用秒级精细刻度
                # 使用1-10秒的间隔
                target_ticks = 6
                interval_seconds = max(1, int(current_range / target_ticks))
                # 标准化间隔到常用值：1, 2, 5, 10秒
                if interval_seconds <= 1:
                    interval_seconds = 1
                elif interval_seconds <= 2:
                    interval_seconds = 2
                elif interval_seconds <= 5:
                    interval_seconds = 5
                elif interval_seconds <= 10:
                    interval_seconds = 10
                else:
                    interval_seconds = 15
                locator = mdates.SecondLocator(interval=interval_seconds)
                formatter = mdates.DateFormatter('%H:%M:%S')
                print(f"[DEBUG update_time_precision] 使用秒级精细刻度，间隔: {interval_seconds}秒")
            
            # 应用新的刻度设置
            self.ax.xaxis.set_major_locator(locator)
            self.ax.xaxis.set_major_formatter(formatter)
            
            # 自动调整刻度标签格式，防止重叠
            if current_range < 120:  # 小于2分钟，使用45度旋转
                self.ax.figure.autofmt_xdate(rotation=45, ha='right')
            else:
                self.ax.figure.autofmt_xdate(rotation=0, ha='center')
                
            print(f"[DEBUG update_time_precision] 时间精度更新完成")
            
        except Exception as e:
            print(f"[ERROR update_time_precision] 更新时间精度时出错: {str(e)}")
            import traceback
            traceback.print_exc()
            
    def reset_to_original_view(self, *args):
        """复位到生成折线图时的初始视图状态"""
        try:
            # 检查是否已生成图表
            if self.original_xlim is None or self.original_ylim is None:
                print("[INFO reset] 尚未生成图表，无法复位")
                return
            
            # 恢复原始坐标轴范围
            self.ax.set_xlim(self.original_xlim)
            self.ax.set_ylim(self.original_ylim)
            
            # 如果是时间轴，恢复时间精度到初始状态
            if self.time_column and pd.api.types.is_datetime64_any_dtype(self.df[self.time_column]):
                self.update_time_precision(self.original_xlim)
            
            # 重新绘制图表
            self.canvas_plot.draw()
            
            print(f"[INFO reset] 已恢复到初始视图")
            print(f"  X轴范围: {self.original_xlim}")
            print(f"  Y轴范围: {self.original_ylim}")
            
        except Exception as e:
            print(f"[ERROR reset] 复位失败: {str(e)}")
            import traceback
            traceback.print_exc()
    
    def on_mouse_press(self, event):
        """处理鼠标按下事件 - 开始拖动"""
        try:
            # 只响应左键
            if event.button != 1:
                return
            
            # 检查是否点击在图表区域内
            if event.inaxes != self.ax:
                return
            
            # 记录拖动起始位置和X轴范围
            self.drag_start = event.x
            self.drag_xlim = self.ax.get_xlim()
            
        except Exception as e:
            print(f"[ERROR mouse_press] {str(e)}")
    
    def on_mouse_release(self, event):
        """处理鼠标释放事件 - 结束拖动"""
        try:
            # 清除拖动状态
            self.drag_start = None
            self.drag_xlim = None
            
        except Exception as e:
            print(f"[ERROR mouse_release] {str(e)}")
    
    def on_mouse_drag(self, event):
        """处理鼠标拖动事件 - 平移图表"""
        try:
            # 检查是否处于拖动状态
            if self.drag_start is None or self.drag_xlim is None:
                return
            
            # 只在图表区域内响应拖动
            if event.inaxes != self.ax:
                return
            
            # 计算拖动距离
            dx = event.x - self.drag_start
            
            # 计算平移量（将像素距离转换为数据坐标）
            xlim = self.ax.get_xlim()
            xlim_range = xlim[1] - xlim[0]
            
            # 获取图表宽度的像素值
            # 获取图表宽度的像素值 - 使用 tkinter 组件宽度
            width_pixels = self.canvas_plot_widget.winfo_width()
            
            # 计算平移比例
            if width_pixels > 0:
                shift_ratio = dx / width_pixels
                data_shift = xlim_range * shift_ratio
                
                # 计算新的X轴范围
                new_xlim = [
                    self.drag_xlim[0] - data_shift,
                    self.drag_xlim[1] - data_shift
                ]
                
                # 限制X轴范围不超过原始范围
                if self.original_xlim is not None:
                    if new_xlim[0] < self.original_xlim[0]:
                        new_xlim = [
                            self.original_xlim[0],
                            self.original_xlim[0] + xlim_range
                        ]
                    elif new_xlim[1] > self.original_xlim[1]:
                        new_xlim = [
                            self.original_xlim[1] - xlim_range,
                            self.original_xlim[1]
                        ]
                
                # 应用新的X轴范围
                self.ax.set_xlim(new_xlim)
                
                # 更新时间精度（如果需要）
                if self.time_column and pd.api.types.is_datetime64_any_dtype(self.df[self.time_column]):
                    self.update_time_precision(new_xlim)
                
                # 重新绘制图表
                self.canvas_plot.draw()
            
        except Exception as e:
            print(f"[ERROR mouse_drag] {str(e)}")
            import traceback
            traceback.print_exc()
    
    def on_plot_click(self, event):
        """处理图表点击事件 - 显示数据点信息"""
        try:
            # 检查是否点击在图表区域内
            if event.inaxes != self.ax:
                return

            # 检查是否有缓存的绘图数据
            if not self.plot_data_cache or 'columns' not in self.plot_data_cache:
                return

            selected_columns = self.plot_data_cache['columns']

            # 获取点击位置
            click_x = event.xdata
            click_y = event.ydata

            if click_x is None or click_y is None:
                return

            # 关键修复：使用matplotlib内置方法将点击的x坐标转换为datetime
            import matplotlib.dates as mdates
            try:
                click_datetime = mdates.num2date(click_x)
                # 转换为不带时区的datetime（naive datetime）
                if hasattr(click_datetime, 'tzinfo') and click_datetime.tzinfo is not None:
                    click_datetime = click_datetime.replace(tzinfo=None)
            except Exception as e:
                print(f"日期转换错误: {e}")
                return

            # 【修复】使用有效数据缓存进行查询
            valid_data = self.plot_data_cache.get('valid_data', {})
            
            if len(selected_columns) == 1:
                # 只勾选一列时，直接使用该列
                closest_col = selected_columns[0]
                
                if closest_col not in valid_data:
                    return
                
                col_data = valid_data[closest_col]
                times = col_data['times']
                values = col_data['values']
                
                # 找到最接近的时间点
                min_diff = float('inf')
                closest_time_idx = None
                
                for i, t in enumerate(times):
                    try:
                        if hasattr(t, 'to_pydatetime'):
                            t_datetime = t.to_pydatetime()
                            if hasattr(t_datetime, 'tzinfo') and t_datetime.tzinfo is not None:
                                t_datetime = t_datetime.replace(tzinfo=None)
                        else:
                            t_datetime = pd.to_datetime(t)
                        
                        time_diff = abs((t_datetime - click_datetime).total_seconds())
                        if time_diff < min_diff:
                            min_diff = time_diff
                            closest_time_idx = i
                            closest_time_val = t
                    except Exception as e:
                        continue
                
                if closest_time_idx is None:
                    return
                
                closest_val = values[closest_time_idx]
                
            else:
                # 勾选多列时，找到数值最接近的列
                best_col = None
                best_val = None
                best_time = None
                min_diff = float('inf')
                
                for col in selected_columns:
                    if col not in valid_data:
                        continue
                    
                    col_data = valid_data[col]
                    times = col_data['times']
                    values = col_data['values']
                    
                    # 找到最接近的时间点
                    for i, t in enumerate(times):
                        try:
                            if hasattr(t, 'to_pydatetime'):
                                t_datetime = t.to_pydatetime()
                                if hasattr(t_datetime, 'tzinfo') and t_datetime.tzinfo is not None:
                                    t_datetime = t_datetime.replace(tzinfo=None)
                            else:
                                t_datetime = pd.to_datetime(t)
                            
                            time_diff = abs((t_datetime - click_datetime).total_seconds())
                            val_diff = abs(values[i] - click_y) if not pd.isna(values[i]) else float('inf')
                            
                            # 综合距离：时间距离 + 数值距离
                            total_diff = time_diff / 60 + val_diff / 100  # 归一化
                            
                            if total_diff < min_diff:
                                min_diff = total_diff
                                best_col = col
                                best_val = values[i]
                                best_time = t
                        except Exception as e:
                            continue
                
                closest_col = best_col
                closest_val = best_val
                closest_time_val = best_time

            if closest_col is None or closest_val is None:
                return

            # 格式化时间显示
            time_val = closest_time_val
            if hasattr(time_val, 'strftime'):
                time_str = time_val.strftime('%Y-%m-%d %H:%M:%S')
            elif isinstance(time_val, str):
                time_str = time_val
            else:
                try:
                    time_str = str(pd.to_datetime(time_val))
                except:
                    time_str = str(time_val)

            # 准备显示文本
            info_text = f"时间: {time_str}\n列名: {closest_col}\n数值: {closest_val:.2f}"

            # 移除旧的标注
            if self.annotation is not None:
                self.annotation.remove()
                self.annotation = None

            # 在图表上添加标注 - 标注框固定在点击位置，不带箭头
            self.annotation = self.ax.annotate(
                info_text,
                xy=(click_x, click_y),
                xytext=(10, 10),
                textcoords='offset points',
                bbox=dict(
                    boxstyle='round,pad=0.5',
                    facecolor='yellow',
                    alpha=0.9,
                    edgecolor='black'
                ),
                fontsize=10,
                fontweight='bold',
                ha='left',
                va='bottom',
                zorder=10 # 确保标注在最上层
            )

            # 重新绘制
            self.canvas_plot.draw()
            
        except Exception as e:
            # 静默处理错误，避免影响用户体验
            pass
            
    def show_legend_popup(self, selected_columns):
        """显示独立图例窗口"""
        legend_window = tk.Toplevel(self.root)
        legend_window.title("图例")
        legend_window.geometry("300x500")
        
        # 【修复】设置为主窗口的临时窗口，防止被遮挡
        legend_window.transient(self.root)
        
        # 添加搜索框
        search_frame = ttk.Frame(legend_window)
        search_frame.pack(fill=tk.X, padx=10, pady=10)
        ttk.Label(search_frame, text="搜索:").pack(side=tk.LEFT)
        search_var = tk.StringVar()
        search_entry = ttk.Entry(search_frame, textvariable=search_var)
        search_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        
        # 创建列表框
        list_frame = ttk.Frame(legend_window)
        list_frame.pack(fill=tk.BOTH, expand=True, padx=10)
        
        scrollbar = ttk.Scrollbar(list_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        listbox = tk.Listbox(list_frame, yscrollcommand=scrollbar.set)
        listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # 添加颜色映射
        color_cycle = plt.rcParams['axes.prop_cycle'].by_key()['color']
        
        def filter_listbox(*args):
            search_text = search_var.get().lower()
            listbox.delete(0, tk.END)
            for idx, col in enumerate(selected_columns):
                if search_text in col.lower():
                    color = color_cycle[idx % len(color_cycle)]
                    # 将matplotlib颜色转换为hex格式
                    if isinstance(color, (tuple, list)):
                        rgb = tuple(int(c * 255) for c in color[:3])
                    elif isinstance(color, str):
                        # 如果已经是颜色名称或hex格式，直接使用
                        rgb = (0, 0, 0) # 默认黑色
                    else:
                        rgb = (0, 0, 0)
                    hex_color = f'#{rgb[0]:02x}{rgb[1]:02x}{rgb[2]:02x}'
                    listbox.insert(tk.END, f"■ {col}")
                    listbox.itemconfig(tk.END, {'fg': hex_color})
        
        search_var.trace('w', filter_listbox)
        filter_listbox()
        
        scrollbar.config(command=listbox.yview)
        
    def save_plot(self):
        """保存图表为图片"""
        file_path = filedialog.asksaveasfilename(
            title="保存图表",
            defaultextension=".png",
            filetypes=[
                ("PNG图片", "*.png"),
                ("JPG图片", "*.jpg"),
                ("PDF文件", "*.pdf"),
                ("SVG文件", "*.svg"),
                ("所有文件", "*.*")
            ]
        )
        
        if file_path:
            try:
                self.figure.savefig(file_path, dpi=300, bbox_inches='tight')
                messagebox.showinfo("成功", f"图表已保存至：{file_path}")
            except Exception as e:
                messagebox.showerror("错误", f"保存图片失败：{str(e)}")
    
    # ========== 常用分组功能方法 ==========
    

    def save_groups(self):
        """保存分组数据（别名方法）"""
        self.save_groups_to_file()
    
    def save_groups_to_file(self):
        """保存常用分组数据到文件"""
        try:
            data = {
                'groups': self.groups,
                'group_columns': self.group_columns
            }
            with open(self.groups_file, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            print("分组数据保存成功")
        except Exception as e:
            print(f"保存分组数据失败: {e}")
            messagebox.showerror("错误", f"保存分组数据失败：{str(e)}")
    
    def update_group_combobox(self):
        """更新分组下拉框列表"""
        group_list = list(self.groups.keys())
        self.group_combobox['values'] = group_list
        if not group_list:
            self.group_var.set('')
            self.current_group = None
            self.clear_group_display()
    
    def clear_group_display(self):
        """清空分组列显示区域"""
        for widget in self.group_scrollable_frame.winfo_children():
            widget.destroy()
        self.group_labels = {}
    

    def create_new_group(self):
        """创建新的常用分组 - 弹出独立的选择表单"""
        # 先弹出对话框输入分组名称
        name_dialog = tk.Toplevel(self.root)
        name_dialog.title("新建分组")
        name_dialog.geometry("400x200")
        name_dialog.transient(self.root)
        name_dialog.grab_set()
        name_dialog.focus_force()  # 强制获取焦点
        
        # 居中显示
        name_dialog.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width() // 2) - 200
        y = self.root.winfo_y() + (self.root.winfo_height() // 2) - 100
        name_dialog.geometry(f"+{x}+{y}")
        
        ttk.Label(name_dialog, text="分组名称:", font=('Arial', 10)).pack(anchor=tk.W, padx=20, pady=(20, 5))
        
        name_var = tk.StringVar()
        name_entry = ttk.Entry(name_dialog, textvariable=name_var, width=40)
        name_entry.pack(padx=20, pady=(0, 20))
        name_entry.focus_set()
        
        group_name_result = {'name': None}
        
        def on_confirm_name():
            group_name = name_var.get().strip()
            if not group_name:
                messagebox.showwarning("警告", "请输入分组名称！")
                return
            
            if group_name in self.groups:
                messagebox.showwarning("警告", "该分组名称已存在！")
                return
            
            group_name_result['name'] = group_name
            name_dialog.destroy()
            
            # 关闭名称对话框后，打开列选择表单
            self.open_column_selection_dialog(group_name)
        
        def on_cancel_name():
            name_dialog.destroy()
        
        btn_frame = ttk.Frame(name_dialog)
        btn_frame.pack(pady=10)
        
        ttk.Button(btn_frame, text="下一步", command=on_confirm_name, width=10).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="取消", command=on_cancel_name, width=10).pack(side=tk.LEFT, padx=5)
        
        # 绑定回车键
        name_entry.bind('<Return>', lambda e: on_confirm_name())
        name_dialog.bind('<Escape>', lambda e: on_cancel_name())
    
    def open_column_selection_dialog(self, group_name):
        """打开列选择表单对话框"""
        # 创建列选择对话框
        dialog = tk.Toplevel(self.root)
        dialog.title(f"选择分组 '{group_name}' 的数据列")
        dialog.geometry("500x600")
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.focus_force()  # 强制获取焦点
        
        # 居中显示
        dialog.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width() // 2) - 250
        y = self.root.winfo_y() + (self.root.winfo_height() // 2) - 300
        dialog.geometry(f"+{x}+{y}")
        
        # 主框架
        main_frame = ttk.Frame(dialog, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 提示信息
        ttk.Label(main_frame, text=f"请勾选要加入 '{group_name}' 分组的数据列：", 
                 font=('Arial', 10, 'bold')).pack(anchor=tk.W, pady=(0, 10))
        
        # 创建可滚动的列选择区域
        scroll_frame = ttk.Frame(main_frame)
        scroll_frame.pack(fill=tk.BOTH, expand=True)
        
        canvas = tk.Canvas(scroll_frame, bg='white')
        scrollbar = ttk.Scrollbar(scroll_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # 存储勾选状态的字典
        column_vars = {}
        
        # 搜索框
        search_frame = ttk.Frame(main_frame)
        search_frame.pack(fill=tk.X, pady=(10, 5))
        ttk.Label(search_frame, text="搜索:").pack(side=tk.LEFT)
        search_var = tk.StringVar()
        search_entry = ttk.Entry(search_frame, textvariable=search_var)
        search_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        
        # 全选/全不选按钮
        select_frame = ttk.Frame(main_frame)
        select_frame.pack(fill=tk.X, pady=5)
        
        def select_all_columns():
            for var in column_vars.values():
                var.set(1)
        
        def deselect_all_columns():
            for var in column_vars.values():
                var.set(0)
        
        ttk.Button(select_frame, text="全选", command=select_all_columns).pack(side=tk.LEFT, padx=2)
        ttk.Button(select_frame, text="全不选", command=deselect_all_columns).pack(side=tk.LEFT, padx=2)
        
        # 显示已选数量
        count_label = ttk.Label(select_frame, text="已选: 0")
        count_label.pack(side=tk.RIGHT)
        
        def update_count():
            selected_count = sum(1 for var in column_vars.values() if var.get() == 1)
            count_label.config(text=f"已选: {selected_count}")
        
        # 创建复选框
        if not self.column_vars:
            ttk.Label(scrollable_frame, text="请先加载Excel文件！", foreground="red").pack(anchor=tk.W, padx=10, pady=10)
        else:
            for col in self.column_vars.keys():
                var = tk.IntVar(value=0)
                column_vars[col] = var
                
                cb = ttk.Checkbutton(
                    scrollable_frame,
                    text=col,
                    variable=var,
                    command=update_count
                )
                cb.pack(anchor=tk.W, padx=5, pady=2)
        
        # 搜索功能
        def filter_columns(*args):
            search_text = search_var.get().lower()
            for widget in scrollable_frame.winfo_children():
                if isinstance(widget, ttk.Checkbutton):
                    col_text = widget.cget("text").lower()
                    if search_text in col_text:
                        widget.pack(anchor=tk.W, padx=5, pady=2)
                    else:
                        widget.pack_forget()
        
        search_var.trace('w', filter_columns)
        
        # 按钮区域
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=(10, 0))
        
        def on_save():
            # 获取勾选的列
            selected_columns = [col for col, var in column_vars.items() if var.get() == 1]
            
            if not selected_columns:
                result = messagebox.askyesno(
                    "确认",
                    f"分组 '{group_name}' 中没有勾选任何列，确定要创建空分组吗？"
                )
                if not result:
                    return
            
            # 创建分组数据
            group_data = {}
            for col in self.column_vars.keys():
                group_data[col] = column_vars[col].get()
            
            self.groups[group_name] = group_data
            self.group_columns[group_name] = selected_columns
            
            # 保存到文件
            self.save_groups_to_file()
            
            # 更新下拉框
            self.update_group_combobox()
            
            # 选中新创建的分组
            self.group_var.set(group_name)
            self.current_group = group_name
            self.update_group_display(group_name)
            
            dialog.destroy()
            messagebox.showinfo("成功", f"分组 '{group_name}' 创建成功！\n包含 {len(selected_columns)} 个数据列。")
        
        def on_cancel():
            dialog.destroy()
        
        ttk.Button(btn_frame, text="保存并创建", command=on_save, width=15).pack(side=tk.RIGHT, padx=5)
        ttk.Button(btn_frame, text="取消", command=on_cancel, width=10).pack(side=tk.RIGHT, padx=5)
    
    def save_current_group(self):
        """保存当前分组"""
        if not self.current_group:
            messagebox.showwarning("警告", "请先选择一个分组！")
            return
        
        if self.current_group not in self.groups:
            messagebox.showwarning("警告", "当前选择的分组不存在！")
            return
        
        # 获取当前勾选的列
        selected_columns = [col for col, var in self.column_vars.items() if var.get() == 1]
        all_columns = list(self.column_vars.keys())
        
        # 保存当前所有列的勾选状态
        current_selections = {}
        for col in all_columns:
            current_selections[col] = self.column_vars[col].get()
        
        self.groups[self.current_group] = current_selections
        self.group_columns[self.current_group] = selected_columns
        
        # 保存到文件
        self.save_groups_to_file()
        
        # 更新分组显示
        self.update_group_display(self.current_group)
        
        messagebox.showinfo("成功", f"分组 '{self.current_group}' 已保存！\n包含 {len(selected_columns)} 个数据列。")
    
    def delete_current_group(self):
        """删除当前分组"""
        if not self.current_group:
            messagebox.showwarning("警告", "请先选择一个分组！")
            return
        
        if self.current_group not in self.groups:
            messagebox.showwarning("警告", "当前选择的分组不存在！")
            return
        
        # 确认对话框
        result = messagebox.askyesno(
            "确认删除",
                f"确定要删除分组 '{self.current_group}' 吗？\n此操作不可恢复！"
        )
        
        if not result:
            return
        
        # 删除分组
        del self.groups[self.current_group]
        if self.current_group in self.group_columns:
            del self.group_columns[self.current_group]
        
        # 保存到文件
        self.save_groups_to_file()
        
        # 更新下拉框和显示
        self.update_group_combobox()
        self.current_group = None
        self.group_var.set('')
        self.clear_group_display()
        
        messagebox.showinfo("成功", "分组已删除！")
    
    def edit_current_group(self):
        """编辑当前分组的列配置"""
        if not self.current_group:
            messagebox.showwarning("警告", "请先选择一个分组！")
            return
        
        if self.current_group not in self.groups:
            messagebox.showwarning("警告", "当前选择的分组不存在！")
            return
        
        # 直接打开编辑表单
        self.open_edit_column_selection_dialog(self.current_group)
    
    def rename_current_group(self):
        """重命名当前分组"""
        if not self.current_group:
            messagebox.showwarning("警告", "请先选择一个分组！")
            return
        
        if self.current_group not in self.groups:
            messagebox.showwarning("警告", "当前选择的分组不存在！")
            return
        
        old_name = self.current_group
        
        # 弹出重命名对话框
        dialog = tk.Toplevel(self.root)
        dialog.title("重命名分组")
        dialog.geometry("400x150")
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.focus_force()  # 强制获取焦点
        
        # 居中显示
        dialog.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width() // 2) - 200
        y = self.root.winfo_y() + (self.root.winfo_height() // 2) - 75
        dialog.geometry(f"+{x}+{y}")
        
        ttk.Label(dialog, text="新分组名称:", font=('Arial', 10)).pack(anchor=tk.W, padx=20, pady=(15, 5))
        
        new_name_var = tk.StringVar(value=old_name)
        new_name_entry = ttk.Entry(dialog, textvariable=new_name_var, width=40)
        new_name_entry.pack(padx=20, pady=(0, 15))
        new_name_entry.select_range(0, tk.END)
        new_name_entry.focus_set()
        
        def on_confirm():
            new_name = new_name_var.get().strip()
            
            if not new_name:
                messagebox.showwarning("警告", "请输入新的分组名称！")
                return
            
            if new_name == old_name:
                dialog.destroy()
                return
            
            if new_name in self.groups:
                messagebox.showwarning("警告", "该分组名称已存在！")
                return
            
            # 重命名分组
            self.groups[new_name] = self.groups.pop(old_name)
            if old_name in self.group_columns:
                self.group_columns[new_name] = self.group_columns.pop(old_name)
            
            # 更新当前分组
            self.current_group = new_name
            
            # 保存到文件
            self.save_groups_to_file()
            
            # 更新下拉框和显示
            self.update_group_combobox()
            self.group_var.set(new_name)
            self.update_group_display(new_name)
            
            dialog.destroy()
            messagebox.showinfo("成功", f"分组已重命名为 '{new_name}'！")
        
        def on_cancel():
            dialog.destroy()
        
        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(pady=5)
        
        ttk.Button(btn_frame, text="确定", command=on_confirm, width=10).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="取消", command=on_cancel, width=10).pack(side=tk.LEFT, padx=5)
        
        # 绑定回车键
        new_name_entry.bind('<Return>', lambda e: on_confirm())
        dialog.bind('<Escape>', lambda e: on_cancel())
    
    def open_edit_column_selection_dialog(self, group_name):
        """打开编辑分组的选择表单对话框"""
        # 获取当前分组的数据
        group_data = self.groups.get(group_name, {})
        
        # 创建编辑对话框
        dialog = tk.Toplevel(self.root)
        dialog.title(f"编辑分组 '{group_name}' 的数据列")
        dialog.geometry("500x600")
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.focus_force()  # 强制获取焦点
        
        # 居中显示
        dialog.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width() // 2) - 250
        y = self.root.winfo_y() + (self.root.winfo_height() // 2) - 300
        dialog.geometry(f"+{x}+{y}")
        
        # 主框架
        main_frame = ttk.Frame(dialog, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 提示信息
        ttk.Label(main_frame, text=f"请编辑分组 '{group_name}' 的数据列：", 
                 font=('Arial', 10, 'bold')).pack(anchor=tk.W, pady=(0, 10))
        
        # 创建可滚动的列选择区域
        scroll_frame = ttk.Frame(main_frame)
        scroll_frame.pack(fill=tk.BOTH, expand=True)
        
        canvas = tk.Canvas(scroll_frame, bg='white')
        scrollbar = ttk.Scrollbar(scroll_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # 存储勾选状态的字典
        column_vars = {}
        
        # 搜索框
        search_frame = ttk.Frame(main_frame)
        search_frame.pack(fill=tk.X, pady=(10, 5))
        ttk.Label(search_frame, text="搜索:").pack(side=tk.LEFT)
        search_var = tk.StringVar()
        search_entry = ttk.Entry(search_frame, textvariable=search_var)
        search_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        
        # 全选/全不选按钮
        select_frame = ttk.Frame(main_frame)
        select_frame.pack(fill=tk.X, pady=5)
        
        def select_all_columns():
            for var in column_vars.values():
                var.set(1)
        
        def deselect_all_columns():
            for var in column_vars.values():
                var.set(0)
        
        ttk.Button(select_frame, text="全选", command=select_all_columns).pack(side=tk.LEFT, padx=2)
        ttk.Button(select_frame, text="全不选", command=deselect_all_columns).pack(side=tk.LEFT, padx=2)
        
        # 显示已选数量
        count_label = ttk.Label(select_frame, text="已选: 0")
        count_label.pack(side=tk.RIGHT)
        
        def update_count():
            selected_count = sum(1 for var in column_vars.values() if var.get() == 1)
            count_label.config(text=f"已选: {selected_count}")
        
        # 创建复选框
        if not self.column_vars:
            ttk.Label(scrollable_frame, text="请先加载Excel文件！", foreground="red").pack(anchor=tk.W, padx=10, pady=10)
        else:
            for col in self.column_vars.keys():
                # 从分组数据中恢复勾选状态
                checked = group_data.get(col, 0)
                var = tk.IntVar(value=checked)
                column_vars[col] = var
                
                cb = ttk.Checkbutton(
                    scrollable_frame,
                    text=col,
                    variable=var,
                    command=update_count
                )
                cb.pack(anchor=tk.W, padx=5, pady=2)
        
        # 更新初始计数
        update_count()
        
        # 搜索功能
        def filter_columns(*args):
            search_text = search_var.get().lower()
            for widget in scrollable_frame.winfo_children():
                if isinstance(widget, ttk.Checkbutton):
                    col_text = widget.cget("text").lower()
                    if search_text in col_text:
                        widget.pack(anchor=tk.W, padx=5, pady=2)
                    else:
                        widget.pack_forget()
        
        search_var.trace('w', filter_columns)
        
        # 按钮区域
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=(10, 0))
        
        def on_save():
            # 获取勾选的列
            selected_columns = [col for col, var in column_vars.items() if var.get() == 1]
            
            # 更新分组数据
            new_group_data = {}
            for col in self.column_vars.keys():
                new_group_data[col] = column_vars[col].get()
            
            self.groups[group_name] = new_group_data
            self.group_columns[group_name] = selected_columns
            
            # 保存到文件
            self.save_groups_to_file()
            
            # 更新显示
            self.update_group_display(group_name)
            
            dialog.destroy()
            messagebox.showinfo("成功", f"分组 '{group_name}' 已保存！\n包含 {len(selected_columns)} 个数据列。")
        
        def on_cancel():
            dialog.destroy()
        
        ttk.Button(btn_frame, text="保存修改", command=on_save, width=15).pack(side=tk.RIGHT, padx=5)
        ttk.Button(btn_frame, text="取消", command=on_cancel, width=10).pack(side=tk.RIGHT, padx=5)
    
    def on_group_selected(self, event):
        """分组选择下拉框改变时的回调"""
        selected = self.group_var.get()
        if not selected:
            self.current_group = None
            self.clear_group_display()
            return
        
        if selected not in self.groups:
            return
        
        self.current_group = selected
        
        # 恢复该分组的勾选状态
        group_data = self.groups[selected]
        if group_data:
            for col, checked in group_data.items():
                if col in self.column_vars:
                    self.column_vars[col].set(checked)
        else:
            # 如果分组为空，清空所有勾选
            for col in self.column_vars:
                self.column_vars[col].set(0)
        
        # 更新"已勾选数据"显示
        
        # 更新分组列显示
        self.update_group_display(selected)
    
    def on_group_canvas_scroll(self, event):
        """处理分组Canvas的滚轮事件"""
        try:
            if event.delta:
                scroll_units = -1 * int(event.delta / 120)
            else:
                if event.num == 4:
                    scroll_units = -1
                elif event.num == 5:
                    scroll_units = 1
                else:
                    return
            
            self.group_canvas.yview_scroll(scroll_units, "units")
        except Exception as e:
            pass
    
    # ========== 数据诊断功能方法 ==========
    
    def on_strategy_selected(self, event):
        """策略选择回调"""
        selected_strategy = self.diagnosis_strategy_var.get()
        
        # 隐藏当前配置框架
        if self.current_config_frame:
            self.current_config_frame.pack_forget()
            self.current_config_frame = None
        
        # 根据选择的策略显示对应的配置框架
        if selected_strategy == '波动剧烈检测':
            self.current_config_frame = self.volatility_config
            self.current_config_frame.pack(fill=tk.X, pady=5)
        elif selected_strategy == '同增同减检测':
            self.current_config_frame = self.correlation_config
            self.current_config_frame.pack(fill=tk.X, pady=5)
        elif selected_strategy == '数据范围异常':
            self.current_config_frame = self.range_config
            self.current_config_frame.pack(fill=tk.X, pady=5)
        elif selected_strategy == '自定义策略...':
            # TODO: 打开自定义策略编辑器
            self.current_config_frame = self.volatility_config
            self.current_config_frame.pack(fill=tk.X, pady=5)
            messagebox.showinfo("提示", "自定义策略功能开发中...")
            return
        
        # 更新下拉框选项
        self.update_diagnosis_comboboxes()
    
    def update_diagnosis_comboboxes(self):
        """更新诊断下拉框的数据列选项"""
        if self.df is None:
            return
        
        # 获取所有可用列（排除时间列）
        available_columns = [col for col in self.df.columns if col != self.time_column]
        
        # 更新各个下拉框
        self.volatility_column_var.set('')
        self.correlation_col1_var.set('')
        self.correlation_col2_var.set('')
        self.range_column_var.set('')
        
        # 找到对应的 Combobox 并更新值
        for widget in self.volatility_config.winfo_children():
            if isinstance(widget, ttk.Combobox):
                widget['values'] = available_columns
                if available_columns:
                    widget.current(0)
        
        for widget in self.correlation_config.winfo_children():
            if isinstance(widget, ttk.Combobox):
                widget['values'] = available_columns
                if available_columns:
                    widget.current(0)
        
        for widget in self.range_config.winfo_children():
            # range_config 现在包含 row1 和 row2 两个 Frame，需要遍历它们的子控件
            if isinstance(widget, ttk.Frame):
                for child in widget.winfo_children():
                    if isinstance(child, ttk.Combobox):
                        child['values'] = available_columns
                        if available_columns:
                            child.current(0)
            elif isinstance(widget, ttk.Combobox):
                widget['values'] = available_columns
                if available_columns:
                    widget.current(0)
    
    def run_diagnosis(self):
        """运行诊断"""
        if self.df is None:
            messagebox.showwarning("警告", "请先加载Excel文件！")
            return
        
        selected_strategy = self.diagnosis_strategy_var.get()
        
        print(f"\n[DEBUG] ========== 开始数据诊断 ==========")
        print(f"[DEBUG] 选择策略: {selected_strategy}")
        
        try:
            if selected_strategy == '波动剧烈检测':
                self.run_volatility_diagnosis()
            elif selected_strategy == '同增同减检测':
                self.run_correlation_diagnosis()
            elif selected_strategy == '数据范围异常':
                self.run_range_diagnosis()
            else:
                messagebox.showinfo("提示", "请选择一个诊断策略")
        except Exception as e:
            print(f"[ERROR] 诊断失败: {e}")
            import traceback
            traceback.print_exc()
            messagebox.showerror("错误", f"诊断失败: {e}")
    
    def run_volatility_diagnosis(self):
        """运行波动剧烈检测"""
        print("[DEBUG] 执行波动剧烈检测")
        
        # 获取参数
        threshold = float(self.volatility_threshold_var.get())
        column = self.volatility_column_var.get()
        
        if not column:
            messagebox.showwarning("警告", "请选择要检测的数据列！")
            return
        
        print(f"[DEBUG] 列: {column}, 阈值: {threshold}%")
        
        # 获取有效数据
        valid_mask = self.df[column].notna()
        valid_data = self.df.loc[valid_mask, column]
        valid_times = self.df.loc[valid_mask, self.time_column]
        
        # 计算相邻时刻的变化率
        # change_rate = (当前值 - 前一个值) / abs(前一个值) * 100
        anomalies = []
        for idx in range(1, len(valid_data)):
            prev_value = valid_data.iloc[idx - 1]
            curr_value = valid_data.iloc[idx]
            
            # 避免除以零
            if prev_value == 0:
                continue
            
            change_rate = abs(curr_value - prev_value) / abs(prev_value) * 100
            
            if change_rate > threshold:
                time_val = valid_times.iloc[idx]
                anomalies.append({
                    'index': idx,
                    'time': time_val,
                    'value': curr_value,
                    'prev_value': prev_value,
                    'change_rate': change_rate
                })
        
        # 存储结果
        self.diagnosis_results = {
            'strategy': '波动剧烈检测',
            'column': column,
            'threshold': threshold,
            'anomalies': anomalies,
            'total_count': len(valid_data),
            'anomaly_count': len(anomalies)
        }
        
        # 更新状态
        if anomalies:
            status_text = f"⚠ 发现 {len(anomalies)} 个异常点"
            self.diagnosis_status_label.config(text=f"状态: {status_text}", foreground="red")
            messagebox.showwarning("诊断结果", 
                f"波动剧烈检测完成！\n\n"
                f"检测列: {column}\n"
                f"阈值: {threshold}%\n"
                f"发现异常点: {len(anomalies)} 个\n\n"
                f"点击'查看报告'查看详细结果")
        else:
            self.diagnosis_status_label.config(text="状态: [√] 数据正常", foreground="green")
            messagebox.showinfo("诊断结果", 
                f"波动剧烈检测完成！\n\n"
                f"检测列: {column}\n"
                f"阈值: {threshold}%\n"
                f"未发现异常点，数据波动正常")
    
    def run_correlation_diagnosis(self):
        """运行同增同减检测"""
        print("[DEBUG] 执行同增同减检测")
        
        col1 = self.correlation_col1_var.get()
        col2 = self.correlation_col2_var.get()
        
        if not col1 or not col2:
            messagebox.showwarning("警告", "请选择两个数据列进行对比！")
            return
        
        if col1 == col2:
            messagebox.showwarning("警告", "请选择两个不同的数据列！")
            return
        
        print(f"[DEBUG] 列1: {col1}, 列2: {col2}")
        
        # 【修复】同时检查两列是否都有值，确保行对应
        valid_mask = self.df[col1].notna() & self.df[col2].notna()
        valid_df = self.df.loc[valid_mask]
        
        if len(valid_df) < 2:
            messagebox.showwarning("警告", "有效数据太少，无法进行同增同减检测！")
            return
        
        # 获取有效数据
        data1 = valid_df[col1].reset_index(drop=True)
        data2 = valid_df[col2].reset_index(drop=True)
        times = valid_df[self.time_column].reset_index(drop=True)
        
        print(f"[DEBUG] 有效数据行数: {len(data1)}")
        
        # 计算变化方向
        diff1 = data1.diff().dropna()
        diff2 = data2.diff().dropna()
        
        # 检测不一致的点
        anomalies = []
        for idx in range(len(diff1)):
            # 判断方向是否一致
            d1 = diff1.iloc[idx]
            d2 = diff2.iloc[idx]
            
            # 【修复】改进的判断逻辑
            # 1. 如果两个变化量都为0，跳过（没有变化）
            # 2. 如果一个为0一个不为0，视为不一致
            # 3. 如果两个符号不同（一正一负），视为不一致
            
            is_anomaly = False
            anomaly_type = ""
            
            if d1 == 0 and d2 == 0:
                # 两个都没有变化，跳过
                continue
            elif d1 == 0 or d2 == 0:
                # 一个变化了一个没变化
                is_anomaly = True
                anomaly_type = "单侧变化" if (d1 == 0 or d2 == 0) else ""
            elif (d1 > 0 and d2 < 0) or (d1 < 0 and d2 > 0):
                # 方向相反
                is_anomaly = True
                anomaly_type = "方向相反"
            
            if is_anomaly:
                time_val = times.iloc[idx + 1]
                anomalies.append({
                    'index': idx + 1,
                    'time': time_val,
                    'value1': data1.iloc[idx + 1],
                    'value2': data2.iloc[idx + 1],
                    'diff1': d1,
                    'diff2': d2,
                    'type': anomaly_type
                })
        
        # 存储结果
        self.diagnosis_results = {
            'strategy': '同增同减检测',
            'column1': col1,
            'column2': col2,
            'anomalies': anomalies,
            'total_count': len(data1),
            'anomaly_count': len(anomalies)
        }
        
        # 更新状态
        if anomalies:
            status_text = f"⚠ 发现 {len(anomalies)} 个不一致点"
            self.diagnosis_status_label.config(text=f"状态: {status_text}", foreground="red")
            messagebox.showwarning("诊断结果", 
                f"同增同减检测完成！\n\n"
                f"列1: {col1}\n"
                f"列2: {col2}\n"
                f"有效数据行: {len(data1)}\n"
                f"发现不一致点: {len(anomalies)} 个\n\n"
                f"点击'查看报告'查看详细结果")
        else:
            self.diagnosis_status_label.config(text="状态: [√] 数据正常", foreground="green")
            messagebox.showinfo("诊断结果", 
                f"同增同减检测完成！\n\n"
                f"列1: {col1}\n"
                f"列2: {col2}\n"
                f"有效数据行: {len(data1)}\n"
                f"未发现不一致点，两列数据变化趋势一致")
            messagebox.showinfo("诊断结果", 
                f"同增同减检测完成！\n\n"
                f"列1: {col1}\n"
                f"列2: {col2}\n"
                f"未发现不一致点，两列数据变化趋势一致")
    
    def run_range_diagnosis(self):
        """运行数据范围异常检测"""
        print("[DEBUG] 执行数据范围异常检测")
        
        column = self.range_column_var.get()
        min_value = float(self.range_min_var.get())
        max_value = float(self.range_max_var.get())
        
        if not column:
            messagebox.showwarning("警告", "请选择要检测的数据列！")
            return
        
        print(f"[DEBUG] 列: {column}, 范围: [{min_value}, {max_value}]")
        
        # 检测异常点
        data = self.df[column].dropna()
        anomalies = []
        
        for idx, value in enumerate(data):
            if value < min_value or value > max_value:
                time_val = self.df[self.time_column].iloc[idx]
                anomalies.append({
                    'index': idx,
                    'time': time_val,
                    'value': value,
                    'type': '低于最小值' if value < min_value else '超过最大值'
                })
        
        # 存储结果
        self.diagnosis_results = {
            'strategy': '数据范围异常',
            'column': column,
            'min_value': min_value,
            'max_value': max_value,
            'anomalies': anomalies,
            'total_count': len(data),
            'anomaly_count': len(anomalies)
        }
        
        # 更新状态
        if anomalies:
            status_text = f"⚠ 发现 {len(anomalies)} 个异常值"
            self.diagnosis_status_label.config(text=f"状态: {status_text}", foreground="red")
            messagebox.showwarning("诊断结果", 
                f"数据范围异常检测完成！\n\n"
                f"检测列: {column}\n"
                f"期望范围: [{min_value}, {max_value}]\n"
                f"发现异常值: {len(anomalies)} 个\n\n"
                f"点击'查看报告'查看详细结果")
        else:
            self.diagnosis_status_label.config(text="状态: [√] 数据正常", foreground="green")
            messagebox.showinfo("诊断结果", 
                f"数据范围异常检测完成！\n\n"
                f"检测列: {column}\n"
                f"期望范围: [{min_value}, {max_value}]\n"
                f"未发现异常值，所有数据都在正常范围内")
    
    def view_diagnosis_report(self):
        """查看诊断报告"""
        if self.diagnosis_results is None:
            messagebox.showwarning("警告", "还没有运行诊断！请先点击'运行诊断'按钮")
            return
        
        # 创建报告窗口
        report_window = tk.Toplevel(self.root)
        report_window.title("数据诊断报告")
        report_window.geometry("700x500")
        report_window.transient(self.root)
        report_window.grab_set()
        report_window.focus_force()
        
        # 主框架
        main_frame = ttk.Frame(report_window, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 标题
        strategy = self.diagnosis_results.get('strategy', '未知策略')
        ttk.Label(main_frame, text=f"诊断策略: {strategy}", 
                 font=('Arial', 12, 'bold')).pack(anchor=tk.W, pady=(0, 10))
        
        # 创建文本框显示报告
        text_frame = ttk.Frame(main_frame)
        text_frame.pack(fill=tk.BOTH, expand=True)
        
        scrollbar = ttk.Scrollbar(text_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        report_text = tk.Text(text_frame, wrap=tk.WORD, yscrollcommand=scrollbar.set)
        report_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=report_text.yview)
        
        # 生成报告内容
        report_content = self.generate_diagnosis_report()
        report_text.insert(tk.END, report_content)
        report_text.config(state=tk.DISABLED)
        
        # 按钮
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=(10, 0))
        
        ttk.Button(btn_frame, text="导出报告", 
                  command=lambda: self.export_diagnosis_report(report_content)).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="关闭", 
                  command=report_window.destroy).pack(side=tk.RIGHT, padx=5)
    
    def generate_diagnosis_report(self):
        """生成诊断报告内容"""
        if self.diagnosis_results is None:
            return "无诊断结果"
        
        strategy = self.diagnosis_results.get('strategy', '未知策略')
        report_lines = [f"{'='*60}", f"数据诊断报告", f"策略: {strategy}", f"{'='*60}", ""]
        
        if strategy == '波动剧烈检测':
            report_lines.extend([
                f"检测列: {self.diagnosis_results.get('column', '')}",
                f"波动阈值: {self.diagnosis_results.get('threshold', 0)}%",
                f"总数据点: {self.diagnosis_results.get('total_count', 0)}",
                f"异常点数量: {self.diagnosis_results.get('anomaly_count', 0)}",
                ""
            ])
            
            anomalies = self.diagnosis_results.get('anomalies', [])
            if anomalies:
                report_lines.append("异常点详情:")
                report_lines.append("-" * 60)
                for i, anomaly in enumerate(anomalies, 1):
                    time_str = str(anomaly['time'])
                    report_lines.append(
                        f"{i}. 时间: {time_str}, 当前值: {anomaly['value']:.2f}, 前一值: {anomaly['prev_value']:.2f}, 变化率: {anomaly['change_rate']:.1f}%"
                    )
        
        elif strategy == '同增同减检测':
            report_lines.extend([
                f"检测列1: {self.diagnosis_results.get('column1', '')}",
                f"检测列2: {self.diagnosis_results.get('column2', '')}",
                f"总数据点: {self.diagnosis_results.get('total_count', 0)}",
                f"不一致点数量: {self.diagnosis_results.get('anomaly_count', 0)}",
                ""
            ])
            
            anomalies = self.diagnosis_results.get('anomalies', [])
            if anomalies:
                report_lines.append("不一致点详情:")
                report_lines.append("-" * 60)
                for i, anomaly in enumerate(anomalies, 1):
                    time_str = str(anomaly['time'])
                    report_lines.append(
                        f"{i}. 时间: {time_str}"
                    )
                    report_lines.append(
                        f"   列1值: {anomaly['value1']:.2f} (变化: {anomaly['diff1']:.2f})"
                    )
                    report_lines.append(
                        f"   列2值: {anomaly['value2']:.2f} (变化: {anomaly['diff2']:.2f})"
                    )
        
        elif strategy == '数据范围异常':
            report_lines.extend([
                f"检测列: {self.diagnosis_results.get('column', '')}",
                f"期望范围: [{self.diagnosis_results.get('min_value', 0)}, {self.diagnosis_results.get('max_value', 0)}]",
                f"总数据点: {self.diagnosis_results.get('total_count', 0)}",
                f"异常值数量: {self.diagnosis_results.get('anomaly_count', 0)}",
                ""
            ])
            
            anomalies = self.diagnosis_results.get('anomalies', [])
            if anomalies:
                report_lines.append("异常值详情:")
                report_lines.append("-" * 60)
                for i, anomaly in enumerate(anomalies, 1):
                    time_str = str(anomaly['time'])
                    report_lines.append(
                        f"{i}. 时间: {time_str}, 值: {anomaly['value']:.2f}, 类型: {anomaly['type']}"
                    )
        
        report_lines.extend(["", "=" * 60, "报告生成时间:", str(pd.Timestamp.now())])
        
        return "\n".join(report_lines)
    
    def export_diagnosis_report(self, report_content):
        """导出诊断报告"""
        file_path = filedialog.asksaveasfilename(
            title="保存诊断报告",
            defaultextension=".txt",
            filetypes=[("文本文件", "*.txt"), ("所有文件", "*.*")]
        )
        
        if file_path:
            try:
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write(report_content)
                messagebox.showinfo("成功", f"诊断报告已保存至：\n{file_path}")
            except Exception as e:
                messagebox.showerror("错误", f"保存报告失败：{e}")
    
    def clear_diagnosis(self):
        """清除诊断结果"""
        self.diagnosis_results = None
        self.diagnosis_status_label.config(text="状态: 未运行", foreground="gray")
        messagebox.showinfo("提示", "诊断结果已清除")
    
    # ==================== 批量分析功能 ====================
    
    def select_batch_folder(self):
        """选择批量分析的文件夹"""
        folder_path = filedialog.askdirectory(title="选择包含Excel文件的文件夹")
        if folder_path:
            self.batch_folder_path = folder_path
            self.batch_folder_label.config(text=folder_path)
            print(f"[INFO] 已选择文件夹: {folder_path}")
    
    def scan_excel_files(self):
        """递归扫描文件夹及其子文件夹下的Excel文件"""
        if not self.batch_folder_path:
            messagebox.showwarning("警告", "请先选择文件夹！")
            return
        
        print(f"[DEBUG] scan_excel_files 开始执行")
        import re
        
        # 清空列表
        self.file_listbox.delete(0, tk.END)
        print(f"[DEBUG] 已清空文件列表")
        
        # 初始化进度
        self.batch_progress_label.config(text="正在遍历文件夹...")
        self.batch_progress_bar['value'] = 0
        self.batch_time_label.config(text="耗时: 0.0秒", foreground="gray")
        self.root.update()
        print(f"[DEBUG] 已初始化进度条")
        
        # 文件命名规则
        pattern = re.compile(r'^(.+?)_(BATTERY_)?REPORT$', re.IGNORECASE)
        
        # 第一步：快速收集所有Excel文件路径，并找出样本文件
        print(f"[DEBUG] 开始 os.walk 遍历文件夹...")
        all_excel_files = []
        battery_sample = None
        pcs_sample = None
        
        for root_dir, dirs, files in os.walk(self.batch_folder_path):
            for filename in files:
                # 过滤条件：
                # 1. 必须是 .xlsx, .xls 或 .csv 文件
                # 2. 排除以 ~$ 开头的 Excel 临时文件（Excel 打开文件时自动创建）
                # 3. 排除以 ~ 开头的其他临时文件（如 Word 的 ~WRL 开头文件）
                if filename.endswith(('.xlsx', '.xls', '.csv')) and not filename.startswith('~'):
                    file_path = os.path.join(root_dir, filename)
                    all_excel_files.append((root_dir, filename))
                    
                    # 提前找出样本文件（用于加载列名）
                    name_without_ext = os.path.splitext(filename)[0]
                    match = pattern.match(name_without_ext)
                    if match:
                        is_battery = match.group(2) is not None
                        if is_battery and battery_sample is None:
                            battery_sample = file_path
                        elif not is_battery and pcs_sample is None:
                            pcs_sample = file_path
        
        total_files = len(all_excel_files)
        print(f"[DEBUG] os.walk 完成，找到 {total_files} 个Excel文件")
        print(f"[DEBUG] 样本文件 - 电池: {battery_sample}, PCS: {pcs_sample}")
        
        # 立即启动后台线程加载列名（不等待扫描完成）
        if battery_sample or pcs_sample:
            import threading
            thread = threading.Thread(
                target=self._load_columns_for_both_types, 
                args=(battery_sample, pcs_sample),
                daemon=True
            )
            thread.start()
        
        if total_files == 0:
            self.batch_progress_label.config(text="未找到数据文件")
            self.file_count_label.config(text="共 0 个文件")
            self.valid_excel_files = []
            print(f"[DEBUG] 未找到文件，函数返回")
            return
        
        # 第二步：逐个处理文件并实时更新进度
        valid_files = []
        print(f"[DEBUG] 开始逐个处理文件...")
        
        # 【优化】每5%进度更新一次UI，减少刷新次数
        last_progress_update = -5
        for idx, (root_dir, filename) in enumerate(all_excel_files):
            # 处理文件
            file_path = os.path.join(root_dir, filename)
            name_without_ext = os.path.splitext(filename)[0]
            
            # 检查是否符合命名规则
            match = pattern.match(name_without_ext)
            if match:
                device_code = match.group(1)
                # 识别文件类型：BATTERY 或 PCS
                is_battery = match.group(2) is not None  # group(2) 是 "BATTERY_" 或 None
                file_type = 'BATTERY' if is_battery else 'PCS'
                
                valid_files.append({
                    'path': file_path,
                    'filename': filename,
                    'device_code': device_code,
                    'relative_path': os.path.relpath(file_path, self.batch_folder_path),
                    'file_type': file_type  # 添加文件类型
                })
            
            # 【优化】每5%进度更新一次UI
            progress = (idx + 1) / total_files * 100
            if progress - last_progress_update >= 5 or idx == total_files - 1:
                last_progress_update = int(progress)
                self.batch_progress_bar['value'] = progress
                self.batch_progress_label.config(text=f"扫描中: {idx + 1}/{total_files}")
                self.root.update_idletasks()
        
        # 统计文件类型
        battery_count = sum(1 for f in valid_files if f['file_type'] == 'BATTERY')
        pcs_count = sum(1 for f in valid_files if f['file_type'] == 'PCS')
        
        print(f"[DEBUG] 文件处理完成，有效文件: {len(valid_files)} (电池: {battery_count}, PCS: {pcs_count})")
        
        # 存储有效文件列表
        self.valid_excel_files = valid_files
        
        # 更新类型统计标签
        self.type_stats_label.config(text=f"(电池: {battery_count}, PCS: {pcs_count})")
        
        # 根据当前选择的数据类型过滤显示
        self._filter_file_list()
        
        # 扫描完成
        self.batch_progress_bar['value'] = 100
        self.root.update()
        
        print(f"[INFO] 扫描完成，共扫描 {total_files} 个数据文件，找到 {len(valid_files)} 个有效文件")
        print(f"[DEBUG] scan_excel_files 函数结束")
    
    def _load_columns_for_both_types(self, battery_sample, pcs_sample):
        """分别加载电池和PCS两种类型的列名（线程安全）
        
        使用 nrows=0 只读取列名，不读取数据，大幅提升速度
        """
        battery_columns = []
        pcs_columns = []
        
        # 加载电池数据列名（只读取列名，不读取数据）
        if battery_sample:
            print(f"[DEBUG] 加载电池列名: {battery_sample}")
            try:
                # 【性能优化】根据文件类型选择读取方式
                file_ext = os.path.splitext(battery_sample)[1].lower()
                if file_ext == '.csv':
                    # CSV 文件
                    for encoding in ['utf-8', 'gbk', 'gb2312', 'latin1']:
                        try:
                            df = pd.read_csv(battery_sample, nrows=0, encoding=encoding)
                            break
                        except UnicodeDecodeError:
                            continue
                else:
                    # Excel 文件：使用快速引擎，nrows=0 只读取列名
                    try:
                        df = pd.read_excel(battery_sample, nrows=0, engine='calamine')
                    except:
                        try:
                            df = pd.read_excel(battery_sample, nrows=0, engine='openpyxl')
                        except:
                            df = pd.read_excel(battery_sample, nrows=0)
                battery_columns = list(df.columns)
                print(f"[DEBUG] 电池列名: {len(battery_columns)} 列")
            except Exception as e:
                print(f"[WARN] 加载电池列名失败: {e}")
        
        # 加载PCS数据列名（只读取列名，不读取数据）
        if pcs_sample:
            print(f"[DEBUG] 加载PCS列名: {pcs_sample}")
            try:
                # 【性能优化】根据文件类型选择读取方式
                file_ext = os.path.splitext(pcs_sample)[1].lower()
                if file_ext == '.csv':
                    # CSV 文件
                    for encoding in ['utf-8', 'gbk', 'gb2312', 'latin1']:
                        try:
                            df = pd.read_csv(pcs_sample, nrows=0, encoding=encoding)
                            break
                        except UnicodeDecodeError:
                            continue
                else:
                    # Excel 文件：使用快速引擎
                    try:
                        df = pd.read_excel(pcs_sample, nrows=0, engine='calamine')
                    except:
                        try:
                            df = pd.read_excel(pcs_sample, nrows=0, engine='openpyxl')
                        except:
                            df = pd.read_excel(pcs_sample, nrows=0)
                pcs_columns = list(df.columns)
                print(f"[DEBUG] PCS列名: {len(pcs_columns)} 列")
            except Exception as e:
                print(f"[WARN] 加载PCS列名失败: {e}")
        
        print(f"[DEBUG] 列名加载完成 - 电池: {len(battery_columns)}, PCS: {len(pcs_columns)}")
        
        # 在主线程中更新UI
        def update_ui():
            # 存储列名
            self.battery_columns = battery_columns
            self.pcs_columns = pcs_columns
            
            # 根据当前选择的数据类型更新列名
            data_type = self.data_type_var.get()
            if data_type == 'BATTERY':
                self.current_columns = battery_columns
            else:
                self.current_columns = pcs_columns
            
            # 更新所有条件的参数界面（重新创建以更新列名）
            condition_count = len(self.condition_types) if hasattr(self, 'condition_types') else 0
            for i in range(condition_count):
                condition_type = self.condition_types[i].get()
                self._create_params_for_type(i, condition_type)
            
            # 更新关键值追踪下拉框的选项（过滤掉时间列）
            if hasattr(self, 'key_value_combos'):
                self.batch_available_columns = self._filter_time_columns(self.current_columns)
                for combo in self.key_value_combos:
                    # combo 现在是容器，需要调用 update_columns 方法
                    if hasattr(combo, 'update_columns'):
                        combo.update_columns(self.batch_available_columns)
            
            # 更新进度
            self.batch_progress_label.config(text=f"扫描完成！电池列: {len(battery_columns)}, PCS列: {len(pcs_columns)}")
            print(f"[INFO] 列名加载完成 - 电池: {len(battery_columns)}, PCS列: {len(pcs_columns)}")
        
        self.root.after(0, update_ui)
    
    def run_batch_analysis(self):
        """执行批量分析"""
        if not hasattr(self, 'valid_excel_files') or not self.valid_excel_files:
            messagebox.showwarning("警告", "请先扫描Excel文件！")
            return
        
        # 获取当前选择的数据类型
        data_type = self.data_type_var.get()
        
        # 获取用户选择的关键值列
        key_columns = self._get_selected_key_columns()
        
        # 收集启用的条件
        active_conditions = []
        condition_count = len(self.condition_enabled)  # 使用实际条件数量
        
        for i in range(condition_count):
            if self.condition_enabled[i].get():
                # 获取条件类型
                condition_type = self.condition_types[i].get()
                
                # 获取列选择
                col1 = self.condition_columns[i][0].get()
                if not col1 and condition_type != '自定义':
                    messagebox.showwarning("警告", f"请为条件{i+1}选择数据列！")
                    return
                
                # 对于自定义类型，检查是否已设置公式
                if condition_type == '自定义':
                    if not hasattr(self, 'custom_formulas') or i >= len(self.custom_formulas) or not self.custom_formulas[i]:
                        messagebox.showwarning("警告", f"请先为条件{i+1}设置自定义公式！")
                        return
                    condition = {
                        'type': condition_type,
                        'custom_formula': self.custom_formulas[i]
                    }
                else:
                    condition = {
                        'type': condition_type,
                        'columns': [
                            self.condition_columns[i][0].get(),
                            self.condition_columns[i][1].get(),
                            self.condition_columns[i][2].get() if len(self.condition_columns[i]) > 2 else '',
                            self.condition_columns[i][3].get() if len(self.condition_columns[i]) > 3 else '',
                            self.condition_columns[i][4].get() if len(self.condition_columns[i]) > 4 else ''
                        ],
                        'operators': [
                            self.condition_operators[i][0].get(),
                            self.condition_operators[i][1].get()
                        ],
                        'compare': self.condition_compare_type[i].get(),
                        'threshold': self.condition_threshold[i].get(),
                        'min_val': self.condition_min_max[i][0].get() if i < len(self.condition_min_max) else "0",
                        'max_val': self.condition_min_max[i][1].get() if i < len(self.condition_min_max) else "10000"
                    }
                active_conditions.append(condition)
        
        if not active_conditions:
            messagebox.showwarning("警告", "请至少启用一个异常条件！")
            return
        
        # 过滤出当前类型的文件
        files_to_analyze = [f for f in self.valid_excel_files if f['file_type'] == data_type]
        
        if not files_to_analyze:
            messagebox.showwarning("警告", f"没有{data_type}类型的文件可分析！")
            return
        
        # 清空结果
        self.result_tree.delete(*self.result_tree.get_children())
        self.batch_results = []
        
        # 初始化暂停控制
        import threading
        import time
        self.pause_event = threading.Event()
        self.pause_event.set()  # 默认不暂停
        self.batch_paused = False
        self.batch_running = True
        
        # 更新按钮状态
        self.start_btn.config(state=tk.DISABLED)
        self.pause_btn.config(state=tk.NORMAL, text="暂停")
        self.stop_btn.config(state=tk.NORMAL)
        self.batch_progress_label.config(text="分析进行中...")
        self.batch_progress_bar['value'] = 0  # 重置进度条
        self.batch_time_label.config(text="耗时: 0.0秒", foreground="gray")  # 重置耗时
        
        # 【性能优化】收集所有需要读取的列名
        columns_needed = set()
        for condition in active_conditions:
            if condition['type'] != '自定义':
                for col in condition.get('columns', []):
                    if col:
                        columns_needed.add(col)
            else:
                # 自定义条件
                custom_formula = condition.get('custom_formula', {})
                for elem in custom_formula.get('elements', []):
                    if elem.get('type') == 'column':
                        columns_needed.add(elem.get('value'))
                for item in custom_formula.get('items', []):
                    if item.get('column'):
                        columns_needed.add(item.get('column'))
        
        # 添加关键值列
        if key_columns:
            columns_needed.update(key_columns)
        
        # 添加时间列（通常是第一列）
        columns_needed.add(None)  # None 表示读取所有列的第一列作为时间列
        
        print(f"[PERF] 需要读取的列: {len(columns_needed)} 个（含时间列）")
        
        # 记录开始时间并启动耗时更新
        self._batch_start_time = time.time()
        self._batch_time_active = True
        self._update_batch_time()
        
        # 【性能优化】使用线程池并行处理多个文件
        from concurrent.futures import ThreadPoolExecutor, as_completed
        import os
        
        total = len(files_to_analyze)
        completed_count = 0
        anomaly_count = 0
        lock = threading.Lock()  # 用于线程安全的计数器
        
        # 更新进度显示
        self.root.after(0, lambda: self._update_batch_progress(0, f"准备并行分析 {total} 个文件..."))
        
        def analyze_file_wrapper(file_info, idx):
            """包装函数，用于并行处理"""
            # 检查暂停状态
            if self.pause_event:
                self.pause_event.wait()
            
            # 检查是否已停止
            if not self.batch_running:
                return None, idx
            
            # 分析文件
            try:
                file_results = self._analyze_single_file(
                    file_info, active_conditions, key_columns, columns_needed
                )
                return file_results, idx
            except Exception as e:
                print(f"[ERROR] 分析文件失败 {file_info['filename']}: {e}")
                return None, idx
        
        def analyze_thread():
            nonlocal completed_count, anomaly_count
            
            # 使用线程池，最多4个线程并行
            max_workers = min(4, total)  # 最多4个线程，或文件数
            
            with ThreadPoolExecutor(max_workers=max_workers) as executor:
                # 提交所有任务
                future_to_file = {
                    executor.submit(analyze_file_wrapper, file_info, idx): file_info
                    for idx, file_info in enumerate(files_to_analyze)
                }
                
                # 处理完成的任务
                for future in as_completed(future_to_file):
                    if not self.batch_running:
                        break
                    
                    file_info = future_to_file[future]
                    try:
                        file_results, idx = future.result()
                        
                        with lock:
                            completed_count += 1
                            current_progress = (completed_count / total) * 100
                            
                            if file_results:
                                anomaly_count += 1
                                self.batch_results.extend(file_results)
                                # 在主线程中更新结果列表
                                self.root.after(0, lambda r=file_results: self._add_results_to_tree(r))
                            
                            # 更新进度
                            self.root.after(0, lambda p=current_progress, c=completed_count, t=total, f=file_info['filename']: 
                                self._update_batch_progress(p, f"已完成 {c}/{t}: {f}"))
                    
                    except Exception as e:
                        print(f"[ERROR] 处理结果失败: {e}")
            
            # 分析完成
            self.batch_running = False
            self._batch_time_active = False
            # 计算总耗时
            total_time = time.time() - self._batch_start_time
            self.root.after(0, lambda: self._batch_analysis_complete(total, anomaly_count, total_time))
        
        thread = threading.Thread(target=analyze_thread)
        thread.daemon = True
        thread.start()
    
    def _update_batch_progress(self, progress, text):
        """更新批量分析进度"""
        self.batch_progress_bar['value'] = progress
        self.batch_progress_label.config(text=text)
        # 【优化】移除 update_idletasks() 调用，让 Tkinter 在事件循环中自动处理更新
    
    def _update_batch_time(self):
        """实时更新批量分析耗时"""
        if not getattr(self, '_batch_time_active', False):
            return
        
        import time
        elapsed = time.time() - self._batch_start_time
        self.batch_time_label.config(text=f"耗时: {elapsed:.1f}秒", foreground="gray")
        
        # 每500ms更新一次
        self._batch_time_timer = self.root.after(500, self._update_batch_time)
    
    def _analyze_single_file(self, file_info, conditions, key_columns=None, columns_needed=None):
        """分析单个文件（性能优化版）
        
        Args:
            file_info: 文件信息字典
            conditions: 异常条件列表
            key_columns: 要追踪的关键值列列表（可选）
            columns_needed: 需要读取的列名集合（可选，用于只读取需要的列）
        """
        import time
        start_time = time.time()
        
        results = []
        relative_path = file_info.get('relative_path', file_info['filename'])
        
        try:
            # 【性能优化】只读取需要的列
            t1 = time.time()
            
            # 首先读取第一列（时间列）+ 需要的列
            usecols = None
            if columns_needed and len(columns_needed) > 1:
                # 过滤掉 None
                cols_list = [c for c in columns_needed if c is not None]
                if cols_list:
                    usecols = cols_list
            
            df = None
            try:
                # 【性能优化】根据文件类型选择读取方式
                file_ext = os.path.splitext(file_info['path'])[1].lower()
                
                if file_ext == '.csv':
                    # CSV 文件：尝试自动检测编码
                    for encoding in ['utf-8', 'gbk', 'gb2312', 'latin1']:
                        try:
                            if usecols:
                                df = pd.read_csv(file_info['path'], encoding=encoding, usecols=usecols)
                            else:
                                df = pd.read_csv(file_info['path'], encoding=encoding)
                            break
                        except UnicodeDecodeError:
                            continue
                    if df is None:
                        raise Exception("无法自动检测CSV编码")
                else:
                    # Excel 文件：使用 usecols 只读取需要的列
                    if usecols:
                        # 先读取第一列
                        df_temp = pd.read_excel(file_info['path'], engine='calamine', nrows=0)
                        first_col = df_temp.columns[0] if len(df_temp.columns) > 0 else None
                        
                        # 构建最终需要读取的列（第一列 + 条件列）
                        final_cols = [first_col] if first_col else []
                        for col in usecols:
                            if col and col != first_col:
                                final_cols.append(col)
                        
                        df = pd.read_excel(file_info['path'], engine='calamine', usecols=final_cols)
                    else:
                        df = pd.read_excel(file_info['path'], engine='calamine')
            except Exception as e:
                if file_ext != '.csv':
                    print(f"[WARN] calamine引擎读取失败，尝试openpyxl: {e}")
                    try:
                        if usecols:
                            df = pd.read_excel(file_info['path'], engine='openpyxl', usecols=usecols)
                        else:
                            df = pd.read_excel(file_info['path'], engine='openpyxl')
                    except:
                        df = pd.read_excel(file_info['path'])
                else:
                    raise e
            
            read_time = time.time() - t1
            
            # 获取时间列（通常是第一列）
            time_column = df.columns[0] if len(df.columns) > 0 else None
            
            # 【性能优化】预先缓存数值转换结果，避免重复转换
            t2 = time.time()
            numeric_cache = {}  # 缓存已转换的数值列
            
            def get_numeric_column(col_name):
                """获取数值化的列（带缓存）"""
                if col_name not in df.columns:
                    return None
                if col_name in numeric_cache:
                    return numeric_cache[col_name]
                
                col_series = df[col_name]
                if pd.api.types.is_numeric_dtype(col_series):
                    numeric_cache[col_name] = col_series
                else:
                    # 尝试转换为数值
                    converted = pd.to_numeric(col_series, errors='coerce')
                    if converted.notna().sum() > 0:
                        numeric_cache[col_name] = converted
                    else:
                        # 无法转换为数值，作为分类数据处理
                        unique_values = col_series.dropna().unique()
                        value_map = {v: i for i, v in enumerate(unique_values)}
                        numeric_cache[col_name] = col_series.map(value_map).astype(float)
                return numeric_cache[col_name]
            
            cache_time = time.time() - t2
            
            # 提取关键值数据（使用缓存的数值列）
            t3 = time.time()
            key_values = {}
            if key_columns:
                for col in key_columns:
                    if col in df.columns:
                        numeric_col = get_numeric_column(col)
                        if numeric_col is not None:
                            non_null = numeric_col.dropna()
                            if len(non_null) > 0:
                                key_values[col] = f"{non_null.mean():.2f}"
                            else:
                                key_values[col] = "无数据"
                        else:
                            key_values[col] = "列不存在"
            key_time = time.time() - t3
            
            # 辅助函数：获取异常点对应的时间
            def get_anomaly_times(indices, df, time_col):
                """根据异常索引获取对应的时间值"""
                if not time_col or time_col not in df.columns:
                    return []
                try:
                    times = df.loc[indices, time_col].tolist()
                    # 转换为字符串，处理各种时间格式
                    time_strs = []
                    for t in times:
                        if pd.isna(t):
                            continue
                        # 如果是 datetime 类型，格式化
                        if hasattr(t, 'strftime'):
                            time_strs.append(t.strftime('%Y-%m-%d %H:%M:%S'))
                        else:
                            time_strs.append(str(t))
                    return time_strs
                except Exception:
                    return []
            
            # 辅助函数：获取异常点的具体值
            def get_anomaly_values(indices, df, columns):
                """根据异常索引获取对应的具体值"""
                values_list = []
                try:
                    for idx in indices:
                        if idx not in df.index:
                            continue
                        values = {}
                        for col in columns:
                            if col in df.columns:
                                val = df.loc[idx, col]
                                if pd.isna(val):
                                    values[col] = "NaN"
                                elif isinstance(val, float):
                                    values[col] = f"{val:.4f}"
                                else:
                                    values[col] = str(val)
                        values_list.append(values)
                except Exception as e:
                    print(f"[WARN] 获取异常值失败: {e}")
                return values_list
            
            # 执行条件检查
            t4 = time.time()
            for idx, condition in enumerate(conditions):
                condition_type = condition['type']
                
                # 兼容旧名称
                old_to_new = {
                    '波动剧烈': '波动剧烈检测',
                    '超出范围': '数据范围异常',
                    '同增同减异常': '同增同减检测'
                }
                condition_type = old_to_new.get(condition_type, condition_type)
                
                # 获取列选择（仅非自定义类型需要）
                if condition_type != '自定义':
                    columns_list = condition.get('columns', ['', '', '', '', ''])
                    col1 = columns_list[0] if len(columns_list) > 0 else ''
                    col2 = columns_list[1] if len(columns_list) > 1 else ''
                    col3 = columns_list[2] if len(columns_list) > 2 else ''
                
                # 根据条件类型执行不同的检查
                if condition_type == '波动剧烈检测':
                    if not col1 or col1 not in df.columns:
                        continue
                    threshold = float(condition['threshold'])
                    result = self._check_volatility_fast(df, col1, threshold, get_numeric_column)
                    if result:
                        # 获取异常时间和具体值
                        anomaly_indices = result.get('indices', [])
                        anomaly_times = get_anomaly_times(anomaly_indices, df, time_column)
                        anomaly_values = get_anomaly_values(anomaly_indices, df, [col1])
                        results.append({
                            'filename': file_info['filename'],
                            'device_code': file_info['device_code'],
                            'relative_path': relative_path,
                            'anomaly_type': '波动剧烈检测',
                            'anomaly_detail': f"列'{col1}'有{result['count']}个点变化率超过{threshold}%阈值",
                            'key_values': key_values,
                            'file_path': file_info['path'],
                            'anomaly_times': anomaly_times,
                            'anomaly_column': col1,
                            'anomaly_values': anomaly_values
                        })
                
                elif condition_type == '数据范围异常':
                    if not col1 or col1 not in df.columns:
                        continue
                    min_val = float(condition['min_val'])
                    max_val = float(condition['max_val'])
                    result = self._check_range_fast(df, col1, min_val, max_val, get_numeric_column)
                    if result:
                        # 获取异常时间和具体值
                        anomaly_indices = result.get('indices', [])
                        anomaly_times = get_anomaly_times(anomaly_indices, df, time_column)
                        anomaly_values = get_anomaly_values(anomaly_indices, df, [col1])
                        results.append({
                            'filename': file_info['filename'],
                            'device_code': file_info['device_code'],
                            'relative_path': relative_path,
                            'anomaly_type': '数据范围异常',
                            'anomaly_detail': f"列'{col1}'有{result['count']}个点超出[{min_val}, {max_val}]范围",
                            'key_values': key_values,
                            'file_path': file_info['path'],
                            'anomaly_times': anomaly_times,
                            'anomaly_column': col1,
                            'anomaly_values': anomaly_values
                        })
                
                elif condition_type == '同增同减检测':
                    # 获取所有非空的列（最多5列）
                    cols = [c for c in condition.get('columns', ['', '', '', '', '']) if c and c in df.columns]
                    
                    if len(cols) < 2:
                        continue
                    
                    # 检查所有列是否同增同减
                    result = self._check_all_columns_correlation(df, cols, get_numeric_column)
                    
                    if result and result['count'] > 0:
                        # 构建异常详情
                        cols_str = '、'.join([f"'{c}'" for c in cols])
                        detail = f"列{cols_str}存在{result['count']}个同增同减异常点"
                        
                        # 获取异常时间和具体值
                        anomaly_indices = result.get('indices', [])
                        anomaly_times = get_anomaly_times(anomaly_indices, df, time_column)
                        anomaly_values = get_anomaly_values(anomaly_indices, df, cols)
                        
                        results.append({
                            'filename': file_info['filename'],
                            'device_code': file_info['device_code'],
                            'relative_path': relative_path,
                            'anomaly_type': '同增同减检测',
                            'anomaly_detail': detail,
                            'key_values': key_values,
                            'file_path': file_info['path'],
                            'anomaly_times': anomaly_times,
                            'anomaly_column': ', '.join(cols),
                            'anomaly_values': anomaly_values
                        })
                
                elif condition_type == '多条件组合':
                    # 多条件组合检测：支持 AND/OR 逻辑组合多个条件
                    mc_data = condition.get('multi_condition', {})
                    col1 = mc_data.get('column', '')
                    logic = mc_data.get('logic', 'AND')
                    cond_list = mc_data.get('conditions', [])
                    
                    if not col1 or col1 not in df.columns or not cond_list:
                        continue
                    
                    # 获取数值列
                    numeric_col = get_numeric_column(col1)
                    if numeric_col is None:
                        continue
                    
                    # 构建条件掩码
                    import numpy as np
                    masks = []
                    cond_details = []
                    
                    for cond in cond_list:
                        op = cond.get('op', '!=')
                        val_str = cond.get('value', '')
                        
                        try:
                            val = float(val_str)
                        except ValueError:
                            continue
                        
                        # 根据运算符构建掩码
                        if op == '=':
                            mask = numeric_col == val
                        elif op == '!=':
                            mask = numeric_col != val
                        elif op == '>':
                            mask = numeric_col > val
                        elif op == '<':
                            mask = numeric_col < val
                        elif op == '>=':
                            mask = numeric_col >= val
                        elif op == '<=':
                            mask = numeric_col <= val
                        else:
                            continue
                        
                        masks.append(mask)
                        cond_details.append(f"{col1}{op}{val}")
                    
                    if not masks:
                        continue
                    
                    # 根据逻辑关系组合掩码
                    if logic == 'AND':
                        # 所有条件都满足才标记为异常
                        combined_mask = masks[0]
                        for m in masks[1:]:
                            combined_mask = combined_mask & m
                    else:  # OR
                        # 任一条件满足就标记为异常
                        combined_mask = masks[0]
                        for m in masks[1:]:
                            combined_mask = combined_mask | m
                    
                    # 只保留有效值的异常
                    combined_mask = combined_mask & numeric_col.notna()
                    anomaly_indices = df.index[combined_mask].tolist()
                    
                    if anomaly_indices:
                        # 构建异常详情
                        logic_str = " 且 " if logic == 'AND' else " 或 "
                        detail = f"列'{col1}'满足条件: {logic_str.join(cond_details)}"
                        
                        # 获取异常时间和具体值
                        anomaly_times = get_anomaly_times(anomaly_indices, df, time_column)
                        anomaly_values = get_anomaly_values(anomaly_indices, df, [col1])
                        
                        results.append({
                            'filename': file_info['filename'],
                            'device_code': file_info['device_code'],
                            'relative_path': relative_path,
                            'anomaly_type': '多条件组合',
                            'anomaly_detail': detail,
                            'key_values': key_values,
                            'file_path': file_info['path'],
                            'anomaly_times': anomaly_times,
                            'anomaly_column': col1,
                            'anomaly_values': anomaly_values
                        })
                
                elif condition_type == '自定义':
                    # 使用新的公式格式
                    custom_formula = condition.get('custom_formula')
                    if not custom_formula:
                        continue
                    
                    compare_op = custom_formula.get('compare', '>')
                    threshold = float(custom_formula.get('threshold', 0))
                    
                    try:
                        # 支持新格式（elements）和旧格式（items）
                        elements = custom_formula.get('elements', [])
                        items = custom_formula.get('items', [])
                        
                        if elements:
                            # 新格式：使用 elements 列表
                            # 检查所有列是否存在
                            valid = True
                            for elem in elements:
                                if elem['type'] == 'column' and elem['value'] not in df.columns:
                                    valid = False
                                    break
                            
                            if not valid:
                                continue
                            
                            # 构建并执行公式（使用缓存的数值化 DataFrame）
                            expr_str = ""
                            has_logic_op = False  # 标记是否包含逻辑运算符
                            for elem in elements:
                                if elem['type'] == 'column':
                                    expr_str += f"get_numeric_column('{elem['value']}')"
                                elif elem['type'] == 'operator':
                                    # 逻辑运算符转换为 pandas 位运算符
                                    if elem['value'] == 'AND':
                                        expr_str += " & "
                                        has_logic_op = True
                                    elif elem['value'] == 'OR':
                                        expr_str += " | "
                                        has_logic_op = True
                                    elif elem['value'] == 'NOT':
                                        expr_str += " ~ "
                                        has_logic_op = True
                                    else:
                                        expr_str += f" {elem['value']} "
                                elif elem['type'] == 'constant':
                                    expr_str += str(elem['value'])
                                elif elem['type'] == 'parenthesis':
                                    expr_str += elem['value']
                            
                            if not expr_str:
                                continue
                            
                            # 安全执行表达式
                            try:
                                result_series = eval(expr_str)
                            except Exception as e:
                                print(f"[ERROR] 公式执行失败: {expr_str}, {e}")
                                continue
                            
                        elif items:
                            # 旧格式：使用 items 列表
                            valid = True
                            for item in items:
                                col = item['column']
                                if col not in df.columns:
                                    valid = False
                                    break
                            
                            if not valid:
                                continue
                            
                            # 计算公式结果（使用缓存的数值列）
                            result_series = None
                            expr_parts = []
                            
                            for i, item in enumerate(items):
                                col = item['column']
                                op = item.get('operator', '+')
                                
                                if i == 0:
                                    expr_parts.append(f"get_numeric_column('{col}')")
                                else:
                                    expr_parts.append(f"{op} get_numeric_column('{col}')")
                            
                            expr_str = " ".join(expr_parts)
                            try:
                                result_series = eval(expr_str)
                            except Exception as e:
                                print(f"[ERROR] 公式执行失败: {expr_str}, {e}")
                                continue
                        else:
                            continue
                        
                        # 执行比较
                        if result_series is not None:
                            if compare_op == '>':
                                anomalies = (result_series > threshold).sum()
                            elif compare_op == '<':
                                anomalies = (result_series < threshold).sum()
                            elif compare_op == '>=':
                                anomalies = (result_series >= threshold).sum()
                            elif compare_op == '<=':
                                anomalies = (result_series <= threshold).sum()
                            elif compare_op == '==':
                                anomalies = (result_series == threshold).sum()
                            elif compare_op == '!=':
                                anomalies = (result_series != threshold).sum()
                            else:
                                anomalies = 0
                            
                            if anomalies > 0:
                                results.append({
                                    'filename': file_info['filename'],
                                    'device_code': file_info['device_code'],
                                    'relative_path': relative_path,
                                    'anomaly_type': '自定义条件',
                                    'anomaly_detail': f"有{anomalies}个点满足条件",
                                    'key_values': key_values,
                                    'file_path': file_info['path']
                                })
                    except Exception as e:
                        print(f"[ERROR] 自定义条件检查失败: {e}")
            
            check_time = time.time() - t4
            total_time = time.time() - start_time
            
            # 输出性能日志（仅当耗时超过1秒时）
            if total_time > 1.0:
                print(f"[PERF] {file_info['filename']}: 总计{total_time:.2f}s (读取:{read_time:.2f}s, 缓存:{cache_time:.2f}s, 关键值:{key_time:.2f}s, 检查:{check_time:.2f}s)")
            
        except Exception as e:
            print(f"[ERROR] 分析文件失败 {file_info['filename']}: {e}")
            import traceback
            traceback.print_exc()
        
        return results
    
    def _check_volatility(self, df, column, threshold):
        """检查波动剧烈（使用向量化操作优化）
        
        检测逻辑：比较相邻时刻的变化率
        - 计算每个数据点相对于前一个数据点的变化百分比
        - 如果变化率超过阈值，则标记为异常
        """
        try:
            col_series = df[column]
            if pd.api.types.is_numeric_dtype(col_series):
                valid_data = col_series.dropna()
            else:
                # 尝试转换为数值
                valid_data = pd.to_numeric(col_series, errors='coerce').dropna()
                if len(valid_data) == 0:
                    # 无法转换为数值，作为分类数据处理
                    unique_values = col_series.dropna().unique()
                    value_map = {v: i for i, v in enumerate(unique_values)}
                    valid_data = col_series.map(value_map).dropna().astype(float)
        except Exception:
            return None
        
        if len(valid_data) < 2:
            return None
        
        # 【性能优化】使用向量化操作计算相邻时刻变化率
        # change_rate = |当前值 - 前一值| / |前一值| * 100
        values = valid_data.values
        prev_values = values[:-1]
        curr_values = values[1:]
        
        # 避免除以零
        with np.errstate(divide='ignore', invalid='ignore'):
            change_rates = np.abs((curr_values - prev_values) / np.abs(prev_values)) * 100
            # 处理前一值为0的情况
            change_rates = np.where(np.isinf(change_rates), np.nan, change_rates)
        
        # 找出变化率超过阈值的点
        anomaly_count = np.sum(change_rates > threshold)
        
        if anomaly_count > 0:
            return {'count': int(anomaly_count)}
        return None
    
    def _check_range(self, df, column, min_val, max_val):
        """检查超出范围"""
        # 处理数据类型
        try:
            col_series = df[column]
            if pd.api.types.is_numeric_dtype(col_series):
                valid_data = col_series.dropna()
            else:
                # 尝试转换为数值
                valid_data = pd.to_numeric(col_series, errors='coerce').dropna()
                if len(valid_data) == 0:
                    # 无法转换为数值，作为分类数据处理
                    unique_values = col_series.dropna().unique()
                    value_map = {v: i for i, v in enumerate(unique_values)}
                    valid_data = col_series.map(value_map).dropna().astype(float)
        except Exception:
            return None
        
        if len(valid_data) == 0:
            return None
        
        out_of_range = valid_data[(valid_data < min_val) | (valid_data > max_val)]
        
        if len(out_of_range) > 0:
            return {'count': len(out_of_range)}
        return None
    
    def _check_correlation(self, df, column1, column2):
        """检查同增同减异常（使用向量化操作优化）"""
        try:
            # 处理第一列
            col1_series = df[column1]
            if pd.api.types.is_numeric_dtype(col1_series):
                col1_numeric = col1_series
            else:
                # 尝试转换为数值
                col1_numeric = pd.to_numeric(col1_series, errors='coerce')
                if col1_numeric.notna().sum() == 0:
                    # 无法转换为数值，作为分类数据处理
                    unique_values = col1_series.dropna().unique()
                    value_map = {v: i for i, v in enumerate(unique_values)}
                    col1_numeric = col1_series.map(value_map).astype(float)
            
            # 处理第二列
            col2_series = df[column2]
            if pd.api.types.is_numeric_dtype(col2_series):
                col2_numeric = col2_series
            else:
                # 尝试转换为数值
                col2_numeric = pd.to_numeric(col2_series, errors='coerce')
                if col2_numeric.notna().sum() == 0:
                    # 无法转换为数值，作为分类数据处理
                    unique_values = col2_series.dropna().unique()
                    value_map = {v: i for i, v in enumerate(unique_values)}
                    col2_numeric = col2_series.map(value_map).astype(float)
            
            valid_mask = col1_numeric.notna() & col2_numeric.notna()
        except Exception as e:
            print(f"[DEBUG] _check_correlation 处理异常: {e}")
            return None
        
        if valid_mask.sum() < 2:
            return None
        
        data1 = col1_numeric[valid_mask].reset_index(drop=True)
        data2 = col2_numeric[valid_mask].reset_index(drop=True)
        
        diff1 = data1.diff().dropna()
        diff2 = data2.diff().dropna()
        
        # 【性能优化】使用向量化操作替代 for 循环
        # 计算异常条件：
        # 1. d1 == 0 且 d2 == 0 -> 正常（跳过）
        # 2. d1 == 0 或 d2 == 0 -> 异常
        # 3. d1 > 0 且 d2 < 0 或 d1 < 0 且 d2 > 0 -> 异常（方向相反）
        
        both_zero = (diff1 == 0) & (diff2 == 0)
        one_zero = (diff1 == 0) | (diff2 == 0)
        opposite_direction = (diff1 * diff2) < 0  # 异号表示方向相反
        
        # 异常条件：不是both_zero 且 (one_zero 或 opposite_direction)
        anomalies = (~both_zero & (one_zero | opposite_direction)).sum()
        
        if anomalies > 0:
            return {'count': int(anomalies)}
        return None
    
    def _check_volatility_fast(self, df, column, threshold, get_numeric_column):
        """检查波动剧烈（使用缓存的数值列）
        
        检测逻辑：比较相邻时刻的变化率
        - 计算每个数据点相对于前一个数据点的变化百分比
        - 如果变化率超过阈值，则标记为异常
        
        Returns:
            dict with 'count' and 'indices' (异常点的原始索引列表) or None
        """
        try:
            numeric_col = get_numeric_column(column)
            if numeric_col is None:
                return None
            
            valid_data = numeric_col.dropna()
            if len(valid_data) < 2:
                return None
            
            # 保留原始索引用于映射
            original_indices = valid_data.index.tolist()
            
            # 计算相邻时刻的变化率
            # change_rate = (当前值 - 前一个值) / abs(前一个值) * 100
            values = valid_data.values
            prev_values = values[:-1]
            curr_values = values[1:]
            
            # 避免除以零
            with np.errstate(divide='ignore', invalid='ignore'):
                change_rates = np.abs((curr_values - prev_values) / np.abs(prev_values)) * 100
                # 处理前一个值为0的情况（变化率为无穷大）
                change_rates = np.where(np.isinf(change_rates), np.nan, change_rates)
            
            # 找出变化率超过阈值的点
            # 第一个数据点没有前一个点，所以从索引1开始
            anomaly_mask = change_rates > threshold
            
            if np.any(anomaly_mask):
                # 获取异常点在原始 df 中的索引（索引+1因为diff会少一个元素）
                anomaly_positions = np.where(anomaly_mask)[0]
                anomaly_indices = [original_indices[i + 1] for i in anomaly_positions]
                return {'count': len(anomaly_indices), 'indices': anomaly_indices}
        except Exception:
            pass
        return None
    
    def _check_range_fast(self, df, column, min_val, max_val, get_numeric_column):
        """检查超出范围（使用缓存的数值列）
        
        Returns:
            dict with 'count' and 'indices' (异常点的原始索引列表) or None
        """
        try:
            numeric_col = get_numeric_column(column)
            if numeric_col is None:
                return None
            
            valid_data = numeric_col.dropna()
            if len(valid_data) == 0:
                return None
            
            out_of_range = (valid_data < min_val) | (valid_data > max_val)
            anomaly_count = out_of_range.sum()
            
            if anomaly_count > 0:
                # 获取异常点在原始 df 中的索引
                anomaly_indices = valid_data[out_of_range].index.tolist()
                return {'count': int(anomaly_count), 'indices': anomaly_indices}
        except Exception:
            pass
        return None
    
    def _check_correlation_fast(self, df, column1, column2, get_numeric_column):
        """检查同增同减异常（使用缓存的数值列）
        
        Returns:
            dict with 'count' and 'indices' (异常点的原始索引列表) or None
        """
        try:
            col1_numeric = get_numeric_column(column1)
            col2_numeric = get_numeric_column(column2)
            
            if col1_numeric is None or col2_numeric is None:
                return None
            
            valid_mask = col1_numeric.notna() & col2_numeric.notna()
            if valid_mask.sum() < 2:
                return None
            
            # 保留原始索引用于映射
            original_indices = col1_numeric[valid_mask].index.tolist()
            
            data1 = col1_numeric[valid_mask].reset_index(drop=True)
            data2 = col2_numeric[valid_mask].reset_index(drop=True)
            
            diff1 = data1.diff().dropna()
            diff2 = data2.diff().dropna()
            
            # 【性能优化】使用向量化操作
            both_zero = (diff1 == 0) & (diff2 == 0)
            one_zero = (diff1 == 0) | (diff2 == 0)
            opposite_direction = (diff1 * diff2) < 0
            
            anomaly_mask = ~both_zero & (one_zero | opposite_direction)
            anomalies = anomaly_mask.sum()
            
            if anomalies > 0:
                # diff 的索引从1开始（因为 diff() 会少一个元素）
                # 所以 diff 结果的索引 i 对应原始数据的索引 i+1
                diff_anomaly_indices = anomaly_mask[anomaly_mask].index.tolist()
                # 映射回原始索引（diff 的第 i 个元素对应原始数据的第 i+1 个元素）
                original_anomaly_indices = [original_indices[i + 1] for i in diff_anomaly_indices if i + 1 < len(original_indices)]
                return {'count': int(anomalies), 'indices': original_anomaly_indices}
        except Exception:
            pass
        return None
    
    def _check_all_columns_correlation(self, df, columns, get_numeric_column):
        """检查所有列是否同增同减
        
        同增同减的定义：所有列必须同时增加、同时减少或同时不变。
        如果有一列的变化方向与其他列不一致，就是异常。
        
        Args:
            df: DataFrame
            columns: 要检查的列名列表（至少2列）
            get_numeric_column: 获取数值列的函数
            
        Returns:
            dict with 'count' and 'indices' (异常点的原始索引列表) or None
        """
        try:
            if len(columns) < 2:
                return None
            
            # 获取所有列的数值数据
            numeric_cols = []
            for col in columns:
                numeric_col = get_numeric_column(col)
                if numeric_col is None:
                    return None
                numeric_cols.append(numeric_col)
            
            # 找到所有列都有有效值的行
            valid_mask = numeric_cols[0].notna()
            for col in numeric_cols[1:]:
                valid_mask = valid_mask & col.notna()
            
            if valid_mask.sum() < 2:
                return None
            
            # 保留原始索引用于映射
            original_indices = numeric_cols[0][valid_mask].index.tolist()
            
            # 计算每列的变化量（diff）
            diffs = []
            for col in numeric_cols:
                diff = col[valid_mask].reset_index(drop=True).diff().dropna()
                diffs.append(diff)
            
            # 检查每一行的变化方向是否一致
            # 方向：增加(>0)、减少(<0)、不变(=0)
            # 异常条件：不是所有列都同方向变化
            
            # 获取每一行的变化方向
            # sign: 1=增加, -1=减少, 0=不变
            import numpy as np
            
            n_rows = len(diffs[0])
            anomaly_indices = []
            
            for i in range(n_rows):
                # 获取该行所有列的变化方向
                directions = []
                for diff in diffs:
                    val = diff.iloc[i]
                    if val > 0:
                        directions.append(1)  # 增加
                    elif val < 0:
                        directions.append(-1)  # 减少
                    else:
                        directions.append(0)  # 不变
                
                # 检查是否所有方向一致
                # 规则：非零的方向必须全部相同
                non_zero_dirs = [d for d in directions if d != 0]
                
                if len(non_zero_dirs) == 0:
                    # 所有列都不变，正常
                    continue
                elif len(non_zero_dirs) == len(directions):
                    # 所有列都有变化，检查方向是否一致
                    if len(set(non_zero_dirs)) > 1:
                        # 方向不一致，异常
                        anomaly_indices.append(i)
                else:
                    # 部分列变化，部分列不变，异常
                    anomaly_indices.append(i)
            
            if len(anomaly_indices) > 0:
                # 映射回原始索引（diff 的第 i 个元素对应原始数据的第 i+1 个元素）
                original_anomaly_indices = [original_indices[i + 1] for i in anomaly_indices if i + 1 < len(original_indices)]
                return {'count': len(original_anomaly_indices), 'indices': original_anomaly_indices}
            
            return None
        except Exception:
            pass
        return None
    
    def _add_results_to_tree(self, results):
        """添加结果到树形列表"""
        for result in results:
            # 获取关键值（最多显示3个）
            key_values = result.get('key_values', {})
            key_cols = list(key_values.keys()) if key_values else []
            kv1 = f"{key_cols[0]}={key_values.get(key_cols[0], '')}" if len(key_cols) > 0 else ""
            kv2 = f"{key_cols[1]}={key_values.get(key_cols[1], '')}" if len(key_cols) > 1 else ""
            kv3 = f"{key_cols[2]}={key_values.get(key_cols[2], '')}" if len(key_cols) > 2 else ""
            
            self.result_tree.insert('', tk.END, values=(
                result['filename'],
                result.get('relative_path', ''),
                result['device_code'],
                result['anomaly_type'],
                result['anomaly_detail'],
                kv1,
                kv2,
                kv3
            ))
    
    def _update_anomaly_time_panel(self):
        """更新异常时间面板，按文件名分组显示异常时间"""
        # 清空现有的异常时间树
        self.anomaly_time_tree.delete(*self.anomaly_time_tree.get_children())
        
        # 按文件名分组
        file_groups = {}
        for result in self.batch_results:
            filename = result['filename']
            if filename not in file_groups:
                file_groups[filename] = []
            file_groups[filename].append(result)
        
        # 按文件名添加到树中
        for filename, results in file_groups.items():
            # 收集该文件的所有异常时间
            all_times = []
            for result in results:
                anomaly_times = result.get('anomaly_times', [])
                anomaly_type = result['anomaly_type']
                anomaly_column = result.get('anomaly_column', '')
                anomaly_values = result.get('anomaly_values', [])
                
                # 将时间、值配对
                for i, time_str in enumerate(anomaly_times):
                    # 获取对应的异常值
                    if i < len(anomaly_values):
                        values_dict = anomaly_values[i]
                        # 构建值字符串（使用 | 分隔，更清晰）
                        value_parts = [f"{col}={val}" for col, val in values_dict.items()]
                        value_str = ' | '.join(value_parts)
                    else:
                        value_str = ''
                    
                    all_times.append({
                        'time': time_str,
                        'type': anomaly_type,
                        'column': anomaly_column,
                        'detail': value_str  # 使用具体值代替重复的描述
                    })
            
            # 如果没有异常时间，跳过
            if not all_times:
                continue
            
            # 创建父节点（文件名）
            # 统计各类型异常数量
            type_counts = {}
            for t in all_times:
                type_counts[t['type']] = type_counts.get(t['type'], 0) + 1
            type_summary = ', '.join([f"{k}: {v}" for k, v in type_counts.items()])
            
            parent_id = self.anomaly_time_tree.insert('', tk.END, text=f"{filename} ({type_summary})", values=('', '', '', f'共{len(all_times)}个异常点'))
            
            # 添加子节点（异常时间）
            for time_info in all_times:
                self.anomaly_time_tree.insert(parent_id, tk.END, text='', values=(
                    time_info['time'],
                    time_info['type'],
                    time_info['column'],
                    time_info['detail']
                ))
    
    def _on_result_double_click(self, event):
        """双击异常文件列表项，跳转到单文件分析"""
        # 获取选中的项
        selection = self.result_tree.selection()
        if not selection:
            return
        
        # 获取选中项的索引
        item_id = selection[0]
        item_index = self.result_tree.index(item_id)
        
        # 从 batch_results 获取对应的结果
        if not hasattr(self, 'batch_results') or item_index >= len(self.batch_results):
            return
        
        result = self.batch_results[item_index]
        file_path = result.get('file_path')
        
        if not file_path or not os.path.exists(file_path):
            messagebox.showerror("错误", f"文件不存在: {file_path}")
            return
        
        # 切换到单文件分析标签页
        self.main_notebook.select(self.single_file_tab)
        
        # 【关键】立即更新 UI 显示加载状态（不要等到后台线程）
        self.file_path_label.config(text=file_path)
        self.progress_label.config(text="正在加载文件...", foreground="blue")
        self.progress_detail.config(text="请稍候，正在读取文件...")
        self.time_label.config(text="耗时: 0.0秒", foreground="gray")
        self.root.update_idletasks()
        
        # 启动模拟进度动画
        self._start_simulated_progress(
            '正在读取文件',
            start_value=5,
            end_value=60,
            duration=20,
            show_time=True
        )
        
        # 记录开始时间
        import time
        self._load_start_time = time.time()
        
        # 在后台线程中加载文件（耗时操作在后台，UI更新在主线程）
        def load_in_background():
            try:
                start_time = time.time()
                
                # 【性能优化】根据文件类型选择读取方式
                file_ext = os.path.splitext(file_path)[1].lower()
                
                if file_ext == '.csv':
                    # CSV 文件：尝试自动检测编码
                    for encoding in ['utf-8', 'gbk', 'gb2312', 'latin1']:
                        try:
                            df = pd.read_csv(file_path, encoding=encoding)
                            break
                        except UnicodeDecodeError:
                            continue
                    if df is None:
                        raise Exception("无法自动检测CSV编码")
                else:
                    # Excel 文件：使用快速引擎加载
                    try:
                        df = pd.read_excel(file_path, engine='calamine')
                    except:
                        try:
                            df = pd.read_excel(file_path, engine='openpyxl')
                        except:
                            df = pd.read_excel(file_path)
                
                read_time = time.time() - start_time
                print(f"[PERF] _quick_load 后台读取耗时: {read_time:.2f}秒")
                
                # 检测时间列
                time_column = None
                for col in df.columns:
                    if '时间' in col or 'time' in col.lower() or 'date' in col.lower():
                        time_column = col
                        break
                
                if time_column is None:
                    time_column = df.columns[0]
                
                # 获取异常列和关键值列（在后台线程准备好）
                anomaly_columns = []
                key_value_columns = []
                
                if hasattr(self, 'batch_results') and self.batch_results:
                    if item_index < len(self.batch_results):
                        result = self.batch_results[item_index]
                        anomaly_detail = result.get('anomaly_detail', '')
                        
                        # 从异常详情中提取列名
                        import re
                        matches = re.findall(r"'([^']+)'", anomaly_detail)
                        anomaly_columns = matches
                        
                        # 获取关键值列
                        key_values = result.get('key_values', {})
                        key_value_columns = list(key_values.keys())
                
                # 在主线程中更新 UI
                def update_ui():
                    # 停止模拟进度
                    self._stop_simulated_progress()
                    
                    self.file_path = file_path
                    self.df = df
                    self.time_column = time_column
                    
                    # 更新进度标签显示加载状态
                    self.progress_label.config(text=f"正在加载: {os.path.basename(file_path)}", foreground="blue")
                    self.root.update_idletasks()
                    
                    # 刷新列列表（现在使用虚拟列表，很快）
                    self.refresh_column_list()
                    
                    # 启用生成按钮
                    self.generate_btn.config(state=tk.NORMAL)
                    self.save_btn.config(state=tk.DISABLED)
                    
                    # 勾选异常列
                    for col_name in anomaly_columns:
                        if col_name in self.column_vars:
                            self.column_vars[col_name].set(1)
                    
                    # 勾选关键值列
                    for col_name in key_value_columns:
                        if col_name in self.column_vars:
                            self.column_vars[col_name].set(1)
                    
                    # 计算总耗时
                    total_time = time.time() - self._load_start_time
                    
                    # 更新状态
                    self._update_progress_ui("准备完成", 100, f'已完成，共 {len(df.columns)} 列数据', total_time)
                    
                    # 自动绘制曲线
                    self.root.after(100, self.generate_plot)
                
                self.root.after(0, update_ui)
                
            except Exception as e:
                import traceback
                traceback.print_exc()
                self.root.after(0, lambda: (
                    self._stop_simulated_progress(),
                    messagebox.showerror("错误", f"加载文件失败: {e}")
                ))
        
        import threading
        thread = threading.Thread(target=load_in_background, daemon=True)
        thread.start()
    
    def _quick_load_file(self, file_path):
        """快速加载文件并准备绘制（从批量分析跳转时调用）"""
        try:
            # 更新文件路径显示
            self.file_path = file_path
            self.file_path_label.config(text=file_path)
            
            # 更新进度标签显示加载状态
            self.progress_label.config(text=f"正在加载: {os.path.basename(file_path)}", foreground="blue")
            self.root.update_idletasks()
            
            # 【性能优化】根据文件类型选择读取方式
            file_ext = os.path.splitext(file_path)[1].lower()
            
            if file_ext == '.csv':
                # CSV 文件：尝试自动检测编码
                for encoding in ['utf-8', 'gbk', 'gb2312', 'latin1']:
                    try:
                        self.df = pd.read_csv(file_path, encoding=encoding)
                        break
                    except UnicodeDecodeError:
                        continue
                if self.df is None:
                    raise Exception("无法自动检测CSV编码")
            else:
                # Excel 文件：使用快速引擎加载
                try:
                    self.df = pd.read_excel(file_path, engine='calamine')
                except:
                    try:
                        self.df = pd.read_excel(file_path, engine='openpyxl')
                    except:
                        self.df = pd.read_excel(file_path)
            
            # 检测时间列
            self.time_column = None
            for col in self.df.columns:
                if '时间' in col or 'time' in col.lower() or 'date' in col.lower():
                    self.time_column = col
                    break
            
            if self.time_column is None:
                self.time_column = self.df.columns[0]
            
            # 刷新列列表（会清空 column_vars 并重新创建）
            self.refresh_column_list()
            
            # 启用生成按钮
            self.generate_btn.config(state=tk.NORMAL)
            self.save_btn.config(state=tk.DISABLED)
            
            # 更新状态
            self.progress_label.config(text=f"已加载: {os.path.basename(file_path)}", foreground="blue")
            
            # 获取批量分析中选择的异常条件列，自动勾选
            if hasattr(self, 'batch_results') and self.batch_results:
                # 获取当前选中行的异常条件
                selection = self.result_tree.selection()
                if selection:
                    item_id = selection[0]
                    item_index = self.result_tree.index(item_id)
                    if item_index < len(self.batch_results):
                        result = self.batch_results[item_index]
                        anomaly_detail = result.get('anomaly_detail', '')
                        
                        # 从异常详情中提取列名（简单解析）
                        # 例如："列'DSP软件版本(10111)'有3个点超过30%阈值"
                        # 或 "列'A'与'B'有5个不一致点"
                        import re
                        # 匹配单引号中的列名
                        matches = re.findall(r"'([^']+)'", anomaly_detail)
                        
                        # 勾选匹配的列
                        for col_name in matches:
                            if col_name in self.column_vars:
                                self.column_vars[col_name].set(1)
            
            # 同时勾选关键值列
            if hasattr(self, 'batch_results') and self.batch_results:
                selection = self.result_tree.selection()
                if selection:
                    item_id = selection[0]
                    item_index = self.result_tree.index(item_id)
                    if item_index < len(self.batch_results):
                        result = self.batch_results[item_index]
                        key_values = result.get('key_values', {})
                        for col_name in key_values.keys():
                            if col_name in self.column_vars:
                                self.column_vars[col_name].set(1)
            
            # 自动绘制曲线
            self.root.after(100, self.generate_plot)  # 延迟一点确保勾选状态已更新
            
        except Exception as e:
            messagebox.showerror("错误", f"快速加载文件失败: {e}")
            import traceback
            traceback.print_exc()
    
    def _batch_analysis_complete(self, total, anomaly_count, total_time=0):
        """批量分析完成"""
        # 停止耗时更新
        self._batch_time_active = False
        if hasattr(self, '_batch_time_timer'):
            try:
                self.root.after_cancel(self._batch_time_timer)
            except:
                pass
        
        self.batch_progress_bar['value'] = 100
        self.batch_progress_label.config(text="分析完成！")
        
        # 显示总耗时
        if total_time > 0:
            self.batch_time_label.config(text=f"耗时: {total_time:.1f}秒", foreground="green")
        
        # 重置按钮状态
        self.batch_running = False
        self.start_btn.config(state=tk.NORMAL)
        self.pause_btn.config(state=tk.DISABLED, text="暂停")
        self.stop_btn.config(state=tk.DISABLED)
        
        normal_count = total - anomaly_count
        stats_text = f"共分析 {total} 个文件\n异常文件: {anomaly_count} 个\n正常文件: {normal_count} 个"
        self.batch_stats_label.config(text=stats_text)
        
        # 更新异常时间面板
        if self.batch_results:
            self._update_anomaly_time_panel()
        
        messagebox.showinfo("分析完成", 
            f"批量分析完成！\n\n"
            f"总文件数: {total}\n"
            f"异常文件: {anomaly_count}\n"
            f"正常文件: {normal_count}")
    
    def export_batch_results(self):
        """导出批量分析结果"""
        if not self.batch_results:
            messagebox.showwarning("警告", "没有分析结果可导出！")
            return
        
        file_path = filedialog.asksaveasfilename(
            title="保存分析结果",
            defaultextension=".xlsx",
            filetypes=[("Excel文件", "*.xlsx"), ("CSV文件", "*.csv"), ("文本文件", "*.txt")]
        )
        
        if file_path:
            try:
                if file_path.endswith('.xlsx'):
                    # 按文件名分组，每个文件一个sheet
                    from openpyxl import Workbook
                    from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
                    
                    # 按文件名分组
                    file_groups = {}
                    for result in self.batch_results:
                        filename = result['filename']
                        if filename not in file_groups:
                            file_groups[filename] = []
                        file_groups[filename].append(result)
                    
                    wb = Workbook()
                    # 删除默认sheet
                    default_sheet = wb.active
                    wb.remove(default_sheet)
                    
                    # 设置样式
                    header_font = Font(bold=True)
                    header_fill = PatternFill(start_color="CCE5FF", end_color="CCE5FF", fill_type="solid")
                    header_alignment = Alignment(horizontal='center', vertical='center')
                    thin_border = Border(
                        left=Side(style='thin'),
                        right=Side(style='thin'),
                        top=Side(style='thin'),
                        bottom=Side(style='thin')
                    )
                    
                    # 为每个文件创建sheet
                    for filename, results in file_groups.items():
                        # sheet名最长31个字符（Excel限制）
                        sheet_name = filename[:31] if len(filename) > 31 else filename
                        ws = wb.create_sheet(title=sheet_name)
                        
                        # 写入标题行
                        headers = ['时间', '异常类型', '异常详情']
                        for col_idx, header in enumerate(headers, 1):
                            cell = ws.cell(row=1, column=col_idx, value=header)
                            cell.font = header_font
                            cell.fill = header_fill
                            cell.alignment = header_alignment
                            cell.border = thin_border
                        
                        # 写入数据行
                        row_idx = 2
                        for result in results:
                            anomaly_times = result.get('anomaly_times', [])
                            anomaly_values = result.get('anomaly_values', [])
                            anomaly_type = result.get('anomaly_type', '')
                            anomaly_column = result.get('anomaly_column', '')
                            
                            # 构建异常详情（参数名=值格式）
                            def format_values(values_dict):
                                """格式化参数值字典为可读字符串"""
                                if not values_dict:
                                    return '-'
                                parts = []
                                for col_name, val in values_dict.items():
                                    parts.append(f"{col_name}={val}")
                                return ' | '.join(parts)  # 使用 | 分隔，更清晰
                            
                            if not anomaly_times:
                                # 没有具体时间点，写入一行汇总
                                time_str = '-'
                                detail_str = result.get('anomaly_detail', '')
                                row_data = [time_str, anomaly_type, detail_str]
                                
                                for col_idx, value in enumerate(row_data, 1):
                                    cell = ws.cell(row=row_idx, column=col_idx, value=value)
                                    cell.border = thin_border
                                    cell.alignment = Alignment(vertical='center', wrap_text=True)
                                row_idx += 1
                            else:
                                # 有具体时间点，展开每行
                                for i, anomaly_time in enumerate(anomaly_times):
                                    # 时间
                                    time_str = anomaly_time.strftime('%Y-%m-%d %H:%M:%S') if hasattr(anomaly_time, 'strftime') else str(anomaly_time)
                                    
                                    # 异常详情：参数名=值, 参数名=值
                                    if i < len(anomaly_values) and anomaly_values[i]:
                                        detail_str = format_values(anomaly_values[i])
                                    else:
                                        detail_str = result.get('anomaly_detail', '')
                                    
                                    row_data = [time_str, anomaly_type, detail_str]
                                    
                                    for col_idx, value in enumerate(row_data, 1):
                                        cell = ws.cell(row=row_idx, column=col_idx, value=value)
                                        cell.border = thin_border
                                        cell.alignment = Alignment(vertical='center', wrap_text=True)
                                    row_idx += 1
                        
                        # 设置列宽
                        ws.column_dimensions['A'].width = 20  # 时间
                        ws.column_dimensions['B'].width = 15  # 异常类型
                        ws.column_dimensions['C'].width = 80  # 异常详情
                    
                    wb.save(file_path)
                    total_rows = sum(len(r.get('anomaly_times', []) or [1]) for r in self.batch_results)
                    messagebox.showinfo("成功", f"结果已保存至：\n{file_path}\n\n共 {len(file_groups)} 个文件，{total_rows} 条记录")
                    
                elif file_path.endswith('.csv'):
                    # CSV格式：时间, 异常类型, 异常详情
                    import csv
                    with open(file_path, 'w', newline='', encoding='utf-8-sig') as f:
                        writer = csv.writer(f)
                        writer.writerow(['时间', '异常类型', '异常详情'])
                        for result in self.batch_results:
                            anomaly_times = result.get('anomaly_times', [])
                            anomaly_values = result.get('anomaly_values', [])
                            anomaly_type = result.get('anomaly_type', '')
                            
                            def format_values(values_dict):
                                if not values_dict:
                                    return '-'
                                parts = []
                                for col_name, val in values_dict.items():
                                    parts.append(f"{col_name}={val}")
                                return ', '.join(parts)
                            
                            if not anomaly_times:
                                writer.writerow(['-', anomaly_type, result.get('anomaly_detail', '')])
                            else:
                                for i, anomaly_time in enumerate(anomaly_times):
                                    time_str = anomaly_time.strftime('%Y-%m-%d %H:%M:%S') if hasattr(anomaly_time, 'strftime') else str(anomaly_time)
                                    if i < len(anomaly_values) and anomaly_values[i]:
                                        detail_str = format_values(anomaly_values[i])
                                    else:
                                        detail_str = result.get('anomaly_detail', '')
                                    writer.writerow([time_str, anomaly_type, detail_str])
                    messagebox.showinfo("成功", f"结果已保存至：\n{file_path}")
                else:
                    # 文本格式
                    with open(file_path, 'w', encoding='utf-8') as f:
                        f.write("时间\t异常类型\t异常详情\n")
                        for result in self.batch_results:
                            anomaly_times = result.get('anomaly_times', [])
                            anomaly_values = result.get('anomaly_values', [])
                            anomaly_type = result.get('anomaly_type', '')
                            
                            def format_values(values_dict):
                                if not values_dict:
                                    return '-'
                                parts = []
                                for col_name, val in values_dict.items():
                                    parts.append(f"{col_name}={val}")
                                return ', '.join(parts)
                            
                            if not anomaly_times:
                                f.write(f"-\t{anomaly_type}\t{result.get('anomaly_detail', '')}\n")
                            else:
                                for i, anomaly_time in enumerate(anomaly_times):
                                    time_str = anomaly_time.strftime('%Y-%m-%d %H:%M:%S') if hasattr(anomaly_time, 'strftime') else str(anomaly_time)
                                    if i < len(anomaly_values) and anomaly_values[i]:
                                        detail_str = format_values(anomaly_values[i])
                                    else:
                                        detail_str = result.get('anomaly_detail', '')
                                    f.write(f"{time_str}\t{anomaly_type}\t{detail_str}\n")
                    messagebox.showinfo("成功", f"结果已保存至：\n{file_path}")
                    
            except Exception as e:
                import traceback
                traceback.print_exc()
                messagebox.showerror("错误", f"导出失败：{e}")
    
    def clear_batch_results(self):
        """清空批量分析结果"""
        self.result_tree.delete(*self.result_tree.get_children())
        self.batch_results = []
        self.batch_progress_bar['value'] = 0
        self.batch_progress_label.config(text="就绪")
        self.batch_time_label.config(text="耗时: 0.0秒", foreground="gray")
        self.batch_stats_label.config(text="尚未开始分析")
        messagebox.showinfo("提示", "结果已清空")
    
    def toggle_pause_batch(self):
        """暂停/继续批量分析"""
        if not self.batch_running:
            return
        
        if self.batch_paused:
            # 继续
            self.batch_paused = False
            self.pause_btn.config(text="暂停")
            self.batch_progress_label.config(text="分析进行中...")
            if self.pause_event:
                self.pause_event.set()  # 释放暂停
        else:
            # 暂停
            self.batch_paused = True
            self.pause_btn.config(text="继续")
            self.batch_progress_label.config(text="已暂停，点击继续...")
            if self.pause_event:
                self.pause_event.clear()  # 设置暂停
    
    def stop_batch_analysis(self):
        """停止批量分析"""
        if not self.batch_running:
            return
        
        if messagebox.askyesno("确认", "确定要停止分析吗？\n已分析的结果将保留。"):
            self.batch_running = False
            self.batch_paused = False
            
            # 停止耗时更新并显示已用耗时
            self._batch_time_active = False
            if hasattr(self, '_batch_time_timer'):
                try:
                    self.root.after_cancel(self._batch_time_timer)
                except:
                    pass
            if hasattr(self, '_batch_start_time'):
                import time
                elapsed = time.time() - self._batch_start_time
                self.batch_time_label.config(text=f"耗时: {elapsed:.1f}秒", foreground="orange")
            
            # 如果是暂停状态，先解除暂停
            if self.pause_event:
                self.pause_event.set()
            
            # 更新按钮状态
            self.start_btn.config(state=tk.NORMAL)
            self.pause_btn.config(state=tk.DISABLED, text="暂停")
            self.stop_btn.config(state=tk.DISABLED)
            self.batch_progress_label.config(text="已停止")


def main():
    root = tk.Tk()
    
    # 设置窗口图标（如果有）
    # root.iconbitmap('icon.ico')
    
    app = ExcelDataVisualizer(root)
    root.mainloop()


if __name__ == "__main__":
    main()

    # ==================== 自定义计算列功能 ====================
    

