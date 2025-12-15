"""
CSV数据分析工具 - tkinter GUI 主程序
通用数据分析与可视化工具，UI与逻辑完全分离
所有分析由 analysis_backend 完成
"""

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
import os
from analysis_backend import DataAnalyzer

# 设置matplotlib中文字体
def setup_matplotlib_fonts():
    """设置matplotlib中文显示字体"""
    font_candidates = [
        'Microsoft YaHei',  # 微软雅黑
        'SimHei',           # 黑体
        'KaiTi',            # 楷体
        'FangSong',         # 仿宋
        'STHeiti',          # 华文黑体
    ]
    
    available_fonts = [f.name for f in fm.fontManager.ttflist]
    
    for font in font_candidates:
        if font in available_fonts:
            plt.rcParams['font.sans-serif'] = [font]
            break
    
    plt.rcParams['axes.unicode_minus'] = False

setup_matplotlib_fonts()


class App:
    """CSV数据分析应用主类"""
    
    def __init__(self, root):
        """
        初始化应用
        
        参数:
            root: tkinter根窗口
        """
        self.root = root
        self.root.title("CSV数据分析工具 v1.1")
        self.root.geometry("1200x800")
        
        # 初始化后端
        self.backend = DataAnalyzer()
        
        # 当前加载的文件路径
        self.current_file = None
        self.last_dir = os.path.dirname(os.path.abspath(__file__))
        
        # 当前画布引用（用于清除）
        self.current_canvas = None
        self.status_var = tk.StringVar(value="未加载数据")
        
        # 建立UI
        self._setup_ui()
    
    
    def _setup_ui(self):
        """构建整个UI布局"""
        
        # ========== 顶部框 (Frame 1): 文件加载 ==========
        top_frame = ttk.Frame(self.root, relief=tk.SUNKEN)
        top_frame.pack(side=tk.TOP, fill=tk.X, padx=5, pady=5)
        
        ttk.Label(top_frame, text="数据文件:", font=("微软雅黑", 10)).pack(side=tk.LEFT, padx=5)
        
        self.file_label = ttk.Label(
            top_frame,
            text="未加载任何文件",
            foreground="red",
            font=("微软雅黑", 9)
        )
        self.file_label.pack(side=tk.LEFT, padx=5)
        
        ttk.Button(top_frame, text="📁 加载CSV文件", command=self._load_csv).pack(side=tk.LEFT, padx=5)
        
        # 添加状态标签
        ttk.Label(top_frame, text=" | ", foreground="gray").pack(side=tk.LEFT, padx=2)
        
        self.status_label = ttk.Label(
            top_frame,
            textvariable=self.status_var,
            foreground="blue",
            font=("微软雅黑", 9)
        )
        self.status_label.pack(side=tk.LEFT, padx=5)
        main_container = ttk.Frame(self.root)
        main_container.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # ========== 左侧框 (Frame 2): 变量选择 ==========
        left_frame = ttk.LabelFrame(main_container, text="变量选择", padding=10)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, padx=5)
        
        # 分类型变量
        cat_header = ttk.Frame(left_frame)
        cat_header.pack(fill=tk.X, anchor=tk.W)
        ttk.Label(cat_header, text="分类型变量 (X轴/分组):", font=("微软雅黑", 10, "bold")).pack(side=tk.LEFT)
        ttk.Button(cat_header, text="✕", width=3, command=self._clear_cat_selection).pack(side=tk.RIGHT)
        
        cat_scrollbar = ttk.Scrollbar(left_frame)
        cat_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.cat_listbox = tk.Listbox(
            left_frame,
            height=10,
            width=25,
            yscrollcommand=cat_scrollbar.set,
            selectmode=tk.SINGLE,
            exportselection=False,
            font=("微软雅黑", 9)
        )
        self.cat_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.cat_listbox.bind('<<ListboxSelect>>', self._on_cat_select)
        self.cat_listbox.bind('<Double-Button-1>', self._clear_cat_selection)
        cat_scrollbar.config(command=self.cat_listbox.yview)
        
        cont_header = ttk.Frame(left_frame)
        cont_header.pack(fill=tk.X, anchor=tk.W, pady=(15, 0))
        ttk.Label(cont_header, text="连续型变量 (Ctrl多选):", font=("微软雅黑", 10, "bold")).pack(side=tk.LEFT)
        ttk.Button(cont_header, text="✕", width=3, command=self._clear_cont_selection).pack(side=tk.RIGHT)
        
        cont_scrollbar = ttk.Scrollbar(left_frame)
        cont_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.cont_listbox = tk.Listbox(
            left_frame,
            height=10,
            width=25,
            yscrollcommand=cont_scrollbar.set,
            selectmode=tk.EXTENDED,  # 多选模式，支持配对t检验等需要选择两个变量的场景
            exportselection=False,
            font=("微软雅黑", 9)
        )
        self.cont_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.cont_listbox.bind('<<ListboxSelect>>', self._on_cont_select)
        self.cont_listbox.bind('<Double-Button-1>', self._clear_cont_selection)
        cont_scrollbar.config(command=self.cont_listbox.yview)
        
        # ========== 右侧框 (Frame 3): 绘图和结果显示 ==========
        right_frame = ttk.Frame(main_container)
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=5)
        
        # 使用 Notebook (标签页)
        self.notebook = ttk.Notebook(right_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True)
        
        # Tab 1: 绘图区域
        plot_tab = ttk.Frame(self.notebook)
        self.notebook.add(plot_tab, text="📊 绘图区")
        
        self.plot_frame = tk.Frame(plot_tab, bg="white")
        self.plot_frame.pack(fill=tk.BOTH, expand=True)
        
        # Tab 2: 统计结果
        result_tab = ttk.Frame(self.notebook)
        self.notebook.add(result_tab, text="📈 统计结果")
        
        result_scrollbar = ttk.Scrollbar(result_tab)
        result_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.text_output = tk.Text(
            result_tab,
            height=30,
            width=80,
            yscrollcommand=result_scrollbar.set,
            font=("Courier New", 9),
            bg="#f5f5f5"
        )
        self.text_output.pack(fill=tk.BOTH, expand=True)
        result_scrollbar.config(command=self.text_output.yview)
        
        # ========== 底部框 (Frame 4): 分析操作按钮 ==========
        button_container = ttk.Frame(self.root)
        button_container.pack(fill=tk.X, padx=5, pady=5)
        
        # ===== 第一行：可视化图表 + 描述统计 =====
        vis_row = ttk.Frame(button_container)
        vis_row.pack(fill=tk.X, pady=2)
        
        # 📊 可视化图表
        vis_frame = ttk.LabelFrame(vis_row, text="📊 可视化图表", padding=5)
        vis_frame.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        
        ttk.Button(vis_frame, text="直方图", command=self._draw_histogram, width=10).pack(side=tk.LEFT, padx=2)
        ttk.Button(vis_frame, text="折线图", command=self._draw_line, width=10).pack(side=tk.LEFT, padx=2)
        ttk.Button(vis_frame, text="柱状图", command=self._draw_bar, width=10).pack(side=tk.LEFT, padx=2)
        ttk.Button(vis_frame, text="饼图", command=self._draw_pie, width=10).pack(side=tk.LEFT, padx=2)
        ttk.Button(vis_frame, text="箱线图", command=self._draw_boxplot, width=10).pack(side=tk.LEFT, padx=2)
        ttk.Button(vis_frame, text="小提琴图", command=self._draw_violin, width=10).pack(side=tk.LEFT, padx=2)
        ttk.Button(vis_frame, text="Q-Q图", command=self._draw_qq, width=10).pack(side=tk.LEFT, padx=2)
        ttk.Button(vis_frame, text="散点图", command=self._draw_scatter, width=10).pack(side=tk.LEFT, padx=2)
        ttk.Button(vis_frame, text="热力图", command=self._draw_correlation_heatmap, width=10).pack(side=tk.LEFT, padx=2)
        
        # ===== 第二行：参数检验 + 非参数检验 =====
        test_row = ttk.Frame(button_container)
        test_row.pack(fill=tk.X, pady=2)
        
        # 🔬 参数检验
        param_frame = ttk.LabelFrame(test_row, text="🔬 参数检验", padding=5)
        param_frame.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        
        ttk.Button(param_frame, text="正态性检验", command=self._run_normality_test, width=12).pack(side=tk.LEFT, padx=2)
        ttk.Button(param_frame, text="独立t检验", command=self._run_t_test, width=12).pack(side=tk.LEFT, padx=2)
        ttk.Button(param_frame, text="配对t检验", command=self._run_paired_t_test, width=12).pack(side=tk.LEFT, padx=2)
        ttk.Button(param_frame, text="ANOVA", command=self._run_anova, width=12).pack(side=tk.LEFT, padx=2)
        ttk.Button(param_frame, text="线性回归", command=self._run_linear_regression, width=12).pack(side=tk.LEFT, padx=2)
        
        # 📉 非参数检验
        nonparam_frame = ttk.LabelFrame(test_row, text="📉 非参数检验", padding=5)
        nonparam_frame.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        
        ttk.Button(nonparam_frame, text="卡方检验", command=self._run_chi_square, width=12).pack(side=tk.LEFT, padx=2)
        ttk.Button(nonparam_frame, text="Mann-Whitney", command=self._run_mann_whitney, width=12).pack(side=tk.LEFT, padx=2)
        ttk.Button(nonparam_frame, text="Kruskal-Wallis", command=self._run_kruskal_wallis, width=12).pack(side=tk.LEFT, padx=2)
        
        # ===== 第三行：数据分析 + 导出操作 =====
        data_row = ttk.Frame(button_container)
        data_row.pack(fill=tk.X, pady=2)
        
        # 📋 数据分析
        data_frame = ttk.LabelFrame(data_row, text="📋 数据分析", padding=5)
        data_frame.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        
        ttk.Button(data_frame, text="描述统计", command=self._show_descriptive_stats, width=12).pack(side=tk.LEFT, padx=2)
        ttk.Button(data_frame, text="批量统计", command=self._show_all_stats, width=12).pack(side=tk.LEFT, padx=2)
        ttk.Button(data_frame, text="相关性分析", command=self._show_correlation_analysis, width=12).pack(side=tk.LEFT, padx=2)
        ttk.Button(data_frame, text="缺失值分析", command=self._show_missing_analysis, width=12).pack(side=tk.LEFT, padx=2)
        
        # 💾 导出操作
        export_frame = ttk.LabelFrame(data_row, text="💾 导出操作", padding=5)
        export_frame.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        
        ttk.Button(export_frame, text="保存图形", command=self._save_plot, width=12).pack(side=tk.LEFT, padx=2)
        ttk.Button(export_frame, text="导出Excel", command=self._export_to_excel, width=12).pack(side=tk.LEFT, padx=2)
        ttk.Button(export_frame, text="清空结果", command=self._clear_output, width=12).pack(side=tk.LEFT, padx=2)
        
        # ===== 第四行：高级分析 =====
        adv_row = ttk.Frame(button_container)
        adv_row.pack(fill=tk.X, pady=2)
        
        # 🚀 高级可视化
        adv_vis_frame = ttk.LabelFrame(adv_row, text="🚀 高级可视化", padding=5)
        adv_vis_frame.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        
        ttk.Button(adv_vis_frame, text="3D散点图", command=self._draw_3d_scatter, width=10).pack(side=tk.LEFT, padx=2)
        ttk.Button(adv_vis_frame, text="3D曲面图", command=self._draw_3d_surface, width=10).pack(side=tk.LEFT, padx=2)
        ttk.Button(adv_vis_frame, text="3D散点(Web)", command=self._draw_3d_scatter_plotly, width=12).pack(side=tk.LEFT, padx=2)
        ttk.Button(adv_vis_frame, text="3D曲面(Web)", command=self._draw_3d_surface_plotly, width=12).pack(side=tk.LEFT, padx=2)
        ttk.Button(adv_vis_frame, text="配对图", command=self._draw_pair_grid, width=10).pack(side=tk.LEFT, padx=2)
        ttk.Button(adv_vis_frame, text="雷达图", command=self._draw_radar, width=10).pack(side=tk.LEFT, padx=2)
        ttk.Button(adv_vis_frame, text="分布对比", command=self._draw_distribution_comparison, width=10).pack(side=tk.LEFT, padx=2)
        
        # 🧠 机器学习
        ml_frame = ttk.LabelFrame(adv_row, text="🧠 机器学习", padding=5)
        ml_frame.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        
        ttk.Button(ml_frame, text="PCA 2D", command=self._draw_pca_2d, width=10).pack(side=tk.LEFT, padx=2)
        ttk.Button(ml_frame, text="PCA 3D", command=self._draw_pca_3d, width=10).pack(side=tk.LEFT, padx=2)
        ttk.Button(ml_frame, text="PCA分析", command=self._show_pca_analysis, width=10).pack(side=tk.LEFT, padx=2)
        ttk.Button(ml_frame, text="K-Means", command=self._draw_kmeans, width=10).pack(side=tk.LEFT, padx=2)
        ttk.Button(ml_frame, text="聚类分析", command=self._show_cluster_analysis, width=10).pack(side=tk.LEFT, padx=2)
        ttk.Button(ml_frame, text="树状图", command=self._draw_dendrogram, width=10).pack(side=tk.LEFT, padx=2)
    
    def _on_cat_select(self, event):
        """分类变量选择时的回调"""
        selection = self.cat_listbox.curselection()
        if selection:
            var_name = self.cat_listbox.get(selection[0])
            self.status_var.set(f"✓ 已选择分类变量: {var_name}")
        else:
            self.status_var.set("未选择分类变量")
    
    def _on_cont_select(self, event):
        """连续变量选择时的回调"""
        selection = self.cont_listbox.curselection()
        if selection:
            var_name = self.cont_listbox.get(selection[0])
            self.status_var.set(f"✓ 已选择连续变量: {var_name}")
        else:
            self.status_var.set("未选择连续变量")
    
    def _clear_cat_selection(self, event=None):
        """清除分类变量的选择"""
        self.cat_listbox.selection_clear(0, tk.END)
        self.status_var.set("已清除分类变量选择")
    
    def _clear_cont_selection(self, event=None):
        """清除连续变量的选择"""
        self.cont_listbox.selection_clear(0, tk.END)
        self.status_var.set("已清除连续变量选择")
    
    def _load_csv(self):
        """加载CSV文件的回调函数"""
        filepath = filedialog.askopenfilename(
            title="选择CSV数据文件",
            filetypes=[("CSV文件", "*.csv"), ("所有文件", "*.*")],
            initialdir=self.last_dir
        )
        
        if not filepath:
            return
        
        # 保存目录路径
        self.last_dir = os.path.dirname(filepath)
        
        try:
            # 调用后端加载数据
            var_dict = self.backend.load_data(filepath)
            self.current_file = filepath
            
            # 更新文件标签
            filename = os.path.basename(filepath)
            self.file_label.config(
                text=f"✅ {filename} (行数: {var_dict['shape'][0]}, 列数: {var_dict['shape'][1]})",
                foreground="green"
            )
            
            # 更新变量列表
            self._update_variable_lists(var_dict)
            
            # 显示加载成功信息
            messagebox.showinfo(
                "成功",
                f"数据加载成功！\n\n"
                f"文件: {filename}\n"
                f"行数: {var_dict['shape'][0]}\n"
                f"列数: {var_dict['shape'][1]}\n"
                f"分类型变量: {len(var_dict['categorical'])}\n"
                f"连续型变量: {len(var_dict['continuous'])}"
            )
        
        except Exception as e:
            messagebox.showerror("错误", f"加载失败: {str(e)}")
    
    
    def _update_variable_lists(self, var_dict):
        """更新分类型和连续型变量列表框"""
        # 清空列表
        self.cat_listbox.delete(0, tk.END)
        self.cont_listbox.delete(0, tk.END)
        
        # 添加分类型变量
        for col in var_dict['categorical']:
            self.cat_listbox.insert(tk.END, col)
        
        # 添加连续型变量
        for col in var_dict['continuous']:
            self.cont_listbox.insert(tk.END, col)
    
    
    def _draw_histogram(self):
        """绘制直方图"""
        if self.backend.df is None:
            messagebox.showwarning("警告", "请先加载数据")
            return
        
        selection = self.cont_listbox.curselection()
        if not selection:
            messagebox.showwarning("警告", "请在右侧连续型变量列表中选择一个变量")
            return
        
        column = self.cont_listbox.get(selection[0])
        
        try:
            fig = self.backend.plot_histogram(column)
            self._embed_figure(fig, self.plot_frame)
            self.notebook.select(0)  # 切换到绘图标签页
        except Exception as e:
            messagebox.showerror("错误", f"绘制失败: {str(e)}")
    
    
    def _draw_qq(self):
        """绘制Q-Q图"""
        if self.backend.df is None:
            messagebox.showwarning("警告", "请先加载数据")
            return
        
        selection = self.cont_listbox.curselection()
        if not selection:
            messagebox.showwarning("警告", "请选择一个连续型变量")
            return
        
        column = self.cont_listbox.get(selection[0])
        
        try:
            fig = self.backend.plot_qq(column)
            self._embed_figure(fig, self.plot_frame)
            self.notebook.select(0)
        except Exception as e:
            messagebox.showerror("错误", f"绘制失败: {str(e)}")
    
    
    def _draw_boxplot(self):
        """绘制箱线图"""
        if self.backend.df is None:
            messagebox.showwarning("警告", "请先加载数据")
            return
        
        cat_sel = self.cat_listbox.curselection()
        cont_sel = self.cont_listbox.curselection()
        
        if not cat_sel:
            messagebox.showwarning("警告", "请在左侧分类型变量列表中选择一个变量")
            return
        if not cont_sel:
            messagebox.showwarning("警告", "请在右侧连续型变量列表中选择一个变量")
            return
        
        x_col = self.cat_listbox.get(cat_sel[0])
        y_col = self.cont_listbox.get(cont_sel[0])
        
        try:
            fig = self.backend.plot_boxplot(x_col, y_col)
            self._embed_figure(fig, self.plot_frame)
            self.notebook.select(0)
        except Exception as e:
            messagebox.showerror("错误", f"绘制失败: {str(e)}")
    
    
    def _draw_violin(self):
        """绘制小提琴图"""
        if self.backend.df is None:
            messagebox.showwarning("警告", "请先加载数据")
            return
        
        cat_sel = self.cat_listbox.curselection()
        cont_sel = self.cont_listbox.curselection()
        
        if not cat_sel:
            messagebox.showwarning("警告", "请在左侧分类型变量列表中选择一个变量")
            return
        if not cont_sel:
            messagebox.showwarning("警告", "请在右侧连续型变量列表中选择一个变量")
            return
        
        x_col = self.cat_listbox.get(cat_sel[0])
        y_col = self.cont_listbox.get(cont_sel[0])
        
        try:
            fig = self.backend.plot_violin(x_col, y_col)
            self._embed_figure(fig, self.plot_frame)
            self.notebook.select(0)
        except Exception as e:
            messagebox.showerror("错误", f"绘制失败: {str(e)}")
    
    
    def _draw_line(self):
        """绘制折线图"""
        if self.backend.df is None:
            messagebox.showwarning("警告", "请先加载数据")
            return
        
        cont_sel = self.cont_listbox.curselection()
        if not cont_sel:
            messagebox.showwarning("警告", "请在右侧连续型变量列表中选择一个变量")
            return
        
        y_col = self.cont_listbox.get(cont_sel[0])
        
        try:
            fig = self.backend.plot_line(y_col)
            self._embed_figure(fig, self.plot_frame)
            self.notebook.select(0)
        except Exception as e:
            messagebox.showerror("错误", f"绘制失败: {str(e)}")
    
    
    def _draw_bar(self):
        """绘制柱状图"""
        if self.backend.df is None:
            messagebox.showwarning("警告", "请先加载数据")
            return
        
        # 优先检查分类变量
        cat_sel = self.cat_listbox.curselection()
        cont_sel = self.cont_listbox.curselection()
        
        if cat_sel:
            col = self.cat_listbox.get(cat_sel[0])
        elif cont_sel:
            col = self.cont_listbox.get(cont_sel[0])
        else:
            messagebox.showwarning("警告", "请选择一个变量（分类或连续均可）")
            return
        
        try:
            fig = self.backend.plot_bar(col)
            self._embed_figure(fig, self.plot_frame)
            self.notebook.select(0)
        except Exception as e:
            messagebox.showerror("错误", f"绘制失败: {str(e)}")
    
    
    def _draw_pie(self):
        """绘制饼图"""
        if self.backend.df is None:
            messagebox.showwarning("警告", "请先加载数据")
            return
        
        cat_sel = self.cat_listbox.curselection()
        if not cat_sel:
            messagebox.showwarning("警告", "请在左侧分类型变量列表中选择一个变量")
            return
        
        x_col = self.cat_listbox.get(cat_sel[0])
        
        try:
            fig = self.backend.plot_pie(x_col)
            self._embed_figure(fig, self.plot_frame)
            self.notebook.select(0)
        except Exception as e:
            messagebox.showerror("错误", f"绘制失败: {str(e)}")
    
    
    def _show_descriptive_stats(self):
        """显示描述性统计"""
        if self.backend.df is None:
            messagebox.showwarning("警告", "请先加载数据")
            return
        
        selection = self.cont_listbox.curselection()
        if not selection:
            messagebox.showwarning("警告", "请在右侧连续型变量列表中选择一个变量")
            return
        
        column = self.cont_listbox.get(selection[0])
        
        try:
            result = self.backend.get_descriptive_stats(column)
            self._display_text_result(result)
        except Exception as e:
            messagebox.showerror("错误", f"计算失败: {str(e)}")
    
    
    def _run_t_test(self):
        """执行t检验"""
        if self.backend.df is None:
            messagebox.showwarning("警告", "请先加载数据")
            return
        
        cat_sel = self.cat_listbox.curselection()
        cont_sel = self.cont_listbox.curselection()
        
        if not cat_sel:
            messagebox.showwarning("警告", "请在左侧分类型变量列表中选择一个变量")
            return
        if not cont_sel:
            messagebox.showwarning("警告", "请在右侧连续型变量列表中选择一个变量")
            return
        
        cat_col = self.cat_listbox.get(cat_sel[0])
        cont_col = self.cont_listbox.get(cont_sel[0])
        
        try:
            result = self.backend.run_t_test(cat_col, cont_col)
            self._display_text_result(result)
        except Exception as e:
            messagebox.showerror("错误", f"检验失败: {str(e)}")
    
    # ==================== 新增学术分析功能 ====================
    
    def _draw_correlation_heatmap(self):
        """绘制相关性热力图"""
        if self.backend.df is None:
            messagebox.showwarning("警告", "请先加载数据")
            return
        
        if len(self.backend.continuous_cols) < 2:
            messagebox.showwarning("警告", "至少需要2个连续变量")
            return
        
        try:
            fig = self.backend.plot_correlation_heatmap()
            self._embed_figure(fig, self.plot_frame)
            self.notebook.select(0)
        except Exception as e:
            messagebox.showerror("错误", f"绘制失败: {str(e)}")
    
    
    def _draw_scatter(self):
        """绘制散点图 - 需要选择两个连续变量"""
        if self.backend.df is None:
            messagebox.showwarning("警告", "请先加载数据")
            return
        
        # 尝试获取连续变量列表中的选择
        cont_sel = self.cont_listbox.curselection()
        
        if len(cont_sel) < 2:
            # 如果选择不足2个，弹出对话框让用户选择
            if len(self.backend.continuous_cols) < 2:
                messagebox.showwarning("警告", "至少需要2个连续变量")
                return
            
            # 使用前两个连续变量作为默认
            x_col = self.backend.continuous_cols[0]
            y_col = self.backend.continuous_cols[1]
            messagebox.showinfo("提示", f"使用默认变量:\nX: {x_col}\nY: {y_col}\n\n提示: 可按Ctrl键多选两个连续变量")
        else:
            x_col = self.cont_listbox.get(cont_sel[0])
            y_col = self.cont_listbox.get(cont_sel[1])
        
        try:
            fig = self.backend.plot_scatter(x_col, y_col)
            self._embed_figure(fig, self.plot_frame)
            self.notebook.select(0)
        except Exception as e:
            messagebox.showerror("错误", f"绘制失败: {str(e)}")
    
    
    def _run_linear_regression(self):
        """执行线性回归"""
        if self.backend.df is None:
            messagebox.showwarning("警告", "请先加载数据")
            return
        
        cont_sel = self.cont_listbox.curselection()
        
        if len(cont_sel) < 2:
            if len(self.backend.continuous_cols) < 2:
                messagebox.showwarning("警告", "至少需要2个连续变量")
                return
            x_col = self.backend.continuous_cols[0]
            y_col = self.backend.continuous_cols[1]
            messagebox.showinfo("提示", f"使用默认变量:\nX(自变量): {x_col}\nY(因变量): {y_col}")
        else:
            x_col = self.cont_listbox.get(cont_sel[0])
            y_col = self.cont_listbox.get(cont_sel[1])
        
        try:
            result = self.backend.linear_regression(x_col, y_col)
            self._display_text_result(result)
            
            # 同时绘制散点图
            fig = self.backend.plot_scatter(x_col, y_col, add_regression=True)
            self._embed_figure(fig, self.plot_frame)
        except Exception as e:
            messagebox.showerror("错误", f"分析失败: {str(e)}")
    
    
    def _run_anova(self):
        """执行单因素方差分析"""
        if self.backend.df is None:
            messagebox.showwarning("警告", "请先加载数据")
            return
        
        cat_sel = self.cat_listbox.curselection()
        cont_sel = self.cont_listbox.curselection()
        
        if not cat_sel:
            messagebox.showwarning("警告", "请在左侧选择分类变量(分组因素)")
            return
        if not cont_sel:
            messagebox.showwarning("警告", "请在右侧选择连续变量(因变量)")
            return
        
        cat_col = self.cat_listbox.get(cat_sel[0])
        cont_col = self.cont_listbox.get(cont_sel[0])
        
        try:
            result = self.backend.one_way_anova(cat_col, cont_col)
            self._display_text_result(result)
        except Exception as e:
            messagebox.showerror("错误", f"分析失败: {str(e)}")
    
    
    def _run_chi_square(self):
        """执行卡方检验"""
        if self.backend.df is None:
            messagebox.showwarning("警告", "请先加载数据")
            return
        
        cat_sel = self.cat_listbox.curselection()
        
        if len(cat_sel) < 2:
            if len(self.backend.categorical_cols) < 2:
                messagebox.showwarning("警告", "至少需要2个分类变量")
                return
            col1 = self.backend.categorical_cols[0]
            col2 = self.backend.categorical_cols[1] if len(self.backend.categorical_cols) > 1 else self.backend.continuous_cols[0]
            messagebox.showinfo("提示", f"使用默认变量:\n变量1: {col1}\n变量2: {col2}")
        else:
            col1 = self.cat_listbox.get(cat_sel[0])
            col2 = self.cat_listbox.get(cat_sel[1])
        
        try:
            result = self.backend.chi_square_test(col1, col2)
            self._display_text_result(result)
        except Exception as e:
            messagebox.showerror("错误", f"检验失败: {str(e)}")
    
    
    def _run_normality_test(self):
        """执行正态性检验"""
        if self.backend.df is None:
            messagebox.showwarning("警告", "请先加载数据")
            return
        
        cont_sel = self.cont_listbox.curselection()
        if not cont_sel:
            messagebox.showwarning("警告", "请在右侧选择一个连续变量")
            return
        
        column = self.cont_listbox.get(cont_sel[0])
        
        try:
            result = self.backend.normality_test(column)
            self._display_text_result(result)
        except Exception as e:
            messagebox.showerror("错误", f"检验失败: {str(e)}")
    
    
    def _run_paired_t_test(self):
        """执行配对t检验"""
        if self.backend.df is None:
            messagebox.showwarning("警告", "请先加载数据")
            return
        
        cont_sel = self.cont_listbox.curselection()
        
        if len(cont_sel) < 2:
            messagebox.showwarning("警告", "请按Ctrl键选择两个连续变量(前测/后测)")
            return
        
        col1 = self.cont_listbox.get(cont_sel[0])
        col2 = self.cont_listbox.get(cont_sel[1])
        
        try:
            result = self.backend.paired_t_test(col1, col2)
            self._display_text_result(result)
        except Exception as e:
            messagebox.showerror("错误", f"检验失败: {str(e)}")
    
    
    def _run_mann_whitney(self):
        """执行Mann-Whitney U检验"""
        if self.backend.df is None:
            messagebox.showwarning("警告", "请先加载数据")
            return
        
        cat_sel = self.cat_listbox.curselection()
        cont_sel = self.cont_listbox.curselection()
        
        if not cat_sel:
            messagebox.showwarning("警告", "请在左侧选择分类变量(二分类)")
            return
        if not cont_sel:
            messagebox.showwarning("警告", "请在右侧选择连续变量")
            return
        
        cat_col = self.cat_listbox.get(cat_sel[0])
        cont_col = self.cont_listbox.get(cont_sel[0])
        
        try:
            result = self.backend.mann_whitney_test(cat_col, cont_col)
            self._display_text_result(result)
        except Exception as e:
            messagebox.showerror("错误", f"检验失败: {str(e)}")
    
    
    def _run_kruskal_wallis(self):
        """执行Kruskal-Wallis检验"""
        if self.backend.df is None:
            messagebox.showwarning("警告", "请先加载数据")
            return
        
        cat_sel = self.cat_listbox.curselection()
        cont_sel = self.cont_listbox.curselection()
        
        if not cat_sel:
            messagebox.showwarning("警告", "请在左侧选择分类变量")
            return
        if not cont_sel:
            messagebox.showwarning("警告", "请在右侧选择连续变量")
            return
        
        cat_col = self.cat_listbox.get(cat_sel[0])
        cont_col = self.cont_listbox.get(cont_sel[0])
        
        try:
            result = self.backend.kruskal_wallis_test(cat_col, cont_col)
            self._display_text_result(result)
        except Exception as e:
            messagebox.showerror("错误", f"检验失败: {str(e)}")
    
    
    def _show_all_stats(self):
        """显示所有变量的批量描述统计"""
        if self.backend.df is None:
            messagebox.showwarning("警告", "请先加载数据")
            return
        
        try:
            result = self.backend.get_all_descriptive_stats()
            self._display_text_result(result)
        except Exception as e:
            messagebox.showerror("错误", f"计算失败: {str(e)}")
    
    
    def _show_missing_analysis(self):
        """显示缺失值分析"""
        if self.backend.df is None:
            messagebox.showwarning("警告", "请先加载数据")
            return
        
        try:
            result = self.backend.missing_value_analysis()
            self._display_text_result(result)
        except Exception as e:
            messagebox.showerror("错误", f"分析失败: {str(e)}")
    
    
    def _export_to_excel(self):
        """导出统计结果到Excel"""
        if self.backend.df is None:
            messagebox.showwarning("警告", "请先加载数据")
            return
        
        filepath = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel文件", "*.xlsx"), ("CSV文件", "*.csv"), ("所有文件", "*.*")],
            initialdir=self.last_dir
        )
        
        if filepath:
            try:
                result = self.backend.export_statistics_to_excel(filepath)
                messagebox.showinfo("导出结果", result)
            except Exception as e:
                messagebox.showerror("错误", f"导出失败: {str(e)}")
    
    
    def _show_correlation_analysis(self):
        """显示相关性分析文本结果"""
        if self.backend.df is None:
            messagebox.showwarning("警告", "请先加载数据")
            return
        
        try:
            result = self.backend.get_correlation_analysis()
            self._display_text_result(result)
        except Exception as e:
            messagebox.showerror("错误", f"分析失败: {str(e)}")
    
    
    def _embed_figure(self, fig, frame):
        """
        将matplotlib图形嵌入到tkinter框架中
        
        参数:
            fig: matplotlib Figure对象
            frame: tkinter Frame对象
        """
        # 清除旧的canvas
        for widget in frame.winfo_children():
            widget.destroy()
        
        # 限制图形最大尺寸，避免撑爆界面
        fig_width, fig_height = fig.get_size_inches()
        max_height = 6  # 最大高度6英寸
        if fig_height > max_height:
            scale = max_height / fig_height
            fig.set_size_inches(fig_width * scale, max_height)
            fig.tight_layout()
        
        # 创建新的canvas并绘制
        canvas = FigureCanvasTkAgg(fig, master=frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        
        # 保存引用以便后续保存图形
        self.current_canvas = canvas
    
    
    def _display_text_result(self, text):
        """显示文本结果"""
        self.text_output.config(state=tk.NORMAL)
        self.text_output.delete(1.0, tk.END)
        self.text_output.insert(tk.END, text)
        self.text_output.config(state=tk.DISABLED)
        self.notebook.select(1)  # 切换到结果标签页
    
    
    def _save_plot(self):
        """保存当前图形"""
        if self.current_canvas is None:
            messagebox.showwarning("警告", "没有可保存的图形")
            return
        
        filepath = filedialog.asksaveasfilename(
            defaultextension=".png",
            filetypes=[("PNG图片", "*.png"), ("PDF文件", "*.pdf"), ("所有文件", "*.*")],
            initialdir=os.path.expanduser("~")
        )
        
        if filepath:
            try:
                self.current_canvas.figure.savefig(filepath, dpi=300, bbox_inches='tight')
                messagebox.showinfo("成功", f"图形已保存到:\n{filepath}")
            except Exception as e:
                messagebox.showerror("错误", f"保存失败: {str(e)}")
    
    
    def _clear_output(self):
        """清空输出"""
        self.text_output.config(state=tk.NORMAL)
        self.text_output.delete(1.0, tk.END)
        self.text_output.config(state=tk.DISABLED)

    # ==================== 高级分析功能 ====================
    
    def _draw_3d_scatter(self):
        """绘制3D散点图"""
        if self.backend.df is None:
            messagebox.showwarning("警告", "请先加载数据")
            return
        
        selections = self.cont_listbox.curselection()
        if len(selections) < 3:
            messagebox.showwarning("提示", 
                "请在左侧【连续型变量】列表中选择3个变量作为X/Y/Z轴\n\n"
                "多选方法：\n"
                "• 按住 Ctrl 点击多个变量\n"
                "• 或按住 Shift 选择连续范围\n\n"
                f"当前已选: {len(selections)} 个，还需选择 {3-len(selections)} 个")
            return
        
        x_col = self.cont_listbox.get(selections[0])
        y_col = self.cont_listbox.get(selections[1])
        z_col = self.cont_listbox.get(selections[2])
        
        # 获取可选的分组变量
        cat_selection = self.cat_listbox.curselection()
        hue_col = self.cat_listbox.get(cat_selection[0]) if cat_selection else None
        
        try:
            fig = self.backend.plot_3d_scatter(x_col, y_col, z_col, hue_col)
            self._embed_figure(fig, self.plot_frame)
            self.notebook.select(0)
        except Exception as e:
            messagebox.showerror("错误", f"绑制失败: {str(e)}")
    
    def _draw_3d_surface(self):
        """绘制3D曲面图"""
        if self.backend.df is None:
            messagebox.showwarning("警告", "请先加载数据")
            return
        
        selections = self.cont_listbox.curselection()
        if len(selections) < 3:
            messagebox.showwarning("提示", 
                "请在左侧【连续型变量】列表中选择3个变量作为X/Y/Z轴\n\n"
                "多选方法：\n"
                "• 按住 Ctrl 点击多个变量\n"
                "• 或按住 Shift 选择连续范围\n\n"
                f"当前已选: {len(selections)} 个，还需选择 {3-len(selections)} 个")
            return
        
        x_col = self.cont_listbox.get(selections[0])
        y_col = self.cont_listbox.get(selections[1])
        z_col = self.cont_listbox.get(selections[2])
        
        try:
            fig = self.backend.plot_3d_surface(x_col, y_col, z_col)
            self._embed_figure(fig, self.plot_frame)
            self.notebook.select(0)
        except Exception as e:
            messagebox.showerror("错误", f"绑制失败: {str(e)}")

    def _draw_3d_scatter_plotly(self):
        """绘制3D散点图 (Web/GPU)"""
        if self.backend.df is None:
            messagebox.showwarning("警告", "请先加载数据")
            return
        
        selections = self.cont_listbox.curselection()
        if len(selections) < 3:
            messagebox.showwarning("提示", "请选择3个连续变量作为X/Y/Z轴 (按住Ctrl多选)")
            return
        
        x_col = self.cont_listbox.get(selections[0])
        y_col = self.cont_listbox.get(selections[1])
        z_col = self.cont_listbox.get(selections[2])
        
        cat_selection = self.cat_listbox.curselection()
        hue_col = self.cat_listbox.get(cat_selection[0]) if cat_selection else None
        
        try:
            filename = self.backend.plot_3d_scatter_plotly(x_col, y_col, z_col, hue_col)
            self.status_var.set(f"已在浏览器中打开: {filename}")
        except Exception as e:
            messagebox.showerror("错误", f"生成失败: {str(e)}")

    def _draw_3d_surface_plotly(self):
        """绘制3D曲面图 (Web/GPU)"""
        if self.backend.df is None:
            messagebox.showwarning("警告", "请先加载数据")
            return
        
        selections = self.cont_listbox.curselection()
        if len(selections) < 3:
            messagebox.showwarning("提示", "请选择3个连续变量作为X/Y/Z轴 (按住Ctrl多选)")
            return
        
        x_col = self.cont_listbox.get(selections[0])
        y_col = self.cont_listbox.get(selections[1])
        z_col = self.cont_listbox.get(selections[2])
        
        try:
            filename = self.backend.plot_3d_surface_plotly(x_col, y_col, z_col)
            self.status_var.set(f"已在浏览器中打开: {filename}")
        except Exception as e:
            messagebox.showerror("错误", f"生成失败: {str(e)}")
    
    def _draw_pair_grid(self):
        """绘制配对图矩阵"""
        if self.backend.df is None:
            messagebox.showwarning("警告", "请先加载数据")
            return
        
        # 获取选中的连续变量
        cont_selections = self.cont_listbox.curselection()
        if len(cont_selections) >= 2:
            columns = [self.cont_listbox.get(i) for i in cont_selections]
        else:
            messagebox.showwarning("警告", "请选择至少2个连续变量进行配对分析\n（按住Ctrl多选）")
            return
        
        # 获取可选的分组变量
        cat_selection = self.cat_listbox.curselection()
        hue_col = self.cat_listbox.get(cat_selection[0]) if cat_selection else None
        
        try:
            fig = self.backend.plot_pair_grid(columns=columns, hue_col=hue_col)
            self._embed_figure(fig, self.plot_frame)
            self.notebook.select(0)
        except Exception as e:
            messagebox.showerror("错误", f"绑制失败: {str(e)}")
    
    def _draw_radar(self):
        """绘制雷达图"""
        if self.backend.df is None:
            messagebox.showwarning("警告", "请先加载数据")
            return
        
        # 获取可选的分组变量
        cat_selection = self.cat_listbox.curselection()
        hue_col = self.cat_listbox.get(cat_selection[0]) if cat_selection else None
        
        try:
            fig = self.backend.plot_radar_chart(group_col=hue_col)
            self._embed_figure(fig, self.plot_frame)
            self.notebook.select(0)
        except Exception as e:
            messagebox.showerror("错误", f"绑制失败: {str(e)}")
    
    def _draw_distribution_comparison(self):
        """绘制分布对比图"""
        if self.backend.df is None:
            messagebox.showwarning("警告", "请先加载数据")
            return
        
        cat_selection = self.cat_listbox.curselection()
        cont_selection = self.cont_listbox.curselection()
        
        if not cat_selection or not cont_selection:
            messagebox.showwarning("警告", "请选择一个分类变量(X)和一个连续变量(Y)")
            return
        
        group_col = self.cat_listbox.get(cat_selection[0])
        value_col = self.cont_listbox.get(cont_selection[0])
        
        try:
            fig = self.backend.plot_distribution_comparison(value_col, group_col)
            self._embed_figure(fig, self.plot_frame)
            self.notebook.select(0)
        except Exception as e:
            messagebox.showerror("错误", f"绑制失败: {str(e)}")
    
    def _draw_pca_2d(self):
        """绘制PCA 2D图"""
        if self.backend.df is None:
            messagebox.showwarning("警告", "请先加载数据")
            return
        
        # 获取可选的分组变量
        cat_selection = self.cat_listbox.curselection()
        hue_col = self.cat_listbox.get(cat_selection[0]) if cat_selection else None
        
        try:
            fig = self.backend.plot_pca_2d(hue_col=hue_col)
            self._embed_figure(fig, self.plot_frame)
            self.notebook.select(0)
        except Exception as e:
            messagebox.showerror("错误", f"PCA分析失败: {str(e)}")
    
    def _draw_pca_3d(self):
        """绘制PCA 3D图"""
        if self.backend.df is None:
            messagebox.showwarning("警告", "请先加载数据")
            return
        
        # 获取可选的分组变量
        cat_selection = self.cat_listbox.curselection()
        hue_col = self.cat_listbox.get(cat_selection[0]) if cat_selection else None
        
        try:
            fig = self.backend.plot_pca_3d(hue_col=hue_col)
            self._embed_figure(fig, self.plot_frame)
            self.notebook.select(0)
        except Exception as e:
            messagebox.showerror("错误", f"PCA分析失败: {str(e)}")
    
    def _show_pca_analysis(self):
        """显示PCA分析详细结果"""
        if self.backend.df is None:
            messagebox.showwarning("警告", "请先加载数据")
            return
        
        try:
            result = self.backend.get_pca_analysis()
            self._display_text_result(result)
        except Exception as e:
            messagebox.showerror("错误", f"PCA分析失败: {str(e)}")
    
    def _draw_kmeans(self):
        """绘制K-Means聚类图"""
        if self.backend.df is None:
            messagebox.showwarning("警告", "请先加载数据")
            return
        
        # 弹出对话框选择聚类数
        from tkinter import simpledialog
        n_clusters = simpledialog.askinteger("K-Means聚类", "请输入聚类数K:", 
                                              initialvalue=3, minvalue=2, maxvalue=10)
        if n_clusters is None:
            return
        
        try:
            fig = self.backend.plot_kmeans_cluster(n_clusters=n_clusters)
            self._embed_figure(fig, self.plot_frame)
            self.notebook.select(0)
        except Exception as e:
            messagebox.showerror("错误", f"聚类分析失败: {str(e)}")
    
    def _show_cluster_analysis(self):
        """显示聚类分析详细结果"""
        if self.backend.df is None:
            messagebox.showwarning("警告", "请先加载数据")
            return
        
        from tkinter import simpledialog
        n_clusters = simpledialog.askinteger("聚类分析", "请输入聚类数K:", 
                                              initialvalue=3, minvalue=2, maxvalue=10)
        if n_clusters is None:
            return
        
        try:
            result = self.backend.get_cluster_analysis(n_clusters=n_clusters)
            self._display_text_result(result)
        except Exception as e:
            messagebox.showerror("错误", f"聚类分析失败: {str(e)}")
    
    def _draw_dendrogram(self):
        """绘制层次聚类树状图"""
        if self.backend.df is None:
            messagebox.showwarning("警告", "请先加载数据")
            return
        
        # 获取可选的标签变量（分类变量）
        cat_selection = self.cat_listbox.curselection()
        label_col = self.cat_listbox.get(cat_selection[0]) if cat_selection else None
        
        try:
            fig = self.backend.plot_dendrogram(label_col=label_col)
            self._embed_figure(fig, self.plot_frame)
            self.notebook.select(0)
            
            if label_col:
                self.status_var.set(f"树状图已生成，标签: {label_col}")
            else:
                self.status_var.set("树状图已生成（选择分类变量X可显示标签）")
        except Exception as e:
            messagebox.showerror("错误", f"绑制失败: {str(e)}")


def main():
    """主函数"""
    root = tk.Tk()
    app = App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
