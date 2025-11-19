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
        ttk.Label(cont_header, text="连续型变量 (Y轴/数值):", font=("微软雅黑", 10, "bold")).pack(side=tk.LEFT)
        ttk.Button(cont_header, text="✕", width=3, command=self._clear_cont_selection).pack(side=tk.RIGHT)
        
        cont_scrollbar = ttk.Scrollbar(left_frame)
        cont_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.cont_listbox = tk.Listbox(
            left_frame,
            height=10,
            width=25,
            yscrollcommand=cont_scrollbar.set,
            selectmode=tk.SINGLE,
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
        button_frame = ttk.LabelFrame(self.root, text="分析操作", padding=10)
        button_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # 第一行按钮
        row1 = ttk.Frame(button_frame)
        row1.pack(fill=tk.X, pady=5)
        
        ttk.Button(
            row1,
            text="📊 直方图 (选Y)",
            command=self._draw_histogram
        ).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(
            row1,
            text="📈 Q-Q图 (选Y)",
            command=self._draw_qq
        ).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(
            row1,
            text="📦 箱线图 (选X、Y)",
            command=self._draw_boxplot
        ).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(
            row1,
            text="🎻 小提琴图 (选X、Y)",
            command=self._draw_violin
        ).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(
            row1,
            text="📉 折线图 (选Y)",
            command=self._draw_line
        ).pack(side=tk.LEFT, padx=5)
        
        # 第二行按钮
        row2 = ttk.Frame(button_frame)
        row2.pack(fill=tk.X, pady=5)
        
        ttk.Button(
            row2,
            text="📊 柱状图 (选X或Y)",
            command=self._draw_bar
        ).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(
            row2,
            text="🥧 饼图 (选X)",
            command=self._draw_pie
        ).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(
            row2,
            text="📋 描述统计 (选Y)",
            command=self._show_descriptive_stats
        ).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(
            row2,
            text="🔬 t检验 (选X、Y)",
            command=self._run_t_test
        ).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(
            row2,
            text="💾 保存图形",
            command=self._save_plot
        ).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(
            row2,
            text="❌ 清空结果",
            command=self._clear_output
        ).pack(side=tk.LEFT, padx=5)
    
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


def main():
    """主函数"""
    root = tk.Tk()
    app = App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
