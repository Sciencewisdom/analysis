"""
分析后端模块 - 不涉及任何GUI代码
包含数据加载、分析和绘图的所有逻辑
"""

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from scipy import stats
from scipy.cluster.hierarchy import dendrogram, linkage
from sklearn.preprocessing import StandardScaler
from sklearn.decomposition import PCA
from sklearn.cluster import KMeans
import warnings
import matplotlib.font_manager as fm
from mpl_toolkits.mplot3d import Axes3D
import plotly.graph_objects as go
import plotly.io as pio
import webbrowser
import os

warnings.filterwarnings('ignore')

# 设置中文字体 - 优先使用系统中文字体
def setup_chinese_fonts():
    """设置中文字体，自动检测系统可用字体"""
    # Windows系统常见的中文字体
    font_candidates = [
        'SimHei',           # 黑体
        'Microsoft YaHei',  # 微软雅黑
        'KaiTi',            # 楷体
        'FangSong',         # 仿宋
        'STHeiti',          # 华文黑体
        'DejaVu Sans'       # 备用英文字体
    ]
    
    available_fonts = [f.name for f in fm.fontManager.ttflist]
    
    for font in font_candidates:
        if font in available_fonts:
            plt.rcParams['font.sans-serif'] = [font]
            plt.rcParams['axes.unicode_minus'] = False
            return font
    
    # 如果都没有找到，使用默认
    plt.rcParams['axes.unicode_minus'] = False
    return None

setup_chinese_fonts()
sns.set_style("whitegrid")


class DataAnalyzer:
    """
    数据分析类 - 核心分析逻辑，独立于GUI
    """
    
    def __init__(self):
        """初始化分析器"""
        self.df = None
        self.filepath = None
        self.continuous_cols = []
        self.categorical_cols = []
    
    
    def load_data(self, filepath):
        """
        加载CSV数据文件
        
        参数:
            filepath (str): CSV文件路径
        
        返回:
            dict: 包含 'continuous' 和 'categorical' 列名列表的字典
                  {
                    'continuous': ['Age', 'BMI', 'SBP', ...],
                    'categorical': ['Sex', ...]
                  }
        
        异常:
            ValueError: 文件读取或处理失败
        """
        try:
            # 使用 gbk 编码处理中文，skiprows=[1] 跳过中文解释行
            self.df = pd.read_csv(
                filepath,
                header=0,
                skiprows=[1],
                encoding='gbk',
                na_values=['', 'nan', 'NaN']
            )
            
            self.filepath = filepath
            
            # 数据清理：移除完全空白的列和行
            self.df = self.df.dropna(how='all', axis=1)
            self.df = self.df.dropna(how='all', axis=0)
            
            # 自动检测列类型
            self._detect_column_types()
            
            return {
                'continuous': self.continuous_cols,
                'categorical': self.categorical_cols,
                'shape': self.df.shape,
                'columns': list(self.df.columns)
            }
        
        except Exception as e:
            raise ValueError(f"加载文件失败: {str(e)}")
    
    
    def _detect_column_types(self):
        """
        自动检测列的类型（连续型 vs 分类型）
        
        规则:
        - 数值型且唯一值 > 10 -> 连续型
        - 对象型或唯一值 <= 10 -> 分类型
        """
        self.continuous_cols = []
        self.categorical_cols = []
        
        for col in self.df.columns:
            # 尝试转换为数值
            if pd.api.types.is_numeric_dtype(self.df[col]):
                # 数值类型：检查唯一值数量
                if self.df[col].nunique() > 10:
                    self.continuous_cols.append(col)
                else:
                    self.categorical_cols.append(col)
            else:
                # 非数值类型视为分类
                self.categorical_cols.append(col)
    
    
    def get_descriptive_stats(self, column):
        """
        获取列的描述性统计
        
        参数:
            column (str): 列名
        
        返回:
            str: 格式化的统计信息字符串
        """
        if self.df is None:
            return "错误: 请先加载数据"
        
        if column not in self.df.columns:
            return f"错误: 列 '{column}' 不存在"
        
        try:
            stats_result = self.df[column].describe()
            
            # 格式化输出
            output = f"\n=== '{column}' 的描述性统计 ===\n"
            output += f"计数 (Count):    {stats_result['count']:.0f}\n"
            output += f"均值 (Mean):     {stats_result['mean']:.4f}\n"
            output += f"标准差 (Std):    {stats_result['std']:.4f}\n"
            output += f"最小值 (Min):    {stats_result['min']:.4f}\n"
            output += f"25% (Q1):       {stats_result['25%']:.4f}\n"
            output += f"中位数 (Median): {stats_result['50%']:.4f}\n"
            output += f"75% (Q3):       {stats_result['75%']:.4f}\n"
            output += f"最大值 (Max):    {stats_result['max']:.4f}\n"
            
            return output
        
        except Exception as e:
            return f"错误: {str(e)}"
    
    
    def plot_histogram(self, column, figsize=(8, 6)):
        """
        绘制直方图（含核密度估计）
        
        参数:
            column (str): 列名
            figsize (tuple): 图形大小
        
        返回:
            matplotlib.figure.Figure: 图形对象
        """
        if self.df is None:
            raise ValueError("请先加载数据")
        
        if column not in self.df.columns:
            raise ValueError(f"列 '{column}' 不存在")
        
        fig, ax = plt.subplots(figsize=figsize)
        
        try:
            sns.histplot(
                data=self.df,
                x=column,
                kde=True,
                stat='density',
                bins=30,
                ax=ax,
                color='skyblue',
                edgecolor='black'
            )
            ax.set_title(f'{column} 的分布', fontsize=14, fontweight='bold')
            ax.set_xlabel(column, fontsize=12)
            ax.set_ylabel('密度', fontsize=12)
            fig.tight_layout()
            return fig
        
        except Exception as e:
            raise ValueError(f"绘制直方图失败: {str(e)}")
    
    
    def plot_boxplot(self, x_col, y_col, figsize=(10, 6)):
        """
        绘制分组箱线图
        
        参数:
            x_col (str): X轴（分类变量）
            y_col (str): Y轴（连续变量）
            figsize (tuple): 图形大小
        
        返回:
            matplotlib.figure.Figure: 图形对象
        """
        if self.df is None:
            raise ValueError("请先加载数据")
        
        if x_col not in self.df.columns or y_col not in self.df.columns:
            raise ValueError(f"列不存在")
        
        fig, ax = plt.subplots(figsize=figsize)
        
        try:
            sns.boxplot(
                data=self.df,
                x=x_col,
                y=y_col,
                ax=ax,
                palette='Set2'
            )
            ax.set_title(f'{y_col} 按 {x_col} 分组', fontsize=14, fontweight='bold')
            ax.set_xlabel(x_col, fontsize=12)
            ax.set_ylabel(y_col, fontsize=12)
            fig.tight_layout()
            return fig
        
        except Exception as e:
            raise ValueError(f"绘制箱线图失败: {str(e)}")
    
    
    def plot_qq(self, column, figsize=(8, 6)):
        """
        绘制Q-Q图（用于检验正态性）
        
        参数:
            column (str): 列名
            figsize (tuple): 图形大小
        
        返回:
            matplotlib.figure.Figure: 图形对象
        """
        if self.df is None:
            raise ValueError("请先加载数据")
        
        if column not in self.df.columns:
            raise ValueError(f"列 '{column}' 不存在")
        
        fig, ax = plt.subplots(figsize=figsize)
        
        try:
            # 计算Q-Q图
            data = self.df[column].dropna()
            stats.probplot(data, dist="norm", plot=ax)
            ax.set_title(f'{column} 的Q-Q图 (正态性检验)', fontsize=14, fontweight='bold')
            fig.tight_layout()
            return fig
        
        except Exception as e:
            raise ValueError(f"绘制Q-Q图失败: {str(e)}")
    
    
    def run_t_test(self, cat_col, cont_col):
        """
        执行独立样本t检验
        
        参数:
            cat_col (str): 分类变量（需为二分类）
            cont_col (str): 连续变量
        
        返回:
            str: 格式化的检验结果字符串
        """
        if self.df is None:
            return "错误: 请先加载数据"
        
        if cat_col not in self.df.columns or cont_col not in self.df.columns:
            return "错误: 列不存在"
        
        try:
            # 获取唯一值
            unique_vals = self.df[cat_col].unique()
            unique_vals = [x for x in unique_vals if pd.notna(x)]
            
            if len(unique_vals) != 2:
                return f"错误: '{cat_col}' 必须是二分类变量，当前有 {len(unique_vals)} 个类别"
            
            # 分离两组数据
            group1_data = self.df[self.df[cat_col] == unique_vals[0]][cont_col].dropna()
            group2_data = self.df[self.df[cat_col] == unique_vals[1]][cont_col].dropna()
            
            if len(group1_data) == 0 or len(group2_data) == 0:
                return "错误: 分组后没有有效数据"
            
            # 执行t检验
            t_stat, p_value = stats.ttest_ind(group1_data, group2_data)
            
            # 格式化结果
            output = f"\n=== t 检验结果 ===\n"
            output += f"分类变量: {cat_col}\n"
            output += f"连续变量: {cont_col}\n"
            output += f"\nGroup 1 ({unique_vals[0]}):\n"
            output += f"  n = {len(group1_data)}\n"
            output += f"  Mean = {group1_data.mean():.4f}\n"
            output += f"  Std = {group1_data.std():.4f}\n"
            output += f"\nGroup 2 ({unique_vals[1]}):\n"
            output += f"  n = {len(group2_data)}\n"
            output += f"  Mean = {group2_data.mean():.4f}\n"
            output += f"  Std = {group2_data.std():.4f}\n"
            output += f"\n检验统计量:\n"
            output += f"  t = {t_stat:.4f}\n"
            output += f"  p-value = {p_value:.6f}\n"
            
            if p_value < 0.05:
                output += f"\n*** 结论: p < 0.05, 两组存在显著差异 ***\n"
            else:
                output += f"\n*** 结论: p >= 0.05, 两组无显著差异 ***\n"
            
            return output
        
        except Exception as e:
            return f"错误: {str(e)}"
    
    
    def plot_violin(self, x_col, y_col, figsize=(10, 6)):
        """
        绘制小提琴图（额外功能）
        
        参数:
            x_col (str): X轴（分类变量）
            y_col (str): Y轴（连续变量）
            figsize (tuple): 图形大小
        
        返回:
            matplotlib.figure.Figure: 图形对象
        """
        if self.df is None:
            raise ValueError("请先加载数据")
        
        fig, ax = plt.subplots(figsize=figsize)
        
        try:
            sns.violinplot(
                data=self.df,
                x=x_col,
                y=y_col,
                ax=ax,
                palette='muted'
            )
            ax.set_title(f'{y_col} 的分布 (按 {x_col} 分组)', fontsize=14, fontweight='bold')
            ax.set_xlabel(x_col, fontsize=12)
            ax.set_ylabel(y_col, fontsize=12)
            fig.tight_layout()
            return fig
        
        except Exception as e:
            raise ValueError(f"绘制小提琴图失败: {str(e)}")
    
    
    def plot_line(self, y_col, x_col=None, figsize=(10, 6)):
        """
        绘制折线图 - 用于展示趋势
        
        参数:
            y_col (str): Y轴列名（连续型变量）
            x_col (str, optional): X轴列名，如果为None则使用索引
            figsize (tuple): 图形尺寸
        
        返回:
            matplotlib.figure.Figure: 绘制的图形对象
        """
        if self.df is None:
            raise ValueError("请先加载数据")
        
        if y_col not in self.df.columns:
            raise ValueError(f"列 '{y_col}' 不存在")
        
        try:
            fig, ax = plt.subplots(figsize=figsize)
            
            if x_col and x_col in self.df.columns:
                ax.plot(self.df[x_col], self.df[y_col], marker='o', linewidth=2, markersize=4)
                ax.set_xlabel(x_col, fontsize=12)
                ax.set_title(f'{y_col} 随 {x_col} 变化趋势', fontsize=14, fontweight='bold')
            else:
                ax.plot(self.df.index, self.df[y_col], marker='o', linewidth=2, markersize=4)
                ax.set_xlabel('样本编号', fontsize=12)
                ax.set_title(f'{y_col} 趋势图', fontsize=14, fontweight='bold')
            
            ax.set_ylabel(y_col, fontsize=12)
            ax.grid(True, alpha=0.3)
            fig.tight_layout()
            return fig
        
        except Exception as e:
            raise ValueError(f"绘制折线图失败: {str(e)}")
    
    
    def plot_bar(self, column, figsize=(10, 6)):
        """
        绘制柱状图 - 用于分类变量的频数统计
        
        参数:
            column (str): 列名（分类变量或连续变量分组）
            figsize (tuple): 图形尺寸
        
        返回:
            matplotlib.figure.Figure: 绘制的图形对象
        """
        if self.df is None:
            raise ValueError("请先加载数据")
        
        if column not in self.df.columns:
            raise ValueError(f"列 '{column}' 不存在")
        
        try:
            fig, ax = plt.subplots(figsize=figsize)
            
            # 判断是连续变量还是分类变量
            if pd.api.types.is_numeric_dtype(self.df[column]) and self.df[column].nunique() > 10:
                # 连续变量：使用直方图风格
                ax.hist(self.df[column].dropna(), bins=30, color='skyblue', 
                       edgecolor='black', alpha=0.7)
                ax.set_title(f'{column} 分布直方图', fontsize=14, fontweight='bold')
                ax.set_ylabel('频数', fontsize=12)
            else:
                # 分类变量：使用柱状图
                value_counts = self.df[column].value_counts()
                bars = ax.bar(range(len(value_counts)), value_counts.values, 
                              color=sns.color_palette('Set2', len(value_counts)))
                
                # 设置x轴标签
                ax.set_xticks(range(len(value_counts)))
                ax.set_xticklabels(value_counts.index, rotation=45, ha='right')
                
                # 在柱子上标注数值
                for bar in bars:
                    height = bar.get_height()
                    ax.text(bar.get_x() + bar.get_width()/2., height,
                           f'{int(height)}',
                           ha='center', va='bottom', fontsize=10)
                
                ax.set_title(f'{column} 分布统计', fontsize=14, fontweight='bold')
                ax.set_ylabel('频数', fontsize=12)
            
            ax.set_xlabel(column, fontsize=12)
            ax.grid(axis='y', alpha=0.3)
            fig.tight_layout()
            return fig
        
        except Exception as e:
            raise ValueError(f"绘制柱状图失败: {str(e)}")
    
    
    def plot_pie(self, column, figsize=(8, 8)):
        """
        绘制饼图 - 用于展示分类变量的占比
        
        参数:
            column (str): 列名（分类变量）
            figsize (tuple): 图形尺寸
        
        返回:
            matplotlib.figure.Figure: 绘制的图形对象
        """
        if self.df is None:
            raise ValueError("请先加载数据")
        
        if column not in self.df.columns:
            raise ValueError(f"列 '{column}' 不存在")
        
        try:
            fig, ax = plt.subplots(figsize=figsize)
            
            # 计算占比
            value_counts = self.df[column].value_counts()
            
            # 绘制饼图
            colors = sns.color_palette('pastel', len(value_counts))
            wedges, texts, autotexts = ax.pie(
                value_counts.values,
                labels=value_counts.index,
                autopct='%1.1f%%',
                colors=colors,
                startangle=90,
                textprops={'fontsize': 11}
            )
            
            # 设置百分比文字样式
            for autotext in autotexts:
                autotext.set_color('white')
                autotext.set_fontweight('bold')
            
            ax.set_title(f'{column} 占比分布', fontsize=14, fontweight='bold', pad=20)
            
            # 添加图例
            ax.legend(
                wedges,
                [f'{label}: {count}' for label, count in zip(value_counts.index, value_counts.values)],
                title=column,
                loc="center left",
                bbox_to_anchor=(1, 0, 0.5, 1)
            )
            
            fig.tight_layout()
            return fig
        
        except Exception as e:
            raise ValueError(f"绘制饼图失败: {str(e)}")
    
    # ==================== 学术分析功能 ====================
    
    def correlation_matrix(self, columns=None):
        """
        计算相关系数矩阵
        
        参数:
            columns (list): 要分析的列名列表，默认为所有连续变量
        
        返回:
            pd.DataFrame: 相关系数矩阵
        """
        if self.df is None:
            raise ValueError("请先加载数据")
        
        if columns is None:
            columns = self.continuous_cols
        
        if len(columns) < 2:
            raise ValueError("至少需要2个连续变量进行相关性分析")
        
        return self.df[columns].corr(method='pearson')
    
    
    def plot_correlation_heatmap(self, columns=None, figsize=(10, 8)):
        """绘制相关性热力图"""
        if self.df is None:
            raise ValueError("请先加载数据")
        
        try:
            corr_matrix = self.correlation_matrix(columns)
            
            fig, ax = plt.subplots(figsize=figsize)
            sns.heatmap(
                corr_matrix,
                annot=True,
                fmt='.3f',
                cmap='RdBu_r',
                center=0,
                vmin=-1,
                vmax=1,
                square=True,
                ax=ax,
                annot_kws={'size': 9}
            )
            ax.set_title('变量相关性热力图 (Pearson)', fontsize=14, fontweight='bold')
            fig.tight_layout()
            return fig
        
        except Exception as e:
            raise ValueError(f"绘制相关性热力图失败: {str(e)}")
    
    
    def get_correlation_analysis(self, columns=None):
        """获取相关性分析的文本结果"""
        if self.df is None:
            return "错误: 请先加载数据"
        
        try:
            if columns is None:
                columns = self.continuous_cols
            
            corr_matrix = self.correlation_matrix(columns)
            
            output = "\n=== 相关性分析结果 (Pearson) ===\n"
            output += f"分析变量: {', '.join(columns)}\n\n"
            
            output += "【强相关变量对 (|r| > 0.5)】\n"
            strong_corr = []
            for i in range(len(columns)):
                for j in range(i+1, len(columns)):
                    r = corr_matrix.iloc[i, j]
                    if abs(r) > 0.5:
                        strong_corr.append((columns[i], columns[j], r))
            
            if strong_corr:
                for var1, var2, r in sorted(strong_corr, key=lambda x: abs(x[2]), reverse=True):
                    strength = "强正相关" if r > 0.7 else ("中等正相关" if r > 0 else ("强负相关" if r < -0.7 else "中等负相关"))
                    output += f"  {var1} ↔ {var2}: r = {r:.4f} ({strength})\n"
            else:
                output += "  无强相关变量对\n"
            
            output += "\n【完整相关矩阵】\n"
            output += corr_matrix.round(4).to_string()
            
            return output
        
        except Exception as e:
            return f"错误: {str(e)}"
    
    
    def plot_scatter(self, x_col, y_col, hue_col=None, add_regression=True, figsize=(10, 6)):
        """绘制散点图（可选回归线）"""
        if self.df is None:
            raise ValueError("请先加载数据")
        
        if x_col not in self.df.columns or y_col not in self.df.columns:
            raise ValueError("指定的列不存在")
        
        try:
            fig, ax = plt.subplots(figsize=figsize)
            
            if add_regression and hue_col is None:
                sns.regplot(data=self.df, x=x_col, y=y_col, ax=ax,
                           scatter_kws={'alpha': 0.6}, line_kws={'color': 'red'})
            else:
                sns.scatterplot(data=self.df, x=x_col, y=y_col, hue=hue_col, ax=ax, alpha=0.6)
            
            ax.set_title(f'{y_col} vs {x_col}', fontsize=14, fontweight='bold')
            ax.set_xlabel(x_col, fontsize=12)
            ax.set_ylabel(y_col, fontsize=12)
            fig.tight_layout()
            return fig
        
        except Exception as e:
            raise ValueError(f"绘制散点图失败: {str(e)}")
    
    
    def linear_regression(self, x_col, y_col):
        """执行简单线性回归分析"""
        if self.df is None:
            return "错误: 请先加载数据"
        
        try:
            valid_data = self.df[[x_col, y_col]].dropna()
            x = valid_data[x_col]
            y = valid_data[y_col]
            
            slope, intercept, r_value, p_value, std_err = stats.linregress(x, y)
            
            output = "\n=== 线性回归分析结果 ===\n"
            output += f"自变量 (X): {x_col}\n"
            output += f"因变量 (Y): {y_col}\n"
            output += f"样本量: n = {len(valid_data)}\n\n"
            
            output += "【回归方程】\n"
            sign = '+' if intercept >= 0 else ''
            output += f"  Y = {slope:.4f} × X {sign}{intercept:.4f}\n\n"
            
            output += "【回归系数】\n"
            output += f"  斜率 (Slope):     {slope:.4f} (SE = {std_err:.4f})\n"
            output += f"  截距 (Intercept): {intercept:.4f}\n\n"
            
            output += "【拟合优度】\n"
            output += f"  相关系数 r:       {r_value:.4f}\n"
            output += f"  决定系数 R²:      {r_value**2:.4f}\n"
            output += f"  p-value:          {p_value:.6f}\n\n"
            
            if p_value < 0.001:
                sig = "极显著 (p < 0.001)"
            elif p_value < 0.01:
                sig = "非常显著 (p < 0.01)"
            elif p_value < 0.05:
                sig = "显著 (p < 0.05)"
            else:
                sig = "不显著 (p ≥ 0.05)"
            
            output += f"【结论】回归关系{sig}\n"
            output += f"  X每增加1个单位，Y平均变化 {slope:.4f} 个单位\n"
            
            return output
        
        except Exception as e:
            return f"错误: {str(e)}"
    
    
    def one_way_anova(self, cat_col, cont_col):
        """执行单因素方差分析 (One-way ANOVA)"""
        if self.df is None:
            return "错误: 请先加载数据"
        
        try:
            groups = self.df.groupby(cat_col)[cont_col].apply(lambda x: x.dropna().values)
            group_names = list(groups.index)
            group_data = [g for g in groups.values if len(g) > 0]
            
            if len(group_data) < 2:
                return "错误: 至少需要2个有效分组"
            
            f_stat, p_value = stats.f_oneway(*group_data)
            
            output = "\n=== 单因素方差分析 (One-way ANOVA) ===\n"
            output += f"分组变量: {cat_col}\n"
            output += f"因变量: {cont_col}\n"
            output += f"组数: {len(group_data)}\n\n"
            
            output += "【各组描述统计】\n"
            for name, data in zip(group_names, group_data):
                output += f"  {name}: n={len(data)}, Mean={np.mean(data):.4f}, SD={np.std(data, ddof=1):.4f}\n"
            
            output += f"\n【ANOVA结果】\n"
            output += f"  F统计量:  {f_stat:.4f}\n"
            output += f"  p-value:  {p_value:.6f}\n\n"
            
            if p_value < 0.001:
                output += "*** 结论: 各组间存在极显著差异 (p < 0.001) ***\n"
            elif p_value < 0.01:
                output += "*** 结论: 各组间存在非常显著差异 (p < 0.01) ***\n"
            elif p_value < 0.05:
                output += "*** 结论: 各组间存在显著差异 (p < 0.05) ***\n"
            else:
                output += "结论: 各组间无显著差异 (p ≥ 0.05)\n"
            
            return output
        
        except Exception as e:
            return f"错误: {str(e)}"
    
    
    def chi_square_test(self, col1, col2):
        """执行卡方独立性检验"""
        if self.df is None:
            return "错误: 请先加载数据"
        
        try:
            contingency_table = pd.crosstab(self.df[col1], self.df[col2])
            chi2, p_value, dof, expected = stats.chi2_contingency(contingency_table)
            
            output = "\n=== 卡方独立性检验 (Chi-Square Test) ===\n"
            output += f"变量1: {col1}\n"
            output += f"变量2: {col2}\n\n"
            
            output += "【列联表 (观察频数)】\n"
            output += contingency_table.to_string()
            output += "\n\n"
            
            output += "【检验结果】\n"
            output += f"  卡方值 χ²:     {chi2:.4f}\n"
            output += f"  自由度 df:     {dof}\n"
            output += f"  p-value:       {p_value:.6f}\n\n"
            
            if p_value < 0.001:
                output += "*** 结论: 两变量间存在极显著关联 (p < 0.001) ***\n"
            elif p_value < 0.01:
                output += "*** 结论: 两变量间存在非常显著关联 (p < 0.01) ***\n"
            elif p_value < 0.05:
                output += "*** 结论: 两变量间存在显著关联 (p < 0.05) ***\n"
            else:
                output += "结论: 两变量间无显著关联 (p ≥ 0.05)\n"
            
            return output
        
        except Exception as e:
            return f"错误: {str(e)}"
    
    
    def normality_test(self, column):
        """执行正态性检验 (Shapiro-Wilk)"""
        if self.df is None:
            return "错误: 请先加载数据"
        
        try:
            data = self.df[column].dropna()
            
            if len(data) > 5000:
                stat, p_value = stats.kstest(data, 'norm', args=(data.mean(), data.std()))
                test_name = "Kolmogorov-Smirnov"
            else:
                stat, p_value = stats.shapiro(data)
                test_name = "Shapiro-Wilk"
            
            output = f"\n=== 正态性检验 ({test_name}) ===\n"
            output += f"变量: {column}\n"
            output += f"样本量: n = {len(data)}\n\n"
            
            output += "【描述统计】\n"
            output += f"  均值:     {data.mean():.4f}\n"
            output += f"  标准差:   {data.std():.4f}\n"
            output += f"  偏度:     {stats.skew(data):.4f}\n"
            output += f"  峰度:     {stats.kurtosis(data):.4f}\n\n"
            
            output += "【检验结果】\n"
            output += f"  统计量W:  {stat:.4f}\n"
            output += f"  p-value:  {p_value:.6f}\n\n"
            
            if p_value < 0.05:
                output += "*** 结论: 数据不服从正态分布 (p < 0.05) ***\n"
                output += "建议: 使用非参数检验方法\n"
            else:
                output += "结论: 数据可认为服从正态分布 (p ≥ 0.05)\n"
                output += "建议: 可使用参数检验方法 (t检验, ANOVA等)\n"
            
            return output
        
        except Exception as e:
            return f"错误: {str(e)}"
    
    
    def paired_t_test(self, col1, col2):
        """执行配对样本t检验"""
        if self.df is None:
            return "错误: 请先加载数据"
        
        try:
            valid_data = self.df[[col1, col2]].dropna()
            data1 = valid_data[col1]
            data2 = valid_data[col2]
            
            t_stat, p_value = stats.ttest_rel(data1, data2)
            diff = data2 - data1
            
            output = "\n=== 配对样本t检验 (Paired t-test) ===\n"
            output += f"变量1 (前测): {col1}\n"
            output += f"变量2 (后测): {col2}\n"
            output += f"配对样本量: n = {len(valid_data)}\n\n"
            
            output += "【描述统计】\n"
            output += f"  {col1}: Mean = {data1.mean():.4f}, SD = {data1.std():.4f}\n"
            output += f"  {col2}: Mean = {data2.mean():.4f}, SD = {data2.std():.4f}\n"
            output += f"  差值:   Mean = {diff.mean():.4f}, SD = {diff.std():.4f}\n\n"
            
            output += "【检验结果】\n"
            output += f"  t统计量: {t_stat:.4f}\n"
            output += f"  p-value: {p_value:.6f}\n\n"
            
            if p_value < 0.05:
                direction = "增加" if diff.mean() > 0 else "减少"
                output += f"*** 结论: 前后存在显著差异 (p < 0.05) ***\n"
                output += f"  平均{direction}: {abs(diff.mean()):.4f}\n"
            else:
                output += "结论: 前后无显著差异 (p ≥ 0.05)\n"
            
            return output
        
        except Exception as e:
            return f"错误: {str(e)}"
    
    
    def mann_whitney_test(self, cat_col, cont_col):
        """执行Mann-Whitney U检验 (非参数版t检验)"""
        if self.df is None:
            return "错误: 请先加载数据"
        
        try:
            unique_vals = self.df[cat_col].dropna().unique()
            if len(unique_vals) != 2:
                return f"错误: '{cat_col}' 必须是二分类变量"
            
            group1 = self.df[self.df[cat_col] == unique_vals[0]][cont_col].dropna()
            group2 = self.df[self.df[cat_col] == unique_vals[1]][cont_col].dropna()
            
            u_stat, p_value = stats.mannwhitneyu(group1, group2, alternative='two-sided')
            
            output = "\n=== Mann-Whitney U 检验 (非参数) ===\n"
            output += f"分组变量: {cat_col}\n"
            output += f"检验变量: {cont_col}\n\n"
            
            output += "【各组描述统计】\n"
            output += f"  {unique_vals[0]}: n={len(group1)}, 中位数={group1.median():.4f}\n"
            output += f"  {unique_vals[1]}: n={len(group2)}, 中位数={group2.median():.4f}\n\n"
            
            output += "【检验结果】\n"
            output += f"  U统计量: {u_stat:.4f}\n"
            output += f"  p-value: {p_value:.6f}\n\n"
            
            if p_value < 0.05:
                output += "*** 结论: 两组存在显著差异 (p < 0.05) ***\n"
            else:
                output += "结论: 两组无显著差异 (p ≥ 0.05)\n"
            
            return output
        
        except Exception as e:
            return f"错误: {str(e)}"
    
    
    def kruskal_wallis_test(self, cat_col, cont_col):
        """执行Kruskal-Wallis H检验 (非参数版ANOVA)"""
        if self.df is None:
            return "错误: 请先加载数据"
        
        try:
            groups = self.df.groupby(cat_col)[cont_col].apply(lambda x: x.dropna().values)
            group_names = list(groups.index)
            group_data = [g for g in groups.values if len(g) > 0]
            
            if len(group_data) < 2:
                return "错误: 至少需要2个有效分组"
            
            h_stat, p_value = stats.kruskal(*group_data)
            
            output = "\n=== Kruskal-Wallis H 检验 (非参数ANOVA) ===\n"
            output += f"分组变量: {cat_col}\n"
            output += f"检验变量: {cont_col}\n"
            output += f"组数: {len(group_data)}\n\n"
            
            output += "【各组描述统计】\n"
            for name, data in zip(group_names, group_data):
                output += f"  {name}: n={len(data)}, 中位数={np.median(data):.4f}\n"
            
            output += f"\n【检验结果】\n"
            output += f"  H统计量: {h_stat:.4f}\n"
            output += f"  p-value: {p_value:.6f}\n\n"
            
            if p_value < 0.05:
                output += "*** 结论: 各组间存在显著差异 (p < 0.05) ***\n"
            else:
                output += "结论: 各组间无显著差异 (p ≥ 0.05)\n"
            
            return output
        
        except Exception as e:
            return f"错误: {str(e)}"
    
    
    def get_all_descriptive_stats(self, columns=None):
        """获取多个变量的描述统计（批量）"""
        if self.df is None:
            return "错误: 请先加载数据"
        
        try:
            if columns is None:
                columns = self.continuous_cols
            
            stats_df = self.df[columns].describe().T
            stats_df['missing'] = self.df[columns].isnull().sum()
            stats_df['missing%'] = (stats_df['missing'] / len(self.df) * 100).round(2)
            
            output = "\n=== 批量描述统计 ===\n"
            output += f"样本量: {len(self.df)}\n"
            output += f"变量数: {len(columns)}\n\n"
            output += stats_df.to_string()
            
            return output
        
        except Exception as e:
            return f"错误: {str(e)}"
    
    
    def missing_value_analysis(self):
        """缺失值分析"""
        if self.df is None:
            return "错误: 请先加载数据"
        
        try:
            missing = self.df.isnull().sum()
            missing_pct = (missing / len(self.df) * 100).round(2)
            
            missing_df = pd.DataFrame({
                '缺失数': missing,
                '缺失比例(%)': missing_pct
            })
            missing_df = missing_df[missing_df['缺失数'] > 0].sort_values('缺失数', ascending=False)
            
            output = "\n=== 缺失值分析 ===\n"
            output += f"总样本量: {len(self.df)}\n"
            output += f"总变量数: {len(self.df.columns)}\n\n"
            
            if len(missing_df) == 0:
                output += "✓ 数据完整，无缺失值\n"
            else:
                output += f"存在缺失的变量数: {len(missing_df)}\n\n"
                output += missing_df.to_string()
            
            return output
        
        except Exception as e:
            return f"错误: {str(e)}"
    
    
    def export_statistics_to_excel(self, filepath, columns=None):
        """导出统计结果到Excel"""
        if self.df is None:
            return "错误: 请先加载数据"
        
        try:
            if columns is None:
                columns = self.continuous_cols
            
            stats_df = self.df[columns].describe().T
            stats_df['missing'] = self.df[columns].isnull().sum()
            stats_df['missing%'] = (stats_df['missing'] / len(self.df) * 100).round(2)
            
            corr_df = self.df[columns].corr()
            
            with pd.ExcelWriter(filepath, engine='openpyxl') as writer:
                stats_df.to_excel(writer, sheet_name='描述统计')
                corr_df.to_excel(writer, sheet_name='相关矩阵')
                self.df[columns].to_excel(writer, sheet_name='原始数据', index=False)
            
            return f"✓ 已成功导出到: {filepath}"
        
        except ImportError:
            try:
                stats_df = self.df[columns].describe().T
                csv_path = filepath.replace('.xlsx', '.csv')
                stats_df.to_csv(csv_path, encoding='utf-8-sig')
                return f"✓ 已导出为CSV: {csv_path}\n(安装openpyxl可导出Excel格式)"
            except Exception as e:
                return f"错误: {str(e)}"
        
        except Exception as e:
            return f"错误: {str(e)}"

    # ==================== 高级分析功能 ====================
    
    def plot_3d_scatter(self, x_col, y_col, z_col, hue_col=None, figsize=(10, 8)):
        """绘制3D散点图"""
        if self.df is None:
            raise ValueError("请先加载数据")
        
        fig = plt.figure(figsize=figsize)
        ax = fig.add_subplot(111, projection='3d')
        
        x = self.df[x_col].dropna()
        y = self.df[y_col].dropna()
        z = self.df[z_col].dropna()
        
        # 对齐索引
        common_idx = x.index.intersection(y.index).intersection(z.index)
        x, y, z = x[common_idx], y[common_idx], z[common_idx]
        
        if hue_col and hue_col in self.df.columns:
            hue = self.df.loc[common_idx, hue_col]
            categories = hue.unique()
            colors = plt.cm.Set1(np.linspace(0, 1, len(categories)))
            
            for cat, color in zip(categories, colors):
                mask = hue == cat
                ax.scatter(x[mask], y[mask], z[mask], c=[color], label=str(cat), alpha=0.7, s=50)
            ax.legend(title=hue_col)
        else:
            scatter = ax.scatter(x, y, z, c=z, cmap='viridis', alpha=0.7, s=50)
            plt.colorbar(scatter, ax=ax, label=z_col, shrink=0.6)
        
        ax.set_xlabel(x_col)
        ax.set_ylabel(y_col)
        ax.set_zlabel(z_col)
        ax.set_title(f'3D散点图: {x_col} vs {y_col} vs {z_col}')
        
        plt.tight_layout()
        return fig
    
    def plot_pca_2d(self, columns=None, hue_col=None, figsize=(10, 8)):
        """PCA降维可视化（2D）"""
        if self.df is None:
            raise ValueError("请先加载数据")
        
        if columns is None:
            columns = self.continuous_cols
        
        if len(columns) < 2:
            raise ValueError("至少需要2个连续变量进行PCA")
        
        # 准备数据
        data = self.df[columns].dropna()
        scaler = StandardScaler()
        scaled_data = scaler.fit_transform(data)
        
        # PCA
        pca = PCA(n_components=2)
        pca_result = pca.fit_transform(scaled_data)
        
        fig, ax = plt.subplots(figsize=figsize)
        
        if hue_col and hue_col in self.df.columns:
            hue = self.df.loc[data.index, hue_col]
            for cat in hue.unique():
                mask = hue == cat
                ax.scatter(pca_result[mask, 0], pca_result[mask, 1], 
                          label=str(cat), alpha=0.7, s=60)
            ax.legend(title=hue_col)
        else:
            ax.scatter(pca_result[:, 0], pca_result[:, 1], alpha=0.7, s=60, c='steelblue')
        
        ax.set_xlabel(f'PC1 ({pca.explained_variance_ratio_[0]*100:.1f}%)')
        ax.set_ylabel(f'PC2 ({pca.explained_variance_ratio_[1]*100:.1f}%)')
        ax.set_title(f'PCA降维可视化\n总解释方差: {sum(pca.explained_variance_ratio_)*100:.1f}%')
        ax.axhline(y=0, color='gray', linestyle='--', alpha=0.3)
        ax.axvline(x=0, color='gray', linestyle='--', alpha=0.3)
        
        plt.tight_layout()
        return fig
    
    def plot_pca_3d(self, columns=None, hue_col=None, figsize=(10, 8)):
        """PCA降维可视化（3D）"""
        if self.df is None:
            raise ValueError("请先加载数据")
        
        if columns is None:
            columns = self.continuous_cols
        
        if len(columns) < 3:
            raise ValueError("至少需要3个连续变量进行3D PCA")
        
        # 准备数据
        data = self.df[columns].dropna()
        scaler = StandardScaler()
        scaled_data = scaler.fit_transform(data)
        
        # PCA
        pca = PCA(n_components=3)
        pca_result = pca.fit_transform(scaled_data)
        
        fig = plt.figure(figsize=figsize)
        ax = fig.add_subplot(111, projection='3d')
        
        if hue_col and hue_col in self.df.columns:
            hue = self.df.loc[data.index, hue_col]
            categories = hue.unique()
            colors = plt.cm.Set1(np.linspace(0, 1, len(categories)))
            
            for cat, color in zip(categories, colors):
                mask = hue == cat
                ax.scatter(pca_result[mask, 0], pca_result[mask, 1], pca_result[mask, 2],
                          c=[color], label=str(cat), alpha=0.7, s=50)
            ax.legend(title=hue_col)
        else:
            ax.scatter(pca_result[:, 0], pca_result[:, 1], pca_result[:, 2],
                      c=pca_result[:, 2], cmap='viridis', alpha=0.7, s=50)
        
        ax.set_xlabel(f'PC1 ({pca.explained_variance_ratio_[0]*100:.1f}%)')
        ax.set_ylabel(f'PC2 ({pca.explained_variance_ratio_[1]*100:.1f}%)')
        ax.set_zlabel(f'PC3 ({pca.explained_variance_ratio_[2]*100:.1f}%)')
        ax.set_title(f'PCA 3D降维可视化\n总解释方差: {sum(pca.explained_variance_ratio_)*100:.1f}%')
        
        plt.tight_layout()
        return fig
    
    def get_pca_analysis(self, columns=None):
        """获取PCA分析详细结果"""
        if self.df is None:
            return "错误: 请先加载数据"
        
        if columns is None:
            columns = self.continuous_cols
        
        data = self.df[columns].dropna()
        scaler = StandardScaler()
        scaled_data = scaler.fit_transform(data)
        
        pca = PCA()
        pca.fit(scaled_data)
        
        result = "=" * 50 + "\n"
        result += "          PCA主成分分析结果\n"
        result += "=" * 50 + "\n\n"
        
        result += "【方差解释比例】\n"
        cumsum = 0
        for i, (var, ratio) in enumerate(zip(pca.explained_variance_, pca.explained_variance_ratio_)):
            cumsum += ratio
            result += f"  PC{i+1}: {ratio*100:6.2f}% (累计: {cumsum*100:6.2f}%)\n"
        
        result += f"\n【主成分载荷矩阵】\n"
        loadings = pd.DataFrame(
            pca.components_.T,
            columns=[f'PC{i+1}' for i in range(len(columns))],
            index=columns
        )
        result += loadings.round(3).to_string()
        
        result += "\n\n【建议】\n"
        cumsum = 0
        for i, ratio in enumerate(pca.explained_variance_ratio_):
            cumsum += ratio
            if cumsum >= 0.8:
                result += f"  前{i+1}个主成分可解释80%以上的方差\n"
                break
        
        return result
    
    def plot_kmeans_cluster(self, columns=None, n_clusters=3, figsize=(12, 5)):
        """K-Means聚类分析可视化"""
        if self.df is None:
            raise ValueError("请先加载数据")
        
        if columns is None:
            columns = self.continuous_cols
        
        if len(columns) < 2:
            raise ValueError("至少需要2个连续变量进行聚类")
        
        # 准备数据
        data = self.df[columns].dropna()
        scaler = StandardScaler()
        scaled_data = scaler.fit_transform(data)
        
        # K-Means聚类
        kmeans = KMeans(n_clusters=n_clusters, random_state=42, n_init=10)
        clusters = kmeans.fit_predict(scaled_data)
        
        # PCA降维用于可视化
        pca = PCA(n_components=2)
        pca_result = pca.fit_transform(scaled_data)
        
        fig, axes = plt.subplots(1, 2, figsize=figsize)
        
        # 左图：聚类结果
        scatter = axes[0].scatter(pca_result[:, 0], pca_result[:, 1], 
                                  c=clusters, cmap='Set1', alpha=0.7, s=60)
        
        # 绘制聚类中心
        centers_pca = pca.transform(kmeans.cluster_centers_)
        axes[0].scatter(centers_pca[:, 0], centers_pca[:, 1], 
                       c='black', marker='X', s=200, edgecolors='white', linewidths=2)
        
        axes[0].set_xlabel(f'PC1 ({pca.explained_variance_ratio_[0]*100:.1f}%)')
        axes[0].set_ylabel(f'PC2 ({pca.explained_variance_ratio_[1]*100:.1f}%)')
        axes[0].set_title(f'K-Means聚类结果 (K={n_clusters})')
        axes[0].legend(*scatter.legend_elements(), title="簇")
        
        # 右图：肘部法则
        inertias = []
        K_range = range(1, min(10, len(data)))
        for k in K_range:
            km = KMeans(n_clusters=k, random_state=42, n_init=10)
            km.fit(scaled_data)
            inertias.append(km.inertia_)
        
        axes[1].plot(K_range, inertias, 'bo-', linewidth=2, markersize=8)
        axes[1].axvline(x=n_clusters, color='red', linestyle='--', label=f'当前K={n_clusters}')
        axes[1].set_xlabel('聚类数K')
        axes[1].set_ylabel('簇内平方和 (Inertia)')
        axes[1].set_title('肘部法则 - 选择最佳K值')
        axes[1].legend()
        axes[1].grid(True, alpha=0.3)
        
        plt.tight_layout()
        return fig
    
    def get_cluster_analysis(self, columns=None, n_clusters=3):
        """获取聚类分析详细结果"""
        if self.df is None:
            return "错误: 请先加载数据"
        
        if columns is None:
            columns = self.continuous_cols
        
        data = self.df[columns].dropna()
        scaler = StandardScaler()
        scaled_data = scaler.fit_transform(data)
        
        kmeans = KMeans(n_clusters=n_clusters, random_state=42, n_init=10)
        clusters = kmeans.fit_predict(scaled_data)
        
        result = "=" * 50 + "\n"
        result += f"       K-Means聚类分析结果 (K={n_clusters})\n"
        result += "=" * 50 + "\n\n"
        
        result += "【各簇样本数】\n"
        unique, counts = np.unique(clusters, return_counts=True)
        for c, n in zip(unique, counts):
            result += f"  簇 {c}: {n} 个样本 ({n/len(clusters)*100:.1f}%)\n"
        
        result += "\n【各簇中心点（标准化后）】\n"
        centers_df = pd.DataFrame(kmeans.cluster_centers_, columns=columns)
        result += centers_df.round(3).to_string()
        
        result += "\n\n【各簇特征均值（原始值）】\n"
        data_with_cluster = data.copy()
        data_with_cluster['Cluster'] = clusters
        cluster_means = data_with_cluster.groupby('Cluster').mean()
        result += cluster_means.round(2).to_string()
        
        result += f"\n\n【模型评估】\n"
        result += f"  簇内平方和 (Inertia): {kmeans.inertia_:.2f}\n"
        
        return result
    
    def plot_dendrogram(self, columns=None, label_col=None, method='ward', figsize=(14, 7)):
        """绘制层次聚类树状图"""
        if self.df is None:
            raise ValueError("请先加载数据")
        
        if columns is None:
            columns = self.continuous_cols
        
        data = self.df[columns].dropna()
        original_index = data.index.tolist()
        
        # 如果样本太多，随机采样
        max_samples = 50
        if len(data) > max_samples:
            data = data.sample(n=max_samples, random_state=42)
            original_index = data.index.tolist()
        
        scaler = StandardScaler()
        scaled_data = scaler.fit_transform(data)
        
        # 准备标签
        if label_col and label_col in self.df.columns:
            labels = [f"{self.df.loc[idx, label_col]}_{i}" for i, idx in enumerate(original_index)]
        else:
            labels = [f"样本{i}" for i in range(len(data))]
        
        fig, ax = plt.subplots(figsize=figsize)
        
        linked = linkage(scaled_data, method=method)
        
        # 根据样本数决定是否显示标签
        show_labels = len(data) <= 50
        
        dn = dendrogram(
            linked, 
            ax=ax, 
            labels=labels if show_labels else None,
            leaf_rotation=45, 
            leaf_font_size=8,
            color_threshold=0.7 * max(linked[:, 2])  # 设置颜色阈值，更好区分聚类
        )
        
        ax.set_title(f'层次聚类树状图 (方法: {method}, 样本数: {len(data)})', fontsize=12)
        ax.set_xlabel('样本' if show_labels else '样本索引', fontsize=10)
        ax.set_ylabel('聚类距离', fontsize=10)
        
        # 添加图例说明颜色含义
        ax.axhline(y=0.7 * max(linked[:, 2]), color='gray', linestyle='--', alpha=0.5, label='聚类阈值')
        ax.legend(loc='upper right')
        
        plt.tight_layout()
        return fig
    
    def plot_pair_grid(self, columns=None, hue_col=None, figsize=(10, 10)):
        """绘制配对图矩阵"""
        if self.df is None:
            raise ValueError("请先加载数据")
        
        if columns is None:
            columns = self.continuous_cols[:4]  # 默认取前4个
        
        if len(columns) < 2:
            raise ValueError("至少需要2个变量进行配对分析")
        
        # 根据变量数量动态调整子图大小
        n_vars = len(columns)
        height = max(1.2, 6 / n_vars)  # 动态计算高度，总高度约6英寸
        
        plot_data = self.df[columns].dropna()
        
        if hue_col and hue_col in self.df.columns:
            plot_data = plot_data.copy()
            plot_data[hue_col] = self.df.loc[plot_data.index, hue_col]
            g = sns.pairplot(plot_data, hue=hue_col, diag_kind='kde', 
                            plot_kws={'alpha': 0.6, 's': 20}, height=height)
        else:
            g = sns.pairplot(plot_data, diag_kind='kde', 
                            plot_kws={'alpha': 0.6, 's': 20}, height=height)
        
        g.fig.suptitle(f'配对图矩阵 ({len(columns)}个变量)', y=1.01, fontsize=10)
        g.fig.tight_layout()
        
        return g.fig
    
    def plot_distribution_comparison(self, column, group_col, figsize=(12, 5)):
        """绘制分布对比图（直方图+密度+箱线图）"""
        if self.df is None:
            raise ValueError("请先加载数据")
        
        fig, axes = plt.subplots(1, 3, figsize=figsize)
        
        groups = self.df[group_col].unique()
        colors = plt.cm.Set1(np.linspace(0, 1, len(groups)))
        
        # 直方图对比
        for group, color in zip(groups, colors):
            data = self.df[self.df[group_col] == group][column].dropna()
            axes[0].hist(data, bins=15, alpha=0.5, label=str(group), color=color, density=True)
        axes[0].set_xlabel(column)
        axes[0].set_ylabel('密度')
        axes[0].set_title('直方图对比')
        axes[0].legend(title=group_col)
        
        # 核密度估计对比
        for group, color in zip(groups, colors):
            data = self.df[self.df[group_col] == group][column].dropna()
            if len(data) > 1:
                data.plot.kde(ax=axes[1], label=str(group), color=color, linewidth=2)
        axes[1].set_xlabel(column)
        axes[1].set_ylabel('密度')
        axes[1].set_title('核密度估计对比')
        axes[1].legend(title=group_col)
        
        # 箱线图对比
        box_data = [self.df[self.df[group_col] == g][column].dropna() for g in groups]
        bp = axes[2].boxplot(box_data, labels=[str(g) for g in groups], patch_artist=True)
        for patch, color in zip(bp['boxes'], colors):
            patch.set_facecolor(color)
            patch.set_alpha(0.6)
        axes[2].set_xlabel(group_col)
        axes[2].set_ylabel(column)
        axes[2].set_title('箱线图对比')
        
        fig.suptitle(f'{column} 按 {group_col} 分组的分布对比', fontsize=12, y=1.02)
        plt.tight_layout()
        return fig
    
    def plot_radar_chart(self, columns=None, group_col=None, figsize=(8, 8)):
        """绘制雷达图（多变量对比）"""
        if self.df is None:
            raise ValueError("请先加载数据")
        
        if columns is None:
            columns = self.continuous_cols[:6]  # 最多6个变量
        
        if len(columns) > 8:
            columns = columns[:8]
        
        fig, ax = plt.subplots(figsize=figsize, subplot_kw=dict(projection='polar'))
        
        # 计算角度
        angles = np.linspace(0, 2 * np.pi, len(columns), endpoint=False).tolist()
        angles += angles[:1]  # 闭合
        
        if group_col and group_col in self.df.columns:
            groups = self.df[group_col].unique()
            colors = plt.cm.Set1(np.linspace(0, 1, len(groups)))
            
            for group, color in zip(groups, colors):
                group_data = self.df[self.df[group_col] == group][columns].mean()
                # 归一化到0-1
                normalized = (group_data - self.df[columns].min()) / (self.df[columns].max() - self.df[columns].min())
                values = normalized.tolist()
                values += values[:1]  # 闭合
                
                ax.plot(angles, values, 'o-', linewidth=2, label=str(group), color=color)
                ax.fill(angles, values, alpha=0.25, color=color)
        else:
            mean_data = self.df[columns].mean()
            normalized = (mean_data - self.df[columns].min()) / (self.df[columns].max() - self.df[columns].min())
            values = normalized.tolist()
            values += values[:1]
            
            ax.plot(angles, values, 'o-', linewidth=2, color='steelblue')
            ax.fill(angles, values, alpha=0.25, color='steelblue')
        
        ax.set_xticks(angles[:-1])
        ax.set_xticklabels(columns)
        ax.set_title('雷达图 (归一化后的均值)', y=1.08)
        
        if group_col:
            ax.legend(loc='upper right', bbox_to_anchor=(1.3, 1.0), title=group_col)
        
        plt.tight_layout()
        return fig
    
    def plot_3d_surface(self, x_col, y_col, z_col, figsize=(8, 6)):
        """绘制3D曲面图（基于数据插值）- 优化版"""
        if self.df is None:
            raise ValueError("请先加载数据")
        
        from scipy.interpolate import griddata
        
        # 获取数据并对齐索引
        data = self.df[[x_col, y_col, z_col]].dropna()
        x = data[x_col].values
        y = data[y_col].values
        z = data[z_col].values
        
        # 减少网格密度以提高性能（从50降到25）
        grid_size = 25
        xi = np.linspace(x.min(), x.max(), grid_size)
        yi = np.linspace(y.min(), y.max(), grid_size)
        xi, yi = np.meshgrid(xi, yi)
        
        # 使用线性插值（比cubic快很多）
        zi = griddata((x, y), z, (xi, yi), method='linear')
        
        fig = plt.figure(figsize=figsize)
        ax = fig.add_subplot(111, projection='3d')
        
        # 减少渲染复杂度
        surf = ax.plot_surface(xi, yi, zi, cmap='viridis', alpha=0.8, 
                               linewidth=0, antialiased=False,  # 关闭抗锯齿
                               rcount=25, ccount=25)  # 限制渲染分辨率
        
        # 只显示部分数据点（如果太多会卡）
        if len(x) > 50:
            idx = np.random.choice(len(x), 50, replace=False)
            ax.scatter(x[idx], y[idx], z[idx], c='red', s=15, alpha=0.6, label=f'数据点(采样50/{len(x)})')
        else:
            ax.scatter(x, y, z, c='red', s=15, alpha=0.6, label='原始数据点')
        
        ax.set_xlabel(x_col, fontsize=9)
        ax.set_ylabel(y_col, fontsize=9)
        ax.set_zlabel(z_col, fontsize=9)
        ax.set_title(f'3D曲面图: {z_col} = f({x_col}, {y_col})', fontsize=10)
        ax.legend(loc='upper left', fontsize=8)
        
        fig.colorbar(surf, shrink=0.5, aspect=10, label=z_col)
        return fig

    def plot_3d_surface_plotly(self, x_col, y_col, z_col):
        """绘制3D曲面图（Plotly GPU加速版）"""
        if self.df is None:
            raise ValueError("请先加载数据")
        
        from scipy.interpolate import griddata
        
        data = self.df[[x_col, y_col, z_col]].dropna()
        x = data[x_col].values
        y = data[y_col].values
        z = data[z_col].values
        
        # GPU渲染可以支持更高分辨率
        grid_size = 100
        xi = np.linspace(x.min(), x.max(), grid_size)
        yi = np.linspace(y.min(), y.max(), grid_size)
        xi, yi = np.meshgrid(xi, yi)
        
        # GPU可以使用cubic插值获得更平滑效果
        zi = griddata((x, y), z, (xi, yi), method='cubic')
        
        fig = go.Figure(data=[go.Surface(z=zi, x=xi, y=yi, colorscale='Viridis')])
        
        fig.add_trace(go.Scatter3d(x=x, y=y, z=z, mode='markers', 
                                   marker=dict(size=3, color='red'),
                                   name='原始数据点'))
        
        fig.update_layout(title=f'3D曲面图 (GPU加速): {z_col} = f({x_col}, {y_col})',
                          scene=dict(xaxis_title=x_col, yaxis_title=y_col, zaxis_title=z_col),
                          autosize=True)
        
        filename = f'3d_surface_{x_col}_{y_col}_{z_col}.html'
        pio.write_html(fig, file=filename, auto_open=True)
        return filename

    def plot_3d_scatter_plotly(self, x_col, y_col, z_col, group_col=None):
        """绘制3D散点图（Plotly GPU加速版）"""
        if self.df is None:
            raise ValueError("请先加载数据")
            
        data = self.df.dropna(subset=[x_col, y_col, z_col])
        
        if group_col:
            fig = go.Figure()
            groups = data[group_col].unique()
            for group in groups:
                subset = data[data[group_col] == group]
                fig.add_trace(go.Scatter3d(
                    x=subset[x_col], y=subset[y_col], z=subset[z_col],
                    mode='markers',
                    marker=dict(size=5),
                    name=str(group)
                ))
        else:
            fig = go.Figure(data=[go.Scatter3d(
                x=data[x_col], y=data[y_col], z=data[z_col],
                mode='markers',
                marker=dict(size=5, color=data[z_col], colorscale='Viridis'),
                text=data.index
            )])
            
        fig.update_layout(title=f'3D散点图 (GPU加速): {x_col} vs {y_col} vs {z_col}',
                          scene=dict(xaxis_title=x_col, yaxis_title=y_col, zaxis_title=z_col),
                          autosize=True)
                          
        filename = f'3d_scatter_{x_col}_{y_col}_{z_col}.html'
        pio.write_html(fig, file=filename, auto_open=True)
        return filename
        
        plt.tight_layout()
        return fig
