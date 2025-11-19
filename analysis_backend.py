"""
分析后端模块 - 不涉及任何GUI代码
包含数据加载、分析和绘图的所有逻辑
"""

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from scipy import stats
import warnings
import matplotlib.font_manager as fm

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
