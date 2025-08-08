import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import numpy as np
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import os
from datetime import datetime
import re

class ExcelCompareTool:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel 文件批量比较工具")
        self.root.geometry("900x700")
        self.root.configure(bg='#f0f0f0')
        
        # 存储文件路径和数据
        self.selected_files = []
        self.file_pairs = []
        self.batch_results = []  # 存储批量比较结果
        
        self.setup_ui()
        
    def setup_ui(self):
        # 主标题
        title_frame = tk.Frame(self.root, bg='#f0f0f0')
        title_frame.pack(pady=20)
        
        title_label = tk.Label(title_frame, text="📊 Excel 文件批量比较工具", 
                              font=('Arial', 20, 'bold'), bg='#f0f0f0', fg='#2c3e50')
        title_label.pack()
        
        subtitle_label = tk.Label(title_frame, 
                                 text="自动匹配文件名-A和-B的文件进行批量比较，生成详细比较报告",
                                 font=('Arial', 12), bg='#f0f0f0', fg='#7f8c8d')
        subtitle_label.pack(pady=5)
        
        # 文件选择区域
        file_frame = tk.LabelFrame(self.root, text="批量选择Excel文件", 
                                  font=('Arial', 12, 'bold'), bg='#f0f0f0')
        file_frame.pack(pady=20, padx=30, fill='both', expand=True)
        
        # 选择文件按钮
        btn_frame = tk.Frame(file_frame, bg='#f0f0f0')
        btn_frame.pack(pady=10)
        
        select_btn = tk.Button(btn_frame, text="📁 选择多个Excel文件", 
                              font=('Arial', 12, 'bold'), bg='#3498db', fg='white',
                              command=self.select_multiple_files)
        select_btn.pack(side=tk.LEFT, padx=5)
        
        clear_btn = tk.Button(btn_frame, text="🗑️ 清空列表", 
                             font=('Arial', 12, 'bold'), bg='#e74c3c', fg='white',
                             command=self.clear_files)
        clear_btn.pack(side=tk.LEFT, padx=5)
        
        # 文件列表显示
        list_frame = tk.Frame(file_frame, bg='#f0f0f0')
        list_frame.pack(fill='both', expand=True, padx=10, pady=10)
        
        # 创建Treeview来显示文件
        columns = ('文件名', '类型', '状态', '配对文件')
        self.file_tree = ttk.Treeview(list_frame, columns=columns, show='headings', height=10)
        
        # 设置列标题
        self.file_tree.heading('文件名', text='文件名')
        self.file_tree.heading('类型', text='类型')
        self.file_tree.heading('状态', text='状态')
        self.file_tree.heading('配对文件', text='配对文件')
        
        # 设置列宽
        self.file_tree.column('文件名', width=300)
        self.file_tree.column('类型', width=80)
        self.file_tree.column('状态', width=100)
        self.file_tree.column('配对文件', width=300)
        
        # 添加滚动条
        scrollbar_y = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.file_tree.yview)
        scrollbar_x = ttk.Scrollbar(list_frame, orient=tk.HORIZONTAL, command=self.file_tree.xview)
        self.file_tree.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)
        
        self.file_tree.pack(side=tk.LEFT, fill='both', expand=True)
        scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
        scrollbar_x.pack(side=tk.BOTTOM, fill=tk.X)
        
        # 比较配置
        config_frame = tk.LabelFrame(self.root, text="比较配置", 
                                    font=('Arial', 12, 'bold'), bg='#f0f0f0')
        config_frame.pack(pady=10, padx=30, fill='x')
        
        range_frame = tk.Frame(config_frame, bg='#f0f0f0')
        range_frame.pack(pady=10)
        
        tk.Label(range_frame, text="数据比较行序:", font=('Arial', 10), bg='#f0f0f0').pack(side=tk.LEFT)
        self.range_entry = tk.Entry(range_frame, font=('Arial', 10), width=20)
        self.range_entry.pack(side=tk.LEFT, padx=10)
        
        # 添加行选择格式提示
        format_label = tk.Label(range_frame, text="(支持格式: '1-100'或'1,3,4,9'或留空比较所有行)", 
                               font=('Arial', 8), bg='#f0f0f0', fg='#7f8c8d')
        format_label.pack(side=tk.LEFT)
        
        # 开始比较按钮
        compare_btn = tk.Button(self.root, text="🔍 开始批量比较", font=('Arial', 14, 'bold'),
                               bg='#27ae60', fg='white', height=2, width=20,
                               command=self.start_batch_comparison)
        compare_btn.pack(pady=20)
        
        # 状态显示区域
        self.status_frame = tk.LabelFrame(self.root, text="操作状态", 
                                         font=('Arial', 12, 'bold'), bg='#f0f0f0')
        self.status_frame.pack(pady=10, padx=30, fill='x')
        
        self.status_text = tk.Text(self.status_frame, height=6, font=('Arial', 10))
        self.status_text.pack(fill='x', padx=10, pady=10)
        
        self.log_message("准备就绪，请选择要批量比较的Excel文件...")
        self.log_message("文件命名规则：基础名称-A.xlsx 和 基础名称-B.xlsx")
    
    def log_message(self, message):
        """在状态区域显示消息"""
        timestamp = datetime.now().strftime('%H:%M:%S')
        self.status_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.status_text.see(tk.END)
        self.root.update()
    
    def select_multiple_files(self):
        """选择多个Excel文件"""
        file_paths = filedialog.askopenfilenames(
            title="选择多个Excel文件",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        
        if file_paths:
            self.selected_files.extend(file_paths)
            # 去重
            self.selected_files = list(set(self.selected_files))
            self.update_file_list()
            self.log_message(f"已选择 {len(file_paths)} 个文件")
    
    def clear_files(self):
        """清空文件列表"""
        self.selected_files = []
        self.file_pairs = []
        self.batch_results = []
        self.update_file_list()
        self.log_message("已清空文件列表")
    
    def extract_base_name_and_type(self, filename):
        """提取文件的基础名称和类型（A或B）"""
        # 移除文件扩展名
        name_without_ext = os.path.splitext(filename)[0]
        
        # 检查是否以-A或-B结尾
        if name_without_ext.endswith('-A'):
            return name_without_ext[:-2], 'A'
        elif name_without_ext.endswith('-B'):
            return name_without_ext[:-2], 'B'
        else:
            return name_without_ext, 'Unknown'
    
    def find_file_pairs(self):
        """查找文件配对"""
        self.file_pairs = []
        file_dict = {}
        
        # 按基础名称分组
        for file_path in self.selected_files:
            filename = os.path.basename(file_path)
            base_name, file_type = self.extract_base_name_and_type(filename)
            
            if base_name not in file_dict:
                file_dict[base_name] = {}
            file_dict[base_name][file_type] = file_path
        
        # 查找配对
        for base_name, files in file_dict.items():
            if 'A' in files and 'B' in files:
                self.file_pairs.append({
                    'base_name': base_name,
                    'file_a': files['A'],
                    'file_b': files['B']
                })
        
        return file_dict
    
    def update_file_list(self):
        """更新文件列表显示"""
        # 清空现有项目
        for item in self.file_tree.get_children():
            self.file_tree.delete(item)
        
        if not self.selected_files:
            return
        
        # 查找配对
        file_dict = self.find_file_pairs()
        
        # 显示文件信息
        for file_path in self.selected_files:
            filename = os.path.basename(file_path)
            base_name, file_type = self.extract_base_name_and_type(filename)
            
            # 确定状态和配对文件
            if base_name in file_dict and 'A' in file_dict[base_name] and 'B' in file_dict[base_name]:
                status = "✅ 已配对"
                if file_type == 'A':
                    pair_file = os.path.basename(file_dict[base_name]['B'])
                elif file_type == 'B':
                    pair_file = os.path.basename(file_dict[base_name]['A'])
                else:
                    pair_file = "无配对"
            else:
                status = "❌ 未配对"
                pair_file = "无配对"
            
            self.file_tree.insert('', 'end', values=(filename, file_type, status, pair_file))
        
        # 更新日志
        paired_count = len(self.file_pairs)
        self.log_message(f"文件分析完成：共 {len(self.selected_files)} 个文件，{paired_count} 对可比较")
    
    def start_batch_comparison(self):
        """开始批量比较"""
        if not self.selected_files:
            messagebox.showerror("错误", "请先选择Excel文件")
            return
        
        if not self.file_pairs:
            messagebox.showerror("错误", "没有找到可配对的文件\n请确保文件名格式为：基础名称-A.xlsx 和 基础名称-B.xlsx")
            return
        
        try:
            self.log_message(f"开始批量比较 {len(self.file_pairs)} 对文件...")
            
            # 清空之前的结果
            self.batch_results = []
            
            # 获取保存目录（使用第一个文件的目录）
            save_dir = os.path.dirname(self.file_pairs[0]['file_a'])
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            
            # 创建批量比较结果目录
            batch_dir = os.path.join(save_dir, f"批量比较结果_{timestamp}")
            os.makedirs(batch_dir, exist_ok=True)
            
            successful_comparisons = 0
            failed_comparisons = 0
            
            # 逐对比较文件
            for i, pair in enumerate(self.file_pairs, 1):
                try:
                    self.log_message(f"正在比较第 {i}/{len(self.file_pairs)} 对：{pair['base_name']}")
                    
                    # 比较单对文件
                    result = self.compare_file_pair(pair, batch_dir)
                    
                    if result:
                        successful_comparisons += 1
                        self.log_message(f"✅ {pair['base_name']} 比较完成")
                        # 将结果添加到批量结果中
                        self.batch_results.append(result)
                    else:
                        failed_comparisons += 1
                        self.log_message(f"❌ {pair['base_name']} 比较失败")
                        
                except Exception as e:
                    failed_comparisons += 1
                    self.log_message(f"❌ {pair['base_name']} 比较出错：{str(e)}")
            
            # 生成批量比较汇总报告
            self.generate_batch_summary(batch_dir, successful_comparisons, failed_comparisons)
            
            # 显示完成消息
            result_msg = f"""批量比较完成！

📊 比较结果:
• 成功比较: {successful_comparisons} 对
• 失败比较: {failed_comparisons} 对
• 总计: {len(self.file_pairs)} 对

📁 结果保存在:
{batch_dir}

是否要打开结果目录？"""
            
            if messagebox.askyesno("批量比较完成", result_msg):
                # 打开结果目录
                if os.name == 'nt':  # Windows
                    os.startfile(batch_dir)
                elif os.name == 'posix':  # macOS/Linux
                    os.system(f'open "{batch_dir}"')
            
        except Exception as e:
            error_msg = f"批量比较失败: {str(e)}"
            self.log_message(error_msg)
            messagebox.showerror("错误", error_msg)
    
    def compare_file_pair(self, pair, save_dir):
        """比较单对文件"""
        try:
            # 读取Excel文件
            df1 = pd.read_excel(pair['file_a'])
            df2 = pd.read_excel(pair['file_b'])
            
            # 解析行范围
            range_text = self.range_entry.get().strip()
            
            # 初始化行索引列表
            rows_to_compare = []
            
            if range_text:
                if '-' in range_text:  # 处理范围格式 "1-100"
                    start, end = map(int, range_text.split('-'))
                    rows_to_compare = list(range(start, end + 1))
                elif ',' in range_text:  # 处理离散行格式 "1,3,4,9"
                    rows_to_compare = [int(x.strip()) for x in range_text.split(',')]
                else:  # 单个数字
                    rows_to_compare = [int(range_text)]
            else:  # 如果为空，比较所有行
                rows_to_compare = list(range(1, min(len(df1) + 1, len(df2) + 1)))
            
            # 转换为0基索引并过滤有效范围
            valid_indices = []
            for row in rows_to_compare:
                idx = row - 1
                if 0 <= idx < len(df1) and 0 <= idx < len(df2):
                    valid_indices.append(idx)
            
            if not valid_indices:
                return None
            
            # 提取指定行进行比较
            df1_compare = df1.iloc[valid_indices].copy()
            df2_compare = df2.iloc[valid_indices].copy()
            original_row_indices = [idx + 1 for idx in valid_indices]
            
            # 计算差异
            differences = self.calculate_differences(df1_compare, df2_compare, original_row_indices)
            
            # 计算统计信息
            min_rows = min(len(df1_compare), len(df2_compare))
            min_cols = min(len(df1_compare.columns), len(df2_compare.columns))
            total_cells = min_rows * min_cols
            diff_count = len(differences)
            similarity = ((total_cells - diff_count) / total_cells * 100) if total_cells > 0 else 100
            
            # 生成单个比较报告
            report_filename = f"{pair['base_name']}_比较报告.xlsx"
            report_path = os.path.join(save_dir, report_filename)
            
            self.generate_single_report(pair, df1_compare, df2_compare, differences, 
                                      original_row_indices, report_path)
            
            # 返回比较结果用于汇总
            return {
                'pair': pair,
                'differences': differences,
                'statistics': {
                    'total_cells': total_cells,
                    'diff_count': diff_count,
                    'similarity': similarity,
                    'compared_rows': min_rows,
                    'compared_cols': min_cols
                }
            }
            
        except Exception as e:
            self.log_message(f"比较 {pair['base_name']} 时出错: {str(e)}")
            return None
    
    def calculate_differences(self, df1, df2, original_row_indices):
        """计算两个DataFrame之间的差异"""
        differences = []
        min_rows = min(len(df1), len(df2))
        min_cols = min(len(df1.columns), len(df2.columns))
        
        for i in range(min_rows):
            original_row = original_row_indices[i]
            for j in range(min_cols):
                val1 = df1.iloc[i, j]
                val2 = df2.iloc[i, j]
                
                if not self.values_equal(val1, val2):
                    differences.append({
                        '原始行号': original_row,
                        '列号': j + 1,
                        '列名': df1.columns[j] if j < len(df1.columns) else f'列{j+1}',
                        '文件A值': self.format_export_value(val1),
                        '文件B值': self.format_export_value(val2),
                        '差异类型': self.get_difference_type(val1, val2)
                    })
        
        return differences
    
    def generate_single_report(self, pair, df1, df2, differences, original_row_indices, save_path):
        """生成单个文件对的比较报告"""
        with pd.ExcelWriter(save_path, engine='openpyxl') as writer:
            # 1. 概览报告
            min_rows = min(len(df1), len(df2))
            min_cols = min(len(df1.columns), len(df2.columns))
            total_cells = min_rows * min_cols
            diff_count = len(differences)
            similarity = ((total_cells - diff_count) / total_cells * 100) if total_cells > 0 else 100
            
            overview_data = {
                '项目': [
                    '基础文件名', '文件A', '文件B', '比较时间', '比较行范围',
                    '文件A行数', '文件A列数', '文件B行数', '文件B列数',
                    '比较的行数', '比较的列数', '不同单元格数', '相似度(%)'
                ],
                '值': [
                    pair['base_name'],
                    os.path.basename(pair['file_a']),
                    os.path.basename(pair['file_b']),
                    datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                    self.range_entry.get().strip() or '所有行',
                    len(df1), len(df1.columns), len(df2), len(df2.columns),
                    min_rows, min_cols, diff_count, f"{similarity:.2f}%"
                ]
            }
            overview_df = pd.DataFrame(overview_data)
            overview_df.to_excel(writer, sheet_name='比较概览', index=False)
            
            # 2. 文件A数据
            df1_export = df1.copy()
            df1_export.insert(0, '原始行号', original_row_indices)
            df1_export.to_excel(writer, sheet_name='文件A数据', index=False)
            
            # 3. 文件B数据
            df2_export = df2.copy()
            df2_export.insert(0, '原始行号', original_row_indices)
            df2_export.to_excel(writer, sheet_name='文件B数据', index=False)
            
            # 4. 差异详情
            if differences:
                diff_df = pd.DataFrame(differences)
                diff_df.to_excel(writer, sheet_name='差异详情', index=False)
            else:
                no_diff_df = pd.DataFrame({'说明': ['两个文件在指定范围内完全相同']})
                no_diff_df.to_excel(writer, sheet_name='差异详情', index=False)
    
    def generate_batch_summary(self, batch_dir, successful, failed):
        """生成批量比较汇总报告"""
        summary_path = os.path.join(batch_dir, "批量比较汇总.xlsx")
        
        with pd.ExcelWriter(summary_path, engine='openpyxl') as writer:
            # 1. 汇总统计
            summary_data = {
                '项目': [
                    '比较时间', '总文件对数', '成功比较', '失败比较', '成功率(%)',
                    '比较行范围', '结果目录'
                ],
                '值': [
                    datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                    len(self.file_pairs), successful, failed,
                    f"{(successful / len(self.file_pairs) * 100):.1f}%" if self.file_pairs else "0%",
                    self.range_entry.get().strip() or '所有行',
                    batch_dir
                ]
            }
            summary_df = pd.DataFrame(summary_data)
            summary_df.to_excel(writer, sheet_name='汇总统计', index=False)
            
            # 2. 文件对概览
            pairs_overview = []
            for result in self.batch_results:
                pair = result['pair']
                stats = result['statistics']
                diff_count = stats['diff_count']
                
                pairs_overview.append({
                    '基础名称': pair['base_name'],
                    '文件A': os.path.basename(pair['file_a']),
                    '文件B': os.path.basename(pair['file_b']),
                    '比较行数': stats['compared_rows'],
                    '比较列数': stats['compared_cols'],
                    '总单元格数': stats['total_cells'],
                    '差异单元格数': diff_count,
                    '相似度(%)': f"{stats['similarity']:.2f}%",
                    '状态': '有差异' if diff_count > 0 else '完全相同',
                    '报告文件': f"{pair['base_name']}_比较报告.xlsx"
                })
            
            pairs_df = pd.DataFrame(pairs_overview)
            pairs_df.to_excel(writer, sheet_name='文件对概览', index=False)
            
            # 3. 所有差异详情汇总
            all_differences = []
            for result in self.batch_results:
                pair = result['pair']
                differences = result['differences']
                
                for diff in differences:
                    all_differences.append({
                        '文件对': pair['base_name'],
                        '文件A': os.path.basename(pair['file_a']),
                        '文件B': os.path.basename(pair['file_b']),
                        '原始行号': diff['原始行号'],
                        '列号': diff['列号'],
                        '列名': diff['列名'],
                        '文件A值': diff['文件A值'],
                        '文件B值': diff['文件B值'],
                        '差异类型': diff['差异类型']
                    })
            
            if all_differences:
                all_diff_df = pd.DataFrame(all_differences)
                all_diff_df.to_excel(writer, sheet_name='所有差异详情', index=False)
            else:
                no_diff_df = pd.DataFrame({'说明': ['所有文件对在指定范围内都完全相同，没有发现任何差异。']})
                no_diff_df.to_excel(writer, sheet_name='所有差异详情', index=False)
            
            # 4. 差异统计分析
            if all_differences:
                # 按文件对统计差异数量
                diff_stats = []
                for result in self.batch_results:
                    pair = result['pair']
                    stats = result['statistics']
                    differences = result['differences']
                    
                    # 按列统计差异
                    col_diff_count = {}
                    for diff in differences:
                        col_name = diff['列名']
                        col_diff_count[col_name] = col_diff_count.get(col_name, 0) + 1
                    
                    # 找出差异最多的列
                    max_diff_col = max(col_diff_count.items(), key=lambda x: x[1]) if col_diff_count else ('无', 0)
                    
                    diff_stats.append({
                        '文件对': pair['base_name'],
                        '总差异数': len(differences),
                        '差异最多的列': max_diff_col[0],
                        '该列差异数': max_diff_col[1],
                        '差异率(%)': f"{(len(differences) / stats['total_cells'] * 100):.2f}%" if stats['total_cells'] > 0 else "0%"
                    })
                
                diff_stats_df = pd.DataFrame(diff_stats)
                diff_stats_df.to_excel(writer, sheet_name='差异统计分析', index=False)
            
            # 5. 有差异的文件列表（仅包含有差异的文件）
            files_with_diff = []
            for result in self.batch_results:
                if result['statistics']['diff_count'] > 0:
                    pair = result['pair']
                    stats = result['statistics']
                    
                    files_with_diff.append({
                        '文件对': pair['base_name'],
                        '文件A路径': pair['file_a'],
                        '文件B路径': pair['file_b'],
                        '差异单元格数': stats['diff_count'],
                        '相似度(%)': f"{stats['similarity']:.2f}%",
                        '详细报告': f"{pair['base_name']}_比较报告.xlsx"
                    })
            
            if files_with_diff:
                files_diff_df = pd.DataFrame(files_with_diff)
                files_diff_df.to_excel(writer, sheet_name='有差异的文件', index=False)
            else:
                no_diff_files_df = pd.DataFrame({'说明': ['所有文件对都完全相同，没有差异。']})
                no_diff_files_df.to_excel(writer, sheet_name='有差异的文件', index=False)
        
        self.log_message(f"详细汇总报告已生成：{summary_path}")
    
    def values_equal(self, val1, val2):
        """比较两个值是否相等，处理NaN和空值"""
        if pd.isna(val1) and pd.isna(val2):
            return True
        if pd.isna(val1) or pd.isna(val2):
            return False
        
        str1 = str(val1).strip()
        str2 = str(val2).strip()
        
        if str1 in ['', 'None', 'nan'] and str2 in ['', 'None', 'nan']:
            return True
        
        return str1 == str2
    
    def format_export_value(self, value):
        """格式化导出值"""
        if pd.isna(value):
            return "[空值]"
        elif str(value).strip() == "":
            return "[空字符串]"
        else:
            return value
    
    def get_difference_type(self, val1, val2):
        """获取差异类型"""
        if pd.isna(val1) and not pd.isna(val2):
            return "文件A为空值"
        elif not pd.isna(val1) and pd.isna(val2):
            return "文件B为空值"
        elif str(val1).strip() == "" and str(val2).strip() != "":
            return "文件A为空字符串"
        elif str(val1).strip() != "" and str(val2).strip() == "":
            return "文件B为空字符串"
        else:
            return "值不同"

def main():
    root = tk.Tk()
    app = ExcelCompareTool(root)
    root.mainloop()

if __name__ == "__main__":
    main()