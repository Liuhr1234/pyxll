# info_window.py
"""模拟信息窗口模块 - 修复多input显示版本，行列互换，修复中文列名问题，增加错误标记显示"""

import tkinter as tk
from tkinter import ttk
import numpy as np

# 特殊错误标记
ERROR_MARKER = "#ERROR!"

class InfoWindow:
    """模拟信息窗口，显示四个子表格：Input数据、Output数据、Input属性、Output属性"""
    
    def __init__(self, sim):
        self.sim = sim
        self.root = tk.Tk()
        self.root.title(f"模拟信息 - 模拟ID: {sim.sim_id}")
        self.root.geometry("1400x800")
        
        # 创建主框架
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 添加标题
        title_label = ttk.Label(main_frame, text=f"模拟信息 - {sim.name}", font=("Arial", 14, "bold"))
        title_label.pack(pady=10)
        
        # 基本信息
        info_frame = ttk.LabelFrame(main_frame, text="基本信息", padding=10)
        info_frame.pack(fill=tk.X, pady=5)
        
        # 获取持续时间
        duration = sim.get_duration() if hasattr(sim, 'get_duration') else 0
        
        ttk.Label(info_frame, text=f"模拟ID: {sim.sim_id}").grid(row=0, column=0, sticky=tk.W, padx=5)
        ttk.Label(info_frame, text=f"抽样方法: {sim.sampling_method}").grid(row=0, column=1, sticky=tk.W, padx=5)
        ttk.Label(info_frame, text=f"迭代次数: {sim.n_iterations:,}").grid(row=0, column=2, sticky=tk.W, padx=5)
        ttk.Label(info_frame, text=f"创建时间: {sim.timestamp}").grid(row=0, column=3, sticky=tk.W, padx=5)
        ttk.Label(info_frame, text=f"持续时间: {duration:.2f}秒").grid(row=0, column=4, sticky=tk.W, padx=5)
        
        # 计算各种类型的Input数量
        distribution_count = 0
        makeinput_count = 0
        nested_count = 0
        
        if hasattr(sim, 'input_attributes'):
            for attrs in sim.input_attributes.values():
                if attrs.get('is_makeinput'):
                    makeinput_count += 1
                elif attrs.get('is_nested'):
                    nested_count += 1
                else:
                    distribution_count += 1
        
        ttk.Label(info_frame, text=f"分布函数: {distribution_count}").grid(row=1, column=0, sticky=tk.W, padx=5)
        ttk.Label(info_frame, text=f"MakeInput: {makeinput_count}").grid(row=1, column=1, sticky=tk.W, padx=5)
        ttk.Label(info_frame, text=f"内嵌分布: {nested_count}").grid(row=1, column=2, sticky=tk.W, padx=5)
        ttk.Label(info_frame, text=f"Output数量: {len(sim.output_cells)}").grid(row=1, column=3, sticky=tk.W, padx=5)
        
        # 数据统计
        actual_input_count = len(sim.input_cache) if hasattr(sim, 'input_cache') else 0
        actual_output_count = len(sim.output_cache) if hasattr(sim, 'output_cache') else 0
        ttk.Label(info_frame, text=f"实际Input数据: {actual_input_count}").grid(row=1, column=4, sticky=tk.W, padx=5)
        
        # 错误统计
        error_count = 0
        if hasattr(sim, 'input_cache'):
            for data in sim.input_cache.values():
                if isinstance(data, np.ndarray):
                    error_count += np.sum(data == ERROR_MARKER)
        
        if hasattr(sim, 'output_cache'):
            for data in sim.output_cache.values():
                if isinstance(data, np.ndarray):
                    error_count += np.sum(data == ERROR_MARKER)
        
        ttk.Label(info_frame, text=f"错误标记数量: {error_count}").grid(row=1, column=5, sticky=tk.W, padx=5)
        
        # 场景信息
        scenario_info = sim.get_scenario_info()
        if scenario_info.get('is_scenario_simulation', False):
            ttk.Label(info_frame, text=f"场景索引: {scenario_info['scenario_index']+1}/{scenario_info['scenario_count']}").grid(row=0, column=5, sticky=tk.W, padx=5)
        
        # 创建Notebook选项卡
        notebook = ttk.Notebook(main_frame)
        notebook.pack(fill=tk.BOTH, expand=True, pady=10)
        
        # Input数据页面（行列互换后）
        input_data_frame = ttk.Frame(notebook)
        notebook.add(input_data_frame, text=f"Input数据 ({actual_input_count}个)")
        self.create_input_data_table_transposed(input_data_frame)
        
        # Output数据页面（行列互换后）
        output_data_frame = ttk.Frame(notebook)
        notebook.add(output_data_frame, text=f"Output数据 ({actual_output_count}个)")
        self.create_output_data_table_transposed(output_data_frame)
        
        # Input属性页面
        input_attrs_frame = ttk.Frame(notebook)
        notebook.add(input_attrs_frame, text=f"Input属性 ({actual_input_count}个)")
        self.create_input_attributes_table(input_attrs_frame)
        
        # Output属性页面
        output_attrs_frame = ttk.Frame(notebook)
        notebook.add(output_attrs_frame, text=f"Output属性({actual_output_count}个)")
        self.create_output_attributes_table(output_attrs_frame)
        
        # 添加关闭按钮
        close_button = ttk.Button(main_frame, text="关闭", command=self.root.destroy)
        close_button.pack(pady=10)
    
    def create_input_data_table_transposed(self, parent):
        """
        创建Input数据表格（行列互换版本）
        行：迭代次数
        列：Input变量
        
        修改：支持内嵌分布函数的显示
        """
        # 获取所有Input数据
        if not hasattr(self.sim, 'input_cache') or not self.sim.input_cache:
            ttk.Label(parent, text="没有Input数据").pack(pady=20)
            return
        
        # 按单元格地址分组Input数据
        cell_groups = {}
        all_input_keys = []
        
        for input_key in self.sim.input_cache.keys():
            # 解析input_key：格式为 "工作表名!单元格地址_序号" 或 "工作表名!单元格地址"（对于MakeInput）
            if '!' in input_key:
                parts = input_key.split('!')
                sheet_name = parts[0]
                cell_with_index = parts[1]
                
                # 检查是否是内嵌分布函数（包含_nested_）
                if '_nested_' in cell_with_index:
                    # 内嵌分布函数：格式为 "B17_nested_1"
                    # 提取原始单元格地址和内嵌索引
                    nested_parts = cell_with_index.split('_nested_')
                    if len(nested_parts) >= 2:
                        cell_addr = nested_parts[0]
                        nested_index = nested_parts[1]
                        index = f"nested_{nested_index}"
                        is_nested = True
                        is_makeinput_only = False
                    else:
                        cell_addr = cell_with_index
                        index = '1'
                        is_nested = False
                        is_makeinput_only = False
                # 检查是否包含序号（下划线后的数字）
                elif '_' in cell_with_index and '_nested_' not in cell_with_index:
                    # 分割单元格地址和序号
                    cell_parts = cell_with_index.rsplit('_', 1)
                    cell_addr = cell_parts[0]
                    index = cell_parts[1] if len(cell_parts) > 1 else '1'
                    is_nested = False
                    is_makeinput_only = False
                else:
                    # 没有序号，可能是MakeInput单元格的最终值
                    cell_addr = cell_with_index
                    index = '1'
                    is_nested = False
                    is_makeinput_only = True
                
                full_cell_addr = f"{sheet_name}!{cell_addr}"
            else:
                # 没有工作表名的情况
                if '_nested_' in input_key:
                    # 内嵌分布函数
                    nested_parts = input_key.split('_nested_')
                    if len(nested_parts) >= 2:
                        cell_addr = nested_parts[0]
                        nested_index = nested_parts[1]
                        index = f"nested_{nested_index}"
                        is_nested = True
                        is_makeinput_only = False
                    else:
                        cell_addr = input_key
                        index = '1'
                        is_nested = False
                        is_makeinput_only = False
                elif '_' in input_key and '_nested_' not in input_key:
                    cell_parts = input_key.rsplit('_', 1)
                    cell_addr = cell_parts[0]
                    index = cell_parts[1] if len(cell_parts) > 1 else '1'
                    is_nested = False
                    is_makeinput_only = False
                else:
                    cell_addr = input_key
                    index = '1'
                    is_nested = False
                    is_makeinput_only = True
                full_cell_addr = cell_addr
            
            # 获取属性，判断类型
            attrs = {}
            if hasattr(self.sim, 'input_attributes') and input_key in self.sim.input_attributes:
                attrs = self.sim.input_attributes[input_key]
            
            is_makeinput = attrs.get('is_makeinput', False)
            is_nested = attrs.get('is_nested', False) or is_nested  # 从属性或解析结果获取
            
            # 创建显示名称
            if is_makeinput and is_makeinput_only:
                # MakeInput单元格的最终值
                display_name = f"{full_cell_addr}_MakeInput"
                if 'name' in attrs and attrs['name']:
                    display_name = f"{attrs['name']} ({full_cell_addr}_MakeInput)"
            elif is_nested:
                # 内嵌分布函数
                # 检查是否是@函数
                is_at_function = attrs.get('is_at_function', False)
                
                if is_at_function:
                    display_name = f"{full_cell_addr}_@嵌套分布{index}"
                else:
                    display_name = f"{full_cell_addr}_嵌套分布{index}"
                
                # 添加分布类型信息
                dist_type = attrs.get('distribution_type', '')
                if dist_type:
                    display_name = f"{display_name}({dist_type})"
            else:
                # 普通分布函数
                display_name = f"{full_cell_addr}_Input{index}"
                if 'name' in attrs and attrs['name']:
                    display_name = f"{attrs['name']} ({full_cell_addr}_Input{index})"
            
            # 分组
            if full_cell_addr not in cell_groups:
                cell_groups[full_cell_addr] = {}
            
            # 使用组合键确保唯一性
            if is_nested:
                group_key = f"nested_{index}"
            elif is_makeinput and is_makeinput_only:
                group_key = "makeinput_final"
            else:
                group_key = f"dist_{index}"
            
            cell_groups[full_cell_addr][group_key] = {
                'input_key': input_key,
                'display_name': display_name,
                'is_makeinput': is_makeinput,
                'is_nested': is_nested,
                'is_makeinput_only': is_makeinput_only,
                'index': index
            }
            
            all_input_keys.append({
                'full_key': input_key,
                'display_name': display_name,
                'cell_addr': full_cell_addr,
                'index': index,
                'is_makeinput': is_makeinput,
                'is_nested': is_nested,
                'is_makeinput_only': is_makeinput_only
            })
        
        # 如果没有数据
        if not cell_groups:
            ttk.Label(parent, text="没有Input数据").pack(pady=20)
            return
        
        # 限制显示的迭代次数：前50次
        max_iterations = min(50, self.sim.n_iterations) if self.sim.n_iterations > 0 else 0
        
        # 限制显示的Input数量：前50个
        max_inputs = min(50, len(all_input_keys))
        
        # 创建Treeview - 行列互换
        columns = ["迭代编号"]
        column_display_names = ["迭代编号"]  # 存储显示名称
        
        # 添加Input列（每个Input一列），使用简单的列标识符
        for i in range(max_inputs):
            input_info = all_input_keys[i]
            col_id = f"col{i+1}"
            columns.append(col_id)
            column_display_names.append(input_info['display_name'])
        
        tree = ttk.Treeview(parent, columns=columns, show='headings', height=min(25, max_iterations))
        
        # 设置列标题和宽度
        # 第一列：迭代编号（使用固定的列标识符）
        tree.heading("#1", text="迭代编号")
        tree.column("#1", width=80)
        
        # 其他列：Input变量
        for i in range(max_inputs):
            col_id = f"col{i+1}"
            display_name = column_display_names[i+1]  # 跳过第一列
            tree.heading(col_id, text=display_name)
            tree.column(col_id, width=120)  # 固定宽度
        
        # 添加滚动条
        vsb = ttk.Scrollbar(parent, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=vsb.set)
        
        hsb = ttk.Scrollbar(parent, orient="horizontal", command=tree.xview)
        tree.configure(xscrollcommand=hsb.set)
        
        # 布局
        tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        
        # 配置网格权重
        parent.grid_rowconfigure(0, weight=1)
        parent.grid_columnconfigure(0, weight=1)
        
        # 插入数据 - 行列互换
        for iteration in range(max_iterations):
            # 准备行数据
            row_data = [f"迭代{iteration+1}"]
            
            # 获取每个Input在当前迭代的值
            for i in range(max_inputs):
                input_info = all_input_keys[i]
                input_key = input_info['full_key']
                
                data = self.sim.input_cache.get(input_key)
                if data is not None and iteration < len(data):
                    value = data[iteration]
                    if isinstance(value, str) and value == ERROR_MARKER:
                        row_data.append(ERROR_MARKER)
                    elif isinstance(value, (int, float, np.number)):
                        try:
                            if not np.isnan(value):
                                row_data.append(f"{value:.4f}")
                            else:
                                row_data.append("NaN")
                        except (TypeError, ValueError):
                            row_data.append(str(value))
                    else:
                        row_data.append(str(value))
                else:
                    row_data.append("")
            
            tree.insert("", tk.END, values=row_data)
        
        # 添加统计信息标签
        total_inputs = len(all_input_keys)
        
        # 统计不同类型Input的数量
        dist_count = sum(1 for x in all_input_keys if not x['is_makeinput'] and not x['is_nested'])
        makeinput_count = sum(1 for x in all_input_keys if x['is_makeinput'] and x['is_makeinput_only'])
        nested_count = sum(1 for x in all_input_keys if x['is_nested'])
        
        stats_text = f"显示 {max_iterations} 次迭代 × {max_inputs} 个Input变量（共 {self.sim.n_iterations} 次迭代，{total_inputs} 个Input）"
        stats_text += f" | 分布函数: {dist_count}, MakeInput: {makeinput_count}, 内嵌分布: {nested_count}"
        
        stats_label = ttk.Label(parent, text=stats_text)
        stats_label.grid(row=2, column=0, pady=5, sticky=tk.W)
    
    def create_output_data_table_transposed(self, parent):
        """
        创建Output数据表格（行列互换版本）
        行：迭代次数
        列：Output变量
        """
        # 获取所有Output数据
        if not hasattr(self.sim, 'output_cache') or not self.sim.output_cache:
            ttk.Label(parent, text="没有Output数据").pack(pady=20)
            return
        
        output_keys = list(self.sim.output_cache.keys())
        if not output_keys:
            ttk.Label(parent, text="没有Output数据").pack(pady=20)
            return
        
        # 限制显示的迭代次数：前50次
        max_iterations = min(50, self.sim.n_iterations) if self.sim.n_iterations > 0 else 0
        
        # 限制显示的Output数量：前50个
        max_outputs = min(50, len(output_keys))
        
        # 创建Treeview - 行列互换
        columns = ["迭代编号"]
        column_display_names = ["迭代编号"]  # 存储显示名称
        
        # 添加Output列（每个Output一列），使用简单的列标识符
        display_names = []  # 存储显示名称
        for i in range(max_outputs):
            output_key = output_keys[i]
            # 创建显示名称
            display_name = output_key
            
            # 获取属性中的名称（如果有）
            if hasattr(self.sim, 'output_attributes') and output_key in self.sim.output_attributes:
                attrs = self.sim.output_attributes[output_key]
                if 'name' in attrs and attrs['name']:
                    display_name = f"{attrs['name']}"
            
            # 使用简单的列标识符
            col_id = f"col{i+1}"
            columns.append(col_id)
            column_display_names.append(display_name)  # 存储用于显示的名称
            display_names.append(display_name)
        
        tree = ttk.Treeview(parent, columns=columns, show='headings', height=min(25, max_iterations))
        
        # 设置列标题和宽度
        # 第一列：迭代编号（使用固定的列标识符）
        tree.heading("#1", text="迭代编号")
        tree.column("#1", width=80)
        
        # 其他列：Output变量
        for i in range(max_outputs):
            col_id = f"col{i+1}"
            display_name = display_names[i]
            tree.heading(col_id, text=display_name)
            tree.column(col_id, width=120)  # 固定宽度
        
        # 添加垂直滚动条
        vsb = ttk.Scrollbar(parent, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=vsb.set)
        
        # 添加水平滚动条
        hsb = ttk.Scrollbar(parent, orient="horizontal", command=tree.xview)
        tree.configure(xscrollcommand=hsb.set)
        
        # 布局
        tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        
        # 配置网格权重
        parent.grid_rowconfigure(0, weight=1)
        parent.grid_columnconfigure(0, weight=1)
        
        # 插入数据 - 行列互换
        for iteration in range(max_iterations):
            # 准备行数据
            row_data = [f"迭代{iteration+1}"]
            
            # 获取每个Output在当前迭代的值
            for i in range(max_outputs):
                output_key = output_keys[i]
                data = self.sim.output_cache[output_key]
                
                if data is not None and iteration < len(data):
                    value = data[iteration]
                    if isinstance(value, str) and value == ERROR_MARKER:
                        row_data.append(ERROR_MARKER)
                    elif isinstance(value, (int, float, np.number)):
                        try:
                            if not np.isnan(value):
                                row_data.append(f"{value:.4f}")
                            else:
                                row_data.append("NaN")
                        except (TypeError, ValueError):
                            row_data.append(str(value))
                    else:
                        row_data.append(str(value))
                else:
                    row_data.append("")
            
            tree.insert("", tk.END, values=row_data)
        
        # 添加统计信息标签
        stats_label = ttk.Label(
            parent, 
            text=f"显示 {max_iterations} 次迭代 × {max_outputs} 个Output变量（共 {self.sim.n_iterations} 次迭代，{len(output_keys)} 个Output）"
        )
        stats_label.grid(row=2, column=0, pady=5, sticky=tk.W)
    
    def create_input_attributes_table(self, parent):
        """创建Input属性表格 - 移除“父MakeInput”和“分布类型”两列，确保兼容DriskIndexMC"""
        # 获取所有Input数据
        if not hasattr(self.sim, 'input_cache') or not self.sim.input_cache:
            ttk.Label(parent, text="没有Input属性数据").pack(pady=20)
            return
        
        # 按单元格地址分组Input数据
        cell_groups = {}
        for input_key in self.sim.input_cache.keys():
            # 解析input_key：格式为 "工作表名!单元格地址_序号" 或 "工作表名!单元格地址_nested_序号"
            if '!' in input_key:
                # 分割工作表名和单元格地址_序号
                parts = input_key.split('!')
                sheet_name = parts[0]
                cell_with_index = parts[1]
                
                # 检查是否是内嵌分布函数
                if '_nested_' in cell_with_index:
                    # 内嵌分布函数
                    nested_parts = cell_with_index.split('_nested_')
                    if len(nested_parts) >= 2:
                        cell_addr = nested_parts[0]
                        nested_index = nested_parts[1]
                        index = f"nested_{nested_index}"
                        is_nested = True
                    else:
                        cell_addr = cell_with_index
                        index = '1'
                        is_nested = False
                # 分割单元格地址和序号
                elif '_' in cell_with_index and '_nested_' not in cell_with_index:
                    cell_parts = cell_with_index.rsplit('_', 1)
                    cell_addr = cell_parts[0]
                    index = cell_parts[1] if len(cell_parts) > 1 else '1'
                    is_nested = False
                else:
                    cell_addr = cell_with_index
                    index = '1'
                    is_nested = False
                
                full_cell_addr = f"{sheet_name}!{cell_addr}"
            else:
                # 没有工作表名的情况
                if '_nested_' in input_key:
                    # 内嵌分布函数
                    nested_parts = input_key.split('_nested_')
                    if len(nested_parts) >= 2:
                        cell_addr = nested_parts[0]
                        nested_index = nested_parts[1]
                        index = f"nested_{nested_index}"
                        is_nested = True
                    else:
                        cell_addr = input_key
                        index = '1'
                        is_nested = False
                elif '_' in input_key and '_nested_' not in input_key:
                    cell_parts = input_key.rsplit('_', 1)
                    cell_addr = cell_parts[0]
                    index = cell_parts[1] if len(cell_parts) > 1 else '1'
                    is_nested = False
                else:
                    cell_addr = input_key
                    index = '1'
                    is_nested = False
                full_cell_addr = cell_addr
            
            # 获取属性
            attrs = {}
            if hasattr(self.sim, 'input_attributes') and input_key in self.sim.input_attributes:
                attrs = self.sim.input_attributes[input_key]
            
            # 更新is_nested标记
            is_nested = attrs.get('is_nested', False) or is_nested
            is_makeinput = attrs.get('is_makeinput', False)
            
            # 分组
            if full_cell_addr not in cell_groups:
                cell_groups[full_cell_addr] = {}
            
            # 使用组合键确保唯一性
            if is_nested:
                group_key = f"nested_{index}"
            elif is_makeinput:
                group_key = "makeinput"
            else:
                group_key = f"dist_{index}"
            
            cell_groups[full_cell_addr][group_key] = {
                'input_key': input_key,
                'index': index,
                'is_nested': is_nested,
                'is_makeinput': is_makeinput,
                'attrs': attrs
            }
        
        # 如果没有数据
        if not cell_groups:
            ttk.Label(parent, text="没有Input属性数据").pack(pady=20)
            return
        
        # 创建Treeview - 移除"父MakeInput"和"分布类型"两列
        columns = ["Input Key", "单元格地址", "Input类型", "Input序号", "名称", "单位", "类别", "其他属性"]
        tree = ttk.Treeview(parent, columns=columns, show='headings', height=min(25, len(self.sim.input_cache)))
        
        # 设置列标题和宽度
        col_widths = {
            "Input Key": 180,
            "单元格地址": 120,
            "Input类型": 100,
            "Input序号": 80,
            "名称": 100,
            "单位": 80,
            "类别": 80,
            "其他属性": 250
        }
        
        for col in columns:
            tree.heading(col, text=col)
            tree.column(col, width=col_widths.get(col, 100))
        
        # 添加垂直滚动条
        vsb = ttk.Scrollbar(parent, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=vsb.set)
        
        # 添加水平滚动条
        hsb = ttk.Scrollbar(parent, orient="horizontal", command=tree.xview)
        tree.configure(xscrollcommand=hsb.set)
        
        # 布局
        tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        
        # 配置网格权重
        parent.grid_rowconfigure(0, weight=1)
        parent.grid_columnconfigure(0, weight=1)
        
        # 插入数据 - 为每个Input创建一行
        for cell_addr, inputs in sorted(cell_groups.items()):
            # 对输入进行排序：先普通分布，再内嵌分布，最后MakeInput
            sorted_inputs = []
            
            # 普通分布函数
            dist_inputs = [(k, v) for k, v in inputs.items() if k.startswith('dist_')]
            dist_inputs.sort(key=lambda x: int(x[1]['index']) if x[1]['index'].isdigit() else 0)
            
            # 内嵌分布函数
            nested_inputs = [(k, v) for k, v in inputs.items() if k.startswith('nested_')]
            nested_inputs.sort(key=lambda x: int(x[1]['index'].split('_')[1]) if len(x[1]['index'].split('_')) > 1 else 0)
            
            # MakeInput
            makeinput_inputs = [(k, v) for k, v in inputs.items() if k == 'makeinput']
            
            sorted_inputs = dist_inputs + nested_inputs + makeinput_inputs
            
            for group_key, input_info in sorted_inputs:
                input_key = input_info['input_key']
                index = input_info['index']
                is_nested = input_info['is_nested']
                is_makeinput = input_info['is_makeinput']
                attrs = input_info['attrs']
                
                # 确定Input类型
                if is_makeinput:
                    input_type = "MakeInput"
                    if index == '1':
                        display_index = "最终值"
                    else:
                        display_index = index
                elif is_nested:
                    input_type = "内嵌分布"
                    # 提取内嵌索引
                    if '_' in index:
                        nested_idx = index.split('_')[1] if len(index.split('_')) > 1 else index
                        display_index = f"内嵌{nested_idx}"
                    else:
                        display_index = f"内嵌{index}"
                else:
                    input_type = "分布函数"
                    display_index = index
                
                # 提取主要属性
                name = attrs.get('name', '')
                units = attrs.get('units', '')
                category = attrs.get('category', '')
                
                # 其他属性（移除已显示的项）
                other_attrs = []
                for attr_key, attr_value in attrs.items():
                    if attr_key not in ['name', 'units', 'category', 'parent_makeinput', 'distribution_type', 'expression', 'is_makeinput', 'is_nested']:
                        if attr_value and str(attr_value).lower() not in ['false', '0', '0.0', '']:
                            other_attrs.append(f"{attr_key}:{attr_value}")
                
                other_attrs_str = "; ".join(other_attrs)
                
                # 插入行，注意移除了 parent_makeinput 和 distribution_type
                tree.insert("", tk.END, values=[
                    input_key, cell_addr, input_type, display_index, name, units, category, other_attrs_str
                ])
        
        # 添加统计信息标签
        total_inputs = sum(len(indices) for indices in cell_groups.values())
        
        # 统计不同类型
        dist_count = sum(1 for inputs in cell_groups.values() for k in inputs.keys() if k.startswith('dist_'))
        nested_count = sum(1 for inputs in cell_groups.values() for k in inputs.keys() if k.startswith('nested_'))
        makeinput_count = sum(1 for inputs in cell_groups.values() for k in inputs.keys() if k == 'makeinput')
        
        stats_text = f"显示 {total_inputs} 个Input属性 | 分布函数: {dist_count}, 内嵌分布: {nested_count}, MakeInput: {makeinput_count}"
        stats_label = ttk.Label(parent, text=stats_text)
        stats_label.grid(row=2, column=0, pady=5, sticky=tk.W)
    
    def create_output_attributes_table(self, parent):
        """创建Output属性表格"""
        # 获取所有Output属性
        if not hasattr(self.sim, 'output_attributes') or not self.sim.output_attributes:
            ttk.Label(parent, text="没有Output属性数据").pack(pady=20)
            return
        
        output_attrs = list(self.sim.output_attributes.items())
        if not output_attrs:
            ttk.Label(parent, text="没有Output属性数据").pack(pady=20)
            return
        
        # 创建Treeview
        columns = ["Output单元格", "名称", "类别", "位置", "单位", "是否日期", "是否离散", "收敛", "其他属性"]
        tree = ttk.Treeview(parent, columns=columns, show='headings', height=min(25, len(output_attrs)))
        
        # 设置列标题
        col_widths = {
            "Output单元格": 150,
            "名称": 120,
            "类别": 100,
            "位置": 60,
            "单位": 80,
            "是否日期": 70,
            "是否离散": 70,
            "收敛": 100,
            "其他属性": 200
        }
        
        for col in columns:
            tree.heading(col, text=col)
            tree.column(col, width=col_widths.get(col, 100))
        
        # 添加垂直滚动条
        vsb = ttk.Scrollbar(parent, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=vsb.set)
        
        # 添加水平滚动条
        hsb = ttk.Scrollbar(parent, orient="horizontal", command=tree.xview)
        tree.configure(xscrollcommand=hsb.set)
        
        # 布局
        tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        
        # 配置网格权重
        parent.grid_rowconfigure(0, weight=1)
        parent.grid_columnconfigure(0, weight=1)
        
        # 插入数据
        for key, attrs in output_attrs:
            if not attrs:
                continue
                
            # 提取主要属性
            name = attrs.get('name', '')
            category = attrs.get('category', '')
            position = str(attrs.get('position', 1))
            units = attrs.get('units', '')
            is_date = str(attrs.get('is_date', False))
            is_discrete = str(attrs.get('is_discrete', False))
            convergence = attrs.get('convergence', '')
            
            # 其他属性
            other_attrs = []
            for attr_key, attr_value in attrs.items():
                if attr_key not in ['name', 'category', 'position', 'units', 'is_date', 'is_discrete', 'convergence']:
                    if attr_value and str(attr_value).lower() not in ['false', '0', '0.0', '']:
                        other_attrs.append(f"{attr_key}:{attr_value}")
            
            other_attrs_str = "; ".join(other_attrs)
            
            tree.insert("", tk.END, values=[
                key, name, category, position, units, is_date, 
                is_discrete, convergence, other_attrs_str
            ])
        
        # 添加统计信息标签
        stats_label = ttk.Label(parent, text=f"显示 {len(output_attrs)} 个Output属性")
        stats_label.grid(row=2, column=0, pady=5, sticky=tk.W)