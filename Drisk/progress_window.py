# progress_window.py
"""进度条窗口模块 - @RISK风格界面，美观且功能正确"""

import tkinter as tk
from tkinter import ttk, messagebox
import time
import threading

class ProgressWindow:
    """@RISK风格进度条窗口 - 经典Windows桌面应用程序风格"""
    
    def __init__(self, title="Drisk-模拟进展（1 CPU）", n_iterations_per_scenario=1000, n_scenarios=1):
        # 先设置变量，但不立即显示窗口
        self.root = None
        self.is_initialized = False
        
        # 场景信息
        self.n_iterations_per_scenario = n_iterations_per_scenario
        self.n_scenarios = n_scenarios
        
        # 总进度信息
        self.total_iterations = n_iterations_per_scenario * n_scenarios
        self.completed_iterations = 0
        
        # 当前场景信息
        self.current_scenario_iteration = 0
        self.current_scenario_index = 0
        
        # 时间相关
        self.start_time = time.time()
        self.scenario_start_time = self.start_time
        
        # 取消相关
        self.cancelled = False
        self.cancel_event = None
        self._closed_by_window_x = False
        
        # 初始化标志，防止初始跳跃
        self._first_update = True
        
        # 延迟初始化窗口
        self._init_window(title)
    
    def _init_window(self, title):
        """延迟初始化窗口，避免初始跳跃"""
        self.root = tk.Tk()
        self.root.title(title)
        
        # 隐藏窗口直到完全初始化
        self.root.withdraw()
        
        # 设置更紧凑的窗口尺寸
        self.root.geometry("500x320")  # 高度调整为320，更窄一些
        
        # 设置窗口背景色与边框一致
        self.root.configure(bg="#D0D0D0")  # 窗口背景使用与边框相同的颜色
        
        # 设置最小尺寸
        self.root.minsize(500, 320)
        
        # 设置窗口置顶
        self.root.attributes('-topmost', True)
        
        # 移除左上角图标（设置空图标）
        try:
            # 创建一个1x1像素的透明图标
            empty_icon = tk.PhotoImage(width=1, height=1)
            self.root.iconphoto(True, empty_icon)
        except:
            pass  # 如果设置图标失败，忽略
        
        # 移除最小化和最大化按钮（只保留关闭按钮）
        self.root.attributes('-toolwindow', 1)  # 移除最小化和最大化按钮
        
        # 获取屏幕尺寸
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        
        # 计算窗口位置（左下角）
        window_width = 500
        window_height = 320
        x_position = 30  # 左侧边距
        y_position = screen_height - window_height - 120  # 底部位置
        
        # 设置窗口位置
        self.root.geometry(f"{window_width}x{window_height}+{x_position}+{y_position}")
        
        # 添加外围细边框（凹陷效果）
        border_frame = tk.Frame(
            self.root,
            bg="#F3F3F3",  # 与窗口背景一致的颜色
            relief=tk.SUNKEN,  # 凹陷效果
            bd=1,  # 很细的边框
            highlightbackground="#F3F3F3",  # 边框高亮颜色
            highlightthickness=0
        )
        border_frame.pack(fill=tk.BOTH, expand=True, padx=1, pady=1)
        
        # 主框架 - 灰色矩形区域，平坦边框，没有凹陷效果
        main_frame = tk.Frame(
            border_frame, 
            bg="#E8E8E8",  # 内容区域背景色
            relief=tk.FLAT,  # 平坦边框，没有凹陷效果
            bd=0,  # 边框宽度为0
            highlightbackground="#E8E8E8",  # 高亮边框颜色与背景一致
            highlightthickness=0  # 高亮边框厚度为0
        )
        main_frame.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)
        
        # 创建一个容器框架，确保所有内容都能显示
        container_frame = tk.Frame(main_frame, bg="#E8E8E8")
        container_frame.pack(fill=tk.BOTH, expand=True)
        
        # 进度条框架 - 紧凑的垂直间距
        progress_frame = tk.Frame(container_frame, bg="#E8E8E8")
        progress_frame.pack(fill=tk.X, padx=10, pady=(6, 10))
        
        # 使用Canvas自定义进度条（更可控）
        self.canvas = tk.Canvas(
            progress_frame,
            height=80,  # 进度条高度调整为80
            bg="#E8E8E8",
            highlightthickness=0,
            relief=tk.FLAT
        )
        self.canvas.pack(fill=tk.X)
        
        # 保存Canvas尺寸
        self.canvas_width = 480  # 调整宽度以适应窗口
        self.canvas_height = 80  # 进度条高度
        
        # 创建进度条背景（白色）
        self.progress_bg = self.canvas.create_rectangle(
            0, 0, self.canvas_width, self.canvas_height,
            fill="white", outline="#A0A0A0", width=1
        )
        
        # 创建进度条前景（深红色）
        self.progress_fg = self.canvas.create_rectangle(
            0, 0, 0, self.canvas_height,  # 初始宽度为0
            fill="#8B0000", outline="#8B0000", width=0  # 暗红色
        )
        
        # 创建文本标签（在进度条上方，透明背景）- 黑色，居中，字体大小11
        self.progress_text_id = self.canvas.create_text(
            self.canvas_width // 2, self.canvas_height // 2,
            text="模拟中 0.00%",
            font=("微软雅黑", 11, "bold"),  # 字体大小改为11
            fill="black",  # 黑色文字
            anchor="center"
        )
        
        # 信息框架 - 使用grid布局实现两列对齐，更紧凑
        info_frame = tk.Frame(container_frame, bg="#E8E8E8")
        info_frame.pack(fill=tk.X, padx=10, pady=(4, 6))
        
        # 标签样式配置
        label_font = ("微软雅黑", 10)  # 中文用微软雅黑
        value_font = ("Arial", 10, "bold")  # 英文用Arial
        
        # 第一行：迭代次数（当前场景）- 删除冒号
        tk.Label(
            info_frame,
            text="迭代次数",
            font=label_font,
            bg="#E8E8E8",  # 灰色背景
            fg="#505050",
            anchor=tk.W
        ).grid(row=0, column=0, sticky=tk.W, padx=(0, 8), pady=4)
        
        # 白色凹陷输入框样式的数值显示
        iter_frame = tk.Frame(
            info_frame,
            bg="white",
            relief=tk.SUNKEN,
            bd=2,
            highlightbackground="#A0A0A0"
        )
        iter_frame.grid(row=0, column=1, sticky=tk.EW, pady=4)
        
        self.iter_label = tk.Label(
            iter_frame,
            text=f"0/{self.n_iterations_per_scenario:,}",
            font=value_font,
            bg="white",
            fg="#000000",
            width=18,
            anchor=tk.W,
            padx=6,
            pady=2
        )
        self.iter_label.pack()
        
        # 第二行：模拟次数（场景数）- 删除冒号
        tk.Label(
            info_frame,
            text="模拟次数",
            font=label_font,
            bg="#E8E8E8",  # 灰色背景
            fg="#505050",
            anchor=tk.W
        ).grid(row=1, column=0, sticky=tk.W, padx=(0, 8), pady=4)
        
        # 白色凹陷输入框
        scenario_frame = tk.Frame(
            info_frame,
            bg="white",
            relief=tk.SUNKEN,
            bd=2,
            highlightbackground="#A0A0A0"
        )
        scenario_frame.grid(row=1, column=1, sticky=tk.EW, pady=4)
        
        self.scenario_label = tk.Label(
            scenario_frame,
            text=f"0/{self.n_scenarios}",
            font=value_font,
            bg="white",
            fg="#000000",
            width=18,
            anchor=tk.W,
            padx=6,
            pady=2
        )
        self.scenario_label.pack()
        
        # 第三行：运行时间（总时间）- 删除冒号
        tk.Label(
            info_frame,
            text="运行时间",
            font=label_font,
            bg="#E8E8E8",  # 灰色背景
            fg="#505050",
            anchor=tk.W
        ).grid(row=2, column=0, sticky=tk.W, padx=(0, 8), pady=4)
        
        # 白色凹陷输入框
        time_frame = tk.Frame(
            info_frame,
            bg="white",
            relief=tk.SUNKEN,
            bd=2,
            highlightbackground="#A0A0A0"
        )
        time_frame.grid(row=2, column=1, sticky=tk.EW, pady=4)
        
        self.time_label = tk.Label(
            time_frame,
            text="00:00:00/00:00:00",
            font=value_font,
            bg="white",
            fg="#000000",
            width=18,
            anchor=tk.W,
            padx=6,
            pady=2
        )
        self.time_label.pack()
        
        # 第四行：迭代/秒（总速度）- 删除冒号
        tk.Label(
            info_frame,
            text="迭代/秒",
            font=label_font,
            bg="#E8E8E8",  # 灰色背景
            fg="#505050",
            anchor=tk.W
        ).grid(row=3, column=0, sticky=tk.W, padx=(0, 8), pady=4)
        
        # 白色凹陷输入框
        speed_frame = tk.Frame(
            info_frame,
            bg="white",
            relief=tk.SUNKEN,
            bd=2,
            highlightbackground="#A0A0A0"
        )
        speed_frame.grid(row=3, column=1, sticky=tk.EW, pady=4)
        
        self.speed_label = tk.Label(
            speed_frame,
            text="0.0",
            font=value_font,
            bg="white",
            fg="#000000",
            width=18,
            anchor=tk.W,
            padx=6,
            pady=2
        )
        self.speed_label.pack()
        
        # 配置grid列的权重
        info_frame.columnconfigure(0, weight=0, minsize=85)
        info_frame.columnconfigure(1, weight=1)
        
        # 绑定窗口关闭事件
        self.root.protocol("WM_DELETE_WINDOW", self._on_window_close)
        
        # 窗口已初始化完成
        self.is_initialized = True
    
    def _on_window_close(self):
        """窗口关闭事件处理"""
        if not self.cancelled:
            response = messagebox.askyesno(
                "确认停止",
                "您确定要关闭窗口并停止模拟吗？\n\n已完成的模拟结果将会保存。",
                parent=self.root
            )
            if response:
                self.cancelled = True
                self._closed_by_window_x = True
                self.canvas.itemconfig(self.progress_text_id, text="窗口关闭")
                
                # 设置取消事件
                if self.cancel_event:
                    self.cancel_event.set()
                
                print("用户关闭窗口，模拟将被中断")
                
                # 立即更新窗口
                self.root.update()
                
                # 等待一小段时间确保消息显示
                time.sleep(0.1)
        
        # 销毁窗口
        try:
            self.root.destroy()
        except:
            pass
    
    def show(self):
        """显示窗口（在初始化完成后调用）"""
        if self.root and not self._first_update:
            self.root.deiconify()
            self.root.update()
    
    def start_new_scenario(self, scenario_index: int):
        """开始新场景"""
        self.current_scenario_index = scenario_index
        self.current_scenario_iteration = 0
        self.scenario_start_time = time.time()
        
        # 更新场景标签
        self.scenario_label.config(text=f"{self.current_scenario_index+1}/{self.n_scenarios}")
        
        # 更新迭代次数标签
        self.iter_label.config(text=f"0/{self.n_iterations_per_scenario:,}")
        
        # 第一次更新时显示窗口
        if self._first_update and self.is_initialized:
            self.root.deiconify()
            self.root.update()
            self._first_update = False
        
        # 强制更新窗口
        self.root.update_idletasks()
    
    def set_cancel_event(self, cancel_event):
        """设置取消事件"""
        self.cancel_event = cancel_event
    
    def update(self, current_iteration_in_scenario):
        """更新进度 - current_iteration_in_scenario是当前场景内的迭代次数"""
        # 如果已经取消，不再更新进度
        if self.cancelled:
            return
        
        # 更新当前场景的迭代次数
        self.current_scenario_iteration = current_iteration_in_scenario
        
        # 计算总已完成迭代次数
        previous_scenarios_iterations = self.current_scenario_index * self.n_iterations_per_scenario
        total_completed_iterations = previous_scenarios_iterations + self.current_scenario_iteration
        
        # 计算总进度百分比（所有场景的总迭代次数）
        if self.total_iterations > 0:
            total_percent = min(100.0, (total_completed_iterations / self.total_iterations) * 100)
        else:
            total_percent = 0
        
        # 计算当前场景的百分比（仅用于显示迭代次数）
        if self.n_iterations_per_scenario > 0:
            scenario_percent = min(100, int((self.current_scenario_iteration / self.n_iterations_per_scenario) * 100))
        else:
            scenario_percent = 0
        
        # 更新Canvas进度条和文本
        progress_width = int(self.canvas_width * total_percent / 100.0)
        
        # 更新进度条前景宽度
        self.canvas.coords(self.progress_fg, 0, 0, progress_width, self.canvas_height)
        
        # 更新进度文本 - 字体大小已改为11
        self.canvas.itemconfig(self.progress_text_id, text=f"模拟中 {total_percent:.2f}%", fill="black")
        
        # 更新迭代次数标签（当前场景）
        self.iter_label.config(text=f"{self.current_scenario_iteration:,}/{self.n_iterations_per_scenario:,}")
        
        # 计算性能信息 - 使用总时间
        current_time = time.time()
        total_elapsed_time = current_time - self.start_time
        
        # 格式化已耗时
        elapsed_hours = int(total_elapsed_time // 3600)
        elapsed_minutes = int((total_elapsed_time % 3600) // 60)
        elapsed_seconds = int(total_elapsed_time % 60)
        elapsed_str = f"{elapsed_hours:02d}:{elapsed_minutes:02d}:{elapsed_seconds:02d}"
        
        # 计算预计总时间（基于总进度）
        if total_completed_iterations > 0 and total_elapsed_time > 0 and self.total_iterations > 0:
            if total_percent > 0:
                # 计算总估计时间
                total_estimated_time = total_elapsed_time / (total_percent / 100.0)
                
                # 格式化总估计时间
                total_hours = int(total_estimated_time // 3600)
                total_minutes = int((total_estimated_time % 3600) // 60)
                total_seconds = int(total_estimated_time % 60)
                total_time_str = f"{total_hours:02d}:{total_minutes:02d}:{total_seconds:02d}"
                
                # 更新时间标签
                self.time_label.config(text=f"{elapsed_str}/{total_time_str}")
            else:
                self.time_label.config(text=f"{elapsed_str}/00:00:00")
        else:
            self.time_label.config(text=f"{elapsed_str}/00:00:00")
        
        # 计算总速度（总迭代次数/总时间）
        if total_elapsed_time > 0 and total_completed_iterations > 0:
            # 计算总速度
            total_speed = total_completed_iterations / total_elapsed_time
            self.speed_label.config(text=f"{total_speed:,.1f}")
        else:
            self.speed_label.config(text="0.0")
        
        # 提高UI更新频率 - 每次迭代都更新，但使用after避免阻塞
        try:
            # 使用after方法避免阻塞，提高响应性
            self.root.after(1, lambda: None)
            self.root.update_idletasks()
            
            # 每10次迭代强制更新一次
            if current_iteration_in_scenario % 10 == 0:
                self.root.update()
        except tk.TclError:
            # 如果窗口已被销毁，停止更新
            pass
    
    def set_message(self, message):
        """设置消息（兼容性方法）"""
        pass
    
    def set_detail_message(self, message):
        """设置详细消息（兼容性方法）"""
        pass
    
    def set_warning_message(self, message):
        """设置警告消息（兼容性方法）"""
        pass
    
    def set_success_message(self, message):
        """设置成功消息（兼容性方法）"""
        pass
    
    def close(self):
        """正常关闭窗口（模拟完成时调用）"""
        try:
            # 先显示完成消息
            self.canvas.coords(self.progress_fg, 0, 0, self.canvas_width, self.canvas_height)
            self.canvas.itemconfig(self.progress_text_id, text="模拟完成 100.00%", fill="black")
            
            # 等待短暂时间显示完成消息
            self.root.update()
            time.sleep(1.0)
            
            self.root.quit()
            self.root.destroy()
        except:
            pass
    
    def is_cancelled(self):
        """检查是否取消"""
        return self.cancelled
    
    def was_closed_by_window_x(self):
        """检查是否通过窗口关闭按钮取消"""
        return self._closed_by_window_x