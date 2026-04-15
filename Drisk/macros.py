# macros.py
"""宏函数模块 - 优化版本，调整批处理大小逻辑，支持DriskMakeInput，删除进度窗口，使用状态栏显示进度"""

from attribute_functions import set_static_mode, get_static_mode
import time
import atexit
import threading
import numpy as np
import traceback
import shutil
import os
from pyxll import xl_macro, xl_func, xlcAlert, xl_on_open
from constants import MIN_ITERATIONS, MAX_ITERATIONS, SAMPLING_MC, SAMPLING_LHC
from com_fixer import _safe_excel_app, _clean_com_cache_completely
from simulation_manager import (
    clear_simulations, create_simulation, get_simulation, 
    get_all_simulations, get_current_sim_id
)
# 在文件顶部添加导入
from model_functions import * # type: ignore
from dependency_tracker import find_all_simulation_cells_in_workbook
from simulation_engine import iterative_simulation_workbook, clear_dependency_cache
from info_window import InfoWindow

# 导入win32com用于批量操作
import win32com.client
import pythoncom

# 特殊错误标记
ERROR_MARKER = "#ERROR!"

def _clean_pycache():
    """
    清理当前目录和子目录中的 __pycache__ 文件夹
    """
    try:
        # 获取当前脚本所在目录
        current_dir = os.path.dirname(os.path.abspath(__file__))
        pycache_count = 0
        
        # 遍历当前目录及所有子目录
        for root, dirs, files in os.walk(current_dir):
            for dir_name in dirs:
                if dir_name == "__pycache__":
                    pycache_dir = os.path.join(root, dir_name)
                    try:
                        shutil.rmtree(pycache_dir)
                        print(f"已删除缓存目录: {pycache_dir}")
                        pycache_count += 1
                    except Exception as e:
                        print(f"删除缓存目录失败 {pycache_dir}: {e}")
        
        if pycache_count > 0:
            print(f"已清理 {pycache_count} 个 __pycache__ 目录")
        else:
            print("未找到需要清理的 __pycache__ 目录")
            
        return pycache_count
        
    except Exception as e:
        print(f"清理缓存时出错: {e}")
        return 0

# 注册退出时自动清理缓存
atexit.register(_clean_pycache)

# 修改 macros.py 中的 DriskSimtable 函数

@xl_func("var*: var", category="Drisk Functions", volatile=True)
def DriskSimtable(*args):
    """
    DriskSimtable函数 - 场景模拟值表
    
    参数:
        *args: 可变参数，可以是数值、数组引用或单元格区域引用
        
    返回:
        静态模式下：返回第一个值
        模拟模式下：根据场景索引返回对应的值
        
    示例:
        =DriskSimtable(1, 2, 3)                 # 直接数值
        =DriskSimtable({1, 2, 3})               # 数组常量
        =DriskSimtable(E17:E19)                 # 单元格区域引用
        =DriskSimtable(A1, A2, A3)              # 多个单元格引用
    """
    try:
        # 检查是否处于静态模式
        from attribute_functions import get_static_mode
        static_mode = get_static_mode()
        
        # 如果没有参数，返回0
        if len(args) == 0:
            return 0
        
        # 处理参数，提取所有值
        values = []
        
        # Excel传递参数的方式：
        # 1. 直接数值: (1.0, 2.0, 3.0) 或单个值
        # 2. 数组常量: (('0.8', '1.0', '1.2'),) 或类似结构
        # 3. 区域引用: ((1.0, 1.5, 2.0),) 或类似结构
        
        # 深度优先遍历所有参数，提取数值
        def extract_values(item):
            if isinstance(item, (list, tuple, np.ndarray)):
                # 如果是容器类型，递归提取
                for sub_item in item:
                    extract_values(sub_item)
            elif isinstance(item, str):
                # 字符串类型，尝试转换为数值
                try:
                    # 清理字符串
                    clean_item = item.strip()
                    # 处理可能的花括号（数组常量）
                    if clean_item.startswith('{') and clean_item.endswith('}'):
                        # 提取花括号内的内容
                        inner = clean_item[1:-1].strip()
                        # 分割逗号
                        parts = [p.strip() for p in inner.split(',') if p.strip()]
                        for part in parts:
                            try:
                                val = float(part)
                                values.append(val)
                            except:
                                pass
                    else:
                        # 普通字符串，尝试转换
                        val = float(clean_item)
                        values.append(val)
                except (ValueError, TypeError):
                    # 不能转换为数值，跳过
                    pass
            else:
                # 其他类型（通常是数值）
                try:
                    if item is not None:
                        val = float(item)
                        values.append(val)
                except (ValueError, TypeError):
                    # 不能转换为数值，跳过
                    pass
        
        # 提取所有参数的值
        for arg in args:
            extract_values(arg)
        
        # 如果没有有效值，返回0
        if not values:
            return 0
        
        # 确定要返回哪个索引的值
        if static_mode:
            # 静态模式：总是返回第一个值
            index_to_return = 0
        else:
            # 模拟模式：获取当前场景索引
            from simulation_manager import get_current_sim_id, get_simulation
            sim_id = get_current_sim_id()
            sim = get_simulation(sim_id) if sim_id else None
            
            if sim and hasattr(sim, 'is_scenario_simulation') and sim.is_scenario_simulation:
                # 多场景模拟：根据场景索引返回值
                scenario_index = getattr(sim, 'scenario_index', 0)
                index_to_return = scenario_index % len(values) if values else 0
            else:
                # 单场景模拟：返回第一个值
                index_to_return = 0
        
        # 返回对应索引的值
        if 0 <= index_to_return < len(values):
            return float(values[index_to_return])
        else:
            return float(values[0]) if values else 0
            
    except Exception as e:
        print(f"DriskSimtable函数出错: {str(e)}")
        import traceback
        traceback.print_exc()
        return 0

@xl_macro()
def DriskFixCOM():
    """
    修复COM类型库缓存问题
    """
    try:
        xlcAlert("正在修复COM类型库缓存问题，这可能需要几秒钟...\n\n如果问题持续，请重启Excel并重新运行此函数。")
        
        print("开始修复COM缓存...")
        success, deleted_dirs = _clean_com_cache_completely()
        
        if success:
            result_msg = "COM缓存修复成功！\n\n"
            if deleted_dirs:
                result_msg += f"已清理的缓存目录:\n"
                for dir_path in deleted_dirs:
                    result_msg += f"- {dir_path}\n"
                result_msg += "\n"
            
            result_msg += "修复措施已生效。\n\n"
            result_msg += "建议：\n"
            result_msg += "1. 如果问题仍然存在，请重启Excel\n"
            result_msg += "2. 然后重新运行DriskExample或其他Drisk函数\n"
            result_msg += "3. 如果问题持续，请检查Python环境和win32com安装"
            
            print(result_msg)
            xlcAlert(result_msg)
        else:
            xlcAlert("COM缓存修复失败。\n\n请尝试：\n1. 重启Excel\n2. 重新安装pywin32: pip install --upgrade pywin32\n3. 检查Python环境")
            
    except Exception as e:
        error_msg = f"修复COM缓存时出错: {str(e)}\n\n{traceback.format_exc()}"
        print(error_msg)
        xlcAlert(f"修复COM缓存时出错:\n{str(e)}")

# 修改StatusBarProgress类，添加ESC键检测功能
class StatusBarProgress:
    """在Excel状态栏显示进度信息，支持ESC键立即停止"""
    
    def __init__(self, app, n_iterations_per_scenario, n_scenarios):
        self.app = app
        self.n_iterations_per_scenario = n_iterations_per_scenario
        self.n_scenarios = n_scenarios
        self.total_iterations = n_iterations_per_scenario * n_scenarios
        self.completed_iterations = 0
        self.current_scenario_iteration = 0
        self.current_scenario_index = 0
        self.start_time = time.time()
        self.scenario_start_time = self.start_time
        self.cancelled = False
        self.esc_key_pressed = False
        self.last_update_time = 0
        self.last_esc_check_time = 0
        
        # 不再使用Windows钩子，改为直接检查
        print(f"状态栏进度显示器初始化完成: {n_iterations_per_scenario}次迭代, {n_scenarios}个场景")
        
    def start_new_scenario(self, scenario_index):
        """开始新场景"""
        self.current_scenario_index = scenario_index
        self.current_scenario_iteration = 0
        self.scenario_start_time = time.time()
        
        # 重置取消状态（但保留ESC键检测）
        self.cancelled = False
        
        # 在状态栏显示场景开始信息
        self.update_status_bar()
        
        print(f"开始场景 {scenario_index+1}/{self.n_scenarios}")
        
    def check_esc_key(self):
        """检查ESC键是否被按下（使用pywin32的GetAsyncKeyState）"""
        try:
            import win32api
            import win32con
            
            # 每100ms检查一次ESC键，避免频繁检查影响性能
            current_time = time.time()
            if current_time - self.last_esc_check_time < 0.1:  # 100ms间隔
                return False
                
            self.last_esc_check_time = current_time
            
            # 检查ESC键状态 (VK_ESCAPE = 0x1B)
            # 使用GetAsyncKeyState获取异步按键状态
            # 高位表示按键是否被按下（1表示按下）
            key_state = win32api.GetAsyncKeyState(win32con.VK_ESCAPE)
            
            # 如果ESC键被按下（高位为1）
            if key_state & 0x8000:
                print("检测到ESC键被按下")
                self.esc_key_pressed = True
                self.cancelled = True
                return True
                
        except Exception as e:
            # 如果无法检测ESC键，记录错误但不中断模拟
            print(f"检查ESC键失败: {str(e)}")
            
        return False
    
    def update(self, current_iteration_in_scenario):
        """更新进度"""
        # 检查ESC键
        if not self.cancelled:
            self.check_esc_key()
        
        # 如果已经取消，不再更新进度
        if self.cancelled:
            return
        
        # 更新当前场景的迭代次数
        self.current_scenario_iteration = current_iteration_in_scenario
        
        # 计算总已完成迭代次数
        previous_scenarios_iterations = self.current_scenario_index * self.n_iterations_per_scenario
        total_completed_iterations = previous_scenarios_iterations + self.current_scenario_iteration
        
        # 每10次迭代或1秒更新一次状态栏，避免频繁更新
        current_time = time.time()
        if current_iteration_in_scenario % 10 == 0 or (current_time - self.last_update_time) > 1.0:
            self.update_status_bar()
            self.last_update_time = current_time
        
        # 检查是否应该取消
        if self.esc_key_pressed:
            self.cancelled = True
            print("模拟因ESC键被取消")
    
    def update_status_bar(self):
        """更新Excel状态栏显示"""
        try:
            # 计算进度百分比
            if self.total_iterations > 0:
                total_completed = self.current_scenario_index * self.n_iterations_per_scenario + self.current_scenario_iteration
                total_percent = min(100.0, (total_completed / self.total_iterations) * 100)
            else:
                total_percent = 0
            
            # 计算场景进度百分比
            if self.n_iterations_per_scenario > 0:
                scenario_percent = min(100, int((self.current_scenario_iteration / self.n_iterations_per_scenario) * 100))
            else:
                scenario_percent = 0
            
            # 计算运行时间
            current_time = time.time()
            total_elapsed_time = current_time - self.start_time
            elapsed_hours = int(total_elapsed_time // 3600)
            elapsed_minutes = int((total_elapsed_time % 3600) // 60)
            elapsed_seconds = int(total_elapsed_time % 60)
            
            # 计算速度
            if total_elapsed_time > 0 and total_completed > 0:
                speed = total_completed / total_elapsed_time
            else:
                speed = 0
            
            # 构建状态栏消息
            if self.n_scenarios > 1:
                # 多场景模式
                status_msg = f"DRISK模拟进展: 总进度 {total_percent:.1f}%, 模拟次数 {self.current_scenario_index+1}/{self.n_scenarios}, "
                status_msg += f"迭代次数 {self.current_scenario_iteration:,}/{self.n_iterations_per_scenario:,} ({scenario_percent}%), "
                status_msg += f"时间 {elapsed_hours:02d}:{elapsed_minutes:02d}:{elapsed_seconds:02d}, "
                status_msg += f"迭代速度 {speed:.1f} 迭代/秒"
            else:
                # 单场景模式
                status_msg = f"DRISK模拟进展: 迭代次数 {self.current_scenario_iteration:,}/{self.n_iterations_per_scenario:,} ({scenario_percent}%), "
                status_msg += f"时间 {elapsed_hours:02d}:{elapsed_minutes:02d}:{elapsed_seconds:02d}, "
                status_msg += f"迭代速度 {speed:.1f} 迭代/秒"
            
            # 添加ESC提示
            if not self.cancelled:
                status_msg += " | 按ESC键立即停止"
            else:
                status_msg += " | 正在停止..."
            
            # 更新Excel状态栏
            self.app.StatusBar = status_msg
            
        except Exception as e:
            print(f"更新状态栏失败: {str(e)}")
    
    def clear(self):
        """清除状态栏"""
        try:
            # 恢复默认状态栏
            self.app.StatusBar = False
            print("状态栏已清除")
        except Exception as e:
            print(f"清除状态栏失败: {str(e)}")
    
    def is_cancelled(self):
        """检查是否取消"""
        return self.cancelled
    
    def was_cancelled_by_esc(self):
        """检查是否通过ESC键取消"""
        return self.esc_key_pressed
    
@xl_macro()
def DriskExample():
    """
    创建增强的测试示例 - 包含多种分布和复杂计算，以及Simtable和DriskMakeInput示例
    """
    try:
        # 使用安全的Excel应用获取方法
        app = _safe_excel_app()
        workbook = app.ActiveWorkbook
        
        # 创建工作表
        try:
            sheet = workbook.Worksheets("Drisk增强测试")
            sheet.Delete()
        except:
            pass
        
        sheet = workbook.Worksheets.Add()
        sheet.Name = "Drisk增强测试"
        
        # 标题和说明
        sheet.Range("A1").Value = "Drisk蒙特卡洛模拟 - 增强测试示例（含Simtable和DriskMakeInput）"
        sheet.Range("A1").Font.Bold = True
        sheet.Range("A1").Font.Size = 14
        
        sheet.Range("A3").Value = "示例说明:"
        sheet.Range("A3").Font.Bold = True
        sheet.Range("A4").Value = "这是一个完整的Drisk模拟测试示例，包含多种分布类型、Simtable函数、DriskMakeInput函数、多输入、多输出和复杂计算。"
        sheet.Range("A5").Value = "运行DriskRunMC()或DriskRunLHC()开始模拟，然后使用统计函数查看结果。"
        
        # ========== 输入区域 ==========
        sheet.Range("A7").Value = "输入区域（多种分布类型）:"
        sheet.Range("A7").Font.Bold = True
        
        sheet.Range("A8").Value = "收入 (正态分布):"
        sheet.Range("B8").Formula = '=DriskNormal(100, 20, DriskName("收入"), DriskUnits("万元"), DriskStatic(110))'
        sheet.Range("C8").Value = "静态值: 110"
        
        sheet.Range("A9").Value = "成本率 (均匀分布):"
        sheet.Range("B9").Formula = '=DriskUniform(0.5, 0.8, DriskName("成本率"), DriskUnits("比例"), DriskSeed(1, 123))'
        sheet.Range("C9").Value = "种子: 123"
        
        # ========== 测试单元格区域引用 ==========
        sheet.Range("E17").Value = 1.0
        sheet.Range("E18").Value = 1.5
        sheet.Range("E19").Value = 2.0
        
        # ========== Simtable区域 ==========
        sheet.Range("A11").Value = "Simtable区域（场景模拟 - 支持单元格区域引用）:"
        sheet.Range("A11").Font.Bold = True
        
        sheet.Range("A12").Value = "场景系数 (Simtable直接数值):"
        sheet.Range("B12").Formula = '=DriskSimtable(1, 2, 3)'
        sheet.Range("C12").Value = "场景值: 1,2,3"
        
        sheet.Range("A13").Value = "价格系数 (Simtable数组引用):"
        sheet.Range("B13").Formula = '=DriskSimtable({0.8, 1.0, 1.2})'
        sheet.Range("C13").Value = "数组: {0.8,1.0,1.2}"
        
        sheet.Range("A14").Value = "区域系数 (Simtable单元格区域引用):"
        sheet.Range("B14").Formula = '=DriskSimtable(E17:E19)'
        sheet.Range("C14").Value = "区域: E17:E19 = 1.0,1.5,2.0"
        
        # ========== MakeInput区域 ==========
        sheet.Range("A16").Value = "MakeInput区域（标记单元格值为输入）:"
        sheet.Range("A16").Font.Bold = True
        
        sheet.Range("A17").Value = "简单MakeInput:"
        sheet.Range("B17").Formula = '=DriskMakeInput(DriskNormal(5, 2), DriskName("正态输入"), DriskUnits("单位"))'
        sheet.Range("C17").Value = "记录DriskNormal(5,2)的值"
        
        sheet.Range("A18").Value = "组合MakeInput:"
        sheet.Range("B18").Formula = '=DriskMakeInput(DriskNormal(4,2)+C1, DriskName("组合输入"), DriskUnits("综合单位"))'
        sheet.Range("C18").Value = "记录DriskNormal(4,2)+C1的值"
        
        sheet.Range("A19").Value = "复杂MakeInput:"
        sheet.Range("B19").Formula = '=DriskMakeInput(DriskNormal(10,3)*DriskUniform(0.5,1.5), DriskName("乘积输入"), DriskCategory("计算输入"))'
        sheet.Range("C19").Value = "记录乘积的值"
        
        # 创建一些测试单元格
        sheet.Range("C1").Value = 100
        
        # ========== 中间计算 ==========
        sheet.Range("A21").Value = "中间计算（复杂公式）:"
        sheet.Range("A21").Font.Bold = True
        
        sheet.Range("A22").Value = "调整后收入 = 收入 × 区域系数:"
        sheet.Range("B22").Formula = '=B8 * B14'
        
        sheet.Range("A23").Value = "调整后成本率 = 成本率 × 价格系数:"
        sheet.Range("B23").Formula = '=B9 * B13'
        
        sheet.Range("A24").Value = "总收入 = 调整后收入 + MakeInput组合:"
        sheet.Range("B24").Formula = '=B22 + B18'
        
        sheet.Range("A25").Value = "总成本 = 总收入 × 调整后成本率:"
        sheet.Range("B25").Formula = '=B24 * B23'
        
        sheet.Range("A26").Value = "净利润 = 总收入 - 总成本:"
        sheet.Range("B26").Formula = '=B24 - B25'
        
        sheet.Range("A27").Value = "净利率 = 净利润 / 总收入:"
        sheet.Range("B27").Formula = '=IF(B24>0, B26/B24, 0)'
        
        # ========== 输出区域 ==========
        sheet.Range("A29").Value = "输出区域（多输出分析）:"
        sheet.Range("A29").Font.Bold = True
        
        sheet.Range("A30").Value = "净利润输出:"
        sheet.Range("B30").Formula = '=DriskOutput("净利润", "财务指标", 1, DriskUnits("万元"), DriskIsDiscrete(FALSE)) + B26'
        
        sheet.Range("A31").Value = "净利率输出:"
        sheet.Range("B31").Formula = '=DriskOutput("净利率", "财务比率", 2, DriskUnits("百分比"), DriskIsDiscrete(FALSE)) + B27'
        
        sheet.Range("A32").Value = "总收入输出:"
        sheet.Range("B32").Formula = '=DriskOutput("总收入", "收入指标", 3, DriskUnits("万元")) + B24'
        
        # ========== 统计区域 ==========
        sheet.Range("A34").Value = "统计函数（自动计算）:"
        sheet.Range("A34").Font.Bold = True
        
        # 净利润统计
        sheet.Range("A35").Value = "净利润统计:"
        sheet.Range("A36").Value = "均值:"; sheet.Range("B36").Formula = "=DriskMean(B30,1)"
        sheet.Range("A37").Value = "标准差:"; sheet.Range("B37").Formula = "=DriskStd(B30,1)"
        sheet.Range("A38").Value = "最小值:"; sheet.Range("B38").Formula = "=DriskMin(B30,1)"
        sheet.Range("A39").Value = "最大值:"; sheet.Range("B39").Formula = "=DriskMax(B30,1)"
        
        # ========== 格式化和美化 ==========
        # 设置列宽
        sheet.Columns("A").ColumnWidth = 30
        sheet.Columns("B").ColumnWidth = 25
        sheet.Columns("C").ColumnWidth = 20
        sheet.Columns("E").ColumnWidth = 10
        
        # 数字格式
        sheet.Range("B8:B9").NumberFormat = "#,##0.00"
        sheet.Range("B12:B14").NumberFormat = "#,##0.00"
        sheet.Range("B17:B19").NumberFormat = "#,##0.00"
        sheet.Range("B22:B27").NumberFormat = "#,##0.00"
        sheet.Range("B30:B32").NumberFormat = "#,##0.00"
        sheet.Range("B36:B39").NumberFormat = "#,##0.00"
        sheet.Range("E17:E19").NumberFormat = "#,##0.00"
        
        # 边框
        sheet.Range("A7:C9").Borders.LineStyle = 1
        sheet.Range("A11:C14").Borders.LineStyle = 1
        sheet.Range("A16:C19").Borders.LineStyle = 1
        sheet.Range("E17:E19").Borders.LineStyle = 1
        sheet.Range("A21:C27").Borders.LineStyle = 1
        sheet.Range("A29:C32").Borders.LineStyle = 1
        sheet.Range("A34:B39").Borders.LineStyle = 1
        
        # 颜色
        sheet.Range("A7:C7").Interior.Color = 0xCCFFCC  # 浅绿色
        sheet.Range("A11:C11").Interior.Color = 0xCCFFFF  # 浅青色（Simtable）
        sheet.Range("A16:C16").Interior.Color = 0xCCCCFF  # 浅蓝色（MakeInput）
        sheet.Range("E17:E19").Interior.Color = 0xCCFFFF  # 浅青色（区域引用）
        sheet.Range("A21:C21").Interior.Color = 0xFFFFCC  # 浅黄色
        sheet.Range("A29:C29").Interior.Color = 0xFFCCCC  # 浅红色
        sheet.Range("A34:B34").Interior.Color = 0xCCCCFF  # 浅蓝色
        
        # 修改说明
        sheet.Range("A41").Value = "功能更新说明:"
        sheet.Range("A41").Font.Bold = True
        
        sheet.Range("A42").Value = "1. Simtable现在支持单元格区域引用：=DriskSimtable(E17:E19)"
        sheet.Range("A43").Value = "2. MakeInput功能简化：不再支持复杂嵌套，只作为独立的输入标记函数"
        sheet.Range("A44").Value = "3. MakeInput示例：=DriskMakeInput(DriskNormal(4,2)+C1, DriskName('组合输入'))"
        sheet.Range("A45").Value = "4. MakeInput与分布函数共享输入编号系统"
        sheet.Range("A46").Value = "5. MakeInput单元格：记录整个单元格的模拟值，不解析内部公式"
        sheet.Range("A47").Value = "6. 运行模拟后，在DriskInfo中查看MakeInput输入"

        # 激活单元格
        sheet.Range("A1").Activate()
        
        # 设置自动计算
        app.Calculation = -4105  # xlCalculationAutomatic
        app.CalculateFull()
        
        alert_msg = (
            "Drisk增强测试示例（含Simtable区域引用和简化MakeInput）已创建成功！\n\n"
            "主要更新:\n"
            "1. Simtable函数支持单元格区域引用:\n"
            "   - =DriskSimtable(E17:E19)  # 引用E17:E19区域的值\n"
            "   - 区域内的值按顺序作为不同场景的值\n"
            "   - 也支持直接数值和数组常量\n\n"
            "2. MakeInput函数简化:\n"
            "   - 不再支持复杂嵌套形式（如DriskNormal()+DriskMakeInput()+DriskNormal()）\n"
            "   - 现在只作为独立的输入标记函数\n"
            "   - 参数格式：第一个是计算公式，后续是属性函数\n"
            "   - 示例：=DriskMakeInput(DriskNormal(4,2)+C1, DriskName('组合输入'))\n"
            "   - 在模拟中记录整个单元格的值\n"
            "   - 与分布函数共享输入编号系统\n\n"
            "运行DriskRunMC()开始模拟，然后使用DriskInfo()查看结果。"
        )
        xlcAlert(alert_msg)
        
    except Exception as e:
        error_msg = f"创建示例时出错: {str(e)}\n\n"
        error_msg += "建议运行DriskFixCOM()函数修复COM缓存问题。"
        xlcAlert(error_msg)

@xl_macro()
def DriskInfo():
    """显示模拟信息窗口，支持DriskMakeInput输入"""
    try:
        app = _safe_excel_app()
        
        try:
            sim_num_in = app.InputBox(Prompt="请输入模拟编号（默认为1）：", Title="DriskInfo - 模拟编号", Type=1, Default=1)
            if sim_num_in is False:
                return
            sim_num = int(sim_num_in)
        except Exception:
            sim_num = 1
        
        sim = get_simulation(sim_num)
        if sim is None:
            xlcAlert(f"模拟 {sim_num} 不存在")
            return
        
        info_window = InfoWindow(sim)
        info_window.root.mainloop()
        
    except Exception as e:
        xlcAlert(f"DriskInfo出错: {e}")

@xl_macro()
def DriskClear():
    """清除所有模拟数据"""
    clear_simulations()
    set_static_mode(True)
    
    try:
        app = _safe_excel_app()
        app.Calculation = -4105  # xlCalculationAutomatic
        app.CalculateFull()
    except:
        pass
    
    xlcAlert("所有模拟数据已清除，已恢复静态模式和自动计算")

@xl_macro()
def DriskClearCache():
    """清除依赖分析缓存"""
    clear_dependency_cache()
    xlcAlert("依赖分析缓存已清除")

@xl_macro()
def DriskListSims():
    """列出所有可用的模拟，包括DriskMakeInput信息"""
    try:
        result = "可用的模拟:\n\n"
        
        sim_cache = get_all_simulations()
        if sim_cache:
            for sim_id, sim in sim_cache.items():
                result += f"模拟 {sim_id}: {sim.name}\n"
                result += f"  抽样方法: {sim.sampling_method}\n"
                result += f"  迭代次数: {sim.n_iterations:,}\n"
                
                # 显示场景信息
                if sim.is_scenario_simulation:
                    result += f"  场景: {sim.scenario_index+1}/{sim.scenario_count}\n"
                
                # 显示输入类型信息
                dist_count = len(sim.distribution_cells) if hasattr(sim, 'distribution_cells') else 0
                makeinput_count = len(sim.makeinput_cells) if hasattr(sim, 'makeinput_cells') else 0
                result += f"  输入类型: {dist_count} 个分布函数, {makeinput_count} 个DriskMakeInput\n"
                
                result += f"  Input数量: {len(sim.all_input_keys)}\n"
                result += f"  Output数量: {len(sim.output_cells)}\n"
                result += f"  工作簿: {sim.workbook_name or '未知'}\n"
                result += f"  创建时间: {sim.timestamp}\n"
                
                if hasattr(sim, 'input_cache') and sim.input_cache:
                    result += f"  输入数据点: {len(sim.input_cache)}\n"
                if hasattr(sim, 'output_cache') and sim.output_cache:
                    result += f"  输出数据点: {len(sim.output_cache)}\n"
                
                result += "\n"
        else:
            result += "没有找到模拟数据\n"
        
        result += "提示: 运行DriskRunMC()或DriskRunLHC()创建新模拟"
        xlcAlert(result)
    except Exception as e:
        xlcAlert(f"列出模拟失败: {e}")

@xl_on_open
def auto_fix_com_on_open(*_args):
    """Excel打开工作簿时自动修复COM类型库缓存问题"""
    try:
        print("Excel启动：自动修复COM缓存...")
        success, deleted_dirs = _clean_com_cache_completely()
        if success:
            print(f"COM缓存自动修复成功，清理了 {len(deleted_dirs)} 个目录")
        else:
            print("COM缓存自动修复失败")
    except Exception as e:
        print(f"自动修复COM缓存时出错: {e}")