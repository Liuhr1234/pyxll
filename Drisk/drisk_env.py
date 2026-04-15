# drisk_env.py
import os

def setup_environment():
    """
    Drisk 全局环境初始化
    必须在任何 PyQt/PySide6 以及 PyXLL 核心模块导入之前调用！
    """
    # 1. 彻底禁用 GPU 硬件加速 (规避 Excel 宿主环境下的黑屏/白屏/闪退)
    # 2. 禁用沙盒模式 (由于我们是作为插件寄宿在 EXCEL.EXE 中，Chromium 沙盒经常会因为权限冲突导致渲染进程崩溃)
    os.environ["QTWEBENGINE_CHROMIUM_FLAGS"] = "--disable-gpu --no-sandbox"
    
    # 3. 强制开启高 DPI 缩放，确保在 4K 或高缩放比显示器下 UI 不糊
    os.environ["QT_ENABLE_HIGHDPI_SCALING"] = "1"
    
    # 也可以在这里统一定义其他环境变量，比如消除一些无用的 Qt 终端警告
    os.environ["QT_LOGGING_RULES"] = "qt.webenginecontext.info=false"

# 模块加载时立即执行
setup_environment()