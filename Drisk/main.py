# main.py
"""Drisk蒙特卡洛模拟系统 - 高性能迭代版本"""

import warnings
warnings.filterwarnings("ignore", message="Inconsistent type specified for")
from pyxll import xl_func, xl_macro, xl_app, xlcAlert, xl_arg
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import traceback
import re
from typing import Dict, List, Optional, Any, Union
import threading
import time
import tkinter as tk
from tkinter import ttk
import sys
import os
import atexit
import pickle
import json
import shutil
import win32com.client
from collections import OrderedDict

# 导入所有模块
from constants import *
from com_fixer import *
from simulation_manager import *
from cell_utils import *
from formula_parser import *
from dependency_tracker import *
from distribution_functions import *
from attribute_functions import *
from statistical_functions import *
from simulation_engine import *
from progress_window import ProgressWindow
from info_window import InfoWindow
from macros import *
from sampling_functions import * # type: ignore
from statistical_functions_theo import *  # type: ignore
from index_functions import *  # type: ignore

# 导入PyXLL API模块
try:
    from pyxll_api import * # type: ignore
    PYXLL_API_AVAILABLE = True
except ImportError:
    PYXLL_API_AVAILABLE = False
    print("警告: PyXLL API模块不可用，高性能功能受限")

# 主函数：用于在PyXLL中注册所有函数
def register_functions():
    """注册所有函数到PyXLL"""
    # 这个函数由PyXLL自动调用，所有标有@xl_func或@xl_macro的函数都会被自动注册
    pass

# 初始化消息
print("=" * 60)
print("Drisk蒙特卡洛模拟框架已加载")
print("=" * 60)

print("\n可用随机函数:")
print("  DriskNormal(均值, 标准差, [属性函数...])")
print("  DriskUniform(最小值, 最大值, [属性函数...])")
print("  DriskGamma(形状参数, 尺度参数, [属性函数...])")
print("  DriskPoisson(λ参数, [属性函数...])")
print("  DriskBeta(α, β, [属性函数...])")
print("  DriskChiSq(自由度, [属性函数...])")
print("  DriskF(自由度1, 自由度2, [属性函数...])")
print("  DriskStudent(自由度, [属性函数...])")
print("  DriskExpon(期望值, [属性函数...])")
print("=" * 60)