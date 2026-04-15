# com_fixer.py
"""COM缓存修复模块 - 增强版（自动检测并修复缓存损坏）"""

import os
import shutil
import time
import win32com.client
import pythoncom
import pywintypes
from pyxll import xl_app

def _clean_com_cache_completely():
    """完全清理 COM 缓存（磁盘 + 内存），并强制重建类型库缓存"""
    try:
        # 获取所有可能的 gen_py 目录
        temp_dir = os.environ.get('TEMP', '')
        local_appdata = os.environ.get('LOCALAPPDATA', '')
        
        # 多个可能的缓存位置
        cache_dirs = [
            os.path.join(temp_dir, 'gen_py'),
            os.path.join(local_appdata, 'Temp', 'gen_py'),
            os.path.join(os.path.expanduser('~'), 'AppData', 'Local', 'Temp', 'gen_py'),
        ]
        
        # 删除所有找到的 gen_py 目录
        deleted_dirs = []
        for cache_dir in cache_dirs:
            if os.path.exists(cache_dir):
                try:
                    print(f"删除 COM 缓存目录: {cache_dir}")
                    shutil.rmtree(cache_dir, ignore_errors=True)
                    deleted_dirs.append(cache_dir)
                except Exception as e:
                    print(f"删除缓存目录 {cache_dir} 失败: {e}")
        
        # 确保 COM 已初始化
        try:
            pythoncom.CoInitialize()
        except:
            pass  # 如果已经初始化，忽略
        
        # 强制重建所有类型库缓存
        try:
            from win32com.client import gencache
            gencache.Rebuild()
            print("已重建 COM 类型库缓存")
        except Exception as e:
            print(f"重建 COM 缓存时出错: {e}")
        
        # 清除 win32com 动态类缓存
        try:
            import win32com.client.dynamic
            if hasattr(win32com.client.dynamic, '_dynamic_classes'):
                win32com.client.dynamic._dynamic_classes.clear()
                print("已清除动态类缓存")
        except:
            pass
        
        return True, deleted_dirs
        
    except Exception as e:
        print(f"清理 COM 缓存失败: {e}")
        return False, []

def _safe_excel_app():
    """安全获取 Excel 应用对象，自动修复 COM 缓存问题"""
    max_retries = 3
    for attempt in range(max_retries):
        try:
            # 首先尝试正常获取
            print(f"第 {attempt+1} 次尝试获取 Excel 应用对象...")
            app = xl_app()
            # 验证应用对象是否有效
            version = app.Version
            print(f"成功获取 Excel 应用对象，版本: {version}")
            return app
        except (AttributeError, pywintypes.com_error, Exception) as e:
            error_str = str(e)
            print(f"获取 Excel 应用对象失败: {error_str}")

            # 检查是否是 COM 缓存问题（扩展关键词列表）
            com_cache_keywords = [
                "CLSIDToPackageMap", "gen_py", "CLSID", "module",
                "has no attribute", "PackageMap", "Dispatch failed",
                "corrupted", "cache"
            ]
            if any(keyword in error_str for keyword in com_cache_keywords):
                print("检测到 COM 缓存问题，执行自动修复...")
                
                # 清理缓存
                success, deleted_dirs = _clean_com_cache_completely()
                
                if success and deleted_dirs:
                    print(f"已清理 {len(deleted_dirs)} 个缓存目录")
                
                # 等待缓存清理完成
                wait_time = 3 if attempt == 0 else 5
                print(f"等待 {wait_time} 秒让缓存清理生效...")
                time.sleep(wait_time)
                
                # 如果是最后一次尝试，尝试使用 win32com 直接获取
                if attempt == max_retries - 1:
                    print("尝试使用 win32com 直接获取 Excel 应用对象...")
                    try:
                        # 确保 COM 已初始化
                        pythoncom.CoInitialize()
                        app = win32com.client.Dispatch("Excel.Application")
                        # 验证应用对象
                        version = app.Version
                        print(f"使用 win32com 成功获取 Excel 应用对象，版本: {version}")
                        return app
                    except Exception as e3:
                        print(f"使用 win32com 直接获取也失败: {e3}")
                        
                        # 最后的尝试：使用 GetActiveObject 获取现有实例
                        try:
                            print("尝试使用 GetActiveObject 获取现有 Excel 实例...")
                            app = win32com.client.GetActiveObject("Excel.Application")
                            version = app.Version
                            print(f"使用 GetActiveObject 成功获取 Excel 应用对象，版本: {version}")
                            return app
                        except Exception as e4:
                            print(f"使用 GetActiveObject 也失败: {e4}")
                            raise Exception(f"所有方法都失败，无法获取 Excel 应用对象: {e4}")
            else:
                # 不是 COM 缓存问题，如果还有重试次数则继续，否则抛出
                if attempt == max_retries - 1:
                    raise
                # 短暂等待后重试
                time.sleep(1)
    
    # 如果所有尝试都失败
    raise Exception("无法获取 Excel 应用对象，所有修复尝试都失败了")

# ==================== 可选：启动时自动清理（建议仅在需要时启用） ====================
# 如果希望每次加载插件时都自动清理缓存，可以取消下面代码的注释。
# 注意：可能会稍微增加启动时间，且可能影响同一机器上其他 Python 进程的 COM 使用。
#
# def _clean_on_startup():
#     try:
#         temp_dir = os.environ.get('TEMP', '')
#         gen_py_path = os.path.join(temp_dir, 'gen_py')
#         if os.path.exists(gen_py_path):
#             shutil.rmtree(gen_py_path, ignore_errors=True)
#             print(f"启动时清理 gen_py 缓存: {gen_py_path}")
#     except:
#         pass
# 
# _clean_on_startup()