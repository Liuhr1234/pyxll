# ui_results_job_controller.py
"""
本模块提供结果视图（Results Views）的异步计算任务调度器 ResultsJobController。

主要功能：
1. 异步计算下发：将耗时的数学计算（如 KDE 核密度估计、百分位数统计、DKW 置信区间计算）转移至后台线程池（QThreadPool），防止阻塞主界面（UI Thread）。
2. 请求防抖（Debounce）：使用 QTimer 拦截用户高频触发的请求（例如连续拖拽范围滑块），只在动作停顿后才真正下发计算任务，节省 CPU 资源。
3. 令牌验证机制（Token-based Invalidation）：通过自增的唯一标识符（Token）管理任务。当后台任务完成并通过 Signal 返回时，控制器会核对该结果的 Token 是否为最新。若为旧请求（即被用户新操作覆盖的“过期结果”），则直接丢弃，彻底消除异步回调引发的数据竞态条件（Race Conditions）。
"""

import inspect
import numpy as np

from PySide6.QtCore import QObject, Signal, Slot, QThreadPool, QTimer

from ui_workers import StatsJob, KDEJob, CDFStatsWorker


# =======================================================
# 1. 异步任务调度控制器
# =======================================================
class ResultsJobController(QObject):
    """
    协调异步统计与 KDE (Kernel Density Estimation) 任务的中心枢纽。
    继承自 QObject 以支持 Qt 的信号槽（Signal/Slot）跨线程通信机制。
    确保最终发射给 UI 层的信号始终只包含“最新且有效的”计算结果。
    """

    # --- 定义跨线程通信信号 ---
    # 发射全局多曲线 KDE 结果：参数为 (X轴网格点数组, Y轴密度字典)
    pdf_kde_done = Signal(object, object)
    # 发射单曲线(如直方图背后的平滑曲线) KDE 结果：参数为 (变量标识CacheKey, X轴数组, Y轴密度字典)
    hist_kde_done = Signal(str, object, object)
    # 发射描述性统计面板结果：参数为 (变量名列表, 统计数据行列表)
    stats_done = Signal(list, list)
    # 发射 CDF 表格统计及置信区间结果：参数为 (格式化后的行数据列表)
    cdf_stats_done = Signal(list)

    def __init__(self, parent=None):
        super().__init__(parent)
        # 初始化 Qt 全局应用线程池，用于派发后台 Worker
        self._pool = QThreadPool(self)

        # ---------------------------------------------------------
        # PDF KDE (概率密度图) 任务调度管线
        # ---------------------------------------------------------
        self._pdf_token = 0                     # 当前最新的有效任务令牌
        self._pdf_debounce = QTimer(self)       # 防抖定时器
        self._pdf_debounce.setSingleShot(True)  # 单次触发模式，每次 start() 都会重置计时
        self._pdf_debounce.setInterval(300)     # 防抖时间：300毫秒
        self._pdf_debounce.timeout.connect(self._exec_pdf_kde)
        self._pdf_args = None                   # 暂存最新一次请求的参数，待定时器触发时使用

        # ---------------------------------------------------------
        # Histogram KDE (直方图附带的密度曲线) 任务调度管线
        # ---------------------------------------------------------
        self._hist_token = 0
        self._hist_debounce = QTimer(self)
        self._hist_debounce.setSingleShot(True)
        self._hist_debounce.setInterval(100)    # 直方图重绘要求响应更快，防抖设为 100毫秒
        self._hist_debounce.timeout.connect(self._exec_hist_kde)
        self._hist_args = None

        # ---------------------------------------------------------
        # Descriptive Stats (右侧描述性统计表格) 任务调度管线
        # ---------------------------------------------------------
        self._stats_token = 0
        self._stats_debounce = QTimer(self)
        self._stats_debounce.setSingleShot(True)
        self._stats_debounce.setInterval(200)   # 统计表格的防抖时间：200毫秒
        self._stats_debounce.timeout.connect(self._exec_stats)
        self._stats_args = None

        # ---------------------------------------------------------
        # CDF Stats (累积分布图表格) 任务调度管线
        # ---------------------------------------------------------
        # CDF 统计通常为即时响应或由其他动作连锁触发，故此处仅使用 Token 防止回调竞态，不设防抖定时器
        self._cdf_token = 0

    # =======================================================
    # 2. 概率密度曲线 (PDF KDE) 调度逻辑
    # =======================================================
    def request_pdf_kde(self, keys, datasets, x_min, x_max, pdf_points, immediate=False):
        """
        接收来自 UI 的 PDF KDE 计算请求。
        参数 immediate: 若为 True 则跳过防抖立即执行（通常用于图表初次加载）。
        """
        self._pdf_args = (keys, datasets, x_min, x_max, pdf_points)
        if immediate:
            self._exec_pdf_kde()
        else:
            self._pdf_debounce.start() # 启动或重置防抖定时器

    def _exec_pdf_kde(self):
        """真正向线程池派发 PDF KDE 任务的内部方法。"""
        if not self._pdf_args:
            return
        keys, datasets, x_min, x_max, pdf_points = self._pdf_args
        
        # 派发新任务前，递增全局 Token。
        # 旧任务如果在后台仍在运行，其持有的 Token 将小于当前最新的 self._pdf_token
        self._pdf_token += 1
        token = self._pdf_token
        
        # 创建后台 Worker
        # 传入一个 lambda 作为 cancel_check 回调：允许底层矩阵运算在发现 Token 过期时提前中断执行，节省算力
        job = KDEJob(token, datasets, keys, float(x_min), float(x_max), int(pdf_points), 8000, lambda: token != self._pdf_token)
        # 将 Worker 的完成信号绑定到本控制器的内部槽函数
        job.signals.done.connect(self._on_pdf_done_internal)
        
        self._pool.start(job)

    @Slot(int, object, object)
    def _on_pdf_done_internal(self, token, x_grid, y_map):
        """
        拦截后台传回的完成信号。
        [关键逻辑]：只有当回传的 token 匹配当前最新的 _pdf_token 时，才向 UI 层发射最终结果。
        被用户后续操作挤掉的“过期”结果将被静默丢弃。
        """
        if token == self._pdf_token:
            self.pdf_kde_done.emit(x_grid, y_map)

    # =======================================================
    # 3. 直方图附加密度曲线 (Hist KDE) 调度逻辑
    # =======================================================
    def request_hist_kde(self, cache_key, data, xmin, xmax, pdf_points, immediate=False):
        """接收直方图专属的 KDE 曲线计算请求。"""
        self._hist_args = (cache_key, data, xmin, xmax, pdf_points)
        if immediate:
            self._exec_hist_kde()
        else:
            self._hist_debounce.start()

    def _exec_hist_kde(self):
        """派发直方图 KDE 任务。结构与 PDF KDE 类似，但针对单条数据序列。"""
        if not self._hist_args:
            return
        cache_key, data, xmin, xmax, pdf_points = self._hist_args
        self._hist_token += 1
        token = self._hist_token
        
        # 将单序列包装为底层框架所需的数据集字典结构
        datasets = {cache_key: np.asarray(data, dtype=float)}
        job = KDEJob(token, datasets, [cache_key], float(xmin), float(xmax), int(pdf_points), 8000, lambda: token != self._hist_token)
        job.signals.done.connect(self._on_hist_done_internal)
        self._pool.start(job)

    @Slot(int, object, object)
    def _on_hist_done_internal(self, token, x_grid, y_map):
        if token == self._hist_token:
            # 提取缓存键名，兜底防护为空的情况
            cache_key = list(y_map.keys())[0] if y_map else "unknown"
            self.hist_kde_done.emit(cache_key, x_grid, y_map)

    # =======================================================
    # 4. 侧边描述性统计表格 (Stats) 调度逻辑
    # =======================================================
    def request_stats(self, keys, sim_id, label_to_cell_key, series_map, left_x, right_x, series_kind="output"):
        """接收统计表格数据刷新请求。将参数缓存并触发防抖计时。"""
        self._stats_args = (keys, sim_id, label_to_cell_key, series_map, left_x, right_x, series_kind)
        self._stats_debounce.start()

    def invalidate_stats(self):
        """
        主动使当前正在计算的统计任务失效。
        通常在 UI 发生重大跳转（如切换分析对象）时调用，以防旧对象的统计数据迟滞刷新到新视图上。
        """
        self._stats_token += 1

    def _exec_stats(self):
        """派发描述性统计运算（如均值、方差、分位数等）任务。"""
        if not self._stats_args:
            return
        keys, sim_id, label_to_cell_key, series_map, left_x, right_x, series_kind = self._stats_args
        self._stats_token += 1
        token = self._stats_token
        
        try:
            # 向后兼容性处理：检查 StatsJob 构造函数的签名，
            # 兼容带有底层大文件直读（基于 sim_id）的新版架构与仅传内存字典（series_map）的旧版架构。
            sig = inspect.signature(StatsJob.__init__)
            if "sim_id" in sig.parameters:
                job = StatsJob(token, sim_id=sim_id, label_to_cell_key=label_to_cell_key, keys=keys, left_x=left_x, right_x=right_x, series_kind=series_kind, cancel_check=lambda: token != self._stats_token)
            else:
                job = StatsJob(token, series_map, keys, left_x, right_x, cancel_check=lambda: token != self._stats_token)
        except Exception:
            # 最兜底的旧版兼容实例化
            job = StatsJob(token, series_map, keys, left_x, right_x)

        job.signals.done.connect(self._on_stats_done_internal)
        self._pool.start(job)

    @Slot(int, list, list)
    def _on_stats_done_internal(self, token, keys, stats_list):
        """拦截并验证统计结果 Token。有效则推送给 UI 渲染表格。"""
        if token == self._stats_token:
            self.stats_done.emit(keys, stats_list)

    # =======================================================
    # 5. CDF 置信区间及统计 (CDF Stats) 调度逻辑
    # =======================================================
    def request_cdf_stats(
        self,
        keys,
        series_map,
        ci_level,
        show_ci,
        engine_mode="fast",
        left_x=None,
        right_x=None,
        sim_id=None,
        label_to_cell_key=None,
        series_kind="output",
    ):
        """
        处理 CDF 图表下方的统计与非参置信区间（DKW 边界等）计算请求。
        该类任务通常响应级别较高，不设前端定时防抖，立即下发；但仍使用 Token 防止晚到的旧结果覆盖。
        """
        self._cdf_token += 1
        token = self._cdf_token
        worker = CDFStatsWorker(
            token, series_map, keys, ci_level, show_ci, engine_mode,
            lambda: token != self._cdf_token, # 允许后台核心算法引擎轮询提前取消计算
            left_x, right_x, sim_id, label_to_cell_key, series_kind
        )
        worker.signals.done.connect(self._on_cdf_done_internal)
        self._pool.start(worker)

    @Slot(int, list)
    def _on_cdf_done_internal(self, token, rows):
        """拦截 CDF 统计计算完成信号。"""
        if token == self._cdf_token:
            self.cdf_stats_done.emit(rows)