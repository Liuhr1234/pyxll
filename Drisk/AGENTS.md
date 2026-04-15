# AGENTS.md

## 仓库说明

这是一个纯 Python 项目。

本仓库的项目结构比较扁平：
- 根目录下直接包含大量 `.py` 文件。
- 所有核心逻辑、功能实现、流程控制基本都在这些顶层 `.py` 文件中。
- 除 Python 文件外，目录中通常只有：
  - `AGENTS.md`
  - `.idea/`
  - `docs/`
  - 图片资源文件夹（如 `icons/`）

这属于该项目的正常结构，不要误判为“项目不完整”或“缺少模块划分”。

---

## 文档与交接基线

本项目当前交接阶段的主文档集（Primary Documentation Set）为：
- `AGENTS.md`
- `docs/CODEBASE_MAP.md`
- `docs/WORKING_NOTES.md`
- `docs/RESULTS_ARCHITECTURE.md`
- `docs/RESULTS_HANDOFF.md`
- `docs/RESULTS_EXTENSION_GUIDE.md`
- `docs/RESULTS_VALIDATION.md`

其中：
- **受保护只读文件基线的权威来源是本文件（`AGENTS.md`）**。
- 其他文档可做引用或摘要，但若出现冲突，以本文件的受保护清单为准。

---

## 你的核心任务

你的任务不是机械记住所有 `.py` 文件内容，而是建立对以下内容的稳定理解：

1. 这个项目是做什么的。
2. 已经实现了哪些主要功能。
3. 哪些顶层 `.py` 文件分别承担什么职责。
4. 新需求到来时，应该优先查看哪些文件。
5. 哪些文件耦合高、脆弱、容易出问题。

后续需求会持续围绕这批 Python 文件迭代。你需要做到：能快速定位相关文件，并延续对项目功能的理解继续工作。

---

## 工作原则

### 1. 先理解，再修改
收到需求后，先不要直接改代码。请先：
- 判断需求可能涉及哪些顶层 `.py` 文件。
- 简要说明准备查看哪些文件。
- 阅读后总结当前实现。
- 再提出修改方案并实施。

### 2. 不要求一次性逐字读完所有文件
项目可能有很多 `.py` 文件，不需要逐字细读全部代码。

优先建立以下理解层次：
- 项目整体用途。
- 主要功能。
- 各顶层 `.py` 文件的大致职责。
- 关键流程由哪些文件串联。
- 常见需求应优先查看哪些文件。

### 3. 不确定时不要猜
如果某个结论没有被代码直接支持，请明确标注为：
- 已确认
- 推测
- 尚不明确，需要继续查看文件

### 4. 默认最小改动
除非用户明确要求，否则默认策略是：
- 优先最小改动。
- 优先沿用现有实现方式。
- 不随意重构。
- 不新增无必要依赖。
- 不随意改动公共行为和核心流程。

### 5. 兼容优先
- 保持关键入口与签名兼容（例如 UI 外部入口、桥接函数、导出合同）。
- 修改后应同步更新主文档集中的对应说明。

---

## 阅读优先级

由于本项目没有明显多层目录结构，请优先从顶层 `.py` 文件中识别以下类型：

1. 主入口文件
2. 主流程控制文件
3. 核心业务逻辑文件
4. 配置相关文件
5. 数据处理/状态处理文件
6. 工具函数文件
7. 与图片资源相关的处理文件

如果无法仅从文件名判断职责，请通过导入关系、主函数、调用链、类/函数命名推断文件定位。

---

## 输出要求

回答尽量结构化，并遵循以下规则。

### 分析需求时
优先说明：
- 本次需求可能影响哪些 `.py` 文件。
- 实际查看了哪些文件。
- 当前实现是什么。
- 准备如何修改。
- 风险点在哪里。

### 理解项目时
尽量引用具体文件名，不要泛泛而谈。

### 判断结论时
明确区分：
- 代码中可直接确认的事实。
- 基于调用关系做出的推测。

---

## 仓库结构判断规则

请特别注意：
- 根目录下有大量 `.py` 文件，是本项目的正常结构。
- 不要因为缺少 `src/`、`app/`、`tests/`、`config/` 等目录，就误判项目结构异常。
- 不要默认认为代码必须按包或多级目录组织。
- 本项目的理解重点应放在“顶层 Python 文件之间的关系”上。

---

## 初始任务（首次接触仓库时）

如果是首次接触该仓库，请先不要修改业务代码，先完成：

1. 扫描根目录重要 `.py` 文件。
2. 判断哪些文件最可能是入口、核心流程、关键功能文件。
3. 总结项目用途和已实现的主要功能。
4. 更新：
   - `docs/CODEBASE_MAP.md`
   - `docs/WORKING_NOTES.md`

其中：
- `CODEBASE_MAP.md` 侧重项目功能、文件职责、主流程、关键文件索引。
- `WORKING_NOTES.md` 侧重风险点、脆弱区域、后续协作建议、优先查看路径。

---

## 语言与注释约定

- 所有回复使用中文。
- 所有生成的 Markdown 文档使用中文。
- **后续新增或修改代码注释默认使用英文**（编码安全与跨环境可读性优先）。
- 除非用户明确要求，否则不要输出英文分析结论。

---

## 项目底层规则：受保护只读文件（Protected Read-Only Files）

以下文件为本项目默认受保护只读文件，作为长期协作基线规则：

1. 允许为理解逻辑而读取、检索、分析这些文件。
2. 未获得用户对“具体文件”的明确批准前，**禁止修改**这些文件。
3. 当根因定位确认必须修改受保护文件时，应先在结论中给出审批请求，再等待用户授权后实施。

受保护文件清单（对应顶层 `.py` 文件）：

- `statistical_functions_theo.py`
- `statistical_functions.py`
- `simulation_manager.py`
- `sampling_functions.py`
- `resolve_functions.py`
- `pyxll_api.py`
- `progress_window.py`
- `numpy_functions.py`
- `main.py`
- `macros.py`
- `model_functions.py`
- `info_window.py`
- `index_functions.py`
- `formula_parser.py`
- `distribution_functions.py`
- `distribution_base.py`
- `dist_trigen.py`
- `dist_triang.py`
- `dist_pearson5.py`
- `dist_pearson6.py`
- `dist_pareto2.py`
- `dist_pareto.py`
- `dist_negbin.py`
- `dist_lognorm2.py`
- `dist_lognorm.py`
- `dist_loglogistic.py`
- `dist_logistic.py`
- `dist_levy.py`
- `dist_laplace.py`
- `dist_kumaraswamy.py`
- `dist_johnsonsu.py`
- `dist_johnsonsb.py`
- `dist_invgauss.py`
- `dist_intuniform.py`
- `dist_hypsecant.py`
- `dist_hypergeo.py`
- `dist_histogrm.py`
- `dist_general.py`
- `dist_geomet.py`
- `dist_frechet.py`
- `dist_fatiguelife.py`
- `dist_extvaluemin.py`
- `dist_extvalue.py`
- `dist_erlang.py`
- `dist_erf.py`
- `dist_duniform.py`
- `dist_doubletriang.py`
- `dist_discrete.py`
- `dist_dagum.py`
- `dist_cumul.py`
- `dist_cauchy.py`
- `dist_binomial.py`
- `dist_bernoulli.py`
- `dist_reciprocal.py`
- `dist_rayleigh.py`
- `dist_pert.py`
- `dist_weibull.py`
- `dist_betageneral.py`
- `dist_betasubj.py`
- `dist_burr12.py`
- `dependency_tracker.py`
- `corrmat_functions.py`
- `constants.py`
- `com_fixer.py`
- `cell_utils.py`
- `cache_functions.py`
- `attribute_functions.py`
