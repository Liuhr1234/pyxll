# -*- coding: utf-8 -*-
"""全局常量定义 - 增加分布支撑集信息"""

import math

PANEL_WBS_OVERRIDES = {}
DISTRIBUTION_PARAM_TOOLTIPS = {
    "Beta": {
        "α1": "是一个形状参数，必须大于 0。",
        "α2": "是一个形状参数，必须大于 0。"
    },
    "BetaGeneral": {
        "α1": "是一个形状参数，必须大于 0。",
        "α2": "是一个形状参数，必须大于 0。",
        "Min": "是分布的最小值，必须小于Max（最大值）。",
        "Max": "是分布的最大值，必须大于Min（最小值）。",
    },
    "BetaSubj": {
        "Min": "是分布的最小值，必须小于Max（最大值）。",
        "M.likely": "是分布的最可能值。",
        "Mean": "是分布的均值。",
        "Max": "是分布的最大值，必须大于Min（最小值）。",
    },
    "Burr12": {
        "γ": "是一个位置参数。",
        "β": "是一个尺度参数，必须大于 0。",
        "α1": "是一个形状参数，必须大于 0。",
        "α2": "是一个形状参数，必须大于 0。",
    },
    "Cauchy": {
        "γ": "是一个位置参数。",
        "β": "是一个尺度参数，必须大于 0。",
    },
    "ChiSq": {
        "V": "代表该分布的自由度，必须为正整数。",
    },
    "Cumul": {
        "Min": "是分布的最小值，必须小于Max（最大值）。",
        "Max": "是分布的最大值，必须大于Min（最小值）。",
        "X-Table": "表示曲线上各点对应的数值，必须严格按数值升序排列指定。",
        "P-Table": "表示曲线上各点对应的累积概率，必须严格按概率升序排列指定，每个点的取值应 >=0 且 <=1。",
    },
    "Dagum": {
        "γ": "是一个位置参数。",
        "β": "是一个尺度参数，必须大于 0。",
        "α1": "是一个形状参数，必须大于 0。",
        "α2": "是一个形状参数，必须大于 0。",
    },
    "DoubleTriang": {
        "Min": "是分布的最小值，必须小于Max（最大值）。",
        "M.likely": "是分布的最可能值。",
        "Max": "是分布的最大值，必须大于Min（最小值）。",
        "Lower P": "下三角分布的概率权重。",
    },
    "Erf": {
        "h": "代表方差参数，必须大于 0。",
    },
    "Erlang": {
        "m": "Gamma分布的整数型参数，必须为正整数。",
        "Beta": "是一个尺度参数，必须大于 0。",
    },
    "Expon": {
        "β": "是分布的均值，必须大于 0。",
    },
    "Extvalue": {
        "A": "是一个位置参数。",
        "B": "是一个尺度参数，必须大于 0。",
    },
    "ExtvalueMin": {
        "A": "是一个位置参数。",
        "B": "是一个尺度参数，必须大于 0。",
    },
    "F": {
        "V1": "是第一个自由度数值，必须是一个正整数。",
        "V2": "是第二个自由度数值，必须是一个正整数。",
    },
    "FatigueLife": {
        "γ": "是一个位置参数。",
        "β": "是一个尺度参数，必须大于 0。",
        "α": "是一个形状参数，必须大于 0。",
    },
    "Frechet": {
        "γ": "是一个位置参数。",
        "β": "是一个尺度参数，必须大于 0。",
        "α": "是一个形状参数，必须大于 0。",
    },
    "Gamma": {
        "α": "是一个形状参数，必须大于 0。",
        "β": "是一个尺度参数，必须大于 0。",
    },
    "General": {
        "Min": "是分布的最小值，必须小于Max（最大值）。",
        "Max": "是分布的最大值，必须大于Min（最小值）。",
        "X-Table": "是每个数据点的取值，必须按升序排列，且必须在该分布的最小值-最大值区间内。",
        "P-Table": "这是 {X} 中每个值对应的概率权重，用于指定该值处概率曲线的相对高度。",
    },
    "Histogrm": {
        "Min": "是分布的最小值，必须小于Max（最大值）。",
        "Max": "是分布的最大值，必须大于Min（最小值）。",
        "P-Table": "是该组内各个相应值的概率权重。必须是一个正数。",
    },
    "HypSecant": {
        "γ": "是一个位置参数。",
        "β": "是一个尺度参数，必须大于 0。",
    },
    "Invgauss": {
        "μ": "是分布的均值，必须大于 0。",
        "λ": "是一个形状参数，必须大于 0。",
    },
    "JohnsonMoments": {
        "μ": "是分布的均值。",
        "σ": "是分布的标准差，必须大于 0。",
        "s": "是分布的斜度。",
        "k": "这是衡量分布尾部厚度的指标，其取值必须大于 1。",
    },
    "JohnsonSB": {
        "α1": "是一个形状参数。",
        "α2": "是一个形状参数，必须大于 0。",
        "a": "是一个连续的边界参数，必须小于 b。",
        "b": "是一个连续的边界参数，必须大于 a。",
    },
    "JohnsonSU": {
        "α1": "是一个形状参数。",
        "α2": "是一个形状参数，必须大于 0。",
        "γ": "是一个位置参数。",
        "β": "是一个尺度参数，必须大于 0。",
    },
    "Kumaraswamy": {
        "α1": "是一个形状参数，必须大于 0。",
        "α2": "是一个形状参数，必须大于 0。",
        "Min": "是分布的最小值，必须小于Max（最大值）。",
        "Max": "是分布的最大值，必须大于Min（最小值）。",
    },
    "Laplace": {
        "μ": "是分布的均值。",
        "σ": "是分布的标准差，必须大于或等于 0。",
    },
    "Levy": {
        "a": "是一个位置参数。",
        "c": "是一个尺度参数，必须大于 0。",
    },
    "Logistic": {
        "α": "是一个位置参数。",
        "β": "是一个尺度参数，必须大于 0。",
    },
    "Loglogistic": {
        "γ": "是一个位置参数。",
        "β": "是一个尺度参数，必须大于 0。",
        "α": "是一个形状参数，必须大于 0。",
    },
    "Lognorm": {
        "μ": "是分布的均值，必须大于 0。",
        "σ": "是分布的标准差，必须大于 0。",
    },
    "Lognorm2": {
        "μ": "是底层正态分布的均值。",
        "σ": "是底层正态分布的标准差，必须大于 0。",
    },
    "Normal": {
        "μ": "是分布的均值。",
        "σ": "是分布的标准差，必须大于 0。",
    },
    "Pareto": {
        "θ": "是一个形状参数，必须大于 0。",
        "α": "是一个尺度参数，必须大于 0。",
    },
    "Pareto2": {
        "b": "是一个尺度参数，必须大于 0。",
        "q": "是一个形状参数，必须大于 0。",
    },
    "Pearson5": {
        "α": "是一个形状参数，必须大于 0。",
        "β": "是一个尺度参数，必须大于 0。",
    },
    "Pearson6": {
        "α1": "是一个形状参数，必须大于 0。",
        "α2": "是一个形状参数，必须大于 0。",
        "β": "是一个尺度参数，必须大于 0。",
    },
    "Pert": {
        "Min": "是分布的最小值，必须小于Max（最大值）。",
        "M.likely": "是分布的最可能值，作为形态参数的计算基准。",
        "Max": "是分布的最大值，必须大于Min（最小值）。",
    },
    "Rayleigh": {
        "b": "是分布的众数，必须大于 0。",
    },
    "Reciprocal": {
        "Min": "是分布的最小值，必须小于或等于Max（最大值）。",
        "Max": "是分布的最大值，必须大于或等于Min（最小值）。",
    },
    "Student": {
        "v": "是分布的自由度，必须是正整数。",
    },
    "Triang": {
        "Min": "是分布的最小值，必须小于Max（最大值）。",
        "M.likely": "是分布的最可能值。",
        "Max": "是分布的最大值，必须大于Min（最小值）。",
    },
    "Trigen": {
        "Bottom": "是计算底部分位数的临界值，表示在该值以下累积概率达到指定百分比。",
        "M.likely": "是分布的最可能值。",
        "Top": "是计算顶部分位数的临界值，表示在该值以上累积概率达到指定百分比。",
        "Bottom%": "指在分布中Bottom左侧区域面积占总面积的百分比。其取值必须大于0到100之间。",
        "Top%": "指在分布中Top左侧区域面积占总面积的百分比。其取值必须大于0到100之间。",
    },
    "Uniform": {
        "Min": "是分布的最小值，必须小于或等于Max（最大值）。",
        "Max": "是分布的最大值，必须大于或等于Min（最小值）。",
    },
    "Weibull": {
        "α": "是一个形状参数，必须大于 0。",
        "β": "是一个尺度参数，必须大于 0。",
    },
    "Bernoulli": {
        "p": "每次试验的成功率，其取值范围必须大于等于0且小于等于1。",
    },
    "Binomial": {
        "n": "试验或抽样的次数。",
        "p": "每次试验的成功概率，其取值范围必须大于等于0且小于等于1。",
    },
    "Discrete": {
        "X-Table": "指定每个可能结果对应的输入值。",
        "P-Table": "用于指定各结果出现概率的权重值。权重值 p 必须满足：所有元素 >=0，且总和 > 0。",
    },
    "DUniform": {
        "X-Table": "指定每个可能结果对应的输入值。",
    },
    "Geomet": {
        "p": "每次试验的成功率，其取值范围必须大于等于0且小于等于1。",
    },
    "Hypergeo": {
        "n": "样本容量，必须为正整数。",
        "D": "表示相关类型的物品数量，必须为正整数。",
        "M": "总体规模，必须为正整数值。",
    },
    "Intuniform": {
        "Min": "是分布的最小值，必须小于或等于Max（最大值）。",
        "Max": "是分布的最大值，必须大于或等于Min（最小值）。",
    },
    "Negbin": {
        "s": "成功次数，必须为正整数。",
        "p": "每次试验的成功概率，其取值范围必须大于等于0且小于等于1。",
    },
    "Poisson": {
        "λ": "是分布的均值，必须大于 0。",
    },
    "Compound": {
        "Frequency": "定义严重程度分布中抽取并累加的样本数量分布。",
        "Severity": "定义描述各样本严重程度的分布。",
        "Deductible": "定义免赔额的可选参数，默认值为不设置免赔额。",
        "Limit": "定义赔偿限额的可选参数，默认值为无限。",
    },
    "Splice": {
        "Left": "是一个分布，或指向包含该分布的单元格的引用。",
        "Right": "是一个分布，或指向包含该分布的单元格的引用。",
        "SplicePoint": "是两个分布拼接在一起的数值。",
    },
}
DISTRIBUTION_CARD_TOOLTIPS = {
    "Bernoulli": {
        "summary": "Bernoulli - 离散、有界分布",
        "detail": "用于模拟具有两种明确结果的事件：是/否、开/关、0/1等。常用于质量控制、可靠性分析、调查抽样及流行病学研究领域。"
    },
    "Binomial": {
        "summary": "Binomial - 离散、有界分布",
        "detail": "用于模拟在一系列独立试验中成功的次数，常见于质量控制、可靠性分析和调查抽样等领域。"
    },
    "Discrete":{
        "summary": "Discrete - 离散、有界分布",
        "detail": "一个通过概率表定义的离散型变量。"
    },
    "DUniform":{
        "summary": "DUniform - 离散、有界分布",
        "detail": "定义具有等概率结果的离散型变量，其取值由数值表指定。"
    },
    "Geomet":{
        "summary": "Geomet - 离散、左端有界分布",
        "detail": "用于模拟在首次成功前所需的试验次数，常见于质量控制和可靠性分析领域。"
    },
    "Hypergeo":{
        "summary": "Hypergeo - 离散、有界分布",
        "detail": "用于模拟在不放回抽样中成功的次数，常见于质量控制和概率论领域。"
    },
    "Intuniform":{
        "summary": "IntUniform - 离散、有界分布",
        "detail": "定义等概率离散变量，通常用于粗略估计"
    },
    "Negbin":{
        "summary": "Negbin - 离散、左端有界分布",
        "detail": "用于模拟在达到指定成功次数前所发生的失败次数，常见于质量控制、可靠性分析与调查抽样领域。"
    },
    "Poisson":{
        "summary": "Poisson - 离散、左端有界分布",
        "detail": "主要用于建模固定时间或空间间隔内独立事件的发生次数，常见于质量控制、可靠性分析及排队论领域。"
    },
    "Beta":{
        "summary": "Beta - 连续、有界分布",
        "detail": "取值区间在0到1之间，常用于表示变量的概率或比例，在贝叶斯推断和顺序统计量分析中具有重要应用价值。"
    },
    "BetaGeneral":{
        "summary": "BetaGeneral - 连续、有界分布",
        "detail": "用于建模两端有界的变量，常用于项目管理和成本分析领域。"
    },
    "BetaSubj":{
        "summary": "BetaSubj - 连续、有界分布",
        "detail": "用于建模两端有界的变量，常用于项目管理和成本分析领域。"
    },
    "Burr12":{
        "summary": "Burr12 - 连续、左端有界分布",
        "detail": "一种灵活的概率分布，用于建模非负随机变量，常用于家庭收入、保险理赔金额、可靠性数据及出行时间分析。"
    },
    "Cauchy":{
        "summary": "Cauchy - 连续、无界分布",
        "detail": "一种对称的概率分布，其尾部远厚重于正态分布，常用于极端事件建模及共振现象分析。"
    },
    "ChiSq":{
        "summary": "ChiSq - 连续、左端有界分布",
        "detail": "伽马分布的一种特例，主要应用于推断统计学和拟合优度检验。"
    },
    "Cumul":{
        "summary": "Cumul - 连续、无界分布",
        "detail": "一种通过数值与累积概率对照表定义的连续变量。"
    },
    "Dagum":{
        "summary": "Dagum - 连续、左端有界分布",
        "detail": "可作为Pareto分布与Lognorm分布的替代分布，其应用包括收入与财富建模、气象数据分析、可靠性研究及生存分析。"
    },
    "DoubleTriang":{
        "summary": "DoubleTriang - 连续、有界分布",
        "detail": "由两个三角分布混合而成，用于建立专家意见模型，常见于项目管理和成本分析领域。"
    },
    "Erf": {
        "summary": "Erf - 连续、无界分布",
        "detail": "作为零均值Normal分布的替代参数化形式，亦称Gauss Error Function。"
    },
    "Erlang":{
        "summary": "Erlang - 连续、左端有界分布",
        "detail": "Gamma分布的一种特例，其形状参数取整数值，常用于排队系统的建模。"
    },
    "Expon":{
        "summary": "Expon - 连续、左端有界分布",
        "detail": "用于建模Poisson过程中事件之间的时间间隔，常见于排队论和可靠性分析。"
    },
    "Extvalue":{
        "summary": "Extvalue - 连续、无界分布",
        "detail": "也被称为Gumbel Ⅰ类型分布，是广义极值分布的一种特例，常应用于金融风险建模、灾害最严重情景预测以及系统失效分析等领域。"
    },
    "ExtvalueMin":{
        "summary": "ExtvalueMin - 连续、无界分布",
        "detail": "也被称为Gumbel Ⅱ类型分布，是广义极值分布的一种特例，常应用于金融风险建模、灾害最严重情景预测以及系统失效分析等领域。"
    },
    "F":{
        "summary": "F - 连续、左端有界分布",
        "detail": "一种常作为检验统计量的零假设分布出现的非对称分布。主要应用于推断统计学领域，尤其在方差分析中较为突出。"
    },
    "FatigueLife":{
        "summary": "FatigueLife - 连续、左端有界分布",
        "detail": "一种常用于分析随时间推移失效情况的右偏分布，也称为Birnbaum-Saunders分布。"
        },
    "Frechet":{
        "summary": "Frechet - 连续、左端有界分布",
        "detail": "是广义极值分布的一种特例，常应用于金融风险建模、灾害最严重情景预测以及系统失效分析等领域。"
    },
    "Gamma":{
        "summary": "Gamma - 连续、左端有界分布",
        "detail": "常用于排队论、可靠性分析和贝叶斯统计学。"
    },  
    "General":{
        "summary": "General - 连续、有界分布",
        "detail": "一种通过数值与概率对照表定义的连续变量。"
    },
    "Histogrm":{
        "summary": "Histogrm - 连续、有界分布",
        "detail": "一种通过概率表定义的连续变量分布，该表格指定了直方图的概率结构。"
    },
    "HypSecant":{
        "summary": "HypSecant - 连续、无界分布",
        "detail": "类似于正态分布，但具有更尖锐的峰态，常用于金融收益率数据的建模分析。"
    },
    "Invgauss":{
        "summary": "Invgauss - 连续、左端有界分布",
        "detail": "主要应用于生存分析和可靠性分析，也被称为Wald分布。"
    },
    "JohnsonSB":{
        "summary": "JohnsonSB - 连续、有界分布",
        "detail": "有界的Johnson分布，适用于专家意见建模、项目管理及成本分析等领域。"
    },
    "JohnsonSU":{
        "summary": "JohnsonSU - 连续、无界分布",
        "detail": "无界的Johnson分布，适用于专家意见建模、项目管理及成本分析等领域。"
    },
    "Kumaraswamy":{
        "summary": "Kumaraswamy - 连续、有界分布",
        "detail": "与Beta分布密切相关，常用于项目管理与成本分析领域。"
    },
     "Laplace":{
        "summary": "Laplace - 连续、无界分布",
        "detail": "又称为双指数分布，是一种对称分布，其峰度较正态分布更为尖锐，常用于极端事件建模。"
    },
     "Levy":{
        "summary": "Levy - 连续、左端有界分布",
        "detail": "是一种用于建模非负随机变量的厚尾分布，广泛应用于金融风险与收益建模领域。"
     },
        "Logistic":{
            "summary": "Logistic - 连续、无界分布",
            "detail": "一种对称分布，其尾部比正态分布重得多，用途包括增长建模和逻辑回归。"
        },
        "Loglogistic":{
                "summary": "Loglogistic - 连续、左端有界分布",
                "detail": "类似于对数正态分布，但尾部更厚，用途包括生存分析以及财富和收入的分布情况分析。"
            },
        "Lognorm":{
            "summary": "Lognorm - 连续、左端有界分布",
            "detail": "一种随机变量，其对数服从正态分布，其用途包括随时间变化的增长情况分析以及可靠性分析。分布的参数是该分布的实际均值和标准差。"
        },
        "Lognorm2":{
            "summary": "Lognorm2 - 连续、左端有界分布",
            "detail": "一种随机变量，其对数服从正态分布。其用途包括对随时间变化的增长情况分析以及可靠性分析。分布的参数是其底层正态分布的均值和标准差。"
        },
        "Normal": {
            "summary": "Normal - 连续、无界分布",
            "detail": "经典的“钟形曲线”分布，是统计学分析中最核心的概率分布，其重要性源于中心极限定理的普适性。"
        },
        "Pareto":{
            "summary": "Pareto - 连续、左端有界分布",
            "detail": "是广义帕累托分布的一个特例，常用于以下领域建模：财富/收入分配、人口统计学、事件间隔时间。"
        },
        "Pareto2":{
            "summary": "Pareto2 - 连续、左端有界分布",
            "detail": "也称为广义帕累托分布，常应用于金融风险建模、灾害最严重情景预测以及系统失效分析等领域。"
        },
        "Pearson5":{
            "summary": "Pearson5 - 连续、左端有界分布",
            "detail": "是Gamma分布与Beta分布的变体，主要用于延迟时间建模和可靠性分析。"
        },
        "Pearson6":{
            "summary": "Pearson6 - 连续、左端有界分布",
            "detail": "是F分布的变体，专为左有界变量的灵活建模而设计。"
        },
        "Pert":{
            "summary": "Pert - 连续、有界分布",
            "detail": "用于建模双边界约束变量，主要用于项目管理和成本分析。"
         },
         "Reciprocal":{
             "summary": "Reciprocal - 连续、有界分布",
             "detail": "是一种双侧有界的右偏态分布，在数值分析、贝叶斯统计及噪声分析中具有重要应用。"
         },
         "Rayleigh":{
             "summary": "Rayleigh - 连续、左端有界分布",
             "detail": "是Weibull分布的一个特例，常用于风速/波浪高度、产品寿命和噪声信号分析。"
         },
          "Student":{
              "summary": "Student - 连续、无界分布",
              "detail": "是一种对称钟形分布，具有比正态分布更厚的尾部，是统计推断中的常用工具。"
          },
          "Triang":{
                "summary": "Triang - 连续、有界分布",
                "detail": "适用于三点估算和专家意见建模，常见于项目管理和成本分析领域。"
          },
          "Trigen":{
                "summary": "Trigen - 连续、有界分布",
                "detail": "适用于三点估算和专家意见建模，常见于项目管理和成本分析领域。"
          },
          "Uniform":{
            "summary": "Uniform - 连续、有界分布",
            "detail": "一种有界连续变量，其所有结果的发生概率相同，通常用于粗略估计。"
          },
          "Weibull":{
            "summary": "Weibull - 连续、左端有界分布",
            "detail": "是一种常用于极端事件建模、产品可靠性评估以及寿命与失效数据分析的灵活概率分布。"
          },
          "Splice":{
            "summary": "Splice - 连续、无界分布",
            "detail": "将两个分布按照你指定的“拼接点”合并在一起。"
          },
          "Compound":{
            "summary": "Compound - 连续、无界分布",
            "detail": "将来自“严重程度”分布的多个样本与“频率”分布进行组合控制，常用于操作风险和精算分析领域。"
          }
}
# 属性函数标记值 - 使用不同的浮点数值
ATTR_NAME_MARKER = 0.0
ATTR_LOC_MARKER = 1.0
ATTR_CATEGORY_MARKER = 0.1
ATTR_COLLECT_MARKER = 0.2
ATTR_CONVERGENCE_MARKER = 0.3
ATTR_COPULA_MARKER = 0.4
ATTR_CORRMAT_MARKER = 0.5
ATTR_FIT_MARKER = 0.6
ATTR_ISDATE_MARKER = 0.7
ATTR_ISDISCRETE_MARKER = 0.8
ATTR_LOCK_MARKER = 0.9
ATTR_SEED_MARKER = 1.1
ATTR_SHIFT_MARKER = 1.2
ATTR_STATIC_MARKER = 1.3
ATTR_TRUNCATE_MARKER = 1.4
ATTR_TRUNCATEP_MARKER = 1.5
ATTR_TRUNCATE2_MARKER = 1.6
ATTR_UNITS_MARKER = 1.7

# 默认模拟参数
DEFAULT_ITERATIONS = 5000
MIN_ITERATIONS = 100
MAX_ITERATIONS = 100000

# 采样方法
SAMPLING_MC = "MC"
SAMPLING_LHC = "LHC"
SAMPLING_SOBOL = "SOBOL"  # 可选，未来扩展

# 批次大小
DEFAULT_BATCH_SIZE = 1000
MAX_BATCH_SIZE = 10000
MIN_BATCH_SIZE = 10

# 分布函数名称
DIST_NORMAL = "DriskNormal"
DIST_UNIFORM = "DriskUniform"
DIST_ERF = "DriskErf"
DIST_EXTVALUE = "DriskExtvalue"
DIST_EXTVALUEMIN = "DriskExtvalueMin"
DIST_FATIGUELIFE = "DriskFatigueLife"
DIST_FRECHET = "DriskFrechet"
DIST_GENERAL = "DriskGeneral"
DIST_HISTOGRM = "DriskHistogrm"
DIST_HYPSECANT = "DriskHypSecant"
DIST_JOHNSONSB = "DriskJohnsonSB"
DIST_JOHNSONSU = "DriskJohnsonSU"
DIST_KUMARASWAMY = "DriskKumaraswamy"
DIST_LAPLACE = "DriskLaplace"
DIST_LOGISTIC = "DriskLogistic"
DIST_LOGLOGISTIC = "DriskLoglogistic"
DIST_LOGNORM = "DriskLognorm"
DIST_LOGNORM2 = "DriskLognorm2"
DIST_BETAGENERAL = "DriskBetaGeneral"
DIST_BETASUBJ = "DriskBetaSubj"
DIST_BURR12 = "DriskBurr12"
DIST_COMPOUND = "DriskCompound"
DIST_SPLICE = "DriskSplice"
DIST_PERT = "DriskPert"
DIST_RECIPROCAL = "DriskReciprocal"
DIST_RAYLEIGH = "DriskRayleigh"
DIST_WEIBULL = "DriskWeibull"
DIST_PEARSON5 = "DriskPearson5"
DIST_PEARSON6 = "DriskPearson6"
DIST_PARETO2 = "DriskPareto2"
DIST_PARETO = "DriskPareto"
DIST_LEVY = "DriskLevy"
DIST_CAUCHY = "DriskCauchy"
DIST_DAGUM = "DriskDagum"
DIST_DOUBLETRIANG = "DriskDoubleTriang"
DIST_GAMMA = "DriskGamma"
DIST_ERLANG = "DriskErlang"
DIST_POISSON = "DriskPoisson"
DIST_BETA = "DriskBeta"
DIST_CHISQ = "DriskChiSq"
DIST_F = "DriskF"
DIST_Student = "DriskStudent"
DIST_EXPON = "DriskExpon"
DIST_INVGAUSS = "DriskInvgauss"
DIST_DUNIFORM = "DriskDUniform"
DIST_GEOMET = "DriskGeomet"
DIST_HYPERGEO = "DriskHypergeo"
DIST_INTUNIFORM = "DriskIntuniform"
DIST_NEGBIN = "DriskNegbin"

# 分布类型常量
DIST_TYPE_NORMAL = "normal"
DIST_TYPE_UNIFORM = "uniform"
DIST_TYPE_ERF = "erf"
DIST_TYPE_EXTVALUE = "extvalue"
DIST_TYPE_EXTVALUEMIN = "extvaluemin"
DIST_TYPE_FATIGUELIFE = "fatiguelife"
DIST_TYPE_FRECHET = "frechet"
DIST_TYPE_GENERAL = "general"
DIST_TYPE_HISTOGRM = "histogrm"
DIST_TYPE_HYPSECANT = "hypsecant"
DIST_TYPE_JOHNSONSB = "johnsonsb"
DIST_TYPE_JOHNSONSU = "johnsonsu"
DIST_TYPE_KUMARASWAMY = "kumaraswamy"
DIST_TYPE_LAPLACE = "laplace"
DIST_TYPE_LOGISTIC = "logistic"
DIST_TYPE_LOGLOGISTIC = "loglogistic"
DIST_TYPE_LOGNORM = "lognorm"
DIST_TYPE_LOGNORM2 = "lognorm2"
DIST_TYPE_BETAGENERAL = "betageneral"
DIST_TYPE_BETASUBJ = "betasubj"
DIST_TYPE_BURR12 = "burr12"
DIST_TYPE_COMPOUND = "compound"
DIST_TYPE_SPLICE = "splice"
DIST_TYPE_PERT = "pert"
DIST_TYPE_RECIPROCAL = "reciprocal"
DIST_TYPE_RAYLEIGH = "rayleigh"
DIST_TYPE_WEIBULL = "weibull"
DIST_TYPE_PEARSON5 = "pearson5"
DIST_TYPE_PEARSON6 = "pearson6"
DIST_TYPE_PARETO2 = "pareto2"
DIST_TYPE_PARETO = "pareto"
DIST_TYPE_LEVY = "levy"
DIST_TYPE_CAUCHY = "cauchy"
DIST_TYPE_DAGUM = "dagum"
DIST_TYPE_DOUBLETRIANG = "doubletriang"
DIST_TYPE_GAMMA = "gamma"
DIST_TYPE_ERLANG = "erlang"
DIST_TYPE_POISSON = "poisson"
DIST_TYPE_BETA = "beta"
DIST_TYPE_CHISQ = "chisq"
DIST_TYPE_F = "f"
DIST_TYPE_Student = "student"
DIST_TYPE_EXPON = "expon"
DIST_TYPE_INVGAUSS = "invgauss"
DIST_TYPE_DUNIFORM = "duniform"
DIST_TYPE_GEOMET = "geomet"
DIST_TYPE_HYPERGEO = "hypergeo"
DIST_TYPE_INTUNIFORM = "intuniform"
DIST_TYPE_NEGBIN = "negbin"

# ==================== 分布注册表（增加支撑集信息） ====================
def _is_distribution_formula_string(value):
    if not isinstance(value, str):
        return False
    text = value.strip()
    if text.startswith("="):
        text = text[1:].strip()
    return text.startswith("Drisk") and "(" in text and text.endswith(")")


DISTRIBUTION_REGISTRY = {
    "DriskNormal": {
        "type": "normal", "min_params": 2, "max_params": 2,
        "param_names": ["μ", "σ"],
        "supports_shift_truncate": True,
        "category": "\u65e0\u754c",  # 无界
        "description":"定义一个经典的“钟形曲线”正态分布，广泛适用于描述大量数据集的统计分布特征。",
        "ui_name": "Normal - \u8fde\u7eed\u3001\u65e0\u754c\u5206\u5e03",  # Normal - 连续、无界分布
        "ui_param_labels": ["μ", "σ"],
        "param_descriptions": [
            "是分布的均值。",
            "是分布的标准差，必须大于0。",
        ],
        "default_params": [0.0, 1.0],
        "is_discrete": False,
        "validate_params": lambda params: params[1] > 0,
        "support": lambda params: (float('-inf'), float('inf'))
    },
    "DriskUniform": {
        "type": "uniform", "min_params": 2, "max_params": 2,
        "param_names": ["Min", "Max"],
        "supports_shift_truncate": True,
        "category": "\u6709\u754c",  # 有界
        "description": "定义一个均匀概率分布。在均匀分布的取值范围内，每个值出现的可能性都是相等的。",
        "ui_name": "Uniform - \u8fde\u7eed\u3001\u6709\u754c\u5206\u5e03",  # Uniform - 连续、有界分布
        "ui_param_labels": ["Min", "Max"],
        "param_descriptions": [
            "是分布的最小值，必须小于或等于Maximum（最大值）。",
            "是分布的最大值，必须大于或等于Minimum（最小值）。",
        ],
        "default_params": [0.0, 1.0],
        "is_discrete": False,
        "validate_params": lambda params: params[1] > params[0],
        "support": lambda params: (params[0], params[1])
    },
    "DriskErf": {
        "type": "erf", "min_params": 1, "max_params": 1,
        "param_names": ["h"],
        "supports_shift_truncate": True,
        "category": "\u65e0\u754c",
        "description": "定义一个具有方差参数H的Gauss Error函数，该分布派生自Normal分布。",
        "ui_name": "Erf - \u8fde\u7eed\u3001\u65e0\u754c\u5206\u5e03",
        "ui_param_labels": ["h"],
        "param_descriptions": [
            "代表方差参数，必须大于0。"
        ],
        "default_params": [1.0],
        "is_discrete": False,
        "validate_params": lambda params: len(params) >= 1 and float(params[0]) > 0,
        "support": lambda params: (float('-inf'), float('inf'))
    },
    "DriskExtvalue": {
        "type": "extvalue", "min_params": 2, "max_params": 2,
        "param_names": ["A", "B"],
        "supports_shift_truncate": True,
        "category": "\u65e0\u754c",
        "description": "定义一个Extvalue分布。",
        "ui_name": "Extvalue - \u8fde\u7eed\u3001\u65e0\u754c\u5206\u5e03",
        "ui_param_labels": ["A", "B"],
        "param_descriptions": [
            "是一个位置参数。",
            "是一个尺度参数，必须大于0。"
        ],
        "default_params": [0.0, 10.0],
        "is_discrete": False,
        "validate_params": lambda params: len(params) >= 2 and float(params[1]) > 0,
        "support": lambda params: (float('-inf'), float('inf'))
    },
    "DriskExtvalueMin": {
        "type": "extvaluemin", "min_params": 2, "max_params": 2,
        "param_names": ["A", "B"],
        "supports_shift_truncate": True,
        "category": "\u65e0\u754c",
        "description": "定义一个ExtvalueMin分布。",
        "ui_name": "ExtvalueMin - \u8fde\u7eed\u3001\u65e0\u754c\u5206\u5e03",
        "ui_param_labels": ["A", "B"],
        "param_descriptions": [
            "是一个位置参数。",
            "是一个尺度参数，必须大于0。"
        ],
        "default_params": [0.0, 10.0],
        "is_discrete": False,
        "validate_params": lambda params: len(params) >= 2 and float(params[1]) > 0,
        "support": lambda params: (float('-inf'), float('inf'))
    },
    "FatigueLife": {
        "type": "fatiguelife", "min_params": 3, "max_params": 3,
        "param_names": ["γ", "β", "α"],
        "supports_shift_truncate": True,
        "category": "\u5de6\u7aef\u6709\u754c",
        "description": "定义一个FatigueLife分布。",
        "ui_name": "FatigueLife - \u8fde\u7eed\u3001\u5de6\u7aef\u6709\u754c\u5206\u5e03",
        "ui_param_labels": ["γ", "β", "α"],
        "param_descriptions": [
            "是一个位置参数。",
            "是一个尺度参数，必须大于0。",
            "是一个形状参数，必须大于 0。"
        ],
        "default_params": [0.0, 10.0, 2.0],
        "is_discrete": False,
        "validate_params": lambda params: (
            len(params) >= 3 and float(params[1]) > 0 and float(params[2]) > 0
        ),
        "support": lambda params: (float(params[0]), float('inf'))
    },
    "DriskFrechet": {
        "type": "frechet", "min_params": 3, "max_params": 3,
        "param_names": ["γ", "β", "α"],
        "supports_shift_truncate": True,
        "category": "\u5de6\u7aef\u6709\u754c",
        "description": "定义一个Frechet分布。",
        "ui_name": "Frechet - \u8fde\u7eed\u3001\u5de6\u7aef\u6709\u754c\u5206\u5e03",
        "ui_param_labels": ["γ", "β", "α"],
        "param_descriptions": [
            "是一个位置参数。",
            "是一个尺度参数，必须大于0。",
            "是一个形状参数，必须大于 0。"
        ],
        "default_params": [0.0, 10.0, 2.0],
        "is_discrete": False,
        "validate_params": lambda params: (
            len(params) >= 3 and float(params[1]) > 0 and float(params[2]) > 0
        ),
        "support": lambda params: (float(params[0]), float('inf'))
    },
    "DriskGeneral": {
        "type": "general", "min_params": 4, "max_params": 4,
        "param_names": ["Min", "Max", "X-Table", "P-Table"],
        "supports_shift_truncate": True,
        "category": "\u6709\u754c",
        "description": "基于指定的 (x, p) 数据对所构建的密度曲线，生成一个广义概率分布。",
        "ui_name": "General - \u8fde\u7eed\u3001\u6709\u754c\u5206\u5e03",
        "ui_param_labels": ["Min", "Max", "X-Table", "P-Table"],
        "param_descriptions": [
            "是分布的最小值，必须小于Maximum（最大值）。",
            "是分布的最大值，必须大于Minimum（最小值）。",
            "是每个数据点的取值，必须按升序排列，且必须落在该分布的最小值-最大值区间内。",
            "这是 {x} 中每个值对应的概率权重，用于指定该值处概率曲线的相对高度。"
        ],
        "default_params": [-1.0, 1.0, "{-0.5,0,0.5}", "{2,3,2}"],
        "is_discrete": False,
        "validate_params": lambda params: len(params) >= 4 and float(params[0]) < float(params[1]),
        "support": lambda params: (float(params[0]), float(params[1]))
    },
    "DriskHistogrm": {
        "type": "histogrm", "min_params": 3, "max_params": 3,
        "param_names": ["Min", "Max", "P-Table"],
        "supports_shift_truncate": True,
        "category": "\u6709\u754c",
        "description": "定义一个用户自定义的直方图分布，该分布在最小值(minimum)和最大值(maximum)之间包含若干等宽区间，且每个区间具有相应的概率权重p。",
        "ui_name": "Histogrm - \u8fde\u7eed\u3001\u6709\u754c\u5206\u5e03",
        "ui_param_labels": ["Min", "Max", "P-Table"],
        "param_descriptions": [
            "是分布的最小值，必须小于Max（最大值）。",
            "是分布的最大值，必须大于Minimum（最小值）。",
            "是该组内各个相应值的概率权重。必须是一个正数。"
        ],
        "default_params": [-1.0, 1.0, "{0.1,0.2,0.4,0.2,0.1}"],
        "is_discrete": False,
        "validate_params": lambda params: len(params) >= 3 and float(params[0]) < float(params[1]),
        "support": lambda params: (float(params[0]), float(params[1]))
    },
    "DriskHypSecant": {
        "type": "hypsecant", "min_params": 2, "max_params": 2,
        "param_names": ["γ", "β"],
        "supports_shift_truncate": True,
        "category": "\u65e0\u754c",
        "description": "定义一个HypSecant分布。",
        "ui_name": "HypSecant - \u8fde\u7eed\u3001\u65e0\u754c\u5206\u5e03",
        "ui_param_labels": ["γ", "β"],
        "param_descriptions": [
            "是一个位置参数。",
            "是一个尺度参数，必须大于0。"
        ],
        "default_params": [0.0, 10.0],
        "is_discrete": False,
        "validate_params": lambda params: len(params) >= 2 and float(params[1]) > 0,
        "support": lambda params: (float('-inf'), float('inf'))
    },
    "DriskJohnsonSB": {
        "type": "johnsonsb", "min_params": 4, "max_params": 4,
        "param_names": ["α1", "α2", "a", "b"],
        "supports_shift_truncate": True,
        "category": "\u6709\u754c",
        "description": "定义一个JohnsonSB（系统有界）分布。",
        "ui_name": "JohnsonSB - \u8fde\u7eed\u3001\u6709\u754c\u5206\u5e03",
        "ui_param_labels": ["α1", "α2", "a", "b"],
        "param_descriptions": [
            "是一个形状参数。",
            "是一个形状参数，必须大于 0。",
            "是一个连续的边界参数，必须小于“B”。",
            "是一个连续的边界参数，必须大于“A”。"
        ],
        "default_params": [1.0, 1.0, -10.0, 100.0],
        "is_discrete": False,
        "validate_params": lambda params: (
            len(params) >= 4 and float(params[1]) > 0 and float(params[3]) > float(params[2])
        ),
        "support": lambda params: (float(params[2]), float(params[3]))
    },
    "DriskJohnsonSU": {
        "type": "johnsonsu", "min_params": 4, "max_params": 4,
        "param_names": ["α1", "α2", "γ", "β"],
        "supports_shift_truncate": True,
        "category": "\u65e0\u754c",
        "description": "定义一个JohnsonSU（系统无界）分布。",
        "ui_name": "JohnsonSU - \u8fde\u7eed\u3001\u65e0\u754c\u5206\u5e03",
        "ui_param_labels": ["α1", "α2", "γ", "β"],
        "param_descriptions": [
            "是一个形状参数。",
            "是一个形状参数，必须大于 0。",
            "是一个位置参数。",
            "是一个尺度参数，必须大于0。"
        ],
        "default_params": [-2.0, 2.0, 0.0, 10.0],
        "is_discrete": False,
        "validate_params": lambda params: (
            len(params) >= 4 and float(params[1]) > 0 and float(params[3]) > 0
        ),
        "support": lambda params: (float("-inf"), float("inf"))
    },
    "DriskKumaraswamy": {
        "type": "kumaraswamy", "min_params": 4, "max_params": 4,
        "param_names": ["α1", "α2", "Min", "Max"],
        "supports_shift_truncate": True,
        "category": "\u6709\u754c",
        "description": "定义一个4参数Kumaraswamy分布。",
        "ui_name": "Kumaraswamy - \u8fde\u7eed\u3001\u6709\u754c\u5206\u5e03",
        "ui_param_labels": ["α1", "α2", "Min", "Max"],
        "param_descriptions": [
            "是一个形状参数，必须大于 0。",
            "是一个形状参数，必须大于 0。",
            "是分布的最小值，必须小于Maximum（最大值）。",
            "是分布的最大值，必须大于Minimum（最小值）。"
        ],
        "default_params": [2.0, 2.0, -10.0, 10.0],
        "is_discrete": False,
        "validate_params": lambda params: (
            len(params) >= 4
            and float(params[0]) > 0
            and float(params[1]) > 0
            and float(params[3]) > float(params[2])
        ),
        "support": lambda params: (float(params[2]), float(params[3]))
    },
    "DriskLaplace": {
        "type": "laplace", "min_params": 2, "max_params": 2,
        "param_names": ["μ", "σ"],
        "supports_shift_truncate": True,
        "category": "\u65e0\u754c",
        "description": "定义一个Laplace分布。",
        "ui_name": "Laplace - \u8fde\u7eed\u3001\u65e0\u754c\u5206\u5e03",
        "ui_param_labels": ["μ", "σ"],
        "param_descriptions": [
            "是分布的均值。",
            "是分布的标准差，必须大于或等于0。"
        ],
        "default_params": [0.0, 10.0],
        "is_discrete": False,
        "validate_params": lambda params: len(params) >= 2 and float(params[1]) > 0,
        "support": lambda params: (float("-inf"), float("inf"))
    },
    "DriskLogistic": {
        "type": "logistic", "min_params": 2, "max_params": 2,
        "param_names": ["α", "β"],
        "supports_shift_truncate": True,
        "category": "\u65e0\u754c",
        "description": "定义一个Logistic分布。",
        "ui_name": "Logistic - \u8fde\u7eed\u3001\u65e0\u754c\u5206\u5e03",
        "ui_param_labels": ["α", "β"],
        "param_descriptions": [
            "是一个位置参数。",
            "是一个尺度参数，必须大于0。"
        ],
        "default_params": [0.0, 10.0],
        "is_discrete": False,
        "validate_params": lambda params: len(params) >= 2 and float(params[1]) > 0,
        "support": lambda params: (float("-inf"), float("inf"))
    },
    "DriskLoglogistic": {
        "type": "loglogistic", "min_params": 3, "max_params": 3,
        "param_names": ["γ", "β", "α"],
        "supports_shift_truncate": True,
        "category": "\u5de6\u7aef\u6709\u754c",
        "description": "定义一个Loglogistic分布。",
        "ui_name": "Loglogistic - \u8fde\u7eed\u3001\u5de6\u7aef\u6709\u754c\u5206\u5e03",
        "ui_param_labels": ["γ", "β", "α"],
        "param_descriptions": [
            "是一个位置参数。",
            "是一个尺度参数，必须大于 0。",
            "是一个形状参数，必须大于0。"
        ],
        "default_params": [100.0, 10.0, 50.0],
        "is_discrete": False,
        "validate_params": lambda params: (
            len(params) >= 3 and float(params[1]) > 0 and float(params[2]) > 0
        ),
        "support": lambda params: (float(params[0]), float("inf"))
    },
    "DriskLognorm": {
        "type": "lognorm", "min_params": 2, "max_params": 2,
        "param_names": ["μ", "σ"],
        "supports_shift_truncate": True,
        "category": "\u5de6\u7aef\u6709\u754c",
        "description": "定义一种对数正态分布，这种形式的对数正态分布的参数为该概率分布的实际均值和标准差。",
        "ui_name": "Lognorm - \u8fde\u7eed\u3001\u5de6\u7aef\u6709\u754c\u5206\u5e03",
        "ui_param_labels": ["μ", "σ"],
        "param_descriptions": [
            "是分布的均值，必须大于0。",
            "是分布的标准差，必须大于0。"
        ],
        "default_params": [10.0, 10.0],
        "is_discrete": False,
        "validate_params": lambda params: (
            len(params) >= 2 and float(params[0]) > 0 and float(params[1]) > 0
        ),
        "support": lambda params: (0.0, float("inf"))
    },
    "DriskLognorm2": {
        "type": "lognorm2", "min_params": 2, "max_params": 2,
        "param_names": ["μ", "σ"],
        "supports_shift_truncate": True,
        "category": "\u5de6\u7aef\u6709\u754c",
        "description": "指定一种对数正态分布，其中输入的均值和标准差等于相应正态分布的均值和标准差。",
        "ui_name": "Lognorm2 - \u8fde\u7eed\u3001\u5de6\u7aef\u6709\u754c\u5206\u5e03",
        "ui_param_labels": ["μ", "σ"],
        "param_descriptions": [
            "是底层正态分布的均值。",
            "是底层正态分布的标准差，必须大于0。"
        ],
        "default_params": [1.263, 1.0],
        "is_discrete": False,
        "validate_params": lambda params: (
            len(params) >= 2 and float(params[1]) > 0
        ),
        "support": lambda params: (0.0, float("inf"))
    },
    "DriskBetaGeneral": {
        "type": "betageneral", "min_params": 4, "max_params": 4,
        "param_names": ["α1", "α2", "Min", "Max"],
        "supports_shift_truncate": True,
        "category": "\u6709\u754c",
        "description": "\u5b9a\u4e49\u4e00\u4e2a\u5177\u6709\u81ea\u5b9a\u4e49\u6700\u5c0f\u503c\u548c\u6700\u5927\u503c\u7684Beta\u5206\u5e03\u3002",
        "ui_name": "BetaGeneral - \u8fde\u7eed\u3001\u6709\u754c\u5206\u5e03",
        "ui_param_labels": ["α1", "α2", "Min", "Max"],
        "param_descriptions": [
            "\u662f\u4e00\u4e2a\u5f62\u72b6\u53c2\u6570\uff0c\u5fc5\u987b\u5927\u4e8e 0\u3002",
            "Alpha_2 \u662f\u4e00\u4e2a\u5f62\u72b6\u53c2\u6570\uff0c\u5fc5\u987b\u5927\u4e8e 0\u3002",
            "\u662f\u5206\u5e03\u7684\u6700\u5c0f\u503c\uff0c\u5fc5\u987b\u5c0f\u4e8eMaximum\uff08\u6700\u5927\u503c\uff09\u3002",
            "\u662f\u5206\u5e03\u7684\u6700\u5927\u503c\uff0c\u5fc5\u987b\u5927\u4e8eMinimum\uff08\u6700\u5c0f\u503c\uff09\u3002"
        ],  
        "default_params": [2.0, 2.0, -10.0, 10.0],
        "is_discrete": False,
        "validate_params": lambda params: (
            len(params) >= 4
            and float(params[0]) > 0
            and float(params[1]) > 0
            and float(params[3]) > float(params[2])
        ),
        "support": lambda params: (float(params[2]), float(params[3]))
    },
    "DriskBetaSubj": {
        "type": "betasubj", "min_params": 4, "max_params": 4,
        "param_names": ["Min", "M.likely", "Mean", "Max"],
        "supports_shift_truncate": True,
        "category": "\u6709\u754c",
        "description": "定义一个具有明确最小值和最大值的Beta分布，其形状参数通过设定的最可能值和均值计算得出。",
        "ui_name": "BetaSubj - \u8fde\u7eed\u3001\u6709\u754c\u5206\u5e03",
        "ui_param_labels": ["Min", "M.likely", "Mean", "Max"],
        "param_descriptions": [
            "是分布的最小值，必须小于Maximum（最大值）。",
            "是分布的最可能值。",
            "是分布的均值。",
            "是分布的最大值，必须大于Minimum（最小值）。"
        ],
        "default_params": [-10.0, 6.0, 2.0, 10.0],
        "is_discrete": False,
        "validate_params": lambda params: (
            len(params) >= 4
            and float(params[3]) > float(params[0])
            and float(params[1]) > float(params[0])
            and float(params[1]) < float(params[3])
            and float(params[2]) > float(params[0])
            and float(params[2]) < float(params[3])
        ),
        "support": lambda params: (float(params[0]), float(params[3]))
    },
    "DriskBurr12": {
        "type": "burr12", "min_params": 4, "max_params": 4,
        "param_names": ["γ", "β", "α1", "α2"],
        "supports_shift_truncate": True,
        "category": "\u5de6\u7aef\u6709\u754c",
        "description": "定义一个Burr 12分布。",
        "ui_name": "Burr12 - \u8fde\u7eed\u3001\u5de6\u7aef\u6709\u754c\u5206\u5e03",
        "ui_param_labels": ["γ", "β", "α1", "α2"],
        "param_descriptions": [
            "是一个位置参数。",
            "是⼀个尺度参数，必须大于 0。",
            "是⼀个形状参数，必须大于 0。",
            "是⼀个形状参数，必须大于 0。"
        ],
        "default_params": [0.0, 1.0, 2.0, 2.0],
        "is_discrete": False,
        "validate_params": lambda params: (
            len(params) >= 4
            and float(params[1]) > 0
            and float(params[2]) > 0
            and float(params[3]) > 0
        ),
        "support": lambda params: (float(params[0]), float("inf"))
    },
    "DriskCompound": {
        "type": "compound", "min_params": 2, "max_params": 4,
        "param_names": ["Frequency", "Severity", "Deductible", "Limit"],
        "supports_shift_truncate": True,
        "category": "\u65e0\u754c",
        "description": "生成基于Severity分布的Frequency样本。",
        "ui_name": "Compound - \u8fde\u7eed\u3001\u65e0\u754c\u5206\u5e03",
        "ui_param_labels": ["Frequency", "Severity", "Deductible", "Limit"],
        "param_descriptions": [
            "定义从严重程度分布中抽取并累加的样本数量分布。",
            "定义描述各样本严重程度的分布。",
            "定义免赔额的可选参数，默认值为不设置免赔额。",
            "定义赔偿限额的可选参数，默认值为无上限。"
        ],
        "default_params": ["DriskPoisson(10)", "DriskLognorm(1000,100)", 0.0, float("inf")],
        "is_discrete": False,
        "validate_params": lambda params: (
            len(params) >= 2
            and _is_distribution_formula_string(params[0])
            and _is_distribution_formula_string(params[1])
            and (len(params) < 3 or float(params[2]) >= 0.0)
            and (len(params) < 4 or float(params[3]) >= 0.0)
        ),
        "support": lambda params: (0.0, float("inf"))
    },
    "DriskSplice": {
        "type": "splice", "min_params": 3, "max_params": 3,
        "param_names": ["Left", "Right", "SplicePoint"],
        "supports_shift_truncate": True,
        "category": "\u62fc\u63a5",
        "description": "在 x 等于Splice点处将分布 #1 和分布 #2 拼接在一起。",
        "ui_name": "Splice - \u62fc\u63a5\u5206\u5e03",
        "ui_param_labels": ["Left", "Right", "SplicePoint"],
        "param_descriptions": [
            "是一个分布，或指向包含该分布的单元格的引用。",
            "是一个分布，或指向包含该分布的单元格的引用。",
            "是两个分布拼接在一起的数值。"
        ],
        "default_params": ["DriskNormal(0,1)", "DriskNormal(2,1)", 1.0],
        "is_discrete": False,
        "validate_params": lambda params: (
            len(params) >= 3
            and _is_distribution_formula_string(params[0])
            and _is_distribution_formula_string(params[1])
            and math.isfinite(float(params[2]))
        ),
        "support": lambda params: (float("-inf"), float("inf"))
    },
    "DriskPert": {
        "type": "pert", "min_params": 3, "max_params": 3,
        "param_names": ["Min", "M.likely", "Max"],
        "supports_shift_truncate": True,
        "category": "\u6709\u754c",
        "description": "定义一个Pert分布（也是Beta分布的一种特殊形式）。",
        "ui_name": "Pert - \u8fde\u7eed\u3001\u6709\u754c\u5206\u5e03",
        "ui_param_labels": ["Min", "M.likely", "Max"],
        "param_descriptions": [
            "是分布的最小值，必须小于Max（最大值）。",
            "是分布的最可能值，常作为形态参数的计算基准。",
            "是分布的最大值，必须大于Min（最小值）。"
        ],
        "default_params": [-10.0, 0.0, 10.0],
        "is_discrete": False,
        "validate_params": lambda params: (
            len(params) >= 3
            and float(params[2]) > float(params[0])
            and float(params[1]) >= float(params[0])
            and float(params[1]) <= float(params[2])
        ),
        "support": lambda params: (float(params[0]), float(params[2]))
    },
    "DriskReciprocal": {
        "type": "reciprocal", "min_params": 2, "max_params": 2,
        "param_names": ["Min", "Max"],
        "supports_shift_truncate": True,
        "category": "\u6709\u754c",
        "description": "定义一个Reciprocal分布。",
        "ui_name": "Reciprocal - \u8fde\u7eed\u3001\u6709\u754c\u5206\u5e03",
        "ui_param_labels": ["Min", "Max"],
        "param_descriptions": [
            "是分布的最小值，必须小于或等于Maximum（最大值）。",
            "是分布的最大值，必须大于或等于Minimum（最小值）。"
        ],
        "default_params": [1.0, 2.0],
        "is_discrete": False,
        "validate_params": lambda params: (
            len(params) >= 2 and float(params[0]) > 0 and float(params[1]) > float(params[0])
        ),
        "support": lambda params: (float(params[0]), float(params[1]))
    },
    "DriskRayleigh": {
        "type": "rayleigh", "min_params": 1, "max_params": 1,
        "param_names": ["b"],
        "supports_shift_truncate": True,
        "category": "\u5de6\u7aef\u6709\u754c",
        "description": "定义一个Rayleigh分布。",
        "ui_name": "Rayleigh - \u8fde\u7eed\u3001\u5de6\u7aef\u6709\u754c\u5206\u5e03",
        "ui_param_labels": ["b"],
        "param_descriptions": [
            "是分布的众数，必须大于0。"
        ],
        "default_params": [10.0],
        "is_discrete": False,
        "validate_params": lambda params: len(params) >= 1 and float(params[0]) > 0,
        "support": lambda params: (0.0, float("inf"))
    },
    "DriskWeibull": {
        "type": "weibull", "min_params": 2, "max_params": 2,
        "param_names": ["α", "β"],
        "supports_shift_truncate": True,
        "category": "\u5de6\u7aef\u6709\u754c",
        "description": "定义Weibull分布，这是一种连续型概率分布，其形状与尺度特性会随参数取值发生显著变化。",
        "ui_name": "Weibull - \u8fde\u7eed\u3001\u5de6\u7aef\u6709\u754c\u5206\u5e03",
        "ui_param_labels": ["α", "β"],
        "param_descriptions": [
            "是一个形状参数，必须大于 0。",
            "是一个尺度参数，必须大于0。"
        ],
        "default_params": [2.0, 10.0],
        "is_discrete": False,
        "validate_params": lambda params: (
            len(params) >= 2 and float(params[0]) > 0 and float(params[1]) > 0
        ),
        "support": lambda params: (0.0, float("inf"))
    },
    "DriskPearson5": {
        "type": "pearson5", "min_params": 2, "max_params": 2,
        "param_names": ["α", "β"],
        "supports_shift_truncate": True,
        "category": "\u5de6\u7aef\u6709\u754c",
        "description": "定义一个Pearson5分布",
        "ui_name": "Pearson5 - \u8fde\u7eed\u3001\u5de6\u7aef\u6709\u754c\u5206\u5e03",
        "ui_param_labels": ["α", "β"],
        "param_descriptions": [
            "是一个形状参数，必须大于 0。",
            "是一个尺度参数，必须大于 0。"
        ],
        "default_params": [3.0, 10.0],
        "is_discrete": False,
        "validate_params": lambda params: (
            len(params) >= 2 and float(params[0]) > 0 and float(params[1]) > 0
        ),
        "support": lambda params: (0.0, float("inf"))
    },
    "DriskPearson6": {
        "type": "pearson6", "min_params": 3, "max_params": 3,
        "param_names": ["α1", "α2", "β"],
        "supports_shift_truncate": True,
        "category": "\u5de6\u7aef\u6709\u754c",
        "description": "定义一个Pearson6分布。",
        "ui_name": "Pearson6 - \u8fde\u7eed\u3001\u5de6\u7aef\u6709\u754c\u5206\u5e03",
        "ui_param_labels": ["α1", "α2", "β"],
        "param_descriptions": [
            "是一个形状参数，必须大于 0。",
            "是一个形状参数，必须大于 0。",
            "是一个尺度参数，必须大于 0。"
        ],
        "default_params": [2.0, 5.0, 10.0],
        "is_discrete": False,
        "validate_params": lambda params: (
            len(params) >= 3
            and float(params[0]) > 0
            and float(params[1]) > 0
            and float(params[2]) > 0
        ),
        "support": lambda params: (0.0, float("inf"))
    },
    "DriskPareto2": {
        "type": "pareto2", "min_params": 2, "max_params": 2,
        "param_names": ["b", "q"],
        "supports_shift_truncate": True,
        "category": "\u5de6\u7aef\u6709\u754c",
        "description": "定义一个Pareto2分布。",
        "ui_name": "Pareto2 - \u8fde\u7eed\u3001\u5de6\u7aef\u6709\u754c\u5206\u5e03",
        "ui_param_labels": ["b", "q"],
        "param_descriptions": [
            "是一个尺度参数，必须大于0。",
            "是一个形状参数，必须大于0。"
        ],
        "default_params": [1.0, 10.0],
        "is_discrete": False,
        "validate_params": lambda params: (
            len(params) >= 2 and float(params[0]) > 0 and float(params[1]) > 0
        ),
        "support": lambda params: (0.0, float("inf"))
    },
    "DriskPareto": {
        "type": "pareto", "min_params": 2, "max_params": 2,
        "param_names": ["θ", "α"],
        "supports_shift_truncate": True,
        "category": "\u5de6\u7aef\u6709\u754c",
        "description": "定义一个Pareto分布。",
        "ui_name": "Pareto - \u8fde\u7eed\u3001\u5de6\u7aef\u6709\u754c\u5206\u5e03",
        "ui_param_labels": ["θ", "α"],
        "param_descriptions": [
            "是一个形状参数，必须大于 0。",
            "是一个尺度参数，必须大于0。"
        ],
        "default_params": [10.0, 1.0],
        "is_discrete": False,
        "validate_params": lambda params: (
            len(params) >= 2 and float(params[0]) > 0 and float(params[1]) > 0
        ),
        "support": lambda params: (float(params[1]), float("inf"))
    },
    "DriskLevy": {
        "type": "levy", "min_params": 2, "max_params": 2,
        "param_names": ["a", "c"],
        "supports_shift_truncate": True,
        "category": "\u5de6\u7aef\u6709\u754c",
        "description": "定义一个Levy分布。",
        "ui_name": "Levy - \u8fde\u7eed\u3001\u5de6\u7aef\u6709\u754c\u5206\u5e03",
        "ui_param_labels": ["a", "c"],
        "param_descriptions": [
            "是一个位置参数。",
            "是一个尺度参数，必须大于0。"
        ],
        "default_params": [0.0, 1.0],
        "is_discrete": False,
        "validate_params": lambda params: len(params) >= 2 and float(params[1]) > 0,
        "support": lambda params: (float(params[0]), float("inf"))
    },
    "DriskCauchy": {
        "type": "cauchy", "min_params": 2, "max_params": 2,
        "param_names": ["γ", "β"],
        "supports_shift_truncate": True,
        "category": "\u65e0\u754c",
        "description": "定义一个Cauchy分布。",
        "ui_name": "Cauchy - \u8fde\u7eed\u3001\u65e0\u754c\u5206\u5e03",
        "ui_param_labels": ["γ", "β"],
        "param_descriptions": [
            "是一个位置参数。",
            "是一个尺度参数，必须大于 0。"
        ],
        "default_params": [0.0, 10.0],
        "is_discrete": False,
        "validate_params": lambda params: len(params) >= 2 and params[1] > 0,
        "support": lambda params: (float('-inf'), float('inf'))
    },
    "DriskDagum": {
        "type": "dagum", "min_params": 4, "max_params": 4,
        "param_names": ["γ", "β", "α1", "α2"],
        "supports_shift_truncate": True,
        "category": "\u5de6\u7aef\u6709\u754c",
        "description": "定义一个Dagum分布。",
        "ui_name": "Dagum - \u8fde\u7eed\u3001\u5de6\u7aef\u6709\u754c\u5206\u5e03",
        "ui_param_labels": ["γ", "β", "α1", "α2"],
        "param_descriptions": [
            "是一个位置参数。",
            "是一个尺度参数，必须大于0。",
            "是一个形状参数，必须大于0。",
            "是一个形状参数，必须大于0。"
        ],
        "default_params": [100.0, 10000.0, 1.0, 1.0],
        "is_discrete": False,
        "validate_params": lambda params: (
            len(params) >= 4
            and params[1] > 0
            and params[2] > 0
            and params[3] > 0
        ),
        "support": lambda params: (float(params[0]), float('inf'))
    },
    "DriskDoubleTriang": {
        "type": "doubletriang", "min_params": 4, "max_params": 4,
        "param_names": ["Min", "M.likely", "Max", "Lower P"],
        "supports_shift_truncate": True,
        "category": "\u6709\u754c",
        "description": "定义一个由三个点及下三角分布概率权重构成的Double Triangular分布，其中最小值与最大值出现的概率为零。",
        "ui_name": "DoubleTriang - \u8fde\u7eed\u3001\u6709\u754c\u5206\u5e03",
        "ui_param_labels": ["Min", "M.likely", "Max", "Lower P"],
        "param_descriptions": [
            "是分布的最小值，必须小于Max（最大值）。",
            "是分布的最可能值。",
            "是分布的最大值，必须大于Min（最小值）。",
            "是下三角分布的概率权重。"
        ],
        "default_params": [-10.0, 0.0, 10.0, 0.4],
        "is_discrete": False,
        "validate_params": lambda params: (
            len(params) >= 4
            and params[0] < params[1] < params[2]
            and 0 <= params[3] <= 1
        ),
        "support": lambda params: (float(params[0]), float(params[2]))
    },
    "DriskPoisson": {
        "type": "poisson", "min_params": 1, "max_params": 1,
        "param_names": ["λ"],
        "supports_shift_truncate": True, "description": "泊松分布",
        "ui_name": "泊松分布 (Poisson)",
        "ui_param_labels": ["λ"],
        "default_params": [5.0],
        "is_discrete": True,
        "description": "定义一个泊松分布，该离散型分布仅返回大于或等于零的整数值。",
        "param_descriptions": [
            "是分布的均值，必须大于0。"
        ],
        "validate_params": lambda params: params[0] > 0,
        "support": lambda params: (0.0, float('inf'))
    },
    "DriskGamma": {
        "type": "gamma", "min_params": 2, "max_params": 2,
        "param_names": ["α", "β"],
        "supports_shift_truncate": True, "description": "定义一个Gamma分布。Gamma分布是一个连续分布。",
        "ui_name": "伽马分布 (Gamma)",
        "ui_param_labels": ["α", "β"],
        "param_descriptions": [
            "是一个形状参数，必须大于 0。",
            "是一个尺度参数，必须大于 0。"
        ],
        "default_params": [2.0, 1.0],
        "is_discrete": False,
        "validate_params": lambda params: params[0] > 0 and params[1] > 0,
        "support": lambda params: (0.0, float('inf'))
    },
    "DriskErlang": {
        "type": "erlang", "min_params": 2, "max_params": 2,
        "param_names": ["m", "β"],
        "supports_shift_truncate": True,
        "category": "\u5de6\u7aef\u6709\u754c",
        "description": "定义一个具有指定m和beta参数的m-erlang分布。",
        "ui_name": "Erlang - \u8fde\u7eed\u3001\u5de6\u7aef\u6709\u754c\u5206\u5e03",
        "ui_param_labels": ["m", "Beta"],
        "param_descriptions": [
            "Gamma分布的整数型参数，必须为正整数。",
            "是一个尺度参数，必须大于 0。"
        ],
        "default_params": [2.0, 10.0],
        "is_discrete": False,
        "validate_params": lambda params: (
            len(params) >= 2
            and abs(float(params[0]) - round(float(params[0]))) <= 1e-9
            and float(params[0]) > 0
            and float(params[1]) > 0
        ),
        "support": lambda params: (0.0, float('inf'))
    },
    "DriskBeta": {
        "type": "beta", "min_params": 2, "max_params": 2,
        "param_names": ["α1", "α2"],
        "supports_shift_truncate": True,
        "description": "使用形状参数 alpha1 和 alpha2 来定义一个Beta分布。",
        "ui_name": "贝塔分布 (Beta)",
        "ui_param_labels": ["α1", "α2"],
        "param_descriptions": [
            "是一个形状参数，必须大于 0。",
            "是一个形状参数，必须大于 0。"
        ],
        "default_params": [2.0, 2.0],
        "is_discrete": False,
        "validate_params": lambda params: params[0] > 0 and params[1] > 0,
        "support": lambda params: (0.0, 1.0)
    },
    "DriskChiSq": {
        "type": "chisq", "min_params": 1, "max_params": 1,
        "param_names": ["V"],
        "supports_shift_truncate": True,
        "description": "定义一个自由度为V的卡方分布。",
        "ui_name": "卡方分布 (Chi-Square)",
        "ui_param_labels": ["V"],
        "param_descriptions": [
            "表示该分布的自由度，必须为正整数。"
        ],
        "default_params": [5.0],
        "is_discrete": False,
        "validate_params": lambda params: params[0] > 0 and params[0].is_integer(),
        "support": lambda params: (0.0, float('inf'))
    },
    "DriskF": {
        "type": "f", "min_params": 2, "max_params": 2,
        "param_names": ["V1", "V2"],
        "supports_shift_truncate": True, 
        "description": "定义一个F分布。",
        "ui_name": "F 分布 (F)",
        "ui_param_labels": ["V1", "V2"],
        "param_descriptions": [
            "是第一个自由度数值，必须是一个正整数。",
            "是第二个自由度数值，必须是一个正整数。"
        ],
        "default_params": [5.0, 10.0],
        "is_discrete": False,
        "validate_params": lambda params: params[0] > 0 and params[1] > 0 and params[0].is_integer() and params[1].is_integer(),
        "support": lambda params: (0.0, float('inf'))
    },
    "DriskStudent": {
        "type": "student", "min_params": 1, "max_params": 1,
        "param_names": ["v"],
        "supports_shift_truncate": True, 
        "description":"定义一个Student分布。",
        "ui_name": "T 分布 (Student-t)",
        "ui_param_labels": ["v"],
        "param_descriptions": ["是分布的自由度，必须是正整数。"],
        "default_params": [10.0],
        "is_discrete": False,
        "validate_params": lambda params: params[0] > 0 and params[0].is_integer(),
        "support": lambda params: (float('-inf'), float('inf'))
    },
    "DriskExpon": {
        "type": "expon", "min_params": 1, "max_params": 1,
        "param_names": ["β"],
        "supports_shift_truncate": True,
        "description": "定义一个具有指定beta值的指数分布。",
        "ui_name": "指数分布 (Exponential)",
        "ui_param_labels": ["β"],
        "param_descriptions": [
            "是分布的均值，必须大于 0。"
        ],
        "default_params": [1.0],
        "is_discrete": False,
        "validate_params": lambda params: params[0] > 0,
        "support": lambda params: (0.0, float('inf'))
    },
    "DriskInvgauss": {
        "type": "invgauss", "min_params": 2, "max_params": 2,
        "param_names": ["μ", "λ"],
        "supports_shift_truncate": True,
        "description": "\u5b9a\u4e49\u4e00\u4e2aInvgauss\u5206\u5e03\u3002",
        "ui_name": "Inverse Gaussian (InvGauss)",
        "ui_param_labels": ["μ", "λ"],
        "param_descriptions": [
            "\u662f\u5206\u5e03\u7684\u5747\u503c\uff0c\u5fc5\u987b\u5927\u4e8e0\u3002",
            "\u662f\u4e00\u4e2a\u5f62\u72b6\u53c2\u6570\uff0c\u5fc5\u987b\u5927\u4e8e 0\u3002"
        ],
        "default_params": [1.0, 10.0],
        "is_discrete": False,
        "validate_params": lambda params: params[0] > 0 and params[1] > 0,
        "support": lambda params: (0.0, float('inf'))
    },
    "DriskGeomet": {
        "type": "geomet", "min_params": 1, "max_params": 1,
        "param_names": ["p"],
        "supports_shift_truncate": True,
        "category": "\u5de6\u7aef\u6709\u754c",
        "description": "定义一个几何分布，其返回值表示在一系列独立试验中首次成功前所经历的失败次数。",
        "ui_name": "Geomet - \u79bb\u6563\u3001\u5de6\u7aef\u6709\u754c\u5206\u5e03",
        "ui_param_labels": ["p"],
        "param_descriptions": ["每次试验的成功概率，其取值范围必须大于等于0且小于等于1。"],
        "default_params": [0.5],
        "is_discrete": True,
        "validate_params": lambda params: params[0] > 0 and params[0] <= 1,
        "support": lambda params: (0.0, float('inf'))
    },
    "DriskHypergeo": {
        "type": "hypergeo", "min_params": 3, "max_params": 3,
        "param_names": ["n", "D", "M"],
        "supports_shift_truncate": True,
        "category": "\u6709\u754c",
        "description": "定义一个超几何分布，该离散型分布仅返回非负整数值。",
        "ui_name": "Hypergeo - \u79bb\u6563\u3001\u6709\u754c\u5206\u5e03",
        "ui_param_labels": [
           "n", "D", "M"
        ],
        "param_descriptions": [
            "样本容量，必须为正整数值。",
            "表示相关类型的物品数量，必须为正整数。",
            "总体规模，必须为正整数值。"
        ],
        "default_params": [10.0, 20.0, 30.0],
        "is_discrete": True,
        "validate_params": lambda params: (
            len(params) >= 3
            and params[0] > 0 and float(params[0]).is_integer()
            and params[1] > 0 and float(params[1]).is_integer()
            and params[2] > 0 and float(params[2]).is_integer()
            and params[0] <= params[2]
            and params[1] <= params[2]
        ),
        "support": lambda params: (
            float(max(0, int(float(params[0])) + int(float(params[1])) - int(float(params[2])))),
            float(min(int(float(params[0])), int(float(params[1]))))
        )
    },
    "DriskIntuniform": {
        "type": "intuniform", "min_params": 2, "max_params": 2,
        "param_names": ["Min", "Max"],
        "supports_shift_truncate": True,
        "category": "\u6709\u754c",
        "description": "定义一个返回最小值和最大值范围内整数的等概率分布。",
        "ui_name": "Intuniform - \u79bb\u6563\u3001\u6709\u754c\u5206\u5e03",
        "ui_param_labels": ["Min", "Max"],
        "param_descriptions": [
            "是分布的最小值，必须小于或等于Maximum（最大值）。",
            "是分布的最大值，必须大于或等于Minimum（最小值）。"
        ],
        "default_params": [0.0, 10.0],
        "is_discrete": True,
        "validate_params": lambda params: (
            len(params) >= 2
            and abs(float(params[0]) - round(float(params[0]))) <= 1e-9
            and abs(float(params[1]) - round(float(params[1]))) <= 1e-9
            and params[0] < params[1]
        ),
        "support": lambda params: (float(int(float(params[0]))), float(int(float(params[1]))))
    },
    "DriskNegbin": {
        "type": "negbin", "min_params": 2, "max_params": 2,
        "param_names": ["s", "p"],
        "supports_shift_truncate": True,
        "category": "\u5de6\u7aef\u6709\u754c",
        "description": "定义一个负二项分布，该离散型分布仅返回大于或等于零的整数值",
        "ui_name": "Negbin - \u79bb\u6563\u3001\u5de6\u7aef\u6709\u754c\u5206\u5e03",
        "ui_param_labels": [
            "s", "p"
        ],
        "param_descriptions": [
            "成功次数，必须为正整数。",
            "每次试验的成功概率，其取值范围必须大于等于0且小于等于1。"
        ],
        "default_params": [1.0, 0.5],
        "is_discrete": True,
        "validate_params": lambda params: (
            len(params) >= 2
            and abs(float(params[0]) - round(float(params[0]))) <= 1e-9
            and params[0] > 0
            and params[1] > 0
            and params[1] <= 1
        ),
        "support": lambda params: (
            0.0,
            0.0 if float(params[1]) >= 1.0 else float('inf')
        )
    },
    "DriskBernoulli": {
        "type": "bernoulli", "min_params": 1, "max_params": 1,
        "param_names": ["p"],
        "supports_shift_truncate": True, 
        "description": "定义一个伯努利分布，其中每次试验的成功概率为p。该离散分布仅返回大于或等于零的整数值。",
        "param_descriptions": ["每次试验的成功概率，其取值范围必须大于等于0且小于等于1。"],
        "ui_name": "伯努利分布 (Bernoulli)",
        "ui_param_labels": ["p"],
        "default_params": [0.5],
        "is_discrete": True,
        "validate_params": lambda params: 0 < params[0] < 1,
        "support": lambda params: (0.0, 1.0)
    },
    "DriskTriang": {
        "type": "triang",
        "min_params": 3,
        "max_params": 3,
        "validate_params": lambda params: params[0] <= params[1] <= params[2],
        "param_names": ["Min", "M.likely", "Max"],
        "supports_shift_truncate": True,
        "description": "定义一个Triangular分布，其最小值与最大值处的发生概率为零。",
        "param_descriptions": [
            "是分布的最小值，必须小于Max（最大值）。",
            "是分布的最可能值。",
            "是分布的最大值，必须大于Min（最小值）。"
        ],
        "ui_name": "三角分布 (Triangular)",
        "ui_param_labels": ["Min", "M.likely", "Max"],
        "default_params": [0.0, 0.5, 1.0],
        "is_discrete": False,
        "support": lambda params: (params[0], params[2])
    },
    "DriskBinomial": {
        "type": "binomial",
        "min_params": 2,
        "max_params": 2,
        "validate_params": lambda params: params[0] > 0 and params[0].is_integer() and 0 < params[1] < 1,
        "param_names": ["n", "p"],
        "supports_shift_truncate": True,
        "description": "定义一个二项分布，其中试验次数为N，每次试验的成功概率为P。该离散分布仅返回大于或等于零的整数值。",
        "param_descriptions  ":[
            "试验或抽样的次数",
            "每次试验的成功概率，其取值范围必须大于等于0且小于等于1。"
        ],
        "ui_name": "二项分布 (Binomial)",
        "ui_param_labels": ["n", "p"],
        "default_params": [10.0, 0.5],
        "is_discrete": True,
        "support": lambda params: (0.0, float(params[0]))
    },
    "DriskTrigen": {
        "type": "trigen",
        "min_params": 5,
        "max_params": 5,
        "validate_params": lambda params: params[0] <= params[1] <= params[2] and 0 <= params[3] < params[4] <= 1,
        "param_names": ["Bottom", "M.likely", "Top", "Bottom%", "Top%"],
        "supports_shift_truncate": True,
        "description": "定义一个三角分布，该分布具有三个点，其中 一个位于最可能值，另外两个位于指定的底部和顶部百分位数处。",
        "param_descriptions": [
            "是计算底部百分位数的临界值，表示在该值以下累积概率达到指定百分比。",
            "是分布的最可能值。",
            "是计算顶部百分位数的临界值，表示在该值以上累积概率达到指定百分比。",
            "指在分布中Bottom左侧区域面积占总面积的百分比，其取值必须介于0到100之间。",
            "指在分布中Top左侧区域面积占总面积的百分比，其取值必须介于0到100之间。"
        ],
        "ui_name": "三参数三角分布 (Trigen)",
        "ui_param_labels": ["Bottom", "M.likely", "Top", "Bottom%", "Top%"],
        "default_params": [0.0, 0.5, 1.0, 0.25, 0.75],
        "is_discrete": False,
        "support": lambda params: (float('-inf'), float('inf'))  # 实际支撑由转换后的三角决定，这里留空
    },
    "DriskCumul": {
        "type": "cumul",
        "min_params": 4,
        "max_params": 4,
        "validate_params": lambda params: True,  # 实际验证在类中
        "param_names": ["Min", "Max", "X-Table", "P-Table"],
        "supports_shift_truncate": True,
        "description": "定义一个由最小值和最大值确定范围、包含n个点的Cumul分布。",
        "ui_name": "累积分布 (Cumul)",
        "ui_param_labels": ["Min", "Max", "X-Table", "P-Table"],
        "param_descriptions": [
            "是分布的最小值，必须小于Maximum（最大值）。",
            "是分布的最大值，必须大于Minimum（最小值）。",
            "表示曲线上各点对应的数值，必须严格按数值升序排列指定。",
            "表示曲线上各点对应的累积概率，必须严格按概率升序排列指定，每个点的取值应>=0且<=1。"
        ],
        "default_params": [-1.0, 1.0, "-0.5,0,0.5", "0.1,0.5,0.9"],
        "is_discrete": False,
        "support": lambda params: (float('-inf'), float('inf'))  # 实际支撑由X数组决定
    },
    "DriskDUniform": {
        "type": "duniform",
        "min_params": 1,
        "max_params": 1,
        "validate_params": lambda params: True,
        "param_names": ["X-Table"],
        "supports_shift_truncate": True,
        "description": "定义一个离散均匀分布，该分布具有任意数量的可能结果，且每个结果发生的概率相等。",
        "param_descriptions": [
            "指定每个可能结果对应的输入值。"
        ],
        "ui_name": "Discrete Uniform (DUniform)",
        "ui_param_labels": ["X-Table"],
        "default_params": ["1,2,3"],
        "is_discrete": True,
        "support": lambda params: (float('-inf'), float('inf'))
    },
    "DriskDiscrete": {
        "type": "discrete",
        "min_params": 2,
        "max_params": 2,
        "validate_params": lambda params: True,
        "param_names": ["X-Table", "P-Table"],
        "supports_shift_truncate": True,
        "description": "定义一个具有指定结果数量的离散分布，可输入任意数量的可能结果。",
        "param_descriptions": [
            "指定每个可能结果对应的输入值。",
            "用于指定各结果出现概率的权重值。权重值p必须满足：所有元素>=0，且总和＞0。"
        ],
        "ui_name": "离散分布 (Discrete)",
        "ui_param_labels": ["X-Table", "P-Table"],
        "default_params": ["1,2,3", "0.2,0.3,0.5"],
        "is_discrete": True,
        "support": lambda params: (float('-inf'), float('inf'))
    },
}

# 分布函数名称列表（从注册表自动生成）
# 旧版面板文案覆盖机制已移除，使用内置分布注册表文案


DISTRIBUTION_FUNCTION_NAMES = list(DISTRIBUTION_REGISTRY.keys())

# 分布类型到函数名的反向映射
DIST_TYPE_TO_FUNC_NAME = {info["type"]: name for name, info in DISTRIBUTION_REGISTRY.items()}

# 辅助函数
def get_distribution_info(func_name):
    """获取分布函数信息"""
    return DISTRIBUTION_REGISTRY.get(func_name, None)

def get_distribution_type(func_name):
    """获取分布类型"""
    info = get_distribution_info(func_name)
    return info["type"] if info else "normal"

def get_all_distribution_names():
    """获取所有分布函数名称"""
    return DISTRIBUTION_FUNCTION_NAMES

def get_all_distribution_types():
    """获取所有分布类型"""
    return [info["type"] for info in DISTRIBUTION_REGISTRY.values()]

def validate_distribution_params(func_name, params):
    """验证分布参数"""
    info = get_distribution_info(func_name)
    if not info:
        return False
    if "validate_params" in info and info["validate_params"]:
        try:
            # 对于 Cumul 和 Discrete，参数是列表，验证函数可能不适用，因此跳过
            if func_name in ["DriskCumul", "DriskDiscrete", "DriskDUniform", "DriskHistogrm"]:
                return True
            return info["validate_params"](params)
        except:
            return False
    return True

def get_distribution_support(func_name, params):
    """获取分布的支撑集 (min, max)"""
    info = get_distribution_info(func_name)
    if not info:
        return (float('-inf'), float('inf'))
    support_func = info.get("support")
    if support_func is None:
        return (float('-inf'), float('inf'))
    try:
        return support_func(params)
    except:
        return (float('-inf'), float('inf'))

# 属性函数名称
ATTR_NAME = "DriskName"
ATTR_LOC = "DriskLoc"
ATTR_CATEGORY = "DriskCategory"
ATTR_COLLECT = "DriskCollect"
ATTR_CONVERGENCE = "DriskConvergence"
ATTR_COPULA = "DriskCopula"
ATTR_CORRMAT = "DriskCorrmat"
ATTR_FIT = "DriskFit"
ATTR_ISDATE = "DriskIsDate"
ATTR_ISDISCRETE = "DriskIsDiscrete"
ATTR_LOCK = "DriskLock"
ATTR_SEED = "DriskSeed"
ATTR_SHIFT = "DriskShift"
ATTR_STATIC = "DriskStatic"
ATTR_TRUNCATE = "DriskTruncate"
ATTR_TRUNCATEP = "DriskTruncateP"
ATTR_TRUNCATE2 = "DriskTruncate2"
ATTR_TRUNCATEP2 = "DriskTruncateP2"
ATTR_UNITS = "DriskUnits"

# 统计函数名称
STAT_MEAN = "DriskMean"
STAT_STD = "DriskStd"
STAT_MIN = "DriskMin"
STAT_MAX = "DriskMax"
STAT_MED = "DriskMed"
STAT_PTOX = "DriskPtoX"
STAT_DATA = "DriskData"

# 随机数生成器类型
RNG_MT19937 = 1          # Mersenne Twister = NumPy的MT19937
RNG_MRG32K3A = 2         # MRG32k3a = NumPy的MRG32k3a
RNG_PCG64 = 3            # PCG64 = NumPy的PCG64
RNG_PHILOX = 4           # Philox = NumPy的Philox
RNG_SFC64 = 5            # SFC64 = NumPy的SFC64
RNG_THREEFRY = 6         # Threefry = NumPy的Threefry
RNG_XOSHIRO256 = 7       # Xoshiro256++ = NumPy的Xoshiro256++
RNG_XOROSHIRO128 = 8     # Xoroshiro128++ = NumPy的Xoroshiro128++

# 默认随机数生成器设置
DEFAULT_RNG_TYPE = RNG_MT19937
DEFAULT_SEED = 42

# 随机数生成器名称映射
RNG_TYPE_NAMES = {
    RNG_MT19937: "Mersenne Twister (MT19937)",
    RNG_MRG32K3A: "MRG32k3a",
    RNG_PCG64: "PCG64",
    RNG_PHILOX: "Philox",
    RNG_SFC64: "SFC64",
    RNG_THREEFRY: "Threefry",
    RNG_XOSHIRO256: "Xoshiro256++",
    RNG_XOROSHIRO128: "Xoroshiro128++",
}
