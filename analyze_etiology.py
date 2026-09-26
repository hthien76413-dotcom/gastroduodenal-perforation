# -*- coding: utf-8 -*-
"""
课题① 儿童胃十二指肠穿孔病因谱十年变迁 —— 按《统计分析计划_SAP_v1》执行全部分析

输入：
  · 裁定合并_一致性与仲裁_v1.xlsx   merge_adjudication.py 的输出（终裁数据集、一致性κ、分歧与仲裁）
  · 胃十二指肠穿孔.xlsx              科研平台导出（出院科别、检验、出院诊断编码）
  · 病案数据调取_填报模板_v1.xlsx     病案室回填的分母（第1、2、5项）与编码核对（第6项）；未回填时相关分析自动跳过
  · 消化道异物_逐月住院人次_v1.xlsx   异物（T18）住院人次，异物穿孔转化率的分母
  · 次要终点提取表_v1.xlsx           临床表现、并发症、再次手术与进食时间；全部纳入病例核对完成且无逻辑错误时才纳入
输出：
  · 分析结果_病因谱_v1.xlsx          SAP 第五节所列各表及趋势分析汇总
  · 图2_逐年病因构成_v1.png          逐年病例数按病因堆叠；有分母时下方另列相对发生率（两幅上下对齐，不用双纵轴）

防误用：终裁数据集中仍有“待仲裁”或“裁定未完成”的病例时，脚本拒绝运行。预试或调试时可加
       --draft 参数，仅分析已定稿病例，结果文件与图上均标注“草稿”，不得引用。

安全约定（同本课题其他脚本）：输出文件若已存在先备份；写盘只在 __main__。
"""
import hashlib
import math
import os
import platform
import shutil
import sys
import warnings
from datetime import datetime

import numpy as np
import pandas as pd
import scipy
import statsmodels
import statsmodels.api as sm
from scipy import stats
from scipy.special import gammaln

import build_data_request_template as req
import build_extraction_sheet as ext
from merge_adjudication import norm

SRC = r'D:\胃十二指肠穿孔\胃十二指肠穿孔.xlsx'
MERGED = r'D:\胃十二指肠穿孔\裁定合并_一致性与仲裁_v1.xlsx'
DIR = r'D:\胃十二指肠穿孔\①儿童胃十二指肠穿孔病因谱十年变迁'
DENOM = DIR + r'\病案数据调取_填报模板_v1.xlsx'
FB = DIR + r'\消化道异物_逐月住院人次_v1.xlsx'
OUT = DIR + r'\分析结果_病因谱_v1.xlsx'
FIG = DIR + r'\图2_逐年病因构成_v1.png'
EXTRACT = r'D:\胃十二指肠穿孔\次要终点提取表_v1.xlsx'

KEY = '科研就诊编号'
SEED = 20260926
N_PERM = 100_000          # 主要分析的置换次数
N_PERM_TAB = 20_000       # 列联表 Fisher–Freeman–Halton 蒙特卡洛次数
T0 = 2021.0               # 连续入院时间的中心
STUDY = (pd.Period('2016-06', 'M'), pd.Period('2026-05', 'M'))
# 科研平台断档；若病案室确认数据完整且科研平台补导成功，改为 None 后重跑（SAP 3.4）
GAP = (pd.Period('2019-07', 'M'), pd.Period('2020-03', 'M'))
PERIODS = [('P1 疫情前', '2016-06', '2019-12'),
           ('P2 疫情期', '2020-01', '2022-12'),
           ('P3 疫情后', '2023-01', '2026-05')]
SA5_START = pd.Period('2020-04', 'M')     # 断档后连续月份起点
FB_START = pd.Period('2016-07', 'M')     # 异物导出起点

CAUSE4 = {'1': '异物相关', '2': '新生儿自发性', '3': '消化性溃疡', '4': '其他', '5': '其他', '6': '其他'}
CAUSE4_ORDER = ['异物相关', '新生儿自发性', '消化性溃疡', '其他']
CAUSE6_ORDER = ['1异物相关', '2新生儿自发性/胃壁肌层缺损', '3消化性溃疡', '4创伤', '5医源性', '6其他/特发性']
MNL_REF = '消化性溃疡'
AGE4_ORDER = ['新生儿', '婴幼儿', '学龄前', '学龄期']
EXCL_ORDER = ['E1部位非胃十二指肠', 'E2无活动性穿孔', 'E3外院术后转入', 'E4同次穿孔再住院', 'E5资料严重缺失']

# 图 2：分类色取 dataviz 参考调色板前 4 槽（相邻 CVD ΔE≥9.1，已用 validate_palette.js 校验）；
# 黄、青两色对浅底对比度＜3:1，故配图例并在末个整年柱旁直接标注类别
FIG_COLORS = {'异物相关': '#2a78d6', '新生儿自发性': '#eb6834', '消化性溃疡': '#1baf7a', '其他': '#eda100'}
FIG_EN = {'异物相关': 'Foreign body', '新生儿自发性': 'Neonatal spontaneous',
          '消化性溃疡': 'Peptic ulcer', '其他': 'Other'}
INK, INK2, GRID, SURFACE = '#0b0b0b', '#52514e', '#e4e3df', '#fcfcfb'


def sha256(path):
    h = hashlib.sha256()
    with open(path, 'rb') as f:
        for blk in iter(lambda: f.read(1 << 20), b''):
            h.update(blk)
    return h.hexdigest()


def in_gap(m):
    return GAP is not None and GAP[0] <= m <= GAP[1]


def study_months():
    """研究窗口内的有效月份（已扣除断档）。"""
    ms = pd.period_range(STUDY[0], STUDY[1], freq='M')
    return pd.PeriodIndex([m for m in ms if not in_gap(m)])


def period_of(m):
    for name, a, b in PERIODS:
        if pd.Period(a, 'M') <= m <= pd.Period(b, 'M'):
            return name
    return ''


def t_of(m):
    """月份中点的连续时间，以 T0 为中心。"""
    return m.year + (m.month - 0.5) / 12 - T0


def age4(days):
    if days <= 28:
        return '新生儿'
    if days < 365 * 3:
        return '婴幼儿'
    if days < 365 * 6:
        return '学龄前'
    return '学龄期'


# ================================================================ 读数
def load_cases(draft):
    fin = pd.read_excel(MERGED, sheet_name='终裁数据集', dtype=object, keep_default_na=False)
    dis = pd.read_excel(MERGED, sheet_name='分歧与仲裁', dtype=object, keep_default_na=False)
    for c in fin.columns:
        if c not in ('年龄_岁', '年龄_天', '住院天数'):
            fin[c] = fin[c].map(norm)
    status = fin['终裁状态'].value_counts().to_dict()
    unfinished = len(fin) - status.get('已定稿', 0)
    if unfinished and not draft:
        sys.exit('终裁数据集尚有 %d 例未定稿（%s）。完成仲裁后再运行；预试可加 --draft。' % (unfinished, status))
    all_ = fin[fin['终裁状态'] == '已定稿'].copy()

    disputed = set(dis.loc[dis['字段'].map(norm) == '病因大类', KEY].map(norm)) if not dis.empty else set()
    all_['病因经仲裁'] = all_[KEY].isin(disputed)
    all_['年龄_天'] = pd.to_numeric(all_['年龄_天'])
    all_['年龄_岁'] = pd.to_numeric(all_['年龄_岁'])
    all_['住院天数'] = pd.to_numeric(all_['住院天数'])
    all_['入院日期'] = pd.to_datetime(all_['入院日期'])
    all_['月'] = all_['入院日期'].dt.to_period('M')
    all_['年'] = all_['入院日期'].dt.year
    all_['t'] = all_['月'].map(t_of)
    all_['时期'] = all_['月'].map(period_of)
    all_['年龄组4'] = all_['年龄_天'].map(age4)

    inc = all_[all_['是否排除'] == '否'].copy()
    bad = [c for c in inc['病因大类'] if c[:1] not in CAUSE4]
    assert not bad, '纳入病例的病因大类取值异常：%s' % bad[:5]
    inc['病因4'] = inc['病因大类'].str[:1].map(CAUSE4)
    inc['异物'] = (inc['病因4'] == '异物相关').astype(int)
    inc['磁性'] = (inc['异物类型'] == '磁性').astype(int)
    out_win = inc[(inc['月'] < STUDY[0]) | (inc['月'] > STUDY[1])]
    assert out_win.empty, '有纳入病例在研究窗口外：%s' % list(out_win[KEY])
    gap_cases = inc[inc['月'].map(in_gap)]
    assert gap_cases.empty, ('断档期内出现纳入病例 %d 例，说明科研平台已补导；请将 GAP 改为 None 后重跑'
                             % len(gap_cases))
    return all_, inc, unfinished, status


def add_clinical(inc):
    """出院科别、入院 48 小时内首次白细胞与 CRP。"""
    x = pd.ExcelFile(SRC)
    head = x.parse('病案首页基本信息')
    head[KEY] = head[KEY].map(norm)
    adm = pd.to_datetime(head.set_index(KEY)['入院日期'])
    inc['出院科别'] = inc[KEY].map(head.set_index(KEY)['出院科别'])

    def first_within_48h(sheet, cols):
        lab = x.parse(sheet)[['就诊编号', '报告时间'] + cols].copy()   # 原表 150+ 列，只取所需
        lab['k'] = lab['就诊编号'].map(norm)
        lab['时间'] = pd.to_datetime(lab['报告时间'], errors='coerce')
        vals = None
        for c in cols:                               # 多个检测项目按顺序取第一个有值者
            v = pd.to_numeric(lab[c], errors='coerce')
            vals = v if vals is None else vals.fillna(v)
        lab['值'] = vals
        lab['入院'] = lab['k'].map(adm)
        ok = lab['值'].notna() & (lab['时间'] >= lab['入院']) & (lab['时间'] <= lab['入院'] + pd.Timedelta(hours=48))
        return lab[ok].sort_values('时间').groupby('k')['值'].first()

    inc['入院白细胞'] = inc[KEY].map(first_within_48h('实验室检查_血常规', ['白细胞计数定量-定量结果']))
    inc['入院CRP'] = inc[KEY].map(first_within_48h('实验室检查_血C反应蛋白（CRP）测定',
                                                   ['C反应蛋白定量-定量结果', '超敏C反应蛋白定量-定量结果']))
    return inc


def add_extraction(inc, draft):
    """术后住院天数与检索集内 30 天再入院由源数据自动计算；提取表变量仅在全部纳入病例核对完成、
    且提取表“逻辑核查”无错误时并入（草稿模式下只并入已核对病例）。返回 (inc, 说明)。"""
    auto = ext.auto_frame(SRC)
    inc['术后住院天数'] = inc[KEY].map(auto['术后住院天数'])
    inc['30天再入院'] = inc[KEY].map(auto['检索集内30天再入院']).fillna('').str.startswith('是')
    if not os.path.exists(EXTRACT):
        return inc, '提取表不存在，临床表现与并发症等变量未纳入'
    e = pd.read_excel(EXTRACT, sheet_name='提取表', dtype=object, keep_default_na=False)
    e[KEY] = e[KEY].map(norm)
    e = e.drop_duplicates(KEY).set_index(KEY)
    done = inc[KEY].map(lambda k: k in e.index and norm(e.at[k, '核对者']) != '')
    chk = pd.read_excel(EXTRACT, sheet_name='逻辑核查', dtype=object, keep_default_na=False)
    n_err = int(((chk['级别'] == '错误') & chk[KEY].map(norm).isin(set(inc[KEY]))).sum()) if not chk.empty else 0
    if (not done.all() or n_err) and not draft:
        return inc, ('提取表尚未就绪（已核对 %d/%d 例，逻辑错误 %d 条），临床表现与并发症等变量未纳入'
                     % (done.sum(), len(inc), n_err))
    for c in ext.MANUAL_COLS:
        inc[c] = [norm(e.at[k, c]) if ok else '' for k, ok in zip(inc[KEY], done)]
    inc['提取已核对'] = done.values
    return inc, '已并入提取表（已核对 %d/%d 例%s）' % (done.sum(), len(inc), '，草稿' if not done.all() or n_err else '')


def _known(v, unknown=('', '不详', '未查', 'NR', '不适用')):
    return ~v.isin(unknown)


def k_of_n(mask, known):
    k, n = int((mask & known).sum()), int(known.sum())
    return '%d/%d (%.1f)' % (k, n, 100 * k / n) if n else '—'


def _matrix(sheet, cols):
    """读病案室填报的年月矩阵表；表头在第 5 行（build_data_request_template 的版式）。"""
    try:
        d = pd.read_excel(DENOM, sheet_name=sheet, header=4)
    except (FileNotFoundError, ValueError):
        return None
    have = [c for c in cols if c in d.columns]
    if not have:
        return None
    num = d[have].apply(pd.to_numeric, errors='coerce')
    if num.notna().sum().sum() == 0:                 # 尚未回填
        return None
    num.index = pd.PeriodIndex([pd.Period(year=int(y), month=int(m), freq='M')
                                for y, m in zip(d['年'], d['月'])])
    return num


def load_denoms():
    """返回 (逐月分母 DataFrame, 可用分母说明, 第6项表)；未回填的分母列为空值。"""
    mon = pd.DataFrame(index=pd.period_range('2016-01', '2026-05', freq='M'))
    notes = {}
    m1 = _matrix('第1项_外科病区住院人次', req.WARDS)
    if m1 is not None:
        miss = m1[req.WARD_CORE].isna().sum().sum()
        notes['核心'] = '第1项核心病区' + ('（有 %d 个空格，已按缺失处理）' % miss if miss else '')
        mon['核心'] = m1[req.WARD_CORE].sum(axis=1, min_count=len(req.WARD_CORE))
        mon['核心+扩展'] = m1[req.WARDS].sum(axis=1, min_count=len(req.WARDS))
    m2 = _matrix('第2项_急诊入院住院人次', req.WARD_CORE)
    if m2 is not None:
        mon['急诊'] = m2[req.WARD_CORE].sum(axis=1, min_count=len(req.WARD_CORE))
    m5 = _matrix('第5项_全院住院总人次', ['全院住院人次'])
    if m5 is not None:
        mon['全院'] = m5['全院住院人次']
    m6 = _matrix('第6项_穿孔相关编码住院人次', req.CODE_GROUPS + ['合计（去重）'])
    for c in ('核心', '核心+扩展', '急诊', '全院'):
        if c not in mon:
            mon[c] = np.nan
    return mon, notes, m6


def load_fb():
    d = pd.read_excel(FB, sheet_name='逐月')
    d = d[d['数据状态'] == '有效']
    d.index = pd.PeriodIndex([pd.Period(year=int(y), month=int(m), freq='M') for y, m in zip(d['年'], d['月'])])
    return d['任一T18_含补回'].astype(float)


# ================================================================ 统计工具
def perm_rxc(x, y, n_perm=N_PERM_TAB, seed=SEED):
    """r×c 表 Fisher–Freeman–Halton 精确检验的蒙特卡洛 p 值（边际固定，按表概率排序）。"""
    x, y = pd.Series(x).reset_index(drop=True), pd.Series(y).reset_index(drop=True)
    ok = x.notna() & y.notna() & (x.astype(str) != '') & (y.astype(str) != '')
    xc = pd.factorize(x[ok])[0]
    yc = pd.factorize(y[ok])[0]
    r, c = xc.max() + 1, yc.max() + 1
    if r < 2 or c < 2:
        return np.nan

    def stat(yy):                                   # Σ log(n_ij!)，越大表概率越小
        return gammaln(np.bincount(xc * c + yy, minlength=r * c) + 1).sum()

    obs = stat(yc)
    rng = np.random.default_rng(seed)
    hits = sum(stat(rng.permutation(yc)) >= obs - 1e-9 for _ in range(n_perm))
    return (hits + 1) / (n_perm + 1)


def ca_trend(y, t, n_perm=N_PERM, seed=SEED):
    """Cochran–Armitage 趋势检验（个体得分 t）：返回 (Z, 渐近 p, 置换 p)。"""
    y, t = np.asarray(y, float), np.asarray(t, float)
    n, n1 = len(y), y.sum()
    if n1 in (0, n):
        return np.nan, np.nan, np.nan
    T = np.dot(t, y)
    e = n1 * t.mean()
    v = n1 * (n - n1) / (n * (n - 1)) * np.sum((t - t.mean()) ** 2)
    z = (T - e) / math.sqrt(v)
    rng = np.random.default_rng(seed)
    hits = 0
    for _ in range(n_perm // 10_000):              # 分批向量化
        P = rng.permuted(np.tile(y, (10_000, 1)), axis=1)
        hits += int(np.sum(np.abs(P @ t - e) >= abs(T - e) - 1e-9))
    return z, 2 * stats.norm.sf(abs(z)), (hits + 1) / (n_perm + 1)


def logit_trend(y, t):
    """logistic 回归，每年 OR 及 95% CI。"""
    try:
        with warnings.catch_warnings():
            warnings.simplefilter('ignore')
            r = sm.Logit(np.asarray(y, float), sm.add_constant(np.asarray(t, float))).fit(disp=0)
        b, se = r.params[1], r.bse[1]
        return math.exp(b), math.exp(b - 1.96 * se), math.exp(b + 1.96 * se), r.pvalues[1], ''
    except Exception as e:                           # 完全分离等
        return np.nan, np.nan, np.nan, np.nan, '模型未能拟合：%s' % type(e).__name__


def mnlogit_trend(cause, t):
    """多分类 logistic 回归，各类病因相对参照类的每年 RRR。"""
    cats = [MNL_REF] + [c for c in CAUSE4_ORDER if c != MNL_REF]
    counts = pd.Series(cause).value_counts()
    if any(counts.get(c, 0) < 10 for c in cats):
        return [], '有病因类别不足 10 例（%s），按 SAP 仅作描述' % counts.to_dict()
    y = pd.Categorical(cause, categories=cats).codes
    try:
        with warnings.catch_warnings():
            warnings.simplefilter('ignore')
            r = sm.MNLogit(y, sm.add_constant(np.asarray(t, float))).fit(disp=0, maxiter=200)
        if not r.mle_retvals.get('converged', True):
            return [], '模型未收敛，按 SAP 仅作描述'
    except Exception as e:
        return [], '模型未能拟合：%s' % type(e).__name__
    rows = []
    params, bse, pv = np.asarray(r.params), np.asarray(r.bse), np.asarray(r.pvalues)
    for j, c in enumerate(cats[1:]):               # 行＝自变量（0 截距、1 时间），列＝非参照类
        b, se, p = params[1, j], bse[1, j], pv[1, j]
        rows.append((c, math.exp(b), math.exp(b - 1.96 * se), math.exp(b + 1.96 * se), p))
    return rows, ''


def count_trend(y, denom, t):
    """月度计数的 Poisson 回归（offset=log 分母）；过度离散（Pearson χ²/df＞1.5）时改负二项。
    返回 (模型, 月数, 事件数, IRR, 下限, 上限, p, 离散度)。"""
    y, denom, t = np.asarray(y, float), np.asarray(denom, float), np.asarray(t, float)
    ok = ~np.isnan(y) & ~np.isnan(denom) & (denom > 0)
    y, denom, t = y[ok], denom[ok], t[ok]
    if len(y) < 12 or y.sum() == 0:
        return ('—', len(y), y.sum()) + (np.nan,) * 5
    X = sm.add_constant(t)
    off = np.log(denom)
    with warnings.catch_warnings():
        warnings.simplefilter('ignore')
        r = sm.GLM(y, X, family=sm.families.Poisson(), offset=off).fit()
        disp = r.pearson_chi2 / r.df_resid
        kind = 'Poisson'
        if disp > 1.5:
            r = sm.NegativeBinomial(y, X, offset=off).fit(disp=0, maxiter=200)
            kind = '负二项'
    b, se, p = r.params[1], r.bse[1], r.pvalues[1]
    return kind, len(y), int(y.sum()), math.exp(b), math.exp(b - 1.96 * se), math.exp(b + 1.96 * se), p, disp


def exact_rate(k, d, per):
    """精确 Poisson 95% CI 的率。"""
    if not d or np.isnan(d):
        return np.nan, np.nan, np.nan
    lo = stats.chi2.ppf(0.025, 2 * k) / 2 if k > 0 else 0.0
    hi = stats.chi2.ppf(0.975, 2 * k + 2) / 2
    return k / d * per, lo / d * per, hi / d * per


def cramers_v(x, y):
    t = pd.crosstab(x, y)
    if min(t.shape) < 2:
        return np.nan
    chi2 = stats.chi2_contingency(t, correction=False)[0]
    return math.sqrt(chi2 / (t.values.sum() * (min(t.shape) - 1)))


def n_pct(mask, n):
    k = int(mask.sum())
    return '%d (%.1f)' % (k, 100 * k / n) if n else '0'


def med_iqr(v):
    v = pd.to_numeric(pd.Series(v), errors='coerce').dropna()
    if v.empty:
        return '—'
    q1, q2, q3 = v.quantile([0.25, 0.5, 0.75])
    return '%.1f (%.1f–%.1f)' % (q2, q1, q3)


def fmt_p(p):
    if p is None or (isinstance(p, float) and np.isnan(p)):
        return '—'
    return '<0.001' if p < 0.001 else '%.3f' % p


def fmt_ci(e, lo, hi, nd=2):
    if e is None or np.isnan(e):
        return '—', '—'
    return ('%.*f' % (nd, e)), ('%.*f–%.*f' % (nd, lo, nd, hi))


# ================================================================ 各表
def flow_table(all_, inc, unfinished):
    rows = [('检索集（住院次数）', len(all_) + unfinished)]
    if unfinished:
        rows.append(('  其中尚未定稿（草稿模式下剔除）', unfinished))
    ex = all_[all_['是否排除'] == '是']
    rows.append(('排除', len(ex)))
    for e in EXCL_ORDER:
        rows.append(('  ' + e, int((ex['排除理由'] == e).sum())))
    rows.append(('纳入集', len(inc)))
    rows.append(('  其中病因经仲裁确定', int(inc['病因经仲裁'].sum())))
    rows.append(('  其中新生儿（≤28 天）', int((inc['年龄组4'] == '新生儿').sum())))
    return pd.DataFrame(rows, columns=['环节', '例数'])


def table1(inc):
    groups = [('总体', inc)] + [(c, inc[inc['病因4'] == c]) for c in CAUSE4_ORDER]
    rows = []

    def row(label, f, p=''):
        rows.append([label] + [f(g) for _, g in groups] + [p])

    row('例数', lambda g: str(len(g)))
    row('男性，n (%)', lambda g: n_pct(g['性别'] == '男性', len(g)),
        fmt_p(perm_rxc(inc['性别'], inc['病因4'])))
    row('年龄（岁），中位数 (IQR)', lambda g: med_iqr(g['年龄_岁']),
        fmt_p(stats.kruskal(*[g['年龄_岁'] for _, g in groups[1:] if len(g)])[1]))
    for a in AGE4_ORDER:
        row('  %s，n (%%)' % a, lambda g, a=a: n_pct(g['年龄组4'] == a, len(g)))
    for col, label in [('住院天数', '住院天数'), ('入院白细胞', '入院白细胞（×10⁹/L）'), ('入院CRP', '入院 CRP（mg/L）')]:
        vals = [pd.to_numeric(g[col], errors='coerce').dropna() for _, g in groups[1:]]
        p = stats.kruskal(*[v for v in vals if len(v)])[1] if sum(len(v) > 0 for v in vals) >= 2 else np.nan
        row(label + '，中位数 (IQR)', lambda g, col=col: med_iqr(g[col]), fmt_p(p))
        row('  缺失，n', lambda g, col=col: str(int(pd.to_numeric(g[col], errors='coerce').isna().sum())))
    if '提取已核对' in inc:                           # 提取表变量：分母为记录明确者（不含不详、未查、NR）
        inc_ok = inc[inc['提取已核对']]
        vals = [pd.to_numeric(g.loc[g['提取已核对'], '主诉病程_小时'], errors='coerce').dropna() for _, g in groups[1:]]
        p = stats.kruskal(*[v for v in vals if len(v)])[1] if sum(len(v) > 0 for v in vals) >= 2 else np.nan
        row('主诉病程（小时），中位数 (IQR)',
            lambda g: med_iqr(pd.to_numeric(g.loc[g['提取已核对'], '主诉病程_小时'], errors='coerce')), fmt_p(p))
        for c, pos in [(s, '是') for s in ext.SYMPTOMS] + [('影像游离气体_治疗前', '有')]:
            kn = _known(inc_ok[c])
            row('%s，n/N (%%)' % c, lambda g, c=c, pos=pos: k_of_n(g[c] == pos, g['提取已核对'] & _known(g[c])),
                fmt_p(perm_rxc(inc_ok.loc[kn, c], inc_ok.loc[kn, '病因4'])))
    return pd.DataFrame(rows, columns=['变量'] + [n for n, _ in groups] + ['P'])


def table2(inc):
    """病因构成按时期；6 类与 4 类，异物再分磁性/非磁性。"""
    per = [p for p, _, _ in PERIODS]
    rows = []
    ns = {p: int((inc['时期'] == p).sum()) for p in per}
    rows.append(['例数'] + [str(ns[p]) for p in per] + [str(len(inc)), ''])
    p4 = perm_rxc(inc['病因4'], inc['时期'])
    for c in CAUSE4_ORDER:
        m = inc['病因4'] == c
        rows.append([c] + [n_pct(m & (inc['时期'] == p), ns[p]) for p in per] + [n_pct(m, len(inc)),
                     fmt_p(p4) if c == CAUSE4_ORDER[0] else ''])
        if c == '异物相关':
            for s in ('磁性', '非磁性'):
                ms = m & (inc['异物类型'] == s)
                rows.append(['  ' + s] + [n_pct(ms & (inc['时期'] == p), ns[p]) for p in per] + [n_pct(ms, len(inc)), ''])
    rows.append(['6 类原始分类（SA3）'] + [''] * (len(per) + 2))
    p6 = perm_rxc(inc['病因大类'], inc['时期'])
    for c in CAUSE6_ORDER:
        m = inc['病因大类'] == c
        rows.append(['  ' + c] + [n_pct(m & (inc['时期'] == p), ns[p]) for p in per] + [n_pct(m, len(inc)),
                     fmt_p(p6) if c == CAUSE6_ORDER[0] else ''])
    return pd.DataFrame(rows, columns=['病因'] + per + ['合计', 'P（Fisher–Freeman–Halton）'])


def table3(inc):
    t = pd.crosstab(inc['病因4'], inc['年龄组4']).reindex(index=CAUSE4_ORDER, columns=AGE4_ORDER, fill_value=0)
    out = t.astype(str)
    for a in AGE4_ORDER:
        n = t[a].sum()
        out[a] = ['%d (%.1f)' % (v, 100 * v / n) if n else '0' for v in t[a]]
    out = out.reset_index().rename(columns={'病因4': '病因'})
    out.loc[len(out)] = ['合计'] + [str(t[a].sum()) for a in AGE4_ORDER]
    stat = pd.DataFrame([['Fisher–Freeman–Halton P', fmt_p(perm_rxc(inc['病因4'], inc['年龄组4'])), '', '', ''],
                         ["Cramér's V", '%.2f' % cramers_v(inc['病因4'], inc['年龄组4']), '', '', '']],
                        columns=out.columns)
    return pd.concat([out, stat], ignore_index=True)


def table5(inc):
    groups = [('总体', inc)] + [(c, inc[inc['病因4'] == c]) for c in CAUSE4_ORDER]
    rows = []
    for col, opts in [('穿孔部位', ['胃', '十二指肠', '胃+十二指肠']),
                      ('手术入路', ['开腹', '腹腔镜', '腔镜中转开腹', '内镜', '未手术']),
                      ('结局', ['治愈出院', '好转出院', '放弃治疗', '院内死亡'])]:
        rows.append([col] + [''] * len(groups))
        for o in opts:
            rows.append(['  ' + o] + [n_pct(g[col] == o, len(g)) for _, g in groups])
    rows.append(['复合不良结局（死亡或放弃治疗）'] +
                [n_pct(g['结局'].isin(['放弃治疗', '院内死亡']), len(g)) for _, g in groups])
    rows.append(['穿孔数目＞1'] + [n_pct(pd.to_numeric(g['穿孔数目'], errors='coerce') > 1, len(g)) for _, g in groups])
    mag = inc[inc['磁性'] == 1]
    rows.append(['磁性异物枚数，中位数 (IQR)'] +
                [med_iqr(pd.to_numeric(g.loc[g['磁性'] == 1, '磁性异物枚数'], errors='coerce')) for _, g in groups])
    rows.append(['  枚数记录为 NR，n'] + [str(int((g.loc[g['磁性'] == 1, '磁性异物枚数'] == 'NR').sum())) for _, g in groups])
    assert len(mag) == int(inc['磁性'].sum())
    rows.append(['术后住院天数，中位数 (IQR)'] + [med_iqr(g['术后住院天数']) for _, g in groups])
    rows.append(['检索集内 30 天再入院'] + [n_pct(g['30天再入院'], len(g)) for _, g in groups])
    if '提取已核对' in inc:
        ok = [(n, g[g['提取已核对']]) for n, g in groups]
        anyc = lambda g: g[ext.COMPLICATIONS].eq('是').any(axis=1) | (g['其他并发症'] != '')
        rows.append(['住院期间并发症（至少一项），n/N (%)'] + [k_of_n(anyc(g), pd.Series(True, index=g.index)) for _, g in ok])
        for c in ext.COMPLICATIONS:
            rows.append(['  %s' % c] + [k_of_n(g[c] == '是', _known(g[c])) for _, g in ok])
        grade = {v: i for i, v in enumerate(ext.CLAVIEN)}
        rows.append(['Clavien–Dindo ≥III，n/N (%)'] +
                    [k_of_n(g['最高Clavien-Dindo分级'].map(grade) >= grade['IIIa'], _known(g['最高Clavien-Dindo分级']))
                     for _, g in ok])
        rows.append(['非计划再次手术（手术病例），n/N (%)'] +
                    [k_of_n(g['非计划再次手术'] == '是', _known(g['非计划再次手术'])) for _, g in ok])
        rows.append(['非手术再干预，n/N (%)'] + [k_of_n(g['非手术再干预'] == '是', _known(g['非手术再干预'])) for _, g in ok])
        for c in ext.FEED:
            rows.append(['%s，中位数 (IQR)' % c] + [med_iqr(pd.to_numeric(g[c], errors='coerce')) for _, g in ok])
            rows.append(['  记录缺失（NR），n/N'] + ['%d/%d' % ((g[c] == 'NR').sum(), len(g)) for _, g in ok])
    return pd.DataFrame(rows, columns=['项目'] + [n for n, _ in groups])


def monthly_frame(inc, mon, fbden):
    months = study_months()
    f = pd.DataFrame(index=months)
    f['t'] = [t_of(m) for m in months]
    f['时期'] = [period_of(m) for m in months]
    f['全部病例'] = inc.groupby('月').size().reindex(months, fill_value=0)
    f['异物相关'] = inc[inc['异物'] == 1].groupby('月').size().reindex(months, fill_value=0)
    f['磁性异物'] = inc[inc['磁性'] == 1].groupby('月').size().reindex(months, fill_value=0)
    for c in ('核心', '核心+扩展', '急诊', '全院'):
        f[c] = mon[c].reindex(months)
    f['T18住院'] = fbden.reindex(months)
    return f


def trend_table(inc, f):
    rows = []

    def add(name, dataset, n, kind, est, lo, hi, p, note=''):
        e, ci = fmt_ci(est, lo, hi)
        rows.append([name, dataset, n, kind, e, ci, fmt_p(p), note])

    # 主要分析与 SA1、SA2
    for label, d in [('主要分析', inc),
                     ('SA1 非新生儿集', inc[inc['年龄组4'] != '新生儿']),
                     ('SA2 一致集', inc[~inc['病因经仲裁']])]:
        z, pa, pp = ca_trend(d['异物'], d['t'])
        orr, lo, hi, _, note = logit_trend(d['异物'], d['t'])
        add(label + '：异物相关占比趋势', '%d 例，其中异物相关 %d' % (len(d), d['异物'].sum()), len(d),
            'OR/年（logistic）', orr, lo, hi, pp,
            ('Cochran–Armitage Z=%.2f，渐近 P=%s，置换 P=%s（%d 次）' % (z, fmt_p(pa), fmt_p(pp), N_PERM)
             if not np.isnan(z) else '无法检验：异物相关为 0 例或全部') + ('；' + note if note else ''))

    # 多分类 logistic
    mrows, note = mnlogit_trend(inc['病因4'], inc['t'])
    if not mrows:
        rows.append(['多分类 logistic：各病因相对%s' % MNL_REF, '纳入集', len(inc), 'RRR/年', '—', '—', '—', note])
    for c, e, lo, hi, p in mrows:
        add('多分类 logistic：%s vs %s' % (c, MNL_REF), '纳入集', len(inc), 'RRR/年', e, lo, hi, p)

    # 相对发生率（SAP 4.3-3、SA4、SA5）
    if f['核心'].notna().any():
        for den in ('核心', '核心+扩展', '急诊', '全院'):
            if not f[den].notna().any():
                continue
            for y in ('全部病例', '异物相关'):
                tag = '次要分析' if den == '核心' else 'SA4'
                kind, nm, ev, e, lo, hi, p, disp = count_trend(f[y], f[den], f['t'])
                add('%s：%s 相对发生率趋势' % (tag, y), '分母=%s住院人次' % den, '%d 个月，%d 例' % (nm, ev),
                    'IRR/年（%s）' % kind, e, lo, hi, p, '离散度 %.2f' % disp if not np.isnan(disp) else '')
        g = f[f.index >= SA5_START]
        kind, nm, ev, e, lo, hi, p, disp = count_trend(g['全部病例'], g['核心'], g['t'])
        add('SA5：全部病例相对发生率（仅断档后连续月份）', '分母=核心住院人次', '%d 个月，%d 例' % (nm, ev),
            'IRR/年（%s）' % kind, e, lo, hi, p)
    else:
        rows.append(['相对发生率趋势（次要分析、SA4、SA5）', '—', '—', '—', '—', '—', '—',
                     '病案室分母尚未回填，已跳过'])

    # 异物穿孔转化率
    g = f[f.index >= FB_START]
    for y in ('异物相关', '磁性异物'):
        kind, nm, ev, e, lo, hi, p, disp = count_trend(g[y], g['T18住院'], g['t'])
        add('次要分析：%s穿孔 / 异物住院 转化率趋势' % ('' if y == '异物相关' else '磁性'),
            '分母=T18 住院人次（自 2016-07）', '%d 个月，%d 例' % (nm, ev), 'IRR/年（%s）' % kind, e, lo, hi, p)
    return pd.DataFrame(rows, columns=['分析', '分析集/分母', '样本', '效应量', '估计', '95% CI', 'P', '备注'])


def rate_table(f):
    rows = []
    for name, a, b in PERIODS:
        g = f[(f.index >= pd.Period(a, 'M')) & (f.index <= pd.Period(b, 'M'))]
        r = [name, len(g), int(g['全部病例'].sum())]
        for y, den, per in [('全部病例', '核心', 1e4), ('异物相关', '核心', 1e4),
                            ('异物相关', 'T18住院', 1e3), ('磁性异物', 'T18住院', 1e3)]:
            gg = g[g[den].notna()]
            e, lo, hi = exact_rate(int(gg[y].sum()), gg[den].sum() if len(gg) else np.nan, per)
            r.append('—' if np.isnan(e) else '%.2f (%.2f–%.2f)' % (e, lo, hi))
        rows.append(r)
    return pd.DataFrame(rows, columns=['时期', '有效月数', '病例数', '全部病例/万核心住院', '异物相关/万核心住院',
                                       '异物相关穿孔/千异物住院', '磁性异物穿孔/千异物住院'])


def completeness_table(m6):
    """第 6 项核对：病案室编码计数 vs 科研平台检索集（全部 107 次，不论是否纳入）。"""
    if m6 is None:
        return pd.DataFrame([['病案室第 6 项尚未回填，已跳过']], columns=['说明'])
    x = pd.ExcelFile(SRC)
    h = x.parse('病案首页基本信息')
    h['月'] = pd.to_datetime(h['入院日期']).dt.to_period('M')
    dx = x.parse('病案出院诊断编目后')
    code, name = dx['诊断疾病编码'].astype(str).str.upper(), dx['诊断疾病名称'].astype(str)
    perf = name.str.contains('穿孔|破裂')
    grp = {
        req.CODE_GROUPS[0]: code.str.startswith('K31.8') & perf & name.str.contains('胃') & ~name.str.contains('十二指肠'),
        req.CODE_GROUPS[1]: code.str.startswith('K31.8') & perf & name.str.contains('十二指肠'),
        req.CODE_GROUPS[2]: code.str.match(r'K2[567]\.[1256]'),
        req.CODE_GROUPS[3]: (code.str.startswith('P78.8') & perf & name.str.contains('胃')) | code.str.startswith('Q40.204'),
        req.CODE_GROUPS[4]: code.str.startswith('S36.3') | (code.str.startswith('S36.4') & name.str.contains('十二指肠')),
    }
    month_of = h.set_index(KEY)['月']
    rows = []
    months = pd.period_range(STUDY[0], STUDY[1], freq='M')
    anyk = set()
    ours = {}
    for g, m in grp.items():
        ks = set(dx.loc[m, KEY])
        anyk |= ks
        ours[g] = pd.Series([month_of[k] for k in ks if k in month_of.index]).value_counts()
    ours['合计（去重）'] = pd.Series([month_of[k] for k in anyk if k in month_of.index]).value_counts()
    for m in months:
        for g in req.CODE_GROUPS + ['合计（去重）']:
            hosp = m6.at[m, g] if m in m6.index else np.nan
            mine = int(ours[g].get(m, 0))
            if pd.notna(hosp) and hosp > mine:
                rows.append([str(m), g, int(hosp), mine, int(hosp) - mine, '断档期' if in_gap(m) else ''])
    out = pd.DataFrame(rows, columns=['月份', '编码组', '病案室计数', '本课题检索集', '差值（疑似遗漏）', '备注'])
    if out.empty:
        out = pd.DataFrame([['各月各组病案室计数均不多于本课题检索集，未见遗漏']], columns=['说明'])
    return out


def figure2(inc, f, path, draft):
    import matplotlib
    matplotlib.use('Agg')
    import matplotlib.pyplot as plt

    years = list(range(STUDY[0].year, STUDY[1].year + 1))
    valid = f.groupby(f.index.year).size().reindex(years, fill_value=0)
    cnt = pd.crosstab(inc['年'], inc['病因4']).reindex(index=years, columns=CAUSE4_ORDER, fill_value=0)
    has_rate = f['核心'].notna().any()

    plt.rcParams.update({'font.family': 'DejaVu Sans', 'font.size': 9, 'axes.edgecolor': INK2,
                         'axes.labelcolor': INK2, 'xtick.color': INK2, 'ytick.color': INK2})
    fig, axes = plt.subplots(2 if has_rate else 1, 1, figsize=(7.0, 5.6 if has_rate else 3.6), sharex=True,
                             gridspec_kw={'height_ratios': [3, 1.6]} if has_rate else None, facecolor=SURFACE)
    axes = np.atleast_1d(axes)
    ax = axes[0]
    x = np.arange(len(years))
    bottom = np.zeros(len(years))
    for c in CAUSE4_ORDER:
        ax.bar(x, cnt[c].values, 0.62, bottom=bottom, color=FIG_COLORS[c], label=FIG_EN[c],
               edgecolor=SURFACE, linewidth=1.5)            # 段间 2px 左右的底色缝
        bottom += cnt[c].values
    # 末根柱右侧直接标注类别（黄、青对比度不足 3:1 的补救）；标签过近时依次上推，避免重叠
    i, base, ys = len(years) - 1, 0.0, []
    for c in CAUSE4_ORDER:
        h = cnt[c].values[i]
        if h > 0:
            ys.append([base + h / 2, c])
        base += h
    gap = max(bottom.max(), 1) * 0.07
    for j in range(1, len(ys)):
        ys[j][0] = max(ys[j][0], ys[j - 1][0] + gap)
    for y, c in ys:
        ax.annotate(FIG_EN[c], (x[i] + 0.4, y), fontsize=7.5, color=INK2, va='center', annotation_clip=False)
    ax.set_xlim(-0.6, len(years) + 0.9)
    ax.set_ylabel('Cases (n)')
    ax.set_facecolor(SURFACE)
    ax.grid(axis='y', color=GRID, linewidth=0.8)
    ax.set_axisbelow(True)
    for s in ('top', 'right'):
        ax.spines[s].set_visible(False)
    ax.legend(ncol=4, frameon=False, loc='lower left', bbox_to_anchor=(0, 1.0), fontsize=8, labelcolor=INK)
    ax.set_title('Gastroduodenal perforation by etiology and year of admission', loc='left', color=INK,
                 fontsize=10, pad=22)
    if has_rate:
        ay = axes[1]
        yr = f.groupby(f.index.year).agg(k=('全部病例', 'sum'), d=('核心', 'sum')).reindex(years)
        rate = yr['k'] / yr['d'] * 1e4
        ay.plot(x, rate.values, color=INK2, linewidth=2, marker='o', markersize=4.5,
                markeredgecolor=SURFACE, markeredgewidth=1.5)
        ay.set_ylabel('Per 10,000 surgical\nadmissions')
        ay.set_ylim(bottom=0)
        ay.set_facecolor(SURFACE)
        ay.grid(axis='y', color=GRID, linewidth=0.8)
        ay.set_axisbelow(True)
        for s in ('top', 'right'):
            ay.spines[s].set_visible(False)
    axes[-1].set_xticks(x)
    axes[-1].set_xticklabels(['%d\n(%d mo)' % (y, valid[y]) if valid[y] < 12 else str(y) for y in years], fontsize=8)
    note = ('Months in parentheses: observed months in partial years (study window Jun 2016–May 2026; '
            'no platform records Jul 2019–Mar 2020).')
    fig.text(0.01, 0.005, note, fontsize=7, color=INK2)
    if draft:
        fig.text(0.99, 0.99, 'DRAFT — not for citation', ha='right', va='top', fontsize=9, color='#d03b3b')
    fig.tight_layout(rect=(0, 0.03, 1, 1))
    fig.savefig(path, dpi=300, facecolor=SURFACE)
    plt.close(fig)


# ================================================================ 写盘
def backup(path):
    if os.path.exists(path):
        base, ext = os.path.splitext(path)
        bak = '%s.backup_%s%s' % (base, datetime.now().strftime('%Y%m%d_%H%M%S'), ext)
        shutil.copy2(path, bak)
        print('已备份旧文件 ->', bak)


def main():
    draft = '--draft' in sys.argv[1:]
    all_, inc, unfinished, status = load_cases(draft)
    inc = add_clinical(inc)
    inc, ext_note = add_extraction(inc, draft)
    mon, den_notes, m6 = load_denoms()
    fbden = load_fb()
    f = monthly_frame(inc, mon, fbden)

    kappa = pd.read_excel(MERGED, sheet_name='一致性κ')
    tabs = {
        '图1_流程': flow_table(all_, inc, unfinished),
        '表1_基本特征': table1(inc),
        '表2_病因×时期': table2(inc),
        '表3_病因×年龄组': table3(inc),
        '表4_趋势分析': trend_table(inc, f),
        '表4附_分期率': rate_table(f),
        '表5_治疗与结局': table5(inc),
        '补充_裁定一致性': kappa,
        '补充_第6项核对': completeness_table(m6),
        '补充_逐月数据': f.assign(月份=f.index.astype(str)).reset_index(drop=True),
    }
    inputs = [MERGED, SRC, FB] + [p for p in (DENOM, EXTRACT) if os.path.exists(p)]
    info = pd.DataFrame(
        ([('★草稿', '以 --draft 运行，结果不得引用。终裁状态：%s（草稿模式仅分析已定稿病例）；次要终点提取表：%s'
                    % (status, ext_note))] if draft else []) + [
            ('生成时间', datetime.now().strftime('%Y-%m-%d %H:%M')),
            ('依据', '统计分析计划_SAP_v1（各表编号与 SAP 第五节一致）'),
            ('纳入集', '%d 例；有效观察月 %d 个（研究窗口 %s 至 %s，扣除断档 %s）'
                       % (len(inc), len(study_months()), STUDY[0], STUDY[1],
                          '%s 至 %s' % GAP if GAP else '无')),
            ('分母', '；'.join('%s：%s' % kv for kv in den_notes.items()) or '病案室分母尚未回填，相对发生率分析已跳过'),
            ('次要终点提取表', ext_note),
            ('随机种子', '%d（主要分析置换 %d 次，列联表蒙特卡洛 %d 次）' % (SEED, N_PERM, N_PERM_TAB)),
            ('软件', 'Python %s；pandas %s；numpy %s；scipy %s；statsmodels %s'
                     % (platform.python_version(), pd.__version__, np.__version__, scipy.__version__,
                        statsmodels.__version__)),
        ] + [('输入 SHA256', '%s  %s' % (sha256(p), p)) for p in inputs],
        columns=['项目', '说明'])

    backup(OUT)
    with pd.ExcelWriter(OUT, engine='openpyxl') as w:
        info.to_excel(w, sheet_name='说明', index=False)
        for name, t in tabs.items():
            t.to_excel(w, sheet_name=name, index=False)
        for name, ws in w.sheets.items():
            for col in ws.columns:
                ws.column_dimensions[col[0].column_letter].width = 18
            ws.column_dimensions['A'].width = 34
            ws.freeze_panes = 'B2'
        w.sheets['说明'].column_dimensions['B'].width = 110
        w.sheets['表4_趋势分析'].column_dimensions['H'].width = 70
    print('已写出 ->', OUT)

    backup(FIG)
    figure2(inc, f, FIG, draft)
    print('已写出 ->', FIG)
    print(tabs['表4_趋势分析'][['分析', '估计', '95% CI', 'P']].to_string(index=False))


if __name__ == '__main__':
    main()
