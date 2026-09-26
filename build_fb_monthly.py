# -*- coding: utf-8 -*-
"""从《10年消化道异物原始数据.xlsx》归纳逐月消化道异物（ICD-10 T18）住院人次，
对应《病案数据调取申请》第 3 项。

口径：
  - 按入院日期归月，按住院人次计（就诊编号唯一，不按患者去重）。
  - 主口径 = 编目后出院诊断中任一编码以 T18 开头。
  - 导出起点为 2016-07-03，此前月份记为缺失（NA），不是 0。
  - 另列兄弟课题（磁性异物十年趋势）分析集 1242 次的逐月数，便于两篇文章对数。
  - 穿孔队列中带 T18 编码、但未被该导出收录的就诊单列「补回」，不混入主口径。

安全约定：目标文件若已存在先备份；写盘只在 __main__。
"""
import hashlib
import os
import shutil
from datetime import datetime

import pandas as pd

FB_XLSX = r'D:\消化道异物\10年消化道异物原始数据.xlsx'
SIB_CSV = r'D:\消化道异物\① 磁性异物的十年趋势 + 手术风险预测模型\derived\analysis_dataset.csv'
SIB_YEARLY = r'D:\消化道异物\① 磁性异物的十年趋势 + 手术风险预测模型\derived\yearly_counts.csv'
PERF_XLSX = r'D:\胃十二指肠穿孔\胃十二指肠穿孔.xlsx'
DST = r'D:\胃十二指肠穿孔\①儿童胃十二指肠穿孔病因谱十年变迁\消化道异物_逐月住院人次_v1.xlsx'

KEY = '科研就诊编号'
PUWAI = ('普外一', '普外二', '胃肠外科', '肝胆外科')   # 普外系统（2024-08 更名前后四个名称）
EXPORT_START = pd.Period('2016-07', 'M')                # 导出实际起点
STUDY = (pd.Period('2016-06', 'M'), pd.Period('2026-05', 'M'))   # 穿孔研究窗口
# 科研数据平台断档：异物导出与穿孔导出在此期间均为 0 次住院（前后月份每月 5–15 次异物住院），
# 2019 年下半年无法用疫情解释，判定为平台数据缺失。计数记为空值，分析时必须扣除这段观察期。
GAP = (pd.Period('2019-07', 'M'), pd.Period('2020-03', 'M'))
ROWS = pd.period_range('2016-01', '2026-06', freq='M')


def sha256(path):
    h = hashlib.sha256()
    with open(path, 'rb') as f:
        for b in iter(lambda: f.read(1 << 20), b''):
            h.update(b)
    return h.hexdigest()


def is_main(s):
    return s.astype(str).str.contains('是|1', na=False)


def load():
    fx = pd.ExcelFile(FB_XLSX)
    h = fx.parse('病案首页基本信息')
    h['入院日期'] = pd.to_datetime(h['入院日期'], errors='coerce')
    h['月'] = h['入院日期'].dt.to_period('M')
    d = fx.parse('病案出院诊断编目后')
    d['T18'] = d['诊断疾病编码'].astype(str).str.upper().str.startswith('T18')
    any_t18 = set(d.loc[d['T18'], KEY])
    main_t18 = set(d.loc[d['T18'] & is_main(d['是否主要诊断']), KEY])
    h['任一T18'] = h[KEY].isin(any_t18)
    h['主诊断T18'] = h[KEY].isin(main_t18)
    h['普外系统'] = h['出院科别'].astype(str).map(lambda s: any(k in s for k in PUWAI))

    sib = pd.read_csv(SIB_CSV, usecols=[KEY])
    h['兄弟课题分析集'] = h[KEY].isin(set(sib[KEY]))

    # 穿孔队列中带 T18 编码、却不在异物导出里的就诊
    px = pd.ExcelFile(PERF_XLSX)
    ph = px.parse('病案首页基本信息')
    ph['入院日期'] = pd.to_datetime(ph['入院日期'], errors='coerce')
    pdx = px.parse('病案出院诊断编目后')
    p_t18 = pdx[pdx['诊断疾病编码'].astype(str).str.upper().str.startswith('T18')]
    cand = ph[ph[KEY].isin(set(p_t18[KEY]))]
    miss = cand[~cand[KEY].isin(set(h[KEY])) & ~cand['住院号'].astype(str).isin(set(h['住院号'].astype(str)))].copy()
    miss['月'] = miss['入院日期'].dt.to_period('M')
    main_dx = pdx[is_main(pdx['是否主要诊断'])].set_index(KEY)['诊断疾病名称']
    miss['主要诊断'] = miss[KEY].map(main_dx)
    miss['T18诊断'] = miss[KEY].map(p_t18.groupby(KEY)['诊断疾病名称'].apply(lambda s: '、'.join(s.astype(str))))
    return h, miss, len(cand)


def monthly(h, miss):
    rows = []
    for m in ROWS:
        covered = m >= EXPORT_START
        gap = GAP[0] <= m <= GAP[1]
        in_study = STUDY[0] <= m <= STUDY[1]
        g = h[h['月'] == m]
        if gap:
            assert len(g) == 0, '断档月 %s 实有 %d 次住院，断档判定需复核' % (m, len(g))
        status = ('导出未覆盖（起于2016-07-03）' if not covered else
                  '平台断档（勿作零值）' if gap else '有效')
        r = {'年': m.year, '月': m.month, '数据状态': status,
             '在穿孔研究窗口内': '是' if in_study else '否'}
        if covered and not gap:
            r.update({
                '全部导出就诊': len(g),
                '兄弟课题分析集': int(g['兄弟课题分析集'].sum()),
                '★出院诊断任一T18': int(g['任一T18'].sum()),
                '其中主要诊断为T18': int(g['主诊断T18'].sum()),
                '其中出院于普外系统': int((g['任一T18'] & g['普外系统']).sum()),
                '穿孔队列漏收T18_补回': int((miss['月'] == m).sum()),
            })
            r['任一T18_含补回'] = r['★出院诊断任一T18'] + r['穿孔队列漏收T18_补回']
        rows.append(r)
    df = pd.DataFrame(rows)
    # 导出覆盖前的月份为空值，会把整列变成浮点；改用可空整数，Excel 里显示为整数
    for c in df.columns[4:]:
        df[c] = df[c].astype('Int64')
    return df


def main():
    h, miss, n_cand = load()
    mon = monthly(h, miss)

    # ---- 自检：逐月合计必须等于总数，任何一条不符即中止 ----
    tot = {'全部导出就诊': len(h), '兄弟课题分析集': int(h['兄弟课题分析集'].sum()),
           '★出院诊断任一T18': int(h['任一T18'].sum()), '其中主要诊断为T18': int(h['主诊断T18'].sum())}
    for c, v in tot.items():
        assert int(mon[c].sum()) == v, '%s 逐月合计 %s ≠ 总数 %s' % (c, mon[c].sum(), v)
    assert tot['全部导出就诊'] == 1244 and tot['兄弟课题分析集'] == 1242, tot
    assert tot['★出院诊断任一T18'] == 1218, tot      # 与兄弟课题 cohort_flow.txt 一致
    assert h['月'].min() == EXPORT_START, h['月'].min()

    # ---- 年度汇总，并与兄弟课题 yearly_counts.csv 对照 ----
    cov = mon[mon['数据状态'] == '有效']
    yr = cov.groupby('年')[['全部导出就诊', '兄弟课题分析集', '★出院诊断任一T18', '其中主要诊断为T18',
                            '其中出院于普外系统', '穿孔队列漏收T18_补回', '任一T18_含补回']].sum().reset_index()
    sy = pd.read_csv(SIB_YEARLY, encoding='utf-8-sig')[['year', 'total_FB']].rename(
        columns={'year': '年', 'total_FB': '兄弟课题yearly_counts'})
    yr.insert(1, '有效月数', yr['年'].map(cov.groupby('年').size()))
    yr = yr.merge(sy, on='年', how='left')
    yr['与兄弟课题年度数一致'] = (yr['兄弟课题分析集'] == yr['兄弟课题yearly_counts']).map({True: '是', False: '否'})

    info = pd.DataFrame([
        ('数据来源', FB_XLSX),
        ('来源文件 SHA256', sha256(FB_XLSX)),
        ('生成时间', datetime.now().strftime('%Y-%m-%d %H:%M')),
        ('对应申请单', '《病案数据调取申请_分母数据》第 3 项：逐月消化道异物（ICD-10 T18）住院人次'),
        ('归月与计数', '按入院日期归月；按住院人次计，同一患儿多次住院分别计入'),
        ('★主口径', '编目后出院诊断中任一编码以 T18 开头（共 %d 次）' % tot['★出院诊断任一T18']),
        ('覆盖范围', '导出起于 2016-07-03、止于 2026-06-30；2016-01 至 2016-06 无数据，记为空（不是 0）。'
                    '穿孔研究窗口首月 2016-06 因此缺失。'),
        ('人群范围', '仅住院患儿；门急诊处理未住院者不在导出内（兄弟课题 README 所述检索条件）。'
                    '本院消化道异物住院患儿 98.5% 出院于普外系统，故可视作全院 T18 住院量。'),
        ('已知遗漏', '穿孔队列中带 T18 编码的 %d 次就诊，有 %d 次未被本导出收录。这些就诊的 T18 均为次要诊断，'
                    '主要诊断为穿孔或脓毒症，门急诊诊断亦无「异物」字样，推测因此漏检。'
                    '漏收的恰是最重的病例；本表仅能补回胃十二指肠穿孔队列中的这几次，'
                    '其他部位穿孔或梗阻等同类漏检无法估计。' % (n_cand, len(miss))),
        ('★平台断档', '2019-07 至 2020-03 共 9 个月，异物导出与胃十二指肠穿孔导出均为 0 次住院，'
                     '而前后月份每月 5–15 次异物住院。2020 年 1–3 月可能叠加武汉封城，但 2019 年下半年'
                     '无法用疫情解释，判定为科研数据平台断档。表中这 9 个月记为空值（不是 0），'
                     '计算发生率时须从观察期中扣除；病案室提供的官方分母若在此期间有数，也不能与本数据的分子配对。'
                     '「逐年汇总」的「有效月数」已扣除这 9 个月：2019 年仅 6 个有效月，2020 年 9 个。'),
        ('导出中无 T18 编码者', '%d 次（多为阑尾炎穿孔、肠梗阻等，应为按文本关键词带入），不计入主口径'
                             % (tot['全部导出就诊'] - tot['★出院诊断任一T18'])),
        ('兄弟课题分析集', '1242 次 = 1244 次剔除 2 次术后伤口处理再入院；列出以便两篇文章对数'),
    ], columns=['项目', '说明'])

    miss_out = miss[[KEY, '入院日期', '出院科别', '主要诊断', 'T18诊断']].copy()
    miss_out['入院日期'] = miss_out['入院日期'].dt.strftime('%Y-%m-%d')

    if os.path.exists(DST):
        bak = DST.replace('.xlsx', '.backup_%s.xlsx' % datetime.now().strftime('%Y%m%d_%H%M%S'))
        shutil.copy2(DST, bak)
        print('已备份旧文件 ->', bak)
    with pd.ExcelWriter(DST, engine='openpyxl') as w:
        info.to_excel(w, sheet_name='说明', index=False)
        mon.to_excel(w, sheet_name='逐月', index=False)
        yr.to_excel(w, sheet_name='逐年汇总', index=False)
        miss_out.to_excel(w, sheet_name='漏收的T18就诊', index=False)
        w.sheets['说明'].column_dimensions['A'].width = 18
        w.sheets['说明'].column_dimensions['B'].width = 110
        for ws in (w.sheets['逐月'], w.sheets['逐年汇总']):
            for col in ws.columns:
                ws.column_dimensions[col[0].column_letter].width = 16
            ws.freeze_panes = 'C2'

    print('已写出 ->', DST)
    print('总数:', tot, '| 漏收补回 %d / 穿孔队列T18就诊 %d' % (len(miss), n_cand))
    print(yr.to_string(index=False))


if __name__ == '__main__':
    main()
