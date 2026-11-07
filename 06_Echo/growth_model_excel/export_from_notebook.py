"""
Export script for the Echo growth-model workbooks.

What it does
------------
1. Runs every code cell of the analysis notebook (default: ../echo_growth_analysis.ipynb)
   inside a temporary folder, so the notebook's figures/clean files are NOT overwritten.
   The workbooks are therefore built from exactly the same objects as notebook section 9
   (retention curves, trends, calibration factor, scenarios, what-if table, backtest).
2. Computes the extra series the workbooks need (contribution of existing cohorts per week,
   weekly new users, the least-squares "equivalent churn" of the simple model, ...).
3. Writes everything to model_data.json, which build_workbooks.py turns into the two .xlsx files.

Usage (from this folder):
    python export_from_notebook.py                      # uses ../echo_growth_analysis.ipynb
    python export_from_notebook.py --notebook other.ipynb --out other.json
    python build_workbooks.py                           # then rebuild the two workbooks

Needs: pandas, numpy, scipy, matplotlib, jinja2, IPython (same as the notebook).
"""
import argparse
import contextlib
import io
import json
import os
import tempfile
from pathlib import Path

import matplotlib
matplotlib.use('Agg')
import numpy as np
import pandas as pd
from scipy.optimize import minimize_scalar

HERE = Path(__file__).resolve().parent


# --------------------------------------------------------------------------------------
# 1. Run the notebook
# --------------------------------------------------------------------------------------
def run_notebook(nb_path: Path, data_dir: Path) -> dict:
    import IPython.display as ipd
    import matplotlib.pyplot as plt
    ipd.display = lambda *a, **k: None          # silence display() calls of the notebook
    nb = json.loads(nb_path.read_text(encoding='utf-8'))
    ns = {'__name__': '__notebook__'}
    tmp = Path(tempfile.mkdtemp(prefix='echo_nb_'))
    os.symlink(data_dir, tmp / 'data')
    old = os.getcwd()
    os.chdir(tmp)
    log = io.StringIO()
    try:
        for i, cell in enumerate(nb['cells']):
            if cell['cell_type'] != 'code':
                continue
            src = ''.join(cell['source'])
            if src.lstrip().startswith(('%', '!')):
                continue
            with contextlib.redirect_stdout(log):
                exec(compile(src, f'<notebook cell {i}>', 'exec'), ns)
            plt.close('all')
    finally:
        os.chdir(old)
    ns['_log'] = log.getvalue()
    return ns


# --------------------------------------------------------------------------------------
# 2. Simple model (version 1) helpers: must replicate the Excel recurrence exactly
# --------------------------------------------------------------------------------------
def simple_path(u0, churn, paid, free, referral=0.0):
    """End users of each week: U(t) = U(t-1) x (1 - churn) x (1 + referral) + paid(t) + free(t)."""
    out, u = [], u0
    for p, f in zip(paid, free):
        u = u * (1 - churn) * (1 + referral) + p + f
        out.append(u)
    return np.array(out)


def fit_churn(u0, target, paid, free):
    """Least squares: churn minimising sum over weeks of (simple model - target)^2."""
    target = np.asarray(target, float)
    res = minimize_scalar(lambda c: ((simple_path(u0, c, paid, free) - target) ** 2).sum(),
                          bounds=(0.0, 0.5), method='bounded', options={'xatol': 1e-10})
    path = simple_path(u0, res.x, paid, free)
    return {'churn': float(res.x), 'sse': float(res.fun),
            'rmse': float(np.sqrt(res.fun / len(target))),
            'mape': float(np.mean(np.abs(path - target) / target)),
            'end_gap': float(path[-1] / target[-1] - 1), 'path': path.tolist()}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--notebook', default=str(HERE.parent / 'echo_growth_analysis.ipynb'))
    ap.add_argument('--data', default=str(HERE.parent / 'data'))
    ap.add_argument('--out', default=str(HERE / 'model_data.json'))
    args = ap.parse_args()

    ns = run_notebook(Path(args.notebook).resolve(), Path(args.data).resolve())
    g = ns  # short alias
    d = {'source_notebook': Path(args.notebook).name}

    # ---------------- constants and A/B-test parameters (notebook sections 5, 7, 9.3) ----------
    d['const'] = {
        'CPM': float(g['CPM']), 'REV_PER_AD': float(g['REV_PER_AD']), 'ADS_PER_HOUR': float(g['ADS_PER_HOUR']),
        'H': int(g['H']), 'REV_RATIO': float(g['REV_RATIO']), 'PHI': float(g['PHI']),
        'APH_RATIO': float(g['APH_RATIO']), 'CALIB': float(g['CALIB']),
        'CAC_PAID': float(g['CAC_PAID']), 'LAST_FULL_WEEK': str(g['LAST_FULL_WEEK'].date()),
        'F_START': str(g['F_START'].date()), 'F_END': str(g['F_END'].date()),
        'actual_wau_last_full_week': float(g['kpi'].WAU.loc[g['LAST_FULL_WEEK']]),
        'model_existing_wau_last_full_week': float(g['nowcast'].WAU.iloc[0]),
        'arpu_week_data': float(g['arpu_w'].loc['all', 'revenue_per_active_week']),
    }
    d['ratio_measured'] = [float(x) for x in g['ratio']]          # R(a), ages 0..H (smoothed)

    # ---------------- retention / value curves by segment (notebook 9.1, build_curves) ----------
    curves = g['curves']
    d['curves'] = {f'{ch}|{int(act)}': {'S': curves[(ch, act)].S.tolist(), 'v': curves[(ch, act)].v.tolist(),
                                        'observed': curves[(ch, act)].observed.astype(int).tolist()}
                   for ch, act in g['GROUPS']}

    # ---------------- trends used to forecast new users (fit_trends) ----------
    tr = g['tr']
    d['trends'] = {'m0': str(tr['m0'].date()), 'cac': float(tr['cac']),
                   'organic_store': [float(x) for x in tr['trends']['organic_store']],
                   'podcast_cross_promo': [float(x) for x in tr['trends']['podcast_cross_promo']],
                   'spend': [float(x) for x in tr['spend_trend']],
                   'act': {k: float(v) for k, v in tr['act'].items()}}
    u, spend = g['u'], g['spend']
    m = u[u.signup_month <= pd.Timestamp('2026-06-01')].groupby(['signup_month', 'acquisition_channel']).size().unstack()
    spend_col = 'spend_imputed' if 'spend_imputed' in spend.columns and spend['spend_imputed'].dtype != bool else 'spend_eur'
    sp = spend.set_index('month')[spend_col]
    d['monthly_history'] = {'month': [str(x.date()) for x in m.index],
                            'organic_store': m.organic_store.astype(int).tolist(),
                            'paid_social': m.paid_social.astype(int).tolist(),
                            'podcast_cross_promo': m.podcast_cross_promo.astype(int).tolist(),
                            'spend': [float(sp.loc[x]) for x in m.index],
                            'spend_column_used': spend_col,
                            'spend_raw': [None if pd.isna(x) else float(x) for x in spend.set_index('month').spend_eur.reindex(m.index)]}

    # ---------------- forecast weeks, new users (base), existing cohorts ----------
    f_weeks = g['f_weeks']
    d['weeks'] = [str(x.date()) for x in f_weeks]
    new, spend_total = g['forecast_new'](tr, '2026-07-01', g['F_END'] + pd.Timedelta(days=6))
    nw = new.groupby(['signup_week', 'channel']).n.sum().unstack().reindex(f_weeks)
    d['new_users_base_weekly'] = {ch: nw[ch].tolist() for ch in g['CHANNELS']}
    d['paid_spend_12m_base'] = float(spend_total)

    exist = g['exist']
    tot = g['simulate'](f_weeks, exist, curves)                      # calib = 1, no scenario effect
    d['existing'] = {'WAU': tot.WAU.tolist(), 'revenue': tot.revenue.tolist(), 'by_segment': {}}
    for ch, act in g['GROUPS']:
        sub = g['simulate'](f_weeks, exist[(exist.channel == ch) & (exist.activated == act)], curves)
        d['existing']['by_segment'][f'{ch}|{int(act)}'] = {'WAU': sub.WAU.tolist(), 'revenue': sub.revenue.tolist()}

    # ---------------- notebook results (targets of the Checks sheets) ----------
    scen = g['scen']
    keys = list(scen.keys())                      # [Pessimistic, Base, Optimistic]
    d['scenarios'] = {}
    for name, short in zip(keys, ['Pessimistic', 'Base', 'Optimistic']):
        s, sp_tot = scen[name]
        d['scenarios'][short] = {'label': name, 'WAU': s.WAU.tolist(), 'revenue': s.revenue.tolist(),
                                 'hours': s['listening hours'].tolist(), 'paid_spend_12m': float(sp_tot)}
    wi = g['wi']
    d['whatif'] = {idx: {c: float(wi.loc[idx, c]) for c in wi.columns} for idx in wi.index}
    # weekly WAU of every what-if line (needed to calibrate the simple model on each of them)
    whatif_kw = dict(g['whatif'])
    d['whatif_paths'] = {}
    for name, kw in whatif_kw.items():
        _, s, sp_tot = g['run'](name, **kw)
        d['whatif_paths'][name] = {'kw': {k: (float(v) if not isinstance(v, bool) else v) for k, v in kw.items()},
                                   'WAU': s.WAU.tolist(), 'revenue': s.revenue.tolist(), 'paid_spend_12m': float(sp_tot)}
    ls = g['ltv_seg']
    d['ltv_seg'] = {ch: {c: float(ls.loc[ch, c]) for c in ls.columns} for ch in ls.index}
    lt = g['ltv']
    d['ltv_channel'] = {ch: {'LTV 52 weeks': float(lt.loc[ch, 'LTV 52 weeks']), 'LTV 104 weeks': float(lt.loc[ch, 'LTV 104 weeks'])}
                        for ch in lt.index}

    # ---------------- backtest (notebook 9.2) ----------
    bt = g['bt']
    d['backtest'] = {'weeks': [str(x.date()) for x in g['bt_weeks']],
                     'actual_WAU': g['act'].WAU.tolist(), 'actual_revenue': g['act'].revenue.tolist(),
                     'A_WAU': g['pred_A'].WAU.tolist(), 'A_revenue': g['pred_A'].revenue.tolist(),
                     'B_WAU': g['pred_B'].WAU.tolist(), 'B_revenue': g['pred_B'].revenue.tolist(),
                     'table': {c: {r: float(bt.loc[r, c]) for r in bt.index} for c in bt.columns}}

    # ---------------- weekly history (actual) ----------
    kpi = g['kpi']
    d['history'] = {'weeks': [str(x.date()) for x in kpi.index], 'WAU': kpi.WAU.astype(float).tolist(),
                    'revenue': kpi.revenue.tolist(), 'hours': kpi['listening hours'].tolist()}
    sw = u.groupby('signup_week').size()
    d['history']['signups_by_signup_week'] = [float(sw.get(x, 0)) for x in kpi.index]

    # ---------------- naive week-to-week churn measured on data (for the README) ----------
    w = g['w']
    act_sets = w.groupby('week_start').user_id.apply(set)
    wks = sorted(act_sets.index)
    lost = [len(act_sets[a] - act_sets[b]) / len(act_sets[a]) for a, b in zip(wks[:-1], wks[1:])]
    back = []
    for i in range(2, len(wks)):
        prev_all = set().union(*[act_sets[x] for x in wks[:i - 1]])
        back.append(len((act_sets[wks[i]] - act_sets[wks[i - 1]]) & prev_all) / len(act_sets[wks[i]]))
    d['naive'] = {'weekly_loss_mean_last26': float(np.mean(lost[-26:])),
                  'weekly_reactivation_share_last26': float(np.mean(back[-26:]))}

    # ======================================================================================
    # 3. Simple model (version 1): parameters and least-squares equivalent churn
    # ======================================================================================
    b_s, a_s = tr['spend_trend']
    t_jul = (2026 - tr['m0'].year) * 12 + 7 - tr['m0'].month          # month index of July 2026
    spend_jul_month = float(np.exp(a_s + b_s * t_jul))
    free_month = lambda t: sum(float(np.exp(a + b * t)) for (b, a) in
                               [tr['trends']['organic_store'], tr['trends']['podcast_cross_promo']])
    months12 = np.arange(t_jul, t_jul + 12)
    fb, fa = np.polyfit(months12 - t_jul, np.log([free_month(t) for t in months12]), 1)
    v1 = {
        'U0': d['const']['actual_wau_last_full_week'],
        'paid_spend_week1': spend_jul_month * 12 / 52,
        'paid_spend_growth_week': float(np.exp(b_s) ** (12 / 52) - 1),
        'paid_spend_growth_month': float(np.exp(b_s) - 1),
        'spend_jul_month': spend_jul_month,
        'free_week1': float(np.exp(fa)) * 12 / 52,
        'free_growth_week': float(np.exp(fb) ** (12 / 52) - 1),
        'free_growth_month': float(np.exp(fb) - 1),
        'free_jul_month': float(np.exp(fa)),
        'cac': float(tr['cac']),
        'arpu_week': d['const']['arpu_week_data'],
    }
    T = np.arange(52)
    def acquisition(spend_mult=1.0):
        paid = v1['paid_spend_week1'] * (1 + v1['paid_spend_growth_week']) ** T * spend_mult / v1['cac']
        free = v1['free_week1'] * (1 + v1['free_growth_week']) ** T
        return paid, free
    paid, free = acquisition()
    v1['fit'] = {}
    for short in ['Base', 'Pessimistic', 'Optimistic']:
        v1['fit'][short] = fit_churn(v1['U0'], d['scenarios'][short]['WAU'], paid, free)
    # what-if lines: activation / aggressive lines are re-calibrated; spend lines keep the base churn
    v1['whatif'] = {}
    for name, p in d['whatif_paths'].items():
        kw = p['kw']
        mult = kw.get('spend_mult', 1.0)
        pa, fr = acquisition(mult)
        if 'spend_mult' in kw:
            c = v1['fit']['Base']['churn']
            path = simple_path(v1['U0'], c, pa, fr)
            v1['whatif'][name] = {'churn': c, 'recalibrated': False, 'spend_mult': mult, 'path': path.tolist(),
                                  'aggressive': False}
        else:
            f = fit_churn(v1['U0'], p['WAU'], pa, fr)
            v1['whatif'][name] = {'churn': f['churn'], 'recalibrated': True, 'spend_mult': mult, 'path': f['path'],
                                  'aggressive': bool(kw.get('aggressive', False)), 'mape': f['mape']}
    # backtest of the simple model: churn fitted on H2 2025 actual WAU, then forecast Jan-Jun 2026
    hist = pd.DataFrame(d['history']).set_index('weeks')
    fit_weeks = [x for x in hist.index if '2025-07-07' <= x <= '2025-12-22']
    u0_fit = float(hist.loc['2025-06-30', 'WAU'])
    sign = hist.loc[fit_weeks, 'signups_by_signup_week'].values
    fbt = fit_churn(u0_fit, hist.loc[fit_weeks, 'WAU'].values, np.zeros(len(sign)), sign)
    arpu_bt = float(hist.loc[fit_weeks, 'revenue'].sum() / hist.loc[fit_weeks, 'WAU'].sum())
    newA = g['new_bt'].groupby('signup_week').n.sum().reindex(g['bt_weeks']).values
    newB = np.array([hist.loc[x, 'signups_by_signup_week'] for x in d['backtest']['weeks']])
    u0_bt = float(hist.loc['2025-12-22', 'WAU'])
    v1['backtest'] = {'fit_weeks': fit_weeks, 'u0_fit': u0_fit, 'fit_signups': sign.tolist(),
                      'churn': fbt['churn'], 'fit_mape': fbt['mape'], 'arpu': arpu_bt, 'u0': u0_bt,
                      'new_A': newA.tolist(), 'new_B': newB.tolist()}
    for lab, nn in [('A', newA), ('B', newB)]:
        pth = simple_path(u0_bt, fbt['churn'], np.zeros(len(nn)), nn)
        a_w = np.array(d['backtest']['actual_WAU']); a_r = np.array(d['backtest']['actual_revenue'])
        v1['backtest'][f'{lab}_mape_wau'] = float(np.mean(np.abs(pth - a_w) / a_w))
        v1['backtest'][f'{lab}_mape_rev'] = float(np.mean(np.abs(pth * arpu_bt - a_r) / a_r))
    # base forecast of the simple model if it were calibrated on history only (no cohort model)
    v1['history_churn_forecast_end'] = float(simple_path(v1['U0'], fbt['churn'], paid, free)[-1])
    d['v1'] = v1

    Path(args.out).write_text(json.dumps(d, indent=1), encoding='utf-8')
    s = d['scenarios']
    print(f"Exported {args.out}")
    print('Notebook scenarios (WAU end / 12m revenue):', {k: (round(v['WAU'][-1]), round(sum(v['revenue']))) for k, v in s.items()})
    print('Calibration factor:', round(d['const']['CALIB'], 4), '| 12m paid spend:', round(d['paid_spend_12m_base']))
    print('V1 equivalent weekly churn:', {k: round(v['churn'], 5) for k, v in v1['fit'].items()},
          '| MAPE:', {k: f"{v['mape']:.2%}" for k, v in v1['fit'].items()})


if __name__ == '__main__':
    main()
