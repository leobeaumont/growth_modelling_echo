"""
Version 2 "COHORT" workbook: faithful replica of notebook section 9 with live formulas.
Called by build_workbooks.py (do not run on its own).

How the live simulation works (one block = one scenario):
  WAU of new users in week t  =  sum over the cohorts j <= t of  N_g(j) x S_g(t - j)
  (N_g(j) = new users of segment g who signed up in week j, S_g(age) = share still active at that age).
  Excel formula: SUMPRODUCT( N(0..t) , S reversed ), with the curve written in reverse order next to the cohorts
  so that cohort j lines up with age t - j. Revenue uses the revenue curve v_g instead of S_g.
"""
from openpyxl import Workbook
from openpyxl.chart import Reference
from openpyxl.comments import Comment
from openpyxl.utils import get_column_letter as L
from openpyxl.worksheet.datavalidation import DataValidation

from build_workbooks import (CENTER, F_BOLD, F_BODY, F_H2, F_HEAD, F_NOTE, FILL_HEAD, FILL_SEL, FILL_SUB, NF_DATE,
                             NF_EUR, NF_EUR2, NF_EUR4, NF_INT, NF_INT1, NF_PCT, NF_PCT2, NF_PCT3, NF_RATIO, NF_X,
                             REF_LABEL, SCEN, SERIES, BOX, add_series, banner, cell, d2, finish, fml, gap_formats,
                             header, imp, inp, line_chart, note, rc, scenario_dropdown, title, widths)

SEGS = [('organic_store', 0), ('organic_store', 1), ('paid_social', 0), ('paid_social', 1),
        ('podcast_cross_promo', 0), ('podcast_cross_promo', 1)]
SEG_LABEL = {0: 'not activated (< 3 shows)', 1: 'activated (≥ 3 shows)'}
CURVE_FIRST, CURVE_LAST = 6, 175            # Curves rows of ages 0..169
R_FIRST, R_LAST = 6, 21                     # ratio R(a), ages 0..15
EX_FIRST = 9                                # Existing rows of weeks t = 0..51
NU_W_FIRST = 40                             # NewUsers weekly rows t = 0..51
NU_D_FIRST, NU_D_LAST = 96, 457             # NewUsers daily rows (362 days)
NU_ACT = {'organic_store': 34, 'paid_social': 35, 'podcast_cross_promo': 36}
NU_BASE_COL = {'organic_store': 'C', 'podcast_cross_promo': 'D', 'paid_social': 'E'}
BLOCK_H = 76                                # rows per simulation block


def s_col(i):
    return L(2 + 2 * i)


def v_col(i):
    return L(3 + 2 * i)


# --------------------------------------------------------------------------------------------- block
PARAM_ROWS = [('mult', 'Paid spend multiplier'), ('cac', 'CAC (€ / paid user)'), ('upl', 'Activation uplift'),
              ('aggr', 'Aggressive push (1 = yes)'), ('r15', 'Retention ratio after week 15'),
              ('rho', 'Revenue ratio per active week'), ('phi', 'Extra weekly retention factor φ (existing users)'),
              ('phiw', 'Erosion duration for existing users (weeks)'), ('aph', 'Ads per listening hour ratio'),
              ('calib', 'Calibration factor (existing cohorts)'), ('rpa', 'Revenue per ad (€)'),
              ('aph0', 'Ads per listening hour (normal push)')]
PARAM_NF = {'mult': '0.00', 'cac': NF_EUR4, 'upl': NF_PCT, 'aggr': '0', 'r15': NF_RATIO, 'rho': NF_RATIO,
            'phi': NF_RATIO, 'phiw': '0', 'aph': NF_RATIO, 'calib': NF_RATIO, 'rpa': NF_EUR4, 'aph0': '0.00'}
PARAM_NOTE = {
    'mult': '× the paid-spend trend (1 = trend continues)',
    'cac': 'paid users = paid spend ÷ CAC',
    'upl': 'extra share of NEW users who follow ≥ 3 shows in week 1',
    'aggr': 'switch: 1 applies the A/B-test effects below, 0 ignores them',
    'r15': 'new users: retention × measured ratio R(age) up to week 15, × this value after',
    'rho': 'aggressive users hear more ads per active week',
    'phi': 'existing users lose (1 − φ) more of their activity each week ...',
    'phiw': '... during this many weeks after the rollout',
    'aph': 'turns revenue back into listening hours',
    'calib': 'existing cohorts × factor so that week 22 Jun 2026 matches the actual WAU',
    'rpa': 'CPM ÷ 1,000',
    'aph0': 'listening hours = revenue ÷ €/ad ÷ ads per hour',
}
COMMON = {'calib': '=Inputs!$C$19', 'rpa': '=Inputs!$C$21', 'aph0': '=Inputs!$C$22'}


def write_block(ws, top, pref, block_title, notes=False):
    """Write one 52-week simulation block. pref: {param: formula}. Returns refs of the summary cells and columns."""
    cell(ws, f'A{top}', block_title, F_H2)
    P = {}
    for i, (key, lab) in enumerate(PARAM_ROWS):
        r = top + 1 + i
        P[key] = f'$C${r}'
        cell(ws, f'A{r}', lab, F_BODY)
        fml(ws, f'C{r}', pref.get(key, COMMON.get(key)), PARAM_NF[key], green=True)
        if notes:
            note(ws, f'D{r}', PARAM_NOTE[key])
    first, last = top + 19, top + 70
    AL_, AM_, AN_, AO_ = 'AL', 'AM', 'AN', 'AO'
    summ = [('WAU in the last week (21 Jun 2027)', f'={AL_}{last}', NF_INT),
            ('12-month ad revenue (€)', f'=SUM({AM_}{first}:{AM_}{last})', NF_EUR),
            ('Average weekly listening hours', f'=AVERAGE({AN_}{first}:{AN_}{last})', NF_INT),
            ('12-month paid spend (€)', f'=SUM({AO_}{first}:{AO_}{last})', NF_EUR),
            ('Revenue − paid spend (€)', f'=K{top+2}-K{top+4}', NF_EUR)]
    out = {}
    for i, (lab, f, nf) in enumerate(summ):
        r = top + 1 + i
        cell(ws, f'H{r}', lab, F_BOLD)
        fml(ws, f'K{r}', f, nf, bold=True, fill=FILL_SEL)
        out[['wau', 'rev', 'hours', 'spend', 'net'][i]] = f'K{r}'
    # segment weights
    cell(ws, f'A{top+14}', 'Share of new users in this segment', F_BOLD)
    cell(ws, f'A{top+15}', 'New users per unit of the base series', F_BOLD)
    if notes:
        note(ws, f'A{top+16}', 'Base series = organic or cross-promo sign-ups (×1), or paid spend in € (× multiplier ÷ CAC). '
                               'Share = activation rate of the channel + uplift (activated) or 1 − that (not activated).')
    wau_cells, rev_cells = [], []
    for i, (ch, act) in enumerate(SEGS):
        c0 = 3 + 5 * i
        cn, cs, cv, cw, cr = (L(c0 + k) for k in range(5))
        a_share = f'MIN(NewUsers!$B${NU_ACT[ch]}+{P["upl"]},1)'
        fml(ws, f'{cn}{top+14}', f'={a_share}' if act else f'=1-{a_share}', NF_PCT)
        fml(ws, f'{cn}{top+15}', f'=IF({P["cac"]}>0,{P["mult"]}/{P["cac"]},0)' if ch == 'paid_social' else '=1', NF_RATIO)
        rc(ws, top + 17, c0, f'{ch} · {SEG_LABEL[act]}', F_HEAD, fill=FILL_HEAD, align=CENTER)
        ws.merge_cells(start_row=top + 17, start_column=c0, end_row=top + 17, end_column=c0 + 4)
        header(ws, top + 18, c0, ['New users', 'S reversed', 'v reversed (€)', 'WAU', 'Revenue (€)'])
        Sc, Vc = s_col(i), v_col(i)
        for t in range(52):
            r = first + t
            base = f'NewUsers!${NU_BASE_COL[ch]}${NU_W_FIRST + t}'
            fml(ws, f'{cn}{r}', f'={base}*{cn}${top+15}*{cn}${top+14}', NF_INT1)
            ratio = f'IF(51-$A{r}<=15,INDEX(Curves!$O${R_FIRST}:$O${R_LAST},52-$A{r}),{P["r15"]})'
            fml(ws, f'{cs}{r}', f'=INDEX(Curves!${Sc}${CURVE_FIRST}:${Sc}${CURVE_LAST},52-$A{r})*IF({P["aggr"]}=1,{ratio},1)', NF_RATIO)
            fml(ws, f'{cv}{r}', f'=INDEX(Curves!${Vc}${CURVE_FIRST}:${Vc}${CURVE_LAST},52-$A{r})*IF({P["aggr"]}=1,{ratio}*{P["rho"]},1)', NF_EUR4)
            for col_out, col_curve in ((cw, cs), (cr, cv)):
                fml(ws, f'{col_out}{r}',
                    f'=SUMPRODUCT(INDEX(${cn}${first}:${cn}${last},1):INDEX(${cn}${first}:${cn}${last},$A{r}+1),'
                    f'INDEX(${col_curve}${first}:${col_curve}${last},52-$A{r}):INDEX(${col_curve}${first}:${col_curve}${last},52))',
                    NF_INT1 if col_out == cw else NF_EUR2)
        wau_cells.append(cw)
        rev_cells.append(cr)
    tail = ['New users: WAU', 'New users: revenue (€)', 'Existing: factor', 'Existing: WAU', 'Existing: revenue (€)',
            'TOTAL WAU', 'TOTAL revenue (€)', 'Listening hours', 'Paid spend (€)']
    header(ws, top + 18, 33, tail)
    header(ws, top + 18, 1, ['t (week)', 'Week starting'])
    for t in range(52):
        r = first + t
        rc(ws, r, 1, t)
        fml(ws, f'B{r}', f'=NewUsers!$B${NU_W_FIRST + t}', NF_DATE, green=True)
        fml(ws, f'AG{r}', '=' + '+'.join(f'{c}{r}' for c in wau_cells), NF_INT)
        fml(ws, f'AH{r}', '=' + '+'.join(f'{c}{r}' for c in rev_cells), NF_EUR2)
        fml(ws, f'AI{r}', f'=IF({P["aggr"]}=1,{P["phi"]}^MIN($A{r}+1,{P["phiw"]}),1)', NF_RATIO)
        fml(ws, f'AJ{r}', f'=Existing!$C${EX_FIRST + t}*{P["calib"]}*AI{r}', NF_INT)
        fml(ws, f'AK{r}', f'=Existing!$D${EX_FIRST + t}*{P["calib"]}*AI{r}*IF({P["aggr"]}=1,{P["rho"]},1)', NF_EUR2)
        fml(ws, f'AL{r}', f'=AG{r}+AJ{r}', NF_INT, bold=True)
        fml(ws, f'AM{r}', f'=AH{r}+AK{r}', NF_EUR2, bold=True)
        fml(ws, f'AN{r}', f'=IF({P["rpa"]}*{P["aph0"]}>0,AM{r}/{P["rpa"]}/({P["aph0"]}*IF({P["aggr"]}=1,{P["aph"]},1)),0)', NF_INT)
        fml(ws, f'AO{r}', f'=NewUsers!$E${NU_W_FIRST + t}*{P["mult"]}', NF_EUR2)
    out.update({'first': first, 'last': last, 'wau_col': 'AL', 'hours_col': 'AN', 'rev_col': 'AM',
                'exist_col': 'AJ', 'new_col': 'AG'})
    return out


def block_widths(ws):
    ws.column_dimensions['A'].width = 40
    ws.column_dimensions['B'].width = 12
    for c in range(3, 42):
        ws.column_dimensions[L(c)].width = 11


# --------------------------------------------------------------------------------------------- workbook
def build_v2(d, ref, path):
    c = d['const']
    wb = Workbook()
    wb.remove(wb.active)

    # ------------------------------------------------------------------ Inputs
    wi = wb.create_sheet('Inputs')
    title(wi, 'Inputs & scenarios — cohort model (version 2)',
          'Edit only the blue cells on yellow (column Custom and the selector C4). Base / Pessimistic / Optimistic reproduce notebook section 9.3.')
    cell(wi, 'A4', 'Selected scenario', F_BOLD)
    inp(wi, 'C4', 'Base', custom=True, comment='Dropdown: Base / Pessimistic / Optimistic / Custom. The Simulation sheet and the summary below follow this cell. '
        'The Scenarios sheet always shows all four.')
    scenario_dropdown(wi, 'C4', SCEN)
    note(wi, 'D4', '← click the cell, then the small arrow, to switch scenario')
    header(wi, 6, 1, ['Parameter', 'Unit'] + SCEN + ['SELECTED', 'Note / source'], height=30)
    ab = {'r15': d['ratio_measured'][-1], 'rho': c['REV_RATIO'], 'phi': c['PHI'], 'aph': c['APH_RATIO']}
    prm = [
        ('mult', 'Paid spend multiplier', '× the spend trend', {s: 1.0 for s in SCEN}, '0.00', None,
         'Notebook run(spend_mult=...). 1 = the paid budget keeps its trend (+2.8 % per month).'),
        ('cac', 'CAC (paid social)', '€ / paid user', None, NF_EUR4, '=NewUsers!$C$31',
         'Median of monthly spend ÷ paid sign-ups (NewUsers C31, = tr["cac"] in notebook 9.1). Custom can override it.'),
        ('upl', 'Activation uplift', 'extra share of new users ≥ 3 shows', {'Base': 0, 'Pessimistic': 0, 'Optimistic': 0.10, 'Custom': 0}, NF_PCT, None,
         'Optimistic = guided onboarding: +10 percentage points of new users follow 3 shows in week 1 (notebook 9.3).'),
        ('aggr', 'Aggressive push for everyone', '1 = yes, 0 = no', {'Base': 0, 'Pessimistic': 1, 'Optimistic': 0, 'Custom': 0}, '0', None,
         'Pessimistic = the growth team\'s plan: ~5 "Don\'t miss out!" pushes a day for every user (notebook 9.3).'),
        ('r15', 'Retention ratio after week 15', 'aggressive ÷ control', None, NF_RATIO, '=Curves!$O$21',
         f"Measured in the A/B test: retention of the aggressive arm ÷ control at week 15 ({ab['r15']:.3f}). Before week 15 the measured curve "
         'R(age) of sheet Curves is used; after week 15 this value is frozen (notebook NEW_RATIO).'),
        ('rho', 'Revenue ratio per active week', 'aggressive ÷ control', {s: ab['rho'] for s in SCEN}, NF_RATIO, None,
         'A/B test, weeks 2-15: ad revenue per active user-week, aggressive ÷ control (notebook REV_RATIO, section 9.3).'),
        ('phi', 'Extra weekly retention factor φ', 'factor per week', {s: ab['phi'] for s in SCEN}, NF_RATIO, None,
         f"Existing users are multiplied by φ each week of exposure (−{1 - ab['phi']:.1%} per week), measured on weeks 8-15 of the A/B test (notebook PHI)."),
        ('phiw', 'Erosion duration for existing users', 'weeks', {s: 26 for s in SCEN}, '0', None,
         'Assumption of the notebook (not tested): the extra erosion of existing users lasts 26 weeks, then stops. See WhatIf for 0 and 52.'),
        ('aph', 'Ads per listening hour ratio', 'aggressive ÷ control', {s: ab['aph'] for s in SCEN}, NF_RATIO, None,
         'A/B test: 9.1 vs 8.5 ads per listening hour (notebook APH_RATIO). Only used to convert revenue into listening hours.'),
    ]
    IR = {}
    for i, (key, lab, unit, vals, nf, link, src) in enumerate(prm):
        r = 7 + i
        IR[key] = r
        cell(wi, f'A{r}', lab, F_BOLD)
        cell(wi, f'B{r}', unit)
        for j, s in enumerate(SCEN):
            ref_ = f'{L(3 + j)}{r}'
            if link and s != 'Custom':
                fml(wi, ref_, link, nf, green=True)
            elif link:
                inp(wi, ref_, ab['r15'] if key == 'r15' else d['trends']['cac'], nf, custom=True)
            else:
                inp(wi, ref_, vals[s], nf, custom=(s == 'Custom'))
        fml(wi, f'G{r}', f'=INDEX(C{r}:F{r},MATCH($C$4,$C$6:$F$6,0))', nf, bold=True, fill=FILL_SEL)
        note(wi, f'H{r}', src)
        wi[f'C{r}'].comment = Comment(src, 'Echo model')
    dv = DataValidation(type='list', formula1='"0,1"', allow_blank=False)
    wi.add_data_validation(dv)
    dv.add(f'F{IR["aggr"]}')
    cell(wi, 'A17', 'Parameters common to all scenarios', F_H2)
    header(wi, 18, 1, ['Parameter', 'Unit', 'Value'])
    cell(wi, 'A19', 'Calibration factor (existing cohorts)', F_BOLD)
    cell(wi, 'B19', 'actual ÷ model')
    fml(wi, 'C19', '=Existing!$C$6', NF_RATIO, green=True)
    note(wi, 'D19', 'Computed on sheet Existing: actual WAU of 22 Jun 2026 ÷ WAU the curves predict for the same users (notebook CALIB).')
    cell(wi, 'A20', 'CPM', F_BOLD)
    cell(wi, 'B20', '€ per 1,000 ads')
    inp(wi, 'C20', c['CPM'], NF_EUR2, comment='Data pack: Echo earns about €9 per 1,000 ad impressions. FIXED: the revenue curves v of sheet Curves were computed at €9, so this cell is only used to turn revenue into listening hours. To test another CPM, re-export the curves.')
    cell(wi, 'A21', 'Revenue per ad impression', F_BOLD)
    cell(wi, 'B21', '€ / ad')
    fml(wi, 'C21', '=C20/1000', NF_EUR4)
    cell(wi, 'A22', 'Ads per listening hour (normal push)', F_BOLD)
    cell(wi, 'B22', 'ads / hour')
    inp(wi, 'C22', c['ADS_PER_HOUR'], '0.00', comment='Median ads per listening hour on the activity data (notebook section 2.5, ADS_PER_HOUR).')
    note(wi, 'D22', 'Imported from the notebook; only used to express revenue as listening hours (North Star Metric).')
    note(wi, 'A23', 'CPM is fixed (blue, no yellow): the revenue curves v (€ per user) were computed at €9 per 1,000 ads in the notebook.')
    cell(wi, 'A25', 'Results of the selected scenario', F_H2)
    header(wi, 26, 1, ['Output', 'Unit', 'Value'])
    res = [('WAU in the last week (21 Jun 2027)', 'users', '=Simulation!K5', NF_INT),
           ('12-month ad revenue', '€', '=Simulation!K6', NF_EUR),
           ('Average weekly listening hours (North Star)', 'hours / week', '=Simulation!K7', NF_INT),
           ('12-month paid spend', '€', '=Simulation!K8', NF_EUR),
           ('Revenue − paid spend', '€', '=Simulation!K9', NF_EUR)]
    for i, (a, b, f, nf) in enumerate(res):
        r = 27 + i
        cell(wi, f'A{r}', a, F_BOLD)
        cell(wi, f'B{r}', b)
        fml(wi, f'C{r}', f, nf, bold=True, green=True)
    widths(wi, {'A': 44, 'B': 30, 'C': 13, 'D': 13, 'E': 13, 'F': 13, 'G': 13, 'H': 120})
    wi.freeze_panes = 'C7'

    # ------------------------------------------------------------------ Curves (imported)
    wc = wb.create_sheet('Curves')
    title(wc, 'Curves — imported data (do not edit)')
    banner(wc, 'A2', 'IMPORTED DATA. Exported from notebook section 9 (build_curves on all users with normal push, data up to 22 Jun 2026) '
                     'by export_from_notebook.py. S = share of the segment\'s users active in week "age" of their life (age 0 = signup week). '
                     'v = ad revenue (€) per acquired user in that week. Italic grey = extrapolated beyond the observed ages (exponential decay, notebook 2.5).',
           merge_to='Q3')
    wc.row_dimensions[2].height = 45
    for i, (ch, act) in enumerate(SEGS):
        rc(wc, 4, 2 + 2 * i, f'{ch} · {SEG_LABEL[act]}', F_HEAD, fill=FILL_HEAD, align=CENTER)
        wc.merge_cells(start_row=4, start_column=2 + 2 * i, end_row=4, end_column=3 + 2 * i)
    header(wc, 5, 1, ['Age (weeks)'] + ['S (share active)', 'v (€ / user)'] * 6, height=30)
    header(wc, 4, 15, ['A/B test'])
    header(wc, 5, 15, ['R(age) aggressive ÷ control'], height=30)
    for a in range(170):
        r = CURVE_FIRST + a
        imp(wc, f'A{r}', a, '0')
        for i, (ch, act) in enumerate(SEGS):
            cv = d['curves'][f'{ch}|{act}']
            ext = not cv['observed'][a]
            imp(wc, f'{s_col(i)}{r}', cv['S'][a], NF_RATIO, italic=ext)
            imp(wc, f'{v_col(i)}{r}', cv['v'][a], NF_EUR4, italic=ext)
    for a, x in enumerate(d['ratio_measured']):
        imp(wc, f'O{R_FIRST + a}', x, NF_RATIO)
    note(wc, f'O{R_LAST + 1}', 'after week 15: frozen value, Inputs row 11')
    wc['O5'].comment = Comment('Smoothed ratio (3-week centred rolling mean) of the share active, aggressive arm ÷ control arm, ages 0-15 '
                               '(notebook section 9.3, variable ratio).', 'Echo model')
    widths(wc, {'A': 11, **{L(k): 12 for k in range(2, 14)}, 'O': 16})
    wc.freeze_panes = 'B6'

    # ------------------------------------------------------------------ Existing (imported + calibration)
    we = wb.create_sheet('Existing')
    title(we, 'Existing cohorts — contribution of users who signed up before 1 July 2026')
    banner(we, 'A2', 'IMPORTED DATA (columns C to P). Precomputed in Python with the notebook function simulate() on the existing cohorts, '
                     'WITHOUT calibration and WITHOUT scenario effect. The Simulation applies the calibration factor and the push effects on top. '
                     'Why imported: these are ~470 weekly cohorts × 6 segments; the formulas would be the same as for new users, only much bigger.',
           merge_to='Q3')
    we.row_dimensions[2].height = 45
    cell(we, 'A4', 'Actual WAU, week of 22 Jun 2026 (data)', F_BOLD)
    imp(we, 'C4', c['actual_wau_last_full_week'], NF_INT, comment='Notebook section 6, kpi.WAU at LAST_FULL_WEEK.')
    cell(we, 'A5', 'WAU predicted by the curves for the same users, same week', F_BOLD)
    imp(we, 'C5', c['model_existing_wau_last_full_week'], NF_INT1, comment='Notebook section 9.3, nowcast (existing cohorts, calibration 1).')
    cell(we, 'A6', 'Calibration factor = actual ÷ model', F_BOLD)
    fml(we, 'C6', '=C4/C5', NF_RATIO, bold=True, fill=FILL_SEL)
    note(we, 'D6', 'Below 1: the curves (estimated on users with normal push) over-predict, mainly because the aggressive-arm users of the test churned faster.')
    seg_h = [f'WAU {ch[:7]}·{act}' for ch, act in SEGS] + [f'Rev {ch[:7]}·{act} (€)' for ch, act in SEGS]
    header(we, 8, 1, ['t (week)', 'Week starting', 'Existing WAU (uncalibrated)', 'Existing revenue (€, uncalibrated)'] + seg_h + ['Check: segments − total'], height=45)
    ex = d['existing']
    for t in range(52):
        r = EX_FIRST + t
        rc(we, r, 1, t)
        imp(we, f'B{r}', d2(d['weeks'][t]), NF_DATE)
        imp(we, f'C{r}', ex['WAU'][t], NF_INT1)
        imp(we, f'D{r}', ex['revenue'][t], NF_EUR2)
        for i, (ch, act) in enumerate(SEGS):
            seg = ex['by_segment'][f'{ch}|{act}']
            imp(we, f'{L(5 + i)}{r}', seg['WAU'][t], NF_INT1)
            imp(we, f'{L(11 + i)}{r}', seg['revenue'][t], NF_EUR2)
        fml(we, f'Q{r}', f'=SUM(E{r}:J{r})-C{r}', '0.000')
    widths(we, {'A': 52, 'B': 13, 'C': 15, 'D': 15, **{L(k): 12 for k in range(5, 17)}, 'Q': 14})
    we.freeze_panes = 'C9'

    # ------------------------------------------------------------------ NewUsers
    wn = wb.create_sheet('NewUsers')
    title(wn, 'New users of the forecast — computed by formulas from the 18-month history')
    note(wn, 'A2', 'Same method as notebook fit_trends / forecast_new: log-linear trend of monthly sign-ups and spend, spread day by day, summed by week.')
    cell(wn, 'A4', 'A. Monthly history (imported from the cleaned data, notebook sections 2 and 9.1)', F_H2)
    header(wn, 5, 1, ['Month', 'Month index', 'Organic sign-ups', 'Paid sign-ups', 'Cross-promo sign-ups', 'Paid spend (€)',
                      'LN organic', 'LN cross-promo', 'LN spend', 'CAC of the month (€)'], height=30)
    mh = d['monthly_history']
    for i, mth in enumerate(mh['month']):
        r = 6 + i
        imp(wn, f'A{r}', d2(mth), 'mmm yyyy')
        fml(wn, f'B{r}', f'=(YEAR(A{r})-2025)*12+MONTH(A{r})-1', '0')
        imp(wn, f'C{r}', mh['organic_store'][i], NF_INT)
        imp(wn, f'D{r}', mh['paid_social'][i], NF_INT)
        imp(wn, f'E{r}', mh['podcast_cross_promo'][i], NF_INT)
        if mh['spend_raw'][i] is None:
            inp(wn, f'F{r}', mh['spend'][i], NF_EUR2, custom=True,
                comment=f"May 2025 is MISSING in marketing_spend.csv. The current notebook fills it with the column mean "
                        f"(€{mh['spend'][i]:,.2f}), which gives a CAC of €{mh['spend'][i] / mh['paid_social'][i]:.2f} that month instead of €2.61. "
                        "Linear interpolation (€1,115) or CAC × paid sign-ups (€1,112) fits the series better. Type 1115 here to see the effect.")
        else:
            imp(wn, f'F{r}', mh['spend'][i], NF_EUR2)
        fml(wn, f'G{r}', f'=LN(C{r})', '0.0000')
        fml(wn, f'H{r}', f'=LN(E{r})', '0.0000')
        fml(wn, f'I{r}', f'=LN(F{r})', '0.0000')
        fml(wn, f'J{r}', f'=F{r}/D{r}', NF_EUR4)
    cell(wn, 'A25', 'B. Trends fitted on the history (log-linear least squares = numpy.polyfit in the notebook)', F_H2)
    header(wn, 26, 1, ['Series', 'Slope (per month)', 'Intercept', 'Monthly growth'])
    for r, lab, col in [(27, 'Organic sign-ups', 'G'), (28, 'Cross-promo sign-ups', 'H'), (29, 'Paid spend (€)', 'I')]:
        cell(wn, f'A{r}', lab, F_BOLD)
        fml(wn, f'B{r}', f'=SLOPE({col}6:{col}23,$B$6:$B$23)', '0.000000')
        fml(wn, f'C{r}', f'=INTERCEPT({col}6:{col}23,$B$6:$B$23)', '0.000000')
        fml(wn, f'D{r}', f'=EXP(B{r})-1', NF_PCT2)
    cell(wn, 'A31', 'CAC (median of the monthly CAC)', F_BOLD)
    fml(wn, 'C31', '=MEDIAN(J6:J23)', NF_EUR4, bold=True, fill=FILL_SEL)
    cell(wn, 'A33', 'Activation rate by channel (share of new users who follow ≥ 3 shows in week 1, normal push)', F_H2)
    for ch, r in NU_ACT.items():
        cell(wn, f'A{r}', ch, F_BOLD)
        imp(wn, f'B{r}', d['trends']['act'][ch], NF_PCT, comment='Imported: notebook tr["act"] (users signed up up to June 2026, excluding the aggressive arm).')
    cell(wn, 'A38', 'C. Weekly new users of the forecast (Base: spend multiplier 1, before the split into segments)', F_H2)
    header(wn, 39, 1, ['t (week)', 'Week starting', 'Organic sign-ups', 'Cross-promo sign-ups', 'Paid spend (€, trend)', 'Paid sign-ups at Base CAC'], height=30)
    for t in range(52):
        r = NU_W_FIRST + t
        rc(wn, r, 1, t)
        fml(wn, f'B{r}', '=DATE(2026,6,29)' if t == 0 else f'=B{r-1}+7', NF_DATE)
        for col, dcol in (('C', 'E'), ('D', 'F'), ('E', 'G')):
            fml(wn, f'{col}{r}', f'=SUMIFS(${dcol}${NU_D_FIRST}:${dcol}${NU_D_LAST},$B${NU_D_FIRST}:$B${NU_D_LAST},$B{r})', NF_INT1 if col != 'E' else NF_EUR2)
        fml(wn, f'F{r}', f'=IF(Inputs!$C$8>0,E{r}/Inputs!$C$8,0)', NF_INT1)
    rT = NU_W_FIRST + 52
    cell(wn, f'A{rT}', 'Total 52 weeks', F_BOLD)
    for col in 'CDEF':
        fml(wn, f'{col}{rT}', f'=SUM({col}{NU_W_FIRST}:{col}{rT-1})', NF_INT if col != 'E' else NF_EUR, bold=True)
    note(wn, f'A{rT+1}', 'Week 0 (29 Jun 2026) only contains 1-5 July: the users of 29-30 June are already in the data (sheet Existing).')
    cell(wn, 'A94', 'D. Daily forecast (helper for table C): monthly trend value ÷ days in the month', F_H2)
    header(wn, 95, 1, ['Date', 'Week starting', 'Month index', 'Days in month', 'Organic / day', 'Cross-promo / day', 'Paid spend / day (€)'], height=30)
    for k in range(NU_D_LAST - NU_D_FIRST + 1):
        r = NU_D_FIRST + k
        fml(wn, f'A{r}', '=DATE(2026,7,1)' if k == 0 else f'=A{r-1}+1', NF_DATE)
        fml(wn, f'B{r}', f'=A{r}-WEEKDAY(A{r},2)+1', NF_DATE)
        fml(wn, f'C{r}', f'=(YEAR(A{r})-2025)*12+MONTH(A{r})-1', '0')
        fml(wn, f'D{r}', f'=DAY(EOMONTH(A{r},0))', '0')
        fml(wn, f'E{r}', f'=EXP($C$27+$B$27*C{r})/D{r}', '0.000')
        fml(wn, f'F{r}', f'=EXP($C$28+$B$28*C{r})/D{r}', '0.000')
        fml(wn, f'G{r}', f'=EXP($C$29+$B$29*C{r})/D{r}', NF_EUR2)
    widths(wn, {'A': 44, 'B': 14, 'C': 14, 'D': 14, 'E': 15, 'F': 15, 'G': 14, 'H': 14, 'I': 12, 'J': 14})

    # ------------------------------------------------------------------ Simulation (selected scenario)
    wsim = wb.create_sheet('Simulation')
    title(wsim, 'Simulation of the selected scenario (52 weeks, July 2026 → June 2027)')
    fml(wsim, 'A2', '="Scenario: "&Inputs!C4', green=True, bold=True)
    note(wsim, 'D2', 'For each segment: WAU(t) = Σ over cohorts j ≤ t of New users(j) × S(t − j). Excel: SUMPRODUCT of the new-user column '
                     '(weeks 0..t) with the reversed curve (ages t..0). Then + existing cohorts × calibration × push factor.')
    sel = {k: f'=Inputs!$G${r}' for k, r in IR.items()}
    SIM = write_block(wsim, 4, sel, 'Parameters (selected scenario) and results', notes=True)
    block_widths(wsim)
    wsim.column_dimensions['D'].width = 11
    wsim.freeze_panes = f'C{SIM["first"]}'
    ch = line_chart('Selected scenario: weekly active users', 'users', None, 24, 10)
    add_series(ch, wsim, 38, SIM['first'], 38, SIM['last'], 'Total WAU', SERIES['Base'])
    add_series(ch, wsim, 36, SIM['first'], 36, SIM['last'], 'Existing users (calibrated)', SERIES['Custom'])
    add_series(ch, wsim, 33, SIM['first'], 33, SIM['last'], 'New users (July 2026 onwards)', SERIES['Optimistic'])
    ch.set_categories(Reference(wsim, min_col=2, min_row=SIM['first'], max_row=SIM['last']))
    wsim.add_chart(ch, f'A{SIM["last"] + 3}')

    # ------------------------------------------------------------------ WhatIf lever table (needed by Engine)
    ww = wb.create_sheet('WhatIf')
    title(ww, 'What-if: one lever at a time (notebook section 9.4, course slide 98)',
          'Each line = Base with ONE lever changed (blue). Aggressive lines take the measured A/B parameters from the Pessimistic column of Inputs (green).')
    lever_cols = ['Line', 'Spend multiplier', 'CAC (€)', 'Activation uplift', 'Aggressive (1/0)', 'Ratio after wk 15',
                  'Revenue ratio', 'φ', 'Erosion weeks', 'Ads/hour ratio']
    header(ww, 4, 1, lever_cols, height=30)
    keys = ['mult', 'cac', 'upl', 'aggr', 'r15', 'rho', 'phi', 'phiw', 'aph']
    nfs = ['0.00', NF_EUR4, NF_PCT, '0', NF_RATIO, NF_RATIO, NF_RATIO, '0', NF_RATIO]
    WI = list(d['whatif'].keys())
    WI_CHANGE = {
        'Base': {},
        'Aggressive push, new users only': {'aggr': 1, 'phiw': 0},
        'Aggressive push, existing users erode 26 weeks': {'aggr': 1, 'phiw': 26},
        'Aggressive push, existing users erode 52 weeks': {'aggr': 1, 'phiw': 52},
        'Activation +5 pp': {'upl': 0.05},
        'Activation +10 pp': {'upl': 0.10},
        'Paid social spend x 1.5': {'mult': 1.5},
        'Paid social stopped': {'mult': 0.0},
    }
    WI_ROW = {}
    for i, name in enumerate(WI):
        r = 5 + i
        WI_ROW[name] = r
        ch_ = WI_CHANGE[name]
        cell(ww, f'A{r}', name, F_BOLD)
        for j, k in enumerate(keys):
            ref_ = f'{L(2 + j)}{r}'
            if k in ch_:
                inp(ww, ref_, ch_[k], nfs[j])
            elif ch_.get('aggr') == 1 and k in ('r15', 'rho', 'phi', 'aph'):
                fml(ww, ref_, f'=Inputs!$D${IR[k]}', nfs[j], green=True)
            else:
                fml(ww, ref_, f'=Inputs!$C${IR[k]}', nfs[j], green=True)

    # ------------------------------------------------------------------ Engine (all scenarios + what-if)
    wen = wb.create_sheet('Engine')
    title(wen, 'Engine — the same simulation block, once per scenario and per what-if line (technical sheet, do not edit)',
          'Used by the sheets Scenarios and WhatIf so that all scenarios are live at the same time. Each block is identical to the Simulation sheet; '
          'only its parameter column differs.')
    ENG = {}
    configs = [(s, f'Scenario: {s}', {k: f'=Inputs!${L(3 + i)}${r}' for k, r in IR.items()}) for i, s in enumerate(SCEN)]
    configs += [(f'WI|{name}', f'What-if: {name}', {k: f'=WhatIf!${L(2 + j)}${WI_ROW[name]}' for j, k in enumerate(keys)}) for name in WI]
    for i, (key, lab, pref) in enumerate(configs):
        ENG[key] = write_block(wen, 4 + i * BLOCK_H, pref, lab)
    block_widths(wen)

    # WhatIf results + activation value
    cell(ww, 'A14', 'Results (live, from sheet Engine)', F_H2)
    header(ww, 15, 1, ['Line', 'WAU June 2027', '12-month revenue (€)', '12-month paid spend (€)', 'Revenue − spend (€)',
                       'Δ (revenue − spend) vs Base (€)'], height=30)
    WIRES = {}
    for i, name in enumerate(WI):
        r = 16 + i
        e = ENG[f'WI|{name}']
        cell(ww, f'A{r}', name, F_BOLD)
        fml(ww, f'B{r}', f'=Engine!{e["wau"]}', NF_INT, green=True)
        fml(ww, f'C{r}', f'=Engine!{e["rev"]}', NF_EUR, green=True)
        fml(ww, f'D{r}', f'=Engine!{e["spend"]}', NF_EUR, green=True)
        fml(ww, f'E{r}', f'=C{r}-D{r}', NF_EUR)
        fml(ww, f'F{r}', f'=E{r}-$E$16', NF_EUR)
        WIRES[name] = r
    cell(ww, 'A26', 'Value of activation: 2-year LTV per acquired user (sum of the v curve over ages 0–103)', F_H2)
    header(ww, 27, 1, ['Channel', 'Not activated (€)', 'Activated ≥ 3 shows (€)', 'Value of activating one user (€)',
                       'Activated LTV ÷ paid CAC'], height=30)
    LTVR = {}
    for i, ch in enumerate(['organic_store', 'paid_social', 'podcast_cross_promo']):
        r = 28 + i
        i0, i1 = SEGS.index((ch, 0)), SEGS.index((ch, 1))
        cell(ww, f'A{r}', ch, F_BOLD)
        fml(ww, f'B{r}', f'=SUM(Curves!{v_col(i0)}{CURVE_FIRST}:{v_col(i0)}{CURVE_FIRST + 103})', NF_EUR2)
        fml(ww, f'C{r}', f'=SUM(Curves!{v_col(i1)}{CURVE_FIRST}:{v_col(i1)}{CURVE_FIRST + 103})', NF_EUR2)
        fml(ww, f'D{r}', f'=C{r}-B{r}', NF_EUR2)
        fml(ww, f'E{r}', f'=IF(Inputs!$C$8>0,C{r}/Inputs!$C$8,0)', NF_X)
        LTVR[ch] = r
    note(ww, 'A32', 'Reading: an activated paid user is worth more than its CAC (ratio > 1), a non-activated one is not: paid social can be fixed by onboarding.')
    widths(ww, {'A': 46, **{L(k): 15 for k in range(2, 11)}})

    # ------------------------------------------------------------------ Scenarios
    wsc = wb.create_sheet('Scenarios')
    title(wsc, 'Scenarios side by side (live)', 'All four scenarios are computed at the same time (sheet Engine). Change a value on Inputs and they update.')
    header(wsc, 4, 1, ['Scenario', 'WAU June 2027', 'Avg weekly listening hours', '12-month ad revenue (€)', '12-month paid spend (€)',
                       'Revenue − paid spend (€)', 'Revenue vs Base', 'WAU vs Base'], height=30)
    for i, s in enumerate(SCEN):
        r = 5 + i
        e = ENG[s]
        cell(wsc, f'A{r}', s, F_BOLD)
        fml(wsc, f'B{r}', f'=Engine!{e["wau"]}', NF_INT, green=True)
        fml(wsc, f'C{r}', f'=Engine!{e["hours"]}', NF_INT, green=True)
        fml(wsc, f'D{r}', f'=Engine!{e["rev"]}', NF_EUR, green=True)
        fml(wsc, f'E{r}', f'=Engine!{e["spend"]}', NF_EUR, green=True)
        fml(wsc, f'F{r}', f'=D{r}-E{r}', NF_EUR)
        fml(wsc, f'G{r}', f'=IF($D$5>0,D{r}/$D$5-1,0)', NF_PCT)
        fml(wsc, f'H{r}', f'=IF($B$5>0,B{r}/$B$5-1,0)', NF_PCT)
    cell(wsc, 'A10', 'Weekly series (for the charts)', F_H2)
    header(wsc, 11, 1, ['Week starting', 'Actual WAU', 'Base', 'Pessimistic', 'Optimistic', 'Custom',
                        'Actual listening hours', 'Base (h)', 'Pessimistic (h)', 'Optimistic (h)', 'Custom (h)'], height=30)
    hist = d['history']
    h_idx = [i for i, w in enumerate(hist['weeks']) if w >= '2025-07-07']
    r = 12
    for i in h_idx:
        imp(wsc, f'A{r}', d2(hist['weeks'][i]), NF_DATE)
        imp(wsc, f'B{r}', hist['WAU'][i], NF_INT)
        imp(wsc, f'G{r}', hist['hours'][i], NF_INT)
        r += 1
    fs = r
    for t in range(52):
        rr = fs + t
        fml(wsc, f'A{rr}', f'=NewUsers!$B${NU_W_FIRST + t}', NF_DATE, green=True)
        for j, s in enumerate(SCEN):
            e = ENG[s]
            fml(wsc, f'{L(3 + j)}{rr}', f'=Engine!{e["wau_col"]}{e["first"] + t}', NF_INT, green=True)
            fml(wsc, f'{L(8 + j)}{rr}', f'=Engine!{e["hours_col"]}{e["first"] + t}', NF_INT, green=True)
    last = fs + 51
    widths(wsc, {'A': 14, **{L(k): 14 for k in range(2, 12)}})
    wsc.freeze_panes = 'B12'
    for k, (lab, ycol0, ytitle, anchor) in enumerate([('Weekly active users: actual and scenarios', 2, 'users', 'M4'),
                                                      ('Weekly listening hours (North Star): actual and scenarios', 7, 'hours / week', 'M27')]):
        ch = line_chart(lab, ytitle, None, 24, 11)
        add_series(ch, wsc, ycol0, 12, ycol0, last, 'Actual', SERIES['Actual'])
        for j, s in enumerate(['Base', 'Pessimistic', 'Optimistic']):
            add_series(ch, wsc, ycol0 + 1 + j, 12, ycol0 + 1 + j, last, s, SERIES[s])
        ch.set_categories(Reference(wsc, min_col=1, min_row=12, max_row=last))
        wsc.add_chart(ch, anchor)

    # ------------------------------------------------------------------ Backtest (static)
    wb_ = wb.create_sheet('Backtest')
    title(wb_, 'Backtest (notebook section 9.2): calibrate on data up to 28 Dec 2025, predict Jan–Jun 2026')
    banner(wb_, 'A2', 'STATIC VALUES exported from the notebook. The backtest model uses curves estimated with data up to 28 Dec 2025; those curves are not '
                      'in this workbook, so these numbers do NOT change when you edit Inputs. Only the error columns are formulas. '
                      'A = new sign-ups forecast from the 2025 trend; B = actual sign-ups.', merge_to='K3')
    wb_.row_dimensions[2].height = 45
    header(wb_, 5, 1, ['Week starting', 'Actual WAU', 'Actual revenue (€)', 'A: WAU', 'A: revenue (€)', 'B: WAU', 'B: revenue (€)',
                       'A: WAU error', 'A: revenue error', 'B: WAU error', 'B: revenue error'], height=30)
    bt = d['backtest']
    n = len(bt['weeks'])
    for i in range(n):
        r = 6 + i
        imp(wb_, f'A{r}', d2(bt['weeks'][i]), NF_DATE)
        for col, key, nf in [('B', 'actual_WAU', NF_INT), ('C', 'actual_revenue', NF_EUR2), ('D', 'A_WAU', NF_INT),
                             ('E', 'A_revenue', NF_EUR2), ('F', 'B_WAU', NF_INT), ('G', 'B_revenue', NF_EUR2)]:
            imp(wb_, f'{col}{r}', bt[key][i], nf)
        fml(wb_, f'H{r}', f'=D{r}/B{r}-1', NF_PCT2)
        fml(wb_, f'I{r}', f'=E{r}/C{r}-1', NF_PCT2)
        fml(wb_, f'J{r}', f'=F{r}/B{r}-1', NF_PCT2)
        fml(wb_, f'K{r}', f'=G{r}/C{r}-1', NF_PCT2)
    rl = 5 + n
    s0 = rl + 2
    header(wb_, s0, 1, ['Metric', 'A: forecast sign-ups', 'B: actual sign-ups'], height=30)
    BTR = {}
    for i, (lab, fa, fb) in enumerate([
            ('MAPE WAU', f'=SUMPRODUCT(ABS(H6:H{rl}))/{n}', f'=SUMPRODUCT(ABS(J6:J{rl}))/{n}'),
            ('MAPE revenue', f'=SUMPRODUCT(ABS(I6:I{rl}))/{n}', f'=SUMPRODUCT(ABS(K6:K{rl}))/{n}'),
            ('bias WAU, last week', f'=H{rl}', f'=J{rl}'),
            ('bias 26-week revenue', f'=SUM(E6:E{rl})/SUM(C6:C{rl})-1', f'=SUM(G6:G{rl})/SUM(C6:C{rl})-1')]):
        r = s0 + 1 + i
        cell(wb_, f'A{r}', lab, F_BOLD)
        fml(wb_, f'B{r}', fa, NF_PCT2, bold=True)
        fml(wb_, f'C{r}', fb, NF_PCT2, bold=True)
        BTR[lab] = r
    widths(wb_, {'A': 22, **{L(k): 14 for k in range(2, 12)}})
    wb_.freeze_panes = 'B6'
    ch = line_chart('Backtest: actual vs predicted weekly active users', 'users', None, 22, 10)
    add_series(ch, wb_, 2, 6, 2, rl, 'Actual', SERIES['Actual'])
    add_series(ch, wb_, 4, 6, 4, rl, 'A: forecast sign-ups', SERIES['Base'])
    add_series(ch, wb_, 6, 6, 6, rl, 'B: actual sign-ups', SERIES['Pessimistic'])
    ch.set_categories(Reference(wb_, min_col=1, min_row=6, max_row=rl))
    wb_.add_chart(ch, 'M5')

    # ------------------------------------------------------------------ Checks
    wk = wb.create_sheet('Checks')
    title(wk, 'Checks: cohort model vs notebook',
          'Gap = this workbook ÷ reference − 1. Target: |gap vs notebook| ≤ 0.5 % (green). The column "original run" shows the numbers of your brief: '
          'the 0.4–1.3 % gaps there come from two cleaning changes made in the notebook since (see README).')
    header(wk, 4, 1, ['Metric', 'Scenario / line', 'Notebook (current version)', 'This workbook (live)', 'Gap vs notebook', 'Status',
                      REF_LABEL, 'Gap vs brief', 'Comment'], height=45)
    rows = []
    for s in ['Pessimistic', 'Base', 'Optimistic']:
        rows.append(('WAU June 2027', s, d['scenarios'][s]['WAU'][-1], f'=Scenarios!B{5 + SCEN.index(s)}', ref['scenarios'][s]['WAU_end'], NF_INT, ''))
    for s in ['Pessimistic', 'Base', 'Optimistic']:
        rows.append(('12-month ad revenue (€)', s, sum(d['scenarios'][s]['revenue']), f'=Scenarios!D{5 + SCEN.index(s)}',
                     ref['scenarios'][s]['revenue_12m'], NF_EUR, ''))
    for s in ['Pessimistic', 'Base', 'Optimistic']:
        rows.append(('Average weekly listening hours', s, sum(d['scenarios'][s]['hours']) / 52, f'=Scenarios!C{5 + SCEN.index(s)}',
                     ref['scenarios'][s]['hours_avg'], NF_INT, ''))
    rows.append(('12-month paid spend (€)', 'all scenarios', d['scenarios']['Base']['paid_spend_12m'], '=Scenarios!E5',
                 ref['scenarios']['Base']['paid_spend_12m'], NF_EUR, 'Brief gap: May 2025 spend imputed with the column mean in the current notebook'))
    rows.append(('Calibration factor (existing cohorts)', '—', c['CALIB'], '=Existing!C6', ref['CALIB'], NF_RATIO, ''))
    for name in WI:
        rows.append(('What-if: WAU June 2027', name, d['whatif'][name]['WAU June 2027'], f'=WhatIf!B{WIRES[name]}',
                     ref['whatif'][name]['WAU June 2027'], NF_INT, ''))
    for name in WI:
        rows.append(('What-if: 12-month revenue (€)', name, d['whatif'][name]['12-month revenue'], f'=WhatIf!C{WIRES[name]}',
                     ref['whatif'][name]['12-month revenue'], NF_EUR, ''))
    for name in WI:
        rows.append(('What-if: 12-month paid spend (€)', name, d['whatif'][name]['12-month paid spend'], f'=WhatIf!D{WIRES[name]}',
                     ref['whatif'][name]['12-month paid spend'], NF_EUR, ''))
    for name in WI:
        rows.append(('What-if: revenue − paid spend (€)', name, d['whatif'][name]['revenue - spend'], f'=WhatIf!E{WIRES[name]}',
                     ref['whatif'][name]['revenue - spend'], NF_EUR, ''))
    for ch in ['organic_store', 'paid_social', 'podcast_cross_promo']:
        for col, key in [('B', 'not activated'), ('C', 'activated (>= 3 shows)'), ('D', 'value of activating one user')]:
            rows.append((f'Activation value: {key} (€)', ch, d['ltv_seg'][ch][key], f'=WhatIf!{col}{LTVR[ch]}', ref['ltv_seg'][ch][key], NF_EUR2, ''))
    for lab, key, sub, col in [('Backtest A: MAPE WAU', 'A: forecast signups', 'MAPE WAU', 'B'), ('Backtest A: MAPE revenue', 'A: forecast signups', 'MAPE revenue', 'B'),
                               ('Backtest B: MAPE WAU', 'B: actual signups', 'MAPE WAU', 'C'), ('Backtest B: MAPE revenue', 'B: actual signups', 'MAPE revenue', 'C')]:
        rows.append((lab, 'Jan–Jun 2026', d['backtest']['table'][key][sub], f'=Backtest!{col}{BTR[sub]}', ref['backtest'][key][sub], NF_PCT2,
                     'Static values (not recomputed by the workbook)'))
    sel_row = 5 + 0
    rows_internal = [
        ('Internal: Simulation = Engine block of the selected scenario (WAU)', '=INDEX(Scenarios!B5:B8,MATCH(Inputs!C4,Scenarios!A5:A8,0))', '=Simulation!K5', NF_INT),
        ('Internal: Simulation = Engine block of the selected scenario (revenue)', '=INDEX(Scenarios!D5:D8,MATCH(Inputs!C4,Scenarios!A5:A8,0))', '=Simulation!K6', NF_EUR),
        ('Internal: existing segments add up to the total (max abs. difference)', '0', '=MAX(MAX(Existing!Q9:Q60),-MIN(Existing!Q9:Q60))', '0.000000'),
        ('Internal: 52-week new users = notebook forecast (organic)', sum(d['new_users_base_weekly']['organic_store']), f'=NewUsers!C{rT}', NF_INT1),
        ('Internal: 52-week new users = notebook forecast (cross-promo)', sum(d['new_users_base_weekly']['podcast_cross_promo']), f'=NewUsers!D{rT}', NF_INT1),
        ('Internal: 52-week new users = notebook forecast (paid, Base)', sum(d['new_users_base_weekly']['paid_social']), f'=NewUsers!F{rT}', NF_INT1),
    ]
    for i, (m, sline, nbv, wbf, refv, nf, cm) in enumerate(rows):
        r = 5 + i
        cell(wk, f'A{r}', m, F_BOLD)
        cell(wk, f'B{r}', sline)
        imp(wk, f'C{r}', nbv, nf)
        fml(wk, f'D{r}', wbf, nf, green=True)
        fml(wk, f'E{r}', f'=IF(C{r}=0,D{r}-C{r},D{r}/C{r}-1)', NF_PCT2)
        fml(wk, f'F{r}', f'=IF(ABS(E{r})<=0.005,"OK","CHECK")')
        imp(wk, f'G{r}', refv, nf)
        fml(wk, f'H{r}', f'=IF(G{r}=0,D{r}-G{r},D{r}/G{r}-1)', NF_PCT2)
        note(wk, f'I{r}', cm)
    last = 4 + len(rows)
    r0 = last + 2
    cell(wk, f'A{r0}', 'Internal consistency checks', F_H2)
    header(wk, r0 + 1, 1, ['Check', '', 'Expected', 'This workbook', 'Gap', 'Status'])
    for i, (m, exp_, got, nf) in enumerate(rows_internal):
        r = r0 + 2 + i
        cell(wk, f'A{r}', m, F_BOLD)
        if isinstance(exp_, str) and exp_.startswith('='):
            fml(wk, f'C{r}', exp_, nf, green=True)
        elif exp_ == '0':
            fml(wk, f'C{r}', '=0', nf)
        else:
            imp(wk, f'C{r}', exp_, nf)
        fml(wk, f'D{r}', got, nf, green=True)
        fml(wk, f'E{r}', f'=IF(C{r}=0,D{r}-C{r},D{r}/C{r}-1)', '0.0000%')
        fml(wk, f'F{r}', f'=IF(ABS(E{r})<=0.005,"OK","CHECK")')
    r_end = r0 + 1 + len(rows_internal)
    cell(wk, f'A{r_end + 2}', 'Number of checks flagged CHECK (target 0)', F_BOLD)
    fml(wk, f'D{r_end + 2}', f'=COUNTIF(F5:F{last},"CHECK")+COUNTIF(F{r0+2}:F{r_end},"CHECK")', '0', bold=True, fill=FILL_SEL)
    wk['C4'].comment = Comment('Imported from the CURRENT notebook (group version, git HEAD) by export_from_notebook.py.', 'Echo model')
    wk['G4'].comment = Comment('Imported from reference_original_run.json: the original notebook run (commit b88fe6a), whose numbers are quoted in the brief.', 'Echo model')
    gap_formats(wk, f'E5:E{last}')
    gap_formats(wk, f'H5:H{last}')
    gap_formats(wk, f'E{r0+2}:E{r_end}')
    widths(wk, {'A': 62, 'B': 44, 'C': 16, 'D': 16, 'E': 11, 'F': 9, 'G': 20, 'H': 11, 'I': 80})
    wk.freeze_panes = 'C5'

    # ------------------------------------------------------------------ Guide
    wg = wb.create_sheet('Guide', 0)
    from build_workbooks import F_TITLE
    lines = [
        ('Echo growth model — Version 2 "COHORT" (faithful to notebook section 9)', F_TITLE),
        ('', None),
        ('WHAT IT ANSWERS', F_BOLD),
        ('Board question: should Echo roll out aggressive push notifications to everyone? This file recomputes the notebook\'s cohort model', F_BODY),
        ('with live Excel formulas: 3 scenarios (Pessimistic = aggressive push for everyone, Base = status quo, Optimistic = guided onboarding) + 8 what-if lines.', F_BODY),
        ('', None),
        ('HOW IT WORKS', F_BOLD),
        ('Users are grouped in 6 segments (3 acquisition channels × activated or not) and in weekly sign-up cohorts.', F_BODY),
        ('WAU(t) = existing cohorts × calibration factor × push factor  +  Σ over segments Σ over new cohorts  New users(cohort) × S_segment(age)', F_BODY),
        ('Revenue(t): same sum with the revenue curve v instead of S.   Listening hours = revenue ÷ €0.009 per ad ÷ ads per listening hour.', F_BODY),
        ('S (share still active at each age) and v (€ per user at each age) come from the data (sheet Curves). New users come from the trends (sheet NewUsers).', F_NOTE),
        ('', None),
        ('COLOR CODE', F_BOLD),
        ('Blue text on yellow = inputs you can change (Inputs column Custom, selector Inputs!C4, the missing May-2025 spend in NewUsers F10).', F_BODY),
        ('Blue text without yellow = scenario values taken from the notebook.   Black = formulas.   Green = values pulled from another sheet.', F_BODY),
        ('Dark grey on light grey = IMPORTED DATA (static values exported from Python by export_from_notebook.py). Italic = extrapolated.', F_BODY),
        ('', None),
        ('SHEETS (in reading order)', F_BOLD),
        ('Inputs — scenario selector (C4), parameters of the 4 scenarios, common parameters, results of the selected scenario.', F_BODY),
        ('Scenarios — the 4 scenarios side by side + charts of WAU and listening hours (all live).', F_BODY),
        ('WhatIf — the 8 one-lever lines of notebook 9.4 + the value of activating a user.', F_BODY),
        ('Simulation — the full week-by-week calculation of the selected scenario (the professor\'s "Model" sheet, by cohort).', F_BODY),
        ('NewUsers — history, fitted trends, daily then weekly new users of the forecast (formulas).', F_BODY),
        ('Curves — imported retention S and revenue v curves by segment and age, and the A/B ratio R(age).', F_BODY),
        ('Existing — imported contribution of the users already in the base on 30 June 2026, and the calibration factor.', F_BODY),
        ('Engine — technical copies of the Simulation block (one per scenario and what-if line). Do not edit.', F_BODY),
        ('Backtest — notebook 9.2 backtest, static values and error formulas.', F_BODY),
        ('Checks — every output compared with the notebook (target ≤ 0.5 %) and with the original run quoted in the brief.', F_BODY),
        ('', None),
        ('Rebuild: python export_from_notebook.py, then python build_workbooks.py (see README_v2_cohort.md).', F_NOTE),
    ]
    for i, (txt, f) in enumerate(lines):
        if f:
            cell(wg, f'B{i+1}', txt, f)
    widths(wg, {'A': 3, 'B': 150})
    order = ['Guide', 'Inputs', 'Scenarios', 'WhatIf', 'Simulation', 'NewUsers', 'Curves', 'Existing', 'Engine', 'Backtest', 'Checks']
    wb._sheets = [wb[n] for n in order]
    finish(wb, path)
