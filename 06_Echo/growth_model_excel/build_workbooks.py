"""
Builds the two Echo growth-model workbooks from model_data.json (written by export_from_notebook.py).

    python export_from_notebook.py     # 1. run the notebook, export the data
    python build_workbooks.py          # 2. write echo_growth_model_v1_simple.xlsx and echo_growth_model_v2_cohort.xlsx

Then open the files in Excel (they recalculate on opening) or run LibreOffice once to store the computed values.
Conventions copied from the professor's 'IMT - Growth Model Simulator.xlsx':
  blue text = inputs (blue on yellow = the cells you are meant to edit), black = formulas,
  green = values pulled from another sheet. Added here: dark grey on light grey = imported data (static values from Python).
"""
import json
from datetime import date, datetime, timedelta
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import LineChart, Reference, Series
from openpyxl.comments import Comment
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.workbook.properties import CalcProperties
from openpyxl.worksheet.datavalidation import DataValidation

HERE = Path(__file__).resolve().parent

# ----------------------------------------------------------------------------- styles
NAVY, GREY_NOTE, BLUE_TXT, GREEN_TXT = '1F2A44', '666666', '0000FF', '008000'
YELLOW, SEL_FILL, IMPORT_FILL, IMPORT_TXT = 'FFFF00', 'E8EAF6', 'F2F2F2', '404040'
SERIES = {'Base': '2A78D6', 'Pessimistic': 'EB6834', 'Optimistic': '1BAF7A', 'Custom': '9A9994', 'Actual': '52514E'}

NF_INT = r'#,##0;\(#,##0\);\-'
NF_INT1 = r'#,##0.0;\(#,##0.0\);\-'
NF_EUR = r'\€#,##0;"(€"#,##0\);\-'
NF_EUR2 = r'\€#,##0.00;"(€"#,##0.00\);\-'
NF_EUR4 = r'\€0.0000'
NF_PCT = r'0.0%;\(0.0%\);\-'
NF_PCT2 = r'0.00%;\-0.00%;0.00%'
NF_PCT3 = r'0.000%'
NF_RATIO = '0.0000'
NF_X = r'0.00\x'
NF_WK = r'0.0" wk"'
NF_DATE = 'dd mmm yyyy'
NF_PP = r'+0.00" pp";-0.00" pp";0.00" pp"'
NF_W = r'"W"0'

F_BODY = Font(name='Arial', size=10)
F_BOLD = Font(name='Arial', size=10, bold=True)
F_TITLE = Font(name='Arial', size=16, bold=True, color=NAVY)
F_H2 = Font(name='Arial', size=12, bold=True, color=NAVY)
F_NOTE = Font(name='Arial', size=9, italic=True, color=GREY_NOTE)
F_HEAD = Font(name='Arial', size=10, bold=True, color='FFFFFF')
F_INPUT = Font(name='Arial', size=10, color=BLUE_TXT)
F_GREEN = Font(name='Arial', size=10, color=GREEN_TXT)
F_GREEN_B = Font(name='Arial', size=10, bold=True, color=GREEN_TXT)
F_IMPORT = Font(name='Arial', size=10, color=IMPORT_TXT)
F_IMPORT_I = Font(name='Arial', size=10, italic=True, color='8C8C8C')
F_BANNER = Font(name='Arial', size=10, bold=True, color='7F4F00')

FILL_HEAD = PatternFill('solid', fgColor=NAVY)
FILL_YEL = PatternFill('solid', fgColor=YELLOW)
FILL_SEL = PatternFill('solid', fgColor=SEL_FILL)
FILL_IMP = PatternFill('solid', fgColor=IMPORT_FILL)
FILL_BANNER = PatternFill('solid', fgColor='FFF2CC')
FILL_SUB = PatternFill('solid', fgColor='D9DEEC')
THIN = Side(style='thin', color='BFBFBF')
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
WRAP = Alignment(wrap_text=True, vertical='top')
CENTER = Alignment(horizontal='center', vertical='center', wrap_text=True)


def cell(ws, ref, value=None, font=F_BODY, nf=None, fill=None, align=None, comment=None, border=None):
    c = ws[ref]
    if value is not None:
        c.value = value
    c.font = font
    if nf:
        c.number_format = nf
    if fill:
        c.fill = fill
    if align:
        c.alignment = align
    if border:
        c.border = border
    if comment:
        c.comment = Comment(comment, 'Echo model')
        c.comment.width, c.comment.height = 320, 140
    return c


def rc(ws, r, col, *a, **k):
    return cell(ws, f'{L(col)}{r}', *a, **k)


def title(ws, text, sub=None):
    cell(ws, 'A1', text, F_TITLE)
    if sub:
        cell(ws, 'A2', sub, F_NOTE)


def header(ws, row, col, labels, height=None):
    for i, lab in enumerate(labels):
        rc(ws, row, col + i, lab, F_HEAD, fill=FILL_HEAD, align=CENTER, border=BOX)
    if height:
        ws.row_dimensions[row].height = height


def inp(ws, ref, value, nf=None, custom=False, comment=None):
    return cell(ws, ref, value, F_INPUT, nf, FILL_YEL if custom else None, comment=comment)


def fml(ws, ref, formula, nf=None, bold=False, green=False, fill=None, comment=None):
    f = (F_GREEN_B if bold else F_GREEN) if green else (F_BOLD if bold else F_BODY)
    return cell(ws, ref, formula, f, nf, fill, comment=comment)


def imp(ws, ref, value, nf=None, italic=False, comment=None):
    return cell(ws, ref, value, F_IMPORT_I if italic else F_IMPORT, nf, FILL_IMP, comment=comment)


def banner(ws, ref, text, merge_to=None):
    c = cell(ws, ref, text, F_BANNER, fill=FILL_BANNER, align=WRAP)
    if merge_to:
        ws.merge_cells(f'{ref}:{merge_to}')
    return c


def note(ws, ref, text):
    return cell(ws, ref, text, F_NOTE)


def widths(ws, spec):
    for col, w in spec.items():
        ws.column_dimensions[col].width = w


def line_chart(title_txt, y_title, x_title=None, width=24, height=11):
    ch = LineChart()
    ch.title = title_txt
    ch.y_axis.title = y_title
    if x_title:
        ch.x_axis.title = x_title
    ch.width, ch.height = width, height
    ch.legend.position = 'b'
    ch.x_axis.delete = False
    ch.y_axis.delete = False
    ch.y_axis.majorGridlines = None
    return ch


def add_series(ch, ws, min_col, min_row, max_col, max_row, name, color, dash=None, width_emu=25400):
    s = Series(Reference(ws, min_col=min_col, min_row=min_row, max_col=max_col, max_row=max_row), title=name)
    s.graphicalProperties.line.solidFill = color
    s.graphicalProperties.line.width = width_emu
    if dash:
        s.graphicalProperties.line.dashStyle = dash
    s.smooth = False
    s.marker.symbol = 'none'
    ch.series.append(s)
    return s


def gap_formats(ws, rng, ok=0.005, warn=0.05):
    """Green if |gap| <= ok, amber if <= warn, red above (conditional formatting on a column of gaps)."""
    first = rng.split(':')[0]
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'AND(ISNUMBER({first}),ABS({first})<={ok})'],
                                                   fill=PatternFill('solid', fgColor='C6EFCE')))
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'AND(ISNUMBER({first}),ABS({first})>{ok},ABS({first})<={warn})'],
                                                   fill=PatternFill('solid', fgColor='FFEB9C')))
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'AND(ISNUMBER({first}),ABS({first})>{warn})'],
                                                   fill=PatternFill('solid', fgColor='FFC7CE')))


def scenario_dropdown(ws, ref, options):
    dv = DataValidation(type='list', formula1='"' + ','.join(options) + '"', allow_blank=False,
                        showDropDown=False, showErrorMessage=True,
                        errorTitle='Unknown scenario', error='Choose one of: ' + ', '.join(options),
                        promptTitle='Scenario', prompt='Choose the scenario shown in the model sheets')
    ws.add_data_validation(dv)
    dv.add(ref)


def finish(wb, path):
    for ws in wb.worksheets:
        ws.sheet_view.zoomScale = 90
    wb.calculation = CalcProperties(fullCalcOnLoad=True)
    wb.save(path)


def d2(s):
    return datetime.strptime(s, '%Y-%m-%d')


SCEN = ['Base', 'Pessimistic', 'Optimistic', 'Custom']
REF_LABEL = 'Original run = numbers in your brief (commit b88fe6a)'


# =============================================================================================
# VERSION 1 : SIMPLE MODEL
# =============================================================================================
def build_v1(d, ref, path):
    v1, c = d['v1'], d['const']
    wb = Workbook()
    wb.remove(wb.active)
    agg = {'Pessimistic'}

    # ------------------------------------------------------------------ Inputs
    wi = wb.create_sheet('Inputs')
    title(wi, 'Inputs & scenarios — simple model (version 1)',
          'Edit only the blue cells on yellow (column Custom) and the scenario selector C4. '
          'Base / Pessimistic / Optimistic reproduce the three scenarios of notebook section 9.3.')
    cell(wi, 'A4', 'Selected scenario', F_BOLD)
    inp(wi, 'C4', 'Base', custom=True, comment='Dropdown: Base / Pessimistic / Optimistic / Custom. '
        'Every sheet that says "selected scenario" (Model, unit economics below) follows this cell.')
    scenario_dropdown(wi, 'C4', SCEN)
    note(wi, 'D4', '← click the cell, then the small arrow, to switch scenario')
    header(wi, 6, 1, ['Parameter', 'Unit'] + SCEN + ['SELECTED', 'Note / source'], height=30)
    fit = v1['fit']
    churn = {s: fit[s if s != 'Custom' else 'Base']['churn'] for s in SCEN}
    arpu = {s: v1['arpu_week'] * (c['REV_RATIO'] if s in agg else 1) for s in SCEN}
    aph = {s: c['ADS_PER_HOUR'] * (c['APH_RATIO'] if s in agg else 1) for s in SCEN}
    rows = [
        ('Starting weekly active users (U₀)', 'users', {s: v1['U0'] for s in SCEN}, NF_INT,
         'Actual WAU in the last full week of data (week of 22 Jun 2026), notebook section 6 (kpi table). Same start as the cohort model after calibration.'),
        ('Paid spend in week 1', '€ / week', {s: v1['paid_spend_week1'] for s in SCEN}, NF_EUR2,
         f"Notebook spend trend for July 2026 (€{v1['spend_jul_month']:,.0f} per month, section 9.1 fit_trends) × 12 / 52."),
        ('Growth of paid spend', '% / week', {s: v1['paid_spend_growth_week'] for s in SCEN}, NF_PCT3,
         f"Notebook spend trend: +{v1['paid_spend_growth_month']:.2%} per month = (1 + monthly growth)^(12/52) − 1 per week."),
        ('CAC (paid social)', '€ / paid user', {s: v1['cac'] for s in SCEN}, NF_EUR2,
         'Median of monthly spend ÷ paid sign-ups, Jan 2025 - Jun 2026 (notebook sections 5.2 and 9.1).'),
        ('Free new users in week 1 (organic + cross-promo)', 'users / week', {s: v1['free_week1'] for s in SCEN}, NF_INT1,
         f"Notebook sign-up trends (organic store + podcast cross-promo) for July 2026: {v1['free_jul_month']:,.0f} per month × 12 / 52. "
         'Same idea as the "Organic sign-ups" line of the professor\'s Exercise sheet.'),
        ('Growth of free new users', '% / week', {s: v1['free_growth_week'] for s in SCEN}, NF_PCT3,
         f"Log-linear fit of the notebook's 12-month free sign-up forecast: +{v1['free_growth_month']:.2%} per month."),
        ('Weekly churn (equivalent)', '% of active users lost / week', churn, NF_PCT3,
         'NOT measured directly: least-squares value that makes this model reproduce the weekly active users of the notebook scenario '
         '(see sheet Calibration). Custom starts at the Base value.'),
        ('Referral rate', 'new users per retained user / week', {s: 0.0 for s in SCEN}, NF_PCT2,
         'Set to 0: Echo has no invitation data, so the K-factor cannot be measured (notebook section 3.4). Kept for compatibility with the professor\'s model.'),
        ('ARPU', '€ / active user / week', arpu, NF_EUR4,
         f"Ad revenue per active user-week of users with normal notifications (notebook section 5.1). Pessimistic = × {c['REV_RATIO']:.3f}: "
         'aggressive users hear more ads per hour (revenue ratio measured in the A/B test, section 9.3).'),
        ('Gross margin', '%', {s: 1.0 for s in SCEN}, NF_PCT,
         '100 %: no cost data in the Echo files, so all values are gross ad revenue (caveat of the notebook).'),
        ('Ads per listening hour', 'ads / hour', aph, '0.00',
         f"Median ads per listening hour (notebook section 2.5). Pessimistic = × {c['APH_RATIO']:.3f} (aggressive arm, section 7.2). Used to turn revenue into listening hours."),
        ('Revenue per ad impression (CPM / 1,000)', '€ / ad', {s: c['REV_PER_AD'] for s in SCEN}, NF_EUR4,
         'Data pack: Echo earns about €9 per 1,000 ad impressions.'),
    ]
    PR = {}
    for i, (lab, unit, vals, nf, src) in enumerate(rows):
        r = 7 + i
        PR[lab] = r
        cell(wi, f'A{r}', lab, F_BOLD)
        cell(wi, f'B{r}', unit, F_BODY)
        for j, s in enumerate(SCEN):
            cm = None
            if lab.startswith('Weekly churn') and s != 'Custom':
                cm = (f"Least-squares fit on the notebook's {s} trajectory: MAPE {fit[s]['mape']:.2%}, "
                      f"gap at week 52 {fit[s]['end_gap']:+.2%}. Source: export_from_notebook.py (fit_churn).")
            inp(wi, f'{L(3 + j)}{r}', vals[s], nf, custom=(s == 'Custom'), comment=cm)
        fml(wi, f'G{r}', f'=INDEX(C{r}:F{r},MATCH($C$4,$C$6:$F$6,0))', nf, bold=True, fill=FILL_SEL)
        note(wi, f'H{r}', src)
    r_u0, r_sp, r_sg, r_cac, r_fr, r_fg, r_ch, r_rf, r_ar, r_gm, r_aph, r_rpa = range(7, 19)

    cell(wi, 'A21', 'Unit economics (selected scenario)', F_H2)
    header(wi, 22, 1, ['Metric', 'Formula', 'Value'])
    ue = [
        ('New paid users in week 1', 'Spend / CAC', f'=IF(G{r_cac}>0,G{r_sp}/G{r_cac},0)', NF_INT1),
        ('All new users in week 1', 'paid + free', f'=C23+G{r_fr}', NF_INT1),
        ('Monthly equivalent churn', '1 − (1 − weekly churn)^(52/12)', f'=1-(1-G{r_ch})^(52/12)', NF_PCT),
        ('Average active lifetime', '1 / weekly churn', f'=IF(G{r_ch}>0,1/G{r_ch},0)', NF_WK),
        ('LTV (simple, gross margin)', 'ARPU × margin / weekly churn', f'=IF(G{r_ch}>0,G{r_ar}*G{r_gm}/G{r_ch},0)', NF_EUR2),
        ('LTV : CAC (simple)', 'LTV / CAC', f'=IF(G{r_cac}>0,C27/G{r_cac},0)', NF_X),
        ('CAC payback', 'CAC / (ARPU × margin), weeks', f'=IF(G{r_ar}*G{r_gm}>0,G{r_cac}/(G{r_ar}*G{r_gm}),0)', NF_WK),
        ('Weekly growth factor (excl. new users)', '(1 − churn) × (1 + referral)', f'=(1-G{r_ch})*(1+G{r_rf})', '0.0000'),
        ('Self-sustaining without new users?', 'factor ≥ 1', '=IF(C30>=1,"YES — the base grows alone","NO — needs new users every week")', None),
    ]
    for i, (a, b, f, nf) in enumerate(ue):
        r = 23 + i
        cell(wi, f'A{r}', a, F_BOLD)
        note(wi, f'B{r}', b)
        fml(wi, f'C{r}', f, nf)

    cell(wi, 'A33', 'Reality check: LTV measured per channel in the notebook (section 5.3, 2 years, gross)', F_H2)
    header(wi, 34, 1, ['Channel', 'LTV 2 years (data)', 'CAC', 'LTV : CAC'])
    for i, ch in enumerate(['paid_social', 'organic_store', 'podcast_cross_promo']):
        r = 35 + i
        cell(wi, f'A{r}', ch, F_BOLD)
        imp(wi, f'B{r}', d['ltv_channel'][ch]['LTV 104 weeks'], NF_EUR2,
            comment='Imported from notebook section 5.3 (table "LTV 104 weeks"), via export_from_notebook.py.')
        if ch == 'paid_social':
            fml(wi, f'C{r}', f'=G{r_cac}', NF_EUR2)
            fml(wi, f'D{r}', f'=IF(C{r}>0,B{r}/C{r},0)', NF_X)
        else:
            note(wi, f'C{r}', 'no media cost')
    note(wi, 'A39', 'Why the two LTVs differ: the simple LTV uses ONE blended churn for every user. Real paid-social users churn much faster in '
                    'their first weeks than the loyal base, so the simple formula overstates the value of a paid user. '
                    'Use the data LTV (row 35) for decisions about paid social.')
    widths(wi, {'A': 44, 'B': 30, 'C': 13, 'D': 13, 'E': 13, 'F': 13, 'G': 13, 'H': 110})
    wi.freeze_panes = 'C7'

    # ------------------------------------------------------------------ Model (selected scenario)
    wm = wb.create_sheet('Model')
    title(wm, 'Week-by-week model (52 weeks, July 2026 → June 2027)')
    fml(wm, 'A2', '="Scenario: "&Inputs!C4', green=True, bold=True)
    note(wm, 'D2', 'End users = Start users × (1 − churn) × (1 + referral) + Paid spend / CAC + Free new users   (same line as the professor\'s model, plus the free sign-ups of his Exercise sheet)')
    cols = ['Week', 'Week starting', 'Start users', 'Churned', 'Retained', 'Paid new users', 'Free new users',
            'Referrals', 'End users (WAU)', 'WoW growth', 'Revenue (€)', 'Gross profit (€)', 'Paid spend (€)',
            'Cumul. GP − spend (€)', 'Listening hours']
    header(wm, 4, 1, cols, height=32)
    rc(wm, 5, 1, 0)
    fml(wm, 'B5', '=DATE(2026,6,22)', NF_DATE, comment='Week 0 = last full week of data (22 Jun 2026): its WAU is the starting point U₀.')
    fml(wm, 'I5', '=Inputs!G7', NF_INT, green=True)
    rc(wm, 5, 14, 0, nf=NF_EUR)
    G = 'Inputs!$G$'
    for k in range(1, 53):
        r = 5 + k
        rc(wm, r, 1, k)
        fml(wm, f'B{r}', f'=B{r-1}+7', NF_DATE)
        fml(wm, f'C{r}', f'=I{r-1}', NF_INT)
        fml(wm, f'D{r}', f'=C{r}*{G}{r_ch}', NF_INT)
        fml(wm, f'E{r}', f'=C{r}-D{r}', NF_INT)
        fml(wm, f'M{r}', f'={G}{r_sp}*(1+{G}{r_sg})^(A{r}-1)', NF_EUR)
        fml(wm, f'F{r}', f'=IF({G}{r_cac}>0,M{r}/{G}{r_cac},0)', NF_INT)
        fml(wm, f'G{r}', f'={G}{r_fr}*(1+{G}{r_fg})^(A{r}-1)', NF_INT)
        fml(wm, f'H{r}', f'=E{r}*{G}{r_rf}', NF_INT)
        fml(wm, f'I{r}', f'=E{r}+F{r}+G{r}+H{r}', NF_INT, bold=True)
        fml(wm, f'J{r}', f'=IF(C{r}>0,I{r}/C{r}-1,0)', NF_PCT)
        fml(wm, f'K{r}', f'=I{r}*{G}{r_ar}', NF_EUR)
        fml(wm, f'L{r}', f'=K{r}*{G}{r_gm}', NF_EUR)
        fml(wm, f'N{r}', f'=N{r-1}+L{r}-M{r}', NF_EUR)
        fml(wm, f'O{r}', f'=IF({G}{r_aph}*{G}{r_rpa}>0,K{r}/{G}{r_rpa}/{G}{r_aph},0)', NF_INT)
    cell(wm, 'A58', 'Total', F_BOLD)
    for col, nf in [('D', NF_INT), ('F', NF_INT), ('G', NF_INT), ('H', NF_INT), ('K', NF_EUR), ('L', NF_EUR), ('M', NF_EUR)]:
        fml(wm, f'{col}58', f'=SUM({col}6:{col}57)', nf, bold=True)
    cell(wm, 'A59', 'Average weekly listening hours (North Star Metric)', F_BOLD)
    fml(wm, 'O59', '=AVERAGE(O6:O57)', NF_INT, bold=True)
    cell(wm, 'A60', 'WAU in the last week (21 Jun 2027)', F_BOLD)
    fml(wm, 'I60', '=I57', NF_INT, bold=True)
    w1 = v1['U0'] * (1 - churn['Base']) + v1['paid_spend_week1'] / v1['cac'] + v1['free_week1']
    note(wm, 'A62', f"Check by hand (Base): week 1 = {v1['U0']:,.0f} × (1 − {churn['Base']:.2%}) + {v1['paid_spend_week1']:,.0f} / {v1['cac']:.2f} "
                    f"+ {v1['free_week1']:,.0f} ≈ {w1:,.0f} users (cohort model of the notebook: {d['scenarios']['Base']['WAU'][0]:,.0f}).")
    widths(wm, {'A': 8, 'B': 13, 'C': 11, 'D': 10, 'E': 11, 'F': 11, 'G': 11, 'H': 10, 'I': 13, 'J': 10,
                'K': 12, 'L': 12, 'M': 12, 'N': 14, 'O': 12})
    wm.freeze_panes = 'C5'
    ch = line_chart('End users (weekly active users), selected scenario', 'users', 'week', 22, 10)
    add_series(ch, wm, 9, 5, 9, 57, 'End users (WAU)', SERIES['Base'])
    ch.set_categories(Reference(wm, min_col=1, min_row=5, max_row=57))
    wm.add_chart(ch, 'Q4')

    # ------------------------------------------------------------------ Scenarios
    wsn = wb.create_sheet('Scenarios')
    title(wsn, 'Scenarios & one-lever sensitivity',
          'Rows 5–8 pull their parameters from Inputs (green). Rows 12–18: Base with ONE lever changed (course slide 98). '
          'Rows 22–29: the 8 what-if lines of notebook section 9.4 replayed in the simple model.')
    cell(wsn, 'A3', 'Lever settings:', F_BOLD)
    note(wsn, 'B3', 'step')
    inp(wsn, 'C3', 0.20, NF_PCT, custom=True, comment='Size of the change applied in rows 12-15 (professor\'s model: 20 %).')
    note(wsn, 'D3', 'spend ×')
    inp(wsn, 'E3', 1.5, '0.00', custom=True, comment='Paid spend multiplier used in row 16 (same as notebook what-if "Paid social spend x 1.5").')
    note(wsn, 'F3', 'referral')
    inp(wsn, 'G3', 0.01, NF_PCT2, custom=True, comment='Hypothetical referral rate for row 18 (Echo has no referral programme: this is a "what if we built one").')
    pcols = ['Scenario', 'U₀ (users)', 'Spend wk 1 (€)', 'Spend growth / wk', 'CAC (€)', 'Free new wk 1', 'Free growth / wk',
             'Churn / wk', 'Referral / wk', 'ARPU (€ / wk)', 'Margin', 'Ads / hour', '€ / ad']
    header(wsn, 4, 1, pcols, height=42)
    W0 = 14                                                   # column N = week 0
    for k in range(53):
        rc(wsn, 4, W0 + k, k, F_HEAD, NF_W, FILL_HEAD, CENTER, border=BOX)
    out_cols = ['WAU week 52', '12-month revenue (€)', '12-month paid spend (€)', 'Revenue − spend (€)',
                'Avg listening hours / wk', 'LTV : CAC (simple)', 'WAU wk 52 vs Base']
    OC = W0 + 53                                              # first output column (BO)
    header(wsn, 4, OC, out_cols, height=42)
    wk1, wk52 = L(W0 + 1), L(W0 + 52)

    def scen_row(r, label, params, font_label=F_BODY):
        """params: list of 12 (value, kind) with kind in {'f','g','i'} formula / green link / input."""
        cell(wsn, f'A{r}', label, font_label)
        nfs = [NF_INT, NF_EUR2, NF_PCT3, NF_EUR2, NF_INT1, NF_PCT3, NF_PCT3, NF_PCT2, NF_EUR4, NF_PCT, '0.00', NF_EUR4]
        for j, (val, kind, *cm) in enumerate(params):
            ref_ = f'{L(2 + j)}{r}'
            if kind == 'g':
                fml(wsn, ref_, val, nfs[j], green=True)
            elif kind == 'i':
                inp(wsn, ref_, val, nfs[j], comment=cm[0] if cm else None)
            else:
                fml(wsn, ref_, val, nfs[j])
        fml(wsn, f'{L(W0)}{r}', f'=B{r}', NF_INT)
        for k in range(1, 53):
            col, prev = L(W0 + k), L(W0 + k - 1)
            fml(wsn, f'{col}{r}', f'={prev}{r}*(1-$H{r})*(1+$I{r})+IF($E{r}>0,$C{r}*(1+$D{r})^({col}$4-1)/$E{r},0)+$F{r}*(1+$G{r})^({col}$4-1)', NF_INT)
        o = [L(OC + i) for i in range(7)]
        fml(wsn, f'{o[0]}{r}', f'={wk52}{r}', NF_INT, bold=True)
        fml(wsn, f'{o[1]}{r}', f'=SUM({wk1}{r}:{wk52}{r})*$J{r}', NF_EUR)
        fml(wsn, f'{o[2]}{r}', f'=IF($D{r}=0,$C{r}*52,$C{r}*((1+$D{r})^52-1)/$D{r})', NF_EUR)
        fml(wsn, f'{o[3]}{r}', f'={o[1]}{r}-{o[2]}{r}', NF_EUR)
        fml(wsn, f'{o[4]}{r}', f'=IF($L{r}*$M{r}>0,{o[1]}{r}/52/$M{r}/$L{r},0)', NF_INT)
        fml(wsn, f'{o[5]}{r}', f'=IF(AND($H{r}>0,$E{r}>0),$J{r}*$K{r}/$H{r}/$E{r},0)', NF_X)
        fml(wsn, f'{o[6]}{r}', f'=IF(${o[0]}$5>0,{o[0]}{r}/${o[0]}$5-1,0)', NF_PCT)

    inputs_rows = [r_u0, r_sp, r_sg, r_cac, r_fr, r_fg, r_ch, r_rf, r_ar, r_gm, r_aph, r_rpa]
    for i, s in enumerate(SCEN):
        colI = L(3 + i)
        scen_row(5 + i, s, [(f'=Inputs!{colI}{x}', 'g') for x in inputs_rows], F_BOLD)
    base = lambda col: (f'={col}$5', 'f')
    cell(wsn, 'A11', 'One-lever sensitivity: Base with ONE lever changed', F_H2)
    levers = [
        ('Base, CAC −20 %', {'E': (f'=E$5*(1-$C$3)', 'f')}),
        ('Base, churn −20 %', {'H': (f'=H$5*(1-$C$3)', 'f')}),
        ('Base, ARPU +20 %', {'J': (f'=J$5*(1+$C$3)', 'f')}),
        ('Base, free sign-ups +20 %', {'F': (f'=F$5*(1+$C$3)', 'f')}),
        ('Base, paid spend × 1.5', {'C': (f'=C$5*$E$3', 'f')}),
        ('Base, paid spend stopped', {'C': (0, 'i', 'Lever: paid spend set to 0 (notebook what-if "Paid social stopped").')}),
        ('Base, referral programme (1 % / wk)', {'I': (f'=$G$3', 'f')}),
    ]
    pletters = [L(2 + j) for j in range(12)]
    for i, (lab, change) in enumerate(levers):
        scen_row(12 + i, lab, [change.get(pl, base(pl)) for pl in pletters])

    cell(wsn, 'A21', 'Notebook what-if lines (section 9.4) replayed in the simple model', F_H2)
    note(wsn, 'H21', 'Churn: re-calibrated by least squares for push / activation lines (blue); spend lines keep the Base churn on purpose, to test how the simple model reacts to spend.')
    WI_ROWS = {}
    for i, (name, info) in enumerate(v1['whatif'].items()):
        r = 22 + i
        WI_ROWS[name] = r
        if name != 'Base' and info['recalibrated']:
            ch_ = (info['churn'], 'i', f"Least-squares churn fitted on the notebook what-if line '{name}' (MAPE {info.get('mape', 0):.2%}).")
        else:
            ch_ = ('=H$5', 'f')
        p = {pl: base(pl) for pl in pletters}
        p['H'] = ch_
        if info['spend_mult'] == 0:
            p['C'] = (0, 'i', 'Paid spend stopped (notebook what-if "Paid social stopped").')
        elif info['spend_mult'] != 1.0:
            p['C'] = ('=C$5*$E$3', 'f')
        if info['aggressive']:
            p['J'] = ('=J$6', 'f')
            p['L'] = ('=L$6', 'f')
        scen_row(r, name, [p[pl] for pl in pletters])
    widths(wsn, {'A': 46, 'B': 9, 'C': 10, 'D': 10, 'E': 9, 'F': 10, 'G': 10, 'H': 10, 'I': 9, 'J': 10, 'K': 8, 'L': 8, 'M': 9})
    for k in range(53):
        wsn.column_dimensions[L(W0 + k)].width = 8
    for i in range(7):
        wsn.column_dimensions[L(OC + i)].width = 14
    wsn.freeze_panes = 'B5'
    ch = line_chart('Weekly active users by scenario (simple model)', 'users', 'week', 30, 12)
    for i, s in enumerate(['Base', 'Pessimistic', 'Optimistic']):
        add_series(ch, wsn, W0, 5 + i, W0 + 52, 5 + i, s, SERIES[s])
    ch.set_categories(Reference(wsn, min_col=W0, min_row=4, max_col=W0 + 52, max_row=4))
    wsn.add_chart(ch, 'A32')

    # ------------------------------------------------------------------ Calibration
    wc = wb.create_sheet('Calibration')
    title(wc, 'Calibration: how the "equivalent churn" was chosen (least squares)',
          'For each scenario we look for the ONE weekly churn that makes the simple model closest to the cohort model of the notebook: '
          'we minimise the sum over the 52 weeks of (simple WAU − cohort WAU)². The scan tables below prove it is the minimum.')
    note(wc, 'A3', 'Scan step (change of churn between two scan rows):')
    inp(wc, 'F3', 0.0005, NF_PCT3, custom=True, comment='0.05 percentage point. Make it smaller to zoom on the minimum.')
    note(wc, 'H3', 'To redo the fit inside Excel: Data > Solver, minimise the SSE cell of a scenario by changing that scenario\'s churn on Inputs row 13.')
    C0 = 3                                                   # column C = week 1
    wkcols = [L(C0 + k) for k in range(52)]
    SSE, RMSE, MAPE, GAP = L(C0 + 52), L(C0 + 53), L(C0 + 54), L(C0 + 55)
    top0 = 26
    CAL = {}
    for b, s in enumerate(['Base', 'Pessimistic', 'Optimistic']):
        t = top0 + b * 31
        srow = 5 + SCEN.index(s)                             # row of the scenario in Scenarios
        cell(wc, f'A{t}', f'{s} scenario', F_H2)
        fml(wc, f'C{t}', f'="Equivalent churn on Inputs: "&TEXT(Inputs!{L(3 + SCEN.index(s))}{r_ch},"0.000%")&" per week"', green=True)
        header(wc, t + 1, 1, ['Row', 'Churn / wk'] + [f'W{k}' for k in range(1, 53)] + ['SSE', 'RMSE (users)', 'MAPE', 'Gap at W52'], height=30)
        cell(wc, f'A{t+2}', 'Cohort model WAU (notebook, imported)', F_BOLD)
        for k in range(52):
            imp(wc, f'{wkcols[k]}{t+2}', d['scenarios'][s]['WAU'][k], NF_INT)
        wc[f'{wkcols[0]}{t+2}'].comment = Comment('Imported: weekly WAU of this notebook scenario (section 9.3, run()), via export_from_notebook.py.', 'Echo model')
        cell(wc, f'A{t+3}', 'Simple model WAU (calibrated churn)', F_BOLD)
        fml(wc, f'B{t+3}', f'=Scenarios!$H${srow}', NF_PCT3, green=True)
        for k in range(52):
            fml(wc, f'{wkcols[k]}{t+3}', f'=Scenarios!{L(W0 + 1 + k)}{srow}', NF_INT, green=True)
        cell(wc, f'A{t+4}', 'Error (simple − cohort) ÷ cohort', F_BODY)
        for k in range(52):
            fml(wc, f'{wkcols[k]}{t+4}', f'={wkcols[k]}{t+3}/{wkcols[k]}{t+2}-1', NF_PCT)
        rng_s, rng_t = f'{wkcols[0]}{t+3}:{wkcols[-1]}{t+3}', f'{wkcols[0]}{t+2}:{wkcols[-1]}{t+2}'
        fml(wc, f'{SSE}{t+3}', f'=SUMXMY2({rng_s},{rng_t})', NF_INT, bold=True)
        fml(wc, f'{RMSE}{t+3}', f'=SQRT({SSE}{t+3}/52)', NF_INT)
        fml(wc, f'{MAPE}{t+3}', f'=SUMPRODUCT(ABS({rng_s}-{rng_t})/{rng_t})/52', NF_PCT2, bold=True)
        fml(wc, f'{GAP}{t+3}', f'={wkcols[-1]}{t+3}/{wkcols[-1]}{t+2}-1', NF_PCT2)
        CAL[s] = {'mape': f'{MAPE}{t+3}', 'rmse': f'{RMSE}{t+3}', 'sse': f'{SSE}{t+3}'}
        cell(wc, f'A{t+6}', 'SSE scan: the same model with the churn moved around the calibrated value', F_BOLD)
        first, last = t + 7, t + 27
        for j in range(21):
            r = first + j
            cell(wc, f'A{r}', f'scan {j - 10:+d}', F_NOTE)
            fml(wc, f'B{r}', f'=$B${t+3}+({j - 10})*$F$3', NF_PCT3)
            for k in range(52):
                prev = f'Scenarios!$B${srow}' if k == 0 else f'{wkcols[k-1]}{r}'
                fml(wc, f'{wkcols[k]}{r}',
                    f'={prev}*(1-$B{r})*(1+Scenarios!$I${srow})+IF(Scenarios!$E${srow}>0,Scenarios!$C${srow}*(1+Scenarios!$D${srow})^{k}/Scenarios!$E${srow},0)'
                    f'+Scenarios!$F${srow}*(1+Scenarios!$G${srow})^{k}', NF_INT)
            fml(wc, f'{SSE}{r}', f'=SUMXMY2({wkcols[0]}{r}:{wkcols[-1]}{r},{rng_t})', NF_INT)
        cell(wc, f'A{t+28}', 'Churn with the lowest SSE in the scan', F_BOLD)
        fml(wc, f'B{t+28}', f'=INDEX(B{first}:B{last},MATCH(MIN({SSE}{first}:{SSE}{last}),{SSE}{first}:{SSE}{last},0))', NF_PCT3, bold=True)
        fml(wc, f'C{t+28}', f'=IF(ABS(B{t+28}-B{t+3})<$F$3/2,"OK: the churn on Inputs is the least-squares minimum","Not the minimum: re-fit the churn")')
    widths(wc, {'A': 40, 'B': 11})
    for k in range(52):
        wc.column_dimensions[wkcols[k]].width = 8
    for col in (SSE, RMSE, MAPE, GAP):
        wc.column_dimensions[col].width = 13
    wc.freeze_panes = 'C5'
    ch = line_chart('Cohort model (dashed) vs simple model (solid), weekly active users', 'users', 'week', 30, 9.5)
    for b, s in enumerate(['Base', 'Pessimistic', 'Optimistic']):
        t = top0 + b * 31
        add_series(ch, wc, C0, t + 2, C0 + 51, t + 2, f'{s} — cohort (notebook)', SERIES[s], dash='dash')
        add_series(ch, wc, C0, t + 3, C0 + 51, t + 3, f'{s} — simple', SERIES[s])
    ch.set_categories(Reference(wc, min_col=C0, min_row=top0 + 1, max_col=C0 + 51, max_row=top0 + 1))
    wc.add_chart(ch, 'A5')

    # ------------------------------------------------------------------ Backtest
    wbk = wb.create_sheet('Backtest')
    bt, nbt = v1['backtest'], d['backtest']
    title(wbk, 'Backtest of the simple model alone (no cohort model)',
          'Churn fitted by least squares on the actual WAU of July–December 2025 only, then the model predicts January–June 2026 '
          '(A: sign-ups forecast from the 2025 trend; B: actual sign-ups). Same test as notebook section 9.2.')
    lab = [('Weekly churn fitted on Jul–Dec 2025 (least squares)', bt['churn'], NF_PCT3, 'i',
            f"Least squares on 25 weeks (fit MAPE {bt['fit_mape']:.2%}). export_from_notebook.py."),
           ('WAU at the start of the fit (week of 30 Jun 2025)', bt['u0_fit'], NF_INT, 'm', 'Actual WAU, notebook section 6 (kpi).'),
           ('WAU at the start of the forecast (week of 22 Dec 2025)', bt['u0'], NF_INT, 'm', 'Actual WAU, notebook section 6 (kpi).'),
           ('Revenue per active user per week, Jul–Dec 2025', bt['arpu'], NF_EUR4, 'm', 'Sum of revenue ÷ sum of WAU over the fit weeks.')]
    for i, (a, v, nf, kind, cm) in enumerate(lab):
        cell(wbk, f'A{4+i}', a, F_BOLD)
        (inp if kind == 'i' else imp)(wbk, f'C{4+i}', v, nf, comment=cm)
    header(wbk, 10, 1, ['Week starting', 'Actual new sign-ups', 'Actual WAU', 'Simple model WAU', 'Error'], height=30)
    for i, wk in enumerate(bt['fit_weeks']):
        r = 11 + i
        imp(wbk, f'A{r}', d2(wk), NF_DATE)
        imp(wbk, f'B{r}', bt['fit_signups'][i], NF_INT)
        imp(wbk, f'C{r}', d['history']['WAU'][d['history']['weeks'].index(wk)], NF_INT)
        prev = '$C$5' if i == 0 else f'D{r-1}'
        fml(wbk, f'D{r}', f'={prev}*(1-$C$4)+B{r}', NF_INT)
        fml(wbk, f'E{r}', f'=D{r}/C{r}-1', NF_PCT)
    rf_last = 11 + len(bt['fit_weeks']) - 1
    cell(wbk, f'A{rf_last+1}', 'Fit MAPE', F_BOLD)
    fml(wbk, f'E{rf_last+1}', f'=SUMPRODUCT(ABS(E11:E{rf_last}))/{len(bt["fit_weeks"])}', NF_PCT2, bold=True)
    h0 = rf_last + 4
    header(wbk, h0, 1, ['Week starting', 'New users A (2025 trend)', 'New users B (actual)', 'Actual WAU', 'Actual revenue (€)',
                        'Simple A: WAU', 'Simple B: WAU', 'Simple A: revenue', 'Simple B: revenue',
                        '|err| A WAU', '|err| B WAU', '|err| A rev', '|err| B rev'], height=42)
    n = len(nbt['weeks'])
    for i in range(n):
        r = h0 + 1 + i
        imp(wbk, f'A{r}', d2(nbt['weeks'][i]), NF_DATE)
        imp(wbk, f'B{r}', bt['new_A'][i], NF_INT)
        imp(wbk, f'C{r}', bt['new_B'][i], NF_INT)
        imp(wbk, f'D{r}', nbt['actual_WAU'][i], NF_INT)
        imp(wbk, f'E{r}', nbt['actual_revenue'][i], NF_EUR)
        pa, pb = ('$C$6', '$C$6') if i == 0 else (f'F{r-1}', f'G{r-1}')
        fml(wbk, f'F{r}', f'={pa}*(1-$C$4)+B{r}', NF_INT)
        fml(wbk, f'G{r}', f'={pb}*(1-$C$4)+C{r}', NF_INT)
        fml(wbk, f'H{r}', f'=F{r}*$C$7', NF_EUR)
        fml(wbk, f'I{r}', f'=G{r}*$C$7', NF_EUR)
        fml(wbk, f'J{r}', f'=ABS(F{r}/D{r}-1)', NF_PCT)
        fml(wbk, f'K{r}', f'=ABS(G{r}/D{r}-1)', NF_PCT)
        fml(wbk, f'L{r}', f'=ABS(H{r}/E{r}-1)', NF_PCT)
        fml(wbk, f'M{r}', f'=ABS(I{r}/E{r}-1)', NF_PCT)
    rl = h0 + n
    s0 = rl + 3
    cell(wbk, f'A{s0}', 'Backtest error (MAPE)', F_H2)
    header(wbk, s0 + 1, 1, ['Metric', 'Simple model (this sheet)', 'Cohort model (notebook 9.2)'], height=30)
    BTV1 = {}
    for i, (lab_, colx, key, sub) in enumerate([('A: MAPE WAU', 'J', 'A: forecast signups', 'MAPE WAU'),
                                                 ('A: MAPE revenue', 'L', 'A: forecast signups', 'MAPE revenue'),
                                                 ('B: MAPE WAU', 'K', 'B: actual signups', 'MAPE WAU'),
                                                 ('B: MAPE revenue', 'M', 'B: actual signups', 'MAPE revenue')]):
        r = s0 + 2 + i
        cell(wbk, f'A{r}', lab_, F_BOLD)
        fml(wbk, f'B{r}', f'=AVERAGE({colx}{h0+1}:{colx}{rl})', NF_PCT2, bold=True)
        imp(wbk, f'C{r}', nbt['table'][key][sub], NF_PCT2, comment='Imported from notebook section 9.2 (bt table).')
        BTV1[lab_] = f'B{r}'
    note(wbk, f'A{s0+7}', 'Reading: with a single churn fitted on 2025, the simple model cannot know that the user base gets older (more loyal users, '
                          'lower churn), so it under-predicts 2026. The cohort model handles this because each cohort keeps its own age.')
    widths(wbk, {'A': 50, 'B': 15, 'C': 15, 'D': 13, 'E': 14, 'F': 13, 'G': 13, 'H': 14, 'I': 14, 'J': 11, 'K': 11, 'L': 11, 'M': 11})
    ch = line_chart('Backtest Jan–Jun 2026: actual vs simple model', 'users', None, 22, 10)
    add_series(ch, wbk, 4, h0 + 1, 4, rl, 'Actual WAU', SERIES['Actual'])
    add_series(ch, wbk, 6, h0 + 1, 6, rl, 'Simple A (forecast sign-ups)', SERIES['Base'])
    add_series(ch, wbk, 7, h0 + 1, 7, rl, 'Simple B (actual sign-ups)', SERIES['Pessimistic'])
    ch.set_categories(Reference(wbk, min_col=1, min_row=h0 + 1, max_row=rl))
    wbk.add_chart(ch, 'O4')

    # ------------------------------------------------------------------ Checks
    wk = wb.create_sheet('Checks')
    title(wk, 'Checks: simple model vs notebook',
          'Gap = this workbook ÷ reference − 1. Colours: green |gap| ≤ 0.5 %, amber ≤ 5 %, red > 5 %. '
          'The simple model is NOT expected to match exactly: the gaps measure what is lost by using one constant churn.')
    header(wk, 4, 1, ['Metric', 'Scenario / line', 'Notebook (current version)', 'This workbook (simple)', 'Gap vs notebook',
                      REF_LABEL, 'Gap vs brief', 'Comment'], height=45)
    rows_ck = []
    srow = {s: 5 + SCEN.index(s) for s in SCEN}
    OCL = [L(OC + i) for i in range(7)]
    for s in ['Pessimistic', 'Base', 'Optimistic']:
        sc = d['scenarios'][s]
        rows_ck.append(('WAU June 2027', s, sc['WAU'][-1], f'=Scenarios!{OCL[0]}{srow[s]}', ref['scenarios'][s]['WAU_end'], NF_INT, ''))
    for s in ['Pessimistic', 'Base', 'Optimistic']:
        sc = d['scenarios'][s]
        rows_ck.append(('12-month ad revenue (€)', s, sum(sc['revenue']), f'=Scenarios!{OCL[1]}{srow[s]}', ref['scenarios'][s]['revenue_12m'], NF_EUR,
                        'ARPU is data-based (€0.1036 / active week), not fitted'))
    for s in ['Pessimistic', 'Base', 'Optimistic']:
        sc = d['scenarios'][s]
        rows_ck.append(('Average weekly listening hours', s, sum(sc['hours']) / 52, f'=Scenarios!{OCL[4]}{srow[s]}', ref['scenarios'][s]['hours_avg'], NF_INT, ''))
    rows_ck.append(('12-month paid spend (€)', 'all scenarios', d['scenarios']['Base']['paid_spend_12m'], f'=Scenarios!{OCL[2]}5',
                    ref['scenarios']['Base']['paid_spend_12m'], NF_EUR, 'Simple model: 52 full weeks; notebook starts on 1 July (week 1 has 5 days)'))
    rows_ck.append(('Calibration factor (existing cohorts)', '—', c['CALIB'], 'not used', ref['CALIB'], '0.000',
                    'The simple model starts directly from the actual WAU (7,457), so it needs no calibration factor'))
    for s in ['Pessimistic', 'Base', 'Optimistic']:
        rows_ck.append(('Fit of the WAU path: MAPE over 52 weeks', s, 0, f'=Calibration!{CAL[s]["mape"]}', None, NF_PCT2,
                        'Average weekly distance to the notebook path (target 0)'))
    for name in d['whatif']:
        r = WI_ROWS[name]
        rows_ck.append(('What-if: WAU June 2027', name, d['whatif'][name]['WAU June 2027'], f'=Scenarios!{OCL[0]}{r}',
                        ref['whatif'][name]['WAU June 2027'], NF_INT, ''))
    for name in d['whatif']:
        r = WI_ROWS[name]
        rows_ck.append(('What-if: 12-month revenue (€)', name, d['whatif'][name]['12-month revenue'], f'=Scenarios!{OCL[1]}{r}',
                        ref['whatif'][name]['12-month revenue'], NF_EUR, ''))
    for name in d['whatif']:
        r = WI_ROWS[name]
        rows_ck.append(('What-if: revenue − paid spend (€)', name, d['whatif'][name]['revenue - spend'], f'=Scenarios!{OCL[3]}{r}',
                        ref['whatif'][name]['revenue - spend'], NF_EUR,
                        'Spend lines: the simple model gives paid users the blended churn, so it overstates their value' if 'spend' in name.lower() or 'stopped' in name else ''))
    for lab_, key, sub in [('A: MAPE WAU', 'A: forecast signups', 'MAPE WAU'), ('A: MAPE revenue', 'A: forecast signups', 'MAPE revenue'),
                           ('B: MAPE WAU', 'B: actual signups', 'MAPE WAU'), ('B: MAPE revenue', 'B: actual signups', 'MAPE revenue')]:
        rows_ck.append(('Backtest ' + lab_, 'Jan–Jun 2026', d['backtest']['table'][key][sub], f'=Backtest!{BTV1[lab_]}',
                        ref['backtest'][key][sub], NF_PCT2, 'Gap = difference of error rates in percentage points (×100). Simple model fitted on 2025 only vs cohort model; lower MAPE is better'))
    for i, (m, sline, nbv, wbf, refv, nf, cm) in enumerate(rows_ck):
        r = 5 + i
        cell(wk, f'A{r}', m, F_BOLD)
        cell(wk, f'B{r}', sline)
        imp(wk, f'C{r}', nbv, nf)
        if wbf == 'not used':
            cell(wk, f'D{r}', 'not used', F_NOTE)
        else:
            fml(wk, f'D{r}', wbf, nf, green=True)
            if m.startswith('Backtest'):          # error rates: compare in percentage points, not in %
                fml(wk, f'E{r}', f'=(D{r}-C{r})*100', NF_PP)
            elif not m.startswith('Fit of'):
                fml(wk, f'E{r}', f'=IF(C{r}=0,"",D{r}/C{r}-1)', NF_PCT2)
        if refv is not None:
            imp(wk, f'F{r}', refv, nf)
            if m.startswith('Backtest'):
                fml(wk, f'G{r}', f'=(D{r}-F{r})*100', NF_PP)
            elif wbf != 'not used' and not m.startswith('Fit of'):
                fml(wk, f'G{r}', f'=IF(F{r}=0,"",D{r}/F{r}-1)', NF_PCT2)
        note(wk, f'H{r}', cm)
    last = 4 + len(rows_ck)
    wk['C4'].comment = Comment('Imported from the CURRENT notebook (group version, git HEAD) by export_from_notebook.py.', 'Echo model')
    wk['F4'].comment = Comment('Imported from reference_original_run.json: the original notebook run, whose numbers are quoted in the brief. '
                               'The current notebook differs by 0.4-1.3 % because of two cleaning changes (see README).', 'Echo model')
    gap_formats(wk, f'E5:E{last}')
    gap_formats(wk, f'G5:G{last}')
    widths(wk, {'A': 40, 'B': 44, 'C': 16, 'D': 16, 'E': 12, 'F': 20, 'G': 12, 'H': 90})
    wk.freeze_panes = 'C5'

    # ------------------------------------------------------------------ Guide (first sheet)
    wg = wb.create_sheet('Guide', 0)
    lines = [
        ('Echo growth model — Version 1 "SIMPLE" (IMT, growth modeling final project)', F_TITLE),
        ('', None),
        ('WHAT IT ANSWERS', F_BOLD),
        ('Board question: should Echo roll out aggressive push notifications to everyone? This file replays the 3 scenarios of notebook section 9.3', F_BODY),
        ('(Pessimistic = aggressive push for everyone, Base = status quo, Optimistic = guided onboarding) with the professor\'s one-line model.', F_BODY),
        ('', None),
        ('HOW IT WORKS', F_BOLD),
        ('Each week:  End users = Start users × (1 − churn) × (1 + referral rate)  +  Paid spend / CAC  +  Free new users', F_BODY),
        ('Same line as the professor\'s Growth Model Simulator, plus the "organic sign-ups" term of his Exercise sheet: 60 % of Echo\'s new users cost nothing.', F_NOTE),
        ('Time step = 1 WEEK (52 steps): the notebook measures and forecasts weekly active users (WAU); a monthly step would mix WAU with monthly churn.', F_NOTE),
        ('Referral rate = 0: Echo has no invitation data, so there is no measured referral loop (the cell is kept so you can test one).', F_NOTE),
        ('Churn is an "equivalent churn": the single weekly value that best reproduces the cohort model of the notebook (least squares, sheet Calibration).', F_NOTE),
        ('', None),
        ('COLOR CODE', F_BOLD),
        ('Blue text on yellow = inputs you can change (Inputs column Custom, the selector Inputs!C4, the lever settings of Scenarios and Calibration).', F_BODY),
        ('Blue text without yellow = scenario values taken from the notebook (change them only if you know why).', F_BODY),
        ('Black = formulas (do not edit).   Green = values pulled from another sheet.   Dark grey on light grey = imported data (static values from Python).', F_BODY),
        ('', None),
        ('SHEETS', F_BOLD),
        ('Inputs — parameters of the 4 scenarios (Base / Pessimistic / Optimistic / Custom), the selected scenario (C4) and unit economics.', F_BODY),
        ('Model — week-by-week detail (52 weeks) of the scenario selected in Inputs!C4, with a chart.', F_BODY),
        ('Scenarios — all scenarios side by side, one-lever sensitivity tests (slide 98), the notebook what-if lines, and a chart.', F_BODY),
        ('Calibration — how each equivalent churn was fitted, with a scan proving it is the least-squares minimum.', F_BODY),
        ('Backtest — the simple model fitted on 2025 only, predicting Jan–Jun 2026 (compare with the cohort model).', F_BODY),
        ('Checks — every output compared with the notebook (current version) and with the original run quoted in the brief.', F_BODY),
        ('', None),
        ('TRY THIS', F_BOLD),
        ('1. Switch Inputs!C4 between Base / Pessimistic / Optimistic and read Model!I57 (WAU after 52 weeks) and Model!K58 (12-month revenue).', F_BODY),
        ('2. In Scenarios, compare the one-lever rows 12–18: which lever moves WAU at week 52 the most?', F_BODY),
        ('3. In Inputs column Custom, set the paid spend (F8) to 0: how many free new users per week keep the base from shrinking?', F_BODY),
        ('4. Look at Inputs rows 27-28 and 35-37: why does the simple LTV say paid social is fine when the data says it loses money?', F_BODY),
        ('', None),
        ('Source: Echo data pack (fictional data), notebook echo_growth_analysis.ipynb (sections 5, 7, 9). Rebuild: export_from_notebook.py then build_workbooks.py.', F_NOTE),
    ]
    for i, (txt, f) in enumerate(lines):
        if f:
            cell(wg, f'B{i+1}', txt, f)
    widths(wg, {'A': 3, 'B': 140})
    finish(wb, path)


if __name__ == '__main__':
    import sys
    data = json.loads((HERE / 'model_data.json').read_text())
    ref_ = json.loads((HERE / 'reference_original_run.json').read_text())
    build_v1(data, ref_, HERE / 'echo_growth_model_v1_simple.xlsx')
    if 'v1only' not in sys.argv:
        from build_v2 import build_v2
        build_v2(data, ref_, HERE / 'echo_growth_model_v2_cohort.xlsx')
    print('Workbooks written.')
