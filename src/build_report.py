from datetime import date
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.formatting.rule import CellIsRule, ColorScaleRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parents[1]
DATA = ROOT / "data"
OUT = ROOT / "support_kpi_report_q3_2026.xlsx"

FONT = "Arial"
NAVY = "1F4E79"
HEADER_FILL = PatternFill("solid", fgColor=NAVY)
HELPER_FILL = PatternFill("solid", fgColor="7F7F7F")
TOTAL_FILL = PatternFill("solid", fgColor="DDEBF7")
# conditional formatting fills need bgColor set as well
RED_FILL = PatternFill("solid", start_color="F8CBAD", end_color="F8CBAD")
GREEN_FILL = PatternFill("solid", start_color="C6EFCE", end_color="C6EFCE")
THIN = Side(style="thin", color="BFBFBF")

PCT, INT_, DEC1, DEC2, DATE_FMT = "0.0%", "#,##0", "0.0", "0.00", "dd mmm yyyy"


def font(**kw):
    return Font(name=FONT, size=kw.pop("size", 10), **kw)


def header_row(ws, row, labels, start_col=1, fill=HEADER_FILL):
    for i, label in enumerate(labels):
        c = ws.cell(row=row, column=start_col + i, value=label)
        c.font = font(bold=True, color="FFFFFF")
        c.fill = fill
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[row].height = 42


def title(ws, text, subtitle=None):
    ws["A1"] = text
    ws["A1"].font = font(size=14, bold=True, color=NAVY)
    if subtitle:
        ws["A2"] = subtitle
        ws["A2"].font = font(italic=True, color="595959")


def widths(ws, mapping):
    for col, w in mapping.items():
        ws.column_dimensions[col].width = w


def fmt(ws, cell_range, number_format=None, bold=False, fill=None):
    for row in ws[cell_range]:
        for c in row:
            if number_format:
                c.number_format = number_format
            c.font = font(bold=bold)
            c.border = Border(top=THIN, bottom=THIN, left=THIN, right=THIN)
            if fill:
                c.fill = fill


def weekly_combo_chart(ws_week, first_w, lw, width):
    bar = BarChart()
    bar.title = "Weekly tickets vs SLA met %"
    bar.add_data(Reference(ws_week, min_col=2, min_row=4, max_row=lw), titles_from_data=True)
    bar.set_categories(Reference(ws_week, min_col=1, min_row=first_w, max_row=lw))
    bar.y_axis.title = "Tickets"
    bar.y_axis.majorGridlines = None
    bar.x_axis.number_format = "dd mmm"
    bar.x_axis.auto = False
    bar.gapWidth = 60
    line = LineChart()
    line.add_data(Reference(ws_week, min_col=6, min_row=4, max_row=lw), titles_from_data=True)
    line.series[0].smooth = False
    line.y_axis.axId = 200
    line.y_axis.title = "SLA met %"
    line.y_axis.number_format = "0%"
    line.y_axis.crosses = "max"
    line.y_axis.delete = False
    bar += line
    chart_style(bar)
    bar.legend.position = "b"
    bar.height, bar.width = 8, width
    return bar


def print_setup(ws):
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True


def chart_style(chart):
    chart.roundedCorners = False
    if chart.legend is not None:
        chart.legend.overlay = False
    chart.x_axis.delete = False
    chart.y_axis.delete = False


def build():
    tickets = pd.read_csv(DATA / "helpdesk_tickets_q3_2026.csv",
                          parse_dates=["created_at", "first_response_at", "resolved_at"])
    agents = pd.read_csv(DATA / "agents.csv", parse_dates=["start_date"])
    n = len(tickets)
    last = n + 1

    wb = Workbook()
    ws_sum = wb.active
    ws_sum.title = "Summary"
    ws_week = wb.create_sheet("Weekly_KPIs")
    ws_agent = wb.create_sheet("Agent_Performance")
    ws_brk = wb.create_sheet("Breakdowns")
    ws_dq = wb.create_sheet("Data_Quality")
    ws_t = wb.create_sheet("Tickets")
    ws_a = wb.create_sheet("Agents")
    ws_set = wb.create_sheet("Settings")

    # ---------- Settings ----------
    title(ws_set, "Settings & targets", "Blue cells are inputs - change them and every report sheet updates.")
    header_row(ws_set, 4, ["Priority", "First response SLA (minutes)"])
    for i, (p, m) in enumerate([("High", 30), ("Medium", 60), ("Low", 240)], start=5):
        ws_set.cell(row=i, column=1, value=p)
        ws_set.cell(row=i, column=2, value=m)
    header_row(ws_set, 10, ["KPI", "Target"])
    targets = [("SLA met %", 0.85, PCT), ("CSAT (1-5)", 4.0, DEC2),
               ("Reopen rate (max)", 0.06, PCT), ("Avg resolution hours (max)", 16, DEC1)]
    for i, (k, v, f) in enumerate(targets, start=11):
        ws_set.cell(row=i, column=1, value=k)
        ws_set.cell(row=i, column=2, value=v).number_format = f
    fmt(ws_set, "A5:B7")
    fmt(ws_set, "A11:B14")
    for r in list(range(5, 8)) + list(range(11, 15)):
        ws_set.cell(row=r, column=2).font = font(color="0000FF")
    ws_set["A16"] = "Targets agreed with the support team lead for H2 2026."
    ws_set["A16"].font = font(italic=True, color="595959")
    widths(ws_set, {"A": 30, "B": 28})
    T_SLA, T_CSAT, T_REOPEN, T_RES = "Settings!$B$11", "Settings!$B$12", "Settings!$B$13", "Settings!$B$14"

    # ---------- Agents ----------
    title(ws_a, "Agent list")
    header_row(ws_a, 3, ["Agent", "Team", "Start date"])
    for i, r in enumerate(agents.itertuples(index=False), start=4):
        ws_a.cell(row=i, column=1, value=r.agent)
        ws_a.cell(row=i, column=2, value=r.team)
        ws_a.cell(row=i, column=3, value=r.start_date.date()).number_format = DATE_FMT
    a_last = 3 + len(agents)
    fmt(ws_a, f"A4:C{a_last}")
    for r in range(4, a_last + 1):
        ws_a.cell(row=r, column=3).number_format = DATE_FMT
    widths(ws_a, {"A": 16, "B": 12, "C": 14})

    # ---------- Tickets (raw export + helper columns) ----------
    raw_cols = list(tickets.columns)
    helper_cols = ["created_date", "week_start", "month", "weekday_no", "first_response_min",
                   "resolution_hrs", "sla_target_min", "sla_met", "csat_clean", "is_duplicate",
                   "team", "resolution_bucket", "is_valid"]
    header_row(ws_t, 1, raw_cols)
    header_row(ws_t, 1, helper_cols, start_col=len(raw_cols) + 1, fill=HELPER_FILL)

    for r, row in enumerate(tickets.itertuples(index=False), start=2):
        for c, val in enumerate(row, start=1):
            if pd.isna(val):
                continue
            if isinstance(val, pd.Timestamp):
                val = val.to_pydatetime()
            ws_t.cell(row=r, column=c, value=val)

        ws_t[f"M{r}"] = f"=INT(B{r})"
        ws_t[f"N{r}"] = f"=M{r}-WEEKDAY(M{r},3)"
        ws_t[f"O{r}"] = f"=DATE(YEAR(M{r}),MONTH(M{r}),1)"
        ws_t[f"P{r}"] = f"=WEEKDAY(M{r},2)"
        ws_t[f"Q{r}"] = f"=ROUND((C{r}-B{r})*1440,1)"
        ws_t[f"R{r}"] = f'=IF(D{r}="","",ROUND((D{r}-B{r})*24,2))'
        ws_t[f"S{r}"] = f"=INDEX(Settings!$B$5:$B$7,MATCH(G{r},Settings!$A$5:$A$7,0))"
        ws_t[f"T{r}"] = f"=IF(Q{r}<=S{r},1,0)"
        ws_t[f"U{r}"] = f'=IF(AND(ISNUMBER(L{r}),L{r}>=1,L{r}<=5),L{r},"")'
        ws_t[f"V{r}"] = f"=IF(COUNTIF($A$2:A{r},A{r})>1,1,0)"
        ws_t[f"W{r}"] = (f'=IF(H{r}="","Unassigned",IFERROR(INDEX(Agents!$B$4:$B${a_last},'
                         f'MATCH(H{r},Agents!$A$4:$A${a_last},0)),"Unknown"))')
        ws_t[f"X{r}"] = (f'=IF(OR(R{r}="",Y{r}=0),"",IF(R{r}<4,"0-4h",IF(R{r}<24,"4-24h",'
                         f'IF(R{r}<72,"1-3 days","3+ days"))))')
        ws_t[f"Y{r}"] = f'=IF(AND(V{r}=0,H{r}<>"",OR(D{r}="",D{r}>=B{r})),1,0)'

    for col in "BCD":
        for c in ws_t[f"{col}2:{col}{last}"]:
            c[0].number_format = "yyyy-mm-dd hh:mm"
    for col in "MNO":
        for c in ws_t[f"{col}2:{col}{last}"]:
            c[0].number_format = "yyyy-mm-dd"
    ws_t.freeze_panes = "B2"
    ws_t.auto_filter.ref = f"A1:Y{last}"
    widths(ws_t, {get_column_letter(i): 13 for i in range(1, 26)})
    widths(ws_t, {"A": 12, "B": 16, "C": 16, "D": 16, "F": 16, "H": 11})
    for row in ws_t.iter_rows(min_row=2, max_row=last):
        for c in row:
            c.font = font()

    def rng(col):
        return f"Tickets!${col}$2:${col}${last}"

    WEEK, MONTH, WDAY, FRM, RES, SLA, CSAT, VALID = (rng(c) for c in "NOPQRTUY")
    CHANNEL, CATEGORY, PRIORITY, AGENT, STATUS, ESC, REOPEN = (rng(c) for c in "EFGHIJK")
    TEAM, BUCKET = rng("W"), rng("X")

    # ---------- Data quality ----------
    title(ws_dq, "Data quality checks", "Run on the raw helpdesk export before any KPI is calculated.")
    header_row(ws_dq, 4, ["Check", "Rows", "How it is handled"])
    checks = [
        ("Rows in export", f"=COUNTA({rng('A')})", ""),
        ("Duplicate ticket IDs (2nd+ occurrence)", f"=SUM({rng('V')})", "Excluded - same ticket exported twice"),
        ("Ticket without assigned agent", f"=COUNTBLANK({AGENT})", "Excluded - reported to helpdesk admin"),
        ("Resolved before created", f'=SUMPRODUCT(({rng("D")}<>"")*({rng("D")}<{rng("B")}))',
         "Excluded - broken timestamp"),
        ("CSAT score outside 1-5", f'=COUNTIFS({rng("L")},">5")+COUNTIFS({rng("L")},"<1")',
         "Ticket kept, score ignored (old 1-10 survey form)"),
        ("Rows used in the report", f"=SUM({VALID})", ""),
        ("Rows excluded", "=B5-B10", ""),
    ]
    for i, (name, formula, note) in enumerate(checks, start=5):
        ws_dq.cell(row=i, column=1, value=name)
        ws_dq.cell(row=i, column=2, value=formula)
        ws_dq.cell(row=i, column=3, value=note)
    fmt(ws_dq, "A5:C11")
    fmt(ws_dq, "B5:B11", INT_)
    fmt(ws_dq, "A10:C11", INT_, bold=True, fill=TOTAL_FILL)
    ws_dq["A12"] = "Excluded share"
    ws_dq["B12"] = "=B11/B5"
    fmt(ws_dq, "A12:C12")
    ws_dq["B12"].number_format = PCT
    widths(ws_dq, {"A": 40, "B": 12, "C": 50})

    # ---------- Weekly KPIs ----------
    title(ws_week, "Weekly support KPIs",
          "Weeks start on Monday. First and last weeks are partial (1-5 Jul, 28-30 Sep).")
    week_headers = ["Week start", "Tickets", "WoW change", "Resolved", "Avg first response (min)",
                    "SLA met %", "Avg resolution (h)", "CSAT", "CSAT responses", "Reopen rate",
                    "Escalation rate", "Technical issue share"]
    header_row(ws_week, 4, week_headers)
    first_w, n_weeks = 5, 14
    for i in range(n_weeks):
        r = first_w + i
        ws_week[f"A{r}"] = "=DATE(2026,6,29)" if i == 0 else f"=A{r-1}+7"
        crit = f"{WEEK},$A{r},{VALID},1"
        ws_week[f"B{r}"] = f"=COUNTIFS({crit})"
        ws_week[f"C{r}"] = "" if i == 0 else f'=IF(B{r-1}=0,"",B{r}/B{r-1}-1)'
        ws_week[f"D{r}"] = f'=COUNTIFS({crit},{STATUS},"Resolved")'
        ws_week[f"E{r}"] = f'=IFERROR(AVERAGEIFS({FRM},{crit}),"")'
        ws_week[f"F{r}"] = f'=IFERROR(AVERAGEIFS({SLA},{crit}),"")'
        ws_week[f"G{r}"] = f'=IFERROR(AVERAGEIFS({RES},{crit},{STATUS},"Resolved"),"")'
        ws_week[f"H{r}"] = f'=IFERROR(AVERAGEIFS({CSAT},{crit}),"")'
        ws_week[f"I{r}"] = f'=COUNTIFS({crit},{CSAT},">=1")'
        ws_week[f"J{r}"] = f'=IFERROR(COUNTIFS({crit},{REOPEN},"Yes")/D{r},"")'
        ws_week[f"K{r}"] = f'=IFERROR(COUNTIFS({crit},{ESC},"Yes")/B{r},"")'
        ws_week[f"L{r}"] = f'=IFERROR(COUNTIFS({crit},{CATEGORY},"Technical issue")/B{r},"")'
    lw = first_w + n_weeks - 1
    tr = lw + 1
    ws_week[f"A{tr}"] = "Q3 total"
    ws_week[f"B{tr}"] = f"=SUM(B{first_w}:B{lw})"
    ws_week[f"D{tr}"] = f"=SUM(D{first_w}:D{lw})"
    ws_week[f"E{tr}"] = f"=AVERAGEIFS({FRM},{VALID},1)"
    ws_week[f"F{tr}"] = f"=AVERAGEIFS({SLA},{VALID},1)"
    ws_week[f"G{tr}"] = f'=AVERAGEIFS({RES},{VALID},1,{STATUS},"Resolved")'
    ws_week[f"H{tr}"] = f"=AVERAGEIFS({CSAT},{VALID},1)"
    ws_week[f"I{tr}"] = f"=SUM(I{first_w}:I{lw})"
    ws_week[f"J{tr}"] = f'=COUNTIFS({VALID},1,{REOPEN},"Yes")/D{tr}'
    ws_week[f"K{tr}"] = f'=COUNTIFS({VALID},1,{ESC},"Yes")/B{tr}'
    ws_week[f"L{tr}"] = f'=COUNTIFS({VALID},1,{CATEGORY},"Technical issue")/B{tr}'

    fmt(ws_week, f"A{first_w}:L{tr}")
    for col, f in zip("ABCDEFGHIJKL", [DATE_FMT, INT_, PCT, INT_, DEC1, PCT, DEC1, DEC2, INT_, PCT, PCT, PCT]):
        for c in ws_week[f"{col}{first_w}:{col}{tr}"]:
            c[0].number_format = f
    fmt(ws_week, f"A{tr}:L{tr}", bold=True, fill=TOTAL_FILL)
    ws_week[f"A{tr}"].number_format = "General"
    for col, f in zip("BCDEFGHIJKL", [INT_, PCT, INT_, DEC1, PCT, DEC1, DEC2, INT_, PCT, PCT, PCT]):
        ws_week[f"{col}{tr}"].number_format = f

    ws_week.conditional_formatting.add(f"F{first_w}:F{lw}", CellIsRule(operator="lessThan", formula=[T_SLA], fill=RED_FILL))
    ws_week.conditional_formatting.add(f"H{first_w}:H{lw}", CellIsRule(operator="lessThan", formula=[T_CSAT], fill=RED_FILL))
    ws_week.conditional_formatting.add(f"J{first_w}:J{lw}", CellIsRule(operator="greaterThan", formula=[T_REOPEN], fill=RED_FILL))
    ws_week.conditional_formatting.add(f"G{first_w}:G{lw}", CellIsRule(operator="greaterThan", formula=[T_RES], fill=RED_FILL))
    ws_week[f"A{tr + 2}"] = "Red cells miss the target set on the Settings sheet."
    ws_week[f"A{tr + 2}"].font = font(italic=True, color="595959")
    ws_week.freeze_panes = "B5"
    widths(ws_week, {"A": 14, **{c: 13 for c in "BCDEFGHIJKL"}})

    ws_week.add_chart(weekly_combo_chart(ws_week, first_w, lw, 22), f"A{tr + 4}")

    # ---------- Agent performance ----------
    title(ws_agent, "Agent performance - Q3 2026", "Valid tickets only. Rank 1 = highest CSAT.")
    header_row(ws_agent, 4, ["Agent", "Team", "Start date", "Tickets", "SLA met %", "Avg resolution (h)",
                             "CSAT", "Reopen rate", "Escalation rate", "CSAT rank"])
    fa = 5
    la = fa + len(agents) - 1
    for i in range(len(agents)):
        r, src = fa + i, 4 + i
        crit = f"{AGENT},$A{r},{VALID},1"
        ws_agent[f"A{r}"] = f"=Agents!A{src}"
        ws_agent[f"B{r}"] = f"=Agents!B{src}"
        ws_agent[f"C{r}"] = f"=Agents!C{src}"
        ws_agent[f"D{r}"] = f"=COUNTIFS({crit})"
        ws_agent[f"E{r}"] = f'=IFERROR(AVERAGEIFS({SLA},{crit}),"")'
        ws_agent[f"F{r}"] = f'=IFERROR(AVERAGEIFS({RES},{crit},{STATUS},"Resolved"),"")'
        ws_agent[f"G{r}"] = f'=IFERROR(AVERAGEIFS({CSAT},{crit}),"")'
        ws_agent[f"H{r}"] = f'=IFERROR(COUNTIFS({crit},{REOPEN},"Yes")/COUNTIFS({crit},{STATUS},"Resolved"),"")'
        ws_agent[f"I{r}"] = f'=IFERROR(COUNTIFS({crit},{ESC},"Yes")/D{r},"")'
        ws_agent[f"J{r}"] = f"=RANK(G{r},$G${fa}:$G${la},0)"
    fmt(ws_agent, f"A{fa}:J{la}")
    for col, f in zip("CDEFGHIJ", [DATE_FMT, INT_, PCT, DEC1, DEC2, PCT, PCT, "0"]):
        for c in ws_agent[f"{col}{fa}:{col}{la}"]:
            c[0].number_format = f
    ws_agent.conditional_formatting.add(f"G{fa}:G{la}", ColorScaleRule(start_type="min", start_color="F8696B",
                                        mid_type="percentile", mid_value=50, mid_color="FFEB84",
                                        end_type="max", end_color="63BE7B"))
    ws_agent.conditional_formatting.add(f"H{fa}:H{la}", ColorScaleRule(start_type="min", start_color="63BE7B",
                                        mid_type="percentile", mid_value=50, mid_color="FFEB84",
                                        end_type="max", end_color="F8696B"))

    tt = la + 3
    ws_agent[f"A{tt - 1}"] = "By team"
    ws_agent[f"A{tt - 1}"].font = font(bold=True, color=NAVY, size=11)
    header_row(ws_agent, tt, ["Team", "Agents", "", "Tickets", "SLA met %", "Avg resolution (h)",
                              "CSAT", "Reopen rate", "Escalation rate"])
    for i, team in enumerate(["Team A", "Team B"]):
        r = tt + 1 + i
        crit = f'{TEAM},$A{r},{VALID},1'
        ws_agent[f"A{r}"] = team
        ws_agent[f"B{r}"] = f"=COUNTIF(Agents!$B$4:$B${a_last},A{r})"
        ws_agent[f"D{r}"] = f"=COUNTIFS({crit})"
        ws_agent[f"E{r}"] = f"=AVERAGEIFS({SLA},{crit})"
        ws_agent[f"F{r}"] = f'=AVERAGEIFS({RES},{crit},{STATUS},"Resolved")'
        ws_agent[f"G{r}"] = f"=AVERAGEIFS({CSAT},{crit})"
        ws_agent[f"H{r}"] = f'=COUNTIFS({crit},{REOPEN},"Yes")/COUNTIFS({crit},{STATUS},"Resolved")'
        ws_agent[f"I{r}"] = f'=COUNTIFS({crit},{ESC},"Yes")/D{r}'
    fmt(ws_agent, f"A{tt + 1}:I{tt + 2}")
    for col, f in zip("BDEFGHI", ["0", INT_, PCT, DEC1, DEC2, PCT, PCT]):
        for c in ws_agent[f"{col}{tt + 1}:{col}{tt + 2}"]:
            c[0].number_format = f
    widths(ws_agent, {"A": 14, "B": 10, "C": 13, **{c: 13 for c in "DEFGHIJ"}})
    ws_agent.freeze_panes = "B5"

    # ---------- Breakdowns ----------
    title(ws_brk, "Breakdowns - Q3 2026", "Valid tickets only.")
    channels = ["Chat", "Email", "Phone"]
    categories = ["Payments", "Technical issue", "Account access", "Product question", "Refund request"]

    def section(row, text):
        ws_brk[f"A{row}"] = text
        ws_brk[f"A{row}"].font = font(bold=True, color=NAVY, size=11)

    # tickets by category x channel
    r0 = 4
    section(r0, "1. Ticket volume by category and channel")
    header_row(ws_brk, r0 + 1, ["Category"] + channels + ["Total", "Share"])
    for i, cat in enumerate(categories):
        r = r0 + 2 + i
        ws_brk[f"A{r}"] = cat
        for j, ch in enumerate(channels):
            col = get_column_letter(2 + j)
            ws_brk[f"{col}{r}"] = f'=COUNTIFS({CATEGORY},$A{r},{CHANNEL},{col}${r0 + 1},{VALID},1)'
        ws_brk[f"E{r}"] = f"=SUM(B{r}:D{r})"
        ws_brk[f"F{r}"] = f"=E{r}/$E${r0 + 2 + len(categories)}"
    rt = r0 + 2 + len(categories)
    ws_brk[f"A{rt}"] = "Total"
    for col in "BCDE":
        ws_brk[f"{col}{rt}"] = f"=SUM({col}{r0 + 2}:{col}{rt - 1})"
    ws_brk[f"F{rt}"] = f"=E{rt}/$E${rt}"
    fmt(ws_brk, f"A{r0 + 2}:F{rt}", INT_)
    fmt(ws_brk, f"A{rt}:F{rt}", INT_, bold=True, fill=TOTAL_FILL)
    for c in ws_brk[f"F{r0 + 2}:F{rt}"]:
        c[0].number_format = PCT

    # category quality
    r1 = rt + 3
    section(r1, "2. Resolution quality by category")
    header_row(ws_brk, r1 + 1, ["Category", "Tickets", "Avg resolution (h)", "Escalation rate", "Reopen rate", "CSAT"])
    for i, cat in enumerate(categories):
        r = r1 + 2 + i
        crit = f"{CATEGORY},$A{r},{VALID},1"
        ws_brk[f"A{r}"] = cat
        ws_brk[f"B{r}"] = f"=COUNTIFS({crit})"
        ws_brk[f"C{r}"] = f'=AVERAGEIFS({RES},{crit},{STATUS},"Resolved")'
        ws_brk[f"D{r}"] = f'=COUNTIFS({crit},{ESC},"Yes")/B{r}'
        ws_brk[f"E{r}"] = f'=COUNTIFS({crit},{REOPEN},"Yes")/COUNTIFS({crit},{STATUS},"Resolved")'
        ws_brk[f"F{r}"] = f"=AVERAGEIFS({CSAT},{crit})"
    r1e = r1 + 1 + len(categories)
    fmt(ws_brk, f"A{r1 + 2}:F{r1e}")
    for col, f in zip("BCDEF", [INT_, DEC1, PCT, PCT, DEC2]):
        for c in ws_brk[f"{col}{r1 + 2}:{col}{r1e}"]:
            c[0].number_format = f

    # SLA by priority x channel
    r2 = r1e + 3
    section(r2, "3. First response SLA met % by priority and channel")
    header_row(ws_brk, r2 + 1, ["Priority"] + channels + ["All channels", "SLA target (min)"])
    for i, pr in enumerate(["High", "Medium", "Low"]):
        r = r2 + 2 + i
        ws_brk[f"A{r}"] = pr
        for j in range(3):
            col = get_column_letter(2 + j)
            ws_brk[f"{col}{r}"] = f"=AVERAGEIFS({SLA},{PRIORITY},$A{r},{CHANNEL},{col}${r2 + 1},{VALID},1)"
        ws_brk[f"E{r}"] = f"=AVERAGEIFS({SLA},{PRIORITY},$A{r},{VALID},1)"
        ws_brk[f"F{r}"] = f"=INDEX(Settings!$B$5:$B$7,MATCH(A{r},Settings!$A$5:$A$7,0))"
    r2e = r2 + 4
    fmt(ws_brk, f"A{r2 + 2}:F{r2e}", PCT)
    for c in ws_brk[f"F{r2 + 2}:F{r2e}"]:
        c[0].number_format = "0"
    ws_brk.conditional_formatting.add(f"B{r2 + 2}:E{r2e}", CellIsRule(operator="lessThan", formula=[T_SLA], fill=RED_FILL))
    ws_brk.conditional_formatting.add(f"B{r2 + 2}:E{r2e}", CellIsRule(operator="greaterThanOrEqual", formula=[T_SLA], fill=GREEN_FILL))

    # weekday
    r3 = r2e + 3
    section(r3, "4. Weekday pattern")
    header_row(ws_brk, r3 + 1, ["Weekday", "Day no.", "Tickets", "Avg first response (min)", "SLA met %"])
    for i, wd in enumerate(["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]):
        r = r3 + 2 + i
        crit = f"{WDAY},$B{r},{VALID},1"
        ws_brk[f"A{r}"] = wd
        ws_brk[f"B{r}"] = i + 1
        ws_brk[f"C{r}"] = f"=COUNTIFS({crit})"
        ws_brk[f"D{r}"] = f"=AVERAGEIFS({FRM},{crit})"
        ws_brk[f"E{r}"] = f"=AVERAGEIFS({SLA},{crit})"
    r3e = r3 + 8
    fmt(ws_brk, f"A{r3 + 2}:E{r3e}")
    for col, f in zip("BCDE", ["0", INT_, DEC1, PCT]):
        for c in ws_brk[f"{col}{r3 + 2}:{col}{r3e}"]:
            c[0].number_format = f
    ws_brk.conditional_formatting.add(f"E{r3 + 2}:E{r3e}", CellIsRule(operator="lessThan", formula=[T_SLA], fill=RED_FILL))

    # resolution time vs CSAT
    r4 = r3e + 3
    section(r4, "5. Does resolution time affect CSAT?")
    header_row(ws_brk, r4 + 1, ["Resolution time", "Resolved tickets", "Rated tickets", "Avg CSAT", "Response rate"])
    for i, b in enumerate(["0-4h", "4-24h", "1-3 days", "3+ days"]):
        r = r4 + 2 + i
        ws_brk[f"A{r}"] = b
        ws_brk[f"B{r}"] = f"=COUNTIFS({BUCKET},$A{r})"
        ws_brk[f"C{r}"] = f'=COUNTIFS({BUCKET},$A{r},{CSAT},">=1")'
        ws_brk[f"D{r}"] = f"=AVERAGEIFS({CSAT},{BUCKET},$A{r})"
        ws_brk[f"E{r}"] = f"=C{r}/B{r}"
    r4e = r4 + 5
    fmt(ws_brk, f"A{r4 + 2}:E{r4e}")
    for col, f in zip("BCDE", [INT_, INT_, DEC2, PCT]):
        for c in ws_brk[f"{col}{r4 + 2}:{col}{r4e}"]:
            c[0].number_format = f

    widths(ws_brk, {"A": 20, "B": 13, "C": 14, "D": 16, "E": 14, "F": 16})

    wd_chart = BarChart()
    wd_chart.title = "SLA met % by weekday"
    wd_chart.add_data(Reference(ws_brk, min_col=5, min_row=r3 + 1, max_row=r3e), titles_from_data=True)
    wd_chart.set_categories(Reference(ws_brk, min_col=1, min_row=r3 + 2, max_row=r3e))
    wd_chart.y_axis.number_format = "0%"
    wd_chart.y_axis.scaling.min = 0
    wd_chart.legend = None
    chart_style(wd_chart)
    wd_chart.height, wd_chart.width = 7, 14
    ws_brk.add_chart(wd_chart, "H4")

    csat_chart = BarChart()
    csat_chart.title = "Avg CSAT by resolution time"
    csat_chart.add_data(Reference(ws_brk, min_col=4, min_row=r4 + 1, max_row=r4e), titles_from_data=True)
    csat_chart.set_categories(Reference(ws_brk, min_col=1, min_row=r4 + 2, max_row=r4e))
    csat_chart.y_axis.scaling.min = 1
    csat_chart.y_axis.scaling.max = 5
    csat_chart.legend = None
    chart_style(csat_chart)
    csat_chart.height, csat_chart.width = 7, 14
    ws_brk.add_chart(csat_chart, "H20")

    # ---------- Summary ----------
    title(ws_sum, "Customer Support Operations - Q3 2026 KPI Report",
          "Period: 1 Jul - 30 Sep 2026  |  Source: helpdesk ticket export (see Tickets sheet)  |  Valid tickets only")
    header_row(ws_sum, 4, ["KPI", "Q3 actual", "Target", "Status", "July", "August", "September"])
    months = [date(2026, 7, 1), date(2026, 8, 1), date(2026, 9, 1)]

    def month_formula(kind, m):
        d = f"DATE({m.year},{m.month},1)"
        crit = f"{MONTH},{d},{VALID},1"
        return {
            "tickets": f"=COUNTIFS({crit})",
            "sla": f"=AVERAGEIFS({SLA},{crit})",
            "frm": f"=AVERAGEIFS({FRM},{crit})",
            "res": f'=AVERAGEIFS({RES},{crit},{STATUS},"Resolved")',
            "csat": f"=AVERAGEIFS({CSAT},{crit})",
            "reopen": f'=COUNTIFS({crit},{REOPEN},"Yes")/COUNTIFS({crit},{STATUS},"Resolved")',
            "esc": f'=COUNTIFS({crit},{ESC},"Yes")/COUNTIFS({crit})',
        }[kind]

    wtr = f"Weekly_KPIs!"
    kpis = [
        ("Tickets received", "tickets", f"={wtr}B{tr}", None, None, INT_),
        ("First response SLA met %", "sla", f"={wtr}F{tr}", f"={T_SLA}", "higher", PCT),
        ("Avg first response (min)", "frm", f"={wtr}E{tr}", None, None, DEC1),
        ("Avg resolution time (h)", "res", f"={wtr}G{tr}", f"={T_RES}", "lower", DEC1),
        ("CSAT (1-5)", "csat", f"={wtr}H{tr}", f"={T_CSAT}", "higher", DEC2),
        ("Reopen rate", "reopen", f"={wtr}J{tr}", f"={T_REOPEN}", "lower", PCT),
        ("Escalation rate", "esc", f"={wtr}K{tr}", None, None, PCT),
    ]
    for i, (name, kind, actual, target, direction, f) in enumerate(kpis, start=5):
        ws_sum[f"A{i}"] = name
        ws_sum[f"B{i}"] = actual
        if target:
            ws_sum[f"C{i}"] = target
            op = ">=" if direction == "higher" else "<="
            ws_sum[f"D{i}"] = f'=IF(B{i}{op}C{i},"On target","Below target")'
        for j, m in enumerate(months):
            ws_sum.cell(row=i, column=5 + j, value=month_formula(kind, m))
    fmt(ws_sum, "A5:G11")
    for i, k in enumerate(kpis, start=5):
        for col in "BCEFG":
            ws_sum[f"{col}{i}"].number_format = k[5]
        ws_sum[f"B{i}"].font = font(bold=True)
    ws_sum.conditional_formatting.add("D5:D11", CellIsRule(operator="equal", formula=['"Below target"'], fill=RED_FILL))
    ws_sum.conditional_formatting.add("D5:D11", CellIsRule(operator="equal", formula=['"On target"'], fill=GREEN_FILL))

    ws_sum["A13"] = "Key findings"
    ws_sum["A13"].font = font(bold=True, color=NAVY, size=12)
    ws_sum["A21"] = "Recommendations"
    ws_sum["A21"].font = font(bold=True, color=NAVY, size=12)
    widths(ws_sum, {"A": 30, "B": 13, "C": 11, "D": 14, "E": 12, "F": 12, "G": 12})

    ws_sum.add_chart(weekly_combo_chart(ws_week, first_w, lw, 17), "I4")
    ws_sum.print_area = "A1:R27"
    ws_week.print_area = f"A1:L{tr + 20}"
    ws_brk.print_area = f"A1:N{r4e}"

    for ws in wb.worksheets:
        ws.sheet_view.showGridLines = ws.title in ("Tickets",)
        if ws.title != "Tickets":
            print_setup(ws)
    ws_sum.sheet_properties.tabColor = NAVY

    return wb


def add_text(wb, findings, recommendations):
    ws = wb["Summary"]
    for start, items in ((14, findings), (22, recommendations)):
        for i, text in enumerate(items):
            r = start + i
            c = ws.cell(row=r, column=1, value=f"{i + 1}. {text}")
            c.font = font()
            c.alignment = Alignment(wrap_text=True, vertical="top")
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=7)
            ws.row_dimensions[r].height = 30


FINDINGS = [
    "SLA met is 83.9% against an 85% target. The gap comes almost entirely from email: only 34% of High and "
    "44% of Medium priority emails got a first response in time, while chat and phone are at 97-100%.",
    "Weekends are understaffed: SLA drops to 71% on Saturday and 75% on Sunday (86-89% on weekdays) and the "
    "average first response is about 2.5 times slower.",
    "The product release in the week of 10 Aug caused a 69% jump in tickets (668 vs ~400). Technical issues "
    "went from ~23% to 42% of volume, resolution time rose to 24.9h and CSAT fell to 3.91.",
    "New agent Lukas M. (joined 3 Aug) has CSAT 3.37 vs 4.10 overall and a 16.5% reopen rate vs 3-8% for the "
    "rest of the team. Without these tickets the reopen rate would be 6.3%, close to the 6% target.",
    "Speed drives satisfaction: tickets solved within 4 hours get CSAT 4.43, tickets open for more than 3 days "
    "get 2.99. Refund requests are the slowest category (47.9h on average, CSAT 3.76).",
    "Data quality: 40 of 5,822 exported rows (0.7%) were excluded - duplicates, tickets without an agent and "
    "broken timestamps. See the Data_Quality sheet.",
]
RECOMMENDATIONS = [
    "Route High and Medium priority emails into a separate queue with a clear owner during each shift.",
    "Add weekend coverage (at least one extra agent on Saturday and Sunday shifts) and track weekend SLA "
    "as a separate KPI.",
    "Agree a release checklist with the product team: support gets release notes and ready answers before "
    "launch, and extra staff for the first 3 days.",
    "Set up a mentoring plan for Lukas M. with Anna K. (best CSAT 4.42, reopen rate 3.0%) and review "
    "reopened tickets weekly.",
    "Review the refund process with the payments team - it is the main source of long-running tickets.",
]

if __name__ == "__main__":
    wb = build()
    add_text(wb, FINDINGS, RECOMMENDATIONS)
    wb.save(OUT)
    print(f"saved {OUT.name}")
