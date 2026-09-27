"""DASHBOARD: profile, selected-date snapshot, macros, 7-day trend, habit KPIs, micronutrients."""
import datetime as dt

from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import PatternFill

from . import styles as st

SHEET = "DASHBOARD"
GREEN, YELLOW, RED = (PatternFill("solid", fgColor=c) for c in ("C6EFCE", "FFEB9C", "FFC7CE"))
ORANGE, BLUE = PatternFill("solid", fgColor="F8CBAD"), PatternFill("solid", fgColor="DDEBF7")
BIG = st.font(14, True, color="1F4E79")


def bar(ws, row, c1, c2, text):
    ws.merge_cells(start_row=row, start_column=c1, end_row=row, end_column=c2)
    for c in range(c1, c2 + 1):
        ws.cell(row=row, column=c).fill = st.FILL_SECTION
    cell = ws.cell(row=row, column=c1, value=text)
    cell.font, cell.alignment = st.font(12, True, color="FFFFFF"), st.CENTER


def kv(ws, row, col, label, formula, fmt=None, big=True):
    st.put(ws, row, col, label, "label")
    c = st.put(ws, row, col + 1, formula, "total", fmt)
    if big:
        c.font = BIG
    return c


def build(wb, ctx):
    ws, R = wb[SHEET], ctx.refs
    ws.sheet_properties.tabColor = "1F4E79"
    st.widths(ws, {"A": 3, "B": 26, "C": 16, "D": 16, "E": 3, "F": 24, "G": 16, "H": 16, "I": 16})
    ws.merge_cells("B2:I2")
    ws["B2"] = '=IFERROR("🏋️ "&UPPER(%s)&"\'S MEGATRACKER – NUTRITION, TRAINING & HABITS","MEGATRACKER")' % R["name"]
    ws["B2"].font, ws["B2"].alignment = st.font(20, True, color="1F4E79"), st.CENTER
    ws.merge_cells("B3:I3")
    ws["B3"] = ('=IFERROR("Goal: "&%s&"  |  Target "&TEXT(%s,"#,##0")&" kcal/day  |  Protein "&TEXT(%s,"0")&" g/day","")'
                % (R["goal"], R["cal_target"], R["protein_target"]))
    ws["B3"].font, ws["B3"].alignment = st.font(11, italic=True, color="595959"), st.CENTER

    # A. Profile & goal
    bar(ws, 5, 2, 4, "👤 PROFILE & GOAL")
    latest_wk = "INDEX(tblBody[Weekly_Change_kg],MATCH(9.99E+307,tblBody[Weight_kg]))"
    for i, (label, f, fmt) in enumerate([
            ("Name", "=" + R["name"], None), ("Age", "=" + R["age"], "0"), ("Height", "=" + R["height"], '0 "cm"'),
            ("Current weight", "=" + R["weight"], '0.0 "kg"'), ("Body fat", "=" + R["body_fat"], '0 "%"'),
            ("Goal", "=" + R["goal"], None), ("Weekly weight change", '=IFERROR(%s,"")' % latest_wk, '+0.00 "kg";-0.00 "kg"')]):
        c = kv(ws, 6 + i, 2, label, f, fmt, big=False)
        ws.merge_cells("C%d:D%d" % (6 + i, 6 + i))
        if label == "Goal":
            c.alignment, c.font = st.WRAP, st.FONT_BODY
            ws.row_dimensions[6 + i].height = 30
    status = ('=IFERROR(IF(NOT(ISNUMBER({w})),"— log 2 weeks of weigh-ins",IF(AND({w}>={lo},{w}<={hi}),"✅ ON TRACK",'
              'IF({w}<{lo},"⚠️ BELOW TARGET RANGE","🔴 ABOVE TARGET RANGE"))),"—")').format(w=latest_wk, lo=R["wk_min"], hi=R["wk_max"])
    kv(ws, 13, 2, "Status", status, big=False)
    ws.merge_cells("C13:D13")
    for word, fill in (("ON TRACK", GREEN), ("BELOW", ORANGE), ("ABOVE", RED)):
        ws.conditional_formatting.add("C13", FormulaRule(formula=['ISNUMBER(SEARCH("%s",$C$13))' % word], fill=fill))

    # B. Snapshot for the selected date
    bar(ws, 5, 6, 9, "📅 SNAPSHOT FOR SELECTED DATE")
    st.put(ws, 6, 6, "Select date →", "label")
    dates = ctx.samples["dates"]
    d = st.put(ws, 6, 7, dates[-1] if dates else dt.date.today(), "input", "YYYY-MM-DD", bold=True)
    st.dropdown(ws, "=" + R["total_dates"], "G6")
    st.put(ws, 6, 8, "pick a logged date", "note")
    snap = [("Calories IN", "tblDailyTotals[Total_Calories],tblDailyTotals[Date]", "#,##0"),
            ("Calories OUT (TDEE)", "tblActivity[TDEE_Day],tblActivity[Date]", "#,##0"),
            ("Surplus / deficit", None, "+#,##0;-#,##0"),
            ("Protein", "tblDailyTotals[Total_Protein_g],tblDailyTotals[Date]", '0 "g"'),
            ("Carbs", "tblDailyTotals[Total_Carbs_g],tblDailyTotals[Date]", '0 "g"'),
            ("Fat", "tblDailyTotals[Total_Fat_g],tblDailyTotals[Date]", '0 "g"')]
    for i, (label, src, fmt) in enumerate(snap):
        f = "=IFERROR(G7-G8,0)" if src is None else "=IFERROR(SUMIFS(%s,$G$6),0)" % src
        kv(ws, 7 + i, 6, label, f, fmt)
    for r, key in ((7, "cal_target"), (10, "protein_target"), (11, "carb_target"), (12, "fat_target")):
        st.put(ws, r, 8, '=IFERROR("target "&TEXT(%s,"#,##0"),"")' % R[key], "note")
    ws.conditional_formatting.add("G9", CellIsRule(operator="lessThan", formula=["0"], fill=BLUE))
    ws.conditional_formatting.add("G9", CellIsRule(operator="between", formula=["100", "400"], fill=GREEN))
    ws.conditional_formatting.add("G9", CellIsRule(operator="greaterThan", formula=["400"], fill=ORANGE))
    st.put(ws, 13, 6, "Surplus colours: 100-400 green, >400 orange, deficit blue", "note")

    # C. Macro split & targets
    bar(ws, 15, 2, 4, "🥧 MACRO SPLIT & TARGETS")
    for i, (label, f, fmt) in enumerate([
            ("% calories from protein", "=IFERROR(G10*4/G7,0)", "0.0%"), ("% from carbs", "=IFERROR(G11*4/G7,0)", "0.0%"),
            ("% from fat", "=IFERROR(G12*9/G7,0)", "0.0%"), ("Target calories", "=" + R["cal_target"], "#,##0"),
            ("Calorie gap (IN - target)", "=IFERROR(G7-C19,\"\")", "+#,##0;-#,##0"),
            ("Target protein", "=" + R["protein_target"], '0 "g"'), ("Protein gap", "=IFERROR(G10-C21,\"\")", '+0 "g";-0 "g"')]):
        kv(ws, 16 + i, 2, label, f, fmt, big=False)

    # D. 7-day trend
    bar(ws, 15, 6, 9, "📈 7-DAY TREND (to selected date)")
    win = '{t},">"&($G$6-7),{d},"<="&$G$6'
    act_win = win.format(t="tblActivity[Date]", d="tblActivity[Date]")
    tot_win = win.format(t="tblDailyTotals[Date]", d="tblDailyTotals[Date]")
    kv(ws, 16, 6, "Avg surplus (kcal/day)", '=IFERROR(AVERAGEIFS(tblActivity[Surplus_Day],%s),"N/A")' % act_win, "+#,##0;-#,##0")
    kv(ws, 17, 6, "Projected change (kg/month)", '=IFERROR(G16*30/%s,"N/A")' % R["kcal_per_kg"], "+0.00;-0.00")
    kv(ws, 18, 6, "Avg calories IN", '=IFERROR(AVERAGEIFS(tblDailyTotals[Total_Calories],%s,tblDailyTotals[Total_Calories],">0"),"N/A")' % tot_win, "#,##0", big=False)
    kv(ws, 19, 6, "Avg protein (g)", '=IFERROR(AVERAGEIFS(tblDailyTotals[Total_Protein_g],%s,tblDailyTotals[Total_Calories],">0"),"N/A")' % tot_win, "0", big=False)
    ws.merge_cells("F20:I22")
    c = ws.cell(row=20, column=6, value='=IFERROR("At your current average surplus of "&TEXT(G16,"+#,##0;-#,##0")&" kcal/day, '
                                         'expected change ≈ "&TEXT(G17,"+0.00;-0.00")&" kg/month.","Log food and activity to see the trend.")')
    c.font, c.alignment = st.font(11, italic=True, color="1F4E79"), st.WRAP

    # E. Last 7 days table
    bar(ws, 24, 2, 9, "🗓 LAST 7 DAYS")
    for col, h in zip([2, 3, 4, 6, 7, 8, 9], ["Date", "Calories IN", "Calories OUT", "Surplus", "Protein (g)", "Sleep (h)", "Screen (h)"]):
        st.headers(ws, 25, col, [h])
    for i in range(7):
        r = 26 + i
        st.put(ws, r, 2, "=$G$6-%d" % (6 - i), "calc", "ddd DD-MMM")
        st.put(ws, r, 3, "=IFERROR(SUMIFS(tblDailyTotals[Total_Calories],tblDailyTotals[Date],B%d),0)" % r, "calc", "#,##0")
        st.put(ws, r, 4, "=IFERROR(SUMIFS(tblActivity[TDEE_Day],tblActivity[Date],B%d),0)" % r, "calc", "#,##0")
        st.put(ws, r, 6, '=IF(OR(C{r}=0,D{r}=0),"",C{r}-D{r})'.format(r=r), "total", "+#,##0;-#,##0")
        st.put(ws, r, 7, "=IFERROR(SUMIFS(tblDailyTotals[Total_Protein_g],tblDailyTotals[Date],B%d),0)" % r, "calc", "0")
        st.put(ws, r, 8, '=IFERROR(AVERAGEIFS(tblHabits[Sleep_hrs],tblHabits[Date],B%d),"")' % r, "calc", "0.0")
        st.put(ws, r, 9, '=IFERROR(AVERAGEIFS(tblHabits[Screen_Time_hrs],tblHabits[Date],B%d),"")' % r, "calc", "0.0")
    for rule in (CellIsRule(operator="lessThan", formula=["0"], fill=BLUE),
                 CellIsRule(operator="between", formula=["100", "400"], fill=GREEN),
                 CellIsRule(operator="greaterThan", formula=["400"], fill=ORANGE)):
        ws.conditional_formatting.add("F26:F32", rule)

    # F. Habit KPIs (new)
    bar(ws, 34, 2, 9, "🧠 HABIT KPIs (7 days to selected date)")
    for col, h in zip([2, 3, 4, 6, 7], ["Habit", "Value", "Unit", "Target", "Status"]):
        st.headers(ws, 35, col, [h])
    hwin = win.format(t="tblHabits[Date]", d="tblHabits[Date]")
    hab = R["habit_col"]
    streak = ('=IFERROR(INDEX(tblHabits[Habit_Streak],MATCH($G$6,tblHabits[Date],0)),'
              'IFERROR(INDEX(tblHabits[Habit_Streak],MATCH(MAX(tblHabits[Date]),tblHabits[Date],0)),0))')
    kpis = [("Screen time (7-day avg)", '=IFERROR(AVERAGEIFS(tblHabits[Screen_Time_hrs],%s),"—")' % hwin, "h/day",
             "=" + R["screen_max"], '"≤ "0.0" h"', '=IF(ISNUMBER(C{r}),IF(C{r}<=F{r},"✅ On target","⚠️ Over limit"),"—")'),
            ("Sleep (7-day avg)", '=IFERROR(AVERAGEIFS(tblHabits[Sleep_hrs],%s),"—")' % hwin, "h/night",
             "=" + R["sleep_min"], '"≥ "0.0" h"', '=IF(ISNUMBER(C{r}),IF(C{r}>=F{r},"✅ On target","⚠️ Below target"),"—")'),
            ("Deep work (7-day avg)", '=IFERROR(AVERAGEIFS(tblHabits[Deep_Work_hrs],%s),"—")' % hwin, "h/day",
             "=" + R["deep_min"], '"≥ "0.0" h"', '=IF(ISNUMBER(C{r}),IF(C{r}>=F{r},"✅ On target","⚠️ Below target"),"—")'),
            ("%s streak" % hab.replace("_", " "), streak, "days",
             '=IFERROR(MAX(tblHabits[Habit_Streak]),0)', '"best "0" d"', '=IF(N(C{r})>=7,"🔥 7+ day streak",IF(N(C{r})>0,"✅ Keep going","⚠️ Start today"))')]
    for i, (label, f, unit, target, tfmt, status_f) in enumerate(kpis):
        r = 36 + i
        st.put(ws, r, 2, label, "label")
        st.put(ws, r, 3, f, "total", "0" if i == 3 else "0.0").font = BIG
        st.put(ws, r, 4, unit, "note")
        st.put(ws, r, 6, target, "calc", tfmt)
        st.put(ws, r, 7, status_f.format(r=r), "calc", bold=True)
        ws.merge_cells("G%d:I%d" % (r, r))
    for word, fill in (("✅", GREEN), ("🔥", GREEN), ("⚠️", YELLOW)):
        ws.conditional_formatting.add("G36:G39", FormulaRule(formula=['ISNUMBER(SEARCH("%s",G36))' % word], fill=fill))

    # G. Micronutrients
    bar(ws, 41, 2, 9, "🧬 MICRONUTRIENTS (selected date, food + supplements)")
    for col, h in zip([2, 3, 4, 6, 7, 8], ["Nutrient", "Intake", "Unit", "Daily target", "% of target", "Status"]):
        st.headers(ws, 42, col, [h])
    labels = ["VitD", "Ca", "Mg", "Fe", "B12", "Om3"]
    for i, (name, unit, _) in enumerate(ctx.cfg["rda"][:6]):
        r = 43 + i
        st.put(ws, r, 2, name, "label")
        st.put(ws, r, 3, "=IFERROR(INDEX(tblMicrosDaily[%s_Total],MATCH($G$6,tblMicrosDaily[Date],0)),0)" % labels[i], "calc", "#,##0.0")
        st.put(ws, r, 4, unit, "note")
        st.put(ws, r, 6, "=" + R["rda"][i], "calc", "#,##0.0")
        st.put(ws, r, 7, "=IFERROR(C%d/F%d,0)" % (r, r), "total", "0%")
        st.put(ws, r, 8, '=IF(G{r}>=1,"✅ ≥100%",IF(G{r}>=0.7,"⚠️ 70-99%","🔴 <70%"))'.format(r=r), "calc", bold=True)
        ws.merge_cells("H%d:I%d" % (r, r))
    ws.conditional_formatting.add("G43:G48", CellIsRule(operator="greaterThanOrEqual", formula=["1"], fill=GREEN))
    ws.conditional_formatting.add("G43:G48", CellIsRule(operator="between", formula=["0.7", "0.9999"], fill=YELLOW))
    ws.conditional_formatting.add("G43:G48", CellIsRule(operator="lessThan", formula=["0.7"], fill=RED))

    ws.merge_cells("B50:I50")
    ws["B50"] = ("Sheets: DASHBOARD → FOOD_DB → DAILY_INPUT → SUPPLEMENTS → ACTIVITY_TDEE → MICROS → "
                 "BODY_LOG → HABITS → SETTINGS → SCENARIOS")
    ws["B50"].font, ws["B50"].alignment = st.FONT_NOTE, st.CENTER
    ws.freeze_panes = "B4"
    ws.print_area = "A1:J50"
    st.landscape(ws, 1)
