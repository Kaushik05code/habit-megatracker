"""ACTIVITY_TDEE: profile + BMR, activity factors, and the daily energy log (tblActivity)."""
from . import styles as st

SHEET = "ACTIVITY_TDEE"
COLUMNS = ["Date", "Planned_Session", "Activity_Level", "Steps", "Extra_Cardio_min", "Extra_Cardio_kcal",
           "Activity_Factor", "Base_TDEE", "TDEE_Day", "Calories_IN", "Surplus_Day", "Surplus_7d_Avg", "Surplus_30d_Avg"]


def build(wb, ctx):
    ws, R = wb[SHEET], ctx.refs
    ws.sheet_properties.tabColor = "ED7D31"
    st.widths(ws, {"A": 3, "B": 24, "C": 17, "D": 16, "E": 10, "F": 17, "G": 17, "H": 15, "I": 12,
                   "J": 12, "K": 13, "L": 13, "M": 15, "N": 16})
    st.title(ws, "🔥 ACTIVITY & TDEE – Energy Balance Engine", "B2:N2")
    st.section(ws, 4, 2, "PROFILE & BMR (linked to SETTINGS)", "ED7D31")
    for i, (label, key, unit) in enumerate([("Age", "age", "years"), ("Sex", "sex", ""), ("Height", "height", "cm"),
                                            ("Current weight", "weight", "kg")]):
        st.put(ws, 5 + i, 2, label, "label")
        st.put(ws, 5 + i, 3, "=" + R[key], "calc", "0.0" if key == "weight" else None)
        st.put(ws, 5 + i, 4, unit, "note")
    st.put(ws, 10, 2, "BMR (Mifflin-St Jeor)", "label")
    st.put(ws, 10, 3, '=IFERROR(10*C8+6.25*C7-5*C5+IF(LEFT(C6,1)="F",-161,5),"")', "total", "0")
    st.put(ws, 10, 4, "kcal/day", "note")
    st.section(ws, 12, 2, "ACTIVITY LEVELS (edit on SETTINGS)", "ED7D31")
    st.headers(ws, 13, 2, ["Activity_Level", "Factor"])
    n_lv = len(ctx.cfg["activity_levels"])
    first_set = int(R["act_table"].split("$")[2].rstrip(":"))
    for i in range(n_lv):
        st.put(ws, 14 + i, 2, "=SETTINGS!$B$%d" % (first_set + i), "calc", align=st.LEFT)
        st.put(ws, 14 + i, 3, "=SETTINGS!$C$%d" % (first_set + i), "calc", "0.000")
    act = "$B$14:$C$%d" % (13 + n_lv)
    r = 15 + n_lv
    st.put(ws, r, 2, "Default activity level", "label")
    st.put(ws, r, 3, "=" + R["act_level"], "calc")
    st.put(ws, r + 1, 2, "Default TDEE", "label")
    st.put(ws, r + 1, 3, '=IFERROR($C$10*VLOOKUP(C%d,%s,2,FALSE),"")' % (r, act), "total", "0")
    R["bmr_live"] = "ACTIVITY_TDEE!$C$10"

    st.section(ws, 4, 6, "ENERGY BALANCE SUMMARY (latest logged date)", "ED7D31")
    summary = [("Latest date", "=IFERROR(MAX(tblActivity[Date]),\"\")", "YYYY-MM-DD"),
               ("7-day avg surplus", '=IFERROR(AVERAGEIFS(tblActivity[Surplus_Day],tblActivity[Date],">"&(G5-7),tblActivity[Date],"<="&G5),"Need more data")', "+0;-0"),
               ("30-day avg surplus", '=IFERROR(AVERAGEIFS(tblActivity[Surplus_Day],tblActivity[Date],">"&(G5-30),tblActivity[Date],"<="&G5),"Need more data")', "+0;-0"),
               ("Projected kg/month (7d)", '=IFERROR(G6*30/%s,"N/A")' % R["kcal_per_kg"], "+0.00;-0.00")]
    for i, (label, f, fmt) in enumerate(summary):
        st.put(ws, 5 + i, 6, label, "label")
        st.put(ws, 5 + i, 7, f, "total", fmt)

    hr = r + 4
    st.section(ws, hr - 1, 2, "DAILY ACTIVITY & ENERGY BALANCE (tblActivity)", "ED7D31")
    st.headers(ws, hr, 2, COLUMNS)
    rows = ctx.samples["activity"]
    n = max(ctx.cfg["layout"]["day_rows"], len(rows) + 30)
    for i in range(n):
        x = hr + 1 + i
        s = rows[i] if i < len(rows) else [None] * 5
        st.put(ws, x, 2, s[0], "input", "YYYY-MM-DD")
        st.put(ws, x, 3, '=IF(B{x}="","",IFERROR(INDEX({s},WEEKDAY(B{x},2)),""))'.format(x=x, s=R["schedule"]), "calc")
        for j, v in enumerate(s[1:]):
            st.put(ws, x, 4 + j, v, "input")
        f = [
            '=IF(B{x}="","",IFERROR(VLOOKUP(D{x},%s,2,FALSE),%s))' % (act, R["act_factor"]),
            '=IF(H{x}="","",IFERROR($C$10*H{x},""))',
            '=IF(I{x}="","",IFERROR(I{x}+N(G{x}),""))',
            '=IF(B{x}="","",IFERROR(SUMIFS(tblDailyTotals[Total_Calories],tblDailyTotals[Date],B{x}),0))',
            '=IF(OR(J{x}="",N(K{x})=0),"",K{x}-J{x})',  # no food logged -> blank, so averages stay honest
            '=IF(L{x}="","",IFERROR(AVERAGEIFS(tblActivity[Surplus_Day],tblActivity[Date],">"&(B{x}-7),tblActivity[Date],"<="&B{x}),""))',
            '=IF(L{x}="","",IFERROR(AVERAGEIFS(tblActivity[Surplus_Day],tblActivity[Date],">"&(B{x}-30),tblActivity[Date],"<="&B{x}),""))',
        ]
        for j, formula in enumerate(f):
            st.put(ws, x, 8 + j, formula.format(x=x), "total" if j in (2, 4) else "calc",
                   "0.000" if j == 0 else ("+0;-0" if j >= 4 else "0"))
    st.table(ws, "tblActivity", "B%d:N%d" % (hr, hr + n))
    st.dropdown(ws, "=" + R["act_names"], "D%d:D%d" % (hr + 1, hr + n))
    ws.freeze_panes = "C3"
    st.landscape(ws)
