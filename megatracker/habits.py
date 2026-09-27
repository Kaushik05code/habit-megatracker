"""HABITS (new): daily habit log (tblHabits) with 7-day averages and streaks."""
import re

from . import styles as st

SHEET = "HABITS"


def habit_column(cfg):
    return re.sub(r"\W+", "_", str(cfg["habits"]["custom_habit"])).strip("_") or "Custom_Habit"


def build(wb, ctx):
    ws, R = wb[SHEET], ctx.refs
    ws.sheet_properties.tabColor = "008080"
    hab = habit_column(ctx.cfg)
    cols = ["Date", "Screen_Time_hrs", "Sleep_hrs", "Water_L", "Steps", "Deep_Work_hrs", hab,
            "Screen_7d_Avg", "Sleep_7d_Avg", "Water_7d_Avg", "Deep_Work_7d_Avg", "Habit_Streak", "Sleep_Streak"]
    st.widths(ws, {"A": 3, "B": 13, "C": 16, "D": 11, "E": 10, "F": 10, "G": 15, "H": max(14, len(hab) + 3),
                   "I": 15, "J": 14, "K": 14, "L": 18, "M": 14, "N": 14})
    st.title(ws, "🧠 HABITS – Daily Habit Log", "B2:N2")
    n3 = ws.cell(row=3, column=2, value="Blue = type here. Sleep comes from BODY_LOG and steps from ACTIVITY_TDEE. "
                                         "%s: 1 = done, 0 = missed. Streaks count consecutive dates." % hab)
    n3.font = st.FONT_NOTE
    st.headers(ws, 4, 2, cols)
    rows = ctx.samples["habits"]
    n = max(ctx.cfg["layout"]["day_rows"], len(rows) + 30)
    avg = '=IF($B{r}="","",IFERROR(AVERAGEIFS(tblHabits[%s],tblHabits[Date],">"&($B{r}-7),tblHabits[Date],"<="&$B{r}),""))'
    for i in range(n):
        r = 5 + i
        s = rows[i] if i < len(rows) else [None] * 5
        st.put(ws, r, 2, s[0], "input", "YYYY-MM-DD")
        st.put(ws, r, 3, s[1], "input", "0.0")
        st.put(ws, r, 4, '=IF(B{r}="","",IFERROR(AVERAGEIFS(tblBody[Sleep_Hours],tblBody[Date],B{r}),""))'.format(r=r), "calc", "0.0")
        st.put(ws, r, 5, s[2], "input", "0.0")
        st.put(ws, r, 6, '=IF(B{r}="","",IFERROR(AVERAGEIFS(tblActivity[Steps],tblActivity[Date],B{r}),""))'.format(r=r), "calc", "#,##0")
        st.put(ws, r, 7, s[3], "input", "0.0")
        st.put(ws, r, 8, s[4], "input", "0")
        for j, col in enumerate(["Screen_Time_hrs", "Sleep_hrs", "Water_L", "Deep_Work_hrs"]):
            st.put(ws, r, 9 + j, (avg % col).format(r=r), "calc", "0.0")
        prev = "IF(IFERROR(B{r}-B{p},0)=1,N({c}{p})+1,1)" if i else "1"
        # Streak: +1 when the habit is done and the previous row is the day before; resets to 0 when missed.
        st.put(ws, r, 13, ('=IF(B{r}="","",IF(H{r}=1,%s,0))' % prev).format(r=r, p=r - 1, c="M"), "total", "0")
        st.put(ws, r, 14, ('=IF(B{r}="","",IF(AND(ISNUMBER(D{r}),N(D{r})>=%s),%s,0))' % (R["sleep_min"], prev))
               .format(r=r, p=r - 1, c="N"), "total", "0")
    st.table(ws, "tblHabits", "B4:N%d" % (4 + n))
    st.dropdown(ws, '"1,0"', "H5:H%d" % (4 + n))
    R["habit_col"] = hab
    ws.freeze_panes = "C5"
    st.landscape(ws)
