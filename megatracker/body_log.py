"""BODY_LOG: weigh-ins, measurements, sleep, stress (tblBody) with 7-day average and weekly change."""
from . import styles as st

SHEET = "BODY_LOG"
COLUMNS = ["Date", "Weight_kg", "Morning_Fast", "Waist_cm", "Chest_cm", "Thigh_cm", "Sleep_Hours",
           "Sleep_Quality_1to5", "Stress_1to5", "Notes", "Weight_7d_Avg", "Weekly_Change_kg"]


def build(wb, ctx):
    ws = wb[SHEET]
    ws.sheet_properties.tabColor = "548235"
    st.widths(ws, {"A": 3, "B": 13, "C": 11, "D": 14, "E": 10, "F": 10, "G": 10, "H": 12, "I": 18,
                   "J": 12, "K": 24, "L": 15, "M": 18})
    st.title(ws, "📊 BODY LOG – Weight, Measurements & Recovery", "B2:M2")
    st.headers(ws, 3, 2, COLUMNS)
    rows = ctx.samples["body"]
    n = max(ctx.cfg["layout"]["day_rows"], len(rows) + 30)
    for i in range(n):
        r = 4 + i
        s = rows[i] if i < len(rows) else [None] * 6
        vals = [s[0], s[1], s[2], None, None, None, s[3], s[4], s[5], None]
        for j, v in enumerate(vals):
            st.put(ws, r, 2 + j, v, "input", "YYYY-MM-DD" if j == 0 else ("0.0" if j in (1, 6) else None))
        # 7-day average of weigh-ins dated (Date-7, Date]; weekly change compares it with the 7 days before.
        st.put(ws, r, 12, '=IF(OR(B%d="",C%d=""),"",IFERROR(AVERAGEIFS(tblBody[Weight_kg],tblBody[Date],">"&(B%d-7),'
                           'tblBody[Date],"<="&B%d),""))' % (r, r, r, r), "calc", "0.00")
        st.put(ws, r, 13, '=IF(L%d="","",IFERROR(L%d-AVERAGEIFS(tblBody[Weight_kg],tblBody[Date],">"&(B%d-14),'
                           'tblBody[Date],"<="&(B%d-7)),""))' % (r, r, r, r), "calc", "+0.00;-0.00")
    st.table(ws, "tblBody", "B3:M%d" % (3 + n))
    st.dropdown(ws, '"Yes,No"', "D4:D%d" % (3 + n))
    ws.freeze_panes = "C4"
    st.landscape(ws)
