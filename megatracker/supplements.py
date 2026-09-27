"""SUPPLEMENTS: master list (tblSuppMaster) and daily log (tblSuppLog) with micronutrient totals."""
from . import styles as st

SHEET = "SUPPLEMENTS"
MASTER = ["Supp_ID", "Name", "Brand", "Default_Dose", "Dose_Units", "Cal_per_Dose", "VitD_IU", "Calcium_mg",
          "Magnesium_mg", "Iron_mg", "VitB12_mcg", "Omega3_mg", "Creatine_g"]
LOG = ["Date", "Supp_ID", "Actual_Doses", "Supp_Calories"] + MASTER[6:]


def build(wb, ctx):
    ws = wb[SHEET]
    ws.sheet_properties.tabColor = "BF8F00"
    st.widths(ws, {"A": 3, "B": 14, "C": 26, "D": 14, "E": 15, "F": 13, "G": 13, "H": 12, "I": 13,
                   "J": 14, "K": 10, "L": 12, "M": 12, "N": 12})
    st.title(ws, "💊 SUPPLEMENTS – Master List & Daily Log", "B2:N2")
    st.section(ws, 3, 2, "SUPPLEMENT MASTER TABLE", "BF8F00")
    st.headers(ws, 4, 2, MASTER)
    supps = ctx.cfg["supplements"]
    n_master = len(supps) + 5
    for i in range(n_master):
        vals = supps[i] if i < len(supps) else [None] * len(MASTER)
        for j, v in enumerate(vals):
            st.put(ws, 5 + i, 2 + j, v, "input", align=st.LEFT if j < 3 else st.CENTER)
    last_m = 4 + n_master
    st.table(ws, "tblSuppMaster", "B4:N%d" % last_m)
    hr = last_m + 3
    st.section(ws, hr - 1, 2, "DAILY SUPPLEMENT LOG", "BF8F00")
    st.headers(ws, hr, 2, LOG)
    rows = ctx.samples["supps"]
    n = max(ctx.cfg["layout"]["supp_log_rows"], len(rows) + 50)
    for i in range(n):
        r = hr + 1 + i
        s = rows[i] if i < len(rows) else [None] * 3
        st.put(ws, r, 2, s[0], "input", "YYYY-MM-DD")
        st.put(ws, r, 3, s[1], "input")
        st.put(ws, r, 4, s[2], "input")
        for j, col in enumerate(["Cal_per_Dose"] + MASTER[6:]):
            st.put(ws, r, 5 + j, '=IF(OR(C%d="",D%d=""),"",IFERROR(INDEX(tblSuppMaster[%s],'
                                 'MATCH(C%d,tblSuppMaster[Supp_ID],0))*D%d,0))' % (r, r, col, r, r), "calc", "0.0")
    st.table(ws, "tblSuppLog", "B%d:L%d" % (hr, hr + n))
    st.dropdown(ws, "=$B$5:$B$%d" % last_m, "C%d:C%d" % (hr + 1, hr + n))
    ws.freeze_panes = "C5"
    st.landscape(ws)
