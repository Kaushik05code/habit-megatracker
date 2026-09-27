"""DAILY_INPUT: food log (tblIntake) scaled from FOOD_DB, plus per-date totals (tblDailyTotals)."""
from openpyxl.utils import get_column_letter

from . import styles as st

SHEET = "DAILY_INPUT"
MACROS = ["Calories", "Protein_g", "Carbs_g", "Fat_g", "Fibre_g"]
MICROS = ["Calcium_mg", "VitaminD_IU", "Magnesium_mg", "Iron_mg", "VitB12_mcg", "Omega3_mg"]
INTAKE = (["Date", "Meal_Label", "Time", "Food_ID", "Description", "Amount_g_ml", "Factor"]
          + ["Row_" + m for m in MACROS] + ["Row_" + m for m in MICROS])
TOTALS = ["Date", "Total_Calories", "Total_Protein_g", "Total_Carbs_g", "Total_Fat_g", "Total_Fibre_g"]
TOT_COL = 21  # column U


def build(wb, ctx):
    ws = wb[SHEET]
    ws.sheet_properties.tabColor = "C00000"
    st.widths(ws, {"A": 3, "B": 13, "C": 14, "D": 8, "E": 17, "F": 22, "G": 13, "H": 9, "T": 3})
    for c in range(9, 20):
        ws.column_dimensions[get_column_letter(c)].width = 13
    for c in range(TOT_COL, TOT_COL + 6):
        ws.column_dimensions[get_column_letter(c)].width = 15
    st.title(ws, "🍽 DAILY INPUT – Food & Drink Log  (daily totals in columns U–Z)", "B2:Z2")
    st.headers(ws, 3, 2, INTAKE)
    rows = ctx.samples["intake"]
    n = max(ctx.cfg["layout"]["intake_rows"], len(rows) + 100)
    look = 'INDEX(tblFood[%s],MATCH(E{r},tblFood[Food_ID],0))'
    for i in range(n):
        r = 4 + i
        s = rows[i] if i < len(rows) else [None] * 6
        for j, v in enumerate(s):
            st.put(ws, r, 2 + j, v, "input", "YYYY-MM-DD" if j == 0 else None, align=st.LEFT if j in (1, 4) else st.CENTER)
        ut = (look % "Unit_Type").format(r=r)
        # Factor: per_100g/per_100ml -> Amount/100, per_unit -> Amount/Unit_Grams.
        st.put(ws, r, 8, '=IF(OR(E{r}="",G{r}=""),"",IFERROR(IF(OR({ut}="per_100g",{ut}="per_100ml"),G{r}/100,'
                         'IF({ut}="per_unit",G{r}/{ug},0)),0))'.format(r=r, ut=ut, ug=(look % "Unit_Grams").format(r=r)),
               "calc", "0.00")
        for j, m in enumerate(MACROS + MICROS):
            st.put(ws, r, 9 + j, '=IF(H{r}="","",IFERROR({v}*H{r},0))'.format(r=r, v=(look % m).format(r=r)),
                   "calc", "0" if j == 0 else "0.0")
    last = 3 + n
    st.table(ws, "tblIntake", "B3:S%d" % last)
    st.dropdown(ws, "=" + ctx.refs["food_ids"], "E4:E%d" % last)
    ws.column_dimensions.group("N", "S", hidden=True)  # per-row micronutrients, used by MICROS

    st.headers(ws, 3, TOT_COL, TOTALS)
    days = ctx.samples["dates"]
    nd = max(ctx.cfg["layout"]["day_rows"], len(days) + 30)
    for i in range(nd):
        r = 4 + i
        st.put(ws, r, TOT_COL, days[i] if i < len(days) else None, "input", "YYYY-MM-DD")
        for j, m in enumerate(MACROS):
            st.put(ws, r, TOT_COL + 1 + j, '=IF($U{r}="","",IFERROR(SUMIFS(tblIntake[Row_{m}],tblIntake[Date],$U{r}),0))'
                   .format(r=r, m=m), "total", "0" if j == 0 else "0.0")
    st.table(ws, "tblDailyTotals", "U3:Z%d" % (3 + nd))
    ctx.refs["total_dates"] = "DAILY_INPUT!$U$4:$U$%d" % (3 + nd)
    ctx.refs["total_rows"] = nd
    ws.freeze_panes = "C4"
    st.landscape(ws)
