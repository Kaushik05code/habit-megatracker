"""MICROS: RDA table (tblRDA) and per-date micronutrients from food + supplements (tblMicrosDaily)."""
from openpyxl.utils import get_column_letter

from . import styles as st

SHEET = "MICROS"
# (short label, tblIntake column, tblSuppLog column)
NUTRIENTS = [("VitD", "Row_VitaminD_IU", "VitD_IU"), ("Ca", "Row_Calcium_mg", "Calcium_mg"),
             ("Mg", "Row_Magnesium_mg", "Magnesium_mg"), ("Fe", "Row_Iron_mg", "Iron_mg"),
             ("B12", "Row_VitB12_mcg", "VitB12_mcg"), ("Om3", "Row_Omega3_mg", "Omega3_mg")]


def build(wb, ctx):
    ws = wb[SHEET]
    ws.sheet_properties.tabColor = "7030A0"
    st.widths(ws, {"A": 3, "B": 14, "C": 10, "D": 11})
    for c in range(5, 27):
        ws.column_dimensions[get_column_letter(c)].width = 11
    st.title(ws, "🧬 MICROS – Vitamins & Minerals (food + supplements)", "B2:Z2")
    st.section(ws, 3, 2, "DAILY TARGETS (edit to match your age and sex)", "7030A0")
    st.headers(ws, 4, 2, ["Nutrient", "Unit", "RDA_Value"])
    rda = ctx.cfg["rda"]
    for i, (name, unit, val) in enumerate(rda):
        st.put(ws, 5 + i, 2, name, "label")
        st.put(ws, 5 + i, 3, unit, "calc")
        st.put(ws, 5 + i, 4, val, "input")
    st.table(ws, "tblRDA", "B4:D%d" % (4 + len(rda)))
    ctx.refs["rda"] = ["MICROS!$D$%d" % (5 + i) for i in range(len(rda))]

    hr = 6 + len(rda) + 2
    st.section(ws, hr - 1, 2, "DAILY MICRONUTRIENT SUMMARY (dates follow DAILY_INPUT totals)", "7030A0")
    cols = ["Date"]
    for lab, _, _ in NUTRIENTS:
        cols += [lab + "_Food", lab + "_Supp", lab + "_Total", lab + "_Pct_RDA"]
    st.headers(ws, hr, 2, cols)
    n = ctx.refs["total_rows"]
    for i in range(n):
        r = hr + 1 + i
        st.put(ws, r, 2, '=IF(DAILY_INPUT!$U${d}="","",DAILY_INPUT!$U${d})'.format(d=4 + i), "calc", "YYYY-MM-DD")
        for k, (lab, food_col, supp_col) in enumerate(NUTRIENTS):
            c = 3 + 4 * k
            L = [get_column_letter(c + j) for j in range(4)]
            st.put(ws, r, c, '=IF($B{r}="","",IFERROR(SUMIFS(tblIntake[{f}],tblIntake[Date],$B{r}),0))'.format(r=r, f=food_col), "calc", "0.0")
            st.put(ws, r, c + 1, '=IF($B{r}="","",IFERROR(SUMIFS(tblSuppLog[{s}],tblSuppLog[Date],$B{r}),0))'.format(r=r, s=supp_col), "calc", "0.0")
            st.put(ws, r, c + 2, '=IF($B{r}="","",N({a}{r})+N({b}{r}))'.format(r=r, a=L[0], b=L[1]), "total", "0.0")
            st.put(ws, r, c + 3, '=IF($B{r}="","",IFERROR({t}{r}/{rda},0))'.format(r=r, t=L[2], rda=ctx.refs["rda"][k]), "calc", "0%")
    st.table(ws, "tblMicrosDaily", "B%d:Z%d" % (hr, hr + n))
    ws.freeze_panes = "C%d" % (hr + 1)
    st.landscape(ws)
