"""SCENARIOS: what-if sandbox using the same Food_IDs and scaling logic as DAILY_INPUT."""
from . import styles as st

SHEET = "SCENARIOS"


def build(wb, ctx):
    ws, R, sc = wb[SHEET], ctx.refs, ctx.cfg["scenario"]
    ws.sheet_properties.tabColor = "00B050"
    st.widths(ws, {"A": 3, "B": 26, "C": 17, "D": 18, "E": 17, "F": 13, "G": 12, "H": 12, "I": 12, "J": 12})
    st.title(ws, "🧪 SCENARIOS – What-If Sandbox", "B2:J2")
    st.section(ws, 4, 2, "SCENARIO INPUTS (blue) AND RESULTS", "00B050")
    st.headers(ws, 5, 2, ["Food", "Food_ID", "Grams_per_Serving", "Servings_per_Day", "Total_g_ml",
                          "Calories", "Protein_g", "Carbs_g", "Fat_g"])
    foods = sc["foods"]
    n = len(foods) + 4
    lk = "INDEX(tblFood[%s],MATCH(C{r},tblFood[Food_ID],0))"
    for i in range(n):
        r = 6 + i
        vals = foods[i] if i < len(foods) else [None] * 4
        for j, v in enumerate(vals):
            st.put(ws, r, 2 + j, v, "input", align=st.LEFT if j == 0 else st.CENTER)
        st.put(ws, r, 6, '=IF(OR(C{r}="",D{r}=""),"",D{r}*IF(E{r}="",1,E{r}))'.format(r=r), "calc", "0")
        factor = 'IF({ut}="per_unit",F{r}/{ug},F{r}/100)'.format(r=r, ut=(lk % "Unit_Type").format(r=r),
                                                                  ug=(lk % "Unit_Grams").format(r=r))
        for j, m in enumerate(["Calories", "Protein_g", "Carbs_g", "Fat_g"]):
            st.put(ws, r, 7 + j, '=IF(F{r}="","",IFERROR({v}*{f},0))'.format(r=r, v=(lk % m).format(r=r), f=factor),
                   "calc", "0" if j == 0 else "0.0")
    tot = 6 + n
    st.put(ws, tot, 2, "TOTAL", "label")
    for j in range(4):
        col = "GHIJ"[j]
        st.put(ws, tot, 7 + j, "=SUM(%s6:%s%d)" % (col, col, tot - 1), "total", "0" if j == 0 else "0.0")
    st.dropdown(ws, "=" + R["food_ids"], "C6:C%d" % (tot - 1))

    r0 = tot + 2
    st.section(ws, r0, 2, "DAY SETTINGS & ENERGY BALANCE", "00B050")
    st.put(ws, r0 + 1, 2, "Activity level", "label")
    st.put(ws, r0 + 1, 3, sc["activity_level"], "input")
    st.dropdown(ws, "=" + R["act_names"], "C%d" % (r0 + 1))
    st.put(ws, r0 + 2, 2, "Sport day? (Yes/No)", "label")
    st.put(ws, r0 + 2, 3, sc["sport_day"], "input")
    st.dropdown(ws, '"Yes,No"', "C%d" % (r0 + 2))
    b = r0 + 3
    items = [("Scenario calories IN", "=G%d" % tot, "0"),
             ("BMR", "=" + R["bmr"], "0"),
             ("Activity factor", '=IFERROR(VLOOKUP(C%d,%s,2,FALSE),%s)' % (r0 + 1, R["act_table"], R["act_factor"]), "0.000"),
             ("Scenario TDEE", '=IFERROR(C%d*C%d,"")' % (b + 1, b + 2), "0"),
             ("Sport-day bonus", '=IF(C%d="Yes",%s,0)' % (r0 + 2, R["sport_bonus"]), "0"),
             ("Adjusted TDEE", '=IFERROR(C%d+C%d,"")' % (b + 3, b + 4), "0"),
             ("Surplus / deficit", '=IFERROR(C%d-C%d,"")' % (b, b + 5), "+0;-0"),
             ("Projected change (kg/month)", '=IFERROR(C%d*30/%s,0)' % (b + 6, R["kcal_per_kg"]), "+0.00;-0.00")]
    for i, (label, f, fmt) in enumerate(items):
        st.put(ws, b + i, 2, label, "label")
        st.put(ws, b + i, 3, f, "total" if i in (0, 6, 7) else "calc", fmt)
    s = b + len(items) + 1
    ws.merge_cells("B%d:J%d" % (s, s + 1))
    c = ws.cell(row=s, column=2, value=(
        '=IFERROR("If you follow this scenario every day, you eat "&TEXT(C{b},"#,##0")&" kcal with "&TEXT(H{t},"0")'
        '&" g protein, "&TEXT(I{t},"0")&" g carbs and "&TEXT(J{t},"0")&" g fat: a surplus of "&TEXT(C{sp},"+#,##0;-#,##0")'
        '&" kcal/day vs TDEE, or about "&TEXT(C{kg},"+0.00;-0.00")&" kg/month.","")').format(b=b, t=tot, sp=b + 6, kg=b + 7))
    c.font, c.alignment = st.font(11, italic=True, color="1F4E79"), st.WRAP
    ws.row_dimensions[s].height = 30
    ws.freeze_panes = "B3"
    st.landscape(ws, 1)
