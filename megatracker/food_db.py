"""FOOD_DB: master food table (tblFood). Every macro and micro lookup reads from here."""
from . import styles as st

SHEET = "FOOD_DB"
COLUMNS = ["Food_ID", "Food_Name", "Category", "Unit_Type", "Unit_Grams", "Calories", "Protein_g", "Carbs_g",
           "Fat_g", "Fibre_g", "Calcium_mg", "VitaminD_IU", "Magnesium_mg", "Iron_mg", "VitB12_mcg", "Omega3_mg"]
SPARE_ROWS = 20


def build(wb, ctx):
    ws = wb[SHEET]
    ws.sheet_properties.tabColor = "2E75B6"
    st.widths(ws, {"A": 3, "B": 16, "C": 34})
    for i in range(3, 17):
        ws.column_dimensions[chr(ord("B") + i)].width = 12
    st.title(ws, "🍽 FOOD DATABASE – Master Nutrition Reference", "B2:Q2")
    st.headers(ws, 3, 2, COLUMNS)
    foods = ctx.cfg["foods"]
    for i in range(len(foods) + SPARE_ROWS):
        vals = foods[i] if i < len(foods) else [None] * len(COLUMNS)
        for j, v in enumerate(vals):
            st.put(ws, 4 + i, 2 + j, v, "input", align=st.LEFT if j < 4 else st.CENTER)
    last = 3 + len(foods) + SPARE_ROWS
    st.table(ws, "tblFood", "B3:Q%d" % last)
    st.dropdown(ws, '"per_100g,per_100ml,per_unit"', "E4:E%d" % last)
    ctx.refs["food_ids"] = "FOOD_DB!$B$4:$B$%d" % last
    note = ws.cell(row=last + 2, column=2, value=(
        "Scaling: per_100g / per_100ml -> factor = Amount / 100; per_unit -> factor = Amount / Unit_Grams. "
        "Row value = nutrient x factor. Add foods in the blank blue rows."))
    note.font = st.FONT_NOTE
    ws.freeze_panes = "C4"
    st.landscape(ws)
