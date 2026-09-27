"""SETTINGS: profile, goals, targets, activity factors, BMR/TDEE, schedule, habit targets."""
from . import styles as st

SHEET = "SETTINGS"


def build(wb, ctx):
    ws, cfg, R = wb[SHEET], ctx.cfg, ctx.refs
    ws.sheet_properties.tabColor = "1F4E79"
    st.widths(ws, {"A": 3, "B": 34, "C": 24, "D": 14, "E": 44})
    st.title(ws, "⚙ SETTINGS – Constants & Preferences", "B2:E2")
    row = [4]

    def block(name, cols=("Value", "Unit", "Notes")):
        st.headers(ws, row[0], 2, [name] + list(cols))
        row[0] += 1

    def line(key, label, value, kind="input", unit="", note="", fmt=None):
        r = row[0]
        st.put(ws, r, 2, label, "label")
        st.put(ws, r, 3, value, kind, fmt)
        st.put(ws, r, 4, unit, "note")
        st.put(ws, r, 5, note, "note")
        if key:
            R[key] = "%s!$C$%d" % (SHEET, r)
        row[0] += 1
        return r

    p, g, t, h = cfg["profile"], cfg["goals"], cfg["training"], cfg["habits"]
    block("PROFILE")
    line("name", "Name", p["name"])
    r = line("sex", "Sex", p["sex"], note="Male or Female")
    st.dropdown(ws, '"Male,Female"', "C%d" % r)
    line("age", "Age", p["age"], unit="years")
    line("height", "Height", p["height_cm"], unit="cm")
    line("start_weight", "Start weight", p["start_weight_kg"], unit="kg", note="Used until BODY_LOG has a weigh-in")
    line("weight", "Current weight", "=IFERROR(LOOKUP(9.99E+307,tblBody[Weight_kg]),%s)" % R["start_weight"],
         "total", "kg", "Latest BODY_LOG weigh-in", "0.0")
    line("body_fat", "Body fat (est.)", p["body_fat_pct"], unit="%")
    row[0] += 1
    block("GOAL")
    line("goal", "Goal", g["goal_text"])
    line("target_weight", "Target weight", g["target_weight_kg"], unit="kg")
    line("timeline", "Timeline", g["timeline_months"], unit="months")
    line("wk_min", "Weekly change: on-track min", g["weekly_change_min_kg"], unit="kg/week")
    line("wk_max", "Weekly change: on-track max", g["weekly_change_max_kg"], unit="kg/week")
    row[0] += 1
    block("TARGET MACROS & SURPLUS")
    line("protein_per_kg", "Protein per kg bodyweight", g["protein_g_per_kg"], unit="g/kg", note="1.6-2.2 g/kg for hypertrophy")
    line("fat_pct", "Fat share of calories", g["fat_pct_of_calories"], unit="%", note="20-35% is typical")
    line("surplus", "Calorie surplus target", g["calorie_surplus_kcal"], unit="kcal/day", note="Negative for a cut")
    line("kcal_per_kg", "kcal per kg body weight", g["kcal_per_kg"], unit="kcal/kg", note="Energy density constant")
    line(None, "Projected change at target", '=IFERROR(%s*30/%s,"")' % (R["surplus"], R["kcal_per_kg"]),
         "calc", "kg/month", "Surplus x 30 / kcal per kg", "+0.00;-0.00")
    row[0] += 1
    st.headers(ws, row[0], 2, ["ACTIVITY LEVEL", "Factor", "", "Description"])
    first = row[0] + 1
    for lvl, factor, desc in cfg["activity_levels"]:
        r = row[0] = row[0] + 1
        st.put(ws, r, 2, lvl, "label")
        st.put(ws, r, 3, factor, "input", "0.000")
        st.put(ws, r, 5, desc, "note")
    R["act_table"] = "%s!$B$%d:$C$%d" % (SHEET, first, row[0])
    R["act_names"] = "%s!$B$%d:$B$%d" % (SHEET, first, row[0])
    row[0] += 1
    r = line("act_level", "Default activity level", t["default_activity_level"])
    st.dropdown(ws, "=" + R["act_names"], "C%d" % r)
    line("act_factor", "Default activity factor", '=IFERROR(VLOOKUP(%s,%s,2,FALSE),1.55)' % (R["act_level"], R["act_table"]),
         "total", "", "Looked up from the table above", "0.000")
    row[0] += 1
    block("ENERGY & DAILY TARGETS")
    line("bmr", "BMR (Mifflin-St Jeor)",
         '=IFERROR(10*%s+6.25*%s-5*%s+IF(LEFT(%s,1)="F",-161,5),"")' % (R["weight"], R["height"], R["age"], R["sex"]),
         "total", "kcal/day", "10 x kg + 6.25 x cm - 5 x age + 5 (male) / -161 (female)", "0")
    line("tdee", "Default TDEE", '=IFERROR(%s*%s,"")' % (R["bmr"], R["act_factor"]), "total", "kcal/day", "BMR x activity factor", "0")
    line("cal_target", "Daily calorie target", '=IFERROR(%s+%s,"")' % (R["tdee"], R["surplus"]), "total", "kcal", "TDEE + surplus", "0")
    line("protein_target", "Daily protein target", '=IFERROR(%s*%s,"")' % (R["weight"], R["protein_per_kg"]), "total", "g", "Weight x g/kg", "0")
    line("fat_target", "Daily fat target", '=IFERROR(%s*%s/100/9,"")' % (R["cal_target"], R["fat_pct"]), "total", "g", "", "0")
    line("carb_target", "Daily carb target", '=IFERROR((%s-%s*4-%s*9)/4,"")' % (R["cal_target"], R["protein_target"], R["fat_target"]),
         "total", "g", "Remaining calories", "0")
    row[0] += 1
    st.headers(ws, row[0], 2, ["TRAINING SCHEDULE", "Planned session", "", "Notes"])
    first = row[0] + 1
    for day in ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]:
        r = row[0] = row[0] + 1
        st.put(ws, r, 2, day, "label")
        st.put(ws, r, 3, t["schedule"].get(day, "Rest"), "input")
    R["schedule"] = "%s!$C$%d:$C$%d" % (SHEET, first, row[0])
    row[0] += 1
    line("sport", "Sport", t["sport_name"])
    line("sport_bonus", "Sport-day bonus", t["sport_day_bonus_kcal"], unit="kcal", note="Used by SCENARIOS")
    row[0] += 1
    block("HABIT TARGETS")
    line("habit_name", "Custom habit (HABITS column)", h["custom_habit"], "calc", note="Rename in config.yaml and rebuild")
    line("screen_max", "Screen time limit", h["screen_time_max_hrs"], unit="hrs/day")
    line("sleep_min", "Sleep target", h["sleep_min_hrs"], unit="hrs/night")
    line("water_min", "Water target", h["water_min_l"], unit="L/day")
    line("deep_min", "Deep work / study target", h["deep_work_min_hrs"], unit="hrs/day")
    row[0] += 1
    st.headers(ws, row[0], 2, ["COLOUR LEGEND", "Meaning"])
    for fill, text in ((st.FILL_INPUT, "Input: type here"), (st.FILL_CALC, "Formula: do not overwrite"),
                       (st.FILL_TOTAL, "Key result / target"), (st.FILL_HEADER, "Header")):
        r = row[0] = row[0] + 1
        c = st.put(ws, r, 2, "", "calc")
        c.fill = fill
        st.put(ws, r, 3, text, "note")
    ws.freeze_panes = "B3"
    st.landscape(ws)
