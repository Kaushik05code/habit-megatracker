"""Deterministic demo rows built from the config (template day, schedule, supplements)."""
import datetime as dt
import random

DAYS = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]


def make_samples(cfg):
    s = cfg["sample"]
    n = int(s.get("days", 0))
    end = s.get("end_date") or dt.date.today()
    rnd = random.Random(s.get("seed", 7))
    dates = [end - dt.timedelta(days=n - 1 - i) for i in range(n)]
    out = {"dates": dates, "intake": [], "activity": [], "supps": [], "body": [], "habits": []}
    tr = cfg["training"]
    w0 = float(cfg["profile"]["start_weight_kg"]) - 0.2
    for i, d in enumerate(dates):
        for item in cfg["template_day"]:
            if item.get("optional") and rnd.random() < 0.5:
                continue
            food = item["rotate"][i % len(item["rotate"])] if item.get("rotate") else item["food"]
            amt = int(round(item["amount"] * rnd.uniform(0.9, 1.1) / 5.0) * 5)
            out["intake"].append([d, item["meal"], item["time"], food, item.get("note", ""), amt])
        session = tr["schedule"].get(DAYS[d.weekday()], "Rest")
        level = tr["session_activity_level"].get(session, tr["default_activity_level"])
        steps = {"Rest": 6000}.get(session, 9000) + rnd.randrange(0, 3000, 100)
        cardio = (30, 150) if session == "Rest" else (0, 0)
        out["activity"].append([d, level, steps, cardio[0], cardio[1]])
        for sp in cfg["supplements"]:
            if not (sp[0] == "FISH_OIL" and rnd.random() < 0.2):
                out["supps"].append([d, sp[0], 1])
        sleep = round(rnd.uniform(6.6, 8.6), 1)
        out["body"].append([d, round(w0 + 0.03 * i + rnd.uniform(-0.15, 0.15), 1), "Yes",
                            sleep, rnd.randint(3, 5), rnd.randint(1, 3)])
        habit = 1 if (i >= n - 6 or rnd.random() < 0.7) else 0
        out["habits"].append([d, round(rnd.uniform(1.8, 4.6), 1), round(rnd.uniform(2.2, 3.6), 1),
                              round(rnd.uniform(1.0, 4.0), 1), habit])
    return out
