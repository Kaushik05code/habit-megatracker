"""Orchestrates the build: creates sheets in display order, then fills them in dependency order."""
from pathlib import Path

from openpyxl import Workbook

from . import (activity_tdee, body_log, daily_input, dashboard, food_db, habits, micros, scenarios,
               settings, supplements)
from .sample_data import make_samples

SHEET_ORDER = ["DASHBOARD", "FOOD_DB", "DAILY_INPUT", "SUPPLEMENTS", "ACTIVITY_TDEE", "MICROS",
               "BODY_LOG", "HABITS", "SETTINGS", "SCENARIOS"]
# SETTINGS first (it defines the cell references others use), DASHBOARD last (it reads everything).
BUILD_ORDER = [settings, food_db, body_log, supplements, daily_input, activity_tdee, micros, habits,
               scenarios, dashboard]


class Context:
    def __init__(self, cfg):
        self.cfg = cfg
        self.samples = make_samples(cfg)
        self.refs = {}  # key -> absolute cell address, e.g. "bmr" -> "SETTINGS!$C$37"


def build_workbook(cfg, out_path):
    wb = Workbook()
    wb.remove(wb.active)
    for name in SHEET_ORDER:
        wb.create_sheet(name)
    ctx = Context(cfg)
    for module in BUILD_ORDER:
        module.build(wb, ctx)
    wb.active = 0
    out = Path(out_path)
    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out)
    return out
