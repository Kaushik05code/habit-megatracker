#!/usr/bin/env python3
"""Build the MegaTracker workbook.

    python build.py [--config config.yaml] [--out output/MegaTracker.xlsx] [--sample-days N]
"""
import argparse

from megatracker.config import load_config
from megatracker.workbook import build_workbook


def main(argv=None):
    ap = argparse.ArgumentParser(description="Build the MegaTracker Excel workbook from a YAML config.")
    ap.add_argument("--config", help="YAML config (default: config.yaml if present, else config.example.yaml)")
    ap.add_argument("--out", default="output/MegaTracker.xlsx", help="output .xlsx path")
    ap.add_argument("--sample-days", type=int, help="days of demo data (0 = empty template)")
    args = ap.parse_args(argv)
    cfg, source = load_config(args.config)
    if args.sample_days is not None:
        cfg["sample"]["days"] = args.sample_days
    out = build_workbook(cfg, args.out)
    print("Built %s (config: %s)" % (out, source))


if __name__ == "__main__":
    main()
