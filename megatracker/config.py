"""Load config.yaml (or a given file) on top of config.example.yaml."""
import copy
from pathlib import Path

import yaml

ROOT = Path(__file__).resolve().parent.parent
EXAMPLE = ROOT / "config.example.yaml"


def _merge(base, override):
    out = copy.deepcopy(base)
    for key, val in (override or {}).items():
        if isinstance(val, dict) and isinstance(out.get(key), dict):
            out[key] = _merge(out[key], val)
        else:
            out[key] = val
    return out


def _read(path):
    with open(path, encoding="utf-8") as fh:
        return yaml.safe_load(fh) or {}


def load_config(path=None):
    """Return (config, source). Missing keys fall back to config.example.yaml."""
    example = _read(EXAMPLE)
    if path is None:
        for cand in (Path("config.yaml"), ROOT / "config.yaml"):
            if cand.exists():
                path = cand
                break
    if path is None:
        cfg, source = example, str(EXAMPLE)
    else:
        cfg, source = _merge(example, _read(path)), str(path)
    food_ids = {f[0] for f in cfg["foods"]}
    for item in cfg["template_day"]:
        for fid in [item["food"]] + list(item.get("rotate", [])):
            if fid not in food_ids:
                raise ValueError("template_day uses unknown Food_ID %r" % fid)
    return cfg, source
