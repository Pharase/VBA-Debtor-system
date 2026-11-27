"""Shared settings for the payment scripts: .env loading, folders and legacy column names.

Copy .env.template to .env and fill it in. Every script imports what it needs from here,
so no path, account or company name is hard-coded in the scripts.
"""
import os

import pandas as pd

HERE = os.path.dirname(os.path.abspath(__file__))


def load_env(path=os.path.join(HERE, ".env")):
    """Read KEY=VALUE lines from .env into the environment (no extra package needed)."""
    if os.path.exists(path):
        with open(path, encoding="utf-8") as f:
            for line in f:
                if "=" in line and not line.lstrip().startswith("#"):
                    key, value = line.strip().split("=", 1)
                    os.environ.setdefault(key.strip(), value.strip())


load_env()

# Local working folder (holds summary/, payment/, processing/program/ ...)
DATA_ROOT = os.environ.get("DATA_ROOT", os.path.join(HERE, "data"))
# Shared team folder (e.g. a synced SharePoint/OneDrive library) where finished files are copied
SHARED_DIR = os.environ.get("SHARED_DIR", os.path.join(DATA_ROOT, "shared"))
LEGACY_COLUMN_PREFIX = os.environ.get("LEGACY_COLUMN_PREFIX", "").lower()


def data_path(*parts):
    """Path inside DATA_ROOT, e.g. data_path("summary", "file.xlsx")."""
    return os.path.join(DATA_ROOT, *[p.strip("\\/") for p in parts if p])


def shared_path(*parts):
    """Path inside SHARED_DIR."""
    return os.path.join(SHARED_DIR, *[p.strip("\\/") for p in parts if p])


def neutral_name(column):
    """<prefix>_code / <prefix>code -> account_code, <prefix>xyz -> xyz. Other names are unchanged."""
    if column is None:
        return column
    name = str(column)
    if not LEGACY_COLUMN_PREFIX or not name.lower().startswith(LEGACY_COLUMN_PREFIX):
        return name
    rest = name[len(LEGACY_COLUMN_PREFIX):].lstrip("_")
    if rest.lower().startswith("code"):
        return "account_" + rest.lower()
    return "account_id" if rest.lower() == "id" else rest


def read_excel(*args, **kwargs):
    """pandas.read_excel, with legacy column names renamed (works for one sheet or a dict of sheets).
    A usecols list may use the new names even when the file still has the old ones."""
    usecols = kwargs.get("usecols")
    if LEGACY_COLUMN_PREFIX and isinstance(usecols, (list, tuple)) and all(isinstance(c, str) for c in usecols):
        kwargs.pop("usecols")
        df = pd.read_excel(*args, **kwargs).rename(columns=neutral_name)
        return df[list(usecols)]
    result = pd.read_excel(*args, **kwargs)
    if isinstance(result, dict):
        return {name: df.rename(columns=neutral_name) for name, df in result.items()}
    return result.rename(columns=neutral_name)
