import io
import csv
import re
import pandas as pd
import numpy as np

ALIASES = {
    "ad_name": ["Ad Name", "AdName", "Name", "Ad Title"],
    "status": ["Status"],
    "sequence": ["Sequence", "Seq"],
    "expense": ["Expense", "Spend", "Cost"],
    "gmv": ["GMV", "Sales", "Revenue"],
    "roas": ["ROAS", "Original ROAS", "Ad ROAS"],
    "items": ["Items Sold", "Orders", "Conversions", "Purchases", "Units Sold"],
}

REQUIRED_FOR_MIN = ["ad_name", "status", "sequence", "expense", "gmv", "roas"]

def _detect_delimiter(file) -> str:
    file.seek(0)
    sample = file.read(4096)
    if isinstance(sample, bytes):
        sample = sample.decode("utf-8", errors="ignore")
    file.seek(0)
    try:
        dialect = csv.Sniffer().sniff(sample)
        return dialect.delimiter
    except Exception:
        return ","

def _standardize_columns(df: pd.DataFrame) -> pd.DataFrame:
    cmap = {}
    lc_cols = {c.lower(): c for c in df.columns}
    for key, candidates in ALIASES.items():
        for cand in candidates:
            if cand in df.columns:
                cmap[cand] = key
                break
            if cand.lower() in lc_cols:
                cmap[lc_cols[cand.lower()]] = key
                break
    return df.rename(columns=cmap)

def _require_columns(df: pd.DataFrame):
    missing = [k for k in REQUIRED_FOR_MIN if k not in df.columns]
    if missing:
        raise ValueError(
            f"Missing required columns: {', '.join(missing)}. Make sure your Shopee Ads export includes them."
        )

def load_shopee_ads_csv(file):
    delim = _detect_delimiter(file)
    file.seek(0)
    df = pd.read_csv(file, delimiter=delim, dtype=str, quotechar='"', engine="python")
    df = _standardize_columns(df)
    _require_columns(df)
    for col in ["expense", "gmv", "roas", "items"]:
        if col in df.columns:
            df[col] = (
                df[col].astype(str)
                .str.replace(",", "", regex=False)
                .str.replace("₱", "", regex=False)
                .str.extract(r"([-+]?[0-9]*\\.?[0-9]+)")[0]
            )
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)
    if "sequence" in df.columns:
        df["sequence"] = df["sequence"].astype(str)
    return df, delim

def load_costing_table(file):
    name = getattr(file, "name", "").lower()
    file.seek(0)
    if name.endswith(".txt"):
        try:
            df = pd.read_csv(file, sep="\\t")
        except Exception:
            file.seek(0)
            df = pd.read_csv(file)
    else:
        df = pd.read_csv(file)
    df.columns = [c.strip().lower() for c in df.columns]
    mapping = {}
    for target in ["product name", "product cost", "srp price"]:
        for c in df.columns:
            if target.replace(" ", "") == c.replace(" ", ""):
                mapping[c] = target.title()
                break
    df = df.rename(columns=mapping)
    if not all(x.title() in df.columns for x in ["Product Name", "Product Cost", "Srp Price"]):
        raise ValueError("Costing file must have columns: Product Name, Product Cost, SRP Price")
    df["Product Cost"] = pd.to_numeric(_
