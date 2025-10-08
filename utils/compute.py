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
    try:
    df = pd.read_csv(file, delimiter=delim, dtype=str, quotechar='"', engine="python")
except Exception:
    file.seek(0)
    df = pd.read_csv(file, delimiter=None, engine="python", dtype=str)

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
    df["Product Cost"] = pd.to_numeric(df["Product Cost"], errors="coerce").fillna(0.0)
    df["Srp Price"] = pd.to_numeric(df["Srp Price"], errors="coerce").fillna(0.0)
    df["Product Key"] = df["Product Name"].str.upper().str.strip()
    return df

SMART_RULES = [
    (r"ARMOR", "BIG ARMOR"),
    (r"GRAPHENE", "GRAPHENE"),
    (r"WART", "WARTS REMOVER"),
    (r"BELO\\W*SHAM", "BELO SHAM 250"),
    (r"BELO\\W*CON", "BELO CON 250"),
    (r"(?:BD|TC).*(?:GOLD)", "BD/TC GOLD"),
    (r"(?:BD|TC).*(?:SILVER)", "BD/TC SILVER"),
]

def _smart_match_product(ad_name: str) -> str | None:
    if not isinstance(ad_name, str):
        return None
    name_up = ad_name.upper()
    for pattern, prod in SMART_RULES:
        if re.search(pattern, name_up, flags=re.IGNORECASE):
            return prod
    return None

def apply_v5_logic(ads_df: pd.DataFrame, costing_df: pd.DataFrame, multiplier: float = 1.25) -> pd.DataFrame:
    df = ads_df.copy()
    _require_columns(df)
    if "items" not in df.columns:
        df["items"] = 0.0
    df["Matched Product"] = df["ad_name"].apply(_smart_match_product)
    df["Product Key"] = df["Matched Product"].fillna("").str.upper()
    cost = costing_df[["Product Key", "Product Name", "Product Cost", "Srp Price"]].copy()
    df = df.merge(cost, on="Product Key", how="left")
    df["Match Status"] = np.where(df["Matched Product"].notna(), "Matched", "Unmatched")
    df["Profit per Item"] = (df["Srp Price"] - df["Product Cost"]).fillna(0.0)
    df["Net Profit"] = df["items"] * df["Profit per Item"] - df["expense"]
    df["Profit Margin %"] = np.where(df["gmv"] > 0, df["Net Profit"] / df["gmv"], 0.0)
    be_denom = (df["Srp Price"] - df["Product Cost"]).replace(0, np.nan)
    df["Break-even ROAS"] = np.where(be_denom.notna(), df["Srp Price"] / be_denom, np.inf)
    df["Suggested ROAS"] = df["Break-even ROAS"] * multiplier
    df["Decision (RUN/OFF)"] = np.where(df["roas"] >= df["Break-even ROAS"], "🟢 RUN", "🔴 OFF")
    cond_win = df["roas"] >= df["Suggested ROAS"]
    cond_opt = (df["roas"] >= df["Break-even ROAS"]) & (df["roas"] < df["Suggested ROAS"])
    df["Decision (WIN/OPT/LOSE)"] = np.select(
        [cond_win, cond_opt], ["🟢 WINNING", "🟡 OPTIMIZE"], default="🔴 LOSING"
    )
    return df

def split_deleted_and_active(df_logic: pd.DataFrame):
    df_deleted = df_logic[df_logic["status"].astype(str).str.upper() == "DELETED"].copy()
    df_active = df_logic[~df_logic.index.isin(df_deleted.index)].copy()
    df_unmatched = df_active[df_active["Match Status"] == "Unmatched"].copy()
    return df_active, df_deleted, df_unmatched

def compute_kpis(df_active: pd.DataFrame) -> dict:
    total_gmv = float(df_active["gmv"].sum()) if "gmv" in df_active else 0.0
    total_expense = float(df_active["expense"].sum()) if "expense" in df_active else 0.0
    total_net_profit = float(df_active["Net Profit"].sum()) if "Net Profit" in df_active else 0.0
    avg_roas = float(df_active["roas"].mean()) if "roas" in df_active else 0.0
    winning_cnt = int((df_active["Decision (WIN/OPT/LOSE)"] == "🟢 WINNING").sum())
    losing_cnt = int((df_active["Decision (WIN/OPT/LOSE)"] == "🔴 LOSING").sum())
    return {
        "total_gmv": total_gmv,
        "total_expense": total_expense,
        "total_net_profit": total_net_profit,
        "avg_roas": avg_roas,
        "winning_cnt": winning_cnt,
        "losing_cnt": losing_cnt,
    }

def style_preview_df(df: pd.DataFrame):
    def color_decision_run(val):
        if "RUN" in str(val): return "background-color:#D5F5E3"
        if "OFF" in str(val): return "background-color:#FADBD8"
        return ""
    def color_decision_status(val):
        if "WINNING" in str(val): return "background-color:#D5F5E3"
        if "OPTIMIZE" in str(val): return "background-color:#FCF3CF"
        if "LOSING" in str(val): return "background-color:#FADBD8"
        return ""
    def color_unmatched(val):
        if val == "Unmatched": return "background-color:#FFF3CD"
        return ""
    styler = df.style
    if "Decision (RUN/OFF)" in df.columns:
        styler = styler.applymap(color_decision_run, subset=["Decision (RUN/OFF)"])
    if "Decision (WIN/OPT/LOSE)" in df.columns:
        styler = styler.applymap(color_decision_status, subset=["Decision (WIN/OPT/LOSE)"])
    if "Match Status" in df.columns:
        styler = styler.applymap(color_unmatched, subset=["Match Status"])
    return styler
