# -*- coding: utf-8 -*-
"""Shared loaders and statistics for the final project builders (docs + notebooks)."""
from pathlib import Path

import numpy as np
import pandas as pd
from scipy import stats

HERE = Path(__file__).resolve().parent
DATA = HERE / "data"
FIG = HERE / "figures"
SURNAME = "Moiseyenko_Svitlana"

URLS = {
    "Control": "https://drive.google.com/uc?export=download&id=1OlboUFclx-9QAakDNUAMzQquH8tNzo07",
    "Test": "https://drive.google.com/uc?export=download&id=1fFIA1tJj5YQs2G7tv244tgzLhpJ5axoG",
}
FILES = {"Control": "control_group.csv", "Test": "test_group.csv"}

RENAME = {
    "Campaign Name": "campaign", "Date": "date", "Spend [USD]": "spend",
    "# of Impressions": "impressions", "Reach": "reach", "# of Website Clicks": "clicks",
    "# of Searches": "searches", "# of View Content": "view_content",
    "# of Add to Cart": "add_to_cart", "# of Purchase": "purchase",
}
STEPS = ["impressions", "reach", "clicks", "searches", "view_content", "add_to_cart", "purchase"]
STEP_UA = {
    "impressions": "Покази", "reach": "Охоплення", "clicks": "Кліки на сайт", "searches": "Пошук на сайті",
    "view_content": "Перегляд контенту", "add_to_cart": "Додавання в кошик", "purchase": "Покупка",
}
METRIC_UA = {
    "spend": "Витрати, $ / день", "impressions": "Покази / день", "reach": "Охоплення / день",
    "clicks": "Кліки / день", "searches": "Пошуки / день", "view_content": "Перегляди контенту / день",
    "add_to_cart": "Додавання в кошик / день", "purchase": "Покупки / день",
    "ctr": "CTR (кліки / покази)", "cpc": "CPC, $ за клік", "cart_rate": "Конверсія клік → кошик",
    "cr_cart": "Конверсія кошик → покупка", "cr_click": "Конверсія клік → покупка", "cpa": "CPA, $ за покупку",
}

C_TEST, C_CONTROL = "#2a78d6", "#eb6834"
C_INK, C_MUTED, C_GRID = "#0b0b0b", "#898781", "#e1e0d9"


def load_group(group: str, data_dir: Path = DATA) -> pd.DataFrame:
    local = data_dir / FILES[group]
    src = local if local.exists() else URLS[group]
    d = pd.read_csv(src, sep=";").rename(columns=RENAME)
    d["date"] = pd.to_datetime(d["date"], format="%d.%m.%Y")
    d["group"] = group
    return d


def load_all(data_dir: Path = DATA) -> pd.DataFrame:
    df = pd.concat([load_group("Control", data_dir), load_group("Test", data_dir)], ignore_index=True)
    df["complete"] = df[STEPS].notna().all(axis=1)
    # a date is usable only if BOTH groups have complete data on it (same period for both arms)
    df["complete"] = df.groupby("date")["complete"].transform("all")
    return df


def add_daily_metrics(df: pd.DataFrame) -> pd.DataFrame:
    d = df.copy()
    d["ctr"] = d.clicks / d.impressions
    d["cpc"] = d.spend / d.clicks
    d["cart_rate"] = d.add_to_cart / d.clicks
    d["cr_cart"] = d.purchase / d.add_to_cart
    d["cr_click"] = d.purchase / d.clicks
    d["cpa"] = d.spend / d.purchase
    return d


def totals(df_complete: pd.DataFrame) -> pd.DataFrame:
    t = df_complete.groupby("group")[["spend"] + STEPS].sum().T[["Control", "Test"]]
    return t


def ratio_table(tot: pd.DataFrame) -> pd.DataFrame:
    rows = {}
    for g in tot.columns:
        s = tot[g]
        rows[g] = {
            "ctr": s.clicks / s.impressions, "cpc": s.spend / s.clicks, "cart_rate": s.add_to_cart / s.clicks,
            "cr_cart": s.purchase / s.add_to_cart, "cr_click": s.purchase / s.clicks, "cpa": s.spend / s.purchase,
        }
    return pd.DataFrame(rows)


def daily_tests(d: pd.DataFrame, metrics) -> pd.DataFrame:
    rows = []
    for m in metrics:
        a = d.loc[d.group == "Test", m].dropna()
        b = d.loc[d.group == "Control", m].dropna()
        w = stats.ttest_ind(a, b, equal_var=False)
        mw = stats.mannwhitneyu(a, b, alternative="two-sided")
        se = np.sqrt(a.var(ddof=1) / len(a) + b.var(ddof=1) / len(b))
        dof = se**4 / ((a.var(ddof=1) / len(a))**2 / (len(a) - 1) + (b.var(ddof=1) / len(b))**2 / (len(b) - 1))
        tcrit = stats.t.ppf(0.975, dof)
        rows.append({
            "metric": m, "Control": b.mean(), "Test": a.mean(), "diff": a.mean() - b.mean(),
            "rel": (a.mean() - b.mean()) / b.mean(), "ci_lo": a.mean() - b.mean() - tcrit * se,
            "ci_hi": a.mean() - b.mean() + tcrit * se, "t": w.statistic, "p_welch": w.pvalue, "p_mw": mw.pvalue,
            "n_test": len(a), "n_control": len(b),
        })
    return pd.DataFrame(rows).set_index("metric")


def two_prop(x1, n1, x2, n2, alpha=0.05):
    """Difference of proportions p1 - p2 (Test - Control): Wald CI + pooled z-test."""
    p1, p2 = x1 / n1, x2 / n2
    diff = p1 - p2
    se = np.sqrt(p1 * (1 - p1) / n1 + p2 * (1 - p2) / n2)
    z = stats.norm.ppf(1 - alpha / 2)
    pp = (x1 + x2) / (n1 + n2)
    se_pooled = np.sqrt(pp * (1 - pp) * (1 / n1 + 1 / n2))
    zstat = diff / se_pooled
    return {"p1": p1, "p2": p2, "diff": diff, "lo": diff - z * se, "hi": diff + z * se,
            "z": zstat, "p": 2 * stats.norm.sf(abs(zstat)), "se": se}


def sample_size_two_prop(p1, p2, alpha=0.05, power=0.8):
    za, zb = stats.norm.ppf(1 - alpha / 2), stats.norm.ppf(power)
    pbar = (p1 + p2) / 2
    n = (za * np.sqrt(2 * pbar * (1 - pbar)) + zb * np.sqrt(p1 * (1 - p1) + p2 * (1 - p2)))**2 / (p1 - p2)**2
    return int(np.ceil(n))
