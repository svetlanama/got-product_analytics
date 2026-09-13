# -*- coding: utf-8 -*-
"""Build Moiseyenko_Svitlana_PA_assignment_3.xlsx: EDA of a synthetic
international genomics / bioinformatics e-commerce store (BioSeq Market)."""
import numpy as np, pandas as pd, datetime as dt, math
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import LineChart, BarChart, ScatterChart, Reference, Series
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.drawing.line import LineProperties
from openpyxl.chart.axis import ChartLines
from openpyxl.chart.marker import Marker
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.worksheet.formula import ArrayFormula

OUT = "/Users/svitlanamoiseyenko/Repos/GOIT/got-product_analytics/hw_03/Moiseyenko_Svitlana_PA_assignment_3.xlsx"
rng = np.random.default_rng(42)

# ----------------------------------------------------------------- palette
BLUE, ORANGE, AQUA, YELLOW, MAGENTA, GREEN, VIOLET, RED = (
    "2A78D6", "EB6834", "1BAF7A", "EDA100", "E87BA4", "008300", "4A3AA7", "E34948")
CAT = [BLUE, ORANGE, AQUA, YELLOW, MAGENTA, GREEN, VIOLET, RED]
GRID, MUTED, INK2 = "E1E0D9", "898781", "52514E"
SEQ_LO, SEQ_HI = "CDE2FB", "2A78D6"

# ----------------------------------------------------------------- dataset
START, END = dt.date(2023, 1, 1), dt.date(2025, 6, 30)
DAYS = pd.date_range(START, END, freq="D")

SEGMENTS = {  # segment: (share, orders/yr mean, discount, qty mult)
    "Academic":  (0.45, 3.0, 0.05, 1.0),
    "Biotech":   (0.20, 4.5, 0.08, 1.4),
    "Pharma":    (0.12, 6.0, 0.13, 2.6),
    "Clinical":  (0.13, 4.0, 0.06, 1.6),
    "CRO":       (0.10, 5.0, 0.10, 2.0),
}
REGIONS = {
    "Europe": (0.40, ["Germany", "United Kingdom", "France", "Netherlands", "Switzerland", "Sweden", "Ukraine", "Poland", "Spain"]),
    "North America": (0.28, ["USA", "USA", "USA", "Canada"]),
    "Asia-Pacific": (0.20, ["Japan", "China", "Australia", "Singapore", "South Korea", "India"]),
    "Latin America": (0.06, ["Brazil", "Mexico", "Argentina"]),
    "Middle East & Africa": (0.06, ["Israel", "UAE", "South Africa"]),
}
PRODUCTS = {  # category: [(name, price)], qty lognormal params, month weights
    "Library Prep Kits": ([("NGS Library Prep Kit (24 rxn)", 1200), ("RNA-seq Library Kit (48 rxn)", 2400), ("16S Amplicon Kit (96 rxn)", 650)], (0.6, 0.5)),
    "DNA/RNA Extraction": ([("Genomic DNA Extraction Kit (250 preps)", 480), ("Total RNA Kit (50 preps)", 320), ("Plasmid Miniprep Kit (250 preps)", 210)], (0.8, 0.6)),
    "Oligos & Primers": ([("Custom Primer Pair", 25), ("Sequencing Adapter Set", 380), ("qPCR Probe (dual-labeled)", 95)], (1.6, 0.8)),
    "Reagents & Enzymes": ([("High-Fidelity DNA Polymerase", 180), ("Reverse Transcriptase", 240), ("dNTP Mix 10 mM", 60)], (1.0, 0.7)),
    "Sequencing Services": ([("WGS 30x (per sample)", 650), ("RNA-seq (per sample)", 280), ("Exome (per sample)", 420)], (2.0, 0.8)),
    "Bioinformatics Software": ([("Genome Analysis Suite - annual license", 3600), ("Variant Caller Pro - annual license", 1800), ("Pipeline Manager - annual license", 950)], (0.1, 0.4)),
    "Cloud Compute": ([("Compute credits 1000 CPU-h", 150), ("GPU credits 100 h", 300), ("Storage 1 TB-year", 120)], (1.2, 0.7)),
    "Training": ([("Bioinformatics Bootcamp seat", 900), ("NGS Data Analysis Workshop", 450)], (0.3, 0.5)),
}
CATS = list(PRODUCTS)
SEG_CAT_PREF = {  # relative preference of categories per segment (same order as CATS)
    "Academic":  [1.0, 1.4, 1.6, 1.3, 0.9, 0.5, 0.8, 1.0],
    "Biotech":   [1.3, 1.0, 1.1, 1.1, 1.0, 1.0, 1.2, 0.5],
    "Pharma":    [1.6, 0.8, 0.7, 0.9, 1.4, 1.4, 1.3, 0.3],
    "Clinical":  [1.2, 1.3, 1.0, 1.0, 1.3, 0.8, 0.4, 0.4],
    "CRO":       [1.4, 1.0, 0.9, 1.0, 1.5, 0.9, 1.0, 0.3],
}
CAT_MONTH = {c: np.ones(12) for c in CATS}
CAT_MONTH["Bioinformatics Software"] = np.array([2.6, 0.6, 0.7, 0.7, 0.6, 0.8, 2.2, 0.6, 0.7, 0.8, 0.7, 0.9])  # Jan/Jul renewals
CAT_MONTH["Training"] = np.array([0.6, 1.3, 1.2, 0.8, 0.9, 0.7, 0.5, 0.6, 1.5, 1.4, 1.0, 0.4])
MONTH_SEAS = np.array([0.85, 0.95, 1.10, 1.00, 1.05, 1.15, 0.85, 0.75, 1.30, 1.15, 1.10, 0.80])
WDAY = np.array([1.10, 1.20, 1.15, 1.10, 0.95, 0.15, 0.10])
DISC_P = {"Academic": [0.45, 0.40, 0.12, 0.03, 0.0], "Biotech": [0.30, 0.35, 0.25, 0.08, 0.02], "Pharma": [0.10, 0.20, 0.35, 0.25, 0.10],
          "Clinical": [0.35, 0.35, 0.22, 0.06, 0.02], "CRO": [0.20, 0.30, 0.30, 0.15, 0.05]}
CHANNELS = (["Web", "Sales rep", "Distributor"], [0.62, 0.23, 0.15])

# customers ------------------------------------------------------------
N_CUST = 900
seg_names = list(SEGMENTS)
cust_seg = rng.choice(seg_names, N_CUST, p=[SEGMENTS[s][0] for s in seg_names])
reg_names = list(REGIONS)
cust_reg = rng.choice(reg_names, N_CUST, p=[REGIONS[r][0] for r in reg_names])
cust_country = np.array([rng.choice(REGIONS[r][1]) for r in cust_reg])
# acquisition dates: growing intake, none in last 2 months so cohorts are meaningful
acq_idx = np.concatenate([rng.integers(0, 120, 300), rng.integers(120, len(DAYS), N_CUST - 300)])  # initial base + steady inflow
acq_idx.sort()
cust_acq = DAYS[acq_idx]
cust_act = np.exp(rng.normal(0, 0.8, N_CUST))                       # heterogeneity of ordering activity
cust_life = rng.exponential(420, N_CUST)                            # days until churn (mean ~14 months)
cust_churn = cust_acq + pd.to_timedelta(cust_life, unit="D")
cust_ids = [f"C{1001+i}" for i in range(N_CUST)]

def day_intensity(d):
    yrs = (d.date() - START).days / 365.25
    trend = 1.18 ** yrs
    m = MONTH_SEAS[d.month - 1]
    w = WDAY[d.weekday()]
    me = 1.3 if d.day >= d.days_in_month - 2 else 1.0
    if d.month == 12 and d.day >= 22: m *= 0.45
    if d.month == 1 and d.day <= 7: m *= 0.6
    return trend * m * w * me

orders = []
oid = 100001
def make_order(d, ci, forced=None):
    global oid
    seg = cust_seg[ci]
    pref = np.array(SEG_CAT_PREF[seg]) * np.array([CAT_MONTH[c][d.month - 1] for c in CATS])
    cat = forced or rng.choice(CATS, p=pref / pref.sum())
    prods, (mu, sig) = PRODUCTS[cat]
    name, price = prods[rng.integers(len(prods))]
    qty = max(1, int(round(np.exp(rng.normal(mu, sig)) * SEGMENTS[seg][3])))
    unit = float(price)  # catalogue price
    disc = float(rng.choice([0.0, 0.05, 0.10, 0.15, 0.20], p=DISC_P[seg]))
    channel = rng.choice(CHANNELS[0], p=CHANNELS[1])
    rev = round(qty * unit * (1 - disc), 2)
    orders.append(dict(order_id=f"O{oid}", order_date=d.date(), customer_id=cust_ids[ci], segment=seg,
                       country=cust_country[ci], region=cust_reg[ci], channel=channel, category=cat,
                       product=name, quantity=qty, unit_price=unit, discount_pct=disc, revenue=rev))
    oid += 1

# first order for every customer on acquisition day
for ci in range(N_CUST):
    make_order(cust_acq[ci], ci)
# subsequent orders
seg_rate = np.array([SEGMENTS[s][1] for s in cust_seg]) * cust_act
BASE = 7.0  # orders per day-equivalent at t0
for d in DAYS:
    active = np.where((cust_acq < d) & (cust_churn > d))[0]
    if len(active) == 0: continue
    lam = BASE * day_intensity(d) * (len(active) / 300) ** 0.5
    n = rng.poisson(lam)
    if n == 0: continue
    w = seg_rate[active]; w = w / w.sum()
    for ci in rng.choice(active, n, p=w):
        make_order(d, ci)

df = pd.DataFrame(orders)

# ---- business anomalies (kept in clean data, explained in the report)
# A1 bulk tender order from a pharma customer
ph = np.where(cust_seg == "Pharma")[0][3]
d1 = pd.Timestamp("2024-11-14")
orders_before = len(df)
make_order(d1, ph, forced="Library Prep Kits")
df = pd.DataFrame(orders)
df.loc[len(df) - 1, ["product", "quantity", "unit_price", "discount_pct", "channel"]] = ["NGS Library Prep Kit (24 rxn)", 200, 1200.0, 0.20, "Sales rep"]
df.loc[len(df) - 1, "revenue"] = round(200 * 1200.0 * 0.80, 2)
# A2 reagents stock-out 2024-04-08..2024-05-05 : drop 90% of reagent orders
m = (df.category == "Reagents & Enzymes") & (df.order_date >= dt.date(2024, 4, 8)) & (df.order_date <= dt.date(2024, 5, 5))
df = df.drop(df[m].sample(frac=0.9, random_state=1).index)
# A3 "Lab Week" promo 2024-11-25..29: extra oligo/extraction orders at 20% discount
promo_days = pd.date_range("2024-11-25", "2024-11-29")
extra = []
for d in promo_days:
    active = np.where((cust_acq < d) & (cust_churn > d))[0]
    for ci in rng.choice(active, 38):
        make_order(d, ci, forced=rng.choice(["Oligos & Primers", "DNA/RNA Extraction"]))
promo = pd.DataFrame(orders[orders_before + 1:])
promo["discount_pct"] = 0.20
promo["revenue"] = (promo.quantity * promo.unit_price * 0.8).round(2)
promo["channel"] = "Web"; promo["is_promo"] = True
df = pd.concat([df, promo], ignore_index=True); df["is_promo"] = df.is_promo.fillna(False).astype(bool)
# A4 web outage 2023-08-15..17: no Web orders
m = (df.order_date >= dt.date(2023, 8, 15)) & (df.order_date <= dt.date(2023, 8, 16))
df = df.drop(df[m].index)

df["order_date"] = pd.to_datetime(df.order_date)
df = df.sort_values(["order_date", "order_id"]).reset_index(drop=True)
df["order_id"] = [f"O{100001+i}" for i in range(len(df))]

# ---- data-quality anomalies -> raw file only
raw = df.copy()
dups = raw.sample(12, random_state=2)                                   # duplicated rows
zero = raw.sample(6, random_state=3).copy(); zero["unit_price"] = 0.0; zero["revenue"] = 0.0; zero["order_id"] = [f"T{900001+i}" for i in range(6)]
ret = raw.sample(15, random_state=4).copy(); ret["quantity"] = -ret.quantity; ret["revenue"] = -ret.revenue
ret["order_id"] = [f"R{800001+i}" for i in range(15)]; ret["order_date"] = ret.order_date + pd.to_timedelta(rng.integers(5, 30, 15), unit="D")
raw = pd.concat([raw, dups, zero, ret], ignore_index=True)
miss_idx = raw.sample(8, random_state=5).index; raw.loc[miss_idx, "country"] = None
typo_idx = raw.sample(3, random_state=6).index; raw.loc[typo_idx, "order_date"] = raw.loc[typo_idx, "order_date"] + pd.DateOffset(years=9)
raw = raw.sort_values(["order_date", "order_id"]).reset_index(drop=True)
raw["order_date"] = raw.order_date.dt.date
# how the clean set is derived from raw (for the report)
raw_n = len(raw)
n_dup = int(raw.duplicated().sum()); n_zero = int((raw.unit_price == 0).sum())
n_ret = int((raw.quantity < 0).sum()); n_miss = int(raw.country.isna().sum())
n_typo = int((pd.to_datetime(raw.order_date) > pd.Timestamp(END)).sum())
# country can be recovered from customer_id -> keep those rows; others removed
clean_n = len(df)

# ---- derived columns
UA_M = ["Січ", "Лют", "Бер", "Кві", "Тра", "Чер", "Лип", "Сер", "Вер", "Жов", "Лис", "Гру"]
UA_W = ["Пн", "Вт", "Ср", "Чт", "Пт", "Сб", "Нд"]
df["year"] = df.order_date.dt.year
df["month"] = df.order_date.dt.month
df["month_name"] = df.month.map(lambda m: UA_M[m - 1])
df["quarter"] = "Q" + df.order_date.dt.quarter.astype(str)
df["year_month"] = df.order_date.dt.strftime("%Y-%m")
df["weekday"] = df.order_date.dt.weekday.map(lambda w: UA_W[w])
df["is_weekend"] = (df.order_date.dt.weekday >= 5).map({True: "Так", False: "Ні"})
first = df.groupby("customer_id").order_date.min()
df["first_order_month"] = df.customer_id.map(first).dt.strftime("%Y-%m")
df["cohort_quarter"] = df.customer_id.map(first).dt.to_period("Q").astype(str)
df["customer_type"] = np.where(df.order_date == df.customer_id.map(first), "Новий", "Повторний")

# ---- customers & RFM
REF = pd.Timestamp(END) + pd.Timedelta(days=1)
cust = df.groupby("customer_id").agg(segment=("segment", "first"), country=("country", "first"), region=("region", "first"),
                                     first_order=("order_date", "min"), last_order=("order_date", "max"),
                                     orders=("order_id", "count"), revenue=("revenue", "sum"), quantity=("quantity", "sum")).reset_index()
cust["recency_days"] = (REF - cust.last_order).dt.days
cust["avg_order_value"] = (cust.revenue / cust.orders).round(2)
def score(s, ascending):
    q = s.rank(method="first", ascending=ascending)
    return pd.qcut(q, 3, labels=[1, 2, 3]).astype(int)
cust["R"] = score(cust.recency_days, ascending=False)  # smaller recency -> higher score
cust["F"] = score(cust.orders, True); cust["M"] = score(cust.revenue, True)
cust["rfm_score"] = cust.R + cust.F + cust.M
cust["rfm_tier"] = pd.cut(cust.rfm_score, [0, 5, 7, 9], labels=["C – разові / сплячі", "B – активні", "A – ключові"]).astype(str)
cust["repeat"] = np.where(cust.orders > 1, "Повторний", "Разовий")
df["rfm_tier"] = df.customer_id.map(cust.set_index("customer_id").rfm_tier)

# ================================================================= analysis
monthly = df.groupby("year_month").agg(orders=("order_id", "count"), revenue=("revenue", "sum"),
                                       customers=("customer_id", "nunique"), quantity=("quantity", "sum")).reset_index()
monthly["aov"] = (monthly.revenue / monthly.orders).round(2)
monthly["ma3"] = monthly.revenue.rolling(3, center=True, min_periods=2).mean().round(2)
monthly["mom_pct"] = (monthly.revenue.pct_change() * 100).round(1)
# robust: exclude the single bulk order for the seasonality index
bulk_id = df.loc[df.revenue.idxmax(), "order_id"]
df_nb = df[(df.order_id != bulk_id) & (~df.is_promo)]   # without the tender order and the Lab Week promo spike
ym_all = df.groupby(["year", "month"]).revenue.sum().unstack(0)
mrev = df_nb.groupby("year_month").revenue.sum()
cma = mrev.rolling(12, center=True).mean().rolling(2).mean().shift(-1)   # centred 12-month MA (2x12)
ratio = (mrev / cma).dropna()
ratio.index = pd.to_datetime(ratio.index).month
seas_idx = ratio.groupby(ratio.index).mean()
seas_idx = (seas_idx / seas_idx.mean()).round(3)
wd = df.groupby(df.order_date.dt.weekday).agg(orders=("order_id", "count"), revenue=("revenue", "sum")).reindex(range(7))
wd.index = UA_W
n_weeks = len(DAYS) / 7
wd["orders_per_day"] = (wd.orders / n_weeks).round(1)
wd["revenue_per_day"] = (wd.revenue / n_weeks).round(0)
qtr = df.groupby(["year", "quarter"]).agg(orders=("order_id", "count"), revenue=("revenue", "sum")).reset_index()
dom = df_nb.groupby(df_nb.order_date.dt.day).revenue.sum()
dom = (dom / df_nb.order_date.dt.day.value_counts().sort_index().index.map(
    lambda d: sum(1 for x in DAYS if x.day == d))).round(0)  # avg revenue per calendar day-of-month
cat_month = df_nb.groupby(["category", "month"]).revenue.sum().unstack(1)
cat_month_idx = cat_month.div(cat_month.mean(axis=1), axis=0).round(2)
cat_year = df.groupby(["category", "year"]).revenue.sum().unstack(1).round(0)

# anomalies: daily revenue, robust rolling z (business days)
daily = df.groupby("order_date").agg(orders=("order_id", "count"), revenue=("revenue", "sum")).reindex(DAYS, fill_value=0)
daily.index.name = "date"
bd = daily[daily.index.weekday < 5].copy()
lg = np.log(bd.revenue.clip(lower=1))
med = lg.rolling(41, center=True, min_periods=20).median()
mad = (lg - med).abs().rolling(41, center=True, min_periods=20).median()
ZT = 3.0
bd["expected"] = np.exp(med).round(0); bd["z"] = (0.6745 * (lg - med) / mad.replace(0, np.nan)).round(2)
bd["upper"] = np.exp(med + ZT * mad / 0.6745).round(0); bd["lower"] = np.exp(med - ZT * mad / 0.6745).round(0)
flag = bd[bd.z.abs() >= ZT].copy()
def explain(d, z):
    if d == pd.Timestamp("2024-11-14"): return "Разове тендерне замовлення Pharma: 200 наборів NGS Library Prep (~$192 тис.)"
    if z > 0 and pd.Timestamp("2024-11-25") <= d <= pd.Timestamp("2024-11-29"): return "Промо-акція Lab Week (-20% на олігонуклеотиди та набори для екстракції)"
    if z < 0 and pd.Timestamp("2023-08-15") <= d <= pd.Timestamp("2023-08-16"): return "Збій ERP/сайту 15–16.08.2023: замовлення не оброблялися"
    if z < 0 and pd.Timestamp("2024-04-08") <= d <= pd.Timestamp("2024-05-05"): return "Дефіцит реагентів (stock-out) — падіння категорії Reagents & Enzymes"
    if z < 0 and (d.month == 12 and d.day >= 22 or (d.month == 1 and d.day <= 7)): return "Святковий період (очікуване зниження)"
    if z > 0 and (d.day >= 28 or (d.month in (6, 9) and d.day >= 20)): return "Кінець місяця / кварталу (закриття бюджетів та грантів) — сезонний ефект"
    if z > 0: return "Кілька великих замовлень одного дня — потребує перевірки"
    return "Низький день — потребує перевірки (не пояснено відомими подіями)"
flag["comment"] = [explain(d, z) for d, z in zip(flag.index, flag.z)]
# order-level outliers by IQR rule
q1, q3 = df.revenue.quantile([0.25, 0.75]); iqr = q3 - q1
out_thr = q3 + 3 * iqr
outliers = df[df.revenue > out_thr].sort_values("revenue", ascending=False)
# monthly reagents stock-out evidence
reag = df[df.category == "Reagents & Enzymes"].groupby("year_month").agg(orders=("order_id", "count"), revenue=("revenue", "sum")).reset_index()

# cohorts
seg_tbl = df.groupby("segment").agg(customers=("customer_id", "nunique"), orders=("order_id", "count"), revenue=("revenue", "sum"),
                                    quantity=("quantity", "sum"), discount=("discount_pct", "mean")).reset_index()
seg_tbl["share_pct"] = (seg_tbl.revenue / seg_tbl.revenue.sum() * 100).round(1)
seg_tbl["aov"] = (seg_tbl.revenue / seg_tbl.orders).round(2)
seg_tbl["orders_per_customer"] = (seg_tbl.orders / seg_tbl.customers).round(2)
seg_tbl["revenue_per_customer"] = (seg_tbl.revenue / seg_tbl.customers).round(0)
seg_tbl = seg_tbl.sort_values("revenue", ascending=False)
reg_tbl = df.groupby("region").agg(customers=("customer_id", "nunique"), orders=("order_id", "count"), revenue=("revenue", "sum")).reset_index()
reg_tbl["share_pct"] = (reg_tbl.revenue / reg_tbl.revenue.sum() * 100).round(1)
reg_tbl["aov"] = (reg_tbl.revenue / reg_tbl.orders).round(2)
reg_tbl["orders_per_customer"] = (reg_tbl.orders / reg_tbl.customers).round(2)
reg_tbl = reg_tbl.sort_values("revenue", ascending=False)
country_tbl = df.groupby("country").agg(customers=("customer_id", "nunique"), orders=("order_id", "count"), revenue=("revenue", "sum")).reset_index().sort_values("revenue", ascending=False).head(12)
rfm_tbl = cust.groupby("rfm_tier").agg(customers=("customer_id", "count"), orders=("orders", "sum"), revenue=("revenue", "sum"),
                                       avg_recency=("recency_days", "mean"), avg_orders=("orders", "mean")).reset_index()
rfm_tbl["share_pct"] = (rfm_tbl.revenue / rfm_tbl.revenue.sum() * 100).round(1)
rfm_tbl = rfm_tbl.sort_values("rfm_tier")
rep_tbl = cust.groupby("repeat").agg(customers=("customer_id", "count"), revenue=("revenue", "sum")).reset_index()
rep_tbl["share_pct"] = (rep_tbl.revenue / rep_tbl.revenue.sum() * 100).round(1)
seg_cat = df.groupby(["segment", "category"]).revenue.sum().unstack(1)
seg_cat_pct = (seg_cat.div(seg_cat.sum(axis=1), axis=0) * 100).round(1)
chan_tbl = df.groupby("channel").agg(orders=("order_id", "count"), revenue=("revenue", "sum")).reset_index()
chan_tbl["share_pct"] = (chan_tbl.revenue / chan_tbl.revenue.sum() * 100).round(1)
# retention by acquisition quarter
cq = df[["customer_id", "cohort_quarter", "order_date"]].copy()
cq["order_q"] = cq.order_date.dt.to_period("Q")
cq["cohort_p"] = pd.PeriodIndex(cq.cohort_quarter, freq="Q")
cq["k"] = (cq.order_q - cq.cohort_p).apply(lambda x: x.n)
ret_cnt = cq.groupby(["cohort_quarter", "k"]).customer_id.nunique().unstack(1)
cohort_size = ret_cnt[0]
ret_pct = (ret_cnt.div(cohort_size, axis=0) * 100).round(1)
maxk = {c: (pd.Period("2025Q2") - pd.Period(c)).n for c in ret_pct.index}
for c in ret_pct.index:
    for k in ret_pct.columns:
        if k > maxk[c]: ret_pct.loc[c, k] = np.nan
    ret_pct.loc[c, [k for k in ret_pct.columns if k <= maxk[c]]] = ret_pct.loc[c, [k for k in ret_pct.columns if k <= maxk[c]]].fillna(0)
cohort_rev = df.groupby("cohort_quarter").revenue.sum()

# descriptive stats (python)
def dstats(s):
    s = s.astype(float)
    return {"Кількість (n)": int(s.count()), "Сума": round(s.sum(), 2), "Середнє": round(s.mean(), 2), "Медіана": round(s.median(), 2),
            "Мода": round(float(s.mode().iloc[0]), 2), "Мінімум": round(s.min(), 2), "Максимум": round(s.max(), 2), "Розмах": round(s.max() - s.min(), 2),
            "Дисперсія": round(s.var(ddof=1), 2), "Стандартне відхилення": round(s.std(ddof=1), 2), "Коеф. варіації, %": round(s.std(ddof=1) / s.mean() * 100, 1),
            "Q1 (25%)": round(s.quantile(0.25), 2), "Q3 (75%)": round(s.quantile(0.75), 2), "IQR": round(s.quantile(0.75) - s.quantile(0.25), 2),
            "Асиметрія (skewness)": round(s.skew(), 3), "Ексцес (kurtosis)": round(s.kurt(), 3)}
STAT_ROWS = list(dstats(df.revenue))
groups = {"Весь датасет": df}
for s in seg_tbl.segment: groups[f"Сегмент: {s}"] = df[df.segment == s]
for r in reg_tbl.region: groups[f"Регіон: {r}"] = df[df.region == r]
for t in sorted(cust.rfm_tier.unique()): groups[f"RFM: {t}"] = df[df.rfm_tier == t]
stats_rev = pd.DataFrame({g: dstats(d.revenue) for g, d in groups.items()})
stats_qty = pd.DataFrame({g: dstats(d.quantity) for g, d in groups.items()})
stats_cust = pd.DataFrame({"Замовлень на клієнта": dstats(cust.orders), "Дохід на клієнта": dstats(cust.revenue), "Давність, днів": dstats(cust.recency_days)})
# histogram bins
bins = [0, 100, 250, 500, 1000, 2500, 5000, 10000, 25000, 50000, 1e9]
labels = ["0–100", "100–250", "250–500", "500–1 000", "1 000–2 500", "2 500–5 000", "5 000–10 000", "10 000–25 000", "25 000–50 000", "> 50 000"]
hist = pd.cut(df.revenue, bins, labels=labels, right=False).value_counts().reindex(labels)
# box plot data per segment
box = pd.DataFrame({s: df[df.segment == s].revenue.describe(percentiles=[.25, .5, .75]) for s in seg_names}).T[["min", "25%", "50%", "75%", "max"]]

print(df.groupby("year").revenue.sum().round()); print("H1 2024 vs H1 2025:", df[(df.year==2024)&(df.month<=6)].revenue.sum().round(), df[(df.year==2025)].revenue.sum().round())
print("orders clean", len(df), "raw", raw_n, "customers", cust.shape[0], "revenue", round(df.revenue.sum()))
print("removed: dup", n_dup, "zero", n_zero, "ret", n_ret, "miss", n_miss, "typo", n_typo)
print(monthly.tail(3)); print("seasonality index\n", seas_idx.round(2).to_dict())
print("weekday\n", wd); print("flags\n", flag[["revenue", "expected", "z", "comment"]])
print("outliers > ", round(out_thr), len(outliers)); print(seg_tbl); print(reg_tbl); print(rfm_tbl); print(rep_tbl)
print(ret_pct); print(stats_rev.iloc[:, :3]); print(hist); print(chan_tbl)

# ================================================================= workbook
wb = Workbook()
H_FILL = PatternFill("solid", fgColor="1C5CAB"); H_FONT = Font(bold=True, color="FFFFFF")
T_FONT = Font(bold=True, size=14); S_FONT = Font(bold=True, size=11, color="1C5CAB")
NOTE_FONT = Font(italic=True, color=INK2, size=9)
thin = Side(style="thin", color=GRID); BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
def header(ws, row, col, names, widths=None):
    for j, n in enumerate(names):
        c = ws.cell(row=row, column=col + j, value=n); c.fill = H_FILL; c.font = H_FONT
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True); c.border = BORDER
def write_df(ws, d, row, col, fmts=None, index=False, index_name=None):
    """write dataframe with header; returns (first_data_row, last_row)."""
    cols = ([index_name or (d.index.name or "")] if index else []) + [str(c) for c in d.columns]
    header(ws, row, col, cols)
    for i, (idx, r) in enumerate(d.iterrows()):
        vals = ([idx] if index else []) + list(r.values)
        for j, v in enumerate(vals):
            if isinstance(v, (np.floating, float)) and (np.isnan(v)): v = None
            elif isinstance(v, np.generic): v = v.item()
            elif isinstance(v, pd.Timestamp): v = v.date()
            c = ws.cell(row=row + 1 + i, column=col + j, value=v); c.border = BORDER
            if fmts and cols[j] in fmts: c.number_format = fmts[cols[j]]
    return row + 1, row + len(d)
def title(ws, text, sub=None):
    ws["A1"] = text; ws["A1"].font = T_FONT
    ws["A2"] = "Moiseyenko Svitlana · Домашнє завдання 3 · Експлоративний аналіз даних"; ws["A2"].font = Font(color=INK2)
    if sub: ws["A3"] = sub; ws["A3"].font = NOTE_FONT
def section(ws, cell, text): ws[cell] = text; ws[cell].font = S_FONT
def note(ws, cell, text): ws[cell] = text; ws[cell].font = NOTE_FONT
def widths(ws, w):
    for i, x in enumerate(w, 1): ws.column_dimensions[get_column_letter(i)].width = x
def style_axes(ch, ytitle=None, xtitle=None, yfmt="#,##0"):
    ch.x_axis.delete = False; ch.y_axis.delete = False
    ch.y_axis.majorGridlines = ChartLines(spPr=GraphicalProperties(ln=LineProperties(solidFill=GRID)))
    ch.x_axis.majorGridlines = None
    ch.y_axis.number_format = yfmt
    if ytitle: ch.y_axis.title = ytitle
    if xtitle: ch.x_axis.title = xtitle
    ch.x_axis.graphicalProperties = GraphicalProperties(ln=LineProperties(solidFill="C3C2B7"))
    ch.y_axis.graphicalProperties = GraphicalProperties(ln=LineProperties(noFill=True))
def line_chart(ws, anchor, ttl, ws_src, cats, series, colors, w=22, h=9, ytitle=None, dashes=None, legend=True):
    ch = LineChart(); ch.title = ttl; ch.width, ch.height = w, h
    for k, (ref, name) in enumerate(series):
        s = Series(ref, title=name); s.smooth = False
        s.graphicalProperties.line.solidFill = colors[k]; s.graphicalProperties.line.width = 19050
        if dashes and dashes[k]: s.graphicalProperties.line.dashStyle = dashes[k]
        s.marker.symbol = "none"; ch.series.append(s)
    ch.set_categories(cats); style_axes(ch, ytitle)
    if legend: ch.legend.position = "b"
    else: ch.legend = None
    ws.add_chart(ch, anchor); return ch
def bar_chart(ws, anchor, ttl, cats, series, colors, w=16, h=8, ytitle=None, horizontal=False, stacked=False, legend=True, yfmt="#,##0"):
    ch = BarChart(); ch.type = "bar" if horizontal else "col"; ch.title = ttl; ch.width, ch.height = w, h
    if stacked: ch.grouping = "stacked"; ch.overlap = 100
    ch.gapWidth = 60
    for k, (ref, name) in enumerate(series):
        s = Series(ref, title=name)
        s.graphicalProperties.solidFill = colors[k]; s.graphicalProperties.line.noFill = True
        ch.series.append(s)
    ch.set_categories(cats); style_axes(ch, ytitle, yfmt=yfmt)
    if legend: ch.legend.position = "b"
    else: ch.legend = None
    ws.add_chart(ch, anchor); return ch
def heat(ws, rng_str, lo=SEQ_LO, hi=SEQ_HI):
    ws.conditional_formatting.add(rng_str, ColorScaleRule(start_type="min", start_color=lo, end_type="max", end_color=hi))

# ---------------------------------------------------------------- Опис
ws = wb.active; ws.title = "Опис"
title(ws, "Домашнє завдання 3. Експлоративний аналіз даних (EDA)")
rows = [
 ("Продукт", "BioSeq Market — міжнародний онлайн-магазин для геномних та біоінформатичних лабораторій: набори для підготовки NGS-бібліотек, реагенти, олігонуклеотиди, послуги секвенування, ліцензії на біоінформатичне ПЗ, хмарні обчислення та навчання."),
 ("Датасет", f"Історія продажів за період {START:%d.%m.%Y} – {END:%d.%m.%Y} (30 місяців). Сирий файл: {raw_n} рядків (аркуш Data_raw); очищений: {clean_n} замовлень від {len(cust)} клієнтів із {df.country.nunique()} країн (аркуш Data). Валюта — USD."),
 ("Походження даних", "Датасет синтетичний, згенерований у Python (numpy/pandas, seed 42) за реалістичними правилами B2B-продажів лабораторних товарів: сезонність бюджетних циклів, робочі дні, зростання ~22 %/рік, продуктові вподобання сегментів, плюс навмисно закладені бізнес-аномалії та помилки у даних, які треба знайти в ході EDA."),
 ("Інструменти", "Python (pandas, numpy) для генерації та розрахунків; Excel — таблиці, вбудовані діаграми, умовне форматування (heatmap) та формули описової статистики (аркуш «Статистика (формули)»)."),
 ("Питання EDA", "1) Яка сезонність попиту (річна, тижнева, внутрішньомісячна, за категоріями)? 2) Які аномалії є у даних і чим вони пояснюються? 3) На які групи (когорти) можна розбити клієнтів і чим вони відрізняються? 4) Які міри описової статистики характеризують замовлення загалом і по когортах?"),
 ("Структура файлу", "Опис → Data_raw (сирі дані) → Data (очищені дані + похідні поля) → Клієнти → Сезонність → Аномалії → Когорти → Статистика (формули) → Статистика (значення) → Висновки."),
]
r = 5
for k, v in rows:
    ws.cell(row=r, column=1, value=k).font = Font(bold=True); c = ws.cell(row=r, column=2, value=v); c.alignment = Alignment(wrap_text=True, vertical="top"); r += 1
section(ws, f"A{r+1}", "Словник полів (аркуш Data)")
fields = [("order_id", "ідентифікатор замовлення"), ("order_date", "дата замовлення"), ("year / month / month_name / quarter / year_month", "похідні поля дати"),
          ("weekday / is_weekend", "день тижня (Пн–Нд) та ознака вихідного"), ("customer_id", "ідентифікатор клієнта"),
          ("segment", "сегмент клієнта: Academic (університети, інститути), Biotech (біотех-компанії), Pharma (фармкомпанії), Clinical (клінічні лабораторії), CRO (контрактні дослідницькі організації)"),
          ("country / region", "країна та регіон клієнта"), ("channel", "канал продажу: Web, Sales rep, Distributor"),
          ("category / product", "категорія та найменування товару"), ("quantity", "кількість одиниць"), ("unit_price", "ціна за одиницю, USD"),
          ("discount_pct", "знижка (частка)"), ("revenue", "дохід замовлення = quantity × unit_price × (1 − discount_pct), USD"),
          ("first_order_month / cohort_quarter", "місяць / квартал першого замовлення клієнта (когорта залучення)"),
          ("customer_type", "Новий (перше замовлення) / Повторний"), ("rfm_tier", "RFM-група клієнта (A – ключові, B – активні, C – разові / сплячі)")]
header(ws, r + 2, 1, ["Поле", "Опис"])
for i, (k, v) in enumerate(fields):
    ws.cell(row=r + 3 + i, column=1, value=k).border = BORDER; c = ws.cell(row=r + 3 + i, column=2, value=v); c.border = BORDER; c.alignment = Alignment(wrap_text=True)
widths(ws, [42, 120])

# ---------------------------------------------------------------- Data_raw
ws = wb.create_sheet("Data_raw")
raw_cols = ["order_id", "order_date", "customer_id", "segment", "country", "region", "channel", "category", "product", "quantity", "unit_price", "discount_pct", "revenue"]
write_df(ws, raw[raw_cols], 1, 1, {"order_date": "yyyy-mm-dd", "unit_price": "#,##0.00", "revenue": "#,##0.00", "discount_pct": "0%"})
ws.freeze_panes = "A2"; ws.auto_filter.ref = f"A1:{get_column_letter(len(raw_cols))}{len(raw)+1}"
widths(ws, [10, 12, 11, 10, 15, 18, 11, 22, 38, 9, 11, 10, 12])

# ---------------------------------------------------------------- Data
ws = wb.create_sheet("Data")
data_cols = ["order_id", "order_date", "year", "month", "month_name", "quarter", "year_month", "weekday", "is_weekend", "customer_id", "segment", "country", "region",
             "channel", "category", "product", "quantity", "unit_price", "discount_pct", "revenue", "first_order_month", "cohort_quarter", "customer_type", "rfm_tier"]
write_df(ws, df[data_cols], 1, 1, {"order_date": "yyyy-mm-dd", "unit_price": "#,##0.00", "revenue": "#,##0.00", "discount_pct": "0%"})
N = len(df) + 1
ws.freeze_panes = "C2"; ws.auto_filter.ref = f"A1:{get_column_letter(len(data_cols))}{N}"
widths(ws, [10, 12, 6, 6, 8, 6, 9, 6, 8, 11, 10, 15, 18, 11, 22, 38, 9, 11, 9, 12, 11, 10, 11, 20])
COL = {c: get_column_letter(i + 1) for i, c in enumerate(data_cols)}
def rngc(c): return f"Data!${COL[c]}$2:${COL[c]}${N}"

# ---------------------------------------------------------------- Клієнти
ws = wb.create_sheet("Клієнти")
title(ws, "Клієнти: агрегація на рівні клієнта та RFM-групи",
      f"Recency — днів від останнього замовлення до {END:%d.%m.%Y}; Frequency — кількість замовлень; Monetary — сумарний дохід. Кожна метрика оцінена 1–3 за терцилями, сума 3–9 → група.")
ccols = ["customer_id", "segment", "country", "region", "first_order", "last_order", "orders", "revenue", "quantity", "avg_order_value", "recency_days", "R", "F", "M", "rfm_score", "rfm_tier", "repeat"]
cs = cust[ccols].sort_values("revenue", ascending=False)
r0, r1 = write_df(ws, cs, 5, 1, {"first_order": "yyyy-mm-dd", "last_order": "yyyy-mm-dd", "revenue": "#,##0", "avg_order_value": "#,##0"})
ws.freeze_panes = "B6"; ws.auto_filter.ref = f"A5:Q{r1}"
widths(ws, [11, 10, 15, 18, 11, 11, 8, 11, 9, 11, 9, 4, 4, 4, 6, 20, 10])
# scatter: orders vs revenue (log-ish visual by nature)
ch = ScatterChart(); ch.title = "Клієнти: кількість замовлень vs сумарний дохід"; ch.style = 13; ch.width, ch.height = 18, 10
s = Series(Reference(ws, min_col=8, min_row=r0, max_row=r1), Reference(ws, min_col=7, min_row=r0, max_row=r1), title="Клієнт")
s.marker = Marker(symbol="circle", size=5); s.marker.graphicalProperties = GraphicalProperties(solidFill=BLUE); s.marker.graphicalProperties.line.noFill = True
s.graphicalProperties.line.noFill = True; ch.series.append(s)
style_axes(ch, "Дохід клієнта, USD", "Кількість замовлень"); ch.legend = None; ch.y_axis.scaling.logBase = 10
ws.add_chart(ch, "S5")
note(ws, "S27", "Вісь Y логарифмічна: дохід клієнтів відрізняється на 3 порядки. Точки праворуч-угорі — ключові клієнти (Pharma/CRO).")

# ---------------------------------------------------------------- Сезонність
ws = wb.create_sheet("Сезонність")
title(ws, "Сезонність: динаміка за місяцями, індекс сезонності, день тижня, день місяця, категорії")
section(ws, "A4", "1. Динаміка за місяцями")
mt = monthly.rename(columns={"year_month": "Місяць", "orders": "Замовлень", "revenue": "Дохід, USD", "customers": "Активних клієнтів", "quantity": "Одиниць", "aov": "Середній чек", "ma3": "Ковзне середнє 3 міс", "mom_pct": "Зміна до попер. міс, %"})
r0, r1 = write_df(ws, mt, 5, 1, {"Дохід, USD": "#,##0", "Середній чек": "#,##0", "Ковзне середнє 3 міс": "#,##0"})
cats = Reference(ws, min_col=1, min_row=r0, max_row=r1)
line_chart(ws, "J5", "Дохід за місяцями, USD (лінія — ковзне середнє 3 міс.)", ws, cats,
           [(Reference(ws, min_col=3, min_row=r0 - 1, max_row=r1), "Дохід"), (Reference(ws, min_col=7, min_row=r0 - 1, max_row=r1), "Ковзне середнє 3 міс")], [BLUE, MUTED], ytitle="USD", dashes=[None, "dash"])
line_chart(ws, "J24", "Кількість замовлень за місяцями", ws, cats, [(Reference(ws, min_col=2, min_row=r0 - 1, max_row=r1), "Замовлень")], [ORANGE], ytitle="замовлень", legend=False)
line_chart(ws, "J43", "Активні клієнти за місяцями", ws, cats, [(Reference(ws, min_col=4, min_row=r0 - 1, max_row=r1), "Клієнтів")], [AQUA], ytitle="клієнтів", legend=False)

r = r1 + 3
section(ws, f"A{r}", "2. Теплова карта: дохід за роками та місяцями (USD) та індекс сезонності")
note(ws, f"A{r+1}", f"Індекс сезонності — метод відношення до центрованого 12-місячного ковзного середнього (ratio-to-moving-average), усереднений за роками; тендерне замовлення {bulk_id} (14.11.2024) і додаткові замовлення промо-акції Lab Week (25–29.11.2024) виключені, щоб не спотворювати листопад.")
hm = ym_all.copy(); hm.index = [f"{m:02d} {UA_M[m-1]}" for m in hm.index]; hm.columns = [str(c) for c in hm.columns]
hm["Індекс сезонності"] = seas_idx.values
hm[["2023", "2024", "2025"]] = hm[["2023", "2024", "2025"]].round(0)
hr0, hr1 = write_df(ws, hm, r + 2, 1, {"2023": "#,##0", "2024": "#,##0", "2025": "#,##0", "Індекс сезонності": "0.00"}, index=True, index_name="Місяць")
heat(ws, f"B{hr0}:D{hr1}")
heat(ws, f"E{hr0}:E{hr1}", "F5E6C8", "EDA100")
bar_chart(ws, f"J{r+2}", "Індекс сезонності за місяцями (1,0 = середній місяць)", Reference(ws, min_col=1, min_row=hr0, max_row=hr1),
          [(Reference(ws, min_col=5, min_row=hr0 - 1, max_row=hr1), "Індекс")], [YELLOW], w=22, h=8, ytitle="індекс", legend=False, yfmt="0.00")

r = hr1 + 3
section(ws, f"A{r}", "3. День тижня")
wt = wd.reset_index().rename(columns={"index": "День", "orders": "Замовлень", "revenue": "Дохід, USD", "orders_per_day": "Замовлень на день", "revenue_per_day": "Дохід на день, USD"})
w0, w1 = write_df(ws, wt, r + 1, 1, {"Дохід, USD": "#,##0", "Дохід на день, USD": "#,##0"})
bar_chart(ws, f"J{r+1}", "Середня кількість замовлень на день тижня", Reference(ws, min_col=1, min_row=w0, max_row=w1),
          [(Reference(ws, min_col=4, min_row=w0 - 1, max_row=w1), "Замовлень/день")], [BLUE], w=14, h=7, ytitle="замовлень", legend=False)
r = w1 + 3
section(ws, f"A{r}", "4. Квартали")
qt = qtr.copy(); qt["Квартал"] = qt.year.astype(str) + " " + qt.quarter; qt = qt[["Квартал", "orders", "revenue"]].rename(columns={"orders": "Замовлень", "revenue": "Дохід, USD"})
q0, q1r = write_df(ws, qt, r + 1, 1, {"Дохід, USD": "#,##0"})
bar_chart(ws, f"J{r+1}", "Дохід за кварталами, USD", Reference(ws, min_col=1, min_row=q0, max_row=q1r),
          [(Reference(ws, min_col=3, min_row=q0 - 1, max_row=q1r), "Дохід")], [BLUE], w=14, h=7, ytitle="USD", legend=False)
r = q1r + 3
section(ws, f"A{r}", "5. День місяця (середній дохід на календарний день, без тендерного замовлення та промо-акції)")
dt_ = dom.reset_index(); dt_.columns = ["День місяця", "Середній дохід, USD"]
d0, d1 = write_df(ws, dt_, r + 1, 1, {"Середній дохід, USD": "#,##0"})
bar_chart(ws, f"J{r+1}", "Середній дохід за днем місяця, USD", Reference(ws, min_col=1, min_row=d0, max_row=d1),
          [(Reference(ws, min_col=2, min_row=d0 - 1, max_row=d1), "Дохід")], [AQUA], w=22, h=7, ytitle="USD", legend=False)
r = d1 + 3
section(ws, f"A{r}", "6. Індекс сезонності за категоріями (дохід місяця / середній місяць категорії, 2023–2025)")
cm = cat_month_idx.copy(); cm.columns = [UA_M[m - 1] for m in cm.columns]
c0, c1 = write_df(ws, cm, r + 1, 1, {c: "0.00" for c in cm.columns}, index=True, index_name="Категорія")
heat(ws, f"B{c0}:M{c1}")
r = c1 + 3
section(ws, f"A{r}", "7. Дохід за категоріями та роками, USD")
cy = cat_year.copy(); cy.columns = [str(c) for c in cy.columns]
y0, y1 = write_df(ws, cy, r + 1, 1, {c: "#,##0" for c in cy.columns}, index=True, index_name="Категорія")
bar_chart(ws, f"J{r+1}", "Дохід за категоріями та роками, USD (2025 — лише H1)", Reference(ws, min_col=1, min_row=y0, max_row=y1),
          [(Reference(ws, min_col=2 + i, min_row=y0 - 1, max_row=y1), str(c)) for i, c in enumerate(cy.columns)], [BLUE, ORANGE, AQUA], w=22, h=9, ytitle="USD", horizontal=True)
widths(ws, [30, 12, 12, 12, 12, 12, 14, 14, 12] + [12] * 12)

# ---------------------------------------------------------------- Аномалії
ws = wb.create_sheet("Аномалії")
title(ws, "Аномалії: помилки у даних, викиди на рівні замовлень, аномальні дні",
      "Метод для днів: робастний z-score на логарифмі доходу = 0,6745·(ln x − ковзна медіана за 41 роб. день) / MAD; аномалія при |z| ≥ 3. Логарифм вирівнює правосторонню асиметрію, тому падіння і сплески оцінюються симетрично. Викиди замовлень — правило Q3 + 3·IQR.")
section(ws, "A5", "1. Аномалії якості даних (Data_raw → Data)")
dq = pd.DataFrame([
    ("Повні дублікати рядків", n_dup, "Видалено (залишено один запис)"),
    ("Тестові замовлення з unit_price = 0 (order_id T…)", n_zero, "Видалено — не є продажами"),
    ("Повернення з від'ємною кількістю (order_id R…)", n_ret, "Виключено з аналізу продажів (аналізуються окремо як повернення)"),
    ("Пропущена країна (country порожнє)", n_miss, "Відновлено за customer_id (кожен клієнт має одну країну)"),
    ("Помилкова дата (рік поза періодом, 2032–2034)", n_typo, "Видалено — неможливо однозначно відновити"),
    ("Разом рядків у сирому файлі / в очищеному", f"{raw_n} / {clean_n}", ""),
], columns=["Тип аномалії", "Кількість", "Дія"])
write_df(ws, dq, 6, 1)
section(ws, "A15", f"2. Викиди на рівні замовлень (revenue > Q3 + 3·IQR = {out_thr:,.0f} USD): {len(outliers)} замовлень ({len(outliers)/len(df)*100:.1f} %)")
note(ws, "A16", "Показані топ-15. Майже всі — великі закупівлі Pharma/CRO/Clinical: це очікувані «важкі хвости» B2B, а не помилки. Виняток — тендерне замовлення 14.11.2024, яке на порядок більше за решту.")
ot = outliers.head(15)[["order_id", "order_date", "customer_id", "segment", "region", "category", "product", "quantity", "unit_price", "discount_pct", "revenue"]]
write_df(ws, ot, 17, 1, {"order_date": "yyyy-mm-dd", "unit_price": "#,##0.00", "revenue": "#,##0", "discount_pct": "0%"})
section(ws, "A35", f"3. Аномальні дні (робочі дні з |z| ≥ 3): {len(flag)} днів")
ft = flag.reset_index()[["date", "orders", "revenue", "expected", "z", "comment"]].rename(columns={"date": "Дата", "orders": "Замовлень", "revenue": "Дохід, USD", "expected": "Очікувано (медіана), USD", "z": "z-score", "comment": "Пояснення"})
f0, f1 = write_df(ws, ft, 36, 1, {"Дата": "yyyy-mm-dd", "Дохід, USD": "#,##0", "Очікувано (медіана), USD": "#,##0", "z-score": "0.00"})
section(ws, f"A{f1+3}", "4. Дефіцит реагентів: категорія Reagents & Enzymes за місяцями")
rg = reag.rename(columns={"year_month": "Місяць", "orders": "Замовлень", "revenue": "Дохід, USD"})
g0, g1 = write_df(ws, rg, f1 + 4, 1, {"Дохід, USD": "#,##0"})
bar_chart(ws, f"E{f1+4}", "Reagents & Enzymes: замовлень за місяцями (провал у квітні 2024 — stock-out)", Reference(ws, min_col=1, min_row=g0, max_row=g1),
          [(Reference(ws, min_col=2, min_row=g0 - 1, max_row=g1), "Замовлень")], [ORANGE], w=22, h=8, ytitle="замовлень", legend=False)
# daily series for chart (business days) in far columns
section(ws, "N5", "5. Щоденний дохід (робочі дні) з очікуваним рівнем і межами аномалії")
dd = bd.reset_index()[["date", "orders", "revenue", "expected", "upper", "lower", "z"]].rename(columns={"date": "Дата", "orders": "Замовлень", "revenue": "Дохід, USD", "expected": "Очікувано", "upper": "Верхня межа", "lower": "Нижня межа", "z": "z-score"})
b0, b1 = write_df(ws, dd, 6, 14, {"Дата": "yyyy-mm-dd", "Дохід, USD": "#,##0", "Очікувано": "#,##0", "Верхня межа": "#,##0", "Нижня межа": "#,##0", "z-score": "0.00"})
ch = line_chart(ws, "A62", "Щоденний дохід (робочі дні): факт, ковзна медіана та межі аномалії", ws, Reference(ws, min_col=14, min_row=b0, max_row=b1),
                [(Reference(ws, min_col=16, min_row=b0 - 1, max_row=b1), "Дохід"), (Reference(ws, min_col=17, min_row=b0 - 1, max_row=b1), "Очікувано (медіана)"),
                 (Reference(ws, min_col=18, min_row=b0 - 1, max_row=b1), "Верхня межа"), (Reference(ws, min_col=19, min_row=b0 - 1, max_row=b1), "Нижня межа")],
                [BLUE, INK2, RED, RED], w=30, h=11, ytitle="USD", dashes=[None, "dash", "sysDot", "sysDot"])
ch.x_axis.number_format = "mmm yy"; ch.x_axis.majorTimeUnit = "months"
ch.y_axis.scaling.max = 60000
note(ws, "A85", "Вісь Y обмежена 60 000 USD, щоб було видно звичайну динаміку; тендерне замовлення 14.11.2024 (~$192 тис.) виходить за межі графіка.")
widths(ws, [46, 12, 11, 10, 10, 20, 22, 36, 9, 11, 10, 12, 3, 12, 10, 12, 12, 12, 12, 9])
ws.column_dimensions["F"].width = 70

# ---------------------------------------------------------------- Когорти
ws = wb.create_sheet("Когорти")
title(ws, "Когорти: сегменти, регіони, RFM-групи, разові vs повторні, когорти залучення")
section(ws, "A4", "1. Сегменти клієнтів")
st = seg_tbl.rename(columns={"segment": "Сегмент", "customers": "Клієнтів", "orders": "Замовлень", "revenue": "Дохід, USD", "quantity": "Одиниць", "discount": "Середня знижка", "share_pct": "Частка доходу, %", "aov": "Середній чек, USD", "orders_per_customer": "Замовлень на клієнта", "revenue_per_customer": "Дохід на клієнта, USD"})
st = st[["Сегмент", "Клієнтів", "Замовлень", "Дохід, USD", "Частка доходу, %", "Середній чек, USD", "Замовлень на клієнта", "Дохід на клієнта, USD", "Середня знижка"]]
s0, s1 = write_df(ws, st, 5, 1, {"Дохід, USD": "#,##0", "Середній чек, USD": "#,##0", "Дохід на клієнта, USD": "#,##0", "Середня знижка": "0.0%"})
bar_chart(ws, "L5", "Дохід за сегментами, USD", Reference(ws, min_col=1, min_row=s0, max_row=s1), [(Reference(ws, min_col=4, min_row=s0 - 1, max_row=s1), "Дохід")], [BLUE], w=13, h=7, legend=False)
bar_chart(ws, "T5", "Середній чек за сегментами, USD", Reference(ws, min_col=1, min_row=s0, max_row=s1), [(Reference(ws, min_col=6, min_row=s0 - 1, max_row=s1), "Середній чек")], [ORANGE], w=13, h=7, legend=False)
r = s1 + 3
section(ws, f"A{r}", "2. Регіони")
rt = reg_tbl.rename(columns={"region": "Регіон", "customers": "Клієнтів", "orders": "Замовлень", "revenue": "Дохід, USD", "share_pct": "Частка доходу, %", "aov": "Середній чек, USD", "orders_per_customer": "Замовлень на клієнта"})
g0, g1 = write_df(ws, rt, r + 1, 1, {"Дохід, USD": "#,##0", "Середній чек, USD": "#,##0"})
bar_chart(ws, f"L{r+1}", "Дохід за регіонами, USD", Reference(ws, min_col=1, min_row=g0, max_row=g1), [(Reference(ws, min_col=4, min_row=g0 - 1, max_row=g1), "Дохід")], [BLUE], w=13, h=7, legend=False)
ct = country_tbl.rename(columns={"country": "Країна", "customers": "Клієнтів", "orders": "Замовлень", "revenue": "Дохід, USD"})
k0, k1 = write_df(ws, ct, r + 1, 10 - 1, {"Дохід, USD": "#,##0"}) if False else (None, None)
r = g1 + 3
section(ws, f"A{r}", "3. Топ-12 країн за доходом")
k0, k1 = write_df(ws, ct, r + 1, 1, {"Дохід, USD": "#,##0"})
bar_chart(ws, f"L{r+1}", "Топ-12 країн за доходом, USD", Reference(ws, min_col=1, min_row=k0, max_row=k1), [(Reference(ws, min_col=4, min_row=k0 - 1, max_row=k1), "Дохід")], [AQUA], w=13, h=9, legend=False, horizontal=True)
r = k1 + 3
section(ws, f"A{r}", "4. RFM-групи клієнтів (розрахунок на аркуші «Клієнти»)")
ft = rfm_tbl.rename(columns={"rfm_tier": "RFM-група", "customers": "Клієнтів", "orders": "Замовлень", "revenue": "Дохід, USD", "avg_recency": "Середня давність, днів", "avg_orders": "Середньо замовлень", "share_pct": "Частка доходу, %"})
f0, f1 = write_df(ws, ft, r + 1, 1, {"Дохід, USD": "#,##0", "Середня давність, днів": "0", "Середньо замовлень": "0.0"})
bar_chart(ws, f"L{r+1}", "Частка доходу за RFM-групами, %", Reference(ws, min_col=1, min_row=f0, max_row=f1), [(Reference(ws, min_col=7, min_row=f0 - 1, max_row=f1), "Частка, %")], [VIOLET], w=13, h=7, legend=False, yfmt="0")
r = f1 + 3
section(ws, f"A{r}", "5. Разові vs повторні клієнти")
pt = rep_tbl.rename(columns={"repeat": "Тип клієнта", "customers": "Клієнтів", "revenue": "Дохід, USD", "share_pct": "Частка доходу, %"})
p0, p1 = write_df(ws, pt, r + 1, 1, {"Дохід, USD": "#,##0"})
r = p1 + 3
section(ws, f"A{r}", "6. Канали продажу")
cht = chan_tbl.rename(columns={"channel": "Канал", "orders": "Замовлень", "revenue": "Дохід, USD", "share_pct": "Частка доходу, %"})
h0, h1 = write_df(ws, cht, r + 1, 1, {"Дохід, USD": "#,##0"})
r = h1 + 3
section(ws, f"A{r}", "7. Структура доходу сегмента за категоріями, % (heatmap)")
sc = seg_cat_pct.copy()
sc0, sc1 = write_df(ws, sc, r + 1, 1, {c: "0.0" for c in sc.columns}, index=True, index_name="Сегмент \\ Категорія")
heat(ws, f"B{sc0}:{get_column_letter(1+len(sc.columns))}{sc1}")
r = sc1 + 3
section(ws, f"A{r}", "8. Когорти залучення (квартал першого замовлення): утримання, % клієнтів когорти, що зробили замовлення у k-му кварталі після першого")
note(ws, f"A{r+1}", "Порожні клітинки — квартал ще не настав. k = 0 — квартал залучення (завжди 100 %). Когорти 2023Q1–Q2 більші за інші, бо включають базу клієнтів, що існувала до початку спостережень (ліва цензура).")
rp = ret_pct.copy(); rp.columns = [f"k={k}" for k in rp.columns]
rp.insert(0, "Дохід когорти, USD", cohort_rev.reindex(rp.index).round(0).values)
rp.insert(0, "Клієнтів", cohort_size.reindex(rp.index).values)
c0, c1 = write_df(ws, rp, r + 2, 1, {"Дохід когорти, USD": "#,##0", **{c: "0.0" for c in rp.columns if c.startswith("k=")}}, index=True, index_name="Когорта")
heat(ws, f"D{c0}:{get_column_letter(3+len(ret_pct.columns))}{c1}")
# retention curve chart: mean retention per k for cohorts with data
rc = ret_pct.mean(axis=0).round(1).reset_index(); rc.columns = ["Кварталів після залучення", "Середнє утримання, %"]
rr0, rr1 = write_df(ws, rc, c1 + 3, 1, {"Середнє утримання, %": "0.0"})
line_chart(ws, f"L{c1+3}", "Середня крива утримання когорт, % клієнтів", ws, Reference(ws, min_col=1, min_row=rr0, max_row=rr1),
           [(Reference(ws, min_col=2, min_row=rr0 - 1, max_row=rr1), "Утримання, %")], [BLUE], w=16, h=8, ytitle="%", legend=False)
widths(ws, [24, 11, 11, 14, 12, 14, 14, 16, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12])

# ---------------------------------------------------------------- Статистика (формули)
ws = wb.create_sheet("Статистика (формули)")
title(ws, "Описова статистика — живі формули Excel над аркушем Data (revenue та quantity)",
      "Формули з IF(...) записані як формули масиву (CSE) і обчислюються в Excel та Google Sheets при відкритті (Apple Numbers формули масиву не підтримує — там ці клітинки показують 0). Загальний стовпчик — звичайні функції. Готові числа для звірки — на аркуші «Статистика (значення)».")
def stats_block(ws, top, field, label, groups_spec):
    """groups_spec: list of (header, key_col or None, key_val or None)"""
    section(ws, f"A{top}", label)
    header(ws, top + 1, 1, ["Показник"] + [g[0] for g in groups_spec])
    V = rngc(field)
    for i, name in enumerate(STAT_ROWS):
        ws.cell(row=top + 2 + i, column=1, value=name).border = BORDER
    for j, (gname, kcol, kval) in enumerate(groups_spec):
        col = get_column_letter(2 + j)
        if kcol is None:
            f = {"Кількість (n)": f"=COUNT({V})", "Сума": f"=SUM({V})", "Середнє": f"=AVERAGE({V})", "Медіана": f"=MEDIAN({V})", "Мода": f"=MODE({V})",
                 "Мінімум": f"=MIN({V})", "Максимум": f"=MAX({V})", "Розмах": f"=MAX({V})-MIN({V})", "Дисперсія": f"=VAR({V})", "Стандартне відхилення": f"=STDEV({V})",
                 "Коеф. варіації, %": f"=STDEV({V})/AVERAGE({V})*100", "Q1 (25%)": f"=QUARTILE({V},1)", "Q3 (75%)": f"=QUARTILE({V},3)",
                 "IQR": f"=QUARTILE({V},3)-QUARTILE({V},1)", "Асиметрія (skewness)": f"=SKEW({V})", "Ексцес (kurtosis)": f"=KURT({V})"}
            arr = set()
        else:
            K = rngc(kcol); cond = f'{K}="{kval}"'; sel = f"IF({cond},{V})"
            f = {"Кількість (n)": f'=COUNTIF({K},"{kval}")', "Сума": f'=SUMIF({K},"{kval}",{V})', "Середнє": f'=AVERAGEIF({K},"{kval}",{V})',
                 "Медіана": f"=MEDIAN({sel})", "Мода": f"=MODE({sel})", "Мінімум": f"=MIN({sel})", "Максимум": f"=MAX({sel})", "Розмах": f"=MAX({sel})-MIN({sel})",
                 "Дисперсія": f"=VAR({sel})", "Стандартне відхилення": f"=STDEV({sel})", "Коеф. варіації, %": f'=STDEV({sel})/AVERAGEIF({K},"{kval}",{V})*100',
                 "Q1 (25%)": f"=QUARTILE({sel},1)", "Q3 (75%)": f"=QUARTILE({sel},3)", "IQR": f"=QUARTILE({sel},3)-QUARTILE({sel},1)",
                 "Асиметрія (skewness)": f"=SKEW({sel})", "Ексцес (kurtosis)": f"=KURT({sel})"}
            arr = {k for k in f if "IF(" in f[k] and not f[k].startswith(("=COUNTIF", "=SUMIF", "=AVERAGEIF"))}
        for i, name in enumerate(STAT_ROWS):
            ref = f"{col}{top+2+i}"
            ws[ref] = ArrayFormula(ref, f[name]) if name in arr else f[name]
            ws[ref].number_format = "0.000" if name in ("Асиметрія (skewness)", "Ексцес (kurtosis)") else ("0.0" if name.startswith("Коеф") else "#,##0.00")
            ws[ref].border = BORDER
    return top + 2 + len(STAT_ROWS)
spec_seg = [("Весь датасет", None, None)] + [(s, "segment", s) for s in seg_tbl.segment]
spec_reg = [(r_, "region", r_) for r_ in reg_tbl.region]
spec_rfm = [(t, "rfm_tier", t) for t in sorted(cust.rfm_tier.unique())]
e = stats_block(ws, 5, "revenue", "A. Дохід замовлення (revenue, USD): загалом та за сегментами", spec_seg)
e = stats_block(ws, e + 2, "revenue", "B. Дохід замовлення (revenue, USD): за регіонами", spec_reg)
e = stats_block(ws, e + 2, "revenue", "C. Дохід замовлення (revenue, USD): за RFM-групами", spec_rfm)
e = stats_block(ws, e + 2, "quantity", "D. Кількість одиниць у замовленні (quantity): загалом та за сегментами", spec_seg)
widths(ws, [26] + [20] * 8)

# ---------------------------------------------------------------- Статистика (значення)
ws = wb.create_sheet("Статистика (значення)")
title(ws, "Описова статистика — значення, розраховані в Python (pandas), для перевірки формул",
      "Дисперсія та стандартне відхилення — вибіркові (n−1), як VAR/STDEV в Excel; асиметрія та ексцес — за формулами Excel SKEW/KURT.")
section(ws, "A5", "A. Дохід замовлення (revenue, USD) — загалом та по когортах")
a0, a1 = write_df(ws, stats_rev, 6, 1, {c: "#,##0.00" for c in stats_rev.columns}, index=True, index_name="Показник")
r = a1 + 3
section(ws, f"A{r}", "B. Кількість одиниць у замовленні (quantity) — загалом та по когортах")
b0, b1 = write_df(ws, stats_qty, r + 1, 1, {c: "#,##0.00" for c in stats_qty.columns}, index=True, index_name="Показник")
r = b1 + 3
section(ws, f"A{r}", "C. Показники на рівні клієнта")
c0, c1 = write_df(ws, stats_cust, r + 1, 1, {c: "#,##0.00" for c in stats_cust.columns}, index=True, index_name="Показник")
r = c1 + 3
section(ws, f"A{r}", "D. Гістограма: розподіл доходу замовлення, USD")
ht = hist.reset_index(); ht.columns = ["Інтервал доходу, USD", "Кількість замовлень"]; ht["Частка, %"] = (ht["Кількість замовлень"] / len(df) * 100).round(1)
h0, h1 = write_df(ws, ht, r + 1, 1)
bar_chart(ws, f"F{r+1}", "Гістограма доходу замовлення (логарифмічні інтервали)", Reference(ws, min_col=1, min_row=h0, max_row=h1),
          [(Reference(ws, min_col=2, min_row=h0 - 1, max_row=h1), "Замовлень")], [BLUE], w=18, h=8, ytitle="замовлень", legend=False)
r = h1 + 3
section(ws, f"A{r}", "E. Box plot за сегментами (п'ятичислова характеристика доходу замовлення, USD)")
note(ws, f"A{r+1}", "Діаграма побудована як стекова гістограма: невидимий сегмент до Q1, потім Q1→медіана та медіана→Q3 («ящик»). Мінімум і максимум («вуса») — у таблиці; вісь Y логарифмічна.")
bx = box.copy(); bx.columns = ["Мінімум", "Q1", "Медіана", "Q3", "Максимум"]
bx["_Q1"] = bx["Q1"]; bx["_Q2-Q1"] = bx["Медіана"] - bx["Q1"]; bx["_Q3-Q2"] = bx["Q3"] - bx["Медіана"]; bx["IQR"] = bx["Q3"] - bx["Q1"]
bx = bx[["Мінімум", "Q1", "Медіана", "Q3", "Максимум", "IQR", "_Q1", "_Q2-Q1", "_Q3-Q2"]].round(2)
x0, x1 = write_df(ws, bx, r + 2, 1, {c: "#,##0" for c in bx.columns}, index=True, index_name="Сегмент")
note(ws, f"H{x1+1}", "← допоміжні стовпчики для діаграми")
ch = bar_chart(ws, f"F{r+2}", "Box plot: дохід замовлення за сегментами (Q1 – медіана – Q3)", Reference(ws, min_col=1, min_row=x0, max_row=x1),
               [(Reference(ws, min_col=8, min_row=x0 - 1, max_row=x1), "до Q1"), (Reference(ws, min_col=9, min_row=x0 - 1, max_row=x1), "Q1 → медіана"),
                (Reference(ws, min_col=10, min_row=x0 - 1, max_row=x1), "медіана → Q3")], [BLUE, "86B6EF", "1C5CAB"], w=18, h=9, stacked=True, ytitle="USD")
ch.series[0].graphicalProperties.noFill = True; ch.series[0].graphicalProperties.solidFill = None
ch.gapWidth = 120
widths(ws, [26] + [16] * 14)

# ---------------------------------------------------------------- Висновки
ws = wb.create_sheet("Висновки")
title(ws, "Висновки експлоративного аналізу")
top3 = seas_idx.sort_values(ascending=False).head(3); low3 = seas_idx.sort_values().head(3)
UA_MF = {1: "січень", 2: "лютий", 3: "березень", 4: "квітень", 5: "травень", 6: "червень", 7: "липень", 8: "серпень", 9: "вересень", 10: "жовтень", 11: "листопад", 12: "грудень"}
best_wd = wd.orders_per_day.idxmax(); wknd_share = wd.loc[["Сб", "Нд"], "orders"].sum() / wd.orders.sum() * 100
me_share = df_nb[df_nb.order_date.dt.day >= 28].revenue.sum() / df_nb.revenue.sum() * 100
sw = cat_month_idx.loc["Bioinformatics Software"]
yoy = monthly[monthly.year_month.str.startswith("2024")].revenue.sum() / monthly[monthly.year_month.str.startswith("2023")].revenue.sum() * 100 - 100
h1 = df[(df.year == 2025)].revenue.sum() / df[(df.year == 2024) & (df.month <= 6)].revenue.sum() * 100 - 100
new_cust = cust.groupby(cust.first_order.dt.year).customer_id.count()
segA = seg_tbl.iloc[0]; segAc = seg_tbl[seg_tbl.segment == "Academic"].iloc[0]; segPh = seg_tbl[seg_tbl.segment == "Pharma"].iloc[0]
rfmA = rfm_tbl[rfm_tbl.rfm_tier.str.startswith("A")].iloc[0]; rfmC = rfm_tbl[rfm_tbl.rfm_tier.str.startswith("C")].iloc[0]
rep = rep_tbl.set_index("repeat")
ov = stats_rev["Весь датасет"]
k1ret = ret_pct[1].dropna().mean(); k4ret = ret_pct[4].dropna().mean()
concl = [
 ("Сезонність", [
   f"Річна: пік попиту — {', '.join(f'{UA_MF[m]} (індекс {v:.2f})' for m, v in top3.items())}: закінчення федерального фінансового року США (30 вересня), витрачання річних бюджетів лабораторій у Q4 та кінець грантових циклів. Провали — {', '.join(f'{UA_MF[m]} ({v:.2f})' for m, v in low3.items())}: відпустки та святковий період. Розмах індексу від {seas_idx.min():.2f} до {seas_idx.max():.2f} означає, що сезонність змінює місячний дохід приблизно вдвічі.",
   f"Тренд: дохід 2024 року на {yoy:.0f} % вищий за 2023 (зростання бази клієнтів), а перше півріччя 2025 — на {h1:.0f} % порівняно з першим півріччям 2024: зростання {'сповільнилося' if h1 >= 5 else 'практично зупинилося'}. Приплив нових клієнтів стабільний ({new_cust.get(2024, 0)} за 2024, {new_cust.get(2025, 0)} за H1 2025), але відтік із віком бази зростає (див. утримання когорт), тому кількість активних клієнтів за місяць вийшла на плато.",
   f"Тижнева: {wknd_share:.1f} % замовлень припадає на вихідні — типовий B2B-профіль; найактивніший день — {best_wd} ({wd.loc[best_wd,'orders_per_day']:.1f} замовлень/день проти {wd.loc['Нд','orders_per_day']:.1f} у неділю).",
   f"Внутрішньомісячна: останні дні місяця (28–31) дають {me_share:.1f} % доходу — ефект закриття замовлень на придбання (PO) наприкінці місяця/кварталу.",
   f"За категоріями: ліцензії на біоінформатичне ПЗ мають власну сезонність поновлень — січень (індекс {sw[1]:.2f}) і липень ({sw[7]:.2f}); навчання — лютий–березень та вересень–жовтень (початок семестрів). Реагенти й набори — рівномірніші, повторюють загальний бюджетний цикл.",
 ]),
 ("Аномалії", [
   f"Якість даних: у сирому файлі знайдено {n_dup} дублікатів, {n_zero} тестових замовлень з нульовою ціною, {n_ret} повернень із від'ємною кількістю, {n_miss} рядків без країни та {n_typo} рядки з помилковою датою; після очищення — {clean_n} замовлень.",
   f"Викиди на рівні замовлень: {len(outliers)} замовлень ({len(outliers)/len(df)*100:.1f} %) перевищують Q3 + 3·IQR = {out_thr:,.0f} USD; це закупівлі Pharma/CRO/Clinical — природний «важкий хвіст» розподілу (асиметрія {ov['Асиметрія (skewness)']:.1f}, ексцес {ov['Ексцес (kurtosis)']:.0f}), а не помилки. Найбільший викид — тендерне замовлення {bulk_id} від 14.11.2024 на 200 наборів NGS Library Prep (~{df.revenue.max()/1000:,.0f} тис. USD), яке саме формує «аномальний» листопад 2024.",
   f"Аномальні дні (робастний z-score, {len(flag)} днів): позитивні — тендер 14.11.2024 та промо-акція Lab Week 25–29.11.2024 (сплеск дрібних web-замовлень зі знижкою 20 %); негативні — збій ERP/сайту 15–16.08.2023 (нульовий дохід у робочі дні) і святкові дні.",
   f"Аномалія на рівні категорії: дефіцит (stock-out) реагентів з 8 квітня по 5 травня 2024 — кількість замовлень Reagents & Enzymes у квітні 2024 впала приблизно у 10 разів і відновилася у травні; на загальному щоденному доході це малопомітно, тому категорії треба моніторити окремо.",
 ]),
 ("Групи (когорти) користувачів", [
   f"За сегментом: найбільший дохід дає {segA.segment} ({segA.share_pct:.1f} %), хоча клієнтів найбільше в Academic ({segAc.customers} клієнтів, {segAc.share_pct:.1f} % доходу). Середній чек Pharma ({segPh.aov:,.0f} USD) у {segPh.aov/segAc.aov:.1f} раза вищий за Academic ({segAc.aov:,.0f} USD), а частота замовлень — {segPh.orders_per_customer:.1f} проти {segAc.orders_per_customer:.1f} на клієнта. Academic — «довгий хвіст» дрібних замовлень з максимальною часткою олігонуклеотидів і наборів для екстракції; Pharma/CRO — набори для NGS-бібліотек, послуги секвенування та ліцензії ПЗ.",
   f"За регіоном: Europe та North America формують {reg_tbl.share_pct.iloc[:2].sum():.0f} % доходу, Asia-Pacific — {reg_tbl.set_index('region').loc['Asia-Pacific','share_pct']:.0f} %; середній чек між регіонами відрізняється мало (від {reg_tbl.aov.min():,.0f} до {reg_tbl.aov.max():,.0f} USD), тобто регіон впливає на обсяг, а не на структуру замовлення.",
   f"За RFM: група A (ключові, {rfmA.customers} клієнтів, {rfmA.customers/len(cust)*100:.0f} % бази) дає {rfmA.share_pct:.1f} % доходу; група C (разові/сплячі, {rfmC.customers} клієнтів) — лише {rfmC.share_pct:.1f} %. Разові клієнти — {rep.loc['Разовий','customers']} ({rep.loc['Разовий','customers']/len(cust)*100:.0f} % бази) і {rep.loc['Разовий','share_pct']:.1f} % доходу.",
   f"За когортою залучення: у середньому {k1ret:.0f} % клієнтів когорти повертаються у наступному кварталі, і {k4ret:.0f} % — через рік; утримання стабільне між когортами, тобто зростання доходу забезпечене головним чином припливом нових клієнтів.",
 ]),
 ("Описова статистика", [
   f"Дохід замовлення: середнє {ov['Середнє']:,.0f} USD, медіана {ov['Медіана']:,.0f} USD, мода {ov['Мода']:,.0f} USD; стандартне відхилення {ov['Стандартне відхилення']:,.0f} USD (коефіцієнт варіації {ov['Коеф. варіації, %']:.0f} %), IQR {ov['IQR']:,.0f} USD. Середнє суттєво більше за медіану — розподіл сильно скошений вправо, тому для планування й порівняння когорт коректніше використовувати медіану та квартилі.",
   f"По когортах: медіана доходу замовлення Pharma — {stats_rev['Сегмент: Pharma']['Медіана']:,.0f} USD проти {stats_rev['Сегмент: Academic']['Медіана']:,.0f} USD в Academic; коефіцієнт варіації найвищий у Pharma ({stats_rev['Сегмент: Pharma']['Коеф. варіації, %']:.0f} %) через тендерне замовлення. Асиметрія і ексцес додатні у всіх групах — важкі праві хвости є нормою для B2B-продажів лабораторних товарів.",
   f"На рівні клієнта: медіана {stats_cust['Замовлень на клієнта']['Медіана']:.0f} замовлень і {stats_cust['Дохід на клієнта']['Медіана']:,.0f} USD доходу, тоді як середнє — {stats_cust['Замовлень на клієнта']['Середнє']:.1f} замовлень і {stats_cust['Дохід на клієнта']['Середнє']:,.0f} USD: невелика частка великих клієнтів визначає результат.",
 ]),
 ("Рекомендації", [
   f"Планувати запаси й маркетинг під бюджетний цикл: нарощувати склад реагентів/наборів перед піковими місяцями ({', '.join(UA_MF[m] for m in sorted(top3.index))}); кампанії поновлення ліцензій ПЗ запускати у грудні та червні, перед хвилями поновлень у січні та липні.",
   "Виділити окремий сегментний підхід: для Pharma/CRO — менеджерські продажі та тендери, для Academic — web-самообслуговування, промо на початку семестрів і кінці грантового року.",
   "Моніторити аномалії на рівні категорій та каналів (а не лише сумарного доходу), щоб не пропускати stock-out або збої сайту; великі разові замовлення виключати з оцінки сезонності.",
   "Працювати з групою C (разові клієнти) — реактиваційні кампанії; з групою A — програми лояльності, адже вона дає більшу частину доходу.",
 ]),
]
import re
def ua_num(t):
    t = re.sub(r"(?<![\d.,])(\d{1,3}(?:,\d{3})+)(\.(\d+))?(?![\d.])", lambda m: m.group(1).replace(",", "\u202f") + ("," + m.group(3) if m.group(2) else ""), t)
    t = re.sub(r"(?<![\d.])(\d+)\.(\d+)(?![\d.])", r"\1,\2", t)
    return t
concl = [(h, [ua_num(i) for i in items]) for h, items in concl]
r = 4
for head, items in concl:
    section(ws, f"A{r}", head); r += 1
    for it in items:
        c = ws.cell(row=r, column=1, value="• " + it); c.alignment = Alignment(wrap_text=True, vertical="top"); ws.row_dimensions[r].height = 15 * (len(it) // 150 + 1) + 4; r += 1
    r += 1
widths(ws, [150])

wb.save(OUT)
print("saved", OUT)
