# -*- coding: utf-8 -*-
"""Build and execute the two notebooks of the final project:
  Task 4 — funnel visualisation (…_Final Project-4.ipynb)
  Task 5 — confidence interval for the add-to-cart → purchase conversion difference (…_Final Project-5.ipynb)
Notebooks are self-contained (no import of fp_common) so they run in Colab as well."""
import nbformat as nbf
from nbclient import NotebookClient

from fp_common import HERE, SURNAME

COMMON_SETUP = r'''
import warnings
from pathlib import Path

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from scipy import stats

warnings.filterwarnings("ignore")
pd.set_option("display.precision", 4)

C_TEST, C_CONTROL = "#2a78d6", "#eb6834"
C_INK, C_MUTED, C_GRID = "#0b0b0b", "#898781", "#e1e0d9"
plt.rcParams.update({
    "figure.dpi": 110, "axes.spines.top": False, "axes.spines.right": False,
    "axes.edgecolor": "#c3c2b7", "axes.labelcolor": C_INK, "xtick.color": C_MUTED, "ytick.color": C_MUTED,
    "axes.grid": True, "grid.color": C_GRID, "grid.linewidth": 0.8, "axes.axisbelow": True, "font.size": 10,
})

URLS = {
    "Control": "https://drive.google.com/uc?export=download&id=1OlboUFclx-9QAakDNUAMzQquH8tNzo07",
    "Test": "https://drive.google.com/uc?export=download&id=1fFIA1tJj5YQs2G7tv244tgzLhpJ5axoG",
}
FILES = {"Control": "data/control_group.csv", "Test": "data/test_group.csv"}
RENAME = {"Campaign Name": "campaign", "Date": "date", "Spend [USD]": "spend", "# of Impressions": "impressions",
          "Reach": "reach", "# of Website Clicks": "clicks", "# of Searches": "searches",
          "# of View Content": "view_content", "# of Add to Cart": "add_to_cart", "# of Purchase": "purchase"}
STEPS = ["impressions", "reach", "clicks", "searches", "view_content", "add_to_cart", "purchase"]
STEP_UA = {"impressions": "Покази", "reach": "Охоплення", "clicks": "Кліки на сайт", "searches": "Пошук на сайті",
           "view_content": "Перегляд контенту", "add_to_cart": "Додавання в кошик", "purchase": "Покупка"}

def load_group(group):
    local = Path(FILES[group])
    d = pd.read_csv(local if local.exists() else URLS[group], sep=";").rename(columns=RENAME)
    d["date"] = pd.to_datetime(d["date"], format="%d.%m.%Y")
    d["group"] = group
    return d

df = pd.concat([load_group("Control"), load_group("Test")], ignore_index=True)
row_ok = df[STEPS].notna().all(axis=1)
# день використовуємо лише тоді, коли повні дані є в ОБОХ групах — щоб порівнювати однакові періоди
df["complete"] = row_ok.groupby(df["date"]).transform("all")
print(f"Рядків: {len(df)}; рядків з пропусками: {(~row_ok).sum()}; "
      f"днів, виключених з аналізу для обох груп: {df.loc[~df.complete, 'date'].nunique()}")
df[~row_ok]
'''

# ===================================================================== TASK 4
def build_task4():
    cells = []
    md = lambda s: cells.append(nbf.v4.new_markdown_cell(s.strip("\n")))
    code = lambda s: cells.append(nbf.v4.new_code_cell(s.strip("\n")))

    md(r"""
# Фінальний проєкт. Завдання №4 — Візуалізація проходження воронки

**Автор:** Мойсеєнко Світлана · **Дисципліна:** Продуктова аналітика та прикладна статистика (GoIT Neoversity)

**Контекст.** Компанія тестує новий підхід до розміщення онлайн-реклами. Контрольна група (`Control Campaign`) працює як раніше, тестова (`Test Campaign`) — за новим підходом. Для кожної групи є поденні дані за 30 днів серпня 2019 р.: витрати, покази, охоплення та проходження воронки продажів (кліки → пошук → перегляд контенту → кошик → покупка).

**Мета ноутбука:** візуалізувати воронку для тестової та контрольної груп, порівняти конверсії між кроками та зробити висновки.

Дані завантажуються з локальної папки `data/` (копії файлів з умови), а за її відсутності — безпосередньо з Google Drive за посиланнями з завдання.
""")
    code(COMMON_SETUP)
    md(r"""
У контрольній групі за **5 серпня** є лише витрати, а показники воронки відсутні. Щоб порівнювати групи за **однаковий період**, цей день виключаємо з аналізу для обох груп — залишається по 29 днів у кожній. Інакше суми за 30 днів тестової групи порівнювалися б із сумами за 29 днів контрольної, що зміщувало б результат на користь тестової групи.
""")
    code(r"""
d = df[df.complete].copy()
tot = d.groupby("group")[["spend"] + STEPS].sum().T[["Control", "Test"]]
tot.index = ["Витрати, $"] + [STEP_UA[s] for s in STEPS]
tot.astype(int)
""")
    md(r"""
## 1. Воронка: абсолютні значення

Класична воронка «ширина = кількість користувачів на кроці». Покази та охоплення — показники реклами, тому воронку сайту починаємо з кліків: саме з цього моменту користувач потрапляє на сайт.
""")
    code(r"""
funnel_steps = ["clicks", "searches", "view_content", "add_to_cart", "purchase"]
F = d.groupby("group")[funnel_steps].sum().T[["Control", "Test"]]

fig, axes = plt.subplots(1, 2, figsize=(11, 4.2), sharey=True)
for ax, (grp, color) in zip(axes, [("Control", C_CONTROL), ("Test", C_TEST)]):
    vals = F[grp].to_numpy()
    top = F.max().max()
    y = np.arange(len(funnel_steps))[::-1]
    ax.barh(y, vals, left=(top - vals) / 2, height=0.62, color=color)
    for yi, v, s in zip(y, vals, funnel_steps):
        share = v / vals[0]
        label = f"{int(v):,}".replace(",", " ") + f"  ({share:.1%})"
        if v > top * 0.6:   # широкий сегмент — підпис усередині
            ax.text(top / 2, yi, label, ha="center", va="center", color="white", fontsize=9, fontweight="bold")
        else:               # вузький сегмент — підпис праворуч від нього
            ax.text((top + v) / 2 + top * 0.02, yi, label, ha="left", va="center", color=C_INK, fontsize=9, fontweight="bold")
    ax.set_xlim(0, top * 1.45)
    ax.set_yticks(y); ax.set_yticklabels([STEP_UA[s] for s in funnel_steps], color=C_INK)
    ax.set_xticks([]); ax.grid(False); ax.spines["bottom"].set_visible(False)
    ax.set_title(f"{grp} — воронка за 29 днів (частка від кліків)", loc="left", color=C_INK, fontsize=11)
plt.tight_layout(); plt.show()
""")
    md(r"""
## 2. Конверсії між кроками (step-to-step) та від кліку до кожного кроку
""")
    code(r"""
def conv_table(F):
    rows = []
    for i in range(1, len(funnel_steps)):
        prev, cur = funnel_steps[i - 1], funnel_steps[i]
        rows.append({"Крок": f"{STEP_UA[prev]} → {STEP_UA[cur]}",
                     "Control": F.loc[cur, "Control"] / F.loc[prev, "Control"],
                     "Test": F.loc[cur, "Test"] / F.loc[prev, "Test"]})
    rows.append({"Крок": "Клік → Покупка (наскрізна)", "Control": F.loc["purchase", "Control"] / F.loc["clicks", "Control"],
                 "Test": F.loc["purchase", "Test"] / F.loc["clicks", "Test"]})
    out = pd.DataFrame(rows).set_index("Крок")
    out["Різниця, п.п."] = (out["Test"] - out["Control"]) * 100
    return out

conv = conv_table(F)
conv.style.format({"Control": "{:.1%}", "Test": "{:.1%}", "Різниця, п.п.": "{:+.1f}"})
""")
    code(r"""
fig, ax = plt.subplots(figsize=(9, 3.8))
labels = conv.index[:-1]
x = np.arange(len(labels)); w = 0.36
for i, (grp, color) in enumerate([("Control", C_CONTROL), ("Test", C_TEST)]):
    bars = ax.bar(x + (i - 0.5) * (w + 0.02), conv[grp].iloc[:-1] * 100, width=w, color=color, label=grp)
    for b in bars:
        ax.text(b.get_x() + b.get_width() / 2, b.get_height() + 1.5, f"{b.get_height():.0f}%", ha="center",
                fontsize=9, color=C_INK)
ax.set_xticks(x); ax.set_xticklabels([l.replace(" → ", "\n→ ") for l in labels], color=C_INK)
ax.set_ylabel("Конверсія кроку, %"); ax.set_ylim(0, 110); ax.grid(axis="x", visible=False)
ax.set_title("Конверсія між сусідніми кроками воронки", loc="left", color=C_INK)
ax.legend(frameon=False, loc="upper left"); plt.tight_layout(); plt.show()
""")
    md(r"""
## 3. Верхня частина воронки: реклама → сайт

Воронка сайту не показує, звідки беруться кліки. Додамо рекламні показники: покази, охоплення, CTR і вартість кліку.
""")
    code(r"""
ad = pd.DataFrame({
    "Control": [tot.loc["Витрати, $", "Control"], tot.loc["Покази", "Control"], tot.loc["Охоплення", "Control"], tot.loc["Кліки на сайт", "Control"]],
    "Test": [tot.loc["Витрати, $", "Test"], tot.loc["Покази", "Test"], tot.loc["Охоплення", "Test"], tot.loc["Кліки на сайт", "Test"]],
}, index=["Витрати, $", "Покази", "Охоплення", "Кліки"])
ad.loc["CTR (кліки / покази)"] = ad.loc["Кліки"] / ad.loc["Покази"]
ad.loc["Частота (покази / охоплення)"] = ad.loc["Покази"] / ad.loc["Охоплення"]
ad.loc["CPC, $"] = ad.loc["Витрати, $"] / ad.loc["Кліки"]
ad.loc["CPA, $ за покупку"] = ad.loc["Витрати, $"] / tot.loc["Покупка"]
ad["Test / Control"] = ad["Test"] / ad["Control"]
ad.round(3)
""")
    md(r"""
## 4. Динаміка воронки по днях

Сумарні цифри можуть приховувати нестабільність. Подивимось на щоденні покупки та ключові конверсії.
""")
    code(r"""
d["ctr"] = d.clicks / d.impressions
d["cart_rate"] = d.add_to_cart / d.clicks
d["cr_cart"] = d.purchase / d.add_to_cart
d["cr_click"] = d.purchase / d.clicks

fig, axes = plt.subplots(2, 2, figsize=(11, 6.2), sharex=True)
panels = [("purchase", "Покупки / день", "{:.0f}"), ("ctr", "CTR", "{:.0%}"),
          ("cart_rate", "Конверсія клік → кошик", "{:.0%}"), ("cr_cart", "Конверсія кошик → покупка", "{:.0%}")]
for ax, (m, title, fmt) in zip(axes.ravel(), panels):
    for grp, color in [("Control", C_CONTROL), ("Test", C_TEST)]:
        s = d[d.group == grp].set_index("date")[m]
        ax.plot(s.index, s.values, color=color, lw=2, label=grp)
    ax.set_title(title, loc="left", color=C_INK, fontsize=10)
    ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda v, _, f=fmt: f.format(v)))
    ax.grid(axis="x", visible=False)
axes[0, 0].legend(frameon=False, loc="upper left", ncol=2)
for ax in axes[1]:
    ax.tick_params(axis="x", rotation=45)
fig.suptitle("Щоденна динаміка ключових показників (серпень 2019)", x=0.01, ha="left", color=C_INK)
plt.tight_layout(); plt.show()

anom = d[d.purchase > d.add_to_cart][["group", "date", "add_to_cart", "purchase"]]
print("Дні, коли покупок більше, ніж додавань у кошик:"); anom
""")
    md(r"""
## Висновки

1. **Новий підхід змінив форму воронки, а не її результат.** За 29 порівнюваних днів тестова група отримала **≈ 14.9 тис. покупок проти ≈ 15.2 тис. у контрольній** (−1.9 %, у середньому ≈ 513 проти ≈ 523 покупок на день — різниця в межах випадкових коливань), хоча витрати тестової групи були на **≈ 12 % вищими** за той самий період.
2. **Верх воронки працює ефективніше.** Тестова кампанія показувалася значно рідше (≈ 2.1 млн показів проти ≈ 3.2 млн) і охоплювала на ≈ 41 % менше людей, але її **CTR удвічі вищий (≈ 8.3 % проти ≈ 4.9 %)**, тож кліків на сайт стало більше (≈ 175 тис. проти ≈ 154 тис., +13 %) за майже однакової вартості кліку (≈ $0.43). Нове розміщення приводить більш «клікабельну» аудиторію.
3. **Середина воронки просідає.** Частка кліків, що доходять до кошика, у тестовій групі **≈ 14.6 % проти ≈ 24.4 %** — найбільша втрата нового підходу. Найсильніше падає крок перегляд контенту → кошик (≈ 46 % проти ≈ 67 %): нова аудиторія частіше клікає й переглядає товари, але рідше знаходить те, що хоче купити.
4. **Низ воронки, навпаки, сильніший.** Ті, хто таки додав товар у кошик, у тестовій групі купують значно частіше: **≈ 58 % проти ≈ 40 %**. Разом ці два ефекти майже компенсують одне одного, і наскрізна конверсія клік → покупка у тестовій групі трохи нижча (≈ 8.5 % проти ≈ 9.8 %).
5. **Вартість покупки (CPA) у тестовій групі вища** (≈ $5.0 проти ≈ $4.4, +14 %) через більші витрати при тій самій кількості покупок. Отже, з погляду ефективності витрат новий підхід **не кращий за поточний**.
6. **Якість даних.** В одному дні контрольної групи (30.08) покупок більше, ніж додавань у кошик, ще в одному (06.08) вони майже рівні — імовірно, похибка трекінгу або покупки без кроку «кошик» (швидка покупка). Також 05.08 у контролі відсутні дані воронки. Такі аномалії треба перевіряти з командою трекінгу до прийняття рішень (див. трекінг-план у Завданні №2).

**Продуктовий висновок:** новий підхід варто не впроваджувати «як є», а доопрацювати середину воронки для нової аудиторії (релевантність посадкової сторінки/каталогу під очікування від реклами). Якщо вдасться підняти конверсію клік → кошик хоча б до рівня контролю, поєднання вищого CTR і вищої конверсії кошик → покупка дасть істотний приріст покупок за ті самі витрати. Статистична перевірка відмінностей — у Завданнях №3 та №5.
""")
    return cells


# ===================================================================== TASK 5
def build_task5():
    cells = []
    md = lambda s: cells.append(nbf.v4.new_markdown_cell(s.strip("\n")))
    code = lambda s: cells.append(nbf.v4.new_code_cell(s.strip("\n")))

    md(r"""
# Фінальний проєкт. Завдання №5 — Довірчий інтервал різниці конверсії «кошик → покупка»

**Автор:** Мойсеєнко Світлана · **Дисципліна:** Продуктова аналітика та прикладна статистика (GoIT Neoversity)

**Запитання.** Наскільки відрізняється конверсія кроку покупки після додавання товару в кошик (`# of Purchase / # of Add to Cart`) між тестовою та контрольною групами, і який 95 % довірчий інтервал цієї різниці?

Розрахуємо інтервал трьома способами, щоб перевірити стійкість результату:

1. **На рівні користувачів (агреговані частки):** різниця двох біноміальних часток, інтервал Вальда та z-тест; додатково — інтервал Ньюкомба (на основі Вілсона), стійкіший для часток.
2. **На рівні днів:** середня щоденна конверсія в кожній групі; t-інтервал Уелча для різниці середніх (день — одиниця спостереження).
3. **Бутстреп по днях** — непараметрична перевірка результату 2.
""")
    code(COMMON_SETUP)
    md(r"""
День 05.08, за який у контрольній групі немає даних воронки, виключаємо для обох груп (залишається по 29 днів). Також перевіримо дні, де покупок більше, ніж додавань у кошик — для таких днів конверсія кроку перевищує 100 %, що фізично неможливо й указує на проблему трекінгу.
""")
    code(r"""
d = df[df.complete].copy()
d["cr_cart"] = d.purchase / d.add_to_cart
anom = d[d.cr_cart > 1][["group", "date", "add_to_cart", "purchase", "cr_cart"]]
print("Дні з конверсією кошик → покупка > 100 %:")
anom
""")
    md(r"""
## 1. Агрегований підхід: різниця часток (користувачі як одиниця)

Усі додавання в кошик за 29 днів розглядаємо як «спроби», а покупки — як «успіхи». Це найпотужніший підхід (десятки тисяч спостережень), але він припускає незалежність подій і однорідність днів.
""")
    code(r"""
def two_prop(x1, n1, x2, n2, alpha=0.05):
    p1, p2 = x1 / n1, x2 / n2
    diff = p1 - p2
    se = np.sqrt(p1 * (1 - p1) / n1 + p2 * (1 - p2) / n2)
    z = stats.norm.ppf(1 - alpha / 2)
    pp = (x1 + x2) / (n1 + n2)
    zstat = diff / np.sqrt(pp * (1 - pp) * (1 / n1 + 1 / n2))
    return {"Test": p1, "Control": p2, "diff": diff, "lo": diff - z * se, "hi": diff + z * se,
            "z": zstat, "p": 2 * stats.norm.sf(abs(zstat))}

def wilson(x, n, alpha=0.05):
    z = stats.norm.ppf(1 - alpha / 2)
    p = x / n; denom = 1 + z**2 / n
    centre = (p + z**2 / (2 * n)) / denom
    half = z * np.sqrt(p * (1 - p) / n + z**2 / (4 * n**2)) / denom
    return centre - half, centre + half

def newcombe(x1, n1, x2, n2, alpha=0.05):
    # Newcombe (1998) hybrid score interval for p1 - p2
    l1, u1 = wilson(x1, n1, alpha); l2, u2 = wilson(x2, n2, alpha)
    p1, p2 = x1 / n1, x2 / n2
    diff = p1 - p2
    lo = diff - np.sqrt((p1 - l1)**2 + (u2 - p2)**2)
    hi = diff + np.sqrt((u1 - p1)**2 + (p2 - l2)**2)
    return lo, hi

T = d[d.group == "Test"]; C = d[d.group == "Control"]
xT, nT = int(T.purchase.sum()), int(T.add_to_cart.sum())
xC, nC = int(C.purchase.sum()), int(C.add_to_cart.sum())
res = two_prop(xT, nT, xC, nC)
nc_lo, nc_hi = newcombe(xT, nT, xC, nC)

print(f"Test:    {xT:>6} покупок / {nT:>6} додавань у кошик = {res['Test']:.4f}")
print(f"Control: {xC:>6} покупок / {nC:>6} додавань у кошик = {res['Control']:.4f}")
print(f"\nРізниця (Test − Control): {res['diff']:+.4f}  ({res['diff']*100:+.2f} п.п.)")
print(f"95% ДІ (Вальд):     [{res['lo']:+.4f}; {res['hi']:+.4f}]")
print(f"95% ДІ (Ньюкомб):   [{nc_lo:+.4f}; {nc_hi:+.4f}]")
print(f"z = {res['z']:.2f}, p-value = {res['p']:.2e}")
""")
    md(r"""
## 2. Поденний підхід: день як одиниця спостереження

Щоденні конверсії відображають природну мінливість між днями (сезонність, зміни ставок, аномалії трекінгу), тому інтервал виходить ширшим і чеснішим щодо невизначеності. Використовуємо t-інтервал Уелча (різні дисперсії, різна кількість днів).
""")
    code(r"""
def welch_ci(a, b, alpha=0.05):
    a, b = np.asarray(a, float), np.asarray(b, float)
    diff = a.mean() - b.mean()
    va, vb = a.var(ddof=1) / len(a), b.var(ddof=1) / len(b)
    se = np.sqrt(va + vb)
    dof = (va + vb)**2 / (va**2 / (len(a) - 1) + vb**2 / (len(b) - 1))
    tcrit = stats.t.ppf(1 - alpha / 2, dof)
    t_stat, p = stats.ttest_ind(a, b, equal_var=False)
    return {"Test": a.mean(), "Control": b.mean(), "diff": diff, "lo": diff - tcrit * se, "hi": diff + tcrit * se,
            "t": t_stat, "dof": dof, "p": p}

daily_all = welch_ci(T.cr_cart, C.cr_cart)
clean = d[d.cr_cart <= 1]
daily_clean = welch_ci(clean.loc[clean.group == "Test", "cr_cart"], clean.loc[clean.group == "Control", "cr_cart"])

summary = pd.DataFrame({
    "Агреговано (Вальд)": {"Test": res["Test"], "Control": res["Control"], "Різниця": res["diff"], "95% ДІ нижня": res["lo"], "95% ДІ верхня": res["hi"], "p-value": res["p"]},
    "Агреговано (Ньюкомб)": {"Test": res["Test"], "Control": res["Control"], "Різниця": res["diff"], "95% ДІ нижня": nc_lo, "95% ДІ верхня": nc_hi, "p-value": res["p"]},
    "По днях (усі дні)": {"Test": daily_all["Test"], "Control": daily_all["Control"], "Різниця": daily_all["diff"], "95% ДІ нижня": daily_all["lo"], "95% ДІ верхня": daily_all["hi"], "p-value": daily_all["p"]},
    "По днях (без аномалії 30.08)": {"Test": daily_clean["Test"], "Control": daily_clean["Control"], "Різниця": daily_clean["diff"], "95% ДІ нижня": daily_clean["lo"], "95% ДІ верхня": daily_clean["hi"], "p-value": daily_clean["p"]},
})
summary.style.format("{:.4f}")
""")
    md(r"""
## 3. Бутстреп по днях (непараметрична перевірка)
""")
    code(r"""
rng = np.random.default_rng(42)
a, b = T.cr_cart.to_numpy(), C.cr_cart.to_numpy()
boot = np.array([rng.choice(a, len(a)).mean() - rng.choice(b, len(b)).mean() for _ in range(10_000)])
b_lo, b_hi = np.percentile(boot, [2.5, 97.5])
print(f"Бутстреп (10 000 повторень): різниця = {boot.mean():+.4f}, 95% ДІ = [{b_lo:+.4f}; {b_hi:+.4f}]")

fig, ax = plt.subplots(figsize=(8, 3.4))
ax.hist(boot, bins=50, color=C_TEST, alpha=0.85, edgecolor="white", linewidth=0.5)
ax.axvline(0, color=C_INK, lw=1.2, ls="--")
ax.axvline(b_lo, color=C_CONTROL, lw=1.5); ax.axvline(b_hi, color=C_CONTROL, lw=1.5)
ax.text(b_lo, ax.get_ylim()[1] * 0.92, f" {b_lo:+.3f}", color=C_CONTROL, fontsize=9)
ax.text(b_hi, ax.get_ylim()[1] * 0.92, f" {b_hi:+.3f}", color=C_CONTROL, fontsize=9)
ax.set_xlabel("Різниця середньої щоденної конверсії кошик → покупка (Test − Control)")
ax.set_ylabel("Кількість вибірок"); ax.grid(axis="x", visible=False)
ax.set_title("Бутстреп-розподіл різниці конверсій (межі — 95% ДІ)", loc="left", color=C_INK)
plt.tight_layout(); plt.show()
""")
    code(r"""
# Порівняння інтервалів на одному графіку
fig, ax = plt.subplots(figsize=(8, 3.2))
items = [("Агреговано (Вальд)", res["diff"], res["lo"], res["hi"]),
         ("Агреговано (Ньюкомб)", res["diff"], nc_lo, nc_hi),
         ("По днях, t-інтервал", daily_all["diff"], daily_all["lo"], daily_all["hi"]),
         ("По днях, без 30.08", daily_clean["diff"], daily_clean["lo"], daily_clean["hi"]),
         ("По днях, бутстреп", boot.mean(), b_lo, b_hi)]
for i, (name, m, lo, hi) in enumerate(items[::-1]):
    ax.plot([lo, hi], [i, i], color=C_TEST, lw=2.5); ax.plot(m, i, "o", color=C_TEST, ms=7)
    ax.text(hi + 0.005, i, f"{m*100:+.1f} п.п. [{lo*100:+.1f}; {hi*100:+.1f}]", va="center", fontsize=9, color=C_INK)
ax.axvline(0, color=C_INK, lw=1, ls="--")
ax.set_yticks(range(len(items))); ax.set_yticklabels([n for n, *_ in items[::-1]], color=C_INK)
ax.set_xlim(-0.05, 0.42); ax.grid(axis="y", visible=False)
ax.xaxis.set_major_formatter(plt.FuncFormatter(lambda v, _: f"{v*100:+.0f} п.п."))
ax.set_title("95% довірчі інтервали різниці конверсії кошик → покупка (Test − Control)", loc="left", color=C_INK)
plt.tight_layout(); plt.show()
""")
    md(r"""
## Висновки

1. **Конверсія «кошик → покупка» у тестовій групі істотно вища.** За агрегованими даними вона становить ≈ 58.3 % проти ≈ 40.2 % у контрольній групі; різниця **≈ +18.1 п.п., 95 % ДІ ≈ [+17.3; +18.9] п.п.** (інтервали Вальда та Ньюкомба практично збігаються), z ≈ 45, p ≈ 0. Нуль у інтервал не входить.
2. **Поденний підхід підтверджує напрям ефекту, але з більшою невизначеністю:** різниця середніх щоденних конверсій ≈ +15.4 п.п., 95 % ДІ ≈ [+3.7; +27.2] п.п. (t-інтервал Уелча, p ≈ 0.011); бутстреп дає практично ті самі межі ([+3.5; +25.8] п.п.). Після виключення аномального дня 30.08 (контроль, конверсія > 100 %, що завищує середнє контрольної групи) різниця зростає до ≈ +19.2 п.п., 95 % ДІ ≈ [+10.1; +28.3] п.п. — інтервал залишається повністю додатним.
3. **Який інтервал використовувати?** Агрегований інтервал надто «впевнений»: він вважає кожне додавання в кошик незалежним спостереженням та ігнорує щоденну мінливість і аномалії трекінгу. Для прийняття рішень коректніше спиратися на **поденний інтервал: приріст конверсії кошик → покупка щонайменше ≈ +4 п.п., найімовірніше ≈ +15–19 п.п.**
4. **Бізнес-інтерпретація.** Нове розміщення реклами приводить аудиторію, яка рідше додає товари в кошик (див. Завдання №4), але, діставшись до кошика, значно частіше завершує покупку. Це свідчить про вищу «купівельну готовність» аудиторії нового каналу. Проте сам по собі цей показник не означає, що кампанія краща загалом: кількість покупок у групах однакова, а витрати в тестовій групі вищі (Завдання №3).
""")
    return cells


def run(cells, out_path):
    nb = nbf.v4.new_notebook()
    nb["cells"] = cells
    nb["metadata"] = {"kernelspec": {"display_name": "Python 3", "language": "python", "name": "python3"},
                      "language_info": {"name": "python"}}
    NotebookClient(nb, timeout=300, kernel_name="python3", resources={"metadata": {"path": str(HERE)}}).execute()
    nbf.write(nb, str(out_path))
    print("written", out_path)


if __name__ == "__main__":
    run(build_task4(), HERE / f"{SURNAME}_Final Project-4.ipynb")
    run(build_task5(), HERE / f"{SURNAME}_Final Project-5.ipynb")
