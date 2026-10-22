"""
preprocess.py - Préparation des données M5 Walmart.

Périmètre : magasins CA_1, TX_1, WI_1 ; top 5 produits (volume total) par magasin
et par catégorie (FOODS, HOBBIES, HOUSEHOLD), soit 3 × 3 × 5 = 45 séries.

Sorties (dossier output/) :
  - series_long.csv         : format long, une ligne par jour × magasin × produit
  - series_excel.xlsx       : un onglet par magasin, tableau date × produit
  - series_prophet.csv      : format Prophet (ds, y, store_id, item_id + features)
  - metrics_summary.xlsx    : gabarit vide, rempli ensuite par compare_all.py
  - top_produits_par_categorie.csv : classement du top 5 par store × catégorie
    (volume, chiffre d'affaires, part de la catégorie, jours sans vente)
  - ventes_ajoutees.csv     : historique cumulé des ventes saisies dans Excel
    (frmGestionProduits -> nouvelles_ventes.csv), réappliqué à chaque preprocessing
"""
import os
import sys
import traceback

import pandas as pd

from utils import (DATA_DIR, OUTPUT_DIR, STORES, LONG_FILE, PROPHET_FILE,
                   add_active_features, future_calendar, step)
from utils import TOP_N_PER_CAT, CATEGORIES

SALES_FILE = os.path.join(DATA_DIR, "sales_train_validation.csv")
CALENDAR_FILE = os.path.join(DATA_DIR, "calendar.csv")
PRICES_FILE = os.path.join(DATA_DIR, "sell_prices.csv")
EXCEL_FILE = os.path.join(OUTPUT_DIR, "series_excel.xlsx")
SUMMARY_FILE = os.path.join(OUTPUT_DIR, "metrics_summary.xlsx")
TOP_CSV = os.path.join(OUTPUT_DIR, "top_produits_par_categorie.csv")
NEW_SALES_FILE = os.path.join(OUTPUT_DIR, "nouvelles_ventes.csv")      # écrit par le VBA
ADDED_SALES_FILE = os.path.join(OUTPUT_DIR, "ventes_ajoutees.csv")     # historique cumulé
NEW_SALES_COLS = ["date", "store_id", "item_id", "cat_id", "sales", "sell_price", "snap", "has_event"]

# Colonne SNAP à utiliser selon l'état du magasin
STATE_SNAP = {"CA": "snap_CA", "TX": "snap_TX", "WI": "snap_WI"}


def check_inputs():
    """Vérifie la présence des fichiers d'entrée."""
    for f in (SALES_FILE, CALENDAR_FILE, PRICES_FILE):
        if not os.path.exists(f):
            raise FileNotFoundError(f"Fichier manquant : {f}")
    os.makedirs(OUTPUT_DIR, exist_ok=True)


def load_sales_top_items():
    """Lit les ventes par morceaux, garde les 3 magasins puis le top N par magasin × catégorie."""
    step("Lecture de sales_train_validation.csv (filtrage des 3 magasins)...")
    chunks = []
    for chunk in pd.read_csv(SALES_FILE, chunksize=5000):
        chunks.append(chunk[chunk["store_id"].isin(STORES)])
    sales = pd.concat(chunks, ignore_index=True).copy()
    day_cols = [c for c in sales.columns if c.startswith("d_")]
    print(f"   {len(sales)} produits trouvés dans {STORES}, {len(day_cols)} jours")

    # Volume total de ventes de chaque produit sur toute la période
    sales["total_sales"] = sales[day_cols].sum(axis=1)

    # Top N par store × catégorie (équilibré entre FOODS, HOBBIES, HOUSEHOLD)
    top = (sales[sales["cat_id"].isin(CATEGORIES)]
                .sort_values("total_sales", ascending=False)
                .groupby(["store_id", "cat_id"], group_keys=False)
                .head(TOP_N_PER_CAT))

    n_series = top.groupby(["store_id", "cat_id"]).size()
    print(f"   Séries sélectionnées : {len(top)} ({TOP_N_PER_CAT} par catégorie × "
          f"{len(CATEGORIES)} catégories × {len(STORES)} stores)")
    print(n_series.to_string())
    return top, day_cols


def to_long(top, day_cols):
    """Passe du format wide (d_1..d_1913) au format long."""
    step("Passage au format long...")
    id_cols = ["item_id", "dept_id", "cat_id", "store_id", "state_id"]
    long = top.melt(id_vars=id_cols, value_vars=day_cols, var_name="d", value_name="sales")
    print(f"   {len(long):,} lignes")
    return long


def merge_calendar(long):
    """Ajoute dates, événements, SNAP du bon état et variables calendaires."""
    step("Fusion avec calendar.csv...")
    cal = pd.read_csv(CALENDAR_FILE, parse_dates=["date"])
    long = long.merge(cal, on="d", how="left")

    # SNAP : on prend le flag de l'état du magasin
    long["snap"] = 0
    for state, col in STATE_SNAP.items():
        mask = long["state_id"] == state
        long.loc[mask, "snap"] = long.loc[mask, col]

    # Événements : nom (concatène les 2 événements possibles) et type principal
    long["has_event"] = long["event_name_1"].notna().astype(int)
    long["event_name"] = long["event_name_1"]
    both = long["event_name_2"].notna()
    long.loc[both, "event_name"] = long.loc[both, "event_name_1"] + "|" + long.loc[both, "event_name_2"]
    long["event_type"] = long["event_type_1"]

    # Variables calendaires
    long["day_of_week"] = long["date"].dt.dayofweek          # 0 = lundi
    long["month"] = long["date"].dt.month
    long["year"] = long["date"].dt.year
    long["week_of_year"] = long["date"].dt.isocalendar().week.astype(int)
    long["is_weekend"] = (long["day_of_week"] >= 5).astype(int)
    long["days_to_monthend"] = (long["date"] + pd.offsets.MonthEnd(0) - long["date"]).dt.days
    return long


def merge_prices(long):
    """Ajoute le prix hebdomadaire de chaque produit dans chaque magasin."""
    step("Fusion avec sell_prices.csv (lecture par morceaux)...")
    items = set(long["item_id"].unique())
    chunks = []
    for chunk in pd.read_csv(PRICES_FILE, chunksize=1_000_000):
        chunks.append(chunk[chunk["store_id"].isin(STORES) & chunk["item_id"].isin(items)])
    prices = pd.concat(chunks, ignore_index=True)
    long = long.merge(prices, on=["store_id", "item_id", "wm_yr_wk"], how="left")

    # Prix manquants (avant la mise en vente) : complétés par le prix connu le plus proche
    n_missing = long["sell_price"].isna().sum()
    long = long.sort_values(["store_id", "item_id", "date"])
    long["sell_price"] = long.groupby(["store_id", "item_id"])["sell_price"].transform(
        lambda s: s.bfill().ffill())
    print(f"   {n_missing} prix manquants complétés (bfill/ffill)")
    return long


def save_long(long):
    """Sauvegarde series_long.csv avec les colonnes demandées."""
    step("Écriture de series_long.csv...")
    cols = ["date", "store_id", "item_id", "cat_id", "dept_id", "sales", "sell_price",
            "snap", "has_event", "event_name", "event_type", "day_of_week", "month",
            "year", "week_of_year", "is_weekend", "days_to_monthend"]
    long = long[cols].sort_values(["store_id", "item_id", "date"]).reset_index(drop=True)
    long.to_csv(LONG_FILE, index=False)
    print(f"   -> {LONG_FILE}")
    return long


def save_excel(long):
    """Un onglet par magasin : dates en lignes, produits en colonnes."""
    step("Écriture de series_excel.xlsx...")
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(bold=True, color="FFFFFF")

    with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl") as writer:
        for store in STORES:
            pivot = (long[long["store_id"] == store]
                     .pivot(index="date", columns="item_id", values="sales"))
            # Colonnes triées par volume décroissant
            pivot = pivot[pivot.sum().sort_values(ascending=False).index]
            pivot.index = pivot.index.date
            pivot.index.name = "date"
            pivot.columns.name = None
            pivot.to_excel(writer, sheet_name=store)

            # Mise en forme : en-têtes colorés, volets figés, largeurs, format date
            ws = writer.sheets[store]
            for cell in ws[1]:
                cell.fill, cell.font = header_fill, header_font
                cell.alignment = Alignment(horizontal="center")
            ws.freeze_panes = "B2"
            ws.column_dimensions["A"].width = 12
            for i in range(2, pivot.shape[1] + 2):
                ws.column_dimensions[get_column_letter(i)].width = 18
            for row in ws.iter_rows(min_row=2, min_col=1, max_col=1):
                row[0].number_format = "yyyy-mm-dd"
            ws.auto_filter.ref = ws.dimensions
    print(f"   -> {EXCEL_FILE}")


def save_prophet(long):
    """Format Prophet : ds, y, store_id, item_id + features actives."""
    step("Écriture de series_prophet.csv...")
    feats = add_active_features(long)
    cols = ["ds", "y", "store_id", "item_id", "sell_price", "price_change_pct",
            "price_rel_mean", "snap", "has_event", "event_type_code",
            "days_to_christmas", "is_holiday_week", "day_of_week", "month",
            "is_weekend", "days_to_monthend"]
    out = feats.rename(columns={"date": "ds", "sales": "y"})[cols]
    out.to_csv(PROPHET_FILE, index=False)
    print(f"   -> {PROPHET_FILE}")


def init_summary():
    """Crée un gabarit metrics_summary.xlsx (rempli par compare_all.py)."""
    step("Création du gabarit metrics_summary.xlsx...")
    template = pd.DataFrame(columns=["Modèle", "Approche", "MAE", "MAPE", "RMSE"])
    with pd.ExcelWriter(SUMMARY_FILE, engine="openpyxl") as writer:
        template.to_excel(writer, sheet_name="Global", index=False)
        for store in STORES:
            template.to_excel(writer, sheet_name=store, index=False)
    print(f"   -> {SUMMARY_FILE} (vide, rempli par compare_all.py)")


def print_stats(long):
    """Statistiques descriptives finales."""
    step("Statistiques")
    n_series = long.groupby(["store_id", "item_id"]).ngroups
    print(f"   Nombre de séries : {n_series}")
    print(f"   Période          : {long['date'].min().date()} -> {long['date'].max().date()} "
          f"({long['date'].nunique()} jours)")

    print("\n   Part de jours à 0 vente (zero_ratio) par magasin :")
    zr = long.groupby("store_id")["sales"].apply(lambda s: (s == 0).mean())
    for store, v in zr.items():
        print(f"     {store} : {v:.1%}")

    # Statistiques par catégorie : HOBBIES et HOUSEHOLD sont beaucoup plus intermittents
    print("\n   Zero ratio par catégorie :")
    zr_cat = long.groupby("cat_id")["sales"].apply(lambda s: (s == 0).mean())
    for cat, v in zr_cat.items():
        print(f"     {cat} : {v:.1%}")

    print("\n   Ventes moyennes par catégorie (unités/jour/série) :")
    mean_cat = long[long.sales > 0].groupby("cat_id")["sales"].mean()
    for cat, v in mean_cat.items():
        print(f"     {cat} : {v:.1f}")

    print(f"\n   Top {TOP_N_PER_CAT} par magasin et par catégorie :")
    vol = long.groupby(["store_id", "cat_id", "item_id"])["sales"].sum().reset_index()
    for (store, cat), d in vol.groupby(["store_id", "cat_id"]):
        items = d.sort_values("sales", ascending=False)["item_id"]
        print(f"     {store} / {cat:<9} : " + ", ".join(items))


def top_products_analysis(long):
    """
    Étude du top N par store × catégorie : rang, volume, chiffre d'affaires, part du
    volume de la catégorie (au sein du top N), jours sans vente, prix moyen.
    Sortie : top_produits_par_categorie.csv.
    """
    step(f"Analyse du top {TOP_N_PER_CAT} produits par store et par catégorie...")
    d = long.assign(revenue=long["sales"] * long["sell_price"])
    tab = (d.groupby(["store_id", "cat_id", "item_id", "dept_id"])
             .agg(ventes_totales=("sales", "sum"), ca_total=("revenue", "sum"),
                  ventes_moy_jour=("sales", "mean"), prix_moyen=("sell_price", "mean"),
                  zero_ratio=("sales", lambda s: (s == 0).mean()))
             .reset_index())
    # Ventes des 365 derniers jours : dynamique récente du produit
    last_year = d[d["date"] > d["date"].max() - pd.Timedelta(days=365)]
    tab = tab.merge(last_year.groupby(["store_id", "item_id"])["sales"].sum()
                    .rename("ventes_12_derniers_mois").reset_index(), on=["store_id", "item_id"])
    tab = tab.sort_values(["store_id", "cat_id", "ventes_totales"], ascending=[True, True, False])
    tab.insert(2, "rang", tab.groupby(["store_id", "cat_id"]).cumcount() + 1)
    tab["part_top_cat_pct"] = (tab["ventes_totales"] /
                               tab.groupby(["store_id", "cat_id"])["ventes_totales"].transform("sum") * 100)
    tab = tab.round({"ca_total": 2, "ventes_moy_jour": 2, "prix_moyen": 2,
                     "zero_ratio": 3, "part_top_cat_pct": 1})
    tab.to_csv(TOP_CSV, index=False)
    print(f"   -> {TOP_CSV}")


# ---------------------------------------------------------------------------
# Nouvelles ventes saisies via frmGestionProduits (Excel)
# ---------------------------------------------------------------------------
def collect_new_sales():
    """
    Fusionne output/nouvelles_ventes.csv (écrit par le VBA) dans l'historique cumulé
    ventes_ajoutees.csv, avec déduplication sur (date, store_id, item_id) : la saisie
    la plus récente l'emporte. series_long.csv est reconstruit depuis les fichiers
    M5 bruts à chaque preprocessing : sans cet historique, les ventes saisies
    seraient perdues au preprocessing suivant.
    NE PAS supprimer nouvelles_ventes.csv ici : le VBA le vide après succès.
    """
    if not os.path.exists(NEW_SALES_FILE):
        return
    nv = pd.read_csv(NEW_SALES_FILE, parse_dates=["date"])
    if nv.empty:
        return
    step(f"Fusion de {len(nv)} nouvelles ventes dans l'historique des saisies...")
    nv = nv.reindex(columns=NEW_SALES_COLS)
    if os.path.exists(ADDED_SALES_FILE):
        hist = pd.read_csv(ADDED_SALES_FILE, parse_dates=["date"])
        merged = pd.concat([hist, nv], ignore_index=True)
    else:
        merged = nv
    before = len(merged)
    merged = merged.drop_duplicates(subset=["date", "store_id", "item_id"], keep="last")
    merged.sort_values(["store_id", "item_id", "date"]).to_csv(ADDED_SALES_FILE, index=False)
    print(f"   {len(merged)} ventes saisies au total ({before - len(merged)} doublons ignorés)")


def apply_added_sales(long):
    """
    Applique ventes_ajoutees.csv au jeu reconstruit : une ligne existante (même date,
    store, produit) est remplacée ; une ligne nouvelle (date future ou nouveau produit)
    est ajoutée avec ses variables calendaires recalculées.
    """
    if not os.path.exists(ADDED_SALES_FILE):
        return long
    add = pd.read_csv(ADDED_SALES_FILE, parse_dates=["date"])
    if add.empty:
        return long
    step(f"Application des {len(add)} ventes saisies dans Excel...")
    keys = ["date", "store_id", "item_id"]
    add["cat_id"] = add["cat_id"].fillna(add["item_id"].str.split("_").str[0])
    add["dept_id"] = add["item_id"].str.rsplit("_", n=1).str[0]
    add["snap"] = add["snap"].fillna(0).astype(int)
    add["has_event"] = add["has_event"].fillna(0).astype(int)

    # Événements : nom / type lus dans calendar.csv quand la date y figure
    cal = future_calendar(add["date"])
    add["event_name"] = add["date"].map(cal["event_name"]).where(add["has_event"] == 1)
    add["event_type"] = add["date"].map(cal["event_type"]).where(add["has_event"] == 1)
    add[["event_name", "event_type"]] = add[["event_name", "event_type"]].replace("none", None)
    add["day_of_week"] = add["date"].dt.dayofweek
    add["month"] = add["date"].dt.month
    add["year"] = add["date"].dt.year
    add["week_of_year"] = add["date"].dt.isocalendar().week.astype(int)
    add["is_weekend"] = (add["day_of_week"] >= 5).astype(int)
    add["days_to_monthend"] = (add["date"] + pd.offsets.MonthEnd(0) - add["date"]).dt.days

    is_existing = add.set_index(keys).index.isin(long.set_index(keys).index)
    n_upd, n_new = int(is_existing.sum()), int((~is_existing).sum())
    # Lignes existantes : on remplace ventes, prix, SNAP et événement saisis
    if n_upd:
        upd = add[is_existing].set_index(keys)[["sales", "sell_price", "snap", "has_event"]]
        long = long.set_index(keys)
        long.update(upd)
        long = long.reset_index()
    # Lignes nouvelles : ajoutées telles quelles
    if n_new:
        long = pd.concat([long, add[~is_existing]], ignore_index=True)
    long = long.sort_values(["store_id", "item_id", "date"]).reset_index(drop=True)
    print(f"   {n_upd} ligne(s) mise(s) à jour, {n_new} ligne(s) ajoutée(s)")

    # Contrôle : un trou de dates fausserait les lags (décalage par ligne)
    gaps = long.groupby(["store_id", "item_id"])["date"].apply(lambda d: (d.diff().dt.days > 1).sum())
    if gaps.sum():
        print(f"   [AVERT] {int((gaps > 0).sum())} série(s) avec des jours manquants : "
              "saisir des ventes pour chaque jour consécutif.")
    return long


def main():
    print("=" * 70)
    print(f" PREPROCESSING M5 - CA_1 / TX_1 / WI_1 - top {TOP_N_PER_CAT} produits par catégorie")
    print("=" * 70)
    check_inputs()
    # Nouvelles ventes saisies dans Excel -> historique cumulé ventes_ajoutees.csv
    collect_new_sales()
    top, day_cols = load_sales_top_items()
    long = to_long(top, day_cols)
    long = merge_calendar(long)
    long = merge_prices(long)
    long = apply_added_sales(long)
    long = save_long(long)
    save_excel(long)
    save_prophet(long)
    init_summary()
    print_stats(long)
    top_products_analysis(long)
    print("\nPreprocessing terminé avec succès.")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"\n[ERREUR] Preprocessing interrompu : {e}")
        traceback.print_exc()
        sys.exit(1)
