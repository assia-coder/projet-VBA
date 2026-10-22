"""
monthly_forecast.py - Prévisions mensuelles sans récursivité, à 1, 2 et 3 mois.

Approche : agréger les ventes journalières en mois, puis entraîner LightGBM
sur les séries mensuelles. Aucun lag n'est jamais une prévision.

Horizons multiples par stratégie DIRECTE : un modèle par horizon h (M+1, M+2, M+3).
Le modèle h n'utilise que des ventes connues h mois avant le mois cible : la
prévision à M+3 ne réutilise donc jamais les prévisions M+1 ou M+2 (pas de
divergence, contrairement à l'approche récursive de la V2).

Intervalles de prévision : deux modèles LightGBM quantiles par horizon
(quantiles 10 % et 90 %), recalibrés par conformal prediction (CQR) sur les
6 derniers mois d'entraînement -> intervalle à 80 %, dont la couverture réelle
est mesurée sur la période de test.

Découpage : 7 derniers mois complets = test, le reste = entraînement
(les mois incomplets du début et de la fin de l'historique sont exclus).

Sorties (output/) :
  - forecasts_monthly.csv         : date, store_id, item_id, cat_id, horizon, y_true, y_pred, y_low, y_high
  - metrics_monthly.csv           : MAE, MAPE, RMSE, WAPE, couverture (global, store, item, category)
  - monthly_forecast_next.csv     : prévision M+1 (1 ligne par série) + bornes + M+2 et M+3
  - monthly_forecast_horizons.csv : format long, 1 ligne par série × horizon
  - monthly_forecast_next.xlsx    : mêmes données, format Excel lisible
  - model_monthly.joblib          : modèle LightGBM M+1 (point)
"""
import os
import sys
import traceback

import joblib
import numpy as np
import pandas as pd

from utils import DATA_DIR, OUTPUT_DIR, load_long, metrics_table, step

N_TEST_MONTHS = 7
HORIZONS = (1, 2, 3)
QUANTILES = (0.10, 0.90)       # intervalle de prévision à 80 %
CALENDAR_FILE = os.path.join(DATA_DIR, "calendar.csv")
STATE_SNAP = {"CA": "snap_CA", "TX": "snap_TX", "WI": "snap_WI"}

MONTHLY_FEATURES = [
    "lag_1", "lag_2", "lag_3", "lag_6", "lag_12",
    "rolling_mean_3", "rolling_mean_6",
    "month_num", "year",
    "sell_price", "snap", "has_event",
    "item_code", "store_code", "cat_code",
]


# ---------------------------------------------------------------------------
# 1. Chargement et agrégation mensuelle
# ---------------------------------------------------------------------------
def build_monthly(df):
    """Agrège les ventes journalières par mois ; exclut les mois incomplets."""
    step("Agrégation mensuelle des ventes journalières...")
    df["month_period"] = df["date"].dt.to_period("M")

    # Mois incomplets (ex. janvier 2011 : 3 jours, avril 2016 : 24 jours) :
    # leur somme serait artificiellement basse et fausserait lags et métriques.
    days_seen = df.groupby("month_period")["date"].nunique()
    full = days_seen[days_seen == days_seen.index.days_in_month].index
    dropped = sorted(set(days_seen.index) - set(full))
    df = df[df["month_period"].isin(full)]
    print(f"   Mois incomplets exclus : {', '.join(str(m) for m in dropped) or 'aucun'}")

    monthly = df.groupby(["store_id", "item_id", "cat_id", "month_period"]).agg(
        sales=("sales", "sum"),
        sell_price=("sell_price", "mean"),
        snap=("snap", "mean"),
        has_event=("has_event", "max"),
    ).reset_index()
    monthly["month_num"] = monthly["month_period"].dt.month
    monthly["year"] = monthly["month_period"].dt.year
    monthly["date"] = monthly["month_period"].dt.to_timestamp()
    print(f"   {monthly['month_period'].nunique()} mois complets "
          f"({monthly['month_period'].min()} -> {monthly['month_period'].max()}), "
          f"{monthly.groupby(['store_id', 'item_id']).ngroups} séries")
    return monthly


# ---------------------------------------------------------------------------
# 2. Features mensuelles (TOUTES basées sur des données réelles agrégées)
# ---------------------------------------------------------------------------
def add_monthly_features(monthly, h=1):
    """
    Features pour l'horizon h. Les noms sont identiques quel que soit h :
    "lag_1" = dernier mois connu au moment de la prévision (mois cible - h),
    "lag_2" = celui d'avant, etc. lag_12 reste le même mois l'année précédente
    (toujours connu tant que h <= 12). Pour h = 1, on retrouve les lags classiques.
    """
    monthly = monthly.sort_values(["store_id", "item_id", "date"]).copy()
    g = monthly.groupby(["store_id", "item_id"])["sales"]
    monthly["lag_1"] = g.shift(h)          # dernier mois connu
    monthly["lag_2"] = g.shift(h + 1)
    monthly["lag_3"] = g.shift(h + 2)
    monthly["lag_6"] = g.shift(h + 5)
    monthly["lag_12"] = g.shift(12)        # même mois l'année d'avant (très puissant)
    # Moyennes glissantes calculées série par série (transform) : sans groupby,
    # la fenêtre déborderait d'un produit sur le suivant.
    monthly["rolling_mean_3"] = g.transform(lambda s: s.shift(h).rolling(3).mean())
    monthly["rolling_mean_6"] = g.transform(lambda s: s.shift(h).rolling(6).mean())
    monthly["item_code"] = monthly["item_id"].astype("category").cat.codes
    monthly["store_code"] = monthly["store_id"].astype("category").cat.codes
    monthly["cat_code"] = monthly["cat_id"].astype("category").cat.codes
    return monthly


# ---------------------------------------------------------------------------
# Modèles : point (régression L2) + 2 quantiles par horizon
# ---------------------------------------------------------------------------
def make_models():
    from lightgbm import LGBMRegressor
    point = LGBMRegressor(n_estimators=400, learning_rate=0.05, num_leaves=31,
                          subsample=0.8, subsample_freq=1, random_state=42, verbose=-1)
    # Modèles quantiles un peu plus simples (moins de feuilles) : plus stables
    quant = {q: LGBMRegressor(objective="quantile", alpha=q, n_estimators=300,
                              learning_rate=0.05, num_leaves=15, min_child_samples=20,
                              subsample=0.8, subsample_freq=1, random_state=42, verbose=-1)
             for q in QUANTILES}
    return point, quant


def raw_interval(quant, X):
    """Bornes brutes des modèles quantiles."""
    return quant[QUANTILES[0]].predict(X), quant[QUANTILES[1]].predict(X)


def calibrate(train, calib_months=6):
    """
    Recalibrage conformal (CQR normalisé) : les quantiles LightGBM bruts sont
    trop étroits (~65-70 % de couverture au lieu de 80 %).
    1. modèles quantiles entraînés sans les `calib_months` derniers mois du train ;
    2. score de chaque ligne de calibration = dépassement de l'intervalle,
       rapporté à sa largeur : s = max(bas - y, y - haut) / (haut - bas + 1) ;
    3. q = quantile 80 % (corrigé n+1) des scores -> intervalle final
       [bas - q·largeur, haut + q·largeur].
    """
    months = np.sort(train["month_period"].unique())
    start = months[-calib_months]
    fit, cal = train[train.month_period < start], train[train.month_period >= start]
    _, quant = make_models()
    for model in quant.values():
        model.fit(fit[MONTHLY_FEATURES], fit["sales"])
    lo, hi = raw_interval(quant, cal[MONTHLY_FEATURES])
    y = cal["sales"].to_numpy()
    scores = np.maximum(lo - y, y - hi) / (hi - lo + 1)
    level = min(1.0, (QUANTILES[1] - QUANTILES[0]) * (1 + 1 / len(scores)))
    return float(np.quantile(scores, level))


def predict_all(models, X):
    """Prévision ponctuelle + bornes calibrées ; garantit bas <= prévision <= haut."""
    point, quant, q = models
    y = np.clip(point.predict(X), 0, None)
    lo, hi = raw_interval(quant, X)
    width = hi - lo + 1
    lo, hi = np.clip(lo - q * width, 0, None), np.clip(hi + q * width, 0, None)
    return y.round(1), np.minimum(lo, y).round(1), np.maximum(hi, y).round(1)


def coverage(d):
    """Part des valeurs réelles comprises dans l'intervalle [y_low, y_high] (en %)."""
    return round(float(((d["y_true"] >= d["y_low"]) & (d["y_true"] <= d["y_high"])).mean() * 100), 1)


def add_coverage(mt, fc):
    """Ajoute la colonne coverage_80 (même découpage que metrics_table)."""
    cov = []
    for _, r in mt.iterrows():
        d = fc
        if r["level"] in ("store", "item", "category"):
            d = d[d["store_id"] == r["store_id"]]
        if r["level"] == "item":
            d = d[d["item_id"] == r["item_id"]]
        if r["level"] == "category":
            d = d[d["cat_id"] == r["item_id"]]
        cov.append(coverage(d))
    return mt.assign(coverage_80=cov)


# ---------------------------------------------------------------------------
# Variables exogènes du mois à prédire
# ---------------------------------------------------------------------------
def calendar_for_month(period):
    """({état: part de jours SNAP}, has_event) lus dans calendar.csv, ou None si mois incomplet."""
    if not os.path.exists(CALENDAR_FILE):
        return None
    cal = pd.read_csv(CALENDAR_FILE, parse_dates=["date"])
    cal = cal[cal["date"].dt.to_period("M") == period]
    if len(cal) < period.days_in_month:      # calendrier ne couvrant pas tout le mois
        return None
    snap = {state: cal[col].mean() for state, col in STATE_SNAP.items()}
    return snap, int(cal["event_name_1"].notna().any())


# ---------------------------------------------------------------------------
# 7. Prévision des 3 PROCHAINS mois (origine = dernier mois complet connu)
# ---------------------------------------------------------------------------
def forecast_next_months(monthly, models_by_h):
    step("Prévision des 3 mois suivants (un modèle direct par horizon)...")
    last_month = monthly["month_period"].max()
    item_codes = monthly.drop_duplicates("item_id").set_index("item_id")["item_code"]
    store_codes = monthly.drop_duplicates("store_id").set_index("store_id")["store_code"]
    cat_codes = monthly.drop_duplicates("cat_id").set_index("cat_id")["cat_code"]

    frames = []
    for h in HORIZONS:
        target = last_month + h
        cal = calendar_for_month(target)
        source = "calendar.csv" if cal else "même mois l'année précédente"
        print(f"   M+{h} = {target} (SNAP / événements : {source})")
        rows = []
        for (store, item, cat), g_series in monthly.groupby(["store_id", "item_id", "cat_id"]):
            g_sorted = g_series.sort_values("date")
            s = g_sorted["sales"]
            if len(s) < 24:
                continue
            # Même mois l'année précédente : sert de lag_12 et de repli pour SNAP / événements
            ref = g_sorted[g_sorted["month_period"] == target - 12].iloc[0]
            rows.append({
                "month": str(target), "horizon": h,
                "store_id": store, "item_id": item, "cat_id": cat,
                # Lags = vraies ventes des derniers mois connus (aucune prévision réinjectée)
                "lag_1": s.iloc[-1], "lag_2": s.iloc[-2], "lag_3": s.iloc[-3],
                "lag_6": s.iloc[-6], "lag_12": ref["sales"],
                "rolling_mean_3": s.iloc[-3:].mean(),
                "rolling_mean_6": s.iloc[-6:].mean(),
                "month_num": target.month, "year": target.year,
                "sell_price": g_sorted["sell_price"].iloc[-1],   # dernier prix connu
                "snap": cal[0][store.split("_")[0]] if cal else ref["snap"],
                "has_event": cal[1] if cal else int(ref["has_event"]),
                "item_code": item_codes[item], "store_code": store_codes[store],
                "cat_code": cat_codes[cat],
                # Historique du même mois, pour la comparaison M-12 / M-24 en console
                "same_month_m12": ref["sales"],
                "same_month_m24": g_sorted.loc[g_sorted["month_period"] == target - 24, "sales"].sum(),
            })
        d = pd.DataFrame(rows)
        d["y_pred"], d["y_low"], d["y_high"] = predict_all(models_by_h[h], d[MONTHLY_FEATURES])
        frames.append(d)
    allh = pd.concat(frames, ignore_index=True)
    allh["sell_price"] = allh["sell_price"].round(2)
    allh["origin"] = str(last_month)

    # Format long : 1 ligne par série × horizon (Power BI, analyses)
    long_cols = ["origin", "horizon", "month", "store_id", "item_id", "cat_id",
                 "y_pred", "y_low", "y_high", "sell_price"]
    allh[long_cols].to_csv(os.path.join(OUTPUT_DIR, "monthly_forecast_horizons.csv"), index=False)
    print(f"   -> {os.path.join(OUTPUT_DIR, 'monthly_forecast_horizons.csv')}")

    # CSV lu par la macro VBA ImporterPrevisionMensuelle. Les 5 premières colonnes
    # sont fixes (month, store_id, item_id, cat_id, y_pred) ; suivent le prix
    # (CA estimé), les bornes de l'intervalle à 80 % et les prévisions M+2 / M+3.
    keys = ["store_id", "item_id"]
    nxt = allh[allh["horizon"] == 1].copy()
    for h in HORIZONS[1:]:
        nxt = nxt.merge(allh.loc[allh["horizon"] == h, keys + ["y_pred"]]
                            .rename(columns={"y_pred": f"y_pred_m{h}"}), on=keys)
    cols = ["month", "store_id", "item_id", "cat_id", "y_pred", "sell_price",
            "y_low", "y_high", "y_pred_m2", "y_pred_m3"]
    out = nxt.sort_values(["store_id", "cat_id", "y_pred"], ascending=[True, True, False])[cols]
    out.to_csv(os.path.join(OUTPUT_DIR, "monthly_forecast_next.csv"), index=False)
    print(f"   -> {os.path.join(OUTPUT_DIR, 'monthly_forecast_next.csv')}")

    # Aussi en Excel pour lecture directe (2 onglets)
    xlsx = os.path.join(OUTPUT_DIR, "monthly_forecast_next.xlsx")
    months = {h: str(last_month + h) for h in HORIZONS}
    with pd.ExcelWriter(xlsx, engine="openpyxl") as w:
        excel = out.assign(ca_estime=(out["y_pred"] * out["sell_price"]).round(2)).rename(
            columns={"y_pred_m2": f"y_pred_{months[2]}", "y_pred_m3": f"y_pred_{months[3]}"})
        excel.to_excel(w, sheet_name="Previsions_Mois_Suivant", index=False)
        pivot = (allh.pivot_table(index=["store_id", "cat_id", "item_id"], columns="month",
                                  values="y_pred").reset_index())
        pivot.columns.name = None
        pivot.to_excel(w, sheet_name="Horizons_3_mois", index=False)
        for ws in w.sheets.values():
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
            for col in "ABCDEFGHIJK":
                ws.column_dimensions[col].width = 14
    print(f"   -> {xlsx}")
    return allh, last_month


def print_summary(allh, last_month, df_daily):
    """Résumé console : volumes par catégorie / store et comparaison historique."""
    step("Résumé de la prévision")
    m1 = allh[allh["horizon"] == 1]
    print(f"   Séries prédites : {len(m1)}")
    print(f"   Mois prédits    : " + ", ".join(f"M+{h} = {last_month + h}" for h in HORIZONS))

    print("\n   Total unités prévues par catégorie et par store (M+1) :")
    pivot = m1.pivot_table(index="cat_id", columns="store_id", values="y_pred",
                           aggfunc="sum", margins=True, margins_name="TOTAL")
    print(pivot.round(0).astype(int).to_string())

    print("\n   Total par catégorie et par horizon (prévision [borne basse - borne haute]) :")
    for cat, d in allh.groupby("cat_id"):
        parts = [f"{last_month + h}: {x.y_pred.sum():>7,.0f} [{x.y_low.sum():,.0f}-{x.y_high.sum():,.0f}]"
                 for h, x in d.groupby("horizon")]
        print(f"   {cat:<10} " + "   ".join(parts))

    # Comparaison avec le même mois les années précédentes (vraies ventes)
    print(f"\n   Comparaison M+1 avec le même mois les années précédentes (unités) :")
    comp = m1.groupby("cat_id")[["y_pred", "same_month_m12", "same_month_m24"]].sum()
    comp.loc["TOTAL"] = comp.sum()
    comp["vs M-12"] = (comp["y_pred"] / comp["same_month_m12"] - 1) * 100
    comp["vs M-24"] = (comp["y_pred"] / comp["same_month_m24"] - 1) * 100
    print(f"   {'':<10} {'Prévu':>9} {'M-12':>9} {'M-24':>9} {'vs M-12':>9} {'vs M-24':>9}")
    for cat, r in comp.iterrows():
        print(f"   {cat:<10} {r['y_pred']:>9,.0f} {r['same_month_m12']:>9,.0f} "
              f"{r['same_month_m24']:>9,.0f} {r['vs M-12']:>+8.1f}% {r['vs M-24']:>+8.1f}%")

    # Contrôle de cohérence : si le mois M+1 est partiellement observé dans les
    # données journalières, on extrapole ces jours au mois entier et on compare.
    nxt = last_month + 1
    part = df_daily[df_daily["date"].dt.to_period("M") == nxt]
    if not part.empty:
        n_days = part["date"].nunique()
        extrap = part["sales"].sum() / n_days * nxt.days_in_month
        prev = m1["y_pred"].sum()
        print(f"\n   Contrôle : {n_days} jours de {nxt} sont déjà observés. "
              f"Extrapolés au mois complet : {extrap:,.0f} unités "
              f"vs {prev:,.0f} prévues ({(prev / extrap - 1) * 100:+.1f} %).")


def main():
    print("=" * 70)
    print(" PREVISION MENSUELLE - LightGBM direct M+1..M+3 (sans récursivité)")
    print("=" * 70)

    # 1. Chargement et agrégation
    step("Chargement de series_long.csv...")
    df = load_long()
    base = build_monthly(df.copy())

    # 3. Découpage : les 7 derniers mois = test (identique pour tous les horizons)
    months = np.sort(base["month_period"].unique())
    cutoff_month = months[-N_TEST_MONTHS]
    print(f"   Test : {cutoff_month} -> {months[-1]} ({N_TEST_MONTHS} mois)")

    models_by_h, fc_all, metrics = {}, [], []
    for h in HORIZONS:
        # 2. Features de l'horizon h  /  4. Entraînement point + quantiles
        step(f"Horizon M+{h} : features, entraînement (point + quantiles 10/90 %), test...")
        monthly = add_monthly_features(base, h)
        train = monthly[monthly.month_period < cutoff_month].dropna(subset=MONTHLY_FEATURES)
        test = monthly[monthly.month_period >= cutoff_month].copy()
        q = calibrate(train)                    # correction de largeur de l'intervalle
        point, quant = make_models()
        point.fit(train[MONTHLY_FEATURES], train["sales"])
        for model in quant.values():
            model.fit(train[MONTHLY_FEATURES], train["sales"])
        models_by_h[h] = (point, quant, q)
        if h == 1:
            featured = monthly                  # features M+1 (codes produits / stores)
        if h == 1:
            joblib.dump(point, os.path.join(OUTPUT_DIR, "model_monthly.joblib"))

        # 5. Prévision sur la période de test
        test["y_pred"], test["y_low"], test["y_high"] = predict_all(models_by_h[h], test[MONTHLY_FEATURES])
        fc = test.rename(columns={"sales": "y_true"})[
            ["date", "store_id", "item_id", "cat_id", "y_true", "y_pred", "y_low", "y_high"]]
        fc.insert(4, "horizon", h)
        fc_all.append(fc)

        # 6. Métriques (niveau catégorie inclus grâce à cat_id) + couverture de l'intervalle
        mt = add_coverage(metrics_table(fc, f"LightGBM M+{h}", "Mensuel"), fc)
        metrics.append(mt.assign(horizon=h))
        g = mt[mt.level == "global"].iloc[0]
        print(f"   {len(train)} lignes d'entraînement | RMSE={g.RMSE:.1f}  MAE={g.MAE:.1f}  "
              f"WAPE={g.WAPE:.1f}%  couverture 80 % = {g.coverage_80:.1f}% "
              f"(calibration q={q:+.2f})")

        # Référence : naïf saisonnier (même mois l'année précédente = lag_12)
        if h == 1:
            naive = fc.assign(y_pred=test["lag_12"].values)
            metrics.append(metrics_table(naive, "Naïf M-12", "Mensuel").assign(horizon=1))

    fc_all = pd.concat(fc_all, ignore_index=True)
    fc_all.to_csv(os.path.join(OUTPUT_DIR, "forecasts_monthly.csv"), index=False)
    mt_all = pd.concat(metrics, ignore_index=True)
    mt_all.to_csv(os.path.join(OUTPUT_DIR, "metrics_monthly.csv"), index=False)
    print(f"\n   -> {os.path.join(OUTPUT_DIR, 'forecasts_monthly.csv')}")
    print(f"   -> {os.path.join(OUTPUT_DIR, 'metrics_monthly.csv')}")

    # Tableau récapitulatif RMSE par catégorie et par modèle
    print("\n   RMSE sur les 7 mois de test (unités / mois / série) :")
    view = mt_all[mt_all["level"].isin(["global", "category"])]
    view = view.assign(niveau=np.where(view["level"] == "global", "GLOBAL",
                                       view["store_id"] + " " + view["item_id"]))
    print(view.pivot_table(index="niveau", columns="model", values="RMSE", sort=False)
              .round(1).to_string())
    g1 = mt_all[(mt_all.level == "global") & (mt_all.model == "LightGBM M+1")].iloc[0]
    gn = mt_all[(mt_all.level == "global") & (mt_all.model == "Naïf M-12")].iloc[0]
    print(f"\n   -> gain LightGBM M+1 vs naïf M-12 : {(1 - g1.RMSE / gn.RMSE) * 100:+.1f} % en RMSE")

    # 7. Prévision des prochains mois + résumé
    allh, last_month = forecast_next_months(featured, models_by_h)
    print_summary(allh, last_month, df)
    print("\nPrévision mensuelle terminée avec succès.")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"\n[ERREUR] monthly_forecast : {e}")
        traceback.print_exc()
        sys.exit(1)
