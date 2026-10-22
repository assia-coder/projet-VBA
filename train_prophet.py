"""
train_prophet.py - Prophet en deux modes sur chaque série (magasin × produit).

  - Mode passif : ds + y uniquement, saisonnalités hebdomadaire et annuelle.
  - Mode actif  : + sell_price, snap, has_event en add_regressor.

Découpage : identique aux modèles ML (1 733 jours train / 180 jours test).

Sorties (output/) :
  - forecasts_prophet_passive.csv
  - forecasts_prophet_active.csv
  - metrics_prophet.csv (les deux modes)
  - forecast_prophet_future_30d.csv : J+1..J+30 après la dernière date connue,
    modèles réentraînés sur tout le dataset (colonne approach = Passif / Actif)
"""
import logging
import os
import sys
import time
import traceback

import numpy as np
import pandas as pd

from utils import (N_TRAIN_DAYS, OUTPUT_DIR, PROPHET_FILE, future_calendar, metrics_table,
                   print_store_metrics, snap_for_store, step)

REGRESSORS = ["sell_price", "snap", "has_event"]

PASSIVE_FILE = os.path.join(OUTPUT_DIR, "forecasts_prophet_passive.csv")
ACTIVE_FILE = os.path.join(OUTPUT_DIR, "forecasts_prophet_active.csv")
METRICS_FILE = os.path.join(OUTPUT_DIR, "metrics_prophet.csv")
FUTURE_FILE = os.path.join(OUTPUT_DIR, "forecast_prophet_future_30d.csv")
FUTURE_HORIZON = 30

# Prophet / cmdstanpy sont très bavards : on coupe leurs logs
logging.getLogger("prophet").setLevel(logging.ERROR)
logging.getLogger("cmdstanpy").setLevel(logging.ERROR)


def fit_predict(train, test, regressors):
    """Entraîne un Prophet sur une série et prévoit la période de test."""
    from prophet import Prophet
    m = Prophet(weekly_seasonality=True, yearly_seasonality=True, daily_seasonality=False)
    for r in regressors:
        m.add_regressor(r)
    m.fit(train[["ds", "y"] + regressors])
    pred = m.predict(test[["ds"] + regressors])
    return np.clip(pred["yhat"].values, 0, None)


def main():
    print("=" * 70)
    print(" ENTRAÎNEMENT PROPHET (passif + actif)")
    print("=" * 70)

    try:
        import prophet  # noqa: F401
    except ImportError:
        raise ImportError("prophet n'est pas installé : pip install prophet")
    # cmdstanpy reconfigure son logger à l'import : on le coupe ensuite
    logging.getLogger("cmdstanpy").disabled = True

    # 1. Chargement
    step("Chargement de series_prophet.csv...")
    if not os.path.exists(PROPHET_FILE):
        raise FileNotFoundError(f"{PROPHET_FILE} introuvable. Lancez d'abord : python preprocess.py")
    df = pd.read_csv(PROPHET_FILE, parse_dates=["ds"])
    series = df.groupby(["store_id", "item_id"])
    print(f"   {len(df):,} lignes, {series.ngroups} séries")

    # 2. Date de coupure train / test
    dates = np.sort(df["ds"].unique())
    cutoff = pd.Timestamp(dates[N_TRAIN_DAYS])
    print(f"   Test à partir du {cutoff.date()} ({len(dates) - N_TRAIN_DAYS} jours)")

    # 3. Boucle sur les séries
    step("Entraînement série par série...")
    out = {"passive": [], "active": []}
    n_fail = 0
    t0 = time.time()
    for i, ((store, item), s) in enumerate(series, start=1):
        train, test = s[s["ds"] < cutoff], s[s["ds"] >= cutoff]
        base = pd.DataFrame({"date": test["ds"].values, "store_id": store,
                             "item_id": item, "y_true": test["y"].values})
        for mode, regs in (("passive", []), ("active", REGRESSORS)):
            try:
                fc = base.copy()
                fc["y_pred"] = fit_predict(train, test, regs).round(3)
                out[mode].append(fc)
            except Exception as e:
                n_fail += 1
                print(f"   [AVERTISSEMENT] {store}/{item} mode {mode} échoué : {e}")
        print(f"   [{i:>2}/{series.ngroups}] {store} / {item}  ({time.time() - t0:.0f}s)", flush=True)

    if not out["passive"] or not out["active"]:
        raise RuntimeError("Aucune série n'a pu être modélisée par Prophet.")

    # 4. Sauvegarde des prévisions
    step("Écriture des prévisions...")
    fc_passive = pd.concat(out["passive"], ignore_index=True)
    fc_active = pd.concat(out["active"], ignore_index=True)
    fc_passive.to_csv(PASSIVE_FILE, index=False)
    fc_active.to_csv(ACTIVE_FILE, index=False)
    print(f"   -> {PASSIVE_FILE}\n   -> {ACTIVE_FILE}")

    # 5. Métriques des deux modes dans un seul fichier
    step("Calcul des métriques...")
    mt_p = metrics_table(fc_passive, model="Prophet", approach="Passif")
    mt_a = metrics_table(fc_active, model="Prophet", approach="Actif")
    print("   Prophet passif :")
    print_store_metrics(mt_p)
    print("   Prophet actif :")
    print_store_metrics(mt_a)
    pd.concat([mt_p, mt_a], ignore_index=True).to_csv(METRICS_FILE, index=False)
    print(f"   -> {METRICS_FILE}")

    if n_fail:
        print(f"\n   {n_fail} ajustement(s) en échec (voir avertissements ci-dessus).")

    # 6. Prévision des 30 prochains jours (Prophet trop lourd à sérialiser :
    #    on stocke directement ses prévisions futures)
    forecast_future(series, pd.Timestamp(dates[-1]))
    print("\nProphet terminé avec succès.")


def forecast_future(series, last_date):
    """
    Réentraîne Prophet sur TOUT le dataset (sans découpage train/test), dans les deux
    modes, et prévoit les 30 jours suivant la dernière date connue.
    Régresseurs futurs : prix = dernier prix connu ; SNAP et événements = calendar.csv
    (connus à l'avance). Sortie : forecast_prophet_future_30d.csv (colonne approach).
    """
    from prophet import Prophet
    step(f"Prévision des {FUTURE_HORIZON} prochains jours avec Prophet (tout le dataset)...")
    future_dates = pd.date_range(last_date + pd.Timedelta(days=1), periods=FUTURE_HORIZON, freq="D")
    cal = future_calendar(future_dates)

    rows = []
    for (store, item), s in series:
        future_df = pd.DataFrame({"ds": future_dates})
        future_df["sell_price"] = float(s["sell_price"].iloc[-1])
        future_df["snap"] = snap_for_store(cal, store).values
        future_df["has_event"] = cal["has_event"].values
        for approach, regs in (("Passif", []), ("Actif", REGRESSORS)):
            try:
                m_full = Prophet(weekly_seasonality=True, yearly_seasonality=True,
                                 daily_seasonality=False)
                for r in regs:
                    m_full.add_regressor(r)
                m_full.fit(s[["ds", "y"] + regs])
                preds = m_full.predict(future_df[["ds"] + regs])
                y_pred = np.clip(preds["yhat"].values, 0, None)
                y_low = np.clip(preds["yhat_lower"].values, 0, None)
                y_high = np.clip(preds["yhat_upper"].values, 0, None)
                for j, (d, yp, yl, yh) in enumerate(zip(future_dates, y_pred, y_low, y_high), 1):
                    rows.append({
                        "date": d.date(), "store_id": store, "item_id": item,
                        "horizon_days": j, "approach": approach,
                        "y_pred": round(float(yp), 3),
                        "y_low": round(float(yl), 3),
                        "y_high": round(float(yh), 3),
                    })
            except Exception as e:
                print(f"   [AVERT] Prophet futur {store}/{item} ({approach}) : {e}")

    if rows:
        pd.DataFrame(rows).to_csv(FUTURE_FILE, index=False)
        print(f"   {future_dates[0].date()} -> {future_dates[-1].date()}")
        print(f"   -> {FUTURE_FILE}")
    else:
        print("   [AVERT] aucune prévision future Prophet n'a pu être calculée.")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"\n[ERREUR] train_prophet : {e}")
        traceback.print_exc()
        sys.exit(1)
