"""
predict_next_30_days.py
=======================
Prédit les 30 prochains jours avec le modèle gagnant identifié par compare_all.py.

- Lit output/winner.json pour savoir quel modèle utiliser.
- Si le gagnant est XGBoost ou LightGBM : charge le .joblib et prédit directement
  chaque horizon J+1..J+30 (approche directe, SANS récursivité : aucune prévision
  n'est jamais réinjectée comme feature).
  Construction des features de la date cible J+h :
    * calendrier (day_of_week, month, year, week_of_year, is_weekend) : date cible ;
    * lags / moyennes glissantes : valeur réelle si elle pointe sur un jour connu
      (ex. lag_28 et lag_35 jusqu'à J+28), sinon état figé au DERNIER JOUR RÉEL ;
    * variables externes (approche active) : SNAP et événements lus dans
      calendar.csv (connus à l'avance), prix = dernier prix connu.
- Si le gagnant est Prophet : reprend forecast_prophet_future_30d.csv (calculé par
  train_prophet.py sur tout le dataset), filtré sur l'approche gagnante.

Usage :
    python predict_next_30_days.py

Sorties :
    output/forecast_next_30_days.csv
        colonnes : date, store_id, item_id, cat_id, horizon_days,
                   y_pred, y_low, y_high, model, approach
    output/status_prediction.txt
        SUCCESS ou ERROR + message
"""
import json
import os
import sys
import traceback

import joblib
import numpy as np
import pandas as pd

from utils import (OUTPUT_DIR, add_active_features, add_passive_features, future_calendar,
                   load_long, snap_for_store, step)

WINNER_FILE = os.path.join(OUTPUT_DIR, "winner.json")
RESULT_FILE = os.path.join(OUTPUT_DIR, "forecast_next_30_days.csv")
STATUS_FILE = os.path.join(OUTPUT_DIR, "status_prediction.txt")
PROPHET_FUTURE_FILE = os.path.join(OUTPUT_DIR, "forecast_prophet_future_30d.csv")
HORIZON = 30
CONF_PCT = 0.12   # intervalle indicatif ±12 % (modèles ML)
KEYS = ["store_id", "item_id"]
OUT_COLS = ["date", "store_id", "item_id", "cat_id", "horizon_days",
            "y_pred", "y_low", "y_high", "model", "approach"]


def encode_codes(df, item_categories, store_categories):
    """
    Recrée item_code et store_code avec l'ordre FIGÉ lors de l'entraînement.
    Si un produit inconnu est présent, lève une ValueError claire.
    """
    item_map = {v: i for i, v in enumerate(item_categories)}
    store_map = {v: i for i, v in enumerate(store_categories)}
    df["item_code"] = df["item_id"].map(item_map)
    df["store_code"] = df["store_id"].map(store_map)
    unknown = df["item_code"].isna() | df["store_code"].isna()
    if unknown.any():
        unknowns = sorted((df.loc[unknown, "store_id"] + "/" + df.loc[unknown, "item_id"]).unique())
        raise ValueError(
            f"Produits inconnus du modèle : {unknowns}. "
            "Réentraîner (commande 'all') après ajout de nouveaux produits.")
    return df


def future_rows(df, last_date):
    """
    Lignes futures (ventes inconnues = NaN) pour chaque série, du lendemain de sa
    dernière date jusqu'à last_date + HORIZON. Prix = dernier prix connu de la série,
    SNAP / événements = calendar.csv.
    """
    end = last_date + pd.Timedelta(days=HORIZON)
    latest = df.groupby(KEYS).tail(1)
    parts = []
    for _, r in latest.iterrows():
        dates = pd.date_range(r["date"] + pd.Timedelta(days=1), end, freq="D")
        parts.append(pd.DataFrame({
            "date": dates, "store_id": r["store_id"], "item_id": r["item_id"],
            "cat_id": r["cat_id"], "dept_id": r.get("dept_id", ""),
            "sales": np.nan, "sell_price": r["sell_price"],
        }))
    fut = pd.concat(parts, ignore_index=True)
    cal = future_calendar(fut["date"])
    fut["has_event"] = fut["date"].map(cal["has_event"]).astype(int)
    fut["event_type"] = fut["date"].map(cal["event_type"])
    fut["event_name"] = fut["date"].map(cal["event_name"])
    fut["snap"] = 0
    for store in fut["store_id"].unique():
        m = fut["store_id"] == store
        fut.loc[m, "snap"] = fut.loc[m, "date"].map(snap_for_store(cal, store)).astype(int).values
    return fut


def predict_ml(bundle, df):
    """Prévision directe J+1..J+30 avec un modèle XGBoost / LightGBM sauvegardé."""
    model = bundle["model"]
    features = bundle["features"]
    approach = bundle["approach"]
    last_date = df["date"].max()

    # 1. Historique + lignes futures, features calculées d'un seul coup
    fut = future_rows(df, last_date)
    ext = pd.concat([df, fut], ignore_index=True)
    ext = add_passive_features(ext)
    if approach == "Actif":
        ext = add_active_features(ext)
    ext = encode_codes(ext, bundle["item_categories"], bundle["store_categories"])

    # 2. État de chaque série au dernier jour réel
    is_future = ext["sales"].isna()
    known = ext[~is_future]
    state = known.groupby(KEYS).tail(1).set_index(KEYS)[features]
    if state.drop(columns=["item_code", "store_code"], errors="ignore").isna().all(axis=1).any():
        raise ValueError(f"Historique insuffisant pour certaines séries au {last_date.date()}. "
                         "Vérifier que series_long.csv est à jour.")

    # 3. Lags pointant sur un jour futur (NaN) -> valeur figée au dernier jour réel.
    #    Jamais une prévision : pas de récursivité.
    pred_df = ext[is_future & (ext["date"] > last_date)].copy()
    idx = pd.MultiIndex.from_frame(pred_df[KEYS])
    for f in features:
        if pred_df[f].isna().any():
            fill = pd.Series(state[f].reindex(idx).values, index=pred_df.index)
            pred_df[f] = pred_df[f].fillna(fill)
    pred_df = pred_df.dropna(subset=features)
    if pred_df.empty:
        raise ValueError(f"Aucune série prévisible au {last_date.date()}.")

    # 4. Prévision (ventes négatives impossibles)
    y = np.clip(model.predict(pred_df[features]), 0, None)
    result = pd.DataFrame({
        "date": pred_df["date"].dt.date.values,
        "store_id": pred_df["store_id"].values,
        "item_id": pred_df["item_id"].values,
        "cat_id": pred_df["cat_id"].values,
        "horizon_days": (pred_df["date"] - last_date).dt.days.values,
        "y_pred": np.round(y, 3),
        "y_low": np.round(y * (1 - CONF_PCT), 3),
        "y_high": np.round(y * (1 + CONF_PCT), 3),
    })
    result["model"] = bundle["model_name"]
    result["approach"] = approach
    return result


def predict_prophet(df, approach):
    """Reprend les prévisions Prophet pré-calculées par train_prophet.py."""
    if not os.path.exists(PROPHET_FUTURE_FILE):
        raise FileNotFoundError(f"{PROPHET_FUTURE_FILE} introuvable. Lancez d'abord train_prophet.py")
    fc = pd.read_csv(PROPHET_FUTURE_FILE, parse_dates=["date"])
    if "approach" in fc.columns:
        fc = fc[fc["approach"] == approach].copy()
    if fc.empty:
        raise ValueError(f"Aucune prévision Prophet ({approach}) dans {PROPHET_FUTURE_FILE}.")
    cat_map = df.drop_duplicates("item_id").set_index("item_id")["cat_id"]
    fc["cat_id"] = fc["item_id"].map(cat_map)
    fc["date"] = fc["date"].dt.date
    fc["model"] = "Prophet"
    fc["approach"] = approach
    return fc


def write_status(status, message=""):
    with open(STATUS_FILE, "w", encoding="utf-8") as f:
        f.write(f"{status}\n{message}\n{pd.Timestamp.now():%Y-%m-%d %H:%M:%S}\n")


def main():
    print("=" * 65)
    print("  PRÉVISION J+1..J+30 — MODÈLE GAGNANT")
    print("=" * 65)

    # 1. Quel modèle utiliser ?
    if not os.path.exists(WINNER_FILE):
        raise FileNotFoundError("winner.json introuvable. Lancez compare_all.py d'abord.")
    with open(WINNER_FILE, "r", encoding="utf-8") as f:
        winner = json.load(f)
    print(f"\n  Gagnant : {winner['label']}  RMSE={winner['RMSE']:.4f}")
    print(f"  Classement établi le : {winner['updated_at']}")

    # 2. Données
    step("Chargement de series_long.csv...")
    df = load_long()
    last_date = df["date"].max()
    print(f"  Dernière date connue : {last_date.date()}")
    print(f"  Prévision : {(last_date + pd.Timedelta(days=1)).date()} "
          f"-> {(last_date + pd.Timedelta(days=HORIZON)).date()}")

    # 3. Prévision selon le type de gagnant
    model_file = os.path.join(OUTPUT_DIR, winner["model_file"])
    if winner["model"] in ("XGBoost", "LightGBM"):
        step(f"Chargement du modèle {winner['label']}...")
        if not os.path.exists(model_file):
            raise FileNotFoundError(f"{model_file} introuvable. Réentraîner le modèle.")
        bundle = joblib.load(model_file)
        if bundle["last_date"] != str(last_date.date()):
            print(f"  [INFO] modèle entraîné jusqu'au {bundle['last_date']}, "
                  f"données jusqu'au {last_date.date()}")
        step("Prévision directe J+1..J+30 (sans récursivité)...")
        result = predict_ml(bundle, df.copy())
    elif winner["model"] == "Prophet":
        step("Chargement des prévisions Prophet pré-calculées...")
        result = predict_prophet(df, winner["approach"])
    else:
        raise ValueError(f"Modèle inconnu : {winner['model']}")

    # 4. Sauvegarde
    result = result.reindex(columns=OUT_COLS).sort_values(["store_id", "item_id", "horizon_days"])
    result.to_csv(RESULT_FILE, index=False)
    n_series = result.groupby(KEYS).ngroups
    print(f"\n  {len(result)} lignes ({n_series} séries × {HORIZON} jours)")
    print(f"  -> {RESULT_FILE}")
    print("\n  Total unités prévues par store :")
    print(result.groupby("store_id")["y_pred"].sum().round(0).astype(int).to_string())

    write_status("SUCCESS", f"{winner['label']} — {n_series} séries")
    print("\n  Terminé avec succès.")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"\n[ERREUR] {e}")
        traceback.print_exc()
        write_status("ERROR", str(e))
        sys.exit(1)
