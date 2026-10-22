"""
train_passive_ml.py - Modèle XGBoost "passif".

Features : uniquement l'historique des ventes (lags, statistiques glissantes),
le calendrier et l'identité de la série. Aucune variable externe.

Découpage : 1 733 premiers jours = entraînement, 180 derniers = test.

Sorties (output/) :
  - forecasts_passive_ml.csv     : date, store_id, item_id, y_true, y_pred
  - metrics_passive_ml.csv       : MAE, MAPE, RMSE (global, par store, par item)
  - model_passive.joblib         : modèle réentraîné sur tout l'historique + ordre des
                                   catégories, utilisé par predict_next_30_days.py
"""
import os
import sys
import traceback

import joblib
import numpy as np
import pandas as pd

from utils import (OUTPUT_DIR, PASSIVE_FEATURES, add_passive_features, load_long,
                   metrics_table, print_store_metrics, split_date, step)

FORECAST_FILE = os.path.join(OUTPUT_DIR, "forecasts_passive_ml.csv")
METRICS_FILE = os.path.join(OUTPUT_DIR, "metrics_passive_ml.csv")
MODEL_FILE = os.path.join(OUTPUT_DIR, "model_passive.joblib")


def save_bundle(model, df):
    """
    Réentraîne le modèle (mêmes hyperparamètres) sur TOUT l'historique, puis le sauvegarde.
    Le modèle évalué ci-dessus n'a jamais vu les 180 derniers jours : pour prévoir
    l'avenir, on utilise toutes les données disponibles.

    L'ordre des catégories est figé : add_passive_features() crée item_code/store_code
    via .astype("category").cat.codes, dont l'ordre (alphabétique) dépend des valeurs
    présentes dans df. Si de nouveaux produits sont ajoutés plus tard, cet ordre peut
    changer : predict_next_30_days.py recrée donc les codes à partir de ces listes.
    """
    step("Réentraînement sur tout l'historique et sauvegarde du modèle...")
    full = df.dropna(subset=PASSIVE_FEATURES)
    final = type(model)(**model.get_params())
    final.fit(full[PASSIVE_FEATURES], full["sales"])

    item_categories = sorted(df["item_id"].unique().tolist())
    store_categories = sorted(df["store_id"].unique().tolist())
    bundle = {
        "model": final,
        "model_name": "XGBoost",
        "approach": "Passif",
        "features": PASSIVE_FEATURES,
        "last_date": str(df["date"].max().date()),
        "item_categories": item_categories,
        "store_categories": store_categories,
    }
    joblib.dump(bundle, MODEL_FILE)
    print(f"   Modèle sauvegardé : {os.path.basename(MODEL_FILE)} "
          f"({len(full):,} lignes, jusqu'au {bundle['last_date']})")


def main():
    print("=" * 70)
    print(" ENTRAÎNEMENT XGBOOST PASSIF")
    print("=" * 70)

    # Import ici pour donner un message clair si la librairie manque
    try:
        from xgboost import XGBRegressor
    except ImportError:
        raise ImportError("xgboost n'est pas installé : pip install xgboost")

    # 1. Chargement des données prétraitées
    step("Chargement de series_long.csv...")
    df = load_long()
    print(f"   {len(df):,} lignes, {df.groupby(['store_id', 'item_id']).ngroups} séries")

    # 2. Construction des features passives
    step("Création des features passives (lags, rolling, calendrier)...")
    df = add_passive_features(df)

    # 3. Découpage train / test
    cutoff = split_date(df)
    train = df[df["date"] < cutoff].dropna(subset=PASSIVE_FEATURES)
    test = df[df["date"] >= cutoff].copy()
    print(f"   Train : {train['date'].min().date()} -> {train['date'].max().date()} ({len(train):,} lignes)")
    print(f"   Test  : {test['date'].min().date()} -> {test['date'].max().date()} ({len(test):,} lignes)")

    # 4. Entraînement
    step("Entraînement XGBoost...")
    model = XGBRegressor(
        n_estimators=600, learning_rate=0.03, max_depth=7,
        subsample=0.8, colsample_bytree=0.8, objective="reg:squarederror",
        importance_type="gain", n_jobs=-1, random_state=42,
    )
    model.fit(train[PASSIVE_FEATURES], train["sales"])

    # 5. Prévision sur les 180 jours de test (ventes négatives impossibles)
    step("Prévision sur la période de test...")
    test["y_pred"] = np.clip(model.predict(test[PASSIVE_FEATURES]), 0, None)
    fc = test.rename(columns={"sales": "y_true"})[["date", "store_id", "item_id", "y_true", "y_pred"]]
    fc["y_pred"] = fc["y_pred"].round(3)
    fc.to_csv(FORECAST_FILE, index=False)
    print(f"   -> {FORECAST_FILE}")

    # 6. Métriques
    step("Calcul des métriques...")
    mt = metrics_table(fc, model="XGBoost", approach="Passif")
    mt.to_csv(METRICS_FILE, index=False)
    print_store_metrics(mt)
    print(f"   -> {METRICS_FILE}")

    # 7. Sauvegarde du modèle pour la prévision J+1..J+30
    save_bundle(model, df)

    print("\nXGBoost passif terminé avec succès.")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"\n[ERREUR] train_passive_ml : {e}")
        traceback.print_exc()
        sys.exit(1)
