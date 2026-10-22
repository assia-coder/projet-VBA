"""
train_active_ml.py - Modèle LightGBM "actif".

Features : toutes les features passives + variables externes
(prix, variation de prix, prix relatif, SNAP de l'état, événements,
jours avant Noël, semaine de fête).

Découpage : identique au passif (1 733 jours train / 180 jours test).

Sorties (output/) :
  - forecasts_active_ml.csv
  - metrics_active_ml.csv
  - model_active.joblib : modèle réentraîné sur tout l'historique + ordre des
                          catégories, utilisé par predict_next_30_days.py
"""
import os
import sys
import traceback

import joblib
import numpy as np

from utils import (ACTIVE_EXTRA_FEATURES, OUTPUT_DIR, PASSIVE_FEATURES, add_active_features,
                   add_passive_features, load_long, metrics_table,
                   print_store_metrics, split_date, step)

FEATURES = PASSIVE_FEATURES + ACTIVE_EXTRA_FEATURES

FORECAST_FILE = os.path.join(OUTPUT_DIR, "forecasts_active_ml.csv")
METRICS_FILE = os.path.join(OUTPUT_DIR, "metrics_active_ml.csv")
MODEL_FILE = os.path.join(OUTPUT_DIR, "model_active.joblib")


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
    final.fit(full[FEATURES], full["sales"])

    item_categories = sorted(df["item_id"].unique().tolist())
    store_categories = sorted(df["store_id"].unique().tolist())
    bundle = {
        "model": final,
        "model_name": "LightGBM",
        "approach": "Actif",
        "features": FEATURES,
        "last_date": str(df["date"].max().date()),
        "item_categories": item_categories,
        "store_categories": store_categories,
    }
    joblib.dump(bundle, MODEL_FILE)
    print(f"   Modèle sauvegardé : {os.path.basename(MODEL_FILE)} "
          f"({len(full):,} lignes, jusqu'au {bundle['last_date']})")


def main():
    print("=" * 70)
    print(" ENTRAÎNEMENT LIGHTGBM ACTIF")
    print("=" * 70)

    try:
        from lightgbm import LGBMRegressor
    except ImportError:
        raise ImportError("lightgbm n'est pas installé : pip install lightgbm")

    # 1. Chargement
    step("Chargement de series_long.csv...")
    df = load_long()
    print(f"   {len(df):,} lignes, {df.groupby(['store_id', 'item_id']).ngroups} séries")

    # 2. Features passives + actives
    step("Création des features passives et externes...")
    df = add_passive_features(df)
    df = add_active_features(df)
    print(f"   {len(FEATURES)} features : {', '.join(ACTIVE_EXTRA_FEATURES)} + passives")

    # 3. Découpage train / test
    cutoff = split_date(df)
    train = df[df["date"] < cutoff].dropna(subset=PASSIVE_FEATURES)
    test = df[df["date"] >= cutoff].copy()
    print(f"   Train : {train['date'].min().date()} -> {train['date'].max().date()} ({len(train):,} lignes)")
    print(f"   Test  : {test['date'].min().date()} -> {test['date'].max().date()} ({len(test):,} lignes)")

    # 4. Entraînement
    step("Entraînement LightGBM...")
    model = LGBMRegressor(
        n_estimators=800, learning_rate=0.03, num_leaves=63,
        subsample=0.8, subsample_freq=1, colsample_bytree=0.8,
        objective="regression", random_state=42, verbose=-1,
    )
    model.fit(train[FEATURES], train["sales"])

    # 5. Prévision
    step("Prévision sur la période de test...")
    test["y_pred"] = np.clip(model.predict(test[FEATURES]), 0, None)
    fc = test.rename(columns={"sales": "y_true"})[["date", "store_id", "item_id", "y_true", "y_pred"]]
    fc["y_pred"] = fc["y_pred"].round(3)
    fc.to_csv(FORECAST_FILE, index=False)
    print(f"   -> {FORECAST_FILE}")

    # 6. Métriques
    step("Calcul des métriques...")
    mt = metrics_table(fc, model="LightGBM", approach="Actif")
    mt.to_csv(METRICS_FILE, index=False)
    print_store_metrics(mt)
    print(f"   -> {METRICS_FILE}")

    # 7. Sauvegarde du modèle pour la prévision J+1..J+30
    save_bundle(model, df)

    print("\nLightGBM actif terminé avec succès.")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"\n[ERREUR] train_active_ml : {e}")
        traceback.print_exc()
        sys.exit(1)
