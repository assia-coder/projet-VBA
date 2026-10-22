"""
utils.py - Fonctions partagées par tous les scripts du pipeline.

Contient : chemins, constantes, calcul des métriques, chargement des données
prétraitées et construction des features "passives" (historique).
"""
import os
import sys

import numpy as np
import pandas as pd

# ---------------------------------------------------------------------------
# Encodage console : évite les erreurs d'accents sous Windows (cp1252)
# ---------------------------------------------------------------------------
try:
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
except Exception:
    pass

# ---------------------------------------------------------------------------
# Chemins et constantes du projet
# ---------------------------------------------------------------------------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")
OUTPUT_DIR = os.path.join(BASE_DIR, "output")

STORES = ["CA_1", "TX_1", "WI_1"]
TOP_N_ITEMS = 10          # conservé pour compatibilité
TOP_N_PER_CAT = 5         # top 5 par catégorie par store (5 × 3 catégories × 3 stores = 45 séries)
CATEGORIES = ["FOODS", "HOBBIES", "HOUSEHOLD"]
N_DAYS = 1913          # nombre total de jours dans sales_train_validation
N_TRAIN_DAYS = 1733    # jours d'entraînement
N_TEST_DAYS = N_DAYS - N_TRAIN_DAYS  # 180 jours de test

LONG_FILE = os.path.join(OUTPUT_DIR, "series_long.csv")
CALENDAR_FILE = os.path.join(DATA_DIR, "calendar.csv")
PROPHET_FILE = os.path.join(OUTPUT_DIR, "series_prophet.csv")

# Encodage du type d'événement (demandé dans le cahier des charges)
EVENT_TYPE_CODES = {"none": 0, "Sporting": 1, "Cultural": 2, "National": 3, "Religious": 4}


def step(msg):
    """Affiche une étape de progression bien visible."""
    print(f"\n>>> {msg}", flush=True)


# ---------------------------------------------------------------------------
# Calendrier des dates futures (prévision J+1..J+30)
# ---------------------------------------------------------------------------
STATE_SNAP = {"CA": "snap_CA", "TX": "snap_TX", "WI": "snap_WI"}


def future_calendar(dates):
    """
    Variables externes CONNUES À L'AVANCE pour des dates futures, lues dans calendar.csv :
    SNAP par état, présence et type d'événement. Les dates absentes du calendrier
    reçoivent snap = 0, has_event = 0, event_type = "none".
    Retourne un DataFrame indexé par date : snap_CA, snap_TX, snap_WI, has_event, event_type, event_name.
    """
    idx = pd.DatetimeIndex(pd.to_datetime(dates)).unique()
    out = pd.DataFrame(index=idx)
    out["snap_CA"] = out["snap_TX"] = out["snap_WI"] = 0
    out["has_event"], out["event_type"], out["event_name"] = 0, "none", "none"
    if os.path.exists(CALENDAR_FILE):
        cal = pd.read_csv(CALENDAR_FILE, parse_dates=["date"]).set_index("date")
        known = idx.intersection(cal.index)
        c = cal.loc[known]
        for col in ("snap_CA", "snap_TX", "snap_WI"):
            out.loc[known, col] = c[col].astype(int)
        out.loc[known, "has_event"] = c["event_name_1"].notna().astype(int)
        out.loc[known, "event_type"] = c["event_type_1"].fillna("none")
        out.loc[known, "event_name"] = c["event_name_1"].fillna("none")
    return out


def snap_for_store(cal, store_id):
    """Colonne SNAP de l'état du magasin (CA_1 -> snap_CA)."""
    return cal[STATE_SNAP.get(str(store_id).split("_")[0], "snap_CA")]


# ---------------------------------------------------------------------------
# Chargement des données prétraitées
# ---------------------------------------------------------------------------
def load_long():
    """Charge series_long.csv (produit par preprocess.py)."""
    if not os.path.exists(LONG_FILE):
        raise FileNotFoundError(
            f"{LONG_FILE} introuvable. Lancez d'abord : python preprocess.py"
        )
    df = pd.read_csv(LONG_FILE, parse_dates=["date"])
    df["event_type"] = df["event_type"].fillna("none")
    df["event_name"] = df["event_name"].fillna("none")
    return df.sort_values(["store_id", "item_id", "date"]).reset_index(drop=True)


def split_date(df):
    """Retourne la première date du jeu de test (les 180 derniers jours)."""
    dates = np.sort(df["date"].unique())
    return pd.Timestamp(dates[N_TRAIN_DAYS])


# ---------------------------------------------------------------------------
# Features passives (uniquement basées sur l'historique des ventes)
# ---------------------------------------------------------------------------
def add_passive_features(df):
    """Ajoute lags, moyennes/écarts-types glissants, calendrier et codes."""
    df = df.sort_values(["store_id", "item_id", "date"]).copy()
    g = df.groupby(["store_id", "item_id"])["sales"]

    # Lags : ventes observées il y a k jours
    for lag in (7, 14, 28, 35):
        df[f"lag_{lag}"] = g.shift(lag)

    # Statistiques glissantes calculées sur lag_28 (pas de fuite d'information)
    g28 = df.groupby(["store_id", "item_id"])["lag_28"]
    df["rolling_mean_7"] = g28.transform(lambda s: s.rolling(7).mean())
    df["rolling_mean_28"] = g28.transform(lambda s: s.rolling(28).mean())
    df["rolling_std_7"] = g28.transform(lambda s: s.rolling(7).std())

    # Variables calendaires (déjà présentes dans series_long, recalcul par sécurité)
    df["day_of_week"] = df["date"].dt.dayofweek
    df["month"] = df["date"].dt.month
    df["year"] = df["date"].dt.year
    df["week_of_year"] = df["date"].dt.isocalendar().week.astype(int)
    df["is_weekend"] = (df["day_of_week"] >= 5).astype(int)

    # Encodage entier des identifiants
    df["item_code"] = df["item_id"].astype("category").cat.codes
    df["store_code"] = df["store_id"].astype("category").cat.codes
    return df


PASSIVE_FEATURES = [
    "lag_7", "lag_14", "lag_28", "lag_35",
    "rolling_mean_7", "rolling_mean_28", "rolling_std_7",
    "day_of_week", "month", "year", "week_of_year", "is_weekend",
    "item_code", "store_code",
]


# ---------------------------------------------------------------------------
# Features actives (variables externes : prix, SNAP, événements)
# ---------------------------------------------------------------------------
def add_active_features(df):
    """Ajoute les features externes à partir de series_long."""
    df = df.sort_values(["store_id", "item_id", "date"]).copy()
    keys = ["store_id", "item_id"]

    # Prix : variation % vs la semaine précédente (même jour, 7 jours avant)
    prev_price = df.groupby(keys)["sell_price"].shift(7)
    df["price_change_pct"] = ((df["sell_price"] - prev_price) / prev_price * 100).fillna(0.0)

    # Prix relatif à la moyenne du produit dans le magasin
    mean_price = df.groupby(keys)["sell_price"].transform("mean")
    df["price_rel_mean"] = df["sell_price"] / mean_price

    # Événements : drapeau 0/1 et type encodé
    df["has_event"] = df["has_event"].astype(int)
    df["event_type_code"] = df["event_type"].fillna("none").map(EVENT_TYPE_CODES).fillna(0).astype(int)

    # Jours restant avant le prochain 25 décembre (0 le jour de Noël)
    xmas = pd.to_datetime(df["date"].dt.year.astype(str) + "-12-25")
    xmas = xmas.where(df["date"] <= xmas, xmas + pd.DateOffset(years=1))
    df["days_to_christmas"] = (xmas - df["date"]).dt.days

    # Semaine Walmart (samedi -> vendredi) contenant un événement National/Religious
    week = df["date"].dt.to_period("W-FRI")
    is_big = df["event_type"].isin(["National", "Religious"])
    big_weeks = set(week[is_big].unique())
    df["is_holiday_week"] = week.isin(big_weeks).astype(int)
    return df


ACTIVE_EXTRA_FEATURES = [
    "sell_price", "price_change_pct", "price_rel_mean", "snap",
    "has_event", "event_type_code", "days_to_christmas", "is_holiday_week",
]


# ---------------------------------------------------------------------------
# Métriques
# ---------------------------------------------------------------------------
def compute_metrics(y_true, y_pred):
    """
    MAE, RMSE, MAPE et WAPE.
    - MAPE ignore les jours à 0 vente (division par 0) : peu fiable sur les séries
      intermittentes (HOBBIES, HOUSEHOLD) où elle dépasse souvent 100 %.
    - WAPE = somme |erreur| / somme ventes réelles : pas de division par 0,
      pondère naturellement par le volume -> métrique relative recommandée.
    """
    y_true = np.asarray(y_true, dtype=float)
    y_pred = np.asarray(y_pred, dtype=float)
    err = y_true - y_pred
    mae = np.mean(np.abs(err))
    rmse = np.sqrt(np.mean(err ** 2))
    mask = y_true > 0
    mape = np.mean(np.abs(err[mask]) / y_true[mask]) * 100 if mask.any() else np.nan
    wape = np.abs(err).sum() / y_true.sum() * 100 if y_true.sum() > 0 else np.nan
    return {"MAE": round(mae, 4), "MAPE": round(mape, 2), "RMSE": round(rmse, 4),
            "WAPE": round(wape, 2)}


def rmsse_by_series(fc, history, n_weight_days=28):
    """
    RMSSE et poids de chaque série (store × produit), au sens de la compétition M5.

    - RMSSE = RMSE du test / RMSE du naïf "veille" sur l'historique d'entraînement
      (calculé à partir de la 1re vente non nulle) : < 1 = mieux que le naïf.
    - Poids = chiffre d'affaires (ventes × prix) des `n_weight_days` derniers jours
      d'entraînement, normalisé à 1.
    fc      : date, store_id, item_id, y_true, y_pred (période de test)
    history : date, store_id, item_id, sales, sell_price (avant la période de test)
    """
    rows = []
    hist = history.sort_values(["store_id", "item_id", "date"])
    last_days = hist["date"].max() - pd.Timedelta(days=n_weight_days - 1)
    for (store, item), h in hist.groupby(["store_id", "item_id"]):
        y = h["sales"].to_numpy(dtype=float)
        nz = np.flatnonzero(y)
        y = y[nz[0]:] if len(nz) else y
        scale = np.mean(np.diff(y) ** 2) if len(y) > 1 else np.nan
        recent = h[h["date"] >= last_days]
        rows.append({"store_id": store, "item_id": item, "scale": scale,
                     "revenue": float((recent["sales"] * recent["sell_price"]).sum())})
    info = pd.DataFrame(rows)
    mse = (fc.assign(se=(fc["y_true"] - fc["y_pred"]) ** 2)
             .groupby(["store_id", "item_id"])["se"].mean().rename("mse").reset_index())
    out = mse.merge(info, on=["store_id", "item_id"], how="left")
    out["rmsse"] = np.sqrt(out["mse"] / out["scale"])
    out["weight"] = out["revenue"] / out["revenue"].sum()
    return out


def wrmsse(per_series):
    """WRMSSE d'un sous-ensemble de séries (poids renormalisés dans le sous-ensemble)."""
    d = per_series.dropna(subset=["rmsse"])
    if d.empty or d["weight"].sum() == 0:
        return np.nan
    return round(float((d["rmsse"] * d["weight"]).sum() / d["weight"].sum()), 4)


def metrics_table(fc, model, approach):
    """
    Construit le tableau de métriques à partir d'un fichier de prévisions
    (colonnes date, store_id, item_id, y_true, y_pred).
    Niveaux : global, par store, par store × item (+ store × catégorie si cat_id présent).
    """
    rows = []
    rows.append({"model": model, "approach": approach, "level": "global",
                 "store_id": "ALL", "item_id": "ALL",
                 **compute_metrics(fc["y_true"], fc["y_pred"])})
    for store, d in fc.groupby("store_id"):
        rows.append({"model": model, "approach": approach, "level": "store",
                     "store_id": store, "item_id": "ALL",
                     **compute_metrics(d["y_true"], d["y_pred"])})
    for (store, item), d in fc.groupby(["store_id", "item_id"]):
        rows.append({"model": model, "approach": approach, "level": "item",
                     "store_id": store, "item_id": item,
                     **compute_metrics(d["y_true"], d["y_pred"])})
    # Niveau catégorie (store × catégorie) : uniquement si la colonne cat_id existe dans fc.
    # La catégorie est rangée dans item_id pour garder le même format de fichier.
    if "cat_id" in fc.columns:
        for (store, cat), d in fc.groupby(["store_id", "cat_id"]):
            rows.append({"model": model, "approach": approach, "level": "category",
                         "store_id": store, "item_id": cat,
                         **compute_metrics(d["y_true"], d["y_pred"])})
    return pd.DataFrame(rows)


def print_store_metrics(mt):
    """Affiche les métriques globales et par store dans la console."""
    view = mt[mt["level"].isin(["global", "store"])][["store_id", "MAE", "MAPE", "RMSE"]]
    print(view.to_string(index=False))
