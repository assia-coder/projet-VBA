"""
generate_app.py - Application Excel/VBA de prevision de la demande (M5 Forecasting)
Version 3.0

Genere Application_Prevision.xlsx (base "admin", sans macros) :
  - Accueil : KPI 100 % formules, dependants de la session (role / magasin)
  - Feuilles de donnees T_* : UNIQUEMENT des tableaux Excel (ListObjects),
    aucun panneau de filtres manuel, aucune troncature de donnees
  - Gestion : liste des tables + emplacements des boutons (poses par build_userforms.py)
  - R_Session (cachee) : role / magasin / utilisateur lus par les formules

Sources :
  output/series_long.csv                  -> T_Transactions (toutes les colonnes, toutes les lignes)
  output/top_produits_par_categorie.csv   -> T_Produits
  data/calendar.csv                       -> T_Calendrier (toutes les colonnes)
  deduit                                  -> T_Categories, T_Magasins, T_Prix

Usage :
    python generate_app.py
    python build_userforms.py      (Windows + Excel)
"""

import sys
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.table import Table, TableColumn, TableStyleInfo

# ------------------------------------------------------------------ Chemins
ROOT = Path(__file__).parent
DATA_DIR = ROOT / "data"
OUTPUT_DIR = ROOT / "output"
VBA_DIR = ROOT / "vba"
XLSX_PATH = ROOT / "Application_Prevision.xlsx"
LOGO_PATH = ROOT / "Walmart-Logo.jpg"
SEARCH_DIRS = [OUTPUT_DIR, DATA_DIR, ROOT]

# ------------------------------------------------------------------ Perimetre
SCOPE_STORES = ["CA_1", "TX_1", "WI_1"]
# None = on garde tout le contenu de top_produits_par_categorie.csv.
# Mettre 5 pour ne garder que les 5 produits les plus vendus PAR MAGASIN.
TOP_N_PER_STORE = None

STORE_INFO = {  # store_id -> (state_id, description)
    "CA_1": ("CA", "California"),
    "TX_1": ("TX", "Texas"),
    "WI_1": ("WI", "Wisconsin"),
}
CAT_INFO = {
    "FOODS": "Produits alimentaires",
    "HOBBIES": "Loisirs et passe-temps",
    "HOUSEHOLD": "Articles pour la maison",
}

# ------------------------------------------------------------------ Palette Walmart
P = {
    "primary": "0071CE", "dark": "004F9E", "accent": "FFC220", "light": "78B9E7",
    "very_light": "E6F1FA", "white": "FFFFFF", "bg": "F5F7FA", "border": "D0D5DD",
    "text": "1A1A1A", "muted": "667085", "success": "12B76A", "warning": "F79009",
    "error": "D92D20",
}
PASSWORD = "walmart2024"

# ------------------------------------------------------------------ Layout des feuilles de donnees
HEADER_ROW = 6      # en-tete du tableau (rows 2-4 : bandeau, 5 : espace)
DATA_START_ROW = 7  # 1re ligne de donnees

DATA_SHEETS = ["T_Produits", "T_Categories", "T_Magasins",
               "T_Transactions", "T_Prix", "T_Calendrier"]
TABLE_NAMES = {
    "T_Produits": "tblProduits", "T_Categories": "tblCategories",
    "T_Magasins": "tblMagasins", "T_Transactions": "tblTransactions",
    "T_Prix": "tblPrix", "T_Calendrier": "tblCalendrier",
}
COLS = {}  # nom de feuille -> liste des en-tetes (rempli a la generation)

# ------------------------------------------------------------------ Layout Accueil / Gestion
# (ancre, libelle, macro VBA, couleur de fond, couleur de texte)
ACCUEIL_BUTTONS = [
    ("B11:C12", "▶  Se connecter", "OpenLogin", P["primary"], P["white"]),
    ("D11:E12", "✕  Se déconnecter", "Logout", P["error"], P["white"]),
    ("B14:C15", "▶  Dashboard transactionnel", "OpenPowerBI", P["primary"], P["white"]),
    ("D14:E15", "▶  Résultats Machine Learning", "OpenML", P["primary"], P["white"]),
    ("F14:G15", "▶  Prévision de la demande", "OpenPrevision", P["primary"], P["white"]),
    ("H14:K15", "▶  Gestion des données", "OpenGestion", P["accent"], P["dark"]),
]
GESTION_TABLES = [  # (libelle, description, feuille, formulaire, macro)
    ("Produits", "Produits suivis (clé magasin × produit)", "T_Produits", "frmProduits", "OpenProduits"),
    ("Catégories", "Catégories de produits", "T_Categories", "frmCategories", "OpenCategories"),
    ("Magasins", "Magasins du périmètre", "T_Magasins", "frmMagasins", "OpenMagasins"),
    ("Transactions", "Historique des ventes journalières", "T_Transactions", "frmTransactions", "OpenTransactions"),
    ("Prix", "Prix hebdomadaires", "T_Prix", "frmPrix", "OpenPrix"),
    ("Calendrier", "Calendrier + événements", "T_Calendrier", "frmCalendrier", "OpenCalendrier"),
]
GESTION_FIRST_TABLE_ROW = 13
GESTION_BUTTONS = [
    ("E5:F6", "←  Accueil", "GoHome", P["primary"], P["white"]),
    ("G5:H6", "✕  Se déconnecter", "Logout", P["error"], P["white"]),
] + [
    (f"H{GESTION_FIRST_TABLE_ROW + i}:I{GESTION_FIRST_TABLE_ROW + i}", f"▶ {t[0]}", t[4],
     P["accent"], P["dark"])
    for i, t in enumerate(GESTION_TABLES)
]

# Largeurs de colonnes des tableaux (defaut : 14)
COL_WIDTHS = {
    "store_id": 11, "item_id": 24, "cat_id": 13, "dept_id": 13, "date": 12,
    "description": 26, "event_name": 22, "event_type": 14, "event_name_1": 24,
    "event_name_2": 24, "event_type_1": 14, "event_type_2": 14, "source": 10,
    "saisi_le": 18, "date_ajout": 12, "weekday": 12, "ventes_12_derniers_mois": 22,
    "part_top_cat_pct": 16, "ventes_moy_jour": 16, "days_to_monthend": 16,
}


# ------------------------------------------------------------------ Helpers styles
def fill(c):
    return PatternFill("solid", fgColor=c)


def fnt(size=11, bold=False, color=None, italic=False):
    return Font(name="Calibri", size=size, bold=bold, italic=italic, color=color or P["text"])


def bord(color=None, style="thin"):
    s = Side(style=style, color=color or P["border"])
    return Border(left=s, right=s, top=s, bottom=s)


def title_banner(ws, rng, text, subtitle=None, bg=None):
    bg = bg or P["primary"]
    ws.merge_cells(rng)
    for row in ws[rng]:
        for c in row:
            c.fill = fill(bg)
    top = ws[rng.split(":")[0]]
    top.value = text
    top.font = fnt(18, bold=True, color=P["white"])
    top.alignment = Alignment(horizontal="left", vertical="center", indent=2)
    if subtitle:
        parts = rng.split(":")
        col_l, row_l = parts[0][0], int(parts[0][1:]) + 1
        col_r, row_r = parts[1][0], int(parts[1][1:]) + 1
        rng2 = f"{col_l}{row_l}:{col_r}{row_r}"
        ws.merge_cells(rng2)
        for row in ws[rng2]:
            for c in row:
                c.fill = fill(bg)
        c = ws[f"{col_l}{row_l}"]
        c.value = subtitle
        c.font = fnt(10, italic=True, color=P["white"])
        c.alignment = Alignment(horizontal="left", vertical="center", indent=2)


def section(ws, rng, text, bg=None):
    bg = bg or P["dark"]
    ws.merge_cells(rng)
    for row in ws[rng]:
        for c in row:
            c.fill = fill(bg)
    c = ws[rng.split(":")[0]]
    c.value = text
    c.font = fnt(12, bold=True, color=P["white"])
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)


def header_row(ws, row, labels, start_col=1):
    for i, lab in enumerate(labels):
        cell = ws.cell(row=row, column=start_col + i, value=lab)
        cell.font = fnt(11, bold=True, color=P["white"])
        cell.fill = fill(P["primary"])
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = bord(P["primary"])


def style_block(ws, rng, bg, border_color=None):
    for row in ws[rng]:
        for c in row:
            c.fill = fill(bg)
            if border_color:
                c.border = bord(border_color)


# ------------------------------------------------------------------ Helpers donnees
def find_file(*names):
    for d in SEARCH_DIRS:
        for n in names:
            p = d / n
            if p.exists():
                return p
    return None


def read_csv_required(label, *names):
    p = find_file(*names)
    if p is None:
        print(f"[ERREUR] Fichier introuvable : {names[0]} (cherché dans {[str(d) for d in SEARCH_DIRS]})")
        sys.exit(1)
    print(f"   lecture {p.relative_to(ROOT) if p.is_relative_to(ROOT) else p}")
    return pd.read_csv(p)


def clean(df):
    """DataFrame -> objets Python natifs, NaN/NaT -> None."""
    return df.astype(object).where(df.notna(), None)


def L(sheet, header):
    """Lettre de colonne Excel d'un en-tete de tableau."""
    return get_column_letter(2 + COLS[sheet].index(header))


def rng(sheet, header, last=5000):
    c = L(sheet, header)
    return f"{sheet}!${c}${DATA_START_ROW}:${c}${last}"


# ------------------------------------------------------------------ Feuilles de donnees (tableaux)
def write_table_sheet(wb, sheet, title, subtitle, df):
    ws = wb.create_sheet(sheet)
    cols = [str(c) for c in df.columns]
    COLS[sheet] = cols
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = P["primary"]
    ws.column_dimensions["A"].width = 3
    for i, c in enumerate(cols):
        ws.column_dimensions[get_column_letter(2 + i)].width = COL_WIDTHS.get(c, 14)

    ws.row_dimensions[2].height = 34
    title_banner(ws, "B2:K2", title, subtitle=subtitle)
    ws.row_dimensions[4].height = 6
    for c in ws["B4:K4"][0]:
        c.fill = fill(P["accent"])

    header_row(ws, HEADER_ROW, cols, start_col=2)
    ws.row_dimensions[HEADER_ROW].height = 28

    # Donnees : ws.append continue sous la ligne d'en-tete (colonne A laissee vide)
    n = 0
    for row in clean(df).itertuples(index=False, name=None):
        ws.append([None] + list(row))
        n += 1
    if n == 0:  # un tableau Excel doit contenir au moins 1 ligne de donnees
        ws.append([None] * (len(cols) + 1))
    last_row = HEADER_ROW + max(n, 1)

    ref = f"B{HEADER_ROW}:{get_column_letter(1 + len(cols))}{last_row}"
    tab = Table(displayName=TABLE_NAMES[sheet], ref=ref,
                tableColumns=[TableColumn(id=i + 1, name=c) for i, c in enumerate(cols)])
    tab.autoFilter = None   # pas de flèches de filtre : filtrage = role, applique par VBA
    tab.tableStyleInfo = TableStyleInfo(name="TableStyleLight1", showRowStripes=True)
    ws.add_table(tab)
    ws.freeze_panes = f"A{DATA_START_ROW}"
    print(f"→ {sheet:<15} {n:>8} lignes × {len(cols)} colonnes")
    return n


def load_produits():
    df = read_csv_required("produits", "top_produits_par_categorie.csv", "top_produits_par_catégorie.csv")
    wanted = ["store_id", "item_id", "cat_id", "dept_id", "rang", "ventes_totales", "ca_total",
              "ventes_moy_jour", "prix_moyen", "zero_ratio", "ventes_12_derniers_mois", "part_top_cat_pct"]
    for c in wanted:
        if c not in df.columns:
            print(f"[warn] colonne absente de top_produits : {c}")
            df[c] = None
    df = df[wanted].copy()
    before = len(df)
    df = df[df["store_id"].isin(SCOPE_STORES)]
    if len(df) != before:
        print(f"   [info] {before - len(df)} ligne(s) hors périmètre {SCOPE_STORES} ignorée(s)")
    if TOP_N_PER_STORE:
        df = (df.sort_values("ventes_totales", ascending=False)
                .groupby("store_id", group_keys=False).head(TOP_N_PER_STORE))
    df = df.drop_duplicates(["store_id", "item_id"]).sort_values(["store_id", "cat_id", "rang"])
    df["date_ajout"] = None
    df["actif"] = True
    df["source"] = "M5"
    return df.reset_index(drop=True)


def load_transactions(produits):
    df = read_csv_required("series_long", "series_long.csv")
    for c in ("date", "store_id", "item_id"):
        if c not in df.columns:
            print(f"[ERREUR] series_long.csv : colonne obligatoire absente : {c}")
            sys.exit(1)
    before = len(df)
    pairs = set(zip(produits["store_id"], produits["item_id"]))
    mask = pd.MultiIndex.from_frame(df[["store_id", "item_id"]]).isin(pairs)
    df = df[mask].copy()
    if len(df) != before:
        print(f"   [info] {before - len(df)} transaction(s) hors périmètre (magasin × produit) ignorée(s)")
    df["date"] = pd.to_datetime(df["date"]).dt.date
    df["source"] = "M5"
    df["saisi_le"] = None
    return df.reset_index(drop=True)


def load_calendar():
    df = read_csv_required("calendar", "calendar.csv")
    df["date"] = pd.to_datetime(df["date"])
    df = df.sort_values("date")
    return df


def build_prix(trans_raw_dates, calendar):
    """Prix hebdomadaires deduits de series_long (1 ligne par magasin × produit × semaine M5)."""
    s = trans_raw_dates[["date", "store_id", "item_id", "sell_price"]].copy()
    s["date"] = pd.to_datetime(s["date"])
    s = s.dropna(subset=["sell_price"])
    if "wm_yr_wk" in calendar.columns:
        s = s.merge(calendar[["date", "wm_yr_wk"]], on="date", how="left")
    else:
        s["wm_yr_wk"] = s["date"].dt.to_period("W-SUN").dt.start_time.dt.strftime("%Y%m%d").astype(int)
    g = (s.sort_values("date")
          .groupby(["store_id", "item_id", "wm_yr_wk"], as_index=False)
          .agg(date=("date", "min"), sell_price=("sell_price", "first")))
    g["date"] = g["date"].dt.date
    g["source"] = "M5"
    g = g[["date", "wm_yr_wk", "store_id", "item_id", "sell_price", "source"]]
    return g.sort_values(["store_id", "item_id", "date"]).reset_index(drop=True)


# ------------------------------------------------------------------ Gestion
def build_gestion(wb):
    print("→ Gestion")
    ws = wb.create_sheet("Gestion")
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = P["accent"]
    for col, w in zip("ABCDEFGHIJK", [3, 30, 30, 14, 16, 16, 20, 14, 14, 14, 14]):
        ws.column_dimensions[col].width = w

    ws.row_dimensions[2].height = 34
    title_banner(ws, "B2:K2", "Gestion des données",
                 subtitle="Tables de l'application : consultation, recherche, saisie contrôlée via UserForms.")

    section(ws, "B4:K4", "  Session utilisateur", bg=P["dark"])
    for r, label, formula in [
        (5, "Statut session :", '=IF(sessRole="","non connecté",sessRoleLabel)'),
        (6, "Utilisateur :", '=IF(sessRole="","—",sessUser)'),
        (7, "Magasin :", '=IF(sessRole="","—",IF(sessStore="ALL","Tous les magasins",sessStore))'),
    ]:
        ws.cell(row=r, column=2, value=label).font = fnt(11, bold=True, color=P["dark"])
        c = ws.cell(row=r, column=3, value=formula)
        c.font = fnt(11, color=P["muted"])
    # zones des boutons (les boutons eux-memes sont poses par build_userforms.py)
    for anchor, label, _m, bg, fg in GESTION_BUTTONS[:2]:
        ws.merge_cells(anchor)
        style_block(ws, anchor, bg, bg)
        c = ws[anchor.split(":")[0]]
        c.value = label
        c.font = fnt(11, bold=True, color=fg)
        c.alignment = Alignment(horizontal="center", vertical="center")

    section(ws, "B11:K11", "  Tables de l'application", bg=P["primary"])
    header_row(ws, 12, ["Table", "Contenu", "Lignes", "dont MANUEL", "Accès", "", "Formulaire", ""], start_col=2)
    ws.merge_cells("F12:G12")
    ws.merge_cells("H12:I12")

    for i, (nom, desc, sheet, _form, _macro) in enumerate(GESTION_TABLES):
        r = GESTION_FIRST_TABLE_ROW + i
        key = L(sheet, COLS[sheet][0])
        src = L(sheet, "source")
        ws.cell(row=r, column=2, value=nom).font = fnt(11, bold=True, color=P["dark"])
        ws.cell(row=r, column=3, value=desc).font = fnt(10)
        # SUBTOTAL(103) ne compte que les lignes visibles : suit le filtre du role connecte
        ws.cell(row=r, column=4,
                value=f"=SUBTOTAL(103,{sheet}!${key}${DATA_START_ROW}:${key}$1048576)").font = fnt(10)
        if "store_id" in COLS[sheet] and sheet != "T_Categories":
            st = L(sheet, "store_id")
            f_manuel = (f'=COUNTIFS({sheet}!${src}${DATA_START_ROW}:${src}$1048576,"MANUEL",'
                        f'{sheet}!${st}${DATA_START_ROW}:${st}$1048576,sessStoreCrit)')
        else:
            f_manuel = f'=COUNTIF({sheet}!${src}${DATA_START_ROW}:${src}$1048576,"MANUEL")'
        ws.cell(row=r, column=5, value=f_manuel).font = fnt(10)
        ws.merge_cells(f"F{r}:G{r}")
        c = ws.cell(row=r, column=6,
                    value='=IF(OR(sessRole="admin",sessRole="hq"),"Lecture / écriture","Lecture seule")')
        c.font = fnt(10, italic=True, color=P["muted"])
        ws.merge_cells(f"H{r}:I{r}")
        style_block(ws, f"H{r}:I{r}", P["accent"], P["accent"])
        c = ws.cell(row=r, column=8, value=f"▶ {nom}")
        c.font = fnt(10, bold=True, color=P["dark"])
        c.alignment = Alignment(horizontal="center", vertical="center")

    section(ws, "B21:K21", "  Règles de cohérence", bg=P["dark"])
    regles = [
        "• Les lignes source = M5 sont protégées : transactions, prix et calendrier en lecture seule.",
        "• Les saisies utilisateur sont en source = MANUEL et restent distinctes du dataset.",
        "• Un produit lié à des transactions ne peut pas être supprimé : il est désactivé (actif = FALSE).",
        "• Une catégorie ou un magasin ne se supprime que s'il n'a plus de produit.",
        "• Gérant de magasin : son magasin uniquement · Analyste HQ : les 3 magasins · Administrateur : accès complet.",
        "• Les modifications manuelles n'alimentent pas les modèles Python (qui lisent output/series_long.csv).",
    ]
    for i, txt in enumerate(regles, start=22):
        ws.merge_cells(f"B{i}:K{i}")
        c = ws.cell(row=i, column=2, value=txt)
        c.font = fnt(10)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[i].height = 28


# ------------------------------------------------------------------ Feuilles annexes
def add_name(wb, name, ref):
    dn = DefinedName(name, attr_text=ref)
    try:
        wb.defined_names[name] = dn
    except TypeError:
        wb.defined_names.append(dn)


def build_session_sheet(wb):
    ws = wb.create_sheet("R_Session")
    ws["A1"], ws["B1"] = "key", "value"
    for i, k in enumerate(["role", "store", "username", "login_time", "role_label", "store_crit"], start=2):
        ws.cell(row=i, column=1, value=k)
    ws["B6"] = ('=IF(B2="admin","Administrateur",IF(B2="hq","Analyste HQ",'
                'IF(B2="store","Gérant de magasin","")))')
    ws["B7"] = '=IF(OR(B3="",B3="ALL"),"*",B3)'   # critere COUNTIFS : "*" = tous les magasins
    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 28
    ws.sheet_state = "hidden"
    for n, ref in [("sessRole", "R_Session!$B$2"), ("sessStore", "R_Session!$B$3"),
                   ("sessUser", "R_Session!$B$4"), ("sessRoleLabel", "R_Session!$B$6"),
                   ("sessStoreCrit", "R_Session!$B$7")]:
        add_name(wb, n, ref)


def build_powerbi(wb):
    ws = wb.create_sheet("PowerBI")
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = P["primary"]
    ws.row_dimensions[2].height = 30
    title_banner(ws, "B2:K2", "PowerBI", subtitle="Dashboard transactionnel — Power BI")
    ws["B4"] = "État :"
    ws["C4"] = '=IF(cfgPowerBI="","Aucun dashboard configuré (Config!C5).","Dashboard configuré.")'
    ws["B5"] = "Lien :"
    ws["C5"] = '=IF(cfgPowerBI="","",HYPERLINK(cfgPowerBI,"▶ Ouvrir le dashboard Power BI"))'
    ws["B7"] = "Procédures :"
    for i, t in enumerate([
        "A · Publier le .pbix dans Power BI, copier l'URL dans Config!C5.",
        "B · Insertion > Compléments Office > Microsoft Power BI (zone C20:L45).",
        "C · Coller le chemin local du .pbix dans Config!C5.",
    ], start=8):
        ws.merge_cells(f"B{i}:K{i}")
        ws.cell(row=i, column=2, value=t).font = fnt(10)


def build_ml(wb):
    ws = wb.create_sheet("ML")
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = P["primary"]
    ws.row_dimensions[2].height = 30
    title_banner(ws, "B2:L2", "ML", subtitle="Résultats Machine Learning — comparaison des modèles")


def build_prevision(wb):
    ws = wb.create_sheet("Prevision")
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = P["primary"]
    ws.row_dimensions[2].height = 30
    title_banner(ws, "B2:O2", "Prevision", subtitle="Prévision de la demande — choisir les paramètres")


def build_config(wb):
    ws = wb.create_sheet("Config")
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["B"].width = 30
    ws.column_dimensions["C"].width = 50
    ws["B1"] = "Configuration"
    ws["B1"].font = fnt(14, bold=True, color=P["dark"])
    ws["B3"] = "PowerBI URL"
    ws["C3"] = ""
    ws["B4"] = "SHEET_PASSWORD"
    ws["C4"] = PASSWORD
    ws["B5"] = "pbi_path"
    ws["C5"] = ""
    add_name(wb, "cfgPowerBI", "Config!$C$5")


def build_r_sheets(wb):
    specs = {
        "R_Previsions": ["date", "store_id", "item_id", "cat_id", "model", "set", "pred", "pi_low", "pi_high"],
        "R_Metriques": ["key", "model", "level", "store", "item", "n", "MAE", "MAPE", "RMSE", "WAPE", "WRMSSE"],
        "R_Resume": ["key", "value"],
        "R_Mensuel": ["date", "horizon", "pred", "store_id", "item_id", "cat_id", "pi_low", "pi_high", ""],
        "R_MetMensuel": ["key", "model", "level", "store", "item", "MAE", "RMSE", "MAPE", "WAPE", "IC80", "n"],
    }
    for name, headers in specs.items():
        ws = wb.create_sheet(name)
        ws.sheet_view.showGridLines = False
        ws.sheet_properties.tabColor = P["muted"]
        for j, h in enumerate(headers, start=1):
            c = ws.cell(row=1, column=j, value=h)
            c.font = fnt(10, bold=True, color=P["white"])
            c.fill = fill(P["muted"])


# ------------------------------------------------------------------ Accueil
def build_accueil(ws):
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = P["primary"]
    for col, w in zip("ABCDEFGHIJK", [3] + [22] * 10):
        ws.column_dimensions[col].width = w

    ws.row_dimensions[2].height = 42
    title_banner(ws, "B2:K2", "Prévision de la demande — M5 Forecasting",
                 subtitle="Walmart · magasins CA_1 · TX_1 · WI_1 · Excel / VBA · Python · Power BI")
    ws.row_dimensions[4].height = 6
    for c in ws["B4:K4"][0]:
        c.fill = fill(P["accent"])

    if LOGO_PATH.exists():
        img = XLImage(str(LOGO_PATH))
        ratio = img.width / img.height if img.height else 1
        img.width = 180
        img.height = int(180 / ratio)
        ws.add_image(img, "B5")
        ws.row_dimensions[5].height = 70
        ws.row_dimensions[6].height = 15

    # ---- KPI : formules dependantes de la session (role / magasin) -> toujours exacts
    ws.row_dimensions[8].height = 26
    ws.row_dimensions[9].height = 44
    ws.row_dimensions[10].height = 10
    p_store = rng("T_Produits", "store_id")
    p_cat = rng("T_Produits", "cat_id")
    p_act = rng("T_Produits", "actif")
    kpis = [
        ("Produits suivis",
         f'=IF(sessRole="","—",COUNTIFS({p_store},sessStoreCrit,{p_act},TRUE))'),
        ("Magasins",
         f'=IF(sessRole="","—",COUNTIFS({rng("T_Magasins", "store_id", 200)},sessStoreCrit))'),
        ("Catégories",
         f'=IF(sessRole="","—",SUMPRODUCT(--(COUNTIFS({p_store},sessStoreCrit,{p_act},TRUE,'
         f'{p_cat},{rng("T_Categories", "cat_id", 200)})>0)))'),
        ("Modèle gagnant",
         '=IF(sessRole="","—",IFERROR(INDEX(R_Resume!$B$2:$B$60,MATCH("winner_label",R_Resume!$A$2:$A$60,0)),"n/d"))'),
    ]
    spans = [("B", "D"), ("E", "G"), ("H", "I"), ("J", "K")]
    for (label, formula), (c1, c2) in zip(kpis, spans):
        ws.merge_cells(f"{c1}8:{c2}8")
        ws.merge_cells(f"{c1}9:{c2}9")
        lc = ws[f"{c1}8"]
        lc.value = label
        lc.font = fnt(10, bold=True, color=P["white"])
        lc.fill = fill(P["primary"])
        lc.alignment = Alignment(horizontal="center", vertical="center")
        vc = ws[f"{c1}9"]
        vc.value = formula
        vc.number_format = "#,##0"
        vc.font = fnt(18, bold=True, color=P["dark"])
        vc.fill = fill(P["light"])
        vc.alignment = Alignment(horizontal="center", vertical="center")
        for row in ws[f"{c1}8:{c2}9"]:
            for c in row:
                c.border = bord(P["primary"])

    # ---- Connexion (lignes 11-12) : boutons poses par build_userforms.py
    ws.row_dimensions[11].height = 22
    ws.row_dimensions[12].height = 22
    for anchor, label, _m, bg, fg in ACCUEIL_BUTTONS[:2]:
        ws.merge_cells(anchor)
        style_block(ws, anchor, bg, bg)
        c = ws[anchor.split(":")[0]]
        c.value = label
        c.font = fnt(11, bold=True, color=fg)
        c.alignment = Alignment(horizontal="center", vertical="center")
    ws.merge_cells("F11:K12")
    style_block(ws, "F11:K12", P["very_light"], P["border"])
    c = ws["F11"]
    c.value = ('=IF(sessRole="","Non connecté — cliquez sur « Se connecter »",'
               '"Connecté : "&sessUser&"  ·  "&sessRoleLabel&"  ·  Magasin : "&IF(sessStore="ALL","Tous",sessStore))')
    c.font = fnt(10, bold=True, color=P["dark"])
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)

    # ---- Navigation (zones des boutons, sans hyperliens : les feuilles sont masquees avant login)
    ws.row_dimensions[13].height = 24
    ws.merge_cells("B13:K13")
    ws["B13"].value = "  Navigation"
    ws["B13"].font = fnt(13, bold=True, color=P["white"])
    ws["B13"].fill = fill(P["dark"])
    ws["B13"].alignment = Alignment(vertical="center")
    ws.row_dimensions[14].height = 60
    for anchor, label, _m, bg, fg in ACCUEIL_BUTTONS[2:]:
        ws.merge_cells(anchor)
        style_block(ws, anchor, bg, bg)
        c = ws[anchor.split(":")[0]]
        c.value = label
        c.font = fnt(11, bold=True, color=fg)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    ws.row_dimensions[17].height = 22
    ws.merge_cells("B17:K17")
    ws["B17"].value = "  Contexte et objectif"
    ws["B17"].font = fnt(12, bold=True, color=P["white"])
    ws["B17"].fill = fill(P["dark"])
    ws["B17"].alignment = Alignment(vertical="center")
    ws.merge_cells("B18:K20")
    ws["B18"] = ("Projet académique de prévision de la demande sur le jeu de données M5 Forecasting (Walmart). "
                 "L'application réunit prévision (modèles entraînés en Python), analyse des ventes (Power BI), "
                 "comparaison des modèles et gestion des tables de référence, dans un seul classeur Excel piloté par VBA.")
    ws["B18"].font = fnt(10)
    ws["B18"].alignment = Alignment(wrap_text=True, vertical="top")
    for row in ws["B18:K20"]:
        for c in row:
            c.fill = fill(P["bg"])
            c.border = bord(P["border"])


# ------------------------------------------------------------------ MAIN
def main():
    print("=== Génération Application_Prevision (v3.0) ===")
    wb = Workbook()
    wb.remove(wb.active)

    # ---- chargement des sources
    produits = load_produits()
    trans = load_transactions(produits)
    calendar = load_calendar()
    prix = build_prix(trans, calendar)

    cal_out = calendar.copy()
    cal_out["date"] = cal_out["date"].dt.date
    cal_out["source"] = "M5"

    # ---- feuilles de donnees (T_Produits d'abord : les formules des autres en dependent)
    write_table_sheet(wb, "T_Produits", "Produits suivis",
                      "Clé magasin × produit — désactivation logique (actif = FALSE) plutôt que suppression.",
                      produits)

    # lettres de colonnes de T_Produits (pour les formules nb_produits)
    p_store, p_cat, p_act = (rng("T_Produits", h) for h in ("store_id", "cat_id", "actif"))

    cats = sorted(produits["cat_id"].dropna().unique())
    cat_cols = ["cat_id", "description", "nb_produits", "source"]
    cat_rows = []
    for i, c in enumerate(cats):
        r = DATA_START_ROW + i
        cat_rows.append([c, CAT_INFO.get(c, ""),
                         f'=COUNTIFS({p_cat},$B{r},{p_store},sessStoreCrit,{p_act},TRUE)', "M5"])
    write_table_sheet(wb, "T_Categories", "Catégories", ", ".join(cats),
                      pd.DataFrame(cat_rows, columns=cat_cols))

    stores = sorted(produits["store_id"].dropna().unique())
    mag_cols = ["store_id", "state_id", "description", "nb_produits", "source"]
    mag_rows = []
    for i, s in enumerate(stores):
        r = DATA_START_ROW + i
        st, desc = STORE_INFO.get(s, (s.split("_")[0], ""))
        mag_rows.append([s, st, desc, f'=COUNTIFS({p_store},$B{r},{p_act},TRUE)', "M5"])
    write_table_sheet(wb, "T_Magasins", "Magasins", ", ".join(stores),
                      pd.DataFrame(mag_rows, columns=mag_cols))

    write_table_sheet(wb, "T_Transactions", "Transactions (historique des ventes)",
                      "Lignes M5 en lecture seule ; saisies manuelles = source MANUEL.", trans)
    write_table_sheet(wb, "T_Prix", "Prix des produits",
                      "Prix hebdomadaires par magasin × produit (déduits de series_long).", prix)
    write_table_sheet(wb, "T_Calendrier", "Calendrier et événements",
                      "Événements manuels = source MANUEL (sans effet sur modèles entraînés).", cal_out)

    # ---- feuilles applicatives
    build_gestion(wb)
    build_powerbi(wb)
    build_ml(wb)
    build_prevision(wb)
    build_r_sheets(wb)
    build_session_sheet(wb)
    build_config(wb)
    ws_accueil = wb.create_sheet("Accueil")
    build_accueil(ws_accueil)

    order = ["Accueil", "PowerBI", "ML", "Prevision", "Gestion",
             "T_Produits", "T_Categories", "T_Magasins",
             "T_Transactions", "T_Prix", "T_Calendrier",
             "R_Previsions", "R_Metriques", "R_Resume", "R_Mensuel",
             "R_MetMensuel", "R_Session", "Config"]
    for i, name in enumerate(order):
        wb.move_sheet(name, offset=(i - wb.sheetnames.index(name)))

    # une seule feuille selectionnee (sinon Excel ouvre les feuilles "groupees")
    for ws in wb.worksheets:
        ws.sheet_view.tabSelected = (ws.title == "Accueil")
    wb.active = wb.sheetnames.index("Accueil")

    wb.save(XLSX_PATH)
    print(f"✔ {XLSX_PATH.name} généré ({len(wb.sheetnames)} feuilles)")
    print("\nÉtape suivante : python build_userforms.py  (Windows + Excel)")


if __name__ == "__main__":
    main()