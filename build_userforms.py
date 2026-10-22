"""
build_userforms.py - Application_Prevision.xlsx -> Application_Prevision.xlsm  (Windows + Excel)

  1. ouvre le .xlsx genere par generate_app.py et l'enregistre en .xlsm (vrai format macro)
  2. injecte les modules VBA (dossier vba/)
  3. cree les UserForms : frmLogin + un formulaire CRUD par table (colonnes lues dans les tableaux)
  4. pose les boutons (formes colorees, avec marge) : Accueil, Gestion et bouton « Accueil » sur toutes les feuilles
  5. ecrit les evenements du classeur (ouverture / activation / enregistrement / fermeture)
  6. verrouille le classeur (seule "Accueil" visible) et enregistre

Pre-requis Excel : Fichier > Options > Centre de gestion de la confidentialite >
Parametres du Centre de gestion de la confidentialite > Parametres des macros >
cocher "Acceder en confiance au modele d'objet du projet VBA".

Usage :  python generate_app.py   puis   python build_userforms.py
"""
import math
import re
import shutil
import sys
import time
import traceback
import uuid
from pathlib import Path

try:
    import win32com.client as win32
except ImportError:
    print("[ERREUR] pywin32 non installé :  pip install pywin32")
    sys.exit(1)

import generate_app as G

ROOT = Path(__file__).parent
XLSX = G.XLSX_PATH
XLSM = ROOT / "Application_Prevision.xlsm"
VBA_DIR = ROOT / "vba"
PWD = G.PASSWORD

REQUIRED = ["ModuleAuth", "ModuleCRUD", "ModuleFilters", "ModuleNavigation", "ModuleProtection"]

# Colonnes jamais saisies a la main dans les formulaires (calculees / renseignees par VBA)
# (prix_moyen, weekday, month, year = formules Excel ; sell_price de T_Prix = formule, sans formulaire)
AUTO_COLS = {"source", "saisi_le", "date_ajout", "nb_produits",
             "prix_moyen", "weekday", "month", "year", "sell_price"}

# nom du formulaire -> (feuille, colonnes cles)
CRUD_FORMS = {
    "Produits":     ("T_Produits",     ["store_id", "item_id"]),
    "Categories":   ("T_Categories",   ["cat_id"]),
    "Magasins":     ("T_Magasins",     ["store_id"]),
    "Transactions": ("T_Transactions", ["date", "store_id", "item_id"]),
    "Calendrier":   ("T_Calendrier",   ["date"]),
}


# ------------------------------------------------------------------ .bas
def read_bas_text(path: Path) -> str:
    raw = None
    for enc in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
        try:
            raw = path.read_text(encoding=enc)
            break
        except UnicodeDecodeError:
            continue
    if raw is None:
        raw = path.read_text(encoding="utf-8", errors="replace")
    raw = re.sub(r"^```[a-zA-Z0-9]*\s*$", "", raw, flags=re.MULTILINE).replace("```", "")
    lines = raw.splitlines()
    i = 0
    while i < len(lines) and lines[i].strip() == "":
        i += 1
    if i < len(lines) and lines[i].strip().startswith("Attribute VB_Name"):
        lines = lines[i + 1:]
    else:
        lines = lines[i:]
    return "\r\n".join(lines).strip() + "\r\n"


def inject_all_modules(vbproj):
    ok = 0
    for name in REQUIRED:
        bas = VBA_DIR / f"{name}.bas"
        code = read_bas_text(bas)
        try:
            try:
                vbproj.VBComponents.Remove(vbproj.VBComponents(name))
            except Exception:
                pass
            c = vbproj.VBComponents.Add(1)
            c.Name = name
            c.CodeModule.AddFromString(code)
            print(f"  + {bas.name}")
            ok += 1
        except Exception as e:
            print(f"  ! {bas.name}: {e}")
    return ok


# ------------------------------------------------------------------ boutons (formes colorees)
def rgb(hexstr: str) -> int:
    r, g, b = int(hexstr[0:2], 16), int(hexstr[2:4], 16), int(hexstr[4:6], 16)
    return r + g * 256 + b * 65536


def add_shape_button(ws, anchor, caption, macro, bg, fg, size=11, pad=4):
    """Bouton = forme coloree posee sur la zone `anchor`, reduite de `pad` points de chaque cote
    (-> espace visible entre deux boutons voisins et par rapport aux cellules colorees)."""
    name = f"btn_{macro}_{anchor}"
    for i in range(ws.Shapes.Count, 0, -1):          # evite les doublons si relance
        try:
            if ws.Shapes(i).Name == name:
                ws.Shapes(i).Delete()
        except Exception:
            pass
    rng = ws.Range(anchor)
    p = min(pad, rng.Height / 4, rng.Width / 4)
    shp = ws.Shapes.AddShape(5, rng.Left + p, rng.Top + p, rng.Width - 2 * p, rng.Height - 2 * p)  # 5 = arrondi
    shp.Name = name
    try:
        shp.Adjustments.Item(1).Value = 0.18
    except Exception:
        pass
    shp.OnAction = macro
    shp.Placement = 2                                # deplace sans redimensionner (xlMove)
    shp.Fill.ForeColor.RGB = rgb(bg)
    shp.Line.ForeColor.RGB = rgb(bg)
    try:
        shp.Shadow.Visible = False
    except Exception:
        pass
    try:
        tf = shp.TextFrame2
        tf.TextRange.Text = caption
        tf.TextRange.Font.Name = "Calibri"
        tf.TextRange.Font.Size = size
        tf.TextRange.Font.Bold = True
        tf.TextRange.Font.Fill.ForeColor.RGB = rgb(fg)
        tf.TextRange.ParagraphFormat.Alignment = 2      # centre
        tf.VerticalAnchor = 3                           # milieu
        tf.WordWrap = True
    except Exception as e:
        print(f"  ! texte du bouton {macro}: {e}")
    return shp


def add_buttons(wb, sheet_name, specs):
    try:
        ws = wb.Sheets(sheet_name)
        ws.Activate()
        for anchor, caption, macro, bg, fg in specs:
            add_shape_button(ws, anchor, caption, macro, bg, fg)
        ws.Range("A1").Select()
        print(f"  + {len(specs)} bouton(s) sur {sheet_name}")
    except Exception as e:
        print(f"  ! boutons de {sheet_name}: {e}")


# ------------------------------------------------------------------ evenements du classeur
THISWORKBOOK_CODE = (
    "Option Explicit\r\n\r\n"
    "Private Sub Workbook_Open()\r\n"
    "    LockWorkbook\r\n"
    "    ThisWorkbook.Saved = True\r\n"
    "End Sub\r\n\r\n"
    "Private Sub Workbook_SheetActivate(ByVal Sh As Object)\r\n"
    "    If IsLoggedIn() Then ApplyRoleFilterToSheet Sh.Name\r\n"
    "End Sub\r\n\r\n"
    "Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)\r\n"
    "    PrepareForSave\r\n"
    "End Sub\r\n\r\n"
    "Private Sub Workbook_BeforeClose(Cancel As Boolean)\r\n"
    "    HandleBeforeClose Cancel\r\n"
    "End Sub\r\n"
)


def set_workbook_events(vbproj):
    cm = vbproj.VBComponents("ThisWorkbook").CodeModule
    if cm.CountOfLines > 0:
        cm.DeleteLines(1, cm.CountOfLines)
    cm.AddFromString(THISWORKBOOK_CODE)
    print("  + événements ThisWorkbook (Open / SheetActivate / BeforeSave / BeforeClose)")


# ------------------------------------------------------------------ UserForms
def create_form(vbproj, final_name, caption, width, height, controls, code):
    try:
        vbproj.VBComponents.Remove(vbproj.VBComponents(final_name))
    except Exception:
        pass
    tmp = f"tmp_{final_name}_{uuid.uuid4().hex[:8]}"
    comp = vbproj.VBComponents.Add(3)
    comp.Name = tmp
    comp.Properties("Caption").Value = caption
    comp.Properties("Width").Value = width
    comp.Properties("Height").Value = height
    try:
        comp.Properties("StartUpPosition").Value = 1       # centre sur Excel
    except Exception:
        pass
    designer = comp.Designer
    for spec in controls:
        try:
            ctrl = designer.Controls.Add(spec["progid"])
            ctrl.Name = spec["name"]
            ctrl.Left = spec.get("left", 10)
            ctrl.Top = spec.get("top", 10)
            ctrl.Width = spec.get("width", 120)
            ctrl.Height = spec.get("height", 20)
            if "caption" in spec:
                ctrl.Caption = spec["caption"]
            try:
                ctrl.Font.Name = "Calibri"
                ctrl.Font.Size = spec.get("font_size", 10)
                if spec.get("bold"):
                    ctrl.Font.Bold = True
            except Exception:
                pass
            for prop, val in spec.get("props", {}).items():
                try:
                    setattr(ctrl, prop, val)
                except Exception as e:
                    print(f"  ! {spec['name']}.{prop}: {e}")
        except Exception as e:
            print(f"  ! ctrl {spec['name']}: {e}")
    comp.CodeModule.AddFromString(code)
    comp.Name = final_name
    print(f"  + form {final_name}")


def build_login(vbproj):
    controls = [
        {"progid": "Forms.Label.1", "name": "lblTitle", "left": 20, "top": 12, "width": 340, "height": 24,
         "caption": "Connexion — Prévision M5", "font_size": 14, "bold": True},
        {"progid": "Forms.Label.1", "name": "lblU", "left": 20, "top": 54, "width": 90, "height": 20,
         "caption": "Utilisateur"},
        {"progid": "Forms.TextBox.1", "name": "txtUser", "left": 120, "top": 52, "width": 220, "height": 22},
        {"progid": "Forms.Label.1", "name": "lblP", "left": 20, "top": 86, "width": 90, "height": 20,
         "caption": "Mot de passe"},
        {"progid": "Forms.TextBox.1", "name": "txtPass", "left": 120, "top": 84, "width": 220, "height": 22},
        {"progid": "Forms.CommandButton.1", "name": "btnOK", "left": 120, "top": 130, "width": 100,
         "height": 28, "caption": "Se connecter", "props": {"Default": True}},
        {"progid": "Forms.CommandButton.1", "name": "btnCancel", "left": 240, "top": 130, "width": 100,
         "height": 28, "caption": "Annuler", "props": {"Cancel": True}},
        {"progid": "Forms.Label.1", "name": "lblErr", "left": 20, "top": 170, "width": 340, "height": 20,
         "caption": "", "props": {"ForeColor": 192}},
    ]
    code = r'''Option Explicit
Private mTries As Long

Private Sub UserForm_Initialize()
    Me.txtPass.PasswordChar = "*"
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnOK_Click()
    Dim role As String, store As String, u As String
    u = LCase$(Trim$(Me.txtUser.Value))
    If Authenticate(u, CStr(Me.txtPass.Value), role, store) Then
        Unload Me
        SetSession role, store, u
        MsgBox "Connecté : " & u & " (profil " & role & ", magasin " & store & ")", vbInformation
    Else
        mTries = mTries + 1
        Me.lblErr.Caption = "Identifiants invalides (" & mTries & "/3)."
        Me.txtPass.Value = ""
        If mTries >= 3 Then Unload Me
    End If
End Sub
'''
    create_form(vbproj, "frmLogin", "Connexion", 400, 230, controls, code)


def vb_array(items):
    return "Array(" + ", ".join(f'"{x}"' for x in items) + ")"


def build_crud_form(vbproj, name, sheet, headers, keys):
    fields = [h for h in headers if h not in AUTO_COLS]
    n = len(fields)
    ncols = 3 if n > 9 else 2
    block = 245
    width = 20 + ncols * block
    list_w = width - 30
    fields_top = 200
    row_h = 28
    controls = [
        {"progid": "Forms.Label.1", "name": "lblT", "left": 10, "top": 8, "width": list_w, "height": 20,
         "caption": f"{name} — {sheet}", "font_size": 12, "bold": True},
        {"progid": "Forms.Label.1", "name": "lblS", "left": 10, "top": 36, "width": 130, "height": 18,
         "caption": "Rechercher (clé) :"},
        {"progid": "Forms.TextBox.1", "name": "txtSearch", "left": 140, "top": 34, "width": 200, "height": 20},
        {"progid": "Forms.Label.1", "name": "lblInfo", "left": 350, "top": 36, "width": max(list_w - 340, 120),
         "height": 18, "caption": "", "font_size": 9},
        {"progid": "Forms.ListBox.1", "name": "lst", "left": 10, "top": 60, "width": list_w, "height": 125},
    ]
    for i, h in enumerate(fields):
        r, c = divmod(i, ncols)
        x, y = 10 + c * block, fields_top + r * row_h
        controls.append({"progid": "Forms.Label.1", "name": f"lbl{i}", "left": x, "top": y + 2,
                         "width": 122, "height": 16, "caption": h, "font_size": 9})
        controls.append({"progid": "Forms.TextBox.1", "name": f"txt{i}", "left": x + 124, "top": y,
                         "width": 100, "height": 20})
    btn_top = fields_top + math.ceil(n / ncols) * row_h + 10
    for j, (nm, cap) in enumerate([("btnAdd", "Ajouter"), ("btnUpd", "Modifier"),
                                   ("btnDel", "Supprimer"), ("btnClose", "Fermer")]):
        controls.append({"progid": "Forms.CommandButton.1", "name": nm, "left": 10 + j * 100,
                         "top": btn_top, "width": 92, "height": 26, "caption": cap})

    code = f'''Option Explicit
Private Const SH As String = "{sheet}"
Private mFields As Variant
Private mKeys As Variant

Private Sub UserForm_Initialize()
    mFields = {vb_array(fields)}
    mKeys = {vb_array(keys)}
    RefreshList
End Sub

Private Sub RefreshList()
    Dim n As Long
    n = FillRecordList(Me.lst, SH, mKeys, CStr(Me.txtSearch.Value))
    Me.lblInfo.Caption = n & " ligne(s) affichée(s) (max " & MAX_LIST_ROWS & ")"
End Sub

Private Function SelectedRow() As Long
    SelectedRow = 0
    If Me.lst.ListIndex >= 0 Then
        SelectedRow = CLng(Me.lst.List(Me.lst.ListIndex, Me.lst.ColumnCount - 1))
    End If
End Function

Private Sub txtSearch_Change()
    RefreshList
End Sub

Private Sub lst_Click()
    If SelectedRow() > 0 Then LoadRecord Me, SH, SelectedRow(), mFields
End Sub

Private Sub btnAdd_Click()
    If AddRecord(Me, SH, mFields, mKeys) Then
        ClearFields Me, UBound(mFields) + 1
        RefreshList
    End If
End Sub

Private Sub btnUpd_Click()
    If SelectedRow() = 0 Then
        MsgBox "Sélectionnez d'abord une ligne dans la liste.", vbExclamation
        Exit Sub
    End If
    If UpdateRecord(Me, SH, mFields, mKeys, SelectedRow()) Then RefreshList
End Sub

Private Sub btnDel_Click()
    If SelectedRow() = 0 Then
        MsgBox "Sélectionnez d'abord une ligne dans la liste.", vbExclamation
        Exit Sub
    End If
    If DeleteRecord(SH, SelectedRow()) Then
        ClearFields Me, UBound(mFields) + 1
        RefreshList
    End If
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub
'''
    create_form(vbproj, "frm" + name, name, width, btn_top + 70, controls, code)


# ------------------------------------------------------------------ verrouillage final
def lock_fallback(wb):
    """Equivalent Python de LockWorkbook (si l'execution de la macro echoue)."""
    wb.Sheets("Accueil").Activate()
    for ws in wb.Worksheets:
        if ws.Name != "Accueil":
            ws.Visible = 2                       # xlSheetVeryHidden
    for ws in wb.Worksheets:
        ws.Protect(PWD, True, True, True, True, False, True, True)
    wb.Protect(PWD, True, False)


# ------------------------------------------------------------------ MAIN
def main():
    if not XLSX.exists():
        print(f"[ERREUR] {XLSX.name} absent. Lancez d'abord generate_app.py.")
        return
    missing = [m for m in REQUIRED if not (VBA_DIR / f"{m}.bas").exists()]
    if missing:
        print(f"[ERREUR] Modules manquants dans vba/ : {missing}")
        return
    if XLSM.exists():
        try:
            XLSM.unlink()
        except Exception as e:
            print(f"[ERREUR] Impossible de remplacer {XLSM.name} (fermez-le dans Excel) : {e}")
            return

    excel = None
    wb = None
    try:
        excel = win32.gencache.EnsureDispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        wb = excel.Workbooks.Open(str(XLSX))
        wb.SaveAs(str(XLSM), FileFormat=52)        # 52 = xlOpenXMLWorkbookMacroEnabled
        print(f"✔ {XLSM.name} créé (format macro)")

        try:
            vbproj = wb.VBProject
            _ = vbproj.VBComponents.Count
        except Exception as e:
            print("[ERREUR] Accès au projet VBA refusé :", e)
            print("  Excel > Fichier > Options > Centre de gestion de la confidentialité >")
            print("  Paramètres des macros > cocher « Accès approuvé au modèle d'objet du projet VBA ».")
            raise SystemExit(1)

        n = inject_all_modules(vbproj)
        print(f"  → {n}/{len(REQUIRED)} modules")
        if n != len(REQUIRED):
            raise RuntimeError("Tous les modules VBA n'ont pas pu être injectés.")

        set_workbook_events(vbproj)

        build_login(vbproj)
        for name, (sheet, keys) in CRUD_FORMS.items():
            lo = wb.Sheets(sheet).ListObjects(1)
            headers = [str(h) for h in lo.HeaderRowRange.Value[0]]
            build_crud_form(vbproj, name, sheet, headers, keys)
            lo.ShowAutoFilter = False              # aucun filtre manuel : tableaux seuls

        # Accueil + Gestion + bouton « Accueil » sur toutes les autres feuilles (voir generate_app.BUTTONS_BY_SHEET)
        for sheet_name, specs in G.BUTTONS_BY_SHEET.items():
            add_buttons(wb, sheet_name, specs)
        wb.Sheets("Accueil").Activate()

        excel.Calculation = -4105
        excel.CalculateFull()

        # Etat verrouille a la livraison : seule "Accueil" visible, session vide
        try:
            excel.Run("LockWorkbook")
            print("✔ LockWorkbook exécuté (code VBA validé)")
        except Exception as e:
            print(f"  ! LockWorkbook (macro) a échoué : {e}\n    → verrouillage Python de secours")
            lock_fallback(wb)
        vis = {ws.Name: ws.Visible for ws in wb.Worksheets}
        print("  feuilles visibles :", [k for k, v in vis.items() if v == -1])

        wb.Save()
        print("✔ Terminé. Ouvrez Application_Prevision.xlsm (activez les macros).")
        print("  Comptes : admin/admin2024 · hq_analyst/hq2024 · tx_manager/tx2024 · "
              "ca_manager/ca2024 · wi_manager/wi2024")
    except SystemExit:
        pass
    except Exception as e:
        print(f"[ERREUR FATALE] {e}")
        traceback.print_exc()
    finally:
        if wb:
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass
        if excel:
            try:
                excel.Quit()
            except Exception:
                pass
        print("Excel fermé.")


if __name__ == "__main__":
    main()