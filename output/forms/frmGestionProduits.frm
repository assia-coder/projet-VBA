VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmGestionProduits
   Caption         =   "Gestion des produits et des ventes"
   ClientHeight    =   9500
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10200
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmGestionProduits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==============================================================================
' frmGestionProduits - Saisie de nouvelles ventes et de nouveaux produits
'   Section A : ventes -> PlgNouvellesVentes (feuille NOUVELLES_VENTES)
'   Section B : produits -> PlgProduits (feuille PRODUITS) + output/produits.csv
'   Bas       : fusion des ventes + réentraînement complet ("all")
'==============================================================================
Option Explicit

' --- Section A : ajouter des ventes ---
Private fraVentes As MSForms.Frame
Private txtDate As MSForms.TextBox
Private WithEvents cmbStore As MSForms.ComboBox
Private cmbItem As MSForms.ComboBox
Private txtSales As MSForms.TextBox
Private txtPrix As MSForms.TextBox
Private chkSnap As MSForms.CheckBox
Private chkEvent As MSForms.CheckBox
Private WithEvents btnAjouterVente As MSForms.CommandButton

' --- Section B : ajouter un produit ---
Private fraProduit As MSForms.Frame
Private txtNouveauItem As MSForms.TextBox
Private cmbNouveauStore As MSForms.ComboBox
Private cmbNouveauCat As MSForms.ComboBox
Private WithEvents btnAjouterProduit As MSForms.CommandButton

' --- Bas ---
Private lblNbVentes As MSForms.Label
Private WithEvents btnLancerEntrainement As MSForms.CommandButton
Private WithEvents btnFermer As MSForms.CommandButton


' -- Construction de l'écran --------------------------------------------------
Private Sub UserForm_Initialize()
    Dim x As Variant, d As Date
    Me.Caption = "Gestion des produits et des ventes"
    Me.Width = 520
    Me.Height = 500
    Me.BackColor = RGB(255, 255, 255)

    ' Section A
    Set fraVentes = AjouterCtl(Me, "Frame", "fraVentes", 12, 8, 490, 196, "A - Ajouter des ventes")
    fraVentes.BackColor = RGB(255, 255, 255)
    AjouterCtl fraVentes, "Label", "lblDate", 12, 16, 100, 14, "Date (JJ/MM/AAAA)"
    Set txtDate = AjouterCtl(fraVentes, "TextBox", "txtDate", 120, 13, 90, 18)
    AjouterCtl fraVentes, "Label", "lblStore", 240, 16, 50, 14, "Magasin"
    Set cmbStore = AjouterCtl(fraVentes, "ComboBox", "cmbStore", 300, 13, 90, 18)
    AjouterCtl fraVentes, "Label", "lblItem", 12, 44, 100, 14, "Produit"
    Set cmbItem = AjouterCtl(fraVentes, "ComboBox", "cmbItem", 120, 41, 170, 18)
    AjouterCtl fraVentes, "Label", "lblSales", 12, 72, 100, 14, "Ventes (unités)"
    Set txtSales = AjouterCtl(fraVentes, "TextBox", "txtSales", 120, 69, 90, 18)
    AjouterCtl fraVentes, "Label", "lblPrix", 240, 72, 50, 14, "Prix ($)"
    Set txtPrix = AjouterCtl(fraVentes, "TextBox", "txtPrix", 300, 69, 90, 18)
    Set chkSnap = AjouterCtl(fraVentes, "CheckBox", "chkSnap", 120, 98, 110, 18, "Jour SNAP")
    Set chkEvent = AjouterCtl(fraVentes, "CheckBox", "chkEvent", 240, 98, 150, 18, "Événement ce jour-là")
    Set btnAjouterVente = AjouterCtl(fraVentes, "CommandButton", "btnAjouterVente", 120, 128, 270, 26, _
                                     "Ajouter la vente")
    AjouterCtl fraVentes, "Label", "lblAideVente", 12, 160, 470, 24, _
               "Les ventes sont mises en attente dans NOUVELLES_VENTES, puis fusionnées à l'historique " & _
               "au prochain entraînement."
    cmbStore.Style = fmStyleDropDownList
    cmbItem.Style = fmStyleDropDownList

    ' Section B
    Set fraProduit = AjouterCtl(Me, "Frame", "fraProduit", 12, 212, 490, 124, "B - Ajouter un produit")
    fraProduit.BackColor = RGB(255, 255, 255)
    AjouterCtl fraProduit, "Label", "lblNouveauItem", 12, 16, 150, 14, "Identifiant (ex. FOODS_3_090)"
    Set txtNouveauItem = AjouterCtl(fraProduit, "TextBox", "txtNouveauItem", 170, 13, 150, 18)
    AjouterCtl fraProduit, "Label", "lblNouveauStore", 12, 44, 100, 14, "Magasin"
    Set cmbNouveauStore = AjouterCtl(fraProduit, "ComboBox", "cmbNouveauStore", 170, 41, 90, 18)
    AjouterCtl fraProduit, "Label", "lblNouveauCat", 276, 44, 60, 14, "Catégorie"
    Set cmbNouveauCat = AjouterCtl(fraProduit, "ComboBox", "cmbNouveauCat", 340, 41, 120, 18)
    Set btnAjouterProduit = AjouterCtl(fraProduit, "CommandButton", "btnAjouterProduit", 170, 72, 220, 26, _
                                       "Ajouter le produit")
    cmbNouveauStore.Style = fmStyleDropDownList
    cmbNouveauCat.Style = fmStyleDropDownList

    ' Bas
    Set lblNbVentes = AjouterCtl(Me, "Label", "lblNbVentes", 12, 348, 490, 16)
    lblNbVentes.Font.Bold = True
    Set btnLancerEntrainement = AjouterCtl(Me, "CommandButton", "btnLancerEntrainement", 12, 372, 340, 32, _
                                           "Fusionner les ventes et relancer l'entraînement")
    btnLancerEntrainement.BackColor = RGB(192, 0, 0)            ' #C00000
    btnLancerEntrainement.ForeColor = RGB(255, 255, 255)
    btnLancerEntrainement.Font.Bold = True
    Set btnFermer = AjouterCtl(Me, "CommandButton", "btnFermer", 402, 372, 100, 32, "Fermer")

    ' Listes depuis PARAMS (B2 stores, B3 catégories)
    For Each x In ListeParam("B2", "CA_1,TX_1,WI_1")
        cmbStore.AddItem x
        cmbNouveauStore.AddItem x
    Next x
    For Each x In ListeParam("B3", "FOODS,HOBBIES,HOUSEHOLD")
        cmbNouveauCat.AddItem x
    Next x
    If cmbNouveauStore.ListCount > 0 Then cmbNouveauStore.ListIndex = 0
    If cmbNouveauCat.ListCount > 0 Then cmbNouveauCat.ListIndex = 0
    If cmbStore.ListCount > 0 Then cmbStore.ListIndex = 0      ' déclenche cmbStore_Change -> cmbItem

    ' Date proposée : lendemain de la dernière date du dataset
    If IsDate(Param("B4")) Then
        d = CDate(Param("B4")) + 1
        txtDate.Value = Format(d, "dd/mm/yyyy")
    End If
    MajNbVentes
End Sub

' Recharge cmbItem : produits actifs du magasin sélectionné (PlgProduits)
Private Sub cmbStore_Change()
    Dim v As Variant, i As Long, cS As Long, cI As Long, cA As Long
    If cmbItem Is Nothing Then Exit Sub
    cmbItem.Clear
    v = LirePlage("PlgProduits")
    If Not IsArray(v) Then Exit Sub
    cS = ColIndex(v, "store_id"): cI = ColIndex(v, "item_id"): cA = ColIndex(v, "actif")
    For i = 2 To UBound(v, 1)
        If CStr(v(i, cS)) = cmbStore.Value Then
            If cA = 0 Then
                cmbItem.AddItem CStr(v(i, cI))
            ElseIf UCase$(CStr(v(i, cA))) <> "FALSE" And UCase$(CStr(v(i, cA))) <> "FAUX" Then
                cmbItem.AddItem CStr(v(i, cI))
            End If
        End If
    Next i
    If cmbItem.ListCount > 0 Then cmbItem.ListIndex = 0
End Sub

Private Function NbNouvellesVentes() As Long
    NbNouvellesVentes = NbLignesPlage("PlgNouvellesVentes")
End Function

Private Sub MajNbVentes()
    lblNbVentes.Caption = NbNouvellesVentes() & " vente(s) en attente de fusion."
End Sub


' -- Section A : ajouter une vente --------------------------------------------
Private Sub btnAjouterVente_Click()
    Dim d As Date, ventes As Double, prix As Double, item As String

    ' Validation champ par champ
    If Not LireDateJJMMAAAA(txtDate.Value, d) Then
        MsgBox "Date invalide : format attendu JJ/MM/AAAA (ex. 25/04/2016).", vbExclamation, "Date"
        txtDate.SetFocus
        Exit Sub
    End If
    If cmbStore.ListIndex < 0 Then
        MsgBox "Magasin : choisissez un magasin dans la liste.", vbExclamation, "Magasin"
        Exit Sub
    End If
    If cmbItem.ListIndex < 0 Then
        MsgBox "Produit : aucun produit suivi pour ce magasin (section B pour en ajouter).", vbExclamation, "Produit"
        Exit Sub
    End If
    If Not LireNombre(txtSales.Value, ventes) Then
        MsgBox "Ventes : nombre >= 0 attendu.", vbExclamation, "Ventes"
        txtSales.SetFocus
        Exit Sub
    End If
    If Not LireNombre(txtPrix.Value, prix) Then
        MsgBox "Prix : nombre >= 0 attendu (ex. 2,98).", vbExclamation, "Prix"
        txtPrix.SetFocus
        Exit Sub
    End If

    ' Ajout dans PlgNouvellesVentes : date, store_id, item_id, cat_id, sales, sell_price, snap, has_event
    item = cmbItem.Value
    AjouterLigne "PlgNouvellesVentes", Array(d, cmbStore.Value, item, CategorieDe(item), ventes, prix, _
                                             IIf(chkSnap.Value, 1, 0), IIf(chkEvent.Value, 1, 0))

    ' Réinitialisation (date suivante proposée pour enchaîner les saisies)
    txtDate.Value = Format(d + 1, "dd/mm/yyyy")
    txtSales.Value = ""
    txtPrix.Value = ""
    chkSnap.Value = False
    chkEvent.Value = False
    MajNbVentes
End Sub


' -- Section B : ajouter un produit -------------------------------------------
Private Sub btnAjouterProduit_Click()
    Dim re As Object, item As String, store As String, v As Variant, i As Long
    item = UCase$(Trim$(txtNouveauItem.Value))
    store = cmbNouveauStore.Value

    ' Format FOODS_x_xxx, HOBBIES_x_xxx ou HOUSEHOLD_x_xxx
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "^(FOODS|HOBBIES|HOUSEHOLD)_\d+_\d{3}$"
    If Not re.Test(item) Then
        MsgBox "Identifiant invalide : format FOODS_x_xxx, HOBBIES_x_xxx ou HOUSEHOLD_x_xxx.", vbExclamation
        Exit Sub
    End If
    If CategorieDe(item) <> cmbNouveauCat.Value Then
        MsgBox "La catégorie choisie (" & cmbNouveauCat.Value & ") ne correspond pas à l'identifiant (" & _
               CategorieDe(item) & ").", vbExclamation
        Exit Sub
    End If

    ' Doublon store + produit
    v = LirePlage("PlgProduits")
    If IsArray(v) Then
        For i = 2 To UBound(v, 1)
            If CStr(v(i, ColIndex(v, "store_id"))) = store And CStr(v(i, ColIndex(v, "item_id"))) = item Then
                MsgBox item & " est déjà suivi dans " & store & ".", vbInformation
                Exit Sub
            End If
        Next i
    End If

    ' Ajout : store_id, item_id, cat_id, date_ajout, actif
    AjouterLigne "PlgProduits", Array(store, item, CategorieDe(item), Date, True)
    ExporterProduits
    txtNouveauItem.Value = ""
    MsgBox item & " ajouté pour " & store & "." & vbCrLf & _
           "Saisissez ses ventes (section A) puis relancez l'entraînement.", vbInformation
    If cmbStore.Value = store Then cmbStore_Change          ' recharge cmbItem
End Sub


' -- Bas : fusion + entraînement ----------------------------------------------
Private Sub btnLancerEntrainement_Click()
    Dim n As Long
    n = NbNouvellesVentes()
    If MsgBox(n & " nouvelles ventes seront fusionnées. Lancer l'entraînement ?", _
              vbYesNo + vbQuestion, "Entraînement") = vbYes Then
        ExporterNouvellesVentes                 ' -> output/nouvelles_ventes.csv (lu par preprocess.py)
        FusionnerNouvellesVentesDonnees         ' PlgNouvellesVentes -> PlgDonnees (dédup date+store+item)
        lblNbVentes.Caption = "Entraînement en cours... (barre d'état d'Excel pour le détail)"
        btnLancerEntrainement.Enabled = False
        DoEvents
        If ModuleApp.LancerEntrainement("all") Then
            ViderNouvellesVentes                ' en-têtes conservés, CSV supprimé
            MsgBox "Entraînement terminé : " & n & " vente(s) intégrée(s).", vbInformation
        Else
            MsgBox "Entraînement interrompu : les ventes restent en attente.", vbExclamation
        End If
        btnLancerEntrainement.Enabled = True
        MajNbVentes
    End If
End Sub

Private Sub btnFermer_Click()
    Unload Me
End Sub
