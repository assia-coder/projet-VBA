VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMetriques
   Caption         =   "Comparaison des modèles"
   ClientHeight    =   7900
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmMetriques"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==============================================================================
' frmMetriques - Comparaison des modèles journaliers (période de test, 180 jours)
' Source : plage nommée PlgMetriques (feuille METRIQUES, fichier metrics_all.csv)
'==============================================================================
Option Explicit

Private WithEvents cmbNiveau As MSForms.ComboBox
Private WithEvents cmbStore As MSForms.ComboBox
Private lstMetriques As MSForms.ListBox
Private lblMeilleur As MSForms.Label
Private WithEvents btnCalculerVBA As MSForms.CommandButton
Private WithEvents btnFermer As MSForms.CommandButton
Private mbPret As Boolean          ' évite les rechargements pendant la construction


' -- Construction de l'écran --------------------------------------------------
Private Sub UserForm_Initialize()
    Dim i As Long, entetes As Variant, x As Variant
    Me.Caption = "Comparaison des modèles"
    Me.Width = 560
    Me.Height = 420
    Me.BackColor = RGB(255, 255, 255)

    AjouterCtl Me, "Label", "lblNiveau", 14, 14, 40, 14, "Niveau"
    Set cmbNiveau = AjouterCtl(Me, "ComboBox", "cmbNiveau", 56, 11, 130, 18)
    AjouterCtl Me, "Label", "lblStore", 204, 14, 50, 14, "Magasin"
    Set cmbStore = AjouterCtl(Me, "ComboBox", "cmbStore", 254, 11, 90, 18)
    cmbNiveau.Style = fmStyleDropDownList
    cmbStore.Style = fmStyleDropDownList

    ' En-têtes de colonnes
    entetes = Array("Modèle", "MAE", "MAPE %", "RMSE", "WAPE %", "WRMSSE")
    x = Array(16, 236, 296, 356, 416, 476)
    For i = 0 To 5
        With AjouterCtl(Me, "Label", "lblEntete" & i, CSng(x(i)), 40, 70, 14, CStr(entetes(i)))
            .Font.Bold = True
        End With
    Next i

    Set lstMetriques = AjouterCtl(Me, "ListBox", "lstMetriques", 14, 56, 522, 250)
    lstMetriques.ColumnCount = 6
    lstMetriques.ColumnWidths = "220;60;60;60;60;60"

    Set lblMeilleur = AjouterCtl(Me, "Label", "lblMeilleur", 14, 312, 522, 30)
    lblMeilleur.Font.Bold = True
    lblMeilleur.ForeColor = RGB(0, 97, 0)
    lblMeilleur.WordWrap = True

    Set btnCalculerVBA = AjouterCtl(Me, "CommandButton", "btnCalculerVBA", 14, 350, 220, 26, _
                                    "Recalculer les métriques en VBA")
    Set btnFermer = AjouterCtl(Me, "CommandButton", "btnFermer", 436, 350, 100, 26, "Fermer")

    cmbNiveau.AddItem "Global": cmbNiveau.AddItem "Par magasin"
    cmbNiveau.AddItem "Par produit": cmbNiveau.AddItem "Par catégorie"
    For Each x In ListeParam("B2", "CA_1,TX_1,WI_1")
        cmbStore.AddItem x
    Next x
    If cmbStore.ListCount > 0 Then cmbStore.ListIndex = 0
    cmbNiveau.ListIndex = 0
    mbPret = True
    cmbNiveau_Change
End Sub

Private Sub cmbNiveau_Change()
    If Not mbPret Then Exit Sub
    cmbStore.Enabled = (cmbNiveau.Value = "Par magasin" Or cmbNiveau.Value = "Par produit")
    ChargerMetriques
End Sub

Private Sub cmbStore_Change()
    If Not mbPret Then Exit Sub
    ChargerMetriques
End Sub

Private Function F(ByVal x As Variant, ByVal fmtTxt As String) As String
    If IsNumeric(x) And Not IsEmpty(x) And CStr(x) <> "" Then F = Format(x, fmtTxt) Else F = "-"
End Function


' Filtre PlgMetriques selon le niveau et le magasin, remplit la liste.
' Le meilleur modèle (RMSE minimal) de chaque groupe est marqué "» ".
Private Sub ChargerMetriques()
    Dim v As Variant, i As Long, n As Long, niveau As String, store As String
    Dim cMo As Long, cAp As Long, cLe As Long, cSt As Long, cIt As Long
    Dim cMae As Long, cMape As Long, cRmse As Long, cWape As Long, cWr As Long
    Dim groupe As String, nom As String, lignes As Object, meilleurs As Object, victoires As Object
    Dim rmse As Double, cle As Variant, best As String, nbBest As Long, nbGroupes As Long

    lstMetriques.Clear
    v = LirePlage("PlgMetriques")
    If Not IsArray(v) Then
        lblMeilleur.Caption = "Données non disponibles : lancez l'entraînement."
        Exit Sub
    End If
    cMo = ColIndex(v, "Modele"): cAp = ColIndex(v, "Approche"): cLe = ColIndex(v, "Level")
    cSt = ColIndex(v, "Store"): cIt = ColIndex(v, "Item")
    cMae = ColIndex(v, "MAE"): cMape = ColIndex(v, "MAPE"): cRmse = ColIndex(v, "RMSE")
    cWape = ColIndex(v, "WAPE"): cWr = ColIndex(v, "WRMSSE")

    Select Case cmbNiveau.Value
        Case "Global": niveau = "global": store = "ALL"
        Case "Par magasin": niveau = "store": store = cmbStore.Value
        Case "Par produit": niveau = "item": store = cmbStore.Value
        Case Else: niveau = "category": store = "ALL"
    End Select

    ' 1er passage : meilleur RMSE de chaque groupe (produit ou catégorie)
    Set meilleurs = CreateObject("Scripting.Dictionary")
    For i = 2 To UBound(v, 1)
        If CStr(v(i, cLe)) = niveau And CStr(v(i, cSt)) = store And IsNumeric(v(i, cRmse)) Then
            groupe = CStr(v(i, cIt))
            rmse = CDbl(v(i, cRmse))
            If Not meilleurs.Exists(groupe) Then
                meilleurs(groupe) = i
            ElseIf rmse < CDbl(v(meilleurs(groupe), cRmse)) Then
                meilleurs(groupe) = i
            End If
        End If
    Next i
    If meilleurs.Count = 0 Then
        lblMeilleur.Caption = "Données non disponibles pour ce niveau."
        Exit Sub
    End If

    ' 2e passage : remplissage de la liste
    For i = 2 To UBound(v, 1)
        If CStr(v(i, cLe)) = niveau And CStr(v(i, cSt)) = store Then
            groupe = CStr(v(i, cIt))
            nom = v(i, cMo) & " (" & v(i, cAp) & ")"
            If niveau = "item" Or niveau = "category" Then nom = groupe & " - " & nom
            If meilleurs.Exists(groupe) Then
                If meilleurs(groupe) = i Then nom = "» " & nom Else nom = "   " & nom
            End If
            lstMetriques.AddItem nom
            lstMetriques.List(n, 1) = F(v(i, cMae), "0.000")
            lstMetriques.List(n, 2) = F(v(i, cMape), "0.0")
            lstMetriques.List(n, 3) = F(v(i, cRmse), "0.000")
            lstMetriques.List(n, 4) = F(v(i, cWape), "0.0")
            If cWr > 0 Then lstMetriques.List(n, 5) = F(v(i, cWr), "0.000")
            n = n + 1
        End If
    Next i

    ' Synthèse : meilleur modèle (ou modèle gagnant le plus souvent)
    If meilleurs.Count = 1 Then
        i = meilleurs.Items()(0)
        lblMeilleur.Caption = "Meilleur modèle : " & v(i, cMo) & " (" & v(i, cAp) & ") - RMSE " & _
                              Format(v(i, cRmse), "0.000")
    Else
        Set victoires = CreateObject("Scripting.Dictionary")
        For Each cle In meilleurs.Keys
            i = meilleurs(cle)
            nom = v(i, cMo) & " (" & v(i, cAp) & ")"
            victoires(nom) = victoires(nom) + 1
        Next cle
        For Each cle In victoires.Keys
            If victoires(cle) > nbBest Then nbBest = victoires(cle): best = cle
        Next cle
        nbGroupes = meilleurs.Count
        lblMeilleur.Caption = "» = meilleur RMSE du groupe. " & best & " gagne " & nbBest & " / " & _
                              nbGroupes & " groupes."
    End If
    If Len(Param("B6")) > 0 Then lblMeilleur.Caption = lblMeilleur.Caption & "  Gagnant global : " & Param("B6") & "."
End Sub


' -- Actions ------------------------------------------------------------------
Private Sub btnCalculerVBA_Click()
    Dim ok As Boolean
    ok = ModuleApp.CalculerMetriquesVBA()
    ChargerMetriques
    lblMeilleur.Caption = lblMeilleur.Caption & IIf(ok, " [VBA validé]", " [écart VBA : voir METRIQUES]")
End Sub

Private Sub btnFermer_Click()
    Unload Me
End Sub
