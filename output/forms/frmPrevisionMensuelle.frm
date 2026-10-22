VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPrevisionMensuelle
   Caption         =   "Prévisions du mois prochain"
   ClientHeight    =   7100
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9400
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmPrevisionMensuelle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==============================================================================
' frmPrevisionMensuelle - Prévisions M+1, M+2, M+3 par magasin et catégorie
' Source : plage nommée PlgMensuel (feuille PREVISION_MENSUELLE)
'==============================================================================
Option Explicit

Private WithEvents cmbStore As MSForms.ComboBox
Private WithEvents cmbCategorie As MSForms.ComboBox
Private lstResultats As MSForms.ListBox
Private WithEvents btnAfficher As MSForms.CommandButton
Private WithEvents btnExporter As MSForms.CommandButton
Private WithEvents btnFermer As MSForms.CommandButton
Private lblStatut As MSForms.Label
Private msMois As String          ' mois M+1 affiché (aaaa-mm), utilisé pour l'export


' -- Construction de l'écran --------------------------------------------------
Private Sub UserForm_Initialize()
    Dim v As Variant, i As Long, entetes As Variant, x As Variant
    Me.Caption = "Prévisions du mois prochain"
    Me.Width = 480
    Me.Height = 380
    Me.BackColor = RGB(255, 255, 255)

    AjouterCtl Me, "Label", "lblStore", 14, 14, 50, 14, "Magasin"
    Set cmbStore = AjouterCtl(Me, "ComboBox", "cmbStore", 62, 11, 80, 18)
    AjouterCtl Me, "Label", "lblCategorie", 156, 14, 56, 14, "Catégorie"
    Set cmbCategorie = AjouterCtl(Me, "ComboBox", "cmbCategorie", 212, 11, 110, 18)
    Set btnAfficher = AjouterCtl(Me, "CommandButton", "btnAfficher", 340, 9, 116, 22, "Afficher")
    cmbStore.Style = fmStyleDropDownList
    cmbCategorie.Style = fmStyleDropDownList

    ' En-têtes de colonnes (la ListBox n'affiche pas d'en-têtes sans RowSource)
    entetes = Array("Produit", "M+1", "Intervalle M+1", "M+2", "M+3")
    x = Array(16, 116, 186, 316, 386)
    For i = 0 To 4
        With AjouterCtl(Me, "Label", "lblEntete" & i, CSng(x(i)), 40, 90, 14, CStr(entetes(i)))
            .Font.Bold = True
        End With
    Next i

    Set lstResultats = AjouterCtl(Me, "ListBox", "lstResultats", 14, 56, 442, 222)
    lstResultats.ColumnCount = 5
    lstResultats.ColumnWidths = "100;70;130;70;60"

    Set lblStatut = AjouterCtl(Me, "Label", "lblStatut", 14, 284, 442, 28)
    lblStatut.WordWrap = True
    lblStatut.ForeColor = RGB(80, 80, 80)

    Set btnExporter = AjouterCtl(Me, "CommandButton", "btnExporter", 14, 318, 150, 24, "Exporter en CSV")
    Set btnFermer = AjouterCtl(Me, "CommandButton", "btnFermer", 356, 318, 100, 24, "Fermer")

    ' Listes depuis PARAMS (B2 stores, B3 catégories) ; CA_1 / FOODS par défaut
    For Each x In ListeParam("B2", "CA_1,TX_1,WI_1")
        cmbStore.AddItem x
    Next x
    For Each x In ListeParam("B3", "FOODS,HOBBIES,HOUSEHOLD")
        cmbCategorie.AddItem x
    Next x
    SelectionnerValeur cmbStore, "CA_1"
    SelectionnerValeur cmbCategorie, "FOODS"

    v = LirePlage("PlgMensuel")
    If IsArray(v) Then
        btnAfficher_Click
    Else
        lblStatut.Caption = "Données non disponibles : lancez la prévision mensuelle " & _
                            "(python excel_bridge.py monthly)."
    End If
End Sub

Private Sub SelectionnerValeur(ByVal cmb As MSForms.ComboBox, ByVal valeur As String)
    Dim i As Long
    For i = 0 To cmb.ListCount - 1
        If cmb.List(i) = valeur Then cmb.ListIndex = i: Exit Sub
    Next i
    If cmb.ListCount > 0 Then cmb.ListIndex = 0
End Sub

' Mois au format aaaa-mm (Excel convertit parfois "2016-05" en date)
Private Function TexteMois(ByVal x As Variant) As String
    If VarType(x) = vbDate Then TexteMois = Format(x, "yyyy-mm") Else TexteMois = CStr(x)
End Function

Private Function N0(ByVal x As Variant) As String
    If IsNumeric(x) And Not IsEmpty(x) Then N0 = Format(x, "0") Else N0 = "-"
End Function


' -- Actions ------------------------------------------------------------------
' Filtre PlgMensuel sur le magasin et la catégorie, remplit la liste
Private Sub btnAfficher_Click()
    Dim v As Variant, i As Long, n As Long, total As Double
    Dim cS As Long, cC As Long, cI As Long, cM As Long
    Dim c1 As Long, c1l As Long, c1h As Long, c2 As Long, c3 As Long

    lstResultats.Clear
    msMois = ""
    v = LirePlage("PlgMensuel")
    If Not IsArray(v) Then
        lblStatut.Caption = "Données non disponibles."
        Exit Sub
    End If
    cM = ColIndex(v, "month"): cS = ColIndex(v, "store_id"): cC = ColIndex(v, "cat_id")
    cI = ColIndex(v, "item_id"): c1 = ColIndex(v, "h1_pred")
    c1l = ColIndex(v, "h1_low"): c1h = ColIndex(v, "h1_high")
    c2 = ColIndex(v, "h2_pred"): c3 = ColIndex(v, "h3_pred")
    If cS = 0 Or cI = 0 Or c1 = 0 Then
        lblStatut.Caption = "Données non disponibles (colonnes attendues absentes)."
        Exit Sub
    End If

    For i = 2 To UBound(v, 1)
        If CStr(v(i, cS)) = cmbStore.Value And CStr(v(i, cC)) = cmbCategorie.Value Then
            lstResultats.AddItem CStr(v(i, cI))
            lstResultats.List(n, 1) = N0(v(i, c1))
            If c1l > 0 And c1h > 0 Then lstResultats.List(n, 2) = "[" & N0(v(i, c1l)) & " - " & N0(v(i, c1h)) & "]"
            If c2 > 0 Then lstResultats.List(n, 3) = N0(v(i, c2))
            If c3 > 0 Then lstResultats.List(n, 4) = N0(v(i, c3))
            If IsNumeric(v(i, c1)) Then total = total + CDbl(v(i, c1))
            If msMois = "" And cM > 0 Then msMois = TexteMois(v(i, cM))
            n = n + 1
        End If
    Next i

    If n = 0 Then
        lblStatut.Caption = "Aucune prévision pour " & cmbStore.Value & " / " & cmbCategorie.Value & "."
    Else
        lblStatut.Caption = "Mois M+1 : " & msMois & ". Total : " & Format(total, "#,##0") & " unités (" & _
                            n & " produits). Intervalle de prévision à 80 %."
    End If
End Sub

' Exporte la liste affichée en CSV (séparateur ";") dans le dossier du classeur
Private Sub btnExporter_Click()
    Dim t() As Variant, i As Long, j As Long, chemin As String
    If lstResultats.ListCount = 0 Then
        lblStatut.Caption = "Rien à exporter : cliquez d'abord sur Afficher."
        Exit Sub
    End If
    If ThisWorkbook.Path = "" Then
        lblStatut.Caption = "Enregistrez d'abord le classeur."
        Exit Sub
    End If
    ReDim t(1 To lstResultats.ListCount + 1, 1 To 5)
    t(1, 1) = "Produit": t(1, 2) = "M+1": t(1, 3) = "Intervalle M+1": t(1, 4) = "M+2": t(1, 5) = "M+3"
    For i = 0 To lstResultats.ListCount - 1
        For j = 0 To 4
            t(i + 2, j + 1) = CStr(lstResultats.List(i, j) & "")
        Next j
    Next i
    chemin = ThisWorkbook.Path & "\prevision_" & cmbStore.Value & "_" & cmbCategorie.Value & "_" & _
             IIf(msMois = "", Format(Date, "yyyy-mm"), msMois) & ".csv"
    EcrireCsv chemin, t, ";"
    lblStatut.Caption = "Exporté : " & chemin
End Sub

Private Sub btnFermer_Click()
    Unload Me
End Sub
