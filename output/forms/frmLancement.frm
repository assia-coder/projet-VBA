VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmLancement
   Caption         =   "Centre de contrôle"
   ClientHeight    =   6300
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7800
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmLancement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==============================================================================
' frmLancement - Centre de contrôle (écran principal de l'application)
' Les contrôles sont créés à l'exécution (ModuleApp.AjouterCtl) : ce fichier .frm
' est du texte pur, importable sans fichier .frx.
'==============================================================================
Option Explicit

' --- Contrôles ---
Private lblTitre As MSForms.Label
Private lblDernierDate As MSForms.Label
Private lblDernier As MSForms.Label
Private lblModele As MSForms.Label
Private WithEvents btnEntrainementAll As MSForms.CommandButton
Private WithEvents btnPredire As MSForms.CommandButton
Private WithEvents btnPrevisionMensuelle As MSForms.CommandButton
Private WithEvents btnGestion As MSForms.CommandButton
Private WithEvents btnMetriques As MSForms.CommandButton
Private lblStatut As MSForms.Label


' -- Construction de l'écran ----------------------------------------------
Private Sub UserForm_Initialize()
    Const L As Single = 14, W As Single = 370
    Me.Caption = "Centre de contrôle"
    Me.Width = 400
    Me.Height = 340
    Me.BackColor = RGB(255, 255, 255)

    Set lblTitre = AjouterCtl(Me, "Label", "lblTitre", L, 12, W, 24, "Prévision de la demande — M5 Walmart")
    lblTitre.Font.Size = 14
    lblTitre.Font.Bold = True
    lblTitre.ForeColor = RGB(30, 86, 49)

    Set lblDernierDate = AjouterCtl(Me, "Label", "lblDernierDate", L, 44, W, 14)
    Set lblDernier = AjouterCtl(Me, "Label", "lblDernier", L, 60, W, 14)
    Set lblModele = AjouterCtl(Me, "Label", "lblModele", L, 76, W, 14)
    lblModele.Font.Bold = True

    Set btnEntrainementAll = AjouterCtl(Me, "CommandButton", "btnEntrainementAll", L, 102, W, 32, _
                                        "Lancer tout l'entraînement")
    btnEntrainementAll.BackColor = RGB(30, 120, 60)
    btnEntrainementAll.ForeColor = RGB(255, 255, 255)
    btnEntrainementAll.Font.Bold = True
    btnEntrainementAll.Font.Size = 10

    Set btnPredire = AjouterCtl(Me, "CommandButton", "btnPredire", L, 142, W, 26, _
                                "Prédire les 30 prochains jours")
    Set btnPrevisionMensuelle = AjouterCtl(Me, "CommandButton", "btnPrevisionMensuelle", L, 174, W, 26, _
                                           "Voir les prévisions mensuelles")
    Set btnGestion = AjouterCtl(Me, "CommandButton", "btnGestion", L, 206, W, 26, _
                                "Gérer les produits / ajouter des ventes")
    Set btnMetriques = AjouterCtl(Me, "CommandButton", "btnMetriques", L, 238, W, 26, "Voir les métriques")

    Set lblStatut = AjouterCtl(Me, "Label", "lblStatut", L, 276, W, 28)
    lblStatut.Font.Italic = True
    lblStatut.ForeColor = RGB(128, 128, 128)
    lblStatut.WordWrap = True

    ChargerInfos
End Sub

' Recharge les informations depuis PARAMS (B4 dernière date, B5 entraînement, B6 modèle)
Private Sub ChargerInfos()
    Dim d As String, e As String, m As String
    d = Param("B4"): e = Param("B5"): m = Param("B6")
    If Len(d) = 0 Then d = "Données non disponibles"
    If Len(e) = 0 Then e = "Données non disponibles"
    If Len(m) = 0 Then m = "Données non disponibles"
    lblDernierDate.Caption = "Dernière date du dataset : " & d
    lblDernier.Caption = "Dernier entraînement : " & e
    lblModele.Caption = "Meilleur modèle : " & m
    If Len(Param("B7")) > 0 Then lblModele.Caption = lblModele.Caption & "  (RMSE " & Param("B7") & ")"
End Sub

' Grise les boutons pendant un calcul Python
Private Sub ActiverBoutons(ByVal actif As Boolean)
    btnEntrainementAll.Enabled = actif
    btnPredire.Enabled = actif
    btnPrevisionMensuelle.Enabled = actif
    btnGestion.Enabled = actif
    btnMetriques.Enabled = actif
End Sub


' -- Actions --------------------------------------------------------------
Private Sub btnEntrainementAll_Click()
    If MsgBox("Lancer tout l'entraînement ?" & vbCrLf & vbCrLf & _
              "Preprocessing, XGBoost, LightGBM, Prophet, comparaison, prévision 30 jours." & vbCrLf & _
              "Durée : plusieurs minutes.", vbQuestion + vbYesNo, "Entraînement") = vbNo Then Exit Sub
    ActiverBoutons False
    lblStatut.Caption = "Entraînement en cours... (barre d'état d'Excel pour le détail)"
    DoEvents
    If ModuleApp.LancerEntrainement("all") Then
        lblStatut.Caption = "Entraînement terminé : " & Format(Now, "dd/mm/yyyy hh:nn") & "."
    Else
        lblStatut.Caption = "Entraînement interrompu (voir output\bridge_log.txt)."
    End If
    ChargerInfos                 ' recharge les labels
    ActiverBoutons True
End Sub

Private Sub btnPredire_Click()
    ActiverBoutons False
    lblStatut.Caption = "Prédiction en cours..."
    DoEvents
    If ModuleApp.LancerEntrainement("predict") Then
        lblStatut.Caption = "Prédiction terminée : feuille PREVISIONS_30J mise à jour."
    Else
        lblStatut.Caption = "Prédiction interrompue."
    End If
    ChargerInfos
    ActiverBoutons True
End Sub

Private Sub btnPrevisionMensuelle_Click()
    frmPrevisionMensuelle.Show
End Sub

Private Sub btnGestion_Click()
    frmGestionProduits.Show
    ChargerInfos                 ' un entraînement a pu être lancé depuis la gestion
End Sub

Private Sub btnMetriques_Click()
    frmMetriques.Show
End Sub
