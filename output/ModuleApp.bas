Attribute VB_Name = "ModuleApp"
'==============================================================================
' ModuleApp - Application de prévision de la demande (M5 Walmart)
'
' Interface 100 % UserForms : frmLancement (centre de contrôle), frmPrevisionMensuelle,
' frmGestionProduits, frmMetriques.
' Calculs Python lancés par excel_bridge.py (dossier parent du classeur), résultats
' relus depuis les CSV de output/.
'
' Import : Alt+F11 > clic droit sur le projet > Importer un fichier > ModuleApp.bas
' Démarrage : dans ThisWorkbook
'     Private Sub Workbook_Open()
'         frmLancement.Show
'     End Sub
'==============================================================================
Option Explicit

' --- Noms des feuilles (générées par generate_workbook.py) ---
Public Const F_DON As String = "DONNEES"
Public Const F_PREV As String = "PREVISION"
Public Const F_PREV30 As String = "PREVISIONS_30J"
Public Const F_METR As String = "METRIQUES"
Public Const F_WIN As String = "WINNER"
Public Const F_MENS As String = "PREVISION_MENSUELLE"
Public Const F_PROD As String = "PRODUITS"
Public Const F_NV As String = "NOUVELLES_VENTES"
Public Const F_PARAMS As String = "PARAMS"
Public Const F_TEST As String = "PREVISIONS_TEST"

' --- Lancement Python ---
Private Const DELAI_MAX As Long = 600          ' secondes (commandes unitaires)
Private Const PAS_POLLING As Long = 3          ' secondes entre deux lectures de status.txt
Private Const DELAI_DEMARRAGE As Long = 45     ' secondes sans status.txt = Python n'a pas démarré

' --- Feuille PREVISION (méthodes VBA classiques) ---
Private Const PREM_LIG As Long = 5
Private Const COL_DATE As Long = 20
Private Const COL_REEL As Long = 21
Private Const COL_XMAS As Long = 22
Private Const COL_NAIF As Long = 23
Private Const COL_MMS As Long = 24
Private Const COL_MMP As Long = 25
Private Const COL_LES As Long = 26
Private Const COL_TEND As Long = 27
Private Const LIG_METH As Long = 15      ' 1re ligne du tableau de comparaison (15 à 20)
Private Const LIG_FUTUR As Long = 24     ' 1re ligne du tableau J+1..J+7
Private Const MARGE_IC As Double = 0.12

Private mbSilencieux As Boolean


'==============================================================================
' LANCEMENT PYTHON
'==============================================================================
' Lance "python excel_bridge.py <commande>" sans console, puis surveille status.txt
' toutes les 3 secondes. Retourne True si Python a terminé avec SUCCESS.
' Utilisable comme une Sub : LancerEntrainement "all"
Public Function LancerEntrainement(ByVal commande As String, Optional ByVal delaiMax As Long = 0) As Boolean
    Dim projet As String, fStatut As String, cmd As String
    Dim statut As String, message As String, t0 As Double, ecoule As Double

    If delaiMax <= 0 Then
        delaiMax = DELAI_MAX
    End If
    If ThisWorkbook.Path = "" Then
        MsgBox "Enregistrez d'abord le classeur (.xlsm) dans le dossier output du projet.", vbExclamation
        Exit Function
    End If
    projet = DossierProjet()
    If Len(Dir(projet & "\excel_bridge.py")) = 0 Then
        MsgBox "excel_bridge.py introuvable dans :" & vbCrLf & projet & vbCrLf & _
               "Le classeur doit rester dans le dossier output du projet.", vbExclamation
        Exit Function
    End If
    fStatut = DossierSorties() & "\status.txt"

    ' Suppression de l'ancien statut : sinon le SUCCESS d'une exécution
    ' précédente serait lu immédiatement, avant la fin du calcul
    On Error Resume Next
    Kill fStatut
    On Error GoTo 0

    ' Lancement asynchrone, console masquée
    cmd = "cmd /c cd /d """ & projet & """ && " & CommandePython() & " excel_bridge.py " & commande
    On Error GoTo ErreurShell
    Shell cmd, vbHide
    On Error GoTo 0

    ' Polling de status.txt jusqu'à SUCCESS / ERROR / délai dépassé
    t0 = Timer
    Do
        Attendre PAS_POLLING
        ecoule = Timer - t0
        If ecoule < 0 Then ecoule = ecoule + 86400        ' passage de minuit
        LireStatut fStatut, statut, message
        Application.StatusBar = "Python '" & commande & "' en cours... " & Format(ecoule, "0") & " s" & _
                                IIf(Len(message) > 0, "  -  " & message, "")
        If statut = "SUCCESS" Or statut = "ERROR" Then Exit Do
        If statut = "" And ecoule > DELAI_DEMARRAGE Then Exit Do
    Loop While ecoule < delaiMax
    Application.StatusBar = False

    Select Case statut
        Case "SUCCESS"
            ' "predict" et "workbook" ne réentraînent pas : la date d'entraînement ne change pas
            ActualiserClasseur (commande <> "predict" And commande <> "workbook")
            LancerEntrainement = True
        Case "ERROR"
            MsgBox "Erreur Python (" & commande & ") :" & vbCrLf & message & vbCrLf & vbCrLf & _
                   "Détail : output\bridge_log.txt", vbExclamation
        Case ""
            MsgBox "Python n'a pas démarré (aucun status.txt après " & DELAI_DEMARRAGE & " s)." & vbCrLf & _
                   "Commande utilisée : " & CommandePython() & vbCrLf & _
                   "Vérifiez PARAMS!B8 ou lancez à la main : python excel_bridge.py " & commande, vbExclamation
        Case Else
            MsgBox "Délai dépassé (" & delaiMax & "s). Vérifier bridge_log.txt.", vbExclamation
    End Select
    Exit Function

ErreurShell:
    Application.StatusBar = False
    MsgBox "Impossible de lancer Python : " & Err.Description, vbCritical
End Function

' Raccourci : prévision J+1..J+30 avec le modèle gagnant (modèles déjà entraînés)
Public Sub LancerPrediction()
    LancerEntrainement "predict"
End Sub

' Attente non bloquante : Excel et les formulaires restent rafraîchis
Private Sub Attendre(ByVal secondes As Long)
    Dim i As Long
    For i = 1 To secondes
        Application.Wait Now + TimeSerial(0, 0, 1)
        DoEvents
    Next i
End Sub


'==============================================================================
' MISE A JOUR DU CLASSEUR DEPUIS output/
'==============================================================================
' Relit les CSV produits par Python et met à jour les feuilles + plages nommées
Public Sub ActualiserClasseur(Optional ByVal apresEntrainement As Boolean = True)
    Dim dossier As String
    dossier = DossierSorties()
    On Error GoTo Erreur
    Application.ScreenUpdating = False
    Application.StatusBar = "Mise à jour du classeur..."

    LireCsv dossier & "\series_long.csv", F_DON, "PlgDonnees", _
            "date,store_id,item_id,cat_id,sales,sell_price,snap,has_event"
    LireCsv dossier & "\forecast_next_30_days.csv", F_PREV30, "PlgPrevisions30j"
    LireCsv dossier & "\metrics_all.csv", F_METR, "PlgMetriques"
    LireCsv dossier & "\previsions_test.csv", F_TEST, "PlgPrevisionsTest"
    LireCsv dossier & "\prevision_mensuelle.csv", F_MENS, "PlgMensuel"
    LireCsv dossier & "\produits.csv", F_PROD, "PlgProduits"
    MajWinner
    MajParams apresEntrainement

    Application.ScreenUpdating = True
    Application.StatusBar = "Classeur mis à jour."
    Exit Sub
Erreur:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    MsgBox "Erreur pendant la mise à jour du classeur : " & Err.Description, vbExclamation
End Sub

' Feuille PARAMS : B4 dernière date du dataset, B5 dernier entraînement, B6/B7 gagnant
Private Sub MajParams(Optional ByVal apresEntrainement As Boolean = True)
    Dim ws As Worksheet, wd As Worksheet, fJson As String, derLig As Long, rmse As String
    Set ws = Feuille(F_PARAMS)
    fJson = DossierSorties() & "\winner.json"
    ws.Range("A4").Value = "Dernière date du dataset"
    ws.Range("A5").Value = "Date du dernier entraînement"
    ws.Range("A6").Value = "Modèle recommandé"
    ws.Range("A7").Value = "RMSE du modèle recommandé"
    If apresEntrainement Then ws.Range("B5").Value = Format(Now, "yyyy-mm-dd hh:nn")
    ws.Range("B6").Value = LireJson(fJson, "label")
    rmse = LireJson(fJson, "RMSE")
    If Len(rmse) > 0 Then ws.Range("B7").Value = Val(rmse) Else ws.Range("B7").Value = ""

    Set wd = Feuille(F_DON)
    derLig = wd.Cells(wd.Rows.Count, 1).End(xlUp).Row
    If derLig >= 2 Then
        ws.Range("B4").Value = Format(Application.WorksheetFunction.Max(wd.Range("A2:A" & derLig)), "yyyy-mm-dd")
    End If
End Sub

' Feuille WINNER depuis winner.json
Private Sub MajWinner()
    Dim ws As Worksheet, f As String, champs As Variant, cles As Variant, i As Long, v As String
    Set ws = Feuille(F_WIN)
    f = DossierSorties() & "\winner.json"
    champs = Array("Modèle gagnant", "Approche", "RMSE", "MAE", "WAPE", "Date du classement")
    cles = Array("label", "approach", "RMSE", "MAE", "WAPE", "updated_at")
    ws.Range("A1:B1").Value = Array("Champ", "Valeur")
    For i = 0 To UBound(champs)
        v = LireJson(f, CStr(cles(i)))
        ws.Cells(i + 2, 1).Value = champs(i)
        If i >= 2 And i <= 4 And Len(v) > 0 Then
            ws.Cells(i + 2, 2).Value = Val(v)
        ElseIf i = 0 And Len(v) = 0 Then
            ws.Cells(i + 2, 2).Value = "Données non disponibles"
        Else
            ws.Cells(i + 2, 2).Value = v
        End If
    Next i
    NommerPlage "PlgWinner", ws.Range("A1:B7")
End Sub

' Copie un CSV dans une feuille à partir de A1 et met à jour la plage nommée.
' colonnes (optionnel) : liste "a,b,c" des colonnes à garder, dans cet ordre.
' Fichier absent : la feuille est vidée et A1 indique "Données non disponibles".
Private Sub LireCsv(ByVal chemin As String, ByVal feuille As String, ByVal plage As String, _
                    Optional ByVal colonnes As String = "")
    Dim ws As Worksheet, wbCsv As Workbook, src As Variant, res As Variant
    Dim nL As Long, nC As Long, i As Long, j As Long, k As Long
    Dim noms() As String, idx() As Long

    Set ws = Feuille(feuille)
    ws.Cells.ClearContents
    If Len(Dir(chemin)) = 0 Then
        ws.Range("A1").Value = "Données non disponibles"
        NommerPlage plage, ws.Range("A1")
        Exit Sub
    End If

    ' Local:=False : séparateur "," et décimale "." quel que soit le paramétrage Windows
    Application.DisplayAlerts = False
    Set wbCsv = Workbooks.Open(Filename:=chemin, ReadOnly:=True, Local:=False)
    src = wbCsv.Worksheets(1).UsedRange.Value
    wbCsv.Close SaveChanges:=False
    Application.DisplayAlerts = True

    If Not IsArray(src) Then
        ws.Range("A1").Value = src
        NommerPlage plage, ws.Range("A1")
        Exit Sub
    End If
    nL = UBound(src, 1)

    If Len(colonnes) = 0 Then
        res = src
        nC = UBound(src, 2)
    Else
        ' Sélection des colonnes par nom d'en-tête
        noms = Split(colonnes, ",")
        nC = UBound(noms) + 1
        ReDim idx(1 To nC)
        For j = 1 To nC
            For k = 1 To UBound(src, 2)
                If LCase$(Trim$(CStr(src(1, k)))) = LCase$(Trim$(noms(j - 1))) Then idx(j) = k
            Next k
        Next j
        ReDim res(1 To nL, 1 To nC)
        For i = 1 To nL
            For j = 1 To nC
                If idx(j) > 0 Then res(i, j) = src(i, idx(j))
            Next j
        Next i
        For j = 1 To nC
            res(1, j) = noms(j - 1)
        Next j
    End If

    ws.Range("A1").Resize(nL, nC).Value = res
    NommerPlage plage, ws.Range("A1").Resize(nL, nC)
End Sub

' Extrait la valeur de "cle" dans un JSON plat (recherche simple par texte).
' Retourne "" si le fichier est absent ou la clé introuvable.
Private Function LireJson(ByVal chemin As String, ByVal cle As String) As String
    Dim txt As String, p As Long, q As Long, v As String, c As String
    txt = LireTexteUtf8(chemin)
    If Len(txt) = 0 Then Exit Function
    p = InStr(1, txt, """" & cle & """", vbBinaryCompare)
    If p = 0 Then Exit Function
    p = InStr(p + Len(cle) + 2, txt, ":")
    If p = 0 Then Exit Function
    v = Trim$(Replace(Replace(Replace(Mid$(txt, p + 1), vbCr, " "), vbLf, " "), vbTab, " "))
    If Left$(v, 1) = """" Then
        q = InStr(2, v, """")
        If q > 1 Then LireJson = Mid$(v, 2, q - 2)
    Else
        For q = 1 To Len(v)
            c = Mid$(v, q, 1)
            If c = "," Or c = "}" Or c = " " Then Exit For
        Next q
        LireJson = Left$(v, q - 1)
    End If
End Function

' Lit un fichier texte UTF-8 (ADODB.Stream) ; repli sur une lecture ANSI
Private Function LireTexteUtf8(ByVal chemin As String) As String
    Dim st As Object, f As Integer, ligne As String
    If Len(Dir(chemin)) = 0 Then Exit Function
    On Error GoTo Repli
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2
    st.Charset = "utf-8"
    st.Open
    st.LoadFromFile chemin
    LireTexteUtf8 = st.ReadText
    st.Close
    Exit Function
Repli:
    On Error GoTo 0
    f = FreeFile
    Open chemin For Input As #f
    Do While Not EOF(f)
        Line Input #f, ligne
        LireTexteUtf8 = LireTexteUtf8 & ligne & vbLf
    Loop
    Close #f
End Function

' Lit status.txt : ligne 1 = statut, ligne 2 = message, ligne 3 = horodatage
Private Sub LireStatut(ByVal chemin As String, ByRef statut As String, ByRef message As String, _
                       Optional ByRef quand As String)
    Dim f As Integer
    statut = "": message = "": quand = ""
    If Len(Dir(chemin)) = 0 Then Exit Sub
    f = FreeFile
    On Error GoTo Fin                       ' fichier éventuellement en cours d'écriture
    Open chemin For Input As #f
    If Not EOF(f) Then Line Input #f, statut
    If Not EOF(f) Then Line Input #f, message
    If Not EOF(f) Then Line Input #f, quand
    statut = Trim$(statut)
Fin:
    Close #f
End Sub


'==============================================================================
' CALCUL DES METRIQUES EN VBA (validation des métriques Python)
'==============================================================================
' Recalcule MAE, MAPE, RMSE, WAPE de chaque modèle sur la période de test et les
' compare aux valeurs Python (niveau global de PlgMetriques). Résultat écrit dans
' METRIQUES, section "Calcul VBA" sous le tableau Python.
' Retourne True si tous les écarts sont <= 0,01 %.
Public Function CalculerMetriquesVBA(Optional ByVal silencieux As Boolean = False) As Boolean
    Dim ws As Worksheet, v As Variant, m As Variant, n As Long, i As Long, k As Long
    Dim cT As Long, cP As Long, cM As Long, cA As Long
    Dim cles As Object, cle As Variant, yT() As Double, yP() As Double, nb As Long
    Dim mv(1 To 4) As Double, py(1 To 4) As Variant, dec As Variant
    Dim ecart As Double, ecartMax As Double, tousOk As Boolean, ok As Boolean
    Dim wm As Worksheet, lig As Long, finTable As Long, rapport As String, src As String

    ' 1. Source : PREVISIONS_30J si elle contient le réel (y_true), sinon la période de test
    v = LirePlage("PlgPrevisions30j")
    If IsArray(v) Then
        If ColIndex(v, "y_true") = 0 Then v = Empty
    End If
    src = F_PREV30
    If Not IsArray(v) Then
        v = LirePlage("PlgPrevisionsTest")
        src = F_TEST
    End If
    If Not IsArray(v) Then
        If Not silencieux Then MsgBox "Données non disponibles : aucune prévision avec valeurs réelles " & _
                                      "(feuille " & F_TEST & ").", vbExclamation
        Exit Function
    End If
    cT = ColIndex(v, "y_true"): cP = ColIndex(v, "y_pred")
    cM = ColIndex(v, "model"): cA = ColIndex(v, "approach")
    n = UBound(v, 1)

    ' 2. Liste des modèles (ordre d'apparition)
    Set cles = CreateObject("Scripting.Dictionary")
    For i = 2 To n
        cle = CStr(v(i, cM)) & "|" & CStr(v(i, cA))
        cles(cle) = cles(cle) + 1
    Next i

    ' 3. Effacement de l'ancienne section VBA puis en-têtes
    Set wm = Feuille(F_METR)
    m = LirePlage("PlgMetriques")
    If IsArray(m) Then finTable = UBound(m, 1) Else finTable = 1
    With wm.Range(wm.Cells(finTable + 1, 1), wm.Cells(finTable + 40, 12))
        .ClearContents
        .Font.Bold = False
    End With
    lig = finTable + 3
    wm.Cells(lig, 1).Value = "Calcul VBA (période de test, source : " & src & ") - " & Format(Now, "yyyy-mm-dd hh:nn")
    wm.Cells(lig, 1).Font.Bold = True
    lig = lig + 1
    wm.Range(wm.Cells(lig, 1), wm.Cells(lig, 11)).Value = Array("Modele", "Approche", "N", "MAE VBA", _
        "MAPE VBA", "RMSE VBA", "WAPE VBA", "RMSE Python", "WAPE Python", "Ecart max %", "Statut")
    wm.Range(wm.Cells(lig, 1), wm.Cells(lig, 11)).Font.Bold = True

    ' Décimales des valeurs Python (utils.compute_metrics) : MAE 4, MAPE 2, RMSE 4, WAPE 2
    dec = Array(0, 4, 2, 4, 2)
    tousOk = True
    For Each cle In cles.Keys
        ' Séries réel / prévu du modèle
        nb = cles(cle)
        ReDim yT(1 To nb): ReDim yP(1 To nb)
        k = 0
        For i = 2 To n
            If CStr(v(i, cM)) & "|" & CStr(v(i, cA)) = cle Then
                k = k + 1
                yT(k) = Nombre(v(i, cT))
                yP(k) = Nombre(v(i, cP))
            End If
        Next i
        mv(1) = CalcMAE(yT, yP)
        mv(2) = CalcMAPE(yT, yP)
        mv(3) = CalcRMSE(yT, yP)
        mv(4) = CalcWAPE(yT, yP)

        ' Valeurs Python correspondantes et écart relatif (VBA arrondi comme Python)
        MetriquesPython m, Split(cle, "|")(0), Split(cle, "|")(1), py
        ecartMax = 0
        ok = True
        For k = 1 To 4
            If EstNombre(py(k)) Then
                If py(k) <> 0 Then
                    ecart = Abs(Round(mv(k), dec(k)) - py(k)) / Abs(py(k)) * 100
                Else
                    ecart = Abs(Round(mv(k), dec(k)))
                End If
                If ecart > ecartMax Then ecartMax = ecart
            Else
                ok = False
            End If
        Next k
        If ecartMax > 0.01 Then ok = False
        If Not ok Then tousOk = False

        lig = lig + 1
        wm.Range(wm.Cells(lig, 1), wm.Cells(lig, 11)).Value = Array( _
            Split(cle, "|")(0), Split(cle, "|")(1), nb, Round(mv(1), 4), Round(mv(2), 2), _
            Round(mv(3), 4), Round(mv(4), 2), py(3), py(4), Round(ecartMax, 5), _
            IIf(ok, "OK", IIf(EstNombre(py(3)), "ECART", "Python absent")))
        wm.Cells(lig, 11).Font.Color = IIf(ok, RGB(0, 97, 0), RGB(192, 0, 0))
        rapport = rapport & Split(cle, "|")(0) & " (" & Split(cle, "|")(1) & ") : RMSE VBA = " & _
                  Format(mv(3), "0.0000") & "  /  Python = " & Fmt(py(3)) & "  ->  " & _
                  IIf(ok, "OK", "ECART") & vbCrLf
    Next cle
    wm.Cells(lig + 1, 1).Value = "Tolérance : écart relatif <= 0,01 % (valeurs VBA arrondies comme Python)."

    CalculerMetriquesVBA = tousOk
    If Not silencieux Then
        MsgBox "Métriques recalculées en VBA (" & src & ") :" & vbCrLf & vbCrLf & rapport & vbCrLf & _
               IIf(tousOk, "Tous les écarts sont <= 0,01 %.", "Au moins un écart > 0,01 % : voir la feuille METRIQUES."), _
               IIf(tousOk, vbInformation, vbExclamation), "Validation des métriques"
    End If
End Function

' MAE, MAPE, RMSE, WAPE "Python" (niveau global) d'un modèle, lus dans PlgMetriques
Private Sub MetriquesPython(ByVal m As Variant, ByVal modele As String, ByVal approche As String, _
                            ByRef py() As Variant)
    Dim i As Long, cMo As Long, cAp As Long, cLe As Long, noms As Variant, k As Long
    For k = 1 To 4
        py(k) = Empty
    Next k
    If Not IsArray(m) Then Exit Sub
    cMo = ColIndex(m, "Modele"): cAp = ColIndex(m, "Approche"): cLe = ColIndex(m, "Level")
    noms = Array("", "MAE", "MAPE", "RMSE", "WAPE")
    For i = 2 To UBound(m, 1)
        If CStr(m(i, cMo)) = modele And CStr(m(i, cAp)) = approche And CStr(m(i, cLe)) = "global" Then
            For k = 1 To 4
                py(k) = m(i, ColIndex(m, CStr(noms(k))))
            Next k
            Exit Sub
        End If
    Next i
End Sub

' MAE = moyenne des |réel - prévu|
Private Function CalcMAE(ByRef yT() As Double, ByRef yP() As Double) As Double
    Dim i As Long, n As Long, s As Double
    n = UBound(yT)
    For i = 1 To n
        s = s + Abs(yT(i) - yP(i))
    Next i
    CalcMAE = s / n
End Function

' MAPE (%) = moyenne des |erreur| / réel sur les jours où réel > 0
Private Function CalcMAPE(ByRef yT() As Double, ByRef yP() As Double) As Double
    Dim i As Long, n As Long, s As Double, cnt As Long
    n = UBound(yT)
    For i = 1 To n
        If yT(i) > 0 Then
            s = s + Abs((yT(i) - yP(i)) / yT(i))
            cnt = cnt + 1
        End If
    Next i
    If cnt > 0 Then CalcMAPE = s / cnt * 100 Else CalcMAPE = 0
End Function

' RMSE = racine de la moyenne des erreurs au carré
Private Function CalcRMSE(ByRef yT() As Double, ByRef yP() As Double) As Double
    Dim i As Long, n As Long, s As Double
    n = UBound(yT)
    For i = 1 To n
        s = s + (yT(i) - yP(i)) ^ 2
    Next i
    CalcRMSE = Sqr(s / n)
End Function

' WAPE (%) = somme |erreur| / somme réel
Private Function CalcWAPE(ByRef yT() As Double, ByRef yP() As Double) As Double
    Dim i As Long, n As Long, sErr As Double, sTrue As Double
    n = UBound(yT)
    For i = 1 To n
        sErr = sErr + Abs(yT(i) - yP(i))
        sTrue = sTrue + yT(i)
    Next i
    If sTrue > 0 Then CalcWAPE = sErr / sTrue * 100 Else CalcWAPE = 0
End Function


'==============================================================================
' OUTILS POUR LES USERFORMS
'==============================================================================
' Crée un contrôle MSForms à l'exécution (les .frm restent du texte pur, sans .frx)
' typ : Label, CommandButton, ComboBox, ListBox, TextBox, CheckBox, Frame
Public Function AjouterCtl(ByVal conteneur As Object, ByVal typ As String, ByVal nom As String, _
        ByVal gauche As Single, ByVal haut As Single, ByVal largeur As Single, ByVal hauteur As Single, _
        Optional ByVal texte As String = "") As Object
    Dim c As Object
    Set c = conteneur.Controls.Add("Forms." & typ & ".1", nom, True)
    c.Left = gauche: c.Top = haut: c.Width = largeur: c.Height = hauteur
    On Error Resume Next                    ' tous les contrôles n'ont pas Caption / BackStyle
    c.Font.Name = "Segoe UI"
    c.Font.Size = 9
    If Len(texte) > 0 Then c.Caption = texte
    If typ = "Label" Or typ = "CheckBox" Then c.BackStyle = 0      ' transparent
    On Error GoTo 0
    Set AjouterCtl = c
End Function

' Valeurs d'une plage nommée (en-têtes compris) ; Empty si absente ou sans données
Public Function LirePlage(ByVal nom As String) As Variant
    Dim r As Range
    On Error Resume Next
    Set r = ThisWorkbook.Names(nom).RefersToRange
    On Error GoTo 0
    If r Is Nothing Then Exit Function
    If r.Rows.Count < 2 Or r.Columns.Count < 2 Then Exit Function
    LirePlage = r.Value
End Function

' Numéro de colonne d'un en-tête dans un tableau (ligne 1) ; 0 si absent
Public Function ColIndex(ByVal v As Variant, ByVal nom As String) As Long
    Dim j As Long
    If Not IsArray(v) Then Exit Function
    For j = 1 To UBound(v, 2)
        If LCase$(Trim$(CStr(v(1, j)))) = LCase$(nom) Then
            ColIndex = j
            Exit Function
        End If
    Next j
End Function

' Liste "a,b,c" d'une cellule de PARAMS (B2 stores, B3 catégories)
Public Function ListeParam(ByVal adresse As String, ByVal defaut As String) As Variant
    Dim s As String
    On Error Resume Next
    s = Trim$(CStr(ThisWorkbook.Worksheets(F_PARAMS).Range(adresse).Value))
    On Error GoTo 0
    If Len(s) = 0 Then s = defaut
    ListeParam = Split(Replace(s, " ", ""), ",")
End Function

' Valeur d'une cellule de PARAMS (texte vide si la feuille manque)
Public Function Param(ByVal adresse As String) As String
    On Error Resume Next
    Param = CStr(ThisWorkbook.Worksheets(F_PARAMS).Range(adresse).Text)
End Function

' Ajoute une ligne à la fin d'une plage nommée et étend la plage
Public Sub AjouterLigne(ByVal nom As String, ByVal valeurs As Variant)
    Dim r As Range, n As Long
    Set r = ThisWorkbook.Names(nom).RefersToRange
    n = UBound(valeurs) - LBound(valeurs) + 1
    r.Worksheet.Cells(r.Row + r.Rows.Count, r.Column).Resize(1, n).Value = valeurs
    NommerPlage nom, r.Resize(r.Rows.Count + 1, Application.WorksheetFunction.Max(n, r.Columns.Count))
End Sub

' Nombre de lignes de données (hors en-têtes) d'une plage nommée
Public Function NbLignesPlage(ByVal nom As String) As Long
    Dim v As Variant
    v = LirePlage(nom)
    If IsArray(v) Then NbLignesPlage = UBound(v, 1) - 1
End Function

' Lecture d'un nombre saisi ("1,25" ou "1.25") ; False si invalide
Public Function LireNombre(ByVal txt As String, ByRef valeur As Double) As Boolean
    Dim re As Object
    txt = Replace(Trim$(txt), ",", ".")
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "^\d+(\.\d+)?$"
    If Not re.Test(txt) Then Exit Function
    valeur = Val(txt)
    LireNombre = True
End Function

' Lecture d'une date JJ/MM/AAAA stricte ; False si invalide
Public Function LireDateJJMMAAAA(ByVal txt As String, ByRef d As Date) As Boolean
    Dim re As Object, p As Variant
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "^\d{2}/\d{2}/\d{4}$"
    txt = Trim$(txt)
    If Not re.Test(txt) Then Exit Function
    p = Split(txt, "/")
    If CLng(p(1)) < 1 Or CLng(p(1)) > 12 Or CLng(p(0)) < 1 Or CLng(p(0)) > 31 Then Exit Function
    d = DateSerial(CLng(p(2)), CLng(p(1)), CLng(p(0)))
    LireDateJJMMAAAA = (Day(d) = CLng(p(0)))      ' rejette 31/02/2016 etc.
End Function

' Catégorie d'un identifiant produit (FOODS_3_090 -> FOODS)
Public Function CategorieDe(ByVal item As String) As String
    CategorieDe = Split(item, "_")(0)
End Function


'==============================================================================
' NOUVELLES VENTES ET PRODUITS (frmGestionProduits)
'==============================================================================
' Écrit PlgNouvellesVentes dans output/nouvelles_ventes.csv (lu par preprocess.py)
Public Function ExporterNouvellesVentes() As Long
    Dim v As Variant
    v = LirePlage("PlgNouvellesVentes")
    If Not IsArray(v) Then Exit Function
    EcrireCsv DossierSorties() & "\nouvelles_ventes.csv", v, ","
    ExporterNouvellesVentes = UBound(v, 1) - 1
End Function

' Vide NOUVELLES_VENTES (en-têtes conservés) et supprime nouvelles_ventes.csv
Public Sub ViderNouvellesVentes()
    Dim r As Range
    Set r = ThisWorkbook.Names("PlgNouvellesVentes").RefersToRange
    If r.Rows.Count > 1 Then r.Offset(1, 0).Resize(r.Rows.Count - 1).ClearContents
    NommerPlage "PlgNouvellesVentes", r.Rows(1)
    On Error Resume Next
    Kill DossierSorties() & "\nouvelles_ventes.csv"
    On Error GoTo 0
End Sub

' Fusionne PlgNouvellesVentes dans PlgDonnees (ajout, ou remplacement si même date + store + produit)
Public Sub FusionnerNouvellesVentesDonnees()
    Dim nv As Variant, rd As Range, dn As Variant, dict As Object
    Dim i As Long, j As Long, cle As String, lig As Long, ws As Worksheet, nAjout As Long
    nv = LirePlage("PlgNouvellesVentes")
    If Not IsArray(nv) Then Exit Sub
    Set rd = ThisWorkbook.Names("PlgDonnees").RefersToRange
    Set ws = rd.Worksheet
    dn = rd.Columns("A:C").Value
    Set dict = CreateObject("Scripting.Dictionary")
    For i = 2 To UBound(dn, 1)
        If IsDate(dn(i, 1)) Then dict(CLng(CDate(dn(i, 1))) & "|" & dn(i, 2) & "|" & dn(i, 3)) = i
    Next i
    lig = rd.Row + rd.Rows.Count
    For i = 2 To UBound(nv, 1)
        cle = CLng(CDate(nv(i, 1))) & "|" & nv(i, 2) & "|" & nv(i, 3)
        If dict.Exists(cle) Then
            For j = 1 To 8
                rd.Cells(dict(cle), j).Value = nv(i, j)
            Next j
        Else
            For j = 1 To 8
                ws.Cells(lig + nAjout, rd.Column + j - 1).Value = nv(i, j)
            Next j
            nAjout = nAjout + 1
        End If
    Next i
    NommerPlage "PlgDonnees", rd.Resize(rd.Rows.Count + nAjout, 8)
End Sub

' Écrit PlgProduits dans output/produits.csv (relu par generate_workbook.py)
Public Sub ExporterProduits()
    Dim v As Variant
    v = LirePlage("PlgProduits")
    If IsArray(v) Then EcrireCsv DossierSorties() & "\produits.csv", v, ","
End Sub

' Écrit un tableau 2D en CSV (dates aaaa-mm-jj, décimale ".")
Public Sub EcrireCsv(ByVal chemin As String, ByVal v As Variant, ByVal sep As String)
    Dim f As Integer, i As Long, j As Long, ligne As String
    f = FreeFile
    Open chemin For Output As #f
    For i = 1 To UBound(v, 1)
        ligne = ""
        For j = 1 To UBound(v, 2)
            If j > 1 Then ligne = ligne & sep
            ligne = ligne & CsvVal(v(i, j), sep)
        Next j
        Print #f, ligne
    Next i
    Close #f
End Sub

Private Function CsvVal(ByVal x As Variant, ByVal sep As String) As String
    Select Case VarType(x)
        Case vbDate
            CsvVal = Format(x, "yyyy-mm-dd")
        Case vbBoolean
            CsvVal = IIf(x, "True", "False")
        Case vbDouble, vbSingle, vbLong, vbInteger, vbCurrency, vbDecimal, vbByte
            CsvVal = Trim$(Str(x))                 ' Str : toujours le point décimal
        Case vbEmpty, vbNull
            CsvVal = ""
        Case Else
            CsvVal = CStr(x)
            If InStr(CsvVal, sep) > 0 Or InStr(CsvVal, """") > 0 Then
                CsvVal = """" & Replace(CsvVal, """", """""") & """"
            End If
    End Select
End Function


'==============================================================================
' PREVISION PASSIVE VBA (5 METHODES CLASSIQUES, feuille PREVISION)
'==============================================================================
Public Sub LancerPrevisions()
    Dim ws As Worksheet, nMMS As Long, nMMP As Long, alpha As Double
    Dim saison As Long, nTest As Long, store As String, produit As String
    Dim dates() As Date, ventes() As Double, n As Long, i As Long, nMin As Long
    Dim sortie() As Variant, ligFin As Long, ligTest As Long, meilleur As String

    Set ws = ThisWorkbook.Worksheets(F_PREV)
    If Not LireParametres(ws, nMMS, nMMP, alpha, saison, nTest, store, produit) Then Exit Sub

    ' 1. Extraction de la série store / produit depuis DONNEES
    n = ExtraireSerie(store, produit, dates, ventes)
    If n = 0 Then
        MsgBox "Aucune donnée trouvée pour " & store & " / " & produit & " dans DONNEES.", vbExclamation
        Exit Sub
    End If
    nMin = nTest + Application.WorksheetFunction.Max(nMMS, nMMP, saison) + 2
    If n < nMin Then
        MsgBox "Série trop courte (" & n & " jours) : il en faut au moins " & nMin & ".", vbExclamation
        Exit Sub
    End If

    On Error GoTo Erreur
    Application.ScreenUpdating = False

    ' 2. Ecriture Date / Réel / is_xmas dans la zone de calcul
    ws.Range(ws.Cells(PREM_LIG, COL_DATE), ws.Cells(ws.Rows.Count, COL_TEND)).ClearContents
    ReDim sortie(1 To n, 1 To 3)
    For i = 1 To n
        sortie(i, 1) = dates(i)
        sortie(i, 2) = ventes(i)
        If Month(dates(i)) = 12 And Day(dates(i)) = 25 Then sortie(i, 3) = 1 Else sortie(i, 3) = 0
    Next i
    ligFin = PREM_LIG + n - 1
    ligTest = ligFin - nTest + 1
    ws.Range(ws.Cells(PREM_LIG, COL_DATE), ws.Cells(ligFin, COL_XMAS)).Value = sortie
    ws.Range(ws.Cells(PREM_LIG, COL_DATE), ws.Cells(ligFin, COL_DATE)).NumberFormat = "dd/mm/yyyy"
    ws.Range(ws.Cells(PREM_LIG, COL_REEL), ws.Cells(ligFin, COL_TEND)).NumberFormat = "0.00"

    ' 3. Les 5 méthodes, les erreurs, les prévisions futures
    NaifSaisonnier ws, ligFin
    MoyenneMobileSimple ws, ligFin
    MoyenneMobilePonderee ws, ligFin
    LissageExponentiel ws, ligFin
    ProjectionTendance ws, ligTest, ligFin
    CalculerErreurs ws, ligTest, ligFin
    RemplirPrevisionsFutures ws, ligFin

    Application.ScreenUpdating = True
    For i = LIG_METH To LIG_METH + 5
        If ws.Cells(i, 6).Value = "MEILLEUR" Then meilleur = ws.Cells(i, 1).Value & " (RMSE = " & Fmt(ws.Cells(i, 5).Value) & ")"
    Next i
    Info "Prévisions calculées pour " & store & " / " & produit & vbCrLf & _
         n & " jours, dont " & nTest & " jours de test." & vbCrLf & vbCrLf & _
         "Meilleure méthode : " & meilleur
    Exit Sub
Erreur:
    Application.ScreenUpdating = True
    MsgBox "Erreur dans LancerPrevisions : " & Err.Description, vbCritical
End Sub

' Lit et valide les paramètres B5:B11 de la feuille PREVISION
Private Function LireParametres(ByVal ws As Worksheet, ByRef nMMS As Long, ByRef nMMP As Long, _
        ByRef alpha As Double, ByRef saison As Long, ByRef nTest As Long, _
        ByRef store As String, ByRef produit As String) As Boolean
    Dim msg As String
    If Not EstEntierPositif(ws.Range("B5").Value) Then msg = msg & "- Fenêtre MMS (B5) : entier >= 1" & vbCrLf
    If Not EstEntierPositif(ws.Range("B6").Value) Then msg = msg & "- Fenêtre MMP (B6) : entier >= 1" & vbCrLf
    If Not EstNombre(ws.Range("B7").Value) Then
        msg = msg & "- Alpha (B7) : nombre entre 0 et 1" & vbCrLf
    ElseIf ws.Range("B7").Value <= 0 Or ws.Range("B7").Value >= 1 Then
        msg = msg & "- Alpha (B7) : nombre entre 0 et 1" & vbCrLf
    End If
    If Not EstEntierPositif(ws.Range("B8").Value) Then msg = msg & "- Saisonnalité (B8) : entier >= 1" & vbCrLf
    If Not EstEntierPositif(ws.Range("B9").Value) Then
        msg = msg & "- Jours de test (B9) : entier >= 7" & vbCrLf
    ElseIf ws.Range("B9").Value < 7 Then
        msg = msg & "- Jours de test (B9) : entier >= 7" & vbCrLf
    End If
    store = Trim$(CStr(ws.Range("B10").Value))
    produit = Trim$(CStr(ws.Range("B11").Value))
    If Not StoreValide(store) Then msg = msg & "- Store (B10) : CA_1, TX_1 ou WI_1" & vbCrLf
    If produit = "" Then msg = msg & "- Produit (B11) : à choisir dans la liste" & vbCrLf
    If msg <> "" Then
        MsgBox "Paramètres invalides :" & vbCrLf & vbCrLf & msg, vbExclamation, "PREVISION"
        Exit Function
    End If
    nMMS = CLng(ws.Range("B5").Value)
    nMMP = CLng(ws.Range("B6").Value)
    alpha = CDbl(ws.Range("B7").Value)
    saison = CLng(ws.Range("B8").Value)
    nTest = CLng(ws.Range("B9").Value)
    LireParametres = True
End Function

' Extrait la série journalière (triée par date, ventes sommées par jour) d'un store / produit
Private Function ExtraireSerie(ByVal store As String, ByVal produit As String, _
        ByRef dates() As Date, ByRef ventes() As Double) As Long
    Dim wd As Worksheet, derLig As Long, donnees As Variant, i As Long, j As Long
    Dim dict As Object, cles As Variant, cle As Long, tmp As Long, n As Long
    Dim cleTri() As Long

    Set wd = ThisWorkbook.Worksheets(F_DON)
    derLig = wd.Cells(wd.Rows.Count, 1).End(xlUp).Row
    If derLig < 2 Then Exit Function
    donnees = wd.Range("A2:E" & derLig).Value
    Set dict = CreateObject("Scripting.Dictionary")
    For i = 1 To UBound(donnees, 1)
        If CStr(donnees(i, 2)) = store Then
            If CStr(donnees(i, 3)) = produit Then
                If IsDate(donnees(i, 1)) Then
                    cle = CLng(CDate(donnees(i, 1)))
                    dict(cle) = dict(cle) + Nombre(donnees(i, 5))
                End If
            End If
        End If
    Next i
    n = dict.Count
    If n = 0 Then Exit Function

    ' Tri des dates (tri par insertion : la série est déjà quasi triée)
    cles = dict.Keys
    ReDim cleTri(1 To n)
    For i = 1 To n
        cleTri(i) = cles(i - 1)
    Next i
    For i = 2 To n
        tmp = cleTri(i)
        j = i - 1
        Do While j >= 1
            If cleTri(j) <= tmp Then Exit Do
            cleTri(j + 1) = cleTri(j)
            j = j - 1
        Loop
        cleTri(j + 1) = tmp
    Next i
    ReDim dates(1 To n)
    ReDim ventes(1 To n)
    For i = 1 To n
        dates(i) = CDate(cleTri(i))
        ventes(i) = dict(cleTri(i))
    Next i
    ExtraireSerie = n
End Function

' Naïf saisonnier : F(t) = y(t - s)
Private Sub NaifSaisonnier(ByVal ws As Worksheet, ByVal ligFin As Long)
    Dim y() As Double, res() As Variant, n As Long, r As Long, s As Long
    s = CLng(ws.Range("B8").Value)
    y = LireColonne(ws, COL_REEL, ligFin)
    n = UBound(y)
    ReDim res(1 To n, 1 To 1)
    For r = s + 1 To n
        res(r, 1) = y(r - s)
    Next r
    EcrireColonne ws, COL_NAIF, res
End Sub

' Moyenne mobile simple sur k jours (somme glissante)
Private Sub MoyenneMobileSimple(ByVal ws As Worksheet, ByVal ligFin As Long)
    Dim y() As Double, res() As Variant, n As Long, r As Long, k As Long, somme As Double
    k = CLng(ws.Range("B5").Value)
    y = LireColonne(ws, COL_REEL, ligFin)
    n = UBound(y)
    ReDim res(1 To n, 1 To 1)
    For r = 1 To n
        If r > k Then res(r, 1) = somme / k
        somme = somme + y(r)
        If r > k Then somme = somme - y(r - k)
    Next r
    EcrireColonne ws, COL_MMS, res
End Sub

' Moyenne mobile pondérée : poids 1..k, le jour le plus récent pèse k
Private Sub MoyenneMobilePonderee(ByVal ws As Worksheet, ByVal ligFin As Long)
    Dim y() As Double, res() As Variant, n As Long, r As Long, j As Long, k As Long
    Dim somme As Double, poids As Double
    k = CLng(ws.Range("B6").Value)
    poids = k * (k + 1) / 2
    y = LireColonne(ws, COL_REEL, ligFin)
    n = UBound(y)
    ReDim res(1 To n, 1 To 1)
    For r = k + 1 To n
        somme = 0
        For j = 1 To k
            somme = somme + j * y(r - k - 1 + j)
        Next j
        res(r, 1) = somme / poids
    Next r
    EcrireColonne ws, COL_MMP, res
End Sub

' Lissage exponentiel simple : F(2) = y(1) ; F(t) = alpha*y(t-1) + (1-alpha)*F(t-1)
Private Sub LissageExponentiel(ByVal ws As Worksheet, ByVal ligFin As Long)
    Dim y() As Double, res() As Variant, n As Long, r As Long, alpha As Double, f As Double
    alpha = CDbl(ws.Range("B7").Value)
    y = LireColonne(ws, COL_REEL, ligFin)
    n = UBound(y)
    ReDim res(1 To n, 1 To 1)
    If n >= 2 Then
        f = y(1)
        res(2, 1) = f
        For r = 3 To n
            f = alpha * y(r - 1) + (1 - alpha) * f
            res(r, 1) = f
        Next r
    End If
    EcrireColonne ws, COL_LES, res
End Sub

' Projection de tendance : régression linéaire sur l'apprentissage, projetée sur le test
Private Sub ProjectionTendance(ByVal ws As Worksheet, ByVal ligTest As Long, ByVal ligFin As Long)
    Dim y() As Double, res() As Variant, n As Long, r As Long, m As Long, a As Double, b As Double
    y = LireColonne(ws, COL_REEL, ligFin)
    n = UBound(y)
    m = ligTest - PREM_LIG
    CoefTendance y, m, a, b
    ReDim res(1 To n, 1 To 1)
    For r = m + 1 To n
        res(r, 1) = a + b * r
    Next r
    EcrireColonne ws, COL_TEND, res
End Sub

' Moindres carrés sur les m premiers points (x = 1..m)
Private Sub CoefTendance(ByRef y() As Double, ByVal m As Long, ByRef a As Double, ByRef b As Double)
    Dim i As Long, sx As Double, sy As Double, sxx As Double, sxy As Double, den As Double
    For i = 1 To m
        sx = sx + i
        sy = sy + y(i)
        sxx = sxx + CDbl(i) * i
        sxy = sxy + i * y(i)
    Next i
    den = m * sxx - sx * sx
    If den = 0 Then b = 0 Else b = (m * sxy - sx * sy) / den
    a = (sy - b * sx) / m
End Sub

' MAE / MAPE / RMSE sur le test (MAPE hors jours à 0 vente et hors Noël is_xmas=1)
Private Sub CalculerErreurs(ByVal ws As Worksheet, ByVal ligTest As Long, ByVal ligFin As Long)
    Dim y() As Double, xm As Variant, f As Variant, j As Long, r As Long, lig As Long
    Dim nb As Long, nm As Long, sae As Double, sse As Double, sape As Double, e As Double
    Dim meilleur As Long, rmseMin As Double, v As Variant

    y = LireColonne(ws, COL_REEL, ligFin)
    xm = ws.Range(ws.Cells(PREM_LIG, COL_XMAS), ws.Cells(ligFin, COL_XMAS)).Value
    For j = 0 To 4
        f = ws.Range(ws.Cells(PREM_LIG, COL_NAIF + j), ws.Cells(ligFin, COL_NAIF + j)).Value
        nb = 0: nm = 0: sae = 0: sse = 0: sape = 0
        For r = ligTest - PREM_LIG + 1 To ligFin - PREM_LIG + 1
            If EstNombre(f(r, 1)) Then
                e = y(r) - CDbl(f(r, 1))
                nb = nb + 1
                sae = sae + Abs(e)
                sse = sse + e * e
                If y(r) > 0 And Nombre(xm(r, 1)) = 0 Then
                    nm = nm + 1
                    sape = sape + Abs(e) / y(r)
                End If
            End If
        Next r
        lig = LIG_METH + j
        If nb > 0 Then
            ws.Cells(lig, 3).Value = Round(sae / nb, 4)
            ws.Cells(lig, 5).Value = Round(Sqr(sse / nb), 4)
        Else
            ws.Cells(lig, 3).Value = "n/d"
            ws.Cells(lig, 5).Value = "n/d"
        End If
        If nm > 0 Then ws.Cells(lig, 4).Value = Round(sape / nm * 100, 2) Else ws.Cells(lig, 4).Value = "n/d"
    Next j

    ' Paramètres utilisés
    ws.Cells(LIG_METH, 2).Value = "s = " & ws.Range("B8").Value
    ws.Cells(LIG_METH + 1, 2).Value = "n = " & ws.Range("B5").Value
    ws.Cells(LIG_METH + 2, 2).Value = "n = " & ws.Range("B6").Value
    ws.Cells(LIG_METH + 3, 2).Value = "alpha = " & Format(ws.Range("B7").Value, "0.00")
    ws.Cells(LIG_METH + 4, 2).Value = "Régression linéaire"

    ' Statut : MEILLEUR pour le plus petit RMSE (6 lignes, ML inclus)
    ws.Range(ws.Cells(LIG_METH, 6), ws.Cells(LIG_METH + 5, 6)).ClearContents
    With ws.Range(ws.Cells(LIG_METH, 1), ws.Cells(LIG_METH + 5, 6))
        .Interior.Color = RGB(255, 255, 255)
        .Font.Bold = False
        .Font.Color = RGB(33, 33, 33)
    End With
    ws.Calculate
    meilleur = 0
    rmseMin = 1E+308
    For lig = LIG_METH To LIG_METH + 5
        v = ws.Cells(lig, 5).Value
        If EstNombre(v) Then
            If CDbl(v) < rmseMin Then
                rmseMin = CDbl(v)
                meilleur = lig
            End If
        End If
    Next lig
    If meilleur > 0 Then
        ws.Cells(meilleur, 6).Value = "MEILLEUR"
        With ws.Range(ws.Cells(meilleur, 1), ws.Cells(meilleur, 6))
            .Interior.Color = RGB(200, 230, 201)
            .Font.Bold = True
            .Font.Color = RGB(27, 94, 32)
        End With
    End If
End Sub

' J+1 à J+7 avec la meilleure méthode passive, intervalle +/- 12 %
Private Sub RemplirPrevisionsFutures(ByVal ws As Worksheet, ByVal ligFin As Long)
    Dim y() As Double, n As Long, meth As Long, lig As Long, rmseMin As Double, v As Variant
    Dim k As Long, j As Long, f As Double, s As Long, nW As Long, somme As Double
    Dim alpha As Double, a As Double, b As Double, m As Long, derDate As Date, noms As Variant

    noms = Array("Naïf saisonnier", "Moyenne mobile simple", "Moyenne mobile pondérée", _
                 "Lissage exponentiel", "Projection tendance")
    y = LireColonne(ws, COL_REEL, ligFin)
    n = UBound(y)
    meth = 0
    rmseMin = 1E+308
    For lig = LIG_METH To LIG_METH + 4
        v = ws.Cells(lig, 5).Value
        If EstNombre(v) Then
            If CDbl(v) < rmseMin Then
                rmseMin = CDbl(v)
                meth = lig - LIG_METH
            End If
        End If
    Next lig
    derDate = ws.Cells(ligFin, COL_DATE).Value
    s = CLng(ws.Range("B8").Value)
    alpha = CDbl(ws.Range("B7").Value)
    m = n - CLng(ws.Range("B9").Value)
    If meth = 4 Then CoefTendance y, m, a, b

    For k = 1 To 7
        Select Case meth
            Case 0
                f = y(n - s + ((k - 1) Mod s) + 1)
            Case 1
                nW = CLng(ws.Range("B5").Value)
                somme = 0
                For j = n - nW + 1 To n
                    somme = somme + y(j)
                Next j
                f = somme / nW
            Case 2
                nW = CLng(ws.Range("B6").Value)
                somme = 0
                For j = 1 To nW
                    somme = somme + j * y(n - nW + j)
                Next j
                f = somme / (nW * (nW + 1) / 2)
            Case 3
                f = alpha * y(n) + (1 - alpha) * Nombre(ws.Cells(ligFin, COL_LES).Value)
            Case Else
                f = a + b * (n + k)
        End Select
        If f < 0 Then f = 0
        lig = LIG_FUTUR + k - 1
        ws.Cells(lig, 1).Value = "J+" & k
        ws.Cells(lig, 2).Value = derDate + k
        ws.Cells(lig, 3).Value = Round(f, 2)
        ws.Cells(lig, 4).Value = Round(f * (1 - MARGE_IC), 2)
        ws.Cells(lig, 5).Value = Round(f * (1 + MARGE_IC), 2)
        ws.Cells(lig, 6).Value = noms(meth)
    Next k
End Sub

' Teste alpha de 0.05 à 0.95 (pas 0.05) et garde le plus petit RMSE sur le test
Public Sub OptimiserAlpha()
    Dim ws As Worksheet, nMMS As Long, nMMP As Long, alpha As Double, saison As Long
    Dim nTest As Long, store As String, produit As String
    Dim dates() As Date, ventes() As Double, n As Long, i As Long, r As Long, nb As Long
    Dim a As Double, f As Double, sse As Double, rmse As Double
    Dim meilleurA As Double, meilleurRmse As Double, ancien As Boolean, rapport As String

    Set ws = ThisWorkbook.Worksheets(F_PREV)
    If Not LireParametres(ws, nMMS, nMMP, alpha, saison, nTest, store, produit) Then Exit Sub
    n = ExtraireSerie(store, produit, dates, ventes)
    If n < nTest + 3 Then
        MsgBox "Série trop courte ou introuvable pour " & store & " / " & produit & ".", vbExclamation
        Exit Sub
    End If

    meilleurRmse = 1E+308
    For i = 1 To 19
        a = Round(i * 0.05, 2)
        f = ventes(1)
        sse = 0
        nb = 0
        For r = 2 To n
            If r >= 3 Then f = a * ventes(r - 1) + (1 - a) * f
            If r > n - nTest Then
                sse = sse + (ventes(r) - f) ^ 2
                nb = nb + 1
            End If
        Next r
        rmse = Sqr(sse / nb)
        rapport = rapport & "alpha = " & Format(a, "0.00") & "   RMSE = " & Format(rmse, "0.000") & vbCrLf
        If rmse < meilleurRmse Then
            meilleurRmse = rmse
            meilleurA = a
        End If
    Next i

    ws.Range("B7").Value = meilleurA
    ancien = mbSilencieux
    mbSilencieux = True
    LancerPrevisions
    mbSilencieux = ancien
    Info "Meilleur alpha : " & Format(meilleurA, "0.00") & " (RMSE test = " & Format(meilleurRmse, "0.000") & ")" & _
         vbCrLf & "Valeur écrite en B7, prévisions relancées." & vbCrLf & vbCrLf & rapport
End Sub


'==============================================================================
' DEMARRAGE ET UTILITAIRES
'==============================================================================
' Appelée par Workbook_Open (ThisWorkbook) : masque les feuilles techniques, ouvre le centre de contrôle
Public Sub DemarrerApplication()
    On Error Resume Next
    ThisWorkbook.Worksheets(F_DON).Visible = xlSheetHidden
    ThisWorkbook.Worksheets(F_TEST).Visible = xlSheetHidden
    On Error GoTo 0
    frmLancement.Show
End Sub

' Feuille existante, ou créée si elle a été supprimée
Private Function Feuille(ByVal nom As String) As Worksheet
    On Error Resume Next
    Set Feuille = ThisWorkbook.Worksheets(nom)
    On Error GoTo 0
    If Feuille Is Nothing Then
        Set Feuille = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        Feuille.Name = nom
    End If
End Function

' (Re)définit une plage nommée au niveau du classeur
Private Sub NommerPlage(ByVal nom As String, ByVal r As Range)
    On Error Resume Next
    ThisWorkbook.Names(nom).Delete
    On Error GoTo 0
    ThisWorkbook.Names.Add Name:=nom, RefersTo:="='" & r.Worksheet.Name & "'!" & r.Address
End Sub

' Dossier racine du projet (contient excel_bridge.py). Le classeur est normalement
' enregistré dans output/, mais on accepte aussi qu'il soit à la racine.
Private Function DossierProjet() As String
    If Len(Dir(ThisWorkbook.Path & "\excel_bridge.py")) > 0 Then
        DossierProjet = ThisWorkbook.Path
    Else
        DossierProjet = Left$(ThisWorkbook.Path, InStrRev(ThisWorkbook.Path, "\") - 1)
    End If
End Function

Public Function DossierSorties() As String
    DossierSorties = DossierProjet() & "\output"
End Function

' Python : chemin écrit par generate_workbook.py dans PARAMS!B8, sinon "py -3", sinon "python"
Private Function CommandePython() As String
    Dim chemin As String
    chemin = Trim$(Param("B8"))
    If Len(chemin) > 0 Then
        If Len(Dir(chemin)) > 0 Then
            CommandePython = """" & chemin & """"
            Exit Function
        End If
    End If
    If Len(Dir(Environ$("WINDIR") & "\py.exe")) > 0 Or _
       Len(Dir(Environ$("LOCALAPPDATA") & "\Programs\Python\Launcher\py.exe")) > 0 Then
        CommandePython = "py -3"
    Else
        CommandePython = "python"
    End If
End Function

Private Sub Info(ByVal texte As String)
    If Not mbSilencieux Then MsgBox texte, vbInformation, "Prévision de la demande"
End Sub

Private Function StoreValide(ByVal store As String) As Boolean
    Select Case store
        Case "CA_1", "TX_1", "WI_1": StoreValide = True
    End Select
End Function

Private Function EstNombre(ByVal v As Variant) As Boolean
    Select Case VarType(v)
        Case vbDouble, vbSingle, vbLong, vbInteger, vbCurrency, vbDecimal, vbByte: EstNombre = True
    End Select
End Function

Private Function EstEntierPositif(ByVal v As Variant) As Boolean
    If EstNombre(v) Then
        If v >= 1 And v = Int(v) Then EstEntierPositif = True
    End If
End Function

Private Function Nombre(ByVal v As Variant) As Double
    If EstNombre(v) Then Nombre = CDbl(v)
End Function

Private Function Fmt(ByVal v As Variant) As String
    If EstNombre(v) Then Fmt = Format(v, "0.00") Else Fmt = "n/d"
End Function

Private Function LireColonne(ByVal ws As Worksheet, ByVal col As Long, ByVal ligFin As Long) As Double()
    Dim v As Variant, res() As Double, i As Long, n As Long
    v = ws.Range(ws.Cells(PREM_LIG, col), ws.Cells(ligFin, col)).Value
    n = UBound(v, 1)
    ReDim res(1 To n)
    For i = 1 To n
        res(i) = Nombre(v(i, 1))
    Next i
    LireColonne = res
End Function

Private Sub EcrireColonne(ByVal ws As Worksheet, ByVal col As Long, ByRef res() As Variant)
    ws.Range(ws.Cells(PREM_LIG, col), ws.Cells(PREM_LIG + UBound(res, 1) - 1, col)).Value = res
End Sub
