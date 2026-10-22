Attribute VB_Name = "ModulePrincipal"
'==============================================================================
' ModulePrincipal - Système de prévision de la demande (M5 Walmart)
' Version 2.0 - Navigation multi-feuilles, raccourcis période, calendrier.
' Généré par generate_dashboard.py
'
' Import : Alt+F11 > Fichier > Importer un fichier... > ModulePrincipal.bas
' Puis coller dans ThisWorkbook :
'     Private Sub Workbook_Open()
'         DemarrerApplication
'     End Sub
'==============================================================================
Option Explicit

' --- Noms des feuilles ---
Private Const F_ACCUEIL As String = "ACCUEIL"
Private Const F_DASH As String = "DASHBOARD"
Private Const F_PREV As String = "PREVISION"
Private Const F_ML As String = "ML_RESULTATS"
Private Const F_GEST As String = "GESTION_DATA"
Private Const F_DON As String = "DONNEES"
Private Const F_PRODUITS As String = "PRODUITS"
Private Const F_CALENDRIER As String = "CALENDRIER"
Private Const F_MAGASINS As String = "MAGASINS"
Private Const F_CATEGORIES As String = "CATEGORIES"
Private Const F_EVENEMENTS As String = "EVENEMENTS"

' --- PREVISION ---
Private Const PREM_LIG As Long = 5
Private Const COL_DATE As Long = 20
Private Const COL_REEL As Long = 21
Private Const COL_XMAS As Long = 22
Private Const COL_NAIF As Long = 23
Private Const COL_MMS As Long = 24
Private Const COL_MMP As Long = 25
Private Const COL_LES As Long = 26
Private Const COL_TEND As Long = 27
Private Const LIG_METH As Long = 15
Private Const LIG_FUTUR As Long = 24
Private Const MARGE_IC As Double = 0.12

' --- GESTION_DATA ---
Private Const LIG_TRANS As Long = 21
Private Const NB_TRANS As Long = 30
Private Const MAX_TRANS As Long = 500
Private Const LIG_PROD As Long = 16
Private Const MAX_PROD As Long = 50

' --- ML_RESULTATS ---
Private Const LIG_ML_T1 As Long = 9
Private Const LIG_ML_T2 As Long = 16
Private Const MAX_ML_T2 As Long = 50

' Mode silencieux (supprime les MsgBox d'information)
Private mbSilencieux As Boolean


'==============================================================================
' NAVIGATION
'==============================================================================
Public Sub AllerAccueil()
    AllerFeuille F_ACCUEIL
End Sub

Public Sub AllerDashboard()
    AllerFeuille F_DASH
End Sub

Public Sub AllerPrevision()
    AllerFeuille F_PREV
End Sub

Public Sub AllerML()
    AllerFeuille F_ML
End Sub

Public Sub AllerGestion()
    AllerFeuille F_GEST
End Sub

Public Sub AllerProduits()
    AllerFeuille F_PRODUITS
End Sub

Public Sub AllerCalendrier()
    AllerFeuille F_CALENDRIER
End Sub

Public Sub AllerMagasins()
    AllerFeuille F_MAGASINS
End Sub

Public Sub AllerCategories()
    AllerFeuille F_CATEGORIES
End Sub

Private Sub AllerFeuille(ByVal nom As String)
    With ThisWorkbook.Worksheets(nom)
        If .Visible <> xlSheetVisible Then .Visible = xlSheetVisible
        .Activate
        .Range("A1").Select
    End With
End Sub


'==============================================================================
' PREVISION PASSIVE (5 METHODES) — identique v1
'==============================================================================
Public Sub LancerPrevisions()
    Dim ws As Worksheet, nMMS As Long, nMMP As Long, alpha As Double
    Dim saison As Long, nTest As Long, store As String, produit As String
    Dim dates() As Date, ventes() As Double, n As Long, i As Long, nMin As Long
    Dim sortie() As Variant, ligFin As Long, ligTest As Long, meilleur As String

    Set ws = ThisWorkbook.Worksheets(F_PREV)
    If Not LireParametres(ws, nMMS, nMMP, alpha, saison, nTest, store, produit) Then Exit Sub

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

    NaifSaisonnier ws, ligFin
    MoyenneMobileSimple ws, ligFin
    MoyenneMobilePonderee ws, ligFin
    LissageExponentiel ws, ligFin
    ProjectionTendance ws, ligTest, ligFin
    CalculerErreurs ws, ligTest, ligFin
    RemplirPrevisionsFutures ws, ligFin
    MettreAJourGraphique ws, ligTest, ligFin, "Réel vs prévisions — " & store & " / " & produit

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

    ws.Cells(LIG_METH, 2).Value = "s = " & ws.Range("B8").Value
    ws.Cells(LIG_METH + 1, 2).Value = "n = " & ws.Range("B5").Value
    ws.Cells(LIG_METH + 2, 2).Value = "n = " & ws.Range("B6").Value
    ws.Cells(LIG_METH + 3, 2).Value = "alpha = " & Format(ws.Range("B7").Value, "0.00")
    ws.Cells(LIG_METH + 4, 2).Value = "Régression linéaire"

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

Private Sub MettreAJourGraphique(ByVal ws As Worksheet, ByVal ligTest As Long, ByVal ligFin As Long, ByVal titre As String)
    Dim cols As Variant, i As Long
    If ws.ChartObjects.Count = 0 Then Exit Sub
    cols = Array(COL_REEL, COL_NAIF, COL_MMS, COL_MMP, COL_LES, COL_TEND)
    With ws.ChartObjects(1).Chart
        For i = 1 To .SeriesCollection.Count
            If i - 1 <= UBound(cols) Then
                .SeriesCollection(i).Values = ws.Range(ws.Cells(ligTest, cols(i - 1)), ws.Cells(ligFin, cols(i - 1)))
                .SeriesCollection(i).XValues = ws.Range(ws.Cells(ligTest, COL_DATE), ws.Cells(ligFin, COL_DATE))
            End If
        Next i
        .HasTitle = True
        .ChartTitle.Text = titre
    End With
End Sub

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
' GESTION DES DONNEES
'==============================================================================
Public Sub AjouterTransaction()
    Dim wg As Worksheet, wd As Worksheet, vDate As Variant, vQte As Variant, vPrix As Variant
    Dim store As String, produit As String, evt As String, msg As String
    Dim lig As Long, d As Date, ligne(1 To 1, 1 To 11) As Variant

    Set wg = ThisWorkbook.Worksheets(F_GEST)
    Set wd = ThisWorkbook.Worksheets(F_DON)
    vDate = wg.Range("C6").Value
    store = Trim$(CStr(wg.Range("C7").Value))
    produit = Trim$(CStr(wg.Range("C8").Value))
    vQte = wg.Range("C9").Value
    vPrix = wg.Range("C10").Value
    evt = Trim$(CStr(wg.Range("C11").Value))

    If Not IsDate(vDate) Then msg = msg & "- Date (C6) : date valide attendue" & vbCrLf
    If Not StoreValide(store) Then msg = msg & "- Store (C7) : CA_1, TX_1 ou WI_1" & vbCrLf
    If produit = "" Then msg = msg & "- Produit (C8) : obligatoire" & vbCrLf
    If Not EstNombre(vQte) Then
        msg = msg & "- Quantité (C9) : nombre positif attendu" & vbCrLf
    ElseIf CDbl(vQte) <= 0 Then
        msg = msg & "- Quantité (C9) : doit être strictement positive" & vbCrLf
    End If
    If Not IsEmpty(vPrix) Then
        If Not EstNombre(vPrix) Then msg = msg & "- Prix (C10) : nombre attendu" & vbCrLf
    End If
    If msg <> "" Then
        MsgBox "Transaction refusée :" & vbCrLf & vbCrLf & msg, vbExclamation, "Validation"
        Exit Sub
    End If

    d = CDate(vDate)
    lig = wd.Cells(wd.Rows.Count, 1).End(xlUp).Row + 1
    ligne(1, 1) = d
    ligne(1, 2) = store
    ligne(1, 3) = produit
    ligne(1, 4) = CategorieProduit(store, produit)
    ligne(1, 5) = CDbl(vQte)
    If Not IsEmpty(vPrix) Then ligne(1, 6) = CDbl(vPrix)
    ligne(1, 7) = 0
    If evt <> "" And UCase$(evt) <> "AUCUN" Then ligne(1, 8) = 1 Else ligne(1, 8) = 0
    ligne(1, 9) = Weekday(d, vbMonday) - 1
    ligne(1, 10) = Month(d)
    ligne(1, 11) = Year(d)
    wd.Range("A" & lig & ":K" & lig).Value = ligne
    wd.Range("A" & lig).NumberFormat = "dd/mm/yyyy"
    wd.Range("F" & lig).NumberFormat = "0.00"

    AfficherDernieresTransactions
    wg.Range("C9:C10").ClearContents
    MsgBox "Transaction ajoutée (ligne " & lig & " de DONNEES) :" & vbCrLf & _
           Format(d, "dd/mm/yyyy") & " - " & store & " - " & produit & " - " & vQte & " unité(s)", _
           vbInformation, "Confirmation"
End Sub

Private Sub AfficherDernieresTransactions()
    Dim wg As Worksheet, wd As Worksheet, derLig As Long, prem As Long, donnees As Variant
    Dim sortie() As Variant, i As Long, j As Long, nb As Long

    Set wg = ThisWorkbook.Worksheets(F_GEST)
    Set wd = ThisWorkbook.Worksheets(F_DON)
    wg.Range(wg.Cells(LIG_TRANS, 2), wg.Cells(LIG_TRANS + MAX_TRANS - 1, 12)).ClearContents
    derLig = wd.Cells(wd.Rows.Count, 1).End(xlUp).Row
    If derLig < 2 Then Exit Sub
    prem = derLig - NB_TRANS + 1
    If prem < 2 Then prem = 2
    donnees = wd.Range("A" & prem & ":K" & derLig).Value
    nb = UBound(donnees, 1)
    ReDim sortie(1 To nb, 1 To 11)
    For i = 1 To nb
        For j = 1 To 11
            sortie(i, j) = donnees(nb - i + 1, j)
        Next j
    Next i
    wg.Range(wg.Cells(LIG_TRANS, 2), wg.Cells(LIG_TRANS + nb - 1, 12)).Value = sortie
    wg.Range("B19").Value = NB_TRANS & " dernières transactions (les plus récentes en haut) — " & _
                            Format(derLig - 1, "#,##0") & " lignes dans DONNEES"
End Sub

Public Sub AppliquerFiltre()
    Dim wg As Worksheet, wd As Worksheet, texte As String, store As String, cat As String
    Dim vDu As Variant, vAu As Variant, derLig As Long, rng As Range, vis As Range, zone As Range
    Dim sortie() As Variant, nb As Long, total As Long, a As Long, r As Long, j As Long
    Dim filtre As Boolean

    Set wg = ThisWorkbook.Worksheets(F_GEST)
    Set wd = ThisWorkbook.Worksheets(F_DON)
    texte = Trim$(CStr(wg.Range("I6").Value))
    store = Trim$(CStr(wg.Range("I7").Value))
    cat = Trim$(CStr(wg.Range("I8").Value))
    vDu = wg.Range("I11").Value
    vAu = wg.Range("I12").Value
    If (Not IsEmpty(vDu) And Not IsDate(vDu)) Or (Not IsEmpty(vAu) And Not IsDate(vAu)) Then
        MsgBox "Période invalide : saisir des dates (jj/mm/aaaa).", vbExclamation
        Exit Sub
    End If
    derLig = wd.Cells(wd.Rows.Count, 1).End(xlUp).Row
    If derLig < 2 Then Exit Sub

    On Error GoTo Erreur
    Application.ScreenUpdating = False
    If wd.AutoFilterMode Then wd.AutoFilterMode = False
    Set rng = wd.Range("A1:K" & derLig)
    If texte <> "" Then
        rng.AutoFilter Field:=3, Criteria1:="*" & texte & "*"
        filtre = True
    End If
    If store <> "" And store <> "Tous" Then
        rng.AutoFilter Field:=2, Criteria1:=store
        filtre = True
    End If
    If cat <> "" And cat <> "Toutes" Then
        rng.AutoFilter Field:=4, Criteria1:=cat
        filtre = True
    End If
    If IsDate(vDu) And IsDate(vAu) Then
        rng.AutoFilter Field:=1, Criteria1:=">=" & CLng(CDate(vDu)), Operator:=xlAnd, _
                       Criteria2:="<=" & CLng(CDate(vAu))
        filtre = True
    ElseIf IsDate(vDu) Then
        rng.AutoFilter Field:=1, Criteria1:=">=" & CLng(CDate(vDu))
        filtre = True
    ElseIf IsDate(vAu) Then
        rng.AutoFilter Field:=1, Criteria1:="<=" & CLng(CDate(vAu))
        filtre = True
    End If
    If Not filtre Then
        Application.ScreenUpdating = True
        AfficherDernieresTransactions
        Info "Aucun critère saisi : affichage des dernières transactions."
        Exit Sub
    End If

    ReDim sortie(1 To MAX_TRANS, 1 To 11)
    On Error Resume Next
    Set vis = wd.Range("A2:A" & derLig).SpecialCells(xlCellTypeVisible)
    On Error GoTo Erreur
    If Not vis Is Nothing Then
        total = vis.Cells.Count
        For a = vis.Areas.Count To 1 Step -1
            Set zone = vis.Areas(a)
            For r = zone.Rows.Count To 1 Step -1
                nb = nb + 1
                For j = 1 To 11
                    sortie(nb, j) = wd.Cells(zone.Cells(r, 1).Row, j).Value
                Next j
                If nb = MAX_TRANS Then Exit For
            Next r
            If nb = MAX_TRANS Then Exit For
        Next a
    End If
    wd.AutoFilterMode = False

    wg.Range(wg.Cells(LIG_TRANS, 2), wg.Cells(LIG_TRANS + MAX_TRANS - 1, 12)).ClearContents
    If nb > 0 Then wg.Range(wg.Cells(LIG_TRANS, 2), wg.Cells(LIG_TRANS + nb - 1, 12)).Value = sortie
    wg.Range("B19").Value = "Filtre appliqué : " & Format(total, "#,##0") & " ligne(s) trouvée(s) — " & _
                            nb & " affichée(s) (max. " & MAX_TRANS & ")"
    Application.ScreenUpdating = True
    Info Format(total, "#,##0") & " ligne(s) correspondent aux critères."
    Exit Sub
Erreur:
    If wd.AutoFilterMode Then wd.AutoFilterMode = False
    Application.ScreenUpdating = True
    MsgBox "Erreur dans AppliquerFiltre : " & Err.Description, vbCritical
End Sub

Public Sub ReinitialisationFiltre()
    Dim wg As Worksheet, wd As Worksheet
    Set wg = ThisWorkbook.Worksheets(F_GEST)
    Set wd = ThisWorkbook.Worksheets(F_DON)
    wg.Range("I6").ClearContents
    wg.Range("I7").Value = "Tous"
    wg.Range("I8").Value = "Toutes"
    wg.Range("I11:I12").ClearContents
    If wd.AutoFilterMode Then wd.AutoFilterMode = False
    AfficherDernieresTransactions
    Info "Filtres réinitialisés."
End Sub

' --- Raccourcis de période ---
Public Sub Filtrer30()
    AppliquerFiltrePeriode 30
End Sub

Public Sub Filtrer90()
    AppliquerFiltrePeriode 90
End Sub

Public Sub Filtrer150()
    AppliquerFiltrePeriode 150
End Sub

Public Sub FiltrerAnnee()
    Dim wg As Worksheet, wd As Worksheet, dateMax As Date
    Set wg = ThisWorkbook.Worksheets(F_GEST)
    Set wd = ThisWorkbook.Worksheets(F_DON)
    dateMax = Application.WorksheetFunction.Max(wd.Columns(1))
    wg.Range("I11").Value = DateSerial(Year(dateMax) - 1, Month(dateMax), Day(dateMax))
    wg.Range("I12").Value = dateMax
    AppliquerFiltre
End Sub

Private Sub AppliquerFiltrePeriode(ByVal jours As Long)
    Dim wg As Worksheet, wd As Worksheet, dateMax As Date
    Set wg = ThisWorkbook.Worksheets(F_GEST)
    Set wd = ThisWorkbook.Worksheets(F_DON)
    dateMax = Application.WorksheetFunction.Max(wd.Columns(1))
    wg.Range("I11").Value = dateMax - jours + 1
    wg.Range("I12").Value = dateMax
    AppliquerFiltre
End Sub

Public Sub AjouterProduit()
    Dim wg As Worksheet, tbl As Variant, store As String, produit As String, cat As String
    Dim i As Long, nb As Long, lig As Long

    Set wg = ThisWorkbook.Worksheets(F_GEST)
    store = Trim$(CStr(wg.Range("O6").Value))
    produit = UCase$(Trim$(CStr(wg.Range("O7").Value)))
    cat = UCase$(Trim$(CStr(wg.Range("O8").Value)))
    If Not StoreValide(store) Then
        MsgBox "Store invalide (O6) : CA_1, TX_1 ou WI_1.", vbExclamation
        Exit Sub
    End If
    If produit = "" Then
        MsgBox "Saisir l'identifiant du produit en O7 (ex. FOODS_3_090).", vbExclamation
        Exit Sub
    End If
    If cat = "" Then cat = CategorieProduit(store, produit)

    tbl = wg.Range(wg.Cells(LIG_PROD, 14), wg.Cells(LIG_PROD + MAX_PROD - 1, 17)).Value
    For i = 1 To MAX_PROD
        If Trim$(CStr(tbl(i, 2))) <> "" Then
            nb = nb + 1
            If CStr(tbl(i, 1)) = store And UCase$(CStr(tbl(i, 2))) = produit Then
                MsgBox produit & " est déjà suivi pour " & store & ".", vbExclamation, "Doublon"
                Exit Sub
            End If
        End If
    Next i
    If nb >= MAX_PROD Then
        MsgBox "Limite atteinte : " & MAX_PROD & " produits suivis au maximum.", vbExclamation
        Exit Sub
    End If

    lig = LIG_PROD + nb
    wg.Cells(lig, 14).Value = store
    wg.Cells(lig, 15).Value = produit
    wg.Cells(lig, 16).Value = cat
    wg.Cells(lig, 17).Value = Date
    wg.Cells(lig, 17).NumberFormat = "dd/mm/yyyy"
    TrierProduits
    VerifierSelections
    wg.Range("O7").ClearContents
    MsgBox "Produit ajouté : " & store & " / " & produit & " (" & nb + 1 & "/" & MAX_PROD & ")." & vbCrLf & _
           "La liste déroulante 'Produit' (PREVISION!B11) est à jour.", vbInformation
End Sub

Public Sub SupprimerProduit()
    Dim wg As Worksheet, tbl As Variant, nouv() As Variant, store As String, produit As String
    Dim i As Long, j As Long, k As Long, trouve As Long

    Set wg = ThisWorkbook.Worksheets(F_GEST)
    store = Trim$(CStr(wg.Range("O6").Value))
    produit = UCase$(Trim$(CStr(wg.Range("O7").Value)))
    tbl = wg.Range(wg.Cells(LIG_PROD, 14), wg.Cells(LIG_PROD + MAX_PROD - 1, 17)).Value
    For i = 1 To MAX_PROD
        If CStr(tbl(i, 1)) = store And UCase$(CStr(tbl(i, 2))) = produit And produit <> "" Then trouve = i
    Next i
    If trouve = 0 Then
        MsgBox "Produit introuvable : saisir le store en O6 et le produit en O7.", vbExclamation
        Exit Sub
    End If
    If MsgBox("Supprimer " & produit & " (" & store & ") de la liste des produits suivis ?" & vbCrLf & _
              "L'historique reste dans DONNEES.", vbYesNo + vbQuestion, "Confirmation") <> vbYes Then Exit Sub

    ReDim nouv(1 To MAX_PROD, 1 To 4)
    For i = 1 To MAX_PROD
        If i <> trouve And Trim$(CStr(tbl(i, 2))) <> "" Then
            k = k + 1
            For j = 1 To 4
                nouv(k, j) = tbl(i, j)
            Next j
        End If
    Next i
    wg.Range(wg.Cells(LIG_PROD, 14), wg.Cells(LIG_PROD + MAX_PROD - 1, 17)).Value = nouv
    VerifierSelections
    wg.Range("O7").ClearContents
    MsgBox produit & " (" & store & ") supprimé. " & k & " produit(s) suivi(s).", vbInformation
End Sub

Private Sub TrierProduits()
    Dim wg As Worksheet
    Set wg = ThisWorkbook.Worksheets(F_GEST)
    With wg.Sort
        .SortFields.Clear
        .SortFields.Add Key:=wg.Range("N16:N65"), Order:=xlAscending
        .SortFields.Add Key:=wg.Range("O16:O65"), Order:=xlAscending
        .SetRange wg.Range("N16:Q65")
        .Header = xlNo
        .Apply
    End With
End Sub

Private Sub VerifierSelections()
    CorrigerSelection ThisWorkbook.Worksheets(F_PREV).Range("B10"), ThisWorkbook.Worksheets(F_PREV).Range("B11")
    CorrigerSelection ThisWorkbook.Worksheets(F_ML).Range("C4"), ThisWorkbook.Worksheets(F_ML).Range("E4")
End Sub

Private Sub CorrigerSelection(ByVal celStore As Range, ByVal celProd As Range)
    Dim tbl As Variant, i As Long, premier As String, trouve As Boolean
    tbl = ThisWorkbook.Worksheets(F_GEST).Range("N16:O65").Value
    For i = 1 To MAX_PROD
        If CStr(tbl(i, 1)) = CStr(celStore.Value) And CStr(tbl(i, 2)) <> "" Then
            If premier = "" Then premier = CStr(tbl(i, 2))
            If CStr(tbl(i, 2)) = CStr(celProd.Value) Then trouve = True
        End If
    Next i
    If Not trouve Then celProd.Value = premier
End Sub

Private Function CategorieProduit(ByVal store As String, ByVal produit As String) As String
    Dim tbl As Variant, i As Long
    tbl = ThisWorkbook.Worksheets(F_GEST).Range("N16:P65").Value
    For i = 1 To MAX_PROD
        If CStr(tbl(i, 2)) = produit And CStr(tbl(i, 3)) <> "" Then
            CategorieProduit = CStr(tbl(i, 3))
            If CStr(tbl(i, 1)) = store Then Exit Function
        End If
    Next i
    If CategorieProduit = "" And InStr(produit, "_") > 0 Then CategorieProduit = Split(produit, "_")(0)
End Function


'==============================================================================
' CALENDRIER
'==============================================================================
Public Sub ApercuCalendrier()
    Dim wc As Worksheet, wd As Worksheet, we As Worksheet
    Dim d As Variant, tbl As Variant, i As Long, j As Long
    Dim evts As String, snapCA As Long, snapTX As Long, snapWI As Long
    Dim wday As String, mois As String, annee As String
    Dim nbTx As Long, totalVentes As Double, dtStr As Date

    Set wc = ThisWorkbook.Worksheets(F_CALENDRIER)
    Set wd = ThisWorkbook.Worksheets(F_DON)
    Set we = ThisWorkbook.Worksheets(F_EVENEMENTS)

    d = wc.Range("C6").Value
    If Not IsDate(d) Then
        MsgBox "Saisir une date valide en C6.", vbExclamation
        Exit Sub
    End If
    dtStr = CDate(d)

    ' Infos calendaires (premier enregistrement correspondant dans DONNEES)
    snapCA = 0: snapTX = 0: snapWI = 0: wday = "": mois = "": annee = "": evts = ""
    tbl = wd.Range(wd.Cells(2, 1), wd.Cells(wd.Cells(wd.Rows.Count, 1).End(xlUp).Row, 11)).Value
    For i = 1 To UBound(tbl, 1)
        If IsDate(tbl(i, 1)) Then
            If CLng(CDate(tbl(i, 1))) = CLng(dtStr) Then
                If wday = "" Then
                    wday = WeekdayName(Weekday(dtStr, vbMonday))
                    mois = MonthName(Month(dtStr))
                    annee = CStr(Year(dtStr))
                    snapCA = Nombre(tbl(i, 7))
                End If
            End If
        End If
    Next i

    ' Événements depuis DONNEES (has_event + noms collectés)
    Dim evtSet As Object
    Set evtSet = CreateObject("Scripting.Dictionary")
    For i = 1 To UBound(tbl, 1)
        If IsDate(tbl(i, 1)) Then
            If CLng(CDate(tbl(i, 1))) = CLng(dtStr) And Nombre(tbl(i, 8)) = 1 Then
                ' pas de nom d'événement dans DONNEES ; signaler via compteur
                evtSet("Evénement signalé") = 1
            End If
        End If
    Next i
    ' Événements personnalisés
    Dim derEv As Long
    derEv = we.Cells(we.Rows.Count, 1).End(xlUp).Row
    If derEv >= 2 Then
        tbl = we.Range("A2:C" & derEv).Value
        For i = 1 To UBound(tbl, 1)
            If IsDate(tbl(i, 1)) Then
                If CLng(CDate(tbl(i, 1))) = CLng(dtStr) Then
                    evtSet(CStr(tbl(i, 2))) = 1
                End If
            End If
        Next i
    End If
    Dim k As Variant
    For Each k In evtSet.Keys
        If evts <> "" Then evts = evts & ", "
        evts = evts & CStr(k)
    Next k
    If evts = "" Then evts = "Aucun événement"

    ' Comptage transactions
    nbTx = Application.WorksheetFunction.CountIfs(wd.Columns(1), CLng(dtStr))
    totalVentes = Application.WorksheetFunction.SumIfs(wd.Columns(5), wd.Columns(1), CLng(dtStr))

    wc.Range("F7").Value = Format(dtStr, "dd/mm/yyyy")
    wc.Range("F8").Value = wday
    wc.Range("F9").Value = mois
    wc.Range("F10").Value = annee
    wc.Range("F11").Value = snapCA
    wc.Range("F12").Value = snapTX
    wc.Range("F13").Value = snapWI
    wc.Range("F14").Value = evts
    wc.Range("F15").Value = nbTx
    wc.Range("F16").Value = totalVentes
End Sub

Public Sub AjouterEvenement()
    Dim wc As Worksheet, we As Worksheet, d As Variant, nom As String, typ As String
    Dim derLig As Long, i As Long, tbl As Variant

    Set wc = ThisWorkbook.Worksheets(F_CALENDRIER)
    Set we = ThisWorkbook.Worksheets(F_EVENEMENTS)
    d = wc.Range("L6").Value
    nom = Trim$(CStr(wc.Range("L7").Value))
    typ = Trim$(CStr(wc.Range("L8").Value))

    If Not IsDate(d) Then
        MsgBox "Date invalide (L6).", vbExclamation
        Exit Sub
    End If
    If nom = "" Then
        MsgBox "Nom obligatoire (L7).", vbExclamation
        Exit Sub
    End If

    ' Doublon ?
    derLig = we.Cells(we.Rows.Count, 1).End(xlUp).Row
    If derLig >= 2 Then
        tbl = we.Range("A2:B" & derLig).Value
        For i = 1 To UBound(tbl, 1)
            If IsDate(tbl(i, 1)) Then
                If CLng(CDate(tbl(i, 1))) = CLng(CDate(d)) And LCase$(CStr(tbl(i, 2))) = LCase$(nom) Then
                    MsgBox "Evénement déjà existant à cette date.", vbExclamation, "Doublon"
                    Exit Sub
                End If
            End If
        Next i
    End If

    If MsgBox("Ajouter l'événement '" & nom & "' (" & typ & ") au " & Format(CDate(d), "dd/mm/yyyy") & " ?", _
              vbYesNo + vbQuestion, "Confirmation") <> vbYes Then Exit Sub

    derLig = we.Cells(we.Rows.Count, 1).End(xlUp).Row + 1
    we.Cells(derLig, 1).Value = CDate(d)
    we.Cells(derLig, 1).NumberFormat = "dd/mm/yyyy"
    we.Cells(derLig, 2).Value = nom
    we.Cells(derLig, 3).Value = typ
    we.Cells(derLig, 4).Value = Now
    we.Cells(derLig, 4).NumberFormat = "dd/mm/yyyy hh:mm"

    ' Mise à jour de la liste des événements
    Dim lig As Long
    For i = 1 To 20
        lig = 13 + i - 1
        If i <= derLig - 1 Then
            wc.Cells(lig, 11).Value = we.Cells(i + 1, 1).Value
            wc.Cells(lig, 11).NumberFormat = "dd/mm/yyyy"
            wc.Cells(lig, 12).Value = we.Cells(i + 1, 2).Value
            wc.Cells(lig, 13).Value = we.Cells(i + 1, 3).Value
            wc.Cells(lig, 14).Value = we.Cells(i + 1, 4).Value
            wc.Cells(lig, 14).NumberFormat = "dd/mm/yyyy hh:mm"
        End If
    Next i

    wc.Range("L7").ClearContents
    MsgBox "Evénement ajouté.", vbInformation
End Sub

Public Sub ReinitialiserCalendrier()
    Dim wc As Worksheet, wd As Worksheet, dateMax As Date
    Set wc = ThisWorkbook.Worksheets(F_CALENDRIER)
    Set wd = ThisWorkbook.Worksheets(F_DON)
    dateMax = Application.WorksheetFunction.Max(wd.Columns(1))
    wc.Range("C6").Value = dateMax
    ApercuCalendrier
End Sub


'==============================================================================
' ML ET EXPORTS
'==============================================================================
Public Sub ImporterResultatsML()
    Dim wd As Worksheet, wm As Worksheet, chemin As String, note As String, meilleur As String
    Dim nP As Long, nA As Long, i As Long, k As Long, r As Long, m As Long, g As Long, lig As Long
    Dim dP() As Date, sP() As String, iP() As String, yP() As Double, pP() As Double
    Dim dA() As Date, sA() As String, iA() As String, yA() As Double, pA() As Double
    Dim cles As Object, grp As Object, cle As String, sortie() As Variant
    Dim acc() As Double, gNiv() As String, gSto() As String, gIt() As String, nG As Long
    Dim idx As Long, yv As Double, e As Double, v As Variant, iGlob As Long
    Dim lstS() As String, lstI() As String, nS As Long, nI As Long, rP As Variant, rA As Variant

    Set wd = ThisWorkbook.Worksheets(F_DON)
    Set wm = ThisWorkbook.Worksheets(F_ML)
    chemin = ThisWorkbook.Path & Application.PathSeparator
    nP = LireCSVPrevisions(chemin & "forecasts_passive_ml.csv", dP, sP, iP, yP, pP)
    nA = LireCSVPrevisions(chemin & "forecasts_active_ml.csv", dA, sA, iA, yA, pA)
    If nP = 0 And nA = 0 Then
        MsgBox "Aucun fichier de prévisions ML trouvé dans :" & vbCrLf & chemin & vbCrLf & _
               "(forecasts_passive_ml.csv / forecasts_active_ml.csv)", vbExclamation
        Exit Sub
    End If

    On Error GoTo Erreur
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Set cles = CreateObject("Scripting.Dictionary")
    ReDim sortie(1 To nP + nA, 1 To 6)
    For i = 1 To nP
        cle = CStr(CLng(dP(i))) & "|" & sP(i) & "|" & iP(i)
        k = k + 1
        cles(cle) = k
        sortie(k, 1) = dP(i): sortie(k, 2) = sP(i): sortie(k, 3) = iP(i)
        sortie(k, 4) = yP(i): sortie(k, 5) = pP(i)
    Next i
    For i = 1 To nA
        cle = CStr(CLng(dA(i))) & "|" & sA(i) & "|" & iA(i)
        If cles.Exists(cle) Then
            r = cles(cle)
        Else
            k = k + 1
            r = k
            cles(cle) = k
            sortie(r, 1) = dA(i): sortie(r, 2) = sA(i): sortie(r, 3) = iA(i): sortie(r, 4) = yA(i)
        End If
        sortie(r, 6) = pA(i)
    Next i

    wd.Range("M2:R" & wd.Rows.Count).ClearContents
    wd.Range("M2").Resize(k, 6).Value = sortie
    wd.Range("M2").Resize(k, 1).NumberFormat = "dd/mm/yyyy"

    Set grp = CreateObject("Scripting.Dictionary")
    ReDim acc(1 To k + 10, 1 To 10)
    ReDim gNiv(1 To k + 10): ReDim gSto(1 To k + 10): ReDim gIt(1 To k + 10)
    For i = 1 To k
        yv = sortie(i, 4)
        For g = 1 To 3
            Select Case g
                Case 1: cle = "G"
                Case 2: cle = "S|" & sortie(i, 2)
                Case Else: cle = "I|" & sortie(i, 2) & "|" & sortie(i, 3)
            End Select
            If Not grp.Exists(cle) Then
                nG = nG + 1
                grp(cle) = nG
                gNiv(nG) = Left$(cle, 1)
                gSto(nG) = CStr(sortie(i, 2))
                gIt(nG) = CStr(sortie(i, 3))
            End If
            idx = grp(cle)
            For m = 0 To 1
                v = sortie(i, 5 + m)
                If Not IsEmpty(v) Then
                    e = yv - CDbl(v)
                    acc(idx, m * 5 + 1) = acc(idx, m * 5 + 1) + 1
                    acc(idx, m * 5 + 2) = acc(idx, m * 5 + 2) + Abs(e)
                    acc(idx, m * 5 + 3) = acc(idx, m * 5 + 3) + e * e
                    If yv > 0 Then
                        acc(idx, m * 5 + 4) = acc(idx, m * 5 + 4) + 1
                        acc(idx, m * 5 + 5) = acc(idx, m * 5 + 5) + Abs(e) / yv
                    End If
                End If
            Next m
        Next g
    Next i

    ReDim lstS(1 To nG): ReDim lstI(1 To nG)
    For g = 1 To nG
        If gNiv(g) = "S" Then
            nS = nS + 1
            lstS(nS) = gSto(g)
        ElseIf gNiv(g) = "I" Then
            nI = nI + 1
            lstI(nI) = gSto(g) & "|" & gIt(g)
        End If
    Next g
    TrierTexte lstS, nS
    TrierTexte lstI, nI

    wm.Range("B9:I12").ClearContents
    iGlob = grp("G")
    wm.Cells(LIG_ML_T1, 2).Value = "Global"
    wm.Cells(LIG_ML_T1, 3).Value = "ALL"
    EcrireTriplet wm, LIG_ML_T1, 4, acc, iGlob, 0
    EcrireTriplet wm, LIG_ML_T1, 7, acc, iGlob, 1
    For i = 1 To nS
        If i > 3 Then Exit For
        lig = LIG_ML_T1 + i
        idx = grp("S|" & lstS(i))
        wm.Cells(lig, 2).Value = "Store"
        wm.Cells(lig, 3).Value = lstS(i)
        EcrireTriplet wm, lig, 4, acc, idx, 0
        EcrireTriplet wm, lig, 7, acc, idx, 1
    Next i

    wm.Range("B16:G65").ClearContents
    wm.Range("I16:J65").ClearContents
    For i = 1 To nI
        If i > MAX_ML_T2 Then Exit For
        lig = LIG_ML_T2 + i - 1
        idx = grp("I|" & lstI(i))
        wm.Cells(lig, 2).Value = gIt(idx)
        wm.Cells(lig, 3).Value = gSto(idx)
        wm.Cells(lig, 4).Value = Metrique(acc, idx, 0, "MAE")
        wm.Cells(lig, 5).Value = Metrique(acc, idx, 0, "RMSE")
        wm.Cells(lig, 6).Value = Metrique(acc, idx, 1, "MAE")
        wm.Cells(lig, 7).Value = Metrique(acc, idx, 1, "RMSE")
        wm.Cells(lig, 9).Value = Metrique(acc, idx, 0, "MAPE")
        wm.Cells(lig, 10).Value = Metrique(acc, idx, 1, "MAPE")
    Next i

    If nP = 0 Then note = "XGBoost (passif) : Modèle non encore entraîné. "
    If nA = 0 Then note = note & "LightGBM (actif) : Modèle non encore entraîné. "
    wm.Range("B6").Value = note & "Dernier import : " & Format(Now, "dd/mm/yyyy hh:nn") & " — " & k & " lignes fusionnées."
    ThisWorkbook.Worksheets(F_ACCUEIL).Range("E14").Value = Now
    Application.Calculation = xlCalculationAutomatic
    Application.Calculate
    Application.ScreenUpdating = True

    rP = Metrique(acc, iGlob, 0, "RMSE")
    rA = Metrique(acc, iGlob, 1, "RMSE")
    If IsEmpty(rA) Then
        meilleur = "XGBoost (passif)"
    ElseIf IsEmpty(rP) Then
        meilleur = "LightGBM (actif)"
    ElseIf rA < rP Then
        meilleur = "LightGBM (actif)"
    Else
        meilleur = "XGBoost (passif)"
    End If
    Info "Import terminé." & vbCrLf & vbCrLf & _
         "Lignes importées : " & nP & " (XGBoost) / " & nA & " (LightGBM)" & vbCrLf & _
         "Séries évaluées : " & nI & vbCrLf & _
         "Meilleur modèle (RMSE global) : " & meilleur
    Exit Sub
Erreur:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Erreur dans ImporterResultatsML : " & Err.Description, vbCritical
End Sub

Private Function LireCSVPrevisions(ByVal fichier As String, ByRef d() As Date, ByRef s() As String, _
        ByRef it() As String, ByRef yv() As Double, ByRef pv() As Double) As Long
    Dim f As Integer, contenu As String, lignes() As String, champs() As String, txt As String
    Dim i As Long, j As Long, n As Long, nL As Long, maxIdx As Long
    Dim iD As Long, iS As Long, iI As Long, iY As Long, iP As Long

    If Len(Dir(fichier)) = 0 Then Exit Function
    f = FreeFile
    Open fichier For Binary Access Read As #f
    contenu = Space$(LOF(f))
    Get #f, , contenu
    Close #f
    lignes = Split(Replace(contenu, vbCr, ""), vbLf)
    nL = UBound(lignes)
    If nL < 1 Then Exit Function

    iD = -1: iS = -1: iI = -1: iY = -1: iP = -1
    champs = Split(lignes(0), ",")
    For j = 0 To UBound(champs)
        Select Case LCase$(Trim$(Replace(champs(j), """", "")))
            Case "date", "ds": iD = j
            Case "store_id": iS = j
            Case "item_id": iI = j
            Case "y_true", "y": iY = j
            Case "y_pred", "yhat": iP = j
        End Select
    Next j
    If iD < 0 Or iS < 0 Or iI < 0 Or iY < 0 Or iP < 0 Then Exit Function
    maxIdx = Application.WorksheetFunction.Max(iD, iS, iI, iY, iP)

    ReDim d(1 To nL): ReDim s(1 To nL): ReDim it(1 To nL): ReDim yv(1 To nL): ReDim pv(1 To nL)
    For i = 1 To nL
        If Len(Trim$(lignes(i))) > 0 Then
            champs = Split(lignes(i), ",")
            If UBound(champs) >= maxIdx Then
                txt = Trim$(champs(iD))
                n = n + 1
                d(n) = DateSerial(CInt(Left$(txt, 4)), CInt(Mid$(txt, 6, 2)), CInt(Mid$(txt, 9, 2)))
                s(n) = Trim$(champs(iS))
                it(n) = Trim$(champs(iI))
                yv(n) = Val(champs(iY))
                pv(n) = Val(champs(iP))
            End If
        End If
    Next i
    If n > 0 Then
        ReDim Preserve d(1 To n): ReDim Preserve s(1 To n): ReDim Preserve it(1 To n)
        ReDim Preserve yv(1 To n): ReDim Preserve pv(1 To n)
    End If
    LireCSVPrevisions = n
End Function

Private Function Metrique(ByRef acc() As Double, ByVal idx As Long, ByVal m As Long, ByVal quoi As String) As Variant
    Dim b As Long
    b = m * 5
    Metrique = Empty
    Select Case quoi
        Case "MAE"
            If acc(idx, b + 1) > 0 Then Metrique = Round(acc(idx, b + 2) / acc(idx, b + 1), 4)
        Case "RMSE"
            If acc(idx, b + 1) > 0 Then Metrique = Round(Sqr(acc(idx, b + 3) / acc(idx, b + 1)), 4)
        Case "MAPE"
            If acc(idx, b + 4) > 0 Then Metrique = Round(acc(idx, b + 5) / acc(idx, b + 4) * 100, 2)
    End Select
End Function

Private Sub EcrireTriplet(ByVal ws As Worksheet, ByVal lig As Long, ByVal col As Long, _
        ByRef acc() As Double, ByVal idx As Long, ByVal m As Long)
    ws.Cells(lig, col).Value = Metrique(acc, idx, m, "MAE")
    ws.Cells(lig, col + 1).Value = Metrique(acc, idx, m, "MAPE")
    ws.Cells(lig, col + 2).Value = Metrique(acc, idx, m, "RMSE")
End Sub

Public Sub ExporterCSV()
    Dim ws As Worksheet, nMMS As Long, nMMP As Long, alpha As Double, saison As Long
    Dim nTest As Long, store As String, produit As String, chemin As String
    Dim dates() As Date, ventes() As Double, n As Long, i As Long, f As Integer

    Set ws = ThisWorkbook.Worksheets(F_PREV)
    If Not LireParametres(ws, nMMS, nMMP, alpha, saison, nTest, store, produit) Then Exit Sub
    If ThisWorkbook.Path = "" Then
        MsgBox "Enregistrez d'abord le classeur.", vbExclamation
        Exit Sub
    End If
    n = ExtraireSerie(store, produit, dates, ventes)
    If n = 0 Then
        MsgBox "Aucune donnée pour " & store & " / " & produit & ".", vbExclamation
        Exit Sub
    End If
    chemin = ThisWorkbook.Path & Application.PathSeparator & "prophet_input.csv"
    f = FreeFile
    Open chemin For Output As #f
    Print #f, "ds,y"
    For i = 1 To n
        Print #f, Format(dates(i), "yyyy-mm-dd") & "," & Replace(CStr(ventes(i)), ",", ".")
    Next i
    Close #f
    MsgBox "Série exportée (" & store & " / " & produit & ", " & n & " lignes) :" & vbCrLf & chemin, _
           vbInformation, "Export CSV pour ML"
End Sub

Public Sub ActualiserTout()
    Dim ancien As Boolean
    ancien = mbSilencieux
    On Error GoTo Erreur
    mbSilencieux = True

    Application.StatusBar = "Actualisation " & Barre(0, 3) & " Réinitialisation des filtres..."
    ReinitialisationFiltre
    Application.StatusBar = "Actualisation " & Barre(1, 3) & " Calcul des prévisions passives..."
    LancerPrevisions
    MsgBox "Progression : " & Barre(2, 3) & vbCrLf & vbCrLf & _
           "1. Filtres réinitialisés" & vbCrLf & "2. Prévisions passives calculées" & vbCrLf & vbCrLf & _
           "Cliquer sur OK pour importer les résultats ML.", vbInformation, "Actualisation en cours"
    Application.StatusBar = "Actualisation " & Barre(2, 3) & " Import des résultats ML..."
    ImporterResultatsML
    ThisWorkbook.Worksheets(F_ACCUEIL).Range("E14").Value = Now

    mbSilencieux = ancien
    Application.StatusBar = False
    MsgBox "Actualisation terminée " & Barre(3, 3) & vbCrLf & vbCrLf & _
           "Filtres, prévisions passives et résultats ML sont à jour.", vbInformation, "Actualiser tout"
    Exit Sub
Erreur:
    mbSilencieux = ancien
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "Erreur pendant l'actualisation : " & Err.Description, vbCritical
End Sub

Public Sub VoirDetailProduit()
    Dim wm As Worksheet, tbl As Variant, store As String, produit As String, i As Long
    Dim msg As String, meilleur As String

    Set wm = ThisWorkbook.Worksheets(F_ML)
    store = Trim$(CStr(wm.Range("C4").Value))
    produit = Trim$(CStr(wm.Range("E4").Value))
    tbl = wm.Range("B16:J65").Value
    For i = 1 To MAX_ML_T2
        If CStr(tbl(i, 1)) = produit And CStr(tbl(i, 2)) = store Then Exit For
    Next i
    If i > MAX_ML_T2 Then
        MsgBox "Aucun résultat ML pour " & store & " / " & produit & "." & vbCrLf & _
               "Lancer ImporterResultatsML ou vérifier les filtres (C4, E4).", vbExclamation
        Exit Sub
    End If
    If EstNombre(tbl(i, 4)) And EstNombre(tbl(i, 6)) Then
        If tbl(i, 6) < tbl(i, 4) Then meilleur = "LightGBM (actif)" Else meilleur = "XGBoost (passif)"
    ElseIf EstNombre(tbl(i, 4)) Then
        meilleur = "XGBoost (passif)"
    ElseIf EstNombre(tbl(i, 6)) Then
        meilleur = "LightGBM (actif)"
    Else
        meilleur = "n/d"
    End If
    msg = "Produit : " & produit & "   Store : " & store & vbCrLf & vbCrLf & _
          "XGBoost (passif)" & vbCrLf & _
          "   MAE = " & Fmt(tbl(i, 3)) & "   MAPE = " & Fmt(tbl(i, 8)) & " %   RMSE = " & Fmt(tbl(i, 4)) & vbCrLf & _
          "LightGBM (actif)" & vbCrLf & _
          "   MAE = " & Fmt(tbl(i, 5)) & "   MAPE = " & Fmt(tbl(i, 9)) & " %   RMSE = " & Fmt(tbl(i, 6)) & vbCrLf & vbCrLf & _
          "Gain RMSE de l'approche active : " & Fmt(tbl(i, 7)) & " %" & vbCrLf & _
          "Meilleur modèle : " & meilleur
    MsgBox msg, vbInformation, "Détail produit"
End Sub


'==============================================================================
' DEMARRAGE ET UTILITAIRES
'==============================================================================
Public Sub DemarrerApplication()
    On Error Resume Next
    InstallerBoutons False
    ThisWorkbook.Worksheets(F_ACCUEIL).Activate
    ThisWorkbook.Worksheets(F_DON).Visible = xlSheetHidden
    ThisWorkbook.Worksheets(F_EVENEMENTS).Visible = xlSheetHidden
    ThisWorkbook.Worksheets("VBA_CODE").Visible = xlSheetHidden
    On Error GoTo 0
End Sub

Public Sub ReinstallerBoutons()
    InstallerBoutons True
End Sub

Private Sub InstallerBoutons(ByVal forcer As Boolean)
    CreerBouton "ACCUEIL", "B8:C10", "AllerDashboard", forcer
    CreerBouton "ACCUEIL", "E8:F10", "AllerPrevision", forcer
    CreerBouton "ACCUEIL", "H8:I10", "AllerML", forcer
    CreerBouton "ACCUEIL", "K8:L10", "AllerGestion", forcer
    CreerBouton "ACCUEIL", "B12:C14", "AllerProduits", forcer
    CreerBouton "ACCUEIL", "E12:F14", "AllerCalendrier", forcer
    CreerBouton "ACCUEIL", "H12:I14", "AllerMagasins", forcer
    CreerBouton "ACCUEIL", "K12:L14", "AllerCategories", forcer
    CreerBouton "ACCUEIL", "E24:I25", "ActualiserTout", forcer
    CreerBouton "PREVISION", "E5:G6", "LancerPrevisions", forcer
    CreerBouton "PREVISION", "E8:G9", "OptimiserAlpha", forcer
    CreerBouton "PREVISION", "E11:G12", "ExporterCSV", forcer
    CreerBouton "ML_RESULTATS", "G4:H5", "ImporterResultatsML", forcer
    CreerBouton "ML_RESULTATS", "J4:K5", "VoirDetailProduit", forcer
    CreerBouton "GESTION_DATA", "E6:F7", "AjouterTransaction", forcer
    CreerBouton "GESTION_DATA", "K6:L7", "AppliquerFiltre", forcer
    CreerBouton "GESTION_DATA", "K9:L10", "ReinitialisationFiltre", forcer
    CreerBouton "GESTION_DATA", "N10:O11", "AjouterProduit", forcer
    CreerBouton "GESTION_DATA", "P10:Q11", "SupprimerProduit", forcer
    CreerBouton "GESTION_DATA", "H13:I13", "Filtrer30", forcer
    CreerBouton "GESTION_DATA", "J13:K13", "Filtrer90", forcer
    CreerBouton "GESTION_DATA", "L13:M13", "Filtrer150", forcer
    CreerBouton "GESTION_DATA", "N13:O13", "FiltrerAnnee", forcer
    CreerBouton "CALENDRIER", "H5:I5", "ApercuCalendrier", forcer
    CreerBouton "CALENDRIER", "H7:I7", "AjouterEvenement", forcer
    CreerBouton "CALENDRIER", "H9:I9", "ReinitialiserCalendrier", forcer
End Sub

Private Sub CreerBouton(ByVal feuille As String, ByVal adresse As String, ByVal macro As String, ByVal forcer As Boolean)
    Dim ws As Worksheet, rng As Range, shp As Shape, nom As String
    Set ws = ThisWorkbook.Worksheets(feuille)
    Set rng = ws.Range(adresse)
    nom = "btn_" & macro & "_" & Replace(adresse, ":", "_")
    On Error Resume Next
    Set shp = ws.Shapes(nom)
    On Error GoTo 0
    If Not shp Is Nothing Then
        If Not forcer Then Exit Sub
        shp.Delete
    End If
    Set shp = ws.Shapes.AddShape(5, rng.Left + 1, rng.Top + 1, rng.Width - 2, rng.Height - 2)
    With shp
        .Name = nom
        .OnAction = macro
        .Placement = xlMoveAndSize
        .Fill.ForeColor.RGB = rng.Cells(1, 1).Interior.Color
        .Line.Visible = 0
        With .TextFrame2
            .VerticalAnchor = 3
            .TextRange.Text = CStr(rng.Cells(1, 1).Value)
            .TextRange.Font.Bold = -1
            .TextRange.Font.Size = 11
            .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
            .TextRange.ParagraphFormat.Alignment = 2
        End With
    End With
End Sub

Private Sub Info(ByVal texte As String)
    If Not mbSilencieux Then MsgBox texte, vbInformation, "Prévision de la demande"
End Sub

Private Function Barre(ByVal k As Long, ByVal n As Long) As String
    Dim plein As Long
    plein = (k * 10) \ n
    Barre = "[" & String(plein, "#") & String(10 - plein, "-") & "] " & k & "/" & n
End Function

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

Private Sub TrierTexte(ByRef t() As String, ByVal n As Long)
    Dim i As Long, j As Long, tmp As String
    For i = 2 To n
        tmp = t(i)
        j = i - 1
        Do While j >= 1
            If t(j) <= tmp Then Exit Do
            t(j + 1) = t(j)
            j = j - 1
        Loop
        t(j + 1) = tmp
    Next i
End Sub
