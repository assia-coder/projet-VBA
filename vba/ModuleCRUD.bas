Attribute VB_Name = "ModuleCRUD"
Option Explicit

' ====================================================================
'  CRUD GENERIQUE utilise par les UserForms (frmProduits, frmPrix, ...)
'  - Travaille sur le tableau (ListObject) de chaque feuille T_*
'  - Colonnes retrouvees par NOM d'en-tete (jamais par position)
'  - Saisies marquees source = MANUEL ; lignes M5 de transactions /
'    prix / calendrier en lecture seule
'  - Les formulaires ne listent que MAX_LIST_ROWS lignes (recherche par cle)
' ====================================================================

Public Const MAX_LIST_ROWS As Long = 500

' ------------------------------------------------------------ utilitaires
Public Function GetLO(ByVal sheetName As String) As ListObject
    Set GetLO = ThisWorkbook.Worksheets(sheetName).ListObjects(1)
End Function

Private Sub BeginEdit(ByVal ws As Worksheet)
    On Error Resume Next
    ws.Unprotect PASSWORD
    On Error GoTo 0
End Sub

Private Sub EndEdit(ByVal ws As Worksheet)
    ProtectSheetForRole ws
End Sub

Private Function CellText(ByVal v As Variant) As String
    If IsError(v) Then Exit Function
    Select Case VarType(v)
        Case vbDate
            CellText = Format$(v, "yyyy-mm-dd")
        Case vbBoolean
            If v Then CellText = "TRUE" Else CellText = "FALSE"
        Case vbEmpty, vbNull
            CellText = ""
        Case Else
            CellText = CStr(v)
    End Select
End Function

Private Function CellAt(ByVal lo As ListObject, ByVal tblRow As Long, ByVal header As String) As String
    Dim c As Long: c = FieldIndex(lo, header)
    If c > 0 Then CellAt = CellText(lo.DataBodyRange.Cells(tblRow, c).Value)
End Function

Private Sub SetByHeader(ByVal lo As ListObject, ByVal tblRow As Long, ByVal header As String, ByVal v As Variant)
    Dim c As Long: c = FieldIndex(lo, header)
    If c > 0 Then lo.DataBodyRange.Cells(tblRow, c).Value = v
End Sub

' Convertit le texte saisi : date, booleen, nombre, sinon texte ; vide -> Empty
Private Function Coerce(ByVal header As String, ByVal txt As String) As Variant
    Dim t As String: t = Trim$(txt)
    If t = "" Then
        If LCase$(header) = "actif" Then Coerce = True Else Coerce = Empty
        Exit Function
    End If
    Select Case UCase$(t)
        Case "TRUE", "VRAI", "OUI"
            Coerce = True: Exit Function
        Case "FALSE", "FAUX", "NON"
            Coerce = False: Exit Function
    End Select
    If LCase$(header) = "date" Then
        If IsDate(t) Then
            Coerce = CDate(t)
            Exit Function
        Else
            Err.Raise vbObjectError + 1001, , "Date invalide : " & t & " (format attendu : aaaa-mm-jj)"
        End If
    End If
    Dim ds As String, alt As String, t2 As String
    ds = Application.International(xlDecimalSeparator)
    If ds = "," Then alt = "." Else alt = ","
    t2 = Replace(t, alt, ds)
    If IsNumeric(t2) Then
        Coerce = CDbl(t2)
    Else
        Coerce = t
    End If
End Function

' Valeurs d'une colonne en tableau 2D 1-base (meme avec une seule ligne)
Private Function ColValues(ByVal lo As ListObject, ByVal idx As Long) As Variant
    Dim rng As Range: Set rng = lo.ListColumns(idx).DataBodyRange
    If rng.Cells.Count = 1 Then
        Dim a(1 To 1, 1 To 1) As Variant
        a(1, 1) = rng.Value
        ColValues = a
    Else
        ColValues = rng.Value
    End If
End Function

' Colonnes cles en texte : tableau (1..n, 1..nbCles)
Private Function LoadKeyArray(ByVal lo As ListObject, ByVal keyHdrs As Variant) As String()
    Dim nK As Long, n As Long, k As Long, r As Long, c As Long
    nK = UBound(keyHdrs) - LBound(keyHdrs) + 1
    n = lo.ListRows.Count
    Dim res() As String
    ReDim res(1 To n, 1 To nK)
    Dim tmp As Variant
    For k = 1 To nK
        c = FieldIndex(lo, CStr(keyHdrs(LBound(keyHdrs) + k - 1)))
        tmp = ColValues(lo, c)
        For r = 1 To n
            res(r, k) = CellText(tmp(r, 1))
        Next r
    Next k
    LoadKeyArray = res
End Function

' ------------------------------------------------------------ liste du formulaire
' Remplit la ListBox (cles + n° de ligne cache). Renvoie le nombre de lignes affichees.
Public Function FillRecordList(ByVal lst As Object, ByVal sheetName As String, _
                               ByVal keyHdrs As Variant, ByVal filterTxt As String) As Long
    Dim lo As ListObject: Set lo = GetLO(sheetName)
    Dim nK As Long: nK = UBound(keyHdrs) - LBound(keyHdrs) + 1
    Dim w As String, k As Long, r As Long
    lst.Clear
    lst.ColumnCount = nK + 1
    For k = 1 To nK: w = w & "85 pt;": Next k
    lst.ColumnWidths = w & "0 pt"
    If lo.DataBodyRange Is Nothing Then Exit Function

    Dim keys() As String: keys = LoadKeyArray(lo, keyHdrs)
    Dim ft As String: ft = LCase$(Trim$(filterTxt))
    Dim n As Long: n = lo.ListRows.Count
    Dim cnt As Long, ok As Boolean, s As String
    For r = 1 To n
        ok = (ft = "")
        If Not ok Then
            s = ""
            For k = 1 To nK: s = s & "|" & keys(r, k): Next k
            ok = (InStr(1, s, ft, vbTextCompare) > 0)
        End If
        If ok Then
            lst.AddItem keys(r, 1)
            For k = 2 To nK
                lst.List(cnt, k - 1) = keys(r, k)
            Next k
            lst.List(cnt, nK) = CStr(r)
            cnt = cnt + 1
            If cnt >= MAX_LIST_ROWS Then Exit For
        End If
    Next r
    FillRecordList = cnt
End Function

Public Sub LoadRecord(ByVal frm As Object, ByVal sheetName As String, _
                      ByVal tblRow As Long, ByVal fields As Variant)
    Dim lo As ListObject: Set lo = GetLO(sheetName)
    Dim i As Long, c As Long
    For i = LBound(fields) To UBound(fields)
        c = FieldIndex(lo, CStr(fields(i)))
        If c > 0 Then
            frm.Controls("txt" & CStr(i - LBound(fields))).Value = _
                CellText(lo.DataBodyRange.Cells(tblRow, c).Value)
        End If
    Next i
End Sub

Public Sub ClearFields(ByVal frm As Object, ByVal n As Long)
    Dim i As Long
    For i = 0 To n - 1
        frm.Controls("txt" & CStr(i)).Value = ""
    Next i
End Sub

' ------------------------------------------------------------ regles
Private Function KeysFilled(ByVal fields As Variant, ByVal vals As Variant, ByVal keyHdrs As Variant) As Boolean
    Dim k As Long, i As Long
    For k = LBound(keyHdrs) To UBound(keyHdrs)
        For i = LBound(fields) To UBound(fields)
            If LCase$(CStr(fields(i))) = LCase$(CStr(keyHdrs(k))) Then
                If IsEmpty(vals(i - LBound(fields))) Then Exit Function
            End If
        Next i
    Next k
    KeysFilled = True
End Function

' N° de ligne (dans le tableau) ayant la meme cle, hors ignoreRow ; 0 si aucune
Private Function FindRowByKey(ByVal lo As ListObject, ByVal fields As Variant, ByVal vals As Variant, _
                              ByVal keyHdrs As Variant, ByVal ignoreRow As Long) As Long
    If lo.DataBodyRange Is Nothing Then Exit Function
    Dim nK As Long: nK = UBound(keyHdrs) - LBound(keyHdrs) + 1
    Dim want() As String: ReDim want(1 To nK)
    Dim k As Long, i As Long, r As Long, match As Boolean
    For k = 1 To nK
        For i = LBound(fields) To UBound(fields)
            If LCase$(CStr(fields(i))) = LCase$(CStr(keyHdrs(LBound(keyHdrs) + k - 1))) Then
                want(k) = CellText(vals(i - LBound(fields)))
                Exit For
            End If
        Next i
    Next k
    Dim keys() As String: keys = LoadKeyArray(lo, keyHdrs)
    For r = 1 To lo.ListRows.Count
        If r <> ignoreRow Then
            match = True
            For k = 1 To nK
                If keys(r, k) <> want(k) Then match = False: Exit For
            Next k
            If match Then FindRowByKey = r: Exit Function
        End If
    Next r
End Function

Private Function IsM5Locked(ByVal lo As ListObject, ByVal sheetName As String, ByVal tblRow As Long) As Boolean
    Select Case sheetName
        Case "T_Transactions", "T_Prix", "T_Calendrier"
            IsM5Locked = (UCase$(CellAt(lo, tblRow, "source")) = "M5")
    End Select
End Function

Private Function CountIn(ByVal sheetName As String, ByVal h1 As String, ByVal v1 As String, _
                         Optional ByVal h2 As String = "", Optional ByVal v2 As String = "") As Long
    Dim lo As ListObject: Set lo = GetLO(sheetName)
    If lo.DataBodyRange Is Nothing Then Exit Function
    If h2 = "" Then
        CountIn = CLng(Application.WorksheetFunction.CountIf( _
            lo.ListColumns(FieldIndex(lo, h1)).DataBodyRange, v1))
    Else
        CountIn = CLng(Application.WorksheetFunction.CountIfs( _
            lo.ListColumns(FieldIndex(lo, h1)).DataBodyRange, v1, _
            lo.ListColumns(FieldIndex(lo, h2)).DataBodyRange, v2))
    End If
End Function

' ------------------------------------------------------------ ecriture
Private Sub WriteRow(ByVal lo As ListObject, ByVal tblRow As Long, ByVal fields As Variant, ByVal vals As Variant)
    Dim i As Long, c As Long, cel As Range
    For i = 0 To UBound(vals)
        c = FieldIndex(lo, CStr(fields(LBound(fields) + i)))
        If c > 0 Then
            Set cel = lo.DataBodyRange.Cells(tblRow, c)
            If VarType(vals(i)) = vbDate Then cel.NumberFormat = "yyyy-mm-dd"
            If IsEmpty(vals(i)) Then cel.ClearContents Else cel.Value = vals(i)
        End If
    Next i
End Sub

' Recopie les formules (ex. nb_produits) de la 1re ligne vers la nouvelle ligne
Private Sub CopyFormulaColumns(ByVal lo As ListObject, ByVal tblRow As Long)
    If tblRow <= 1 Then Exit Sub
    Dim c As Long
    For c = 1 To lo.ListColumns.Count
        If lo.DataBodyRange.Cells(1, c).HasFormula Then
            lo.DataBodyRange.Cells(tblRow, c).FormulaR1C1 = lo.DataBodyRange.Cells(1, c).FormulaR1C1
        End If
    Next c
End Sub

Private Function ReadVals(ByVal frm As Object, ByVal fields As Variant) As Variant
    Dim n As Long: n = UBound(fields) - LBound(fields) + 1
    Dim vals() As Variant: ReDim vals(0 To n - 1)
    Dim i As Long
    For i = 0 To n - 1
        vals(i) = Coerce(CStr(fields(LBound(fields) + i)), CStr(frm.Controls("txt" & CStr(i)).Value))
    Next i
    ReadVals = vals
End Function

Public Function AddRecord(ByVal frm As Object, ByVal sheetName As String, _
                          ByVal fields As Variant, ByVal keyHdrs As Variant) As Boolean
    Dim ws As Worksheet
    On Error GoTo EH
    Dim lo As ListObject: Set lo = GetLO(sheetName)
    Set ws = lo.Parent
    Dim vals As Variant: vals = ReadVals(frm, fields)
    If Not KeysFilled(fields, vals, keyHdrs) Then
        MsgBox "Renseignez tous les champs clés (" & Join(keyHdrs, ", ") & ").", vbExclamation
        Exit Function
    End If
    If FindRowByKey(lo, fields, vals, keyHdrs, 0) > 0 Then
        MsgBox "Un enregistrement avec la même clé existe déjà.", vbExclamation
        Exit Function
    End If
    BeginEdit ws
    Dim lr As ListRow: Set lr = lo.ListRows.Add
    Dim idx As Long: idx = lr.Index
    WriteRow lo, idx, fields, vals
    SetByHeader lo, idx, "source", "MANUEL"
    SetByHeader lo, idx, "saisi_le", Format$(Now, "yyyy-mm-dd hh:nn:ss")
    If FieldIndex(lo, "date_ajout") > 0 Then
        lo.DataBodyRange.Cells(idx, FieldIndex(lo, "date_ajout")).NumberFormat = "yyyy-mm-dd"
        SetByHeader lo, idx, "date_ajout", Date
    End If
    CopyFormulaColumns lo, idx
    EndEdit ws
    Application.Calculate
    AddRecord = True
    MsgBox "Enregistrement ajouté (ligne " & idx & " du tableau).", vbInformation
    Exit Function
EH:
    Dim msg As String: msg = Err.Description
    On Error Resume Next
    If Not ws Is Nothing Then EndEdit ws
    On Error GoTo 0
    MsgBox "Ajout impossible : " & msg, vbExclamation
End Function

Public Function UpdateRecord(ByVal frm As Object, ByVal sheetName As String, ByVal fields As Variant, _
                             ByVal keyHdrs As Variant, ByVal tblRow As Long) As Boolean
    Dim ws As Worksheet
    On Error GoTo EH
    Dim lo As ListObject: Set lo = GetLO(sheetName)
    Set ws = lo.Parent
    If tblRow < 1 Or tblRow > lo.ListRows.Count Then
        MsgBox "Sélectionnez d'abord une ligne dans la liste.", vbExclamation
        Exit Function
    End If
    If IsM5Locked(lo, sheetName, tblRow) Then
        MsgBox "Les lignes M5 de cette table sont protégées (lecture seule).", vbExclamation
        Exit Function
    End If
    Dim vals As Variant: vals = ReadVals(frm, fields)
    If Not KeysFilled(fields, vals, keyHdrs) Then
        MsgBox "Renseignez tous les champs clés (" & Join(keyHdrs, ", ") & ").", vbExclamation
        Exit Function
    End If
    If FindRowByKey(lo, fields, vals, keyHdrs, tblRow) > 0 Then
        MsgBox "Un autre enregistrement possède déjà cette clé.", vbExclamation
        Exit Function
    End If
    BeginEdit ws
    WriteRow lo, tblRow, fields, vals
    EndEdit ws
    Application.Calculate
    UpdateRecord = True
    MsgBox "Enregistrement mis à jour.", vbInformation
    Exit Function
EH:
    Dim msg As String: msg = Err.Description
    On Error Resume Next
    If Not ws Is Nothing Then EndEdit ws
    On Error GoTo 0
    MsgBox "Modification impossible : " & msg, vbExclamation
End Function

Public Function DeleteRecord(ByVal sheetName As String, ByVal tblRow As Long) As Boolean
    Dim ws As Worksheet
    On Error GoTo EH
    Dim lo As ListObject: Set lo = GetLO(sheetName)
    Set ws = lo.Parent
    If tblRow < 1 Or tblRow > lo.ListRows.Count Then
        MsgBox "Sélectionnez d'abord une ligne dans la liste.", vbExclamation
        Exit Function
    End If
    If IsM5Locked(lo, sheetName, tblRow) Then
        MsgBox "Les lignes M5 de cette table sont protégées (lecture seule).", vbExclamation
        Exit Function
    End If
    Dim n As Long
    Select Case sheetName
        Case "T_Produits"
            n = CountIn("T_Transactions", "store_id", CellAt(lo, tblRow, "store_id"), _
                        "item_id", CellAt(lo, tblRow, "item_id"))
            If n > 0 Then
                If MsgBox("Ce produit possède " & n & " transaction(s) : il sera désactivé " & _
                          "(actif = FALSE) au lieu d'être supprimé. Continuer ?", _
                          vbQuestion + vbYesNo) <> vbYes Then Exit Function
                BeginEdit ws
                SetByHeader lo, tblRow, "actif", False
                EndEdit ws
                Application.Calculate
                DeleteRecord = True
                MsgBox "Produit désactivé.", vbInformation
                Exit Function
            End If
        Case "T_Categories"
            n = CountIn("T_Produits", "cat_id", CellAt(lo, tblRow, "cat_id"))
            If n > 0 Then
                MsgBox "Suppression impossible : " & n & " produit(s) utilisent cette catégorie.", vbExclamation
                Exit Function
            End If
        Case "T_Magasins"
            n = CountIn("T_Produits", "store_id", CellAt(lo, tblRow, "store_id"))
            If n > 0 Then
                MsgBox "Suppression impossible : " & n & " produit(s) sont rattachés à ce magasin.", vbExclamation
                Exit Function
            End If
    End Select
    If MsgBox("Supprimer définitivement cette ligne ?", vbQuestion + vbYesNo) <> vbYes Then Exit Function
    BeginEdit ws
    lo.ListRows(tblRow).Delete
    EndEdit ws
    Application.Calculate
    DeleteRecord = True
    MsgBox "Enregistrement supprimé.", vbInformation
    Exit Function
EH:
    Dim msg As String: msg = Err.Description
    On Error Resume Next
    If Not ws Is Nothing Then EndEdit ws
    On Error GoTo 0
    MsgBox "Suppression impossible : " & msg, vbExclamation
End Function
