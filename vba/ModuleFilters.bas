Attribute VB_Name = "ModuleFilters"
Option Explicit

' ====================================================================
'  FILTRAGE PAR ROLE (seul mecanisme d'acces aux lignes)
'  - admin / hq : tout afficher
'  - store      : lignes de son magasin (T_Transactions, T_Produits, T_Prix,
'                 T_Magasins) ; categories ayant au moins 1 produit dans le magasin
'  Le filtre est pose sur le tableau (ListObject) avec les fleches de filtre
'  MASQUEES : l'utilisateur ne peut pas le modifier (feuille protegee).
' ====================================================================

Public Function TableSheets() As Variant
    TableSheets = Array("T_Transactions", "T_Produits", "T_Prix", _
                        "T_Calendrier", "T_Categories", "T_Magasins")
End Function

Public Function IsTableSheet(ByVal nm As String) As Boolean
    Dim v As Variant
    For Each v In TableSheets()
        If CStr(v) = nm Then IsTableSheet = True: Exit Function
    Next v
End Function

Public Function FieldIndex(ByVal lo As ListObject, ByVal header As String) As Long
    Dim i As Long
    For i = 1 To lo.ListColumns.Count
        If LCase$(lo.ListColumns(i).Name) = LCase$(header) Then FieldIndex = i: Exit Function
    Next i
End Function

Public Sub ApplyRoleFilters()
    Dim v As Variant
    For Each v In TableSheets()
        ApplyRoleFilterToSheet CStr(v)
    Next v
End Sub

' Appele a la connexion et a chaque activation de feuille (Workbook_SheetActivate)
Public Sub ApplyRoleFilterToSheet(ByVal sheetName As String)
    If Not IsTableSheet(sheetName) Then Exit Sub
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    If ws.ListObjects.Count = 0 Then Exit Sub

    Dim role As String, store As String, fld As String, crit As String
    role = UserRole(): store = UserStore()
    If role = ROLE_STORE And store <> "" Then
        Select Case sheetName
            Case "T_Transactions", "T_Produits", "T_Prix", "T_Magasins"
                fld = "store_id": crit = "=" & store
            Case "T_Categories"
                fld = "nb_produits": crit = ">0"
        End Select
    End If

    If fld = "" Then
        If IsTableFiltered(ws) Then ClearTableFilter ws
    Else
        If Not IsTableFiltered(ws) Then ApplyTableFilter ws, fld, crit
    End If
End Sub

Private Function IsTableFiltered(ByVal ws As Worksheet) As Boolean
    Dim lo As ListObject
    On Error Resume Next
    Set lo = ws.ListObjects(1)
    If Not lo Is Nothing Then
        If lo.ShowAutoFilter Then IsTableFiltered = lo.AutoFilter.FilterMode
    End If
    On Error GoTo 0
End Function

Private Sub ApplyTableFilter(ByVal ws As Worksheet, ByVal fld As String, ByVal crit As String)
    Dim lo As ListObject, target As Long, i As Long
    On Error GoTo EH
    Set lo = ws.ListObjects(1)
    target = FieldIndex(lo, fld)
    If target = 0 Then Err.Raise vbObjectError + 2001, , "Colonne introuvable : " & fld
    ws.Unprotect PASSWORD
    If lo.ShowAutoFilter Then
        If lo.AutoFilter.FilterMode Then lo.AutoFilter.ShowAllData
    End If
    ' fleches masquees sur toutes les colonnes (aucun filtre manuel possible)
    For i = 1 To lo.ListColumns.Count
        If i <> target Then lo.Range.AutoFilter Field:=i, VisibleDropDown:=False
    Next i
    lo.Range.AutoFilter Field:=target, Criteria1:=crit, VisibleDropDown:=False
    ProtectSheetForRole ws
    Exit Sub
EH:
    ' securite : si le filtre de role n'a pas pu etre pose, la feuille est masquee
    MsgBox "Le filtre de la feuille " & ws.Name & " n'a pas pu être appliqué (" & Err.Description & _
           "). La feuille est masquée par sécurité.", vbExclamation
    HideSheetSafe ws
End Sub

Public Sub ClearTableFilter(ByVal ws As Worksheet)
    Dim lo As ListObject
    On Error Resume Next
    ws.Unprotect PASSWORD
    Set lo = ws.ListObjects(1)
    If Not lo Is Nothing Then
        If lo.ShowAutoFilter Then
            If lo.AutoFilter.FilterMode Then lo.AutoFilter.ShowAllData
            lo.ShowAutoFilter = False
        End If
    End If
    ProtectSheetForRole ws
    On Error GoTo 0
End Sub

Public Sub ResetAllFilters()
    Dim v As Variant, ws As Worksheet
    For Each v In TableSheets()
        Set ws = Nothing
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(CStr(v))
        On Error GoTo 0
        If Not ws Is Nothing Then ClearTableFilter ws
    Next v
End Sub
