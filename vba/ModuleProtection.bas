Attribute VB_Name = "ModuleProtection"
Option Explicit

' ====================================================================
'  PROTECTION / VISIBILITE DES FEUILLES
'  - Avant connexion : seule "Accueil" est visible (les autres sont VeryHidden)
'  - hq / store      : feuilles de l'application visibles, protegees (lecture)
'  - admin           : toutes les feuilles visibles, NON protegees (styles libres)
'  - A l'enregistrement et a la fermeture : le classeur est toujours
'    rendu a l'etat verrouille (session videe, feuilles masquees)
' ====================================================================

Public Const PASSWORD As String = "walmart2024"
Public Const SESSION_SHEET As String = "R_Session"
Public Const HOME_SHEET As String = "Accueil"

Private mRestoreRole As String
Private mRestoreStore As String
Private mRestoreUser As String
Private mRestoreSheet As String

' Type de feuille : home / session / admin / user
Public Function SheetKind(ByVal nm As String) As String
    Select Case nm
        Case "Accueil"
            SheetKind = "home"
        Case "R_Session"
            SheetKind = "session"
        Case "Config", "R_Previsions", "R_Metriques", "R_Resume", "R_Mensuel", "R_MetMensuel"
            SheetKind = "admin"
        Case Else
            SheetKind = "user"
    End Select
End Function

Public Function MustBeVisible(ByVal nm As String, ByVal role As String) As Boolean
    Select Case SheetKind(nm)
        Case "home"
            MustBeVisible = True
        Case "session"
            MustBeVisible = False
        Case "admin"
            MustBeVisible = (role = ROLE_ADMIN)
        Case Else
            MustBeVisible = (role <> "")
    End Select
End Function

' Affiche / masque les feuilles selon le role courant
Public Sub ApplySheetVisibility()
    Dim ws As Worksheet, role As String
    role = UserRole()
    On Error Resume Next
    ThisWorkbook.Unprotect PASSWORD
    On Error GoTo 0
    ' 1) afficher d'abord (une feuille visible doit toujours exister)
    For Each ws In ThisWorkbook.Worksheets
        If MustBeVisible(ws.Name, role) Then
            If ws.Visible <> xlSheetVisible Then ws.Visible = xlSheetVisible
        End If
    Next ws
    ' 2) masquer le reste
    For Each ws In ThisWorkbook.Worksheets
        If Not MustBeVisible(ws.Name, role) Then
            If ws.Visible <> xlSheetVeryHidden Then ws.Visible = xlSheetVeryHidden
        End If
    Next ws
    ' 3) structure verrouillee sauf pour l'admin
    If role <> ROLE_ADMIN Then
        On Error Resume Next
        ThisWorkbook.Protect Password:=PASSWORD, Structure:=True, Windows:=False
        On Error GoTo 0
    End If
End Sub

' Protege une feuille selon le role (admin = aucune protection)
Public Sub ProtectSheetForRole(ByVal ws As Worksheet)
    On Error Resume Next
    ws.Unprotect PASSWORD
    If UserRole() <> ROLE_ADMIN Then
        ws.Protect Password:=PASSWORD, DrawingObjects:=True, Contents:=True, Scenarios:=True, _
                   UserInterfaceOnly:=True, AllowFormattingColumns:=True, AllowFormattingRows:=True
    End If
    On Error GoTo 0
End Sub

Public Sub ProtectAllSheets()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        ProtectSheetForRole ws
    Next ws
End Sub

' Masque une feuille en urgence (ex. filtre de role impossible a appliquer)
Public Sub HideSheetSafe(ByVal ws As Worksheet)
    On Error Resume Next
    ThisWorkbook.Unprotect PASSWORD
    ThisWorkbook.Worksheets(HOME_SHEET).Activate
    ws.Visible = xlSheetVeryHidden
    If UserRole() <> ROLE_ADMIN Then ThisWorkbook.Protect Password:=PASSWORD, Structure:=True, Windows:=False
    On Error GoTo 0
End Sub

' Retour a l'etat verrouille : session videe, filtres retires, feuilles masquees
Public Sub LockWorkbook()
    Dim wasSaved As Boolean: wasSaved = ThisWorkbook.Saved
    Dim prevEv As Boolean: prevEv = Application.EnableEvents
    Dim prevSu As Boolean: prevSu = Application.ScreenUpdating
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    ClearSessionCells
    ResetAllFilters
    ApplySheetVisibility
    ProtectAllSheets
    On Error Resume Next
    ThisWorkbook.Worksheets(HOME_SHEET).Activate
    ThisWorkbook.Worksheets(HOME_SHEET).Range("A1").Select
    On Error GoTo 0
    Application.Calculate
    Application.ScreenUpdating = prevSu
    Application.EnableEvents = prevEv
    ThisWorkbook.Saved = wasSaved
End Sub

' --- Enregistrement : toujours en etat verrouille, puis restauration de la session
Public Sub PrepareForSave()
    If Not IsLoggedIn() Then Exit Sub
    mRestoreRole = UserRole()
    mRestoreStore = UserStore()
    mRestoreUser = CurrentUserName()
    On Error Resume Next
    mRestoreSheet = ActiveSheet.Name
    On Error GoTo 0
    LockWorkbook
    Application.OnTime Now, "'" & ThisWorkbook.Name & "'!RestoreAfterSave"
End Sub

Public Sub RestoreAfterSave()
    If mRestoreRole = "" Then Exit Sub
    Dim r As String, s As String, u As String, sh As String
    r = mRestoreRole: s = mRestoreStore: u = mRestoreUser: sh = mRestoreSheet
    mRestoreRole = ""
    SetSession r, s, u
    On Error Resume Next
    If sh <> "" Then ThisWorkbook.Worksheets(sh).Activate
    On Error GoTo 0
End Sub

' --- Fermeture
Public Sub HandleBeforeClose(ByRef Cancel As Boolean)
    Dim dirty As Boolean: dirty = Not ThisWorkbook.Saved
    Dim role As String: role = UserRole()
    Dim doSave As Boolean
    If role = "" Then Exit Sub          ' non connecte : Excel gere l'invite habituelle
    If dirty And (role = ROLE_ADMIN Or role = ROLE_HQ) Then
        Select Case MsgBox("Enregistrer les modifications avant de fermer ?", _
                           vbYesNoCancel + vbQuestion, "Fermeture")
            Case vbCancel
                Cancel = True
                Exit Sub
            Case vbYes
                doSave = True
        End Select
    End If
    LockWorkbook
    If doSave Then
        ThisWorkbook.Save
    Else
        ThisWorkbook.Saved = True
    End If
End Sub
