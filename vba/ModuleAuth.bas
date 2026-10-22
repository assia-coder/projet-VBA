Attribute VB_Name = "ModuleAuth"
Option Explicit

' ====================================================================
'  AUTHENTIFICATION / SESSION
'  La session est stockee dans la feuille cachee R_Session (B2:B5).
'  Les KPI de l'Accueil sont des formules lisant ces cellules :
'  ils sont donc toujours coherents avec l'utilisateur connecte.
' ====================================================================

Public Const ROLE_ADMIN As String = "admin"
Public Const ROLE_HQ As String = "hq"
Public Const ROLE_STORE As String = "store"

Public Function IsLoggedIn() As Boolean
    IsLoggedIn = (UserRole() <> "")
End Function

Private Function SessionValue(ByVal addr As String) As String
    Dim v As Variant
    On Error Resume Next
    v = ThisWorkbook.Worksheets(SESSION_SHEET).Range(addr).Value
    SessionValue = Trim$(CStr(v))
    On Error GoTo 0
End Function

Public Function UserRole() As String
    UserRole = SessionValue("B2")
End Function

Public Function UserStore() As String
    UserStore = SessionValue("B3")
End Function

Public Function CurrentUserName() As String
    CurrentUserName = SessionValue("B4")
End Function

' Verifie les identifiants. Renvoie le role et le magasin (ALL = tous).
Public Function Authenticate(ByVal usr As String, ByVal pwd As String, _
                             ByRef role As String, ByRef store As String) As Boolean
    role = "": store = ""
    Select Case LCase$(Trim$(usr))
        Case "admin":      If pwd = "admin2024" Then role = ROLE_ADMIN: store = "ALL"
        Case "hq_analyst": If pwd = "hq2024" Then role = ROLE_HQ: store = "ALL"
        Case "tx_manager": If pwd = "tx2024" Then role = ROLE_STORE: store = "TX_1"
        Case "ca_manager": If pwd = "ca2024" Then role = ROLE_STORE: store = "CA_1"
        Case "wi_manager": If pwd = "wi2024" Then role = ROLE_STORE: store = "WI_1"
    End Select
    Authenticate = (role <> "")
End Function

Private Sub WriteSessionCells(ByVal role As String, ByVal store As String, ByVal usr As String)
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(SESSION_SHEET)
    On Error Resume Next
    ws.Unprotect PASSWORD
    On Error GoTo 0
    ws.Range("B2").Value = role
    ws.Range("B3").Value = store
    ws.Range("B4").Value = usr
    ws.Range("B5").Value = Format$(Now, "yyyy-mm-dd hh:nn:ss")
End Sub

Public Sub ClearSessionCells()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(SESSION_SHEET)
    On Error Resume Next
    ws.Unprotect PASSWORD
    On Error GoTo 0
    ws.Range("B2:B5").ClearContents
End Sub

' Ouvre la session : feuilles visibles + filtres du role + protection
Public Sub SetSession(ByVal role As String, ByVal store As String, ByVal usr As String)
    Dim wasSaved As Boolean: wasSaved = ThisWorkbook.Saved
    Dim prevEv As Boolean: prevEv = Application.EnableEvents
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    WriteSessionCells role, store, usr
    ApplySheetVisibility
    ApplyRoleFilters
    ProtectAllSheets
    Application.Calculate
    Application.ScreenUpdating = True
    Application.EnableEvents = prevEv
    ThisWorkbook.Saved = wasSaved
End Sub

Public Sub EndSession()
    LockWorkbook
End Sub

Public Sub DoLogout()
    If Not IsLoggedIn() Then
        MsgBox "Aucune session active.", vbInformation
        Exit Sub
    End If
    If MsgBox("Se déconnecter ?", vbQuestion + vbYesNo) <> vbYes Then Exit Sub
    EndSession
    MsgBox "Déconnecté.", vbInformation
End Sub
