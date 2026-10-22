Attribute VB_Name = "ModuleNavigation"
Option Explicit

' ====================================================================
'  NAVIGATION (macros assignees aux boutons de l'Accueil et de Gestion)
' ====================================================================

Public Sub GoHome()
    ThisWorkbook.Worksheets(HOME_SHEET).Activate
End Sub

Public Function ShowLoginForm() As Boolean
    frmLogin.Show
    ShowLoginForm = IsLoggedIn()
End Function

' Bouton "Se connecter"
Public Sub OpenLogin()
    If IsLoggedIn() Then
        If MsgBox("Déjà connecté (" & CurrentUserName() & "). Changer d'utilisateur ?", _
                  vbQuestion + vbYesNo) <> vbYes Then Exit Sub
        EndSession
    End If
    ShowLoginForm
End Sub

' Bouton "Se deconnecter"
Public Sub Logout()
    DoLogout
End Sub

' Ouvre une feuille ; si personne n'est connecte, affiche d'abord le formulaire de connexion
Public Sub GoToSheet(ByVal sheetName As String)
    If Not IsLoggedIn() Then
        MsgBox "Veuillez vous connecter pour accéder à cette section.", vbInformation, "Connexion requise"
        If Not ShowLoginForm() Then Exit Sub
    End If
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    If ws Is Nothing Then
        MsgBox "Feuille introuvable : " & sheetName, vbExclamation
    ElseIf ws.Visible <> xlSheetVisible Then
        MsgBox "Cette feuille n'est pas accessible avec votre profil.", vbExclamation
    Else
        ws.Activate
    End If
    On Error GoTo 0
End Sub

Public Sub OpenPowerBI():   GoToSheet "PowerBI":   End Sub
Public Sub OpenML():        GoToSheet "ML":        End Sub
Public Sub OpenPrevision(): GoToSheet "Prevision": End Sub
Public Sub OpenGestion():   GoToSheet "Gestion":   End Sub

' Formulaires de saisie (administrateur et HQ uniquement)
Public Sub OpenProduits():     OpenCrudForm "frmProduits":     End Sub
Public Sub OpenCategories():   OpenCrudForm "frmCategories":   End Sub
Public Sub OpenMagasins():     OpenCrudForm "frmMagasins":     End Sub
Public Sub OpenTransactions(): OpenCrudForm "frmTransactions": End Sub
Public Sub OpenPrix():         OpenCrudForm "frmPrix":         End Sub
Public Sub OpenCalendrier():   OpenCrudForm "frmCalendrier":   End Sub

Private Sub OpenCrudForm(ByVal formName As String)
    Dim r As String: r = UserRole()
    If r <> ROLE_ADMIN And r <> ROLE_HQ Then
        MsgBox "Accès en écriture réservé à l'administrateur et au HQ.", vbExclamation
        Exit Sub
    End If
    Dim f As Object
    Set f = VBA.UserForms.Add(formName)
    f.Show
    Set f = Nothing
End Sub
