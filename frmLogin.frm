VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmLogin 
   Caption         =   "UserForm1"
   ClientHeight    =   4815
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8355.001
   OleObjectBlob   =   "frmLogin.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    On Error Resume Next
    Me.StartUpPosition = 1
    Me.caption = "GLOBAL SOFT - Login"
    txtPassword.PasswordChar = "*"
    If Not btnShowPwd Is Nothing Then btnShowPwd.caption = "Show"
    LoadUserDropdown
    On Error GoTo 0
End Sub

Private Sub LoadUserDropdown()
    On Error Resume Next
    cmbRole.Clear
    cmbRole.AddItem "ADMIN"
    cmbRole.AddItem "DEVELOPER"

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("User_Permission")
    If ws Is Nothing Then Exit Sub

    Dim lastRow As Long, i As Long, uname As String
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    For i = 2 To lastRow
        uname = UCase(Trim(ws.Cells(i, 1).value & ""))
        If uname <> "" And uname <> "ADMIN" And uname <> "DEVELOPER" And _
           uname <> "YES" And uname <> "NO" And uname <> "TRUE" And uname <> "FALSE" Then
            cmbRole.AddItem uname
        End If
    Next i

    If cmbRole.ListCount > 0 Then cmbRole.ListIndex = 0
    On Error GoTo 0
End Sub

Private Sub btnLogin_Click()
    On Error GoTo LoginError

    Dim selectedUser As String, enteredPass As String, found As Boolean
    selectedUser = UCase(Trim(cmbRole.value & ""))
    enteredPass = Trim(txtPassword.value & "")
    found = False

    If selectedUser = "" Then
        MsgBox "Please select a user!", vbExclamation, "Login"
        cmbRole.SetFocus
        Exit Sub
    End If
    If enteredPass = "" Then
        MsgBox "Please enter password!", vbExclamation, "Login"
        txtPassword.SetFocus
        Exit Sub
    End If

    ' Check User_Permission sheet: A=Name, B=Password, C=Role
    Dim ws As Worksheet, lastRow As Long, i As Long
    Set ws = ThisWorkbook.Sheets("User_Permission")
    If Not ws Is Nothing Then
        lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
        For i = 2 To lastRow
            If UCase(Trim(ws.Cells(i, 1).value & "")) = selectedUser Then
                If Trim(ws.Cells(i, 2).value & "") = enteredPass Then
                    found = True
                    gUserName = Trim(ws.Cells(i, 1).value & "")
                    gUserRole = Trim(ws.Cells(i, 3).value & "")
                    gUserRow = i
                    CurrentUserName = gUserName
                    CurrentUserRole = gUserRole
                    IsLoggedIn = True
                    Exit For
                End If
            End If
        Next i
    End If

    ' Fallback hardcoded
    If Not found Then
        If selectedUser = "ADMIN" And UCase(enteredPass) = "ADMIN123" Then
            found = True
            gUserName = "ADMIN": gUserRole = "ADMIN": gUserRow = 0
            CurrentUserName = "ADMIN": CurrentUserRole = "ADMIN"
            IsLoggedIn = True
        ElseIf selectedUser = "DEVELOPER" And UCase(enteredPass) = "DEV123" Then
            found = True
            gUserName = "DEVELOPER": gUserRole = "DEVELOPER": gUserRow = 0
            CurrentUserName = "DEVELOPER": CurrentUserRole = "DEVELOPER"
            IsLoggedIn = True
        End If
    End If

    If Not found Then
        MsgBox "Wrong user name or password!", vbCritical, "Login Failed"
        txtPassword.value = ""
        txtPassword.SetFocus
        Exit Sub
    End If

    ' SUCCESS - just unload. Workbook_Open will show dashboard.
    Unload Me
    Exit Sub

LoginError:
    MsgBox "Login Error: " & Err.Description, vbCritical, "Error"
End Sub

Private Sub btnCancel_Click()
    IsLoggedIn = False
    Unload Me
End Sub

Private Sub btnShowPwd_Click()
    On Error Resume Next
    If txtPassword.PasswordChar = "*" Then
        txtPassword.PasswordChar = ""
        btnShowPwd.caption = "Hide"
    Else
        txtPassword.PasswordChar = "*"
        btnShowPwd.caption = "Show"
    End If
    On Error GoTo 0
End Sub

Private Sub txtPassword_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then btnLogin_Click
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        IsLoggedIn = False
    End If
End Sub

