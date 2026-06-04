VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmUserPermission 
   Caption         =   "UserForm1"
   ClientHeight    =   9525.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14250
   OleObjectBlob   =   "frmUserPermission.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmUserPermission"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_NewUserName As String   ' ??? ???? ?? ??? ??? ?????
Private m_IsNewMode As Boolean    ' ??? ??? ?? ????

'========================================
' USERFORM INITIALIZE
'========================================
Private Sub UserForm_Initialize()
    On Error Resume Next
    Me.caption = "User Permission Manager"
    m_IsNewMode = False
    m_NewUserName = ""
    LoadUserList
    ClearPermissions
    On Error GoTo 0
End Sub

'========================================
' LOAD USER LIST INTO COMBOBOX
'========================================
Private Sub LoadUserList()
    On Error Resume Next
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim userName As String
    
    Set ws = ThisWorkbook.Sheets("User_Permission")
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    
    cmbUsers.Clear
    
    For i = 2 To lastRow
        userName = Trim(ws.Cells(i, 1).value & "")
        If userName <> "" Then
            cmbUsers.AddItem userName
        End If
    Next i
    
    ' === ADD NEW USER OPTION ===
    cmbUsers.AddItem "+ Add New User"
    
    If cmbUsers.ListCount > 0 Then cmbUsers.ListIndex = 0
    On Error GoTo 0
End Sub

'========================================
' WHEN USER SELECTED FROM DROPDOWN
'========================================
Private Sub cmbUsers_Change()
    On Error Resume Next
    
    If cmbUsers.value = "+ Add New User" Then
        ' === INPUTBOX ?? ??? ??? ?? ===
        Dim newName As String
        newName = InputBox("Enter New User Name:", "Add New User")
        
        If Trim(newName) = "" Then
            ' Cancelled or empty — reset to first user
            If cmbUsers.ListCount > 1 Then cmbUsers.ListIndex = 0
            On Error GoTo 0
            Exit Sub
        End If
        
        ' === CHECK DUPLICATE ===
        Dim ws As Worksheet
        Dim lastRow As Long, i As Long
        Set ws = ThisWorkbook.Sheets("User_Permission")
        lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
        
        For i = 2 To lastRow
            If UCase(Trim(ws.Cells(i, 1).value & "")) = UCase(Trim(newName)) Then
                MsgBox "User '" & newName & "' already exists!" & vbCrLf & _
                       "Please select from dropdown.", vbExclamation, "Duplicate"
                If cmbUsers.ListCount > 1 Then cmbUsers.ListIndex = 0
                On Error GoTo 0
                Exit Sub
            End If
        Next i
        
        ' === NEW USER MODE ON ===
        m_IsNewMode = True
        m_NewUserName = Trim(newName)
        
        txtPassword.text = ""
        txtPassword.PasswordChar = "*"
        Call ClearPermissions
        Call EnablePermissions
        
        Me.caption = "User Permission Manager - ADDING: " & m_NewUserName
        txtPassword.SetFocus
        
    ElseIf cmbUsers.value <> "" Then
        ' === EXISTING USER LOAD ===
        m_IsNewMode = False
        m_NewUserName = ""
        LoadUserData cmbUsers.value
    End If
    
    On Error GoTo 0
End Sub

'========================================
' LOAD USER DATA FROM SHEET
'========================================
Private Sub LoadUserData(ByVal userName As String)
    On Error Resume Next
    
    ' === SAFETY: Admin cannot load DEVELOPER data ===
    If UCase(CurrentUserRole) = "ADMIN" And UCase(Trim(userName)) = "DEVELOPER" Then
        MsgBox "You cannot modify DEVELOPER account!", vbExclamation
        Call ClearPermissions
        Call DisablePermissions
        Exit Sub
    End If
    
    On Error GoTo 0
    
    ' === SAFE BACKUP CLEAR ===
    On Error Resume Next
    Me.Controls("chkBackup").value = False
    On Error GoTo 0
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("User_Permission")
    Dim lastRow As Long, i As Long
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    
    For i = 2 To lastRow
        If UCase(Trim(ws.Cells(i, 1).value & "")) = UCase(userName) Then
            txtPassword.text = ws.Cells(i, 2).value & ""
            txtPassword.PasswordChar = "*"
            
            chkDashboard.value = (ws.Cells(i, 4).value & "" = "YES")
            chkNewEntry.value = (ws.Cells(i, 5).value & "" = "YES")
            chkAssignWork.value = (ws.Cells(i, 6).value & "" = "YES")
            chkServiceUpdate.value = (ws.Cells(i, 7).value & "" = "YES")
            chkCustomerDB.value = (ws.Cells(i, 8).value & "" = "YES")
            chkEngineerReg.value = (ws.Cells(i, 9).value & "" = "YES")
            chkReports.value = (ws.Cells(i, 10).value & "" = "YES")
            chkSettings.value = (ws.Cells(i, 11).value & "" = "YES")
            
            ' === SAFE BACKUP LOAD ===
            On Error Resume Next
            Me.Controls("chkBackup").value = (ws.Cells(i, 12).value & "" = "YES")
            On Error GoTo 0
            
            Call EnablePermissions
            Exit For
        End If
    Next i
    
    Me.caption = "User Permission Manager"
End Sub

'========================================
' CLEAR ALL CHECKBOXES
'========================================
Private Sub ClearPermissions()
    On Error Resume Next
    chkDashboard.value = False
    chkNewEntry.value = False
    chkAssignWork.value = False
    chkServiceUpdate.value = False
    chkCustomerDB.value = False
    chkEngineerReg.value = False
    chkReports.value = False
    chkSettings.value = False
    
    ' === SAFE BACKUP CLEAR ===
    Me.Controls("chkBackup").value = False
    
    On Error GoTo 0
End Sub

'========================================
' DISABLE ALL CHECKBOXES
'========================================
Private Sub DisablePermissions()
    On Error Resume Next
    chkDashboard.enabled = False
    chkNewEntry.enabled = False
    chkAssignWork.enabled = False
    chkServiceUpdate.enabled = False
    chkCustomerDB.enabled = False
    chkEngineerReg.enabled = False
    chkReports.enabled = False
    chkSettings.enabled = False
    
    ' === SAFE BACKUP DISABLE ===
    Me.Controls("chkBackup").enabled = False
    
    On Error GoTo 0
End Sub

'========================================
' ENABLE ALL CHECKBOXES
'========================================
Private Sub EnablePermissions()
    On Error Resume Next
    chkDashboard.enabled = True
    chkNewEntry.enabled = True
    chkAssignWork.enabled = True
    chkServiceUpdate.enabled = True
    chkCustomerDB.enabled = True
    chkEngineerReg.enabled = True
    chkReports.enabled = True
    chkSettings.enabled = True
    
    ' === SAFE BACKUP ENABLE ===
    Me.Controls("chkBackup").enabled = True
    
    On Error GoTo 0
End Sub

'========================================
' SELECT ALL BUTTON
'========================================
Private Sub btnSelectAll_Click()
    On Error Resume Next
    chkDashboard.value = True
    chkNewEntry.value = True
    chkAssignWork.value = True
    chkServiceUpdate.value = True
    chkCustomerDB.value = True
    chkEngineerReg.value = True
    chkReports.value = True
    chkSettings.value = True
    
    ' === SAFE BACKUP SELECT ===
    Me.Controls("chkBackup").value = True
    
    On Error GoTo 0
End Sub

'========================================
' DESELECT ALL BUTTON
'========================================
Private Sub btnDeselectAll_Click()
    On Error Resume Next
    ClearPermissions
    On Error GoTo 0
End Sub

'========================================
' SAVE PERMISSIONS TO SHEET
'========================================
Private Sub btnSave_Click()
    On Error GoTo SaveError
    
    Dim saveName As String
    
    ' === GET USERNAME ===
    If m_IsNewMode And m_NewUserName <> "" Then
        saveName = m_NewUserName
    ElseIf cmbUsers.value <> "" And cmbUsers.value <> "+ Add New User" Then
        saveName = Trim(cmbUsers.value)
    Else
        MsgBox "Please select or enter a user name!", vbExclamation, "Required"
        Exit Sub
    End If
    
    If Trim(txtPassword.text) = "" Then
        MsgBox "Please enter Password!", vbExclamation, "Required"
        txtPassword.SetFocus
        Exit Sub
    End If
    
    ' === SAFETY: Admin cannot modify DEVELOPER ===
    If UCase(CurrentUserRole) = "ADMIN" And UCase(saveName) = "DEVELOPER" Then
        MsgBox "Only DEVELOPER can modify DEVELOPER account!", vbCritical
        Exit Sub
    End If
    
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim foundRow As Long
    
    Set ws = ThisWorkbook.Sheets("User_Permission")
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    foundRow = 0
    
    ' === CHECK EXISTING ===
    For i = 2 To lastRow
        If UCase(Trim(ws.Cells(i, 1).value & "")) = UCase(saveName) Then
            foundRow = i
            Exit For
        End If
    Next i
    
    ' === SAVE DATA ===
    If foundRow = 0 Then
        ' === NEW USER ===
        foundRow = lastRow + 1
        ws.Cells(foundRow, 1).value = saveName
        ws.Cells(foundRow, 3).value = "USER"   ' Default role
        MsgBox "New user '" & saveName & "' created successfully!", vbInformation, "Added"
        m_IsNewMode = False
        m_NewUserName = ""
    Else
        ' === UPDATE EXISTING ===
        MsgBox "Permissions updated for " & saveName & "!", vbInformation, "Saved"
    End If
    
    ' === SAVE PASSWORD & PERMISSIONS ===
    ws.Cells(foundRow, 2).value = txtPassword.text
    
    ws.Cells(foundRow, "D").value = IIf(chkDashboard.value, "YES", "NO")
    ws.Cells(foundRow, "E").value = IIf(chkNewEntry.value, "YES", "NO")
    ws.Cells(foundRow, "F").value = IIf(chkAssignWork.value, "YES", "NO")
    ws.Cells(foundRow, "G").value = IIf(chkServiceUpdate.value, "YES", "NO")
    ws.Cells(foundRow, "H").value = IIf(chkCustomerDB.value, "YES", "NO")
    ws.Cells(foundRow, "I").value = IIf(chkEngineerReg.value, "YES", "NO")
    ws.Cells(foundRow, "J").value = IIf(chkReports.value, "YES", "NO")
    ws.Cells(foundRow, "K").value = IIf(chkSettings.value, "YES", "NO")
    
    ' === SAFE BACKUP SAVE ===
    On Error Resume Next
    ws.Cells(foundRow, "L").value = IIf(Me.Controls("chkBackup").value, "YES", "NO")
    On Error GoTo 0
    
    ' === REFRESH DROPDOWN ===
    If foundRow = lastRow + 1 Then
        ' New user added — refresh list and select it
        LoadUserList
        cmbUsers.value = saveName
    End If
    
    Me.caption = "User Permission Manager"
    
    Exit Sub
    
SaveError:
    MsgBox "Error saving permissions: " & Err.Description, vbCritical, "Save Error"
End Sub

'========================================
' SHOW PASSWORD BUTTON
'========================================
Private Sub btnShowPassword_Click()
    On Error Resume Next
    If txtPassword.PasswordChar = "*" Then
        txtPassword.PasswordChar = ""
        btnShowPassword.caption = "Hide Password"
    Else
        txtPassword.PasswordChar = "*"
        btnShowPassword.caption = "Show Password"
    End If
    On Error GoTo 0
End Sub

'========================================
' CLOSE BUTTON
'========================================
Private Sub btnClose_Click()
    Unload Me
End Sub

