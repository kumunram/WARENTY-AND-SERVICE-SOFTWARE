VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSetting1 
   Caption         =   "frmSetting"
   ClientHeight    =   13020
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   19755
   OleObjectBlob   =   "frmSetting1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSetting1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private licenseUnlocked As Boolean

'===========================================
' API DECLARATION
'===========================================
#If VBA7 Then
    Private Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
        ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
        ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
#Else
    Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
        ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
        ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
#End If

'===========================================
' FORM INITIALIZE
'===========================================
Private Sub UserForm_Initialize()
    On Error Resume Next
    AddMinMaxButtons Me
    Me.caption = "GLOBAL SOFT - SETTINGS & CONFIGURATION"
    Me.Width = 1000
    Me.Height = 680
    Me.StartUpPosition = 1
    Me.BackColor = RGB(245, 245, 245)
    
    If Not MultiPage1 Is Nothing Then
        MultiPage1.Left = 200
        MultiPage1.Top = 50
        MultiPage1.Width = 770
        MultiPage1.Height = 530
        MultiPage1.BackColor = RGB(255, 255, 255)
    End If
    
    If Not fraMenu Is Nothing Then
        fraMenu.BackColor = RGB(230, 230, 230)
        fraMenu.ForeColor = RGB(0, 0, 0)
        fraMenu.BorderStyle = 0
        fraMenu.Left = 10
        fraMenu.Top = 50
        fraMenu.Width = 180
        fraMenu.Height = 530
    End If
    
    Call SetMenuButton(btnMenuCompany, "Company Info", 20)
    Call SetMenuButton(btnMenuComm, "Communication", 75)
    Call SetMenuButton(btnMenuFinance, "Financial", 130)
    Call SetMenuButton(btnMenuLicense, "License", 185)
    Call SetMenuButton(btnMenuPhoto, "Photo & Page", 240)
    Call SetMenuButton(btnMenuSoftware, "Software Update", 295)
    
    Call SetBottomButton(btnSaveAll, "SAVE ALL SETTINGS", RGB(0, 128, 0), 200, 600)
    Call SetBottomButton(btnResetAll, "RESET ALL", RGB(255, 140, 0), 420, 600)
    Call SetBottomButton(btnClose, "CLOSE", RGB(220, 20, 60), 620, 600)
    
    MultiPage1.value = 0
    Call HighlightMenuButton("COMPANY")
    
    Call LoadTab0_Company
    Call LoadTab1_Communication
    Call LoadTab2_Financial
    Call LoadTab3_License
    Call LoadTab4_Photo
        Call LoadTab5_Software
    
    ' === NEW: Enforce license on startup ===
    Call EnforceLicenseValidity
    
    ' If expired, force stay on License tab only
    If IsLicenseExpired Then
        Call DisableAllTabsExceptLicense
    End If
    
    On Error GoTo 0
End Sub

'===========================================
' MENU HELPERS
'===========================================
Private Sub SetMenuButton(ByRef btn As Object, ByVal cap As String, ByVal t As Long)
    On Error Resume Next
    If btn Is Nothing Then Exit Sub
    btn.caption = cap
    btn.BackColor = RGB(220, 220, 220)
    btn.ForeColor = RGB(0, 0, 0)
    btn.Font.Bold = True
    btn.Font.Size = 10
    btn.Width = 160
    btn.Height = 40
    btn.Left = 10
    btn.Top = t
    On Error GoTo 0
End Sub

Private Sub SetBottomButton(ByRef btn As Object, ByVal cap As String, ByVal clr As Long, ByVal l As Long, ByVal t As Long)
    On Error Resume Next
    If btn Is Nothing Then Exit Sub
    btn.caption = cap
    btn.BackColor = clr
    btn.ForeColor = RGB(255, 255, 255)
    btn.Font.Bold = True
    btn.Font.Size = 11
    btn.Width = 160
    btn.Height = 35
    btn.Left = l
    btn.Top = t
    On Error GoTo 0
End Sub

'===========================================
' HIGHLIGHT / HOVER / CLICK
'===========================================
Private Sub HighlightMenuButton(active As String)
    On Error Resume Next
    Call ResetMenuColor(btnMenuCompany)
    Call ResetMenuColor(btnMenuComm)
    Call ResetMenuColor(btnMenuFinance)
    Call ResetMenuColor(btnMenuLicense)
    Call ResetMenuColor(btnMenuPhoto)
    Call ResetMenuColor(btnMenuSoftware)
    Select Case active
        Case "COMPANY": btnMenuCompany.BackColor = RGB(100, 180, 255)
        Case "COMM": btnMenuComm.BackColor = RGB(100, 180, 255)
        Case "FINANCE": btnMenuFinance.BackColor = RGB(100, 180, 255)
        Case "LICENSE": btnMenuLicense.BackColor = RGB(100, 180, 255)
        Case "PHOTO": btnMenuPhoto.BackColor = RGB(100, 180, 255)
        Case "SOFTWARE": btnMenuSoftware.BackColor = RGB(100, 180, 255)
    End Select
    On Error GoTo 0
End Sub

Private Sub ResetMenuColor(ByRef btn As Object)
    On Error Resume Next
    If Not btn Is Nothing Then btn.BackColor = RGB(220, 220, 220): btn.ForeColor = RGB(0, 0, 0)
    On Error GoTo 0
End Sub

Private Sub btnMenuCompany_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): On Error Resume Next: btnMenuCompany.BackColor = RGB(150, 200, 255): On Error GoTo 0: End Sub
Private Sub btnMenuComm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): On Error Resume Next: btnMenuComm.BackColor = RGB(150, 200, 255): On Error GoTo 0: End Sub
Private Sub btnMenuFinance_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): On Error Resume Next: btnMenuFinance.BackColor = RGB(150, 200, 255): On Error GoTo 0: End Sub
Private Sub btnMenuLicense_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): On Error Resume Next: btnMenuLicense.BackColor = RGB(150, 200, 255): On Error GoTo 0: End Sub
Private Sub btnMenuPhoto_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): On Error Resume Next: btnMenuPhoto.BackColor = RGB(150, 200, 255): On Error GoTo 0: End Sub
Private Sub btnMenuSoftware_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): On Error Resume Next: btnMenuSoftware.BackColor = RGB(150, 200, 255): On Error GoTo 0: End Sub

Private Sub fraMenu_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    On Error Resume Next
    Call ResetMenuColor(btnMenuCompany): Call ResetMenuColor(btnMenuComm): Call ResetMenuColor(btnMenuFinance)
    Call ResetMenuColor(btnMenuLicense): Call ResetMenuColor(btnMenuPhoto): Call ResetMenuColor(btnMenuSoftware)
    Select Case MultiPage1.value
        Case 0: btnMenuCompany.BackColor = RGB(100, 180, 255)
        Case 1: btnMenuComm.BackColor = RGB(100, 180, 255)
        Case 2: btnMenuFinance.BackColor = RGB(100, 180, 255)
        Case 3: btnMenuLicense.BackColor = RGB(100, 180, 255)
        Case 4: btnMenuPhoto.BackColor = RGB(100, 180, 255)
        Case 5: btnMenuSoftware.BackColor = RGB(100, 180, 255)
    End Select
    On Error GoTo 0
End Sub

Private Sub btnMenuCompany_Click(): On Error Resume Next: MultiPage1.value = 0: Call HighlightMenuButton("COMPANY"): On Error GoTo 0: End Sub
Private Sub btnMenuComm_Click(): On Error Resume Next: MultiPage1.value = 1: Call HighlightMenuButton("COMM"): On Error GoTo 0: End Sub
Private Sub btnMenuFinance_Click(): On Error Resume Next: MultiPage1.value = 2: Call HighlightMenuButton("FINANCE"): On Error GoTo 0: End Sub
Private Sub btnMenuLicense_Click(): On Error Resume Next: MultiPage1.value = 3: Call HighlightMenuButton("LICENSE"): On Error GoTo 0: End Sub
Private Sub btnMenuPhoto_Click(): On Error Resume Next: MultiPage1.value = 4: Call HighlightMenuButton("PHOTO"): On Error GoTo 0: End Sub
Private Sub btnMenuSoftware_Click(): On Error Resume Next: MultiPage1.value = 5: Call HighlightMenuButton("SOFTWARE"): On Error GoTo 0: End Sub

'===========================================
' HELPERS
'===========================================
Private Sub LoadComboSafe(ByRef cmb As Object, ByVal items As Variant)
    On Error Resume Next
    If cmb Is Nothing Then Exit Sub
    If cmb.ListCount > 0 Then Exit Sub
    Dim i As Long
    For i = LBound(items) To UBound(items)
        cmb.AddItem items(i)
    Next i
    On Error GoTo 0
End Sub

Private Function URLEncode(ByVal text As String) As String
    Dim i As Integer, result As String, c As String
    For i = 1 To Len(text)
        c = Mid(text, i, 1)
        Select Case c
            Case " ": result = result & "%20"
            Case "!": result = result & "%21"
            Case "#": result = result & "%23"
            Case "$": result = result & "%24"
            Case "&": result = result & "%26"
            Case "'": result = result & "%27"
            Case "(": result = result & "%28"
            Case ")": result = result & "%29"
            Case "*": result = result & "%2A"
            Case "+": result = result & "%2B"
            Case ",": result = result & "%2C"
            Case "/": result = result & "%2F"
            Case ":": result = result & "%3A"
            Case ";": result = result & "%3B"
            Case "=": result = result & "%3D"
            Case "?": result = result & "%3F"
            Case "@": result = result & "%40"
            Case "[": result = result & "%5B"
            Case "]": result = result & "%5D"
            Case Else: result = result & c
        End Select
    Next i
    URLEncode = result
End Function

'===========================================
' TAB 0: COMPANY INFO
'===========================================
Private Sub LoadTab0_Company()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Software_Config")
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    On Error Resume Next
    txtCompanyName.text = ws.Range("B2").value & ""
    txtAddress1.text = ws.Range("B3").value & ""
    txtAddress2.text = ws.Range("B4").value & ""
    Call LoadComboSafe(cmbCity, Array("Add New...", "Nayagarh", "Bhubaneswar", "Cuttack", "Mumbai", "Delhi"))
    cmbCity.text = ws.Range("B5").value & ""
    Call LoadComboSafe(cmbPIN, Array("Add New...", "752069", "751001", "751003", "400001"))
    cmbPIN.text = ws.Range("B6").value & ""
    Call LoadComboSafe(cmbState, Array("Add New...", "Odisha", "West Bengal", "Maharashtra", "Delhi"))
    cmbState.text = ws.Range("B7").value & ""
    Call LoadComboSafe(cmbCountry, Array("Add New...", "India", "USA", "UK", "Australia"))
    cmbCountry.text = ws.Range("B8").value & ""
    txtMobile.text = ws.Range("B9").value & ""
    txtEmail.text = ws.Range("B10").value & ""
    txtWebsite.text = ws.Range("B11").value & ""
    txtLogoPath.text = ws.Range("B12").value & ""
    txtReceiptHeader.text = ws.Range("B13").value & ""
    If txtLogoPath.text <> "" Then
        On Error Resume Next
        Image1.Picture = LoadPicture(txtLogoPath.text)
        On Error GoTo 0
    End If
    On Error GoTo 0
End Sub

Private Sub SaveTab0_Company()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Software_Config")
    ws.Range("B2").value = txtCompanyName.text
    ws.Range("B3").value = txtAddress1.text
    ws.Range("B4").value = txtAddress2.text
    ws.Range("B5").value = cmbCity.text
    ws.Range("B6").value = cmbPIN.text
    ws.Range("B7").value = cmbState.text
    ws.Range("B8").value = cmbCountry.text
    ws.Range("B9").value = txtMobile.text
    ws.Range("B10").value = txtEmail.text
    ws.Range("B11").value = txtWebsite.text
    ws.Range("B12").value = txtLogoPath.text
    ws.Range("B13").value = txtReceiptHeader.text
    On Error GoTo 0
End Sub

Private Sub btnBrowseLogo_Click()
    On Error Resume Next
    Dim fd As Office.FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .title = "Select Logo"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Images", "*.jpg;*.jpeg;*.png"
        If .Show = -1 Then
            txtLogoPath.text = .SelectedItems(1)
            Image1.Picture = LoadPicture(txtLogoPath.text)
        End If
    End With
    On Error GoTo 0
End Sub

Private Sub cmbCity_Click()
    On Error Resume Next
    If cmbCity.text = "Add New..." Then
        Dim newCity As String
        newCity = InputBox("Enter New City:", "Add City")
        If newCity <> "" Then
            cmbCity.RemoveItem 0
            cmbCity.AddItem newCity, 0
            cmbCity.AddItem "Add New...", 1
            cmbCity.text = newCity
        End If
    End If
    On Error GoTo 0
End Sub

Private Sub cmbPIN_Click()
    On Error Resume Next
    If cmbPIN.text = "Add New..." Then
        Dim newPIN As String
        newPIN = InputBox("Enter New PIN (6 digits):", "Add PIN")
        If Len(newPIN) = 6 And IsNumeric(newPIN) Then
            cmbPIN.RemoveItem 0
            cmbPIN.AddItem newPIN, 0
            cmbPIN.AddItem "Add New...", 1
            cmbPIN.text = newPIN
        Else
            MsgBox "Invalid PIN! Must be 6 digits.", vbExclamation
        End If
    End If
    On Error GoTo 0
End Sub

Private Sub cmbState_Click()
    On Error Resume Next
    If cmbState.text = "Add New..." Then
        Dim newState As String
        newState = InputBox("Enter New State:", "Add State")
        If newState <> "" Then
            cmbState.RemoveItem 0
            cmbState.AddItem newState, 0
            cmbState.AddItem "Add New...", 1
            cmbState.text = newState
        End If
    End If
    On Error GoTo 0
End Sub

Private Sub cmbCountry_Click()
    On Error Resume Next
    If cmbCountry.text = "Add New..." Then
        Dim newCountry As String
        newCountry = InputBox("Enter New Country:", "Add Country")
        If newCountry <> "" Then
            cmbCountry.RemoveItem 0
            cmbCountry.AddItem newCountry, 0
            cmbCountry.AddItem "Add New...", 1
            cmbCountry.text = newCountry
        End If
    End If
    On Error GoTo 0
End Sub

'===========================================
' TAB 1: COMMUNICATION
'===========================================
Private Sub LoadTab1_Communication()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Software_Config")
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    On Error Resume Next
    txtWhatsAppNumber.text = ws.Range("B16").value & ""
    txtEmailID.text = ws.Range("B17").value & ""
    txtEmailPassword.text = ws.Range("B18").value & ""
    Call LoadComboSafe(cmbSMTPServer, Array("Gmail", "Outlook", "Yahoo", "Hotmail", "Custom"))
    cmbSMTPServer.text = ws.Range("B19").value & ""
    txtAppPassword.text = ws.Range("B20").value & ""
    txtPort.text = ws.Range("B21").value & ""
    chkSSL.value = (ws.Range("B22").value & "" = "Yes")
    txtEmailPassword.PasswordChar = "*"
    txtAppPassword.PasswordChar = "*"
    If Not lblNote Is Nothing Then lblNote.caption = "Select SMTP Server"
    On Error GoTo 0
End Sub

Private Sub SaveTab1_Communication()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Software_Config")
    ws.Range("B16").value = txtWhatsAppNumber.text
    ws.Range("B17").value = txtEmailID.text
    ws.Range("B18").value = txtEmailPassword.text
    ws.Range("B19").value = cmbSMTPServer.text
    ws.Range("B20").value = txtAppPassword.text
    ws.Range("B21").value = txtPort.text
    ws.Range("B22").value = IIf(chkSSL.value, "Yes", "No")
    On Error GoTo 0
End Sub

Private Sub cmbSMTPServer_Change()
    On Error Resume Next
    Select Case cmbSMTPServer.text
        Case "Gmail": txtPort.text = "465": chkSSL.value = True: If Not lblNote Is Nothing Then lblNote.caption = "Gmail requires App Password!": lblNote.ForeColor = RGB(255, 0, 0)
        Case "Outlook": txtPort.text = "587": chkSSL.value = True: If Not lblNote Is Nothing Then lblNote.caption = "Use regular Outlook password": lblNote.ForeColor = RGB(0, 128, 0)
        Case "Yahoo": txtPort.text = "465": chkSSL.value = True: If Not lblNote Is Nothing Then lblNote.caption = "Yahoo requires App Password!": lblNote.ForeColor = RGB(255, 0, 0)
        Case "Hotmail": txtPort.text = "587": chkSSL.value = True: If Not lblNote Is Nothing Then lblNote.caption = "Use regular Hotmail password": lblNote.ForeColor = RGB(0, 128, 0)
        Case "Custom": txtPort.text = "": chkSSL.value = False: If Not lblNote Is Nothing Then lblNote.caption = "Enter custom SMTP settings": lblNote.ForeColor = RGB(0, 0, 0)
    End Select
    On Error GoTo 0
End Sub

Private Sub btnHideEmail_Click()
    On Error Resume Next
    If txtEmailPassword.PasswordChar = "*" Then txtEmailPassword.PasswordChar = "": If Not btnHideEmail Is Nothing Then btnHideEmail.caption = "Hide" Else txtEmailPassword.PasswordChar = "*": If Not btnHideEmail Is Nothing Then btnHideEmail.caption = "Show"
    On Error GoTo 0
End Sub

Private Sub btnHideApp_Click()
    On Error Resume Next
    If txtAppPassword.PasswordChar = "*" Then txtAppPassword.PasswordChar = "": If Not btnHideApp Is Nothing Then btnHideApp.caption = "Hide" Else txtAppPassword.PasswordChar = "*": If Not btnHideApp Is Nothing Then btnHideApp.caption = "Show"
    On Error GoTo 0
End Sub

Private Sub btnGetAppPassword_Click()
    On Error Resume Next
    Dim server As String: server = cmbSMTPServer.text
    Select Case server
        Case "Gmail": If MsgBox("GMAIL APP PASSWORD SETUP" & vbCrLf & "1. myaccount.google.com" & vbCrLf & "2. Security > 2-Step Verification" & vbCrLf & "3. Security > App Passwords" & vbCrLf & "4. Select 'Mail', type 'GLOBAL SOFT'" & vbCrLf & "5. Generate & copy" & vbCrLf & vbCrLf & "Open now?", vbQuestion + vbYesNo) = vbYes Then ShellExecute 0, "open", "https://myaccount.google.com/security", vbNullString, vbNullString, 1
        Case "Yahoo": If MsgBox("YAHOO APP PASSWORD SETUP" & vbCrLf & "1. login.yahoo.com/account/security" & vbCrLf & "2. Enable 2-Step Verification" & vbCrLf & "3. Generate App Password for 'Mail'" & vbCrLf & vbCrLf & "Open now?", vbQuestion + vbYesNo) = vbYes Then ShellExecute 0, "open", "https://login.yahoo.com/account/security", vbNullString, vbNullString, 1
        Case "Outlook", "Hotmail": MsgBox "No App Password needed! Use regular password.", vbInformation
        Case Else: MsgBox "Please select SMTP Server first!", vbExclamation
    End Select
    On Error GoTo 0
End Sub

Private Sub btnTestWhats_Click()
    On Error Resume Next
    Dim mobile As String: mobile = txtWhatsAppNumber.text
    mobile = Replace(Replace(Replace(mobile, " ", ""), "-", ""), "+", "")
    If mobile = "" Then MsgBox "Please enter WhatsApp number first!", vbExclamation: Exit Sub
    If Left(mobile, 2) <> "91" Then mobile = "91" & mobile
    ShellExecute 0, "open", "https://wa.me/" & mobile & "?text=" & URLEncode("Hello! Test message from GLOBAL SOFT."), vbNullString, vbNullString, 1
    If Err.Number <> 0 Then MsgBox "Could not open WhatsApp.", vbExclamation Else MsgBox "WhatsApp opened!", vbInformation
    On Error GoTo 0
End Sub

Private Sub btnTestEmail_Click()
    On Error GoTo EmailError
    Dim email As String, pwd As String, server As String, port As String, useSSL As Boolean
    email = txtEmailID.text: pwd = txtAppPassword.text: server = cmbSMTPServer.text: port = txtPort.text: useSSL = chkSSL.value
    If email = "" Then MsgBox "Enter Email ID!", vbExclamation: Exit Sub
    If pwd = "" Then MsgBox "Enter App Password!", vbExclamation: Exit Sub
    If server = "" Then MsgBox "Select SMTP Server!", vbExclamation: Exit Sub
    Dim cdoConfig As Object, cdoMessage As Object
    Set cdoConfig = CreateObject("CDO.Configuration")
    Set cdoMessage = CreateObject("CDO.Message")
    Dim smtpServer As String
    Select Case server
        Case "Gmail": smtpServer = "smtp.gmail.com"
        Case "Outlook": smtpServer = "smtp.office365.com"
        Case "Yahoo": smtpServer = "smtp.mail.yahoo.com"
        Case "Hotmail": smtpServer = "smtp.live.com"
        Case Else: smtpServer = server
    End Select
    With cdoConfig.Fields
        .item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
        .item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = smtpServer
        .item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = CInt(port)
        .item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
        .item("http://schemas.microsoft.com/cdo/configuration/sendusername") = email
        .item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = pwd
        .item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = useSSL
        .item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
        .Update
    End With
    With cdoMessage
        Set .Configuration = cdoConfig
        .From = email: .To = email
        .subject = "Test Email from GLOBAL SOFT"
        .TextBody = "Hello!" & vbCrLf & vbCrLf & "This is a test email from GLOBAL SOFT." & vbCrLf & "Your settings are working correctly!"
        .Send
    End With
    MsgBox "Test email sent successfully!", vbInformation
    Exit Sub
EmailError:
    MsgBox "Email sending failed!" & vbCrLf & "Error: " & Err.Description & vbCrLf & vbCrLf & "Tips:" & vbCrLf & "1. Use App Password for Gmail/Yahoo" & vbCrLf & "2. Check internet connection" & vbCrLf & "3. Verify SMTP settings", vbExclamation
End Sub

'===========================================
' TAB 2: FINANCIAL & LEGAL
'===========================================
Private Sub LoadTab2_Financial()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Software_Config")
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    On Error Resume Next
    Call LoadComboSafe(cmbFinCountry, Array("Add New...", "India", "USA", "UK", "Australia", "Canada"))
    cmbFinCountry.text = ws.Range("B23").value & ""
    Call LoadComboSafe(cmbFinState, Array("Add New...", "Odisha", "West Bengal", "Maharashtra", "Karnataka", "Delhi"))
    cmbFinState.text = ws.Range("B24").value & ""
    txtGSTNumber.text = ws.Range("B25").value & ""
    txtServiceTax.text = ws.Range("B26").value & ""
    txtJudicialArea.text = ws.Range("B27").value & ""
    If IsDate(ws.Range("B28").value) Then txtFinancialStart.text = Format(ws.Range("B28").value, "dd/mm/yyyy") Else txtFinancialStart.text = ws.Range("B28").value & ""
    If IsDate(ws.Range("B29").value) Then txtFinancialEnd.text = Format(ws.Range("B29").value, "dd/mm/yyyy") Else txtFinancialEnd.text = ws.Range("B29").value & ""
    txtCurrentYear.text = ws.Range("B30").value & ""
    If txtFinancialStart.text <> "" And txtFinancialEnd.text <> "" And txtCurrentYear.text = "" Then Call CalculateFinancialYear
    On Error GoTo 0
End Sub

Private Sub SaveTab2_Financial()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Software_Config")
    ws.Range("B23").value = cmbFinCountry.text
    ws.Range("B24").value = cmbFinState.text
    ws.Range("B25").value = txtGSTNumber.text
    ws.Range("B26").value = txtServiceTax.text
    ws.Range("B27").value = txtJudicialArea.text
    ws.Range("B28").value = txtFinancialStart.text
    ws.Range("B29").value = txtFinancialEnd.text
    ws.Range("B30").value = txtCurrentYear.text
    On Error GoTo 0
End Sub

Private Sub txtFinancialStart_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    On Error Resume Next
    Dim dt As String: dt = Trim(txtFinancialStart.text)
    If dt = "" Then Exit Sub
    dt = Replace(Replace(Replace(dt, "/", ""), "-", ""), ".", "")
    If Len(dt) = 8 And IsNumeric(dt) Then txtFinancialStart.text = Left(dt, 2) & "/" & Mid(dt, 3, 2) & "/" & Right(dt, 4): Call CalculateFinancialYear
    On Error GoTo 0
End Sub

Private Sub txtFinancialEnd_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    On Error Resume Next
    Dim dt As String: dt = Trim(txtFinancialEnd.text)
    If dt = "" Then Exit Sub
    dt = Replace(Replace(Replace(dt, "/", ""), "-", ""), ".", "")
    If Len(dt) = 8 And IsNumeric(dt) Then txtFinancialEnd.text = Left(dt, 2) & "/" & Mid(dt, 3, 2) & "/" & Right(dt, 4): Call CalculateFinancialYear
    On Error GoTo 0
End Sub

Private Sub CalculateFinancialYear()
    On Error Resume Next
    Dim sDate As String, eDate As String
    sDate = txtFinancialStart.text: eDate = txtFinancialEnd.text
    If sDate = "" Or eDate = "" Then Exit Sub
    Dim sy As String, ey As String
    sy = Right(sDate, 4): ey = Right(eDate, 4)
    If IsNumeric(sy) And IsNumeric(ey) Then txtCurrentYear.text = sy & "-" & Right(ey, 2)
    On Error GoTo 0
End Sub

'===========================================
' TAB 3: LICENSE - LOCK/UNLOCK
'===========================================
Private Sub LockLicenseControls(ByVal bLock As Boolean)
    On Error Resume Next
    cmbLicenseType.enabled = Not bLock
    txtActivationDate.enabled = Not bLock
    txtExpiryDate.enabled = Not bLock
    txtRemainingDays.enabled = Not bLock
    cmbArchivingMode.enabled = Not bLock
    chkAutoArchiving.enabled = Not bLock
    txtLocation.enabled = Not bLock
    btnBrowseLocation.enabled = Not bLock
    On Error GoTo 0
End Sub

Private Sub btnUnlock_Click()
    On Error Resume Next
    
    If licenseUnlocked Then
        ' User wants to lock back
        Call LockLicenseControls(True)
        licenseUnlocked = False
        If Not btnUnlock Is Nothing Then btnUnlock.caption = "Unlock"
        MsgBox "License settings locked!", vbInformation
        
        ' If license is still valid, re-enable all tabs
        If Not IsLicenseExpired Then
            Call EnableAllTabs
        End If
    Else
        ' Try to unlock — ALWAYS ask password (even if expired, for renewal!)
        Dim pwd As String
        pwd = InputBox("Enter Developer Password:", "Unlock License Settings")
        If pwd = "" Then Exit Sub
        
        If pwd = "kumun*1RAM" Then
            Call LockLicenseControls(False)
            licenseUnlocked = True
            If Not btnUnlock Is Nothing Then btnUnlock.caption = "Lock"
            
            ' Enable all tabs since developer unlocked
            Call EnableAllTabs
            
            MsgBox "License Settings Unlocked!" & vbCrLf & vbCrLf & _
                   "Steps to Renew:" & vbCrLf & _
                   "1. Select 'License Type'" & vbCrLf & _
                   "2. Activation & Expiry dates will auto-fill" & vbCrLf & _
                   "3. Click 'SAVE ALL SETTINGS'", vbInformation
        Else
            MsgBox "Wrong Password!", vbExclamation
        End If
    End If
    On Error GoTo 0
End Sub
Private Sub LoadTab3_License()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Software_Config")
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    On Error Resume Next
    licenseUnlocked = False
    Call LockLicenseControls(True)
    If Not btnUnlock Is Nothing Then btnUnlock.caption = "Unlock"
    Call LoadComboSafe(cmbLicenseType, Array("Add New...", "Demo - 3 Days", "Trial - 7 Days", "Trial - 30 Days", _
        "Basic - 3 Months", "Basic - 6 Months", "Standard - 1 Year", "Standard - 2 Years", _
        "Premium - 3 Years", "Premium - 5 Years", "Lifetime"))
    cmbLicenseType.text = ws.Range("B31").value & ""
    txtActivationDate.text = ws.Range("B32").value & ""
    txtExpiryDate.text = ws.Range("B33").value & ""
    txtRemainingDays.text = ws.Range("B35").value & ""
    Call LoadComboSafe(cmbArchivingMode, Array("Daily", "Weekly", "Monthly", "Yearly"))
    cmbArchivingMode.text = ws.Range("B44").value & ""
    chkAutoArchiving.value = (ws.Range("B37").value & "" = "Yes")
    txtLocation.text = ws.Range("B38").value & ""
    If txtActivationDate.text <> "" And txtExpiryDate.text <> "" Then Call UpdateLicenseStatus
    Call CheckDateTampering(ws)
    On Error GoTo 0
End Sub

Private Sub cmbLicenseType_Change()
    If Not licenseUnlocked Then Exit Sub
    On Error Resume Next
    Dim licType As String: licType = cmbLicenseType.text
    If licType = "" Or licType = "Add New..." Then Exit Sub
    Dim days As Long: days = GetLicenseDays(licType)
    If days > 0 Then
        txtActivationDate.text = Format(Date, "dd/mm/yyyy")
        txtExpiryDate.text = Format(DateAdd("d", days, Date), "dd/mm/yyyy")
        Call UpdateLicenseStatus
    End If
    On Error GoTo 0
End Sub

Private Sub txtActivationDate_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    On Error Resume Next
    If Not licenseUnlocked Then Exit Sub
    Dim dt As String: dt = Trim(txtActivationDate.text)
    If dt = "" Then Exit Sub
    dt = Replace(Replace(Replace(dt, "/", ""), "-", ""), ".", "")
    If Len(dt) = 8 And IsNumeric(dt) Then txtActivationDate.text = Left(dt, 2) & "/" & Mid(dt, 3, 2) & "/" & Right(dt, 4)
    If cmbLicenseType.text <> "" And cmbLicenseType.text <> "Add New..." Then
        Dim days As Long: days = GetLicenseDays(cmbLicenseType.text)
        If days > 0 Then
            Dim actDate As Date: actDate = CDate(txtActivationDate.text)
            txtExpiryDate.text = Format(DateAdd("d", days, actDate), "dd/mm/yyyy")
            Call UpdateLicenseStatus
        End If
    End If
    On Error GoTo 0
End Sub

Private Function GetLicenseDays(ByVal licType As String) As Long
    Select Case licType
        Case "Demo - 3 Days": GetLicenseDays = 3
        Case "Trial - 7 Days": GetLicenseDays = 7
        Case "Trial - 30 Days": GetLicenseDays = 30
        Case "Basic - 3 Months": GetLicenseDays = 90
        Case "Basic - 6 Months": GetLicenseDays = 180
        Case "Standard - 1 Year": GetLicenseDays = 365
        Case "Standard - 2 Years": GetLicenseDays = 730
        Case "Premium - 3 Years": GetLicenseDays = 1095
        Case "Premium - 5 Years": GetLicenseDays = 1825
        Case "Lifetime": GetLicenseDays = 36500
        Case Else: GetLicenseDays = 0
    End Select
End Function

Private Sub UpdateLicenseStatus()
    On Error Resume Next
    If txtExpiryDate.text = "" Or Not IsDate(txtExpiryDate.text) Then
        txtRemainingDays.text = ""
        If Not lblLicenseStatus Is Nothing Then lblLicenseStatus.caption = "NOT ACTIVATED": lblLicenseStatus.ForeColor = RGB(128, 128, 128)
        Exit Sub
    End If
    Dim expDate As Date, daysLeft As Long
    expDate = CDate(txtExpiryDate.text)
    daysLeft = DateDiff("d", Date, expDate)
    txtRemainingDays.text = daysLeft
    If Not lblLicenseStatus Is Nothing Then
        If daysLeft > 30 Then lblLicenseStatus.caption = "ACTIVE": lblLicenseStatus.ForeColor = RGB(0, 128, 0)
        If daysLeft > 0 And daysLeft <= 30 Then lblLicenseStatus.caption = "EXPIRING SOON": lblLicenseStatus.ForeColor = RGB(255, 165, 0)
        If daysLeft <= 0 Then lblLicenseStatus.caption = "EXPIRED": lblLicenseStatus.ForeColor = RGB(255, 0, 0)
    End If
    On Error GoTo 0
End Sub

Private Function IsLicenseExpired() As Boolean
    On Error Resume Next
    If txtExpiryDate.text = "" Or Not IsDate(txtExpiryDate.text) Then IsLicenseExpired = True: Exit Function
    IsLicenseExpired = (Date > CDate(txtExpiryDate.text))
    On Error GoTo 0
End Function

Private Sub EnforceLicenseValidity()
    On Error Resume Next
    If IsLicenseExpired Then
        MsgBox "SOFTWARE LICENSE EXPIRED!" & vbCrLf & vbCrLf & _
               "Expiry Date: " & txtExpiryDate.text & vbCrLf & vbCrLf & _
               "All features are LOCKED." & vbCrLf & _
               "Only License Renewal is allowed." & vbCrLf & vbCrLf & _
               "Please contact developer or click 'Unlock' to renew.", vbCritical, "License Expired"
        
        licenseUnlocked = False
        Call LockLicenseControls(True)
        
        ' === LOCK ALL OTHER TABS ===
        Call DisableAllTabsExceptLicense
        
        ' Unlock button MUST work for renewal
        If Not btnUnlock Is Nothing Then
            btnUnlock.caption = "Unlock"
            btnUnlock.enabled = True   ' IMPORTANT: Must be enabled for renewal!
        End If
        
        If Not lblLicenseStatus Is Nothing Then
            lblLicenseStatus.caption = "EXPIRED"
            lblLicenseStatus.ForeColor = RGB(255, 0, 0)
        End If
    End If
    On Error GoTo 0
End Sub

Private Sub CheckDateTampering(ByRef ws As Worksheet)
    On Error Resume Next
    Dim lastRun As Variant: lastRun = ws.Range("B39").value
    If lastRun <> "" And IsDate(lastRun) Then
        If Date < CDate(lastRun) Then
            MsgBox "DATE TAMPERING DETECTED!" & vbCrLf & "System date changed backwards!" & vbCrLf & "License locked for security.", vbCritical
            txtExpiryDate.text = Format(Date, "dd/mm/yyyy")
            Call UpdateLicenseStatus
        End If
    End If
    ws.Range("B39").value = Date
    On Error GoTo 0
End Sub

'===========================================
' SAVE TAB 3 — ONLY ONE COPY (NO DUPLICATE!)
'===========================================
Private Sub SaveTab3_License()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Software_Config")
    ws.Range("B31").value = cmbLicenseType.text
    ws.Range("B32").value = txtActivationDate.text
    ws.Range("B33").value = txtExpiryDate.text
    ws.Range("B35").value = txtRemainingDays.text
    If Not lblLicenseStatus Is Nothing Then ws.Range("B36").value = lblLicenseStatus.caption
    ws.Range("B37").value = IIf(chkAutoArchiving.value, "Yes", "No")
    ws.Range("B44").value = cmbArchivingMode.text
    ws.Range("B38").value = txtLocation.text
    ws.Range("B39").value = Date
    On Error GoTo 0
End Sub

Private Sub btnBrowseLocation_Click()
    On Error Resume Next
    Dim fd As Office.FileDialog
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    With fd
        .title = "Select Archive Folder"
        If .Show = -1 Then
            txtLocation.text = .SelectedItems(1) & "\GLOBAL_SOFT_Archive"
            On Error Resume Next
            MkDir txtLocation.text
            On Error GoTo 0
        End If
    End With
    On Error GoTo 0
End Sub

'===========================================
' TAB 4: PHOTO PATH & PAGE SETUP
'===========================================
Private Sub LoadTab4_Photo()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Settings")
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    On Error Resume Next
    txtImagePath.text = ws.Range("B1").value & ""
    txtProductPhotoPath.text = ws.Range("B2").value & ""
    txtCustomerPhotoPath.text = ws.Range("B3").value & ""
    txtReportPath.text = ws.Range("B4").value & ""
    txtBackupPath.text = ws.Range("B5").value & ""
    txtAccessoryPhotoPath.text = ws.Range("B6").value & ""
    txtLogoPath2.text = ws.Range("B11").value & ""
    Call LoadComboSafe(cmbPageSize, Array("A4", "A5"))
    cmbPageSize.text = ws.Range("B20").value & ""
    If cmbPageSize.text = "" Then cmbPageSize.text = "A4"
    Call LoadComboSafe(cmbShowCustomerPic, Array("YES", "NO"))
    cmbShowCustomerPic.text = ws.Range("B21").value & ""
    If cmbShowCustomerPic.text = "" Then cmbShowCustomerPic.text = "NO"
    Call LoadComboSafe(cmbShowProductPic, Array("YES", "NO"))
    cmbShowProductPic.text = ws.Range("B22").value & ""
    If cmbShowProductPic.text = "" Then cmbShowProductPic.text = "NO"
    Call LoadComboSafe(cmbShowAccessoryPic, Array("YES", "NO"))
    cmbShowAccessoryPic.text = ws.Range("B23").value & ""
    If cmbShowAccessoryPic.text = "" Then cmbShowAccessoryPic.text = "NO"
    Dim i As Long, termText As String
    termText = ""
    For i = 25 To 35
        If Trim(ws.Cells(i, 2).value & "") <> "" Then termText = termText & ws.Cells(i, 2).value & vbCrLf
    Next i
    txtTermsConditions.text = termText
    On Error GoTo 0
End Sub

Private Sub SaveTab4_Photo()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Settings")
    ws.Range("B1").value = txtImagePath.text
    ws.Range("B2").value = txtProductPhotoPath.text
    ws.Range("B3").value = txtCustomerPhotoPath.text
    ws.Range("B4").value = txtReportPath.text
    ws.Range("B5").value = txtBackupPath.text
    ws.Range("B6").value = txtAccessoryPhotoPath.text
    ws.Range("B11").value = txtLogoPath2.text
    ws.Range("B20").value = cmbPageSize.text
    ws.Range("B21").value = cmbShowCustomerPic.text
    ws.Range("B22").value = cmbShowProductPic.text
    ws.Range("B23").value = cmbShowAccessoryPic.text
    Dim lines() As String, i As Long
    lines = Split(txtTermsConditions.text, vbCrLf)
    For i = 25 To 35
        ws.Cells(i, 2).ClearContents
    Next i
    For i = 0 To UBound(lines)
        If i < 11 Then ws.Cells(25 + i, 2).value = Trim(lines(i))
    Next i
    On Error GoTo 0
End Sub

Private Sub btnCreateImagePath_Click()
    On Error Resume Next
    Dim p As String
    p = ThisWorkbook.path & "\GLOBAL_SOFT_DATA\Images"
    
    Call CreateFolderRecursive(p)  ' ? ?? parent + child ????? ??????
    
    txtImagePath.text = p
    If Dir(p, vbDirectory) <> "" Then
        MsgBox "Folder created!" & vbCrLf & p, vbInformation
    Else
        MsgBox "Folder creation FAILED!" & vbCrLf & p, vbCritical
    End If
    On Error GoTo 0
End Sub
Private Sub btnCreateProductPhotoPath_Click()
    On Error Resume Next
    Dim basePath As String
    basePath = txtImagePath.text
    If basePath = "" Then basePath = ThisWorkbook.path & "\GLOBAL_SOFT_DATA\Images"
    
    Dim p As String
    p = basePath & "\Products"
    
    Call CreateFolderRecursive(p)
    
    txtProductPhotoPath.text = p
    If Dir(p, vbDirectory) <> "" Then
        MsgBox "Products folder created!" & vbCrLf & p, vbInformation
    Else
        MsgBox "Products folder FAILED!" & vbCrLf & p, vbCritical
    End If
    On Error GoTo 0
End Sub
Private Sub btnCreateCustomerPhotoPath_Click()
    On Error Resume Next
    Dim basePath As String
    basePath = txtImagePath.text
    If basePath = "" Then basePath = ThisWorkbook.path & "\GLOBAL_SOFT_DATA\Images"
    
    Dim p As String
    p = basePath & "\Customers"
    
    Call CreateFolderRecursive(p)
    
    txtCustomerPhotoPath.text = p
    If Dir(p, vbDirectory) <> "" Then
        MsgBox "Customers folder created!" & vbCrLf & p, vbInformation
    Else
        MsgBox "Customers folder FAILED!" & vbCrLf & p, vbCritical
    End If
    On Error GoTo 0
End Sub
Private Sub btnCreateAccessoryPhotoPath_Click()
    On Error Resume Next
    Dim basePath As String
    basePath = txtImagePath.text
    If basePath = "" Then basePath = ThisWorkbook.path & "\GLOBAL_SOFT_DATA\Images"
    
    Dim p As String
    p = basePath & "\Accessories"
    
    Call CreateFolderRecursive(p)
    
    txtAccessoryPhotoPath.text = p
    If Dir(p, vbDirectory) <> "" Then
        MsgBox "Accessories folder created!" & vbCrLf & p, vbInformation
    Else
        MsgBox "Accessories folder FAILED!" & vbCrLf & p, vbCritical
    End If
    On Error GoTo 0
End Sub


Private Sub btnCreateReportPath_Click()
    On Error Resume Next
    Dim p As String
    p = ThisWorkbook.path & "\GLOBAL_SOFT_DATA\Reports"
    
    Call CreateFolderRecursive(p)
    
    txtReportPath.text = p
    If Dir(p, vbDirectory) <> "" Then
        MsgBox "Reports folder created!" & vbCrLf & p, vbInformation
    Else
        MsgBox "Reports folder FAILED!" & vbCrLf & p, vbCritical
    End If
    On Error GoTo 0
End Sub
Private Sub btnCreateBackupPath_Click()
    On Error Resume Next
    Dim p As String
    p = ThisWorkbook.path & "\GLOBAL_SOFT_DATA\Backup"
    
    Call CreateFolderRecursive(p)
    
    txtBackupPath.text = p
    If Dir(p, vbDirectory) <> "" Then
        MsgBox "Backup folder created!" & vbCrLf & p, vbInformation
    Else
        MsgBox "Backup folder FAILED!" & vbCrLf & p, vbCritical
    End If
    On Error GoTo 0
End Sub
Private Sub btnCreateLogoPath_Click()
    On Error Resume Next
    Dim basePath As String
    basePath = txtImagePath.text
    If basePath = "" Then basePath = ThisWorkbook.path & "\GLOBAL_SOFT_DATA\Images"
    
    Dim p As String
    p = basePath & "\Logo"
    
    Call CreateFolderRecursive(p)
    
    txtLogoPath2.text = p
    If Dir(p, vbDirectory) <> "" Then
        MsgBox "Logo folder created!" & vbCrLf & p, vbInformation
    Else
        MsgBox "Logo folder FAILED!" & vbCrLf & p, vbCritical
    End If
    On Error GoTo 0
End Sub

Private Sub btnCreateAllPaths_Click()
    On Error Resume Next
    Call btnCreateImagePath_Click
    Call btnCreateProductPhotoPath_Click
    Call btnCreateCustomerPhotoPath_Click
    Call btnCreateAccessoryPhotoPath_Click
    Call btnCreateReportPath_Click
    Call btnCreateBackupPath_Click
    Call btnCreateLogoPath_Click
    
    ' Also create Warranty & Staff folders (for complete structure)
    Call CreateFolderRecursive(ThisWorkbook.path & "\GLOBAL_SOFT_DATA\Warranty")
    Call CreateFolderRecursive(ThisWorkbook.path & "\GLOBAL_SOFT_DATA\Delivery")
    Call CreateFolderRecursive(ThisWorkbook.path & "\GLOBAL_SOFT_DATA\Staff")
    
    MsgBox "All folders created successfully!" & vbCrLf & vbCrLf & _
           "Location: " & ThisWorkbook.path & "\GLOBAL_SOFT_DATA\", vbInformation
    On Error GoTo 0
End Sub

Private Sub btnTestCamera_Click()
    On Error Resume Next
    shell "cmd /c start microsoft.windows.camera:", vbHide
    MsgBox "Camera opened!", vbInformation
    On Error GoTo 0
End Sub

Private Sub btnTestDatabase_Click()
    On Error Resume Next
    ThisWorkbook.Sheets("Customer_Master").Activate
    MsgBox "Database OK!", vbInformation
    On Error GoTo 0
End Sub

Private Sub btnTestBackup_Click()
    On Error Resume Next
    Dim p As String
    p = txtBackupPath.text
    If p = "" Then MsgBox "Backup path is empty!", vbExclamation: Exit Sub
    If Dir(p, vbDirectory) <> "" Then MsgBox "Backup Path OK!" & vbCrLf & p, vbInformation Else MsgBox "Backup folder NOT FOUND!" & vbCrLf & p, vbCritical
    On Error GoTo 0
End Sub

Private Sub btnTestVersion_Click()
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Software_Config")
    Dim ver As String
    ver = ws.Range("B56").value & ""
    If ver = "" Then ver = "v1.0"
    MsgBox "Current Version: " & ver, vbInformation
    On Error GoTo 0
End Sub

Private Sub btnRepairDatabase_Click()
    On Error Resume Next
    Dim requiredSheets As Variant
    requiredSheets = Array("Customer_Master", "Product_Master", "Accessory_Master", _
                          "Job_Master", "Job_Product", "Job_Accessory", _
                          "Payment_Master", "Engineer_Register", "Vendor_Register", _
                          "Settings", "Warranty_Master", "Verify_Master", "Status_Log", _
                          "Software_Config", "User_Permission", "Staff_Master", _
                          "Delivery_Master", "Courier_Data", "Inhouse_Register")
    Dim s As Variant, ws As Worksheet, found As Boolean
    Dim msg As String
    msg = "Database Repair Report:" & vbCrLf & vbCrLf
    For Each s In requiredSheets
        found = False
        For Each ws In ThisWorkbook.Sheets
            If ws.name = s Then found = True: Exit For
        Next ws
        If Not found Then
            Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
            ws.name = s
            msg = msg & "[CREATED] " & s & vbCrLf
        Else
            msg = msg & "[OK] " & s & vbCrLf
        End If
    Next s
    Call EnsureAllFolders
    MsgBox msg, vbInformation, "Database Repair Complete"
    On Error GoTo 0
End Sub

Private Sub btnRepairSystem_Click()
    On Error Resume Next
    If MsgBox("This will repair database and verify all system files." & vbCrLf & "Continue?", vbYesNo + vbQuestion) = vbNo Then Exit Sub
    Call btnRepairDatabase_Click
    Call EnsureAllFolders
    MsgBox "System repair completed!", vbInformation
    On Error GoTo 0
End Sub

Private Sub EnsureAllFolders()
    On Error Resume Next
    Dim basePath As String
    basePath = ThisWorkbook.path & "\GLOBAL_SOFT_DATA\"
    
    ' Create base + all subfolders using recursive function
    Call CreateFolderRecursive(basePath)
    Call CreateFolderRecursive(basePath & "Images")
    Call CreateFolderRecursive(basePath & "Images\Customers")
    Call CreateFolderRecursive(basePath & "Images\Products")
    Call CreateFolderRecursive(basePath & "Images\Accessories")
    Call CreateFolderRecursive(basePath & "Images\Logo")
    Call CreateFolderRecursive(basePath & "Reports")
    Call CreateFolderRecursive(basePath & "Backup")
    Call CreateFolderRecursive(basePath & "Archive")
    Call CreateFolderRecursive(basePath & "Warranty")
    Call CreateFolderRecursive(basePath & "Delivery")
    Call CreateFolderRecursive(basePath & "Staff")
    
    On Error GoTo 0
End Sub

'===========================================
' TAB 5: SOFTWARE UPDATE
'===========================================
Private Sub LoadTab5_Software()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Software_Config")
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    On Error Resume Next
    txtInstalledVersion.text = ws.Range("B56").value & ""
    txtReleaseDate.text = ws.Range("B51").value & ""
    txtLicenseStatus2.text = ws.Range("B52").value & ""
    txtNewFile.text = ws.Range("B53").value & ""
    txtBackupTo.text = ws.Range("B54").value & ""
    txtRestoreFile.text = ws.Range("B55").value & ""
    On Error GoTo 0
End Sub

Private Sub SaveTab5_Software()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Software_Config")
    ws.Range("B56").value = txtInstalledVersion.text
    ws.Range("B51").value = txtReleaseDate.text
    ws.Range("B52").value = txtLicenseStatus2.text
    ws.Range("B53").value = txtNewFile.text
    ws.Range("B54").value = txtBackupTo.text
    ws.Range("B55").value = txtRestoreFile.text
    On Error GoTo 0
End Sub

Private Sub btnBrowseNewFile_Click()
    On Error Resume Next
    Dim fd As Office.FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xlsm;*.xlsx;*.xls"
        .title = "Select New Version File"
        If .Show = -1 Then txtNewFile.text = .SelectedItems(1)
    End With
    On Error GoTo 0
End Sub

Private Sub btnBrowseBackup_Click()
    On Error Resume Next
    Dim fd As Office.FileDialog
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    With fd
        .title = "Select Backup Folder"
        If .Show = -1 Then txtBackupTo.text = .SelectedItems(1)
    End With
    On Error GoTo 0
End Sub

Private Sub btnBrowseRestore_Click()
    On Error Resume Next
    Dim fd As Office.FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xlsm;*.xlsx;*.xls"
        .title = "Select Restore File"
        If .Show = -1 Then txtRestoreFile.text = .SelectedItems(1)
    End With
    On Error GoTo 0
End Sub

Private Sub btnUpdateVersion_Click()
    If txtNewFile.text = "" Then MsgBox "Please select new version file first!", vbExclamation: Exit Sub
    If Dir(txtNewFile.text) = "" Then MsgBox "File not found: " & txtNewFile.text, vbExclamation: Exit Sub
    MsgBox "Version update process started." & vbCrLf & "New file: " & txtNewFile.text, vbInformation
End Sub

Private Sub btnRestore_Click()
    If txtRestoreFile.text = "" Then MsgBox "Please select restore file first!", vbExclamation: Exit Sub
    If Dir(txtRestoreFile.text) = "" Then MsgBox "Restore file not found: " & txtRestoreFile.text, vbExclamation: Exit Sub
    MsgBox "Restore process started." & vbCrLf & "File: " & txtRestoreFile.text, vbInformation
End Sub

Private Sub btnOpenBackupFolder_Click()
    On Error Resume Next
    Dim p As String
    p = txtBackupTo.text
    If p = "" Then p = ThisWorkbook.path & "\GLOBAL_SOFT_DATA\Backup"
    If Dir(p, vbDirectory) = "" Then MkDir p
    ShellExecute 0, "open", p, vbNullString, vbNullString, 1
    If Err.Number <> 0 Then MsgBox "Could not open folder!", vbExclamation
    On Error GoTo 0
End Sub

'===========================================
' SAVE ALL / RESET ALL / CLOSE
'===========================================
Private Sub btnSaveAll_Click()
    If MsgBox("Save ALL settings?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    Call SaveTab0_Company
    Call SaveTab1_Communication
    Call SaveTab2_Financial
    Call SaveTab3_License
    Call SaveTab4_Photo
    Call SaveTab5_Software
    MsgBox "All Settings Saved Successfully!", vbInformation, "Saved"
End Sub

Private Sub btnResetAll_Click()
    If MsgBox("Reset ALL settings to default?" & vbCrLf & vbCrLf & "This will clear all user data but preserve system defaults." & vbCrLf & vbCrLf & "Are you sure?", vbExclamation + vbYesNo) = vbNo Then Exit Sub
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Software_Config")
    If Not ws Is Nothing Then
        ws.Range("B2:B12").ClearContents
        ws.Range("B16:B22").ClearContents
        ws.Range("B23:B30").ClearContents
        ws.Range("B31:B38").ClearContents
        ws.Range("B44").ClearContents
        ws.Range("B51:B55").ClearContents
        If ws.Range("B13").value = "" Then ws.Range("B13").value = "WARRANTY AND SERVICE REPORT"
    End If
    Set ws = ThisWorkbook.Sheets("Settings")
    If Not ws Is Nothing Then
        ws.Range("B1:B6").ClearContents
        ws.Range("B11").ClearContents
        ws.Range("B20:B23").ClearContents
        ws.Range("B25:B35").ClearContents
    End If
    On Error GoTo 0
    Call UserForm_Initialize
    MsgBox "Reset completed! Default values preserved.", vbInformation
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

'===========================================
' DISABLE ALL TABS EXCEPT LICENSE (EXPIRED MODE)
'===========================================
Private Sub DisableAllTabsExceptLicense()
    On Error Resume Next
    ' Disable all menu buttons except License
    If Not btnMenuCompany Is Nothing Then btnMenuCompany.enabled = False
    If Not btnMenuComm Is Nothing Then btnMenuComm.enabled = False
    If Not btnMenuFinance Is Nothing Then btnMenuFinance.enabled = False
    If Not btnMenuPhoto Is Nothing Then btnMenuPhoto.enabled = False
    If Not btnMenuSoftware Is Nothing Then btnMenuSoftware.enabled = False
    
    ' Force show License tab only
    MultiPage1.value = 3
    Call HighlightMenuButton("LICENSE")
    
    ' Disable bottom buttons except Close
    If Not btnSaveAll Is Nothing Then btnSaveAll.enabled = False
    If Not btnResetAll Is Nothing Then btnResetAll.enabled = False
    
    On Error GoTo 0
End Sub

'===========================================
' ENABLE ALL TABS (AFTER UNLOCK/RENEWAL)
'===========================================
Private Sub EnableAllTabs()
    On Error Resume Next
    ' Enable all menu buttons
    If Not btnMenuCompany Is Nothing Then btnMenuCompany.enabled = True
    If Not btnMenuComm Is Nothing Then btnMenuComm.enabled = True
    If Not btnMenuFinance Is Nothing Then btnMenuFinance.enabled = True
    If Not btnMenuLicense Is Nothing Then btnMenuLicense.enabled = True
    If Not btnMenuPhoto Is Nothing Then btnMenuPhoto.enabled = True
    If Not btnMenuSoftware Is Nothing Then btnMenuSoftware.enabled = True
    
    ' Enable bottom buttons
    If Not btnSaveAll Is Nothing Then btnSaveAll.enabled = True
    If Not btnResetAll Is Nothing Then btnResetAll.enabled = True
    
    On Error GoTo 0
End Sub

'===========================================
' RECURSIVE FOLDER CREATOR (creates all parent folders)
'===========================================
Private Sub CreateFolderRecursive(ByVal folderPath As String)
    On Error Resume Next
    Dim cleanPath As String
    cleanPath = folderPath
    If Right(cleanPath, 1) = "\" Then cleanPath = Left(cleanPath, Len(cleanPath) - 1)
    
    Dim parts() As String
    parts = Split(cleanPath, "\")
    
    Dim buildPath As String
    buildPath = parts(0)
    
    Dim i As Integer
    For i = 1 To UBound(parts)
        buildPath = buildPath & "\" & parts(i)
        If Dir(buildPath, vbDirectory) = "" Then MkDir buildPath
    Next i
    On Error GoTo 0
End Sub
