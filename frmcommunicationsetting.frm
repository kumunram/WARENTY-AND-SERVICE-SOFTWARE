VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmcommunicationsetting 
   Caption         =   "COMMUNICATION SETTINGS"
   ClientHeight    =   11085
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10305
   OleObjectBlob   =   "frmcommunicationsetting.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmcommunicationsetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'===========================================
' SHELLEXECUTE DECLARE
'===========================================
Private Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
    ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
    ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long



'===========================================
' RESET TO DEFAULT
'===========================================
Private Sub btnReset_Click()
    Dim result As VbMsgBoxResult
    result = MsgBox("Reset to default values?" & vbCrLf & _
                    "Unsaved changes will be lost.", vbQuestion + vbYesNo, "Reset")
    
    If result = vbNo Then Exit Sub
    
    ' Clear WhatsApp
    Me.txtWhatsAppNumber.value = ""
    
    ' Clear Email
    Me.txtEmailID.value = ""
    Me.txtEmailPassword.value = ""
    Me.cmbSMTPServer.value = ""
    Me.txtAppPassword.value = ""
    Me.txtPort.value = ""
    Me.chkSSL.value = False
    
    ' Clear Note
    Me.lblNote.caption = "Select SMTP Server"
    Me.lblNote.ForeColor = RGB(0, 0, 0)
    
    MsgBox "Reset to Default completed!", vbInformation
End Sub

'===========================================
' CLOSE BUTTON
'===========================================
Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnSaveComm_Click()

End Sub



Private Sub Frame1_Click()

End Sub

'===========================================
' FORM INITIALIZE
'===========================================
Private Sub UserForm_Initialize()
    LoadSMTPComboBox
    LoadCommunicationSettings
    
    ' App Password auto hide
    Me.txtAppPassword.PasswordChar = "*"
End Sub

'===========================================
' LOAD SMTP SERVER COMBOBOX
'===========================================
Private Sub LoadSMTPComboBox()
    With Me.cmbSMTPServer
        .Clear
        .AddItem "Gmail"
        .AddItem "Outlook"
        .AddItem "Yahoo"
        .AddItem "Hotmail"
        .AddItem "Custom"
    End With
End Sub

'===========================================
' LOAD COMMUNICATION SETTINGS
'===========================================
Private Sub LoadCommunicationSettings()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Software_Config")
    On Error GoTo 0
    
    If ws Is Nothing Then Exit Sub
    
    ' WhatsApp Settings
    Me.txtWhatsAppNumber.value = ws.Range("B16").value
    
    ' Email Settings
    Me.txtEmailID.value = ws.Range("B17").value
    Me.txtEmailPassword.value = ws.Range("B18").value
    Me.cmbSMTPServer.value = ws.Range("B19").value
    Me.txtAppPassword.value = ws.Range("B20").value
    Me.txtPort.value = ws.Range("B21").value
    Me.chkSSL.value = (ws.Range("B22").value = "Yes")
    
    ' Hide passwords
    Me.txtEmailPassword.PasswordChar = "*"
    Me.txtAppPassword.PasswordChar = "*"
End Sub

'===========================================
' SMTP SERVER CHANGE - Auto Fill Port + Instructions
'===========================================
Private Sub cmbSMTPServer_Change()
    Dim server As String
    server = Me.cmbSMTPServer.value
    
    Select Case server
        Case "Gmail"
            Me.txtPort.value = "465"
            Me.chkSSL.value = True
            Me.lblNote.caption = "?? Gmail requires App Password! Click 'Get App Password'"
            Me.lblNote.ForeColor = RGB(255, 0, 0)
            
        Case "Outlook"
            Me.txtPort.value = "587"
            Me.chkSSL.value = True
            Me.lblNote.caption = "? Use regular Outlook password"
            Me.lblNote.ForeColor = RGB(0, 128, 0)
            
        Case "Yahoo"
            Me.txtPort.value = "465"
            Me.chkSSL.value = True
            Me.lblNote.caption = "?? Yahoo requires App Password! Click 'Get App Password'"
            Me.lblNote.ForeColor = RGB(255, 0, 0)
            
        Case "Hotmail"
            Me.txtPort.value = "587"
            Me.chkSSL.value = True
            Me.lblNote.caption = "? Use regular Hotmail password"
            Me.lblNote.ForeColor = RGB(0, 128, 0)
            
        Case "Custom"
            Me.txtPort.value = ""
            Me.chkSSL.value = False
            Me.lblNote.caption = "Enter custom SMTP settings"
            Me.lblNote.ForeColor = RGB(0, 0, 0)
    End Select
End Sub

'===========================================
' SHOW/HIDE EMAIL PASSWORD
'===========================================
Private Sub btnHideEmail_Click()
    If Me.txtEmailPassword.PasswordChar = "*" Then
        Me.txtEmailPassword.PasswordChar = ""
        Me.btnHideEmail.caption = "Hide"
    Else
        Me.txtEmailPassword.PasswordChar = "*"
        Me.btnHideEmail.caption = "Show"
    End If
End Sub

'===========================================
' SHOW/HIDE APP PASSWORD
'===========================================
Private Sub btnHideApp_Click()
    If Me.txtAppPassword.PasswordChar = "*" Then
        Me.txtAppPassword.PasswordChar = ""
        Me.btnHideApp.caption = "Hide"
    Else
        Me.txtAppPassword.PasswordChar = "*"
        Me.btnHideApp.caption = "Show"
    End If
End Sub

'===========================================
' GET APP PASSWORD BUTTON
'===========================================
Private Sub btnGetAppPassword_Click()
    Dim server As String
    server = Me.cmbSMTPServer.value
    
    If server = "" Then
        MsgBox "Please select SMTP Server first!", vbExclamation
        Exit Sub
    End If
    
    Select Case server
        Case "Gmail"
            ShowGmailInstructions
        Case "Outlook"
            MsgBox "Outlook uses regular password. No App Password needed!", vbInformation
        Case "Yahoo"
            ShowYahooInstructions
        Case "Hotmail"
            MsgBox "Hotmail uses regular password. No App Password needed!", vbInformation
        Case Else
            MsgBox "Please select SMTP Server!", vbExclamation
    End Select
End Sub

'===========================================
' GMAIL INSTRUCTIONS
'===========================================
Private Sub ShowGmailInstructions()
    Dim msg As String
    
    msg = "=== GMAIL APP PASSWORD SETUP ===" & vbCrLf & vbCrLf & _
          "Step 1: Open browser and go to:" & vbCrLf & _
          "        myaccount.google.com" & vbCrLf & vbCrLf & _
          "Step 2: Click 'Security' ? '2-Step Verification' ? Enable" & vbCrLf & vbCrLf & _
          "Step 3: Go back to Security ? 'App Passwords'" & vbCrLf & vbCrLf & _
          "Step 4: Select 'Mail' and 'Other (Custom name)'" & vbCrLf & _
          "        Type: GLOBAL SOFT" & vbCrLf & vbCrLf & _
          "Step 5: Click 'Generate'" & vbCrLf & vbCrLf & _
          "Step 6: Copy the 16-digit password" & vbCrLf & _
          "        Example: abcd efgh ijkl mnop" & vbCrLf & vbCrLf & _
          "Step 7: Paste it in 'APP Password' field" & vbCrLf & vbCrLf & _
          "Open Gmail Settings now?"
    
    If MsgBox(msg, vbQuestion + vbYesNo, "Gmail App Password") = vbYes Then
        ShellExecute 0, "open", "https://myaccount.google.com/security", vbNullString, vbNullString, 1
    End If
End Sub

'===========================================
' YAHOO INSTRUCTIONS
'===========================================
Private Sub ShowYahooInstructions()
    Dim msg As String
    
    msg = "=== YAHOO APP PASSWORD SETUP ===" & vbCrLf & vbCrLf & _
          "Step 1: Go to: login.yahoo.com/account/security" & vbCrLf & vbCrLf & _
          "Step 2: Enable '2-Step Verification'" & vbCrLf & vbCrLf & _
          "Step 3: Generate 'App Password'" & vbCrLf & vbCrLf & _
          "Step 4: Select 'Other App' ? Type 'Mail'" & vbCrLf & vbCrLf & _
          "Step 5: Copy password and paste in 'APP Password' field" & vbCrLf & vbCrLf & _
          "Open Yahoo Security now?"
    
    If MsgBox(msg, vbQuestion + vbYesNo, "Yahoo Setup") = vbYes Then
        ShellExecute 0, "open", "https://login.yahoo.com/account/security", vbNullString, vbNullString, 1
    End If
End Sub
'===========================================
' TEST WHATSAPP BUTTON - With Choice
'===========================================
Private Sub btnTestWhats_Click()
    Dim mobile As String
    Dim testMsg As String
    
    mobile = Me.txtWhatsAppNumber.value
    mobile = Replace(mobile, " ", "")
    mobile = Replace(mobile, "-", "")
    
    If mobile = "" Then
        MsgBox "Enter WhatsApp number!", vbExclamation
        Exit Sub
    End If
    
    testMsg = "Hello! This is a test message from GLOBAL SOFT."
    
    ' User ?? ????
    Dim result As VbMsgBoxResult
    result = MsgBox("Open WhatsApp in:" & vbCrLf & vbCrLf & _
                    "YES = Desktop App (if installed)" & vbCrLf & _
                    "NO = Web Browser", vbQuestion + vbYesNoCancel, "WhatsApp")
    
    If result = vbCancel Then Exit Sub
    
    If result = vbYes Then
        ' Desktop App Try ???
        ShellExecute 0, "open", "whatsapp://send?phone=91" & mobile & "&text=" & EncodeURL(testMsg), vbNullString, vbNullString, 1
    Else
        ' Browser ??? ????
        Dim url As String
        url = "https://wa.me/91" & mobile & "?text=" & EncodeURL(testMsg)
        ShellExecute 0, "open", url, vbNullString, vbNullString, 1
    End If
End Sub
'===========================================
' URL ENCODE FUNCTION
'===========================================
Private Function EncodeURL(ByVal text As String) As String
    Dim i As Integer
    Dim result As String
    
    For i = 1 To Len(text)
        Select Case Mid(text, i, 1)
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
            Case Else: result = result & Mid(text, i, 1)
        End Select
    Next i
    
    EncodeURL = result
End Function
'===========================================
' TEST EMAIL - Gmail (CDO Method)
'===========================================
Private Sub btnTestEmail_Click()
    Dim email As String
    Dim appPassword As String
    
    email = Me.txtEmailID.value
    appPassword = Me.txtAppPassword.value
    
    If email = "" Or appPassword = "" Then
        MsgBox "Enter Email and App Password!", vbExclamation
        Exit Sub
    End If
    
    ' CDO ?? email ???? - Outlook ???? ??????!
    SendEmailCDO email, appPassword
End Sub

Private Sub SendEmailCDO(ByVal email As String, ByVal password As String)
    On Error GoTo ErrorHandler
    
    Dim cdoConfig As Object
    Dim cdoMessage As Object
    
    Set cdoConfig = CreateObject("CDO.Configuration")
    Set cdoMessage = CreateObject("CDO.Message")
    
    With cdoConfig.Fields
        .item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
        .item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.gmail.com"
        .item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 465  ' Changed from 587
        .item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
        .item("http://schemas.microsoft.com/cdo/configuration/sendusername") = email
        .item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = password
        .item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
        .item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
        .Update
    End With
    
    With cdoMessage
        Set .Configuration = cdoConfig
        .From = email
        .To = email
        .subject = "Test Email from GLOBAL SOFT"
        .TextBody = "Hello! This is a test email."
        .Send
    End With
    
    MsgBox "Email sent successfully!", vbInformation
    Exit Sub
    
ErrorHandler:
    MsgBox "Failed: " & Err.Description & vbCrLf & vbCrLf & _
           "Try:" & vbCrLf & _
           "1. Enable 'Less secure app access' in Gmail" & vbCrLf & _
           "2. Or use Hotmail/Outlook account", vbExclamation
End Sub

'===========================================
' SEND TEST EMAIL
'===========================================
Private Sub SendTestEmail(ByVal email As String, ByVal password As String, _
                          ByVal server As String, ByVal port As String, ByVal ssl As Boolean)
    
    On Error GoTo ErrorHandler
    
    Dim cdoConfig As Object
    Dim cdoMessage As Object
    
    Set cdoConfig = CreateObject("CDO.Configuration")
    Set cdoMessage = CreateObject("CDO.Message")
    
    ' SMTP Configuration
    With cdoConfig.Fields
        .item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
        .item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = GetSMTPServer(server)
        .item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = CInt(port)
        .item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
        .item("http://schemas.microsoft.com/cdo/configuration/sendusername") = email
        .item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = password
        .item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = ssl
        .Update
    End With
    
    ' Email Message
    With cdoMessage
        Set .Configuration = cdoConfig
        .From = email
        .To = email
        .subject = "Test Email from GLOBAL SOFT"
        .TextBody = "Hello!" & vbCrLf & vbCrLf & _
                    "This is a test email from GLOBAL SOFT." & vbCrLf & vbCrLf & _
                    "Your email settings are working correctly!" & vbCrLf & vbCrLf & _
                    "Regards," & vbCrLf & "GLOBAL SOFT Team"
        .Send
    End With
    
    MsgBox "Test email sent successfully!" & vbCrLf & "Check your inbox.", vbInformation
    Exit Sub
    
ErrorHandler:
    MsgBox "Email failed!" & vbCrLf & "Error: " & Err.Description, vbExclamation
End Sub

'===========================================
' GET SMTP SERVER ADDRESS
'===========================================
Private Function GetSMTPServer(ByVal serverName As String) As String
    Select Case serverName
        Case "Gmail": GetSMTPServer = "smtp.gmail.com"
        Case "Outlook": GetSMTPServer = "smtp.office365.com"
        Case "Yahoo": GetSMTPServer = "smtp.mail.yahoo.com"
        Case "Hotmail": GetSMTPServer = "smtp.live.com"
        Case Else: GetSMTPServer = serverName
    End Select
End Function
'===========================================
' SAVE ALL SETTINGS
'===========================================
Private Sub btnSave_Click()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Software_Config")
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "Software_Config sheet not found!", vbExclamation
        Exit Sub
    End If
    
    ' WhatsApp Settings
    ws.Range("B16").value = Me.txtWhatsAppNumber.value
    
    ' Email Settings
    ws.Range("B17").value = Me.txtEmailID.value
    ws.Range("B18").value = Me.txtEmailPassword.value
    ws.Range("B19").value = Me.cmbSMTPServer.value
    ws.Range("B20").value = Me.txtAppPassword.value
    ws.Range("B21").value = Me.txtPort.value
    ws.Range("B22").value = IIf(Me.chkSSL.value, "Yes", "No")
    
    MsgBox "All Settings Saved Successfully!", vbInformation, "Saved"
End Sub



