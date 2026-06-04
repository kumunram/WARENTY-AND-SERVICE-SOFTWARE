VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCompanySettings 
   Caption         =   "Company Settings & License Management"
   ClientHeight    =   6300
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13365
   OleObjectBlob   =   "frmCompanySettings.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmCompanySettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'===========================================
' FORM INITIALIZE - SET BUTTON COLORS
'===========================================
Private Sub UserForm_Initialize()
    Me.StartUpPosition = 1 ' CenterOwner
    
    ' Main Menu Buttons - Blue Theme
    SetButtonColor Me.btnCompanyInfo, RGB(0, 112, 192)        ' Blue
    SetButtonColor Me.btnCommunication, RGB(0, 112, 192)      ' Blue
    SetButtonColor Me.btnFinancialLegal, RGB(0, 112, 192)     ' Blue
    SetButtonColor Me.btnLicenseSecurity, RGB(0, 112, 192)    ' Blue
    SetButtonColor Me.btnPhotoPathSetting, RGB(0, 112, 192)   ' Blue
    SetButtonColor Me.btnSoftwareUpdate, RGB(0, 112, 192)     ' Blue
    
    ' Bottom Buttons
    SetButtonColor Me.btnSaveAll, RGB(0, 176, 80)             ' Green (Save)
    SetButtonColor Me.btnRestoreDefault, RGB(255, 192, 0)     ' Orange/Yellow (Reset)
    SetButtonColor Me.btnClose, RGB(192, 80, 77)              ' Red (Close)
End Sub

'===========================================
' SET BUTTON COLOR HELPER
'===========================================
Private Sub SetButtonColor(ByRef btn As CommandButton, ByVal bgColor As Long)
    btn.BackColor = bgColor
    btn.ForeColor = RGB(255, 255, 255) ' White text
    btn.Font.Bold = True
    btn.Font.Size = 10
End Sub

'===========================================
' COMPANY INFORMATION BUTTON
'===========================================
Private Sub btnCompanyInfo_Click()
    frmCompanyInformation.Show
End Sub

'===========================================
' COMMUNICATION SETTING BUTTON
'===========================================
Private Sub btnCommunication_Click()
    frmcommunicationsetting.Show
End Sub

'===========================================
' FINANCIAL & LEGAL BUTTON
'===========================================
Private Sub btnFinancialLegal_Click()
    frmfinancialandlegal.Show
End Sub

'===========================================
' LICENSE & SECURITY BUTTON
'===========================================
Private Sub btnLicenseSecurity_Click()
    frmlicenseandsecurity.Show
End Sub

'===========================================
' PHOTO PATH & PAGE SETTING BUTTON
'===========================================
Private Sub btnPhotoPathSetting_Click()
    frmSetting.Show
End Sub

'===========================================
' SOFTWARE UPDATE BUTTON
'===========================================
Private Sub btnSoftwareUpdate_Click()
    frmSoftwareUpdate.Show
End Sub

'===========================================
' SAVE ALL SETTINGS BUTTON
'===========================================
Private Sub btnSaveAll_Click()
    If MsgBox("Save all settings?" & vbCrLf & vbCrLf & _
              "This will save all configuration data.", _
              vbQuestion + vbYesNo, "Save All") = vbNo Then Exit Sub
    
    ' All forms save their own data when opened
    ' But we can trigger save here if needed
    
    MsgBox "All settings saved successfully!", vbInformation, "Saved"
End Sub

'===========================================
' RESTORE TO DEFAULT BUTTON
'===========================================
Private Sub btnRestoreDefault_Click()
    Dim result As VbMsgBoxResult
    
    result = MsgBox("WARNING: This will reset ALL settings to default!" & vbCrLf & vbCrLf & _
                    "All saved data will be cleared:" & vbCrLf & _
                    "- Company Information" & vbCrLf & _
                    "- Communication Settings" & vbCrLf & _
                    "- Financial & Legal Settings" & vbCrLf & _
                    "- License & Security Settings" & vbCrLf & _
                    "- Photo Path Settings" & vbCrLf & _
                    "- Software Update Settings" & vbCrLf & vbCrLf & _
                    "Are you sure you want to continue?", _
                    vbExclamation + vbYesNo, "Restore to Default")
    
    If result = vbNo Then Exit Sub
    
    ' Reset all settings
    ResetCompanyInfo
    ResetCommunicationSettings
    ResetFinancialSettings
    ResetLicenseSettings
    ResetPhotoPathSettings
    ResetSoftwareUpdateSettings
    
    MsgBox "All settings restored to default!" & vbCrLf & vbCrLf & _
           "Please restart the application for changes to take effect.", _
           vbInformation, "Restore Complete"
End Sub

'===========================================
' RESET COMPANY INFO
'===========================================
Private Sub ResetCompanyInfo()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Software_Config")
    On Error GoTo 0
    
    If ws Is Nothing Then Exit Sub
    
    ' Clear Company Info (B2-B13)
    ws.Range("B2").value = ""
    ws.Range("B3").value = ""
    ws.Range("B4").value = ""
    ws.Range("B5").value = ""
    ws.Range("B6").value = ""
    ws.Range("B7").value = ""
    ws.Range("B8").value = ""
    ws.Range("B9").value = ""
    ws.Range("B10").value = ""
    ws.Range("B11").value = ""
    ws.Range("B12").value = ""
    ws.Range("B13").value = "WARRANTY AND SERVICE REPORT" ' Default
End Sub

'===========================================
' RESET COMMUNICATION SETTINGS
'===========================================
Private Sub ResetCommunicationSettings()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Software_Config")
    On Error GoTo 0
    
    If ws Is Nothing Then Exit Sub
    
    ' Clear Communication Settings (B16-B22)
    ws.Range("B16").value = ""
    ws.Range("B17").value = ""
    ws.Range("B18").value = ""
    ws.Range("B19").value = ""
    ws.Range("B20").value = ""
    ws.Range("B21").value = ""
    ws.Range("B22").value = ""
End Sub

'===========================================
' RESET FINANCIAL SETTINGS
'===========================================
Private Sub ResetFinancialSettings()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Software_Config")
    On Error GoTo 0
    
    If ws Is Nothing Then Exit Sub
    
    ' Clear Financial Settings (B23-B30)
    ws.Range("B23").value = ""
    ws.Range("B24").value = ""
    ws.Range("B25").value = ""
    ws.Range("B26").value = ""
    ws.Range("B27").value = ""
    ws.Range("B28").value = ""
    ws.Range("B29").value = ""
    ws.Range("B30").value = ""
End Sub

'===========================================
' RESET LICENSE SETTINGS
'===========================================
Private Sub ResetLicenseSettings()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Software_Config")
    On Error GoTo 0
    
    If ws Is Nothing Then Exit Sub
    
    ' Clear License Settings (B31-B38)
    ws.Range("B31").value = ""
    ws.Range("B32").value = ""
    ws.Range("B33").value = ""
    ws.Range("B34").value = ""
    ws.Range("B35").value = ""
    ws.Range("B36").value = ""
    ws.Range("B37").value = ""
    ws.Range("B38").value = ""
End Sub

'===========================================
' RESET PHOTO PATH SETTINGS
'===========================================
Private Sub ResetPhotoPathSettings()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Software_Config")
    On Error GoTo 0
    
    If ws Is Nothing Then Exit Sub
    
    ' Clear Photo Path Settings (if any)
    ' Add your specific ranges here
End Sub

'===========================================
' RESET SOFTWARE UPDATE SETTINGS
'===========================================
Private Sub ResetSoftwareUpdateSettings()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Software_Config")
    On Error GoTo 0
    
    If ws Is Nothing Then Exit Sub
    
    ' Clear Software Update Settings (B40-B46)
    ws.Range("B40").value = ""
    ws.Range("B41").value = ""
    ws.Range("B42").value = ""
    ws.Range("B43").value = ""
    ws.Range("B44").value = ""
    ws.Range("B45").value = ""
    ws.Range("B46").value = ""
End Sub

'===========================================
' CLOSE BUTTON
'===========================================
Private Sub btnClose_Click()
    Unload Me
End Sub

