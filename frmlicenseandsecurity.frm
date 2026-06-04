VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmlicenseandsecurity 
   Caption         =   " LICENSE & SECURITY (PASSWORD PROTECTED)"
   ClientHeight    =   11760
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9525.001
   OleObjectBlob   =   "frmlicenseandsecurity.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmlicenseandsecurity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'===========================================
' DEVELOPER PASSWORD
'===========================================
Private Const DEV_PASSWORD As String = "kumun*1RAM"

'===========================================
' FORM VARIABLES
'===========================================
Private isUnlocked As Boolean

'===========================================
' FORM INITIALIZE ??? Label ?? Caption ??? ???
'===========================================
Private Sub UserForm_Initialize()
    isUnlocked = False
    LockAllControls
    LoadLicenseTypes
    LoadArchivingModes
    SetDefaultArchiveLocation
    LoadLicenseSettings
    ' Check date tampering first
    If IsDateTampered Then
        ' Lock everything if tampered
        LockAllControls
        Me.btnUnlock.enabled = False
        Me.lblStatus.caption = "DATE TAMPERED!"
        Me.lblStatus.ForeColor = RGB(255, 0, 0)
        MsgBox "License validation failed due to date tampering!", vbCritical
        Exit Sub
    End If
    ' Archive Location Note
    Me.lblArchiveNote.caption = "Note: Select folder for auto-archiving old data" & vbCrLf & _
                                "      Recommended: D:\GLOBAL_SOFT\Archive"
    Me.lblArchiveNote.Font.Size = 10
    Me.lblArchiveNote.ForeColor = RGB(255, 255, 255) ' Gray
End Sub

'===========================================
' LOCK ALL CONTROLS
'===========================================
Private Sub LockAllControls()
    Me.cmbLicenseType.enabled = False
    Me.txtActivationDate.enabled = False
    Me.txtExpiryDate.enabled = False
    Me.txtRemainingDays.enabled = False
    Me.cmbArchivingMode.enabled = False
    Me.chkAutoArchiving.enabled = False
    Me.txtLocation.enabled = False
    Me.btnBrowseLocation.enabled = False
    Me.btnSave.enabled = False
    Me.btnReset.enabled = False
    
    Me.lblStatus.caption = "LOCKED"
    Me.lblStatus.ForeColor = RGB(255, 0, 0)
    Me.btnUnlock.caption = "Unlock"
End Sub

'===========================================
' UNLOCK ALL CONTROLS
'===========================================
Private Sub UnlockAllControls()
    Me.cmbLicenseType.enabled = True
    Me.txtActivationDate.enabled = True
    Me.cmbArchivingMode.enabled = True
    Me.chkAutoArchiving.enabled = True
    Me.txtLocation.enabled = True
    Me.btnBrowseLocation.enabled = True
    Me.btnSave.enabled = True
    Me.btnReset.enabled = True
    
    Me.txtExpiryDate.enabled = False
    Me.txtRemainingDays.enabled = False
    
    Me.lblStatus.caption = "UNLOCKED"
    Me.lblStatus.ForeColor = RGB(0, 128, 0)
    Me.btnUnlock.caption = "Lock"
    isUnlocked = True
End Sub
'===========================================
' UNLOCK BUTTON
'===========================================
Private Sub btnUnlock_Click()
    If Not isUnlocked Then
        Dim pwd As String
        pwd = InputBox("Enter Developer Password:", "Unlock")
        If pwd = "" Then Exit Sub
        
        If pwd = DEV_PASSWORD Then
            UnlockAllControls
            MsgBox "Unlocked!", vbInformation
        Else
            MsgBox "Wrong password!", vbExclamation
        End If
    Else
        LockAllControls
        isUnlocked = False
    End If
End Sub

'===========================================
' LOAD LICENSE TYPES
'===========================================
Private Sub LoadLicenseTypes()
    With Me.cmbLicenseType
        .Clear
        .AddItem "Add New..."
        .AddItem "Demo - 3 Days"
        .AddItem "Demo - 7 Days"
        .AddItem "Trial - 7 Days"
        .AddItem "Trial - 15 Days"
        .AddItem "Trial - 30 Days"
        .AddItem "Basic - 3 Months"
        .AddItem "Basic - 6 Months"
        .AddItem "Standard - 1 Year"
        .AddItem "Standard - 2 Years"
        .AddItem "Premium - 3 Years"
        .AddItem "Premium - 5 Years"
        .AddItem "Lifetime"
    End With
End Sub
'===========================================
' LOAD ARCHIVING MODES
'===========================================
Private Sub LoadArchivingModes()
    With Me.cmbArchivingMode
        .Clear
        .AddItem "Daily"
        .AddItem "Weekly"
        .AddItem "Monthly"
        .AddItem "Yearly"
    End With
End Sub

'===========================================
' LICENSE TYPE CHANGE - WITH ADD NEW
'===========================================
Private Sub cmbLicenseType_Change()
    If Not isUnlocked Then Exit Sub
    
    Dim licenseType As String
    licenseType = Me.cmbLicenseType.value
    
    ' If "Add New..." selected
    If licenseType = "Add New..." Then
        ' Ask for custom details
        Dim customName As String
        Dim durationType As String
        Dim durationValue As Integer
        Dim totalDays As Integer
        
        ' Step 1: License Name
        customName = InputBox("Enter License Name:", "Add Custom License", "Custom License")
        If customName = "" Then
            Me.cmbLicenseType.value = ""
            Exit Sub
        End If
        
        ' Step 2: Duration Type (Days/Months/Years)
        Dim choice As Integer
        choice = MsgBox("Select Duration Type:" & vbCrLf & vbCrLf & _
                       "YES = Days" & vbCrLf & _
                       "NO = Months" & vbCrLf & _
                       "CANCEL = Years", vbYesNoCancel + vbQuestion, "Duration Type")
        
        If choice = vbCancel Then ' Years
            durationValue = CInt(InputBox("Enter number of Years:", "Years", "1"))
            totalDays = durationValue * 365
            
        ElseIf choice = vbNo Then ' Months
            durationValue = CInt(InputBox("Enter number of Months:", "Months", "6"))
            totalDays = durationValue * 30
            
        Else ' Days
            durationValue = CInt(InputBox("Enter number of Days:", "Days", "45"))
            totalDays = durationValue
        End If
        
        If durationValue <= 0 Then
            Me.cmbLicenseType.value = ""
            Exit Sub
        End If
        
        ' Create display text
        Dim displayText As String
        If choice = vbCancel Then
            displayText = customName & " - " & durationValue & " Year"
            If durationValue > 1 Then displayText = displayText & "s"
        ElseIf choice = vbNo Then
            displayText = customName & " - " & durationValue & " Month"
            If durationValue > 1 Then displayText = displayText & "s"
        Else
            displayText = customName & " - " & durationValue & " Day"
            If durationValue > 1 Then displayText = displayText & "s"
        End If
        
        ' Add to list
        Me.cmbLicenseType.RemoveItem 0  ' Remove "Add New..."
        Me.cmbLicenseType.AddItem displayText, 0
        Me.cmbLicenseType.AddItem "Add New...", 1
        Me.cmbLicenseType.value = displayText
        
        ' Calculate dates
        CalculateDates totalDays
        
        MsgBox "Custom License Added:" & vbCrLf & displayText & vbCrLf & _
               "Total Days: " & totalDays, vbInformation
        Exit Sub
    End If
    
    ' For existing types
    Dim days As Integer
    days = GetDaysFromType(licenseType)
    If days > 0 Then CalculateDates days
End Sub

'===========================================
' GET DAYS FROM LICENSE TYPE
'===========================================
Private Function GetDaysFromType(ByVal licType As String) As Integer
    Select Case licType
        Case "Demo - 3 Days": GetDaysFromType = 3
        Case "Demo - 7 Days", "Trial - 7 Days": GetDaysFromType = 7
        Case "Trial - 15 Days": GetDaysFromType = 15
        Case "Trial - 30 Days": GetDaysFromType = 30
        Case "Basic - 3 Months": GetDaysFromType = 90
        Case "Basic - 6 Months": GetDaysFromType = 180
        Case "Standard - 1 Year": GetDaysFromType = 365
        Case "Standard - 2 Years": GetDaysFromType = 730
        Case "Premium - 3 Years": GetDaysFromType = 1095
        Case "Premium - 5 Years": GetDaysFromType = 1825
        Case "Lifetime": GetDaysFromType = 36500
        Case Else: GetDaysFromType = 0
    End Select
End Function

'===========================================
' CALCULATE DATES
'===========================================
Private Sub CalculateDates(ByVal days As Integer)
    Dim actDate As Date
    Dim expDate As Date
    
    ' Get activation date
    If Me.txtActivationDate.value = "" Then
        actDate = Date
        Me.txtActivationDate.value = Format(Date, "dd/mm/yyyy")
    Else
        actDate = CDate(Me.txtActivationDate.value)
    End If
    
    ' Calculate expiry date
    expDate = DateAdd("d", days, actDate)
    Me.txtExpiryDate.value = Format(expDate, "dd/mm/yyyy")
    
    ' Update remaining days
    UpdateRemainingDays
End Sub

'===========================================
' ACTIVATION DATE EXIT
'===========================================
Private Sub txtActivationDate_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Not isUnlocked Then Exit Sub
    
    Dim dt As String
    dt = Trim(Me.txtActivationDate.value)
    If dt = "" Then Exit Sub
    
    ' Format DDMMYYYY to DD/MM/YYYY
    dt = Replace(Replace(dt, "/", ""), "-", "")
    If Len(dt) = 8 And IsNumeric(dt) Then
        Me.txtActivationDate.value = Left(dt, 2) & "/" & Mid(dt, 3, 2) & "/" & Right(dt, 4)
        
        ' Recalculate
        If Me.cmbLicenseType.value <> "" Then
            CalculateDates GetDaysFromType(Me.cmbLicenseType.value)
        End If
    End If
    
    UpdateRemainingDays
End Sub

'===========================================
' UPDATE REMAINING DAYS
'===========================================
Private Sub UpdateRemainingDays()
    If Me.txtExpiryDate.value = "" Then
        Me.txtRemainingDays.value = "0"
        Me.lblStatus.caption = "NOT ACTIVATED"
        Me.lblStatus.ForeColor = RGB(128, 128, 128)
        Exit Sub
    End If
    
    Dim expDate As Date
    Dim daysLeft As Integer
    
    expDate = CDate(Me.txtExpiryDate.value)
    daysLeft = DateDiff("d", Date, expDate)
    
    Me.txtRemainingDays.value = daysLeft
    
    If daysLeft > 30 Then
        Me.lblStatus.caption = "ACTIVE"
        Me.lblStatus.ForeColor = RGB(0, 128, 0)
    ElseIf daysLeft > 0 Then
        Me.lblStatus.caption = "EXPIRING SOON"
        Me.lblStatus.ForeColor = RGB(255, 165, 0)
    Else
        Me.lblStatus.caption = "EXPIRED"
        Me.lblStatus.ForeColor = RGB(255, 0, 0)
        MsgBox "LICENSE EXPIRED!" & vbCrLf & "Contact developer.", vbCritical
    End If
End Sub

'===========================================
' BROWSE LOCATION BUTTON
'===========================================
Private Sub btnBrowseLocation_Click()
    Dim fd As Office.FileDialog
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    
    With fd
        .title = "Select Archive Folder Location"
        .AllowMultiSelect = False
        
        If .Show = -1 Then
            Dim selectedPath As String
            selectedPath = .SelectedItems(1)
            
            ' Auto create Archive subfolder
            selectedPath = selectedPath & "\GLOBAL_SOFT_Archive"
            
            ' Create folder if not exists
            If Dir(selectedPath, vbDirectory) = "" Then
                MkDir selectedPath
            End If
            
            Me.txtLocation.value = selectedPath
            
            MsgBox "Archive folder created at:" & vbCrLf & selectedPath, vbInformation
        End If
    End With
End Sub

'===========================================
' LOAD LICENSE SETTINGS - FIXED
'===========================================
Private Sub LoadLicenseSettings()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Software_Config")
    On Error GoTo 0
    
    If ws Is Nothing Then Exit Sub
    
    ' License Type
    If ws.Range("B31").value <> "" Then
        Me.cmbLicenseType.value = CStr(ws.Range("B31").value)
    End If
    
    ' Activation Date
    If ws.Range("B32").value = "" Then
        Me.txtActivationDate.value = Format(Date, "dd/mm/yyyy")
    Else
        Me.txtActivationDate.value = CStr(ws.Range("B32").value)
    End If
    
    ' Expiry Date
    If ws.Range("B33").value <> "" Then
        Me.txtExpiryDate.value = CStr(ws.Range("B33").value)
    End If
    
    ' Remaining Days
    If ws.Range("B35").value <> "" Then
        Me.txtRemainingDays.value = CStr(ws.Range("B35").value)
    End If
    
    ' STATUS (B36)
    ' Auto Archiving CheckBox (B37) - YES/NO
    Me.chkAutoArchiving.value = (CStr(ws.Range("B37").value) = "Yes")
    
    ' Archiving Mode ComboBox (B44) - Daily/Weekly/Monthly/Yearly
    If ws.Range("B44").value <> "" Then
        Me.cmbArchivingMode.value = CStr(ws.Range("B44").value)
    End If
    
    ' Location
    If ws.Range("B38").value <> "" Then
        Me.txtLocation.value = CStr(ws.Range("B38").value)
    End If
    
    ' Calculate if dates exist
    If Me.txtExpiryDate.value <> "" Then
        UpdateRemainingDays
    End If
End Sub

'===========================================
' SAVE BUTTON - FIXED
'===========================================
Private Sub btnSave_Click()
    If Not isUnlocked Then
        MsgBox "Please unlock first!", vbExclamation
        Exit Sub
    End If
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Software_Config")
    
    ' Save all fields
    ws.Range("B31").value = Me.cmbLicenseType.value      ' License_Type
    ws.Range("B32").value = Me.txtActivationDate.value   ' Activation_Date
    ws.Range("B33").value = Me.txtExpiryDate.value       ' Expiry_Date
    ws.Range("B34").value = ""                           ' Renew_Date (auto)
    ws.Range("B35").value = Me.txtRemainingDays.value    ' Days_Remaining
    ws.Range("B36").value = Me.lblStatus.caption         ' Status
    
    ' B37 = Auto Archiving (Yes/No) - Checkbox
    ws.Range("B37").value = IIf(Me.chkAutoArchiving.value, "Yes", "No")
    
    ' B44 = Archiving Mode (Daily/Weekly/Monthly/Yearly) - ComboBox
    ws.Range("B44").value = Me.cmbArchivingMode.value
    
    ' B38 = Archive Location
    ws.Range("B38").value = Me.txtLocation.value
    
    MsgBox "License settings saved!" & vbCrLf & _
           "Auto Archiving: " & IIf(Me.chkAutoArchiving.value, "Yes", "No") & vbCrLf & _
           "Archiving Mode: " & Me.cmbArchivingMode.value, vbInformation
End Sub

'===========================================
' RESET BUTTON - FIXED
'===========================================
Private Sub btnReset_Click()
    If Not isUnlocked Then
        MsgBox "Please unlock first!", vbExclamation
        Exit Sub
    End If
    
    If MsgBox("Reset all?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    
    ' Clear fields
    Me.cmbLicenseType.value = ""
    Me.txtActivationDate.value = ""
    Me.txtExpiryDate.value = ""
    Me.txtRemainingDays.value = "0"
    Me.cmbArchivingMode.value = ""
    Me.chkAutoArchiving.value = False
    Me.txtLocation.value = ""
    Me.lblStatus.caption = "NOT ACTIVATED"
    Me.lblStatus.ForeColor = RGB(128, 128, 128)
    
    MsgBox "Reset completed!", vbInformation
End Sub

'===========================================
' CLOSE BUTTON
'===========================================
Private Sub btnClose_Click()
    Unload Me
End Sub
'===========================================
' CHECK SYSTEM DATE TAMPERING
'===========================================
Private Function IsDateTampered() As Boolean
    Dim ws As Worksheet
    Dim lastRunDate As Date
    Dim today As Date
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Software_Config")
    On Error GoTo 0
    
    If ws Is Nothing Then
        IsDateTampered = False
        Exit Function
    End If
    
    ' Get last run date from hidden storage
    lastRunDate = CDate(ws.Range("B32").value) ' Hidden cell for last run
    today = Date
    
    ' Check if system date is before last run date
    If today < lastRunDate Then
        ' Date tampering detected!
        IsDateTampered = True
        
        MsgBox "SYSTEM DATE TAMPERING DETECTED!" & vbCrLf & vbCrLf & _
               "Last run date: " & Format(lastRunDate, "dd/mm/yyyy") & vbCrLf & _
               "Current date: " & Format(today, "dd/mm/yyyy") & vbCrLf & vbCrLf & _
               "Please correct your system date." & vbCrLf & _
               "Contact: support@globalsoft.com", vbCritical, "Security Alert"
    Else
        ' Update last run date
        ws.Range("B32").value = today
        IsDateTampered = False
    End If
End Function
'===========================================
' GET INTERNET TIME (NTP)
'===========================================
Private Function GetInternetTime() As Date
    On Error Resume Next
    
    Dim xmlhttp As Object
    Dim response As String
    Dim serverTime As Date
    
    Set xmlhttp = CreateObject("MSXML2.XMLHTTP")
    
    ' Use worldtimeapi.org
    xmlhttp.Open "GET", "http://worldtimeapi.org/api/ip", False
    xmlhttp.Send
    
    If xmlhttp.status = 200 Then
        response = xmlhttp.responseText
        
        ' Parse datetime from JSON response
        Dim datetimeStr As String
        Dim pos As Integer
        
        pos = InStr(response, """datetime"":""")
        If pos > 0 Then
            datetimeStr = Mid(response, pos + 12, 19) ' Extract YYYY-MM-DDTHH:MM:SS
            serverTime = CDate(Replace(Left(datetimeStr, 10), "-", "/"))
            GetInternetTime = serverTime
            Exit Function
        End If
    End If
    
    ' If internet fails, return system date
    GetInternetTime = Date
    
    Set xmlhttp = Nothing
End Function

'===========================================
' CHECK WITH INTERNET TIME
'===========================================
Private Function IsLicenseValid() As Boolean
    Dim internetDate As Date
    Dim systemDate As Date
    Dim expiryDate As Date
    
    systemDate = Date
    internetDate = GetInternetTime()
    expiryDate = CDate(Me.txtExpiryDate.value)
    
    ' Check if system date matches internet date (±1 day allowed)
    If Abs(systemDate - internetDate) > 1 Then
        MsgBox "System date does not match internet time!" & vbCrLf & _
               "Please check your system date.", vbExclamation
        IsLicenseValid = False
        Exit Function
    End If
    
    ' Check expiry
    If systemDate > expiryDate Then
        MsgBox "License Expired!", vbCritical
        IsLicenseValid = False
        Exit Function
    End If
    
    IsLicenseValid = True
End Function

'===========================================
' SET DEFAULT ARCHIVE LOCATION
'===========================================
Private Sub SetDefaultArchiveLocation()
    Dim defaultPath As String
    defaultPath = Environ("USERPROFILE") & "\Documents\GLOBAL_SOFT_Archive"
    
    If Dir(defaultPath, vbDirectory) = "" Then
        On Error Resume Next
        MkDir defaultPath
        On Error GoTo 0
    End If
    
    Me.txtLocation.value = defaultPath
End Sub
