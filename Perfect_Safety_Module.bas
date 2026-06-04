Attribute VB_Name = "Perfect_Safety_Module"
Option Explicit

'=================== CONSTANTS (????? Interval, Email ????) ===================
Public Const BACKUP_INTERVAL As String = "06:00:00"    '6 ????
Public Const RETENTION_DAYS As Long = 30               '30 ???
Public Const AUTO_SAVE_INTERVAL As String = "00:00:30" '30 ?????

Public NextBackupTime As Date
Public NextAutoSaveTime As Date

'=================== MAIN CONTROL FUNCTIONS ===================

'?? FILE OPEN ?? ?? ?????
Public Sub SystemStartup()
    On Error Resume Next
    
    EnsureAllFolders
    CreateBackup "OPEN"
    
    '??? User ?? Cloud Enable ???? ?? ?? ?? ????
    If IsCloudEnabled Then
        SendCloudBackup "OPEN"
    End If
    
    CleanupOldBackups
    ScheduleNextBackup
    StartAutoSaveTimer
    CompleteDatabaseRepair
    
    LogAction "SYSTEM_STARTED | Cloud=" & IIf(IsCloudEnabled, "ON", "OFF")
    On Error GoTo 0
End Sub

'?? FILE CLOSE ?? ?? ?????
Public Sub SystemShutdown()
    On Error Resume Next
    
    CancelScheduledBackup
    StopAutoSaveTimer
    CreateBackup "CLOSE"
    
    If IsCloudEnabled Then
        SendCloudBackup "CLOSE"
    End If
    
    ThisWorkbook.Save
    LogAction "SYSTEM_CLOSED"
    On Error GoTo 0
End Sub

'?? ?? 6 ???? ??? AUTO ?????
Public Sub ScheduledBackupRunner()
    On Error Resume Next
    CreateBackup "AUTO_6HR"
    
    If IsCloudEnabled Then
        SendCloudBackup "AUTO_6HR"
    End If
    
    CleanupOldBackups
    ScheduleNextBackup
    On Error GoTo 0
End Sub

'=================== LOCAL BACKUP (ZIP) ===================

Private Sub CreateBackup(TriggerType As String)
    Dim fso As Object, shell As Object
    Dim dataPath As String, backupPath As String
    Dim tempFolder As String, zipFile As String
    Dim timeStamp As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    dataPath = ThisWorkbook.path & "\GLOBAL_SOFT_DATA\"
    backupPath = dataPath & "AutoBackup\"
    timeStamp = Format(Now, "yyyymmdd_hhmmss")
    
    'Folder ????
    If Not fso.FolderExists(backupPath) Then fso.CreateFolder backupPath
    
    'Temp ??? Copy ???
    tempFolder = backupPath & "Temp_" & timeStamp & "\"
    fso.CreateFolder tempFolder
    
    ThisWorkbook.SaveCopyAs tempFolder & ThisWorkbook.name
    
    If fso.FolderExists(dataPath & "Images") Then
        fso.CopyFolder dataPath & "Images", tempFolder & "Images"
    End If
    If fso.FolderExists(dataPath & "Reports") Then
        fso.CopyFolder dataPath & "Reports", tempFolder & "Reports"
    End If
    
    'ZIP ????
    zipFile = backupPath & "GlobalSoft_" & timeStamp & "_" & TriggerType & ".zip"
    Set shell = CreateObject("WScript.Shell")
    shell.Run "powershell -command ""Compress-Archive -Path '" & tempFolder & "*' -DestinationPath '" & zipFile & "' -Force""", 0, True
    
    fso.DeleteFolder tempFolder, True
    LogAction "BACKUP_CREATED: " & fso.GetFileName(zipFile)
    
    Set fso = Nothing
    Set shell = Nothing
End Sub

'=================== CLOUD BACKUP (Settings ?? ???? ??) ===================

Private Sub SendCloudBackup(TriggerType As String)
    On Error GoTo CloudError
    Dim cdoConfig As Object, cdoMsg As Object
    Dim fso As Object, backupPath As String, LatestFile As String
    Dim gmailUser As String, gmailPass As String
    
    'Settings ?? ??
    gmailUser = GetSettingValue("Gmail_User")
    gmailPass = DecryptPassword(GetSettingValue("Gmail_Password"))
    
    If gmailUser = "" Or gmailPass = "" Then
        LogError "Cloud Failed: No credentials"
        Exit Sub
    End If
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    backupPath = ThisWorkbook.path & "\GLOBAL_SOFT_DATA\AutoBackup\"
    LatestFile = GetLatestBackupFile(backupPath)
    If LatestFile = "" Then Exit Sub
    
    'Email ????
    Set cdoConfig = CreateObject("CDO.Configuration")
    With cdoConfig.Fields
        .item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.gmail.com"
        .item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 465
        .item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
        .item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
        .item("http://schemas.microsoft.com/cdo/configuration/sendusername") = gmailUser
        .item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = gmailPass
        .item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
        .Update
    End With
    
    Set cdoMsg = CreateObject("CDO.Message")
    With cdoMsg
        Set .Configuration = cdoConfig
        .To = gmailUser
        .From = gmailUser
        .subject = "GlobalSoft Backup [" & TriggerType & "] " & Format(Now, "dd-mmm hh:mm")
        .TextBody = "Backup at " & Now & " from " & Environ("COMPUTERNAME")
        .AddAttachment LatestFile
        .Send
    End With
    
    SaveSettingValue "Cloud_LastBackup", Now
    LogAction "CLOUD_SENT: " & TriggerType
    
    Set cdoMsg = Nothing
    Set cdoConfig = Nothing
    Exit Sub
    
CloudError:
    LogError "Cloud Failed: " & Err.Description
End Sub

'=================== SETTINGS & HELPERS (???) ===================

Public Function IsCloudEnabled() As Boolean
    IsCloudEnabled = (GetSettingValue("Cloud_Enabled") = "TRUE")
End Function

Public Function GetSettingValue(key As String) As String
    Dim ws As Worksheet, f As Range
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Settings")
    Set f = ws.Range("A:A").Find(key, LookAt:=xlWhole)
    If Not f Is Nothing Then GetSettingValue = CStr(ws.Cells(f.row, 2).value)
    On Error GoTo 0
End Function

Public Sub SaveSettingValue(key As String, value As Variant)
    Dim ws As Worksheet, f As Range, lastRow As Long
    Set ws = ThisWorkbook.Sheets("Settings")
    Set f = ws.Range("A:A").Find(key, LookAt:=xlWhole)
    
    If f Is Nothing Then
        lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row + 1
        ws.Cells(lastRow, 1).value = key
        ws.Cells(lastRow, 2).value = value
    Else
        ws.Cells(f.row, 2).value = value
    End If
End Sub

'Simple Encryption (ASCII × 3)
Public Function EncryptPassword(pwd As String) As String
    Dim i As Integer, result As String
    For i = 1 To Len(pwd)
        result = result & Asc(Mid(pwd, i, 1)) * 3 & "-"
    Next i
    EncryptPassword = result
End Function

Public Function DecryptPassword(enc As String) As String
    Dim arr() As String, i As Integer, result As String
    If enc = "" Then DecryptPassword = "": Exit Function
    arr = Split(enc, "-")
    For i = LBound(arr) To UBound(arr) - 1
        If IsNumeric(arr(i)) Then result = result & Chr(Int(CInt(arr(i)) / 3))
    Next i
    DecryptPassword = result
End Function

'=================== TIMER & CLEANUP ===================

Private Sub ScheduleNextBackup()
    NextBackupTime = Now + TimeValue(BACKUP_INTERVAL)
    Application.OnTime NextBackupTime, "ScheduledBackupRunner"
    LogAction "NEXT_BACKUP: " & Format(NextBackupTime, "dd-mm hh:mm")
End Sub

Private Sub CancelScheduledBackup()
    On Error Resume Next
    If NextBackupTime > 0 Then
        Application.OnTime NextBackupTime, "ScheduledBackupRunner", , False
    End If
End Sub

Private Sub CleanupOldBackups()
    Dim fso As Object, folder As Object, file As Object
    Dim backupPath As String, cutoffDate As Date
    Dim deletedCount As Long
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    backupPath = ThisWorkbook.path & "\GLOBAL_SOFT_DATA\AutoBackup\"
    If Not fso.FolderExists(backupPath) Then Exit Sub
    
    Set folder = fso.GetFolder(backupPath)
    cutoffDate = Now - RETENTION_DAYS
    
    For Each file In folder.files
        If LCase(fso.GetExtensionName(file.name)) = "zip" Then
            If file.DateCreated < cutoffDate Then
                file.Delete True
                deletedCount = deletedCount + 1
                LogAction "DELETED_OLD: " & file.name
            End If
        End If
    Next
    
    If deletedCount > 0 Then LogAction "CLEANUP: " & deletedCount & " files"
    Set fso = Nothing
End Sub

Private Sub StartAutoSaveTimer()
    NextAutoSaveTime = Now + TimeValue(AUTO_SAVE_INTERVAL)
    Application.OnTime NextAutoSaveTime, "AutoSaveRunner"
End Sub

Private Sub StopAutoSaveTimer()
    On Error Resume Next
    If NextAutoSaveTime > 0 Then
        Application.OnTime NextAutoSaveTime, "AutoSaveRunner", , False
    End If
End Sub

Public Sub AutoSaveRunner()
    On Error Resume Next
    ThisWorkbook.Save
    StartAutoSaveTimer
End Sub

Private Function GetLatestBackupFile(folderPath As String) As String
    Dim fso As Object, folder As Object, file As Object
    Dim latestDate As Date, LatestFile As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folderPath) Then Exit Function
    
    Set folder = fso.GetFolder(folderPath)
    latestDate = CDate("01-01-1900")
    
    For Each file In folder.files
        If LCase(fso.GetExtensionName(file.name)) = "zip" Then
            If file.DateLastModified > latestDate Then
                latestDate = file.DateLastModified
                LatestFile = file.path
            End If
        End If
    Next
    
    GetLatestBackupFile = LatestFile
    Set fso = Nothing
End Function

'=================== DATABASE REPAIR (??? ?????) ===================

Public Sub CompleteDatabaseRepair()
    On Error Resume Next
    RepairSheet "Customer_Master", Array("CustomerID", "CustomerName", "Mobile", "Email", "Address", "City", "State", "Pincode", "PhotoPath", "CreateDate", "LastUpdate")
    RepairSheet "Product_Master", Array("ProductID", "ProductName", "Company", "Model", "Category", "Description", "Price")
    RepairSheet "Accessory_Master", Array("AccessoryID", "AccessoryName", "Category", "Price", "Stock")
    RepairSheet "Job_Master", Array("EntryID", "CustomerID", "EntryDate", "Status", "LastUpdate", "Priority", "AssignedTo", "Problem", "Solution")
    RepairSheet "Job_Product", Array("EntryID", "CustomerID", "EntryType", "Status", "LastStatusDate", "ProductType", "Company", "Model", "SerialNumber", "PurchaseDate", "WarrantyYear", "WarrantyStatus", "Problem", "Remark", "Picture", "CreateDate", "VerifyType", "VerifyName", "TotalAmount", "ReceivedAmount", "PendingAmount", "PaymentMode", "ReceivedBy")
    RepairSheet "Job_Accessory", Array("EntryID", "AccessoryName", "Status", "SerialNo", "WorkingStatus", "PhotoPath", "Price")
    RepairSheet "Payment_Master", Array("PaymentID", "EntryID", "CustomerID", "TransType", "TransMode", "Amount", "TransDate", "Remark", "CreatedBy")
    RepairSheet "Engineer_Register", Array("EngineerID", "Name", "Mobile", "Email", "Specialization", "Status", "JoinDate")
    RepairSheet "Vendor_Register", Array("VendorID", "Name", "Mobile", "Email", "Company", "Status", "JoinDate")
    RepairSheet "Settings", Array("SettingName", "Value", "Description")
    SetDefaultSettings
    LogAction "DATABASE_REPAIR_DONE"
    On Error GoTo 0
End Sub

Private Sub RepairSheet(sheetName As String, headers As Variant)
    Dim ws As Worksheet, i As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
        ws.name = sheetName
        LogAction "CREATED_SHEET: " & sheetName
    End If
    
    If ws.Cells(1, 1).value = "" Then
        For i = LBound(headers) To UBound(headers)
            ws.Cells(1, i + 1).value = headers(i)
        Next i
        With ws.Rows(1)
            .Font.Bold = True
            .Interior.Color = RGB(68, 114, 196)
            .Font.Color = RGB(255, 255, 255)
        End With
        LogAction "REPAIRED_HEADERS: " & sheetName
    End If
End Sub

Private Sub SetDefaultSettings()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Settings")
    If ws.Range("B1").value = "" Then ws.Range("B1").value = ThisWorkbook.path & "\GLOBAL_SOFT_DATA\Images"
    If ws.Range("B2").value = "" Then ws.Range("B2").value = ThisWorkbook.path & "\GLOBAL_SOFT_DATA\Images\Customers"
    If ws.Range("B3").value = "" Then ws.Range("B3").value = ThisWorkbook.path & "\GLOBAL_SOFT_DATA\Images\Products"
End Sub

Private Sub EnsureAllFolders()
    Dim fso As Object, base As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    base = ThisWorkbook.path & "\GLOBAL_SOFT_DATA\"
    If Not fso.FolderExists(base) Then fso.CreateFolder base
    If Not fso.FolderExists(base & "Images") Then fso.CreateFolder base & "Images"
    If Not fso.FolderExists(base & "Images\Customers") Then fso.CreateFolder base & "Images\Customers"
    If Not fso.FolderExists(base & "Images\Products") Then fso.CreateFolder base & "Images\Products"
    If Not fso.FolderExists(base & "Images\Accessories") Then fso.CreateFolder base & "Images\Accessories"
    If Not fso.FolderExists(base & "AutoBackup") Then fso.CreateFolder base & "AutoBackup"
    If Not fso.FolderExists(base & "Reports") Then fso.CreateFolder base & "Reports"
    Set fso = Nothing
End Sub

Public Sub LogAction(msg As String)
    On Error Resume Next
    Dim fNum As Integer, logFile As String
    logFile = ThisWorkbook.path & "\GLOBAL_SOFT_DATA\activity.log"
    fNum = FreeFile
    Open logFile For Append As #fNum
    Print #fNum, Format(Now, "yyyy-mm-dd hh:mm:ss") & " | " & msg
    Close #fNum
End Sub

Public Sub LogError(errorSource As String, Optional errorDesc As String = "")
    If errorDesc = "" Then
        LogAction "ERROR: " & errorSource
    Else
        LogAction "ERROR in " & errorSource & ": " & errorDesc
    End If
End Sub
Sub FixControlNames()
    '?? Code Module ??? Run ???, Form ??? ????
    Dim ctrl As Object
    For Each ctrl In frmBackup.Controls
        MsgBox "Control Name: " & ctrl.name & " Type: " & TypeName(ctrl)
    Next
End Sub
' Perfect_Safety_Module ?? End ?? ????:
Public Function IsFeatureEnabled(featureName As String) As Boolean
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Settings")
    
    Select Case UCase(featureName)
        Case "PHOTOS"
            IsFeatureEnabled = (UCase(ws.Range("C16").value) = "TRUE")
        Case "CLOUD"
            IsFeatureEnabled = (UCase(ws.Range("C17").value) = "TRUE")
        Case "PAYMENT"
            IsFeatureEnabled = (UCase(ws.Range("C18").value) = "TRUE")
        Case "DEVMODE"
            IsFeatureEnabled = (UCase(ws.Range("C19").value) = "TRUE")
        Case Else
            IsFeatureEnabled = False
    End Select
    On Error GoTo 0
End Function
Public Sub ProtectSheets()
    ' Data rows ?? Protect ????, ????? Headers ?? Allow ????
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Sheets
        If ws.name Like "*Master*" Or ws.name Like "*Job*" Then
            ws.Protect UserInterfaceOnly:=True
        End If
    Next
End Sub
