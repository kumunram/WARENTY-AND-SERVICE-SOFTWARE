VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmBackup 
   Caption         =   "Database Backup"
   ClientHeight    =   11430
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9765.001
   OleObjectBlob   =   "frmBackup.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmBackup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const BACKUP_PREFIX As String = "GITS_BACKUP_"
Private Const KEEP_BACKUPS As Integer = 5

'========================================
' FORM LOAD
'========================================
Private Sub UserForm_Initialize()
    Call CreateMainDataFolders
    AddMinMaxButtons Me
    Me.StartUpPosition = 1
    
    On Error Resume Next
    txtBackupLocation.value = CStr(ThisWorkbook.Sheets("Settings").Range("B10").value)
    If txtBackupLocation.value = "" Then
        txtBackupLocation.value = ThisWorkbook.path & "\WARRENTY AND SERVICE BACKUP\"
    End If
    On Error GoTo 0
    
    UpdateReferenceCheckboxes
    LoadCloudSettings
End Sub

Private Sub LoadCloudSettings()
    On Error Resume Next
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("Settings")
    Dim gmail As String, enabled As String, pass As String
    
    gmail = CStr(ws.Range("B12").value)
    pass = CStr(ws.Range("B13").value)
    enabled = UCase(CStr(ws.Range("B14").value))
    
    If gmail <> "" And gmail <> "Empty" Then txtGmailID.value = gmail
    
    If pass <> "" And pass <> "Empty" Then
        txtPassword.value = "****************"
        txtPassword.Tag = "SAVED:" & CleanPasswordChars(pass)
    End If
    
    chkEnableCloud.value = (enabled = "TRUE")
    
    If gmail <> "" And enabled = "TRUE" Then
        lblCloudStatus.caption = "Gmail Configured & Enabled"
        lblCloudStatus.ForeColor = RGB(0, 128, 0)
    ElseIf gmail <> "" Then
        lblCloudStatus.caption = "Gmail Configured (Disabled)"
        lblCloudStatus.ForeColor = RGB(255, 165, 0)
    Else
        lblCloudStatus.caption = "Gmail Not Configured"
        lblCloudStatus.ForeColor = RGB(255, 0, 0)
    End If
    On Error GoTo 0
End Sub

'========================================
' AUTO-CREATE DATA FOLDERS
'========================================
Private Sub CreateMainDataFolders()
    On Error Resume Next
    Dim basePath As String
    basePath = ThisWorkbook.path & "\GLOBAL_SOFT_DATA\"
    
    If Dir(basePath, vbDirectory) = "" Then MkDir basePath
    If Dir(basePath & "Product\", vbDirectory) = "" Then MkDir basePath & "Product\"
    If Dir(basePath & "Accessories\", vbDirectory) = "" Then MkDir basePath & "Accessories\"
    If Dir(basePath & "Customers\", vbDirectory) = "" Then MkDir basePath & "Customers\"
    If Dir(basePath & "Logo\", vbDirectory) = "" Then MkDir basePath & "Logo\"
    If Dir(basePath & "Warranty\", vbDirectory) = "" Then MkDir basePath & "Warranty\"
    If Dir(basePath & "Delivery\", vbDirectory) = "" Then MkDir basePath & "Delivery\"
    If Dir(basePath & "Staff\", vbDirectory) = "" Then MkDir basePath & "Staff\"
    On Error GoTo 0
End Sub

'========================================
' PASSWORD HELPER
'========================================
Private Function CleanPasswordChars(ByVal raw As String) As String
    Dim cleaned As String, i As Integer, c As String
    cleaned = ""
    For i = 1 To Len(raw)
        c = Mid(raw, i, 1)
        If (c >= "A" And c <= "Z") Or (c >= "a" And c <= "z") Or (c >= "0" And c <= "9") Then
            cleaned = cleaned & c
        End If
    Next i
    CleanPasswordChars = cleaned
End Function

Private Function GetCleanPassword() As String
    Dim raw As String
    If Left(txtPassword.Tag, 6) = "SAVED:" Then
        raw = Mid(txtPassword.Tag, 7)
    ElseIf txtPassword.value = "****************" Or txtPassword.value = "???????????????" Then
        On Error Resume Next
        raw = CStr(ThisWorkbook.Sheets("Settings").Range("B13").value)
        On Error GoTo 0
    Else
        raw = txtPassword.value
    End If
    GetCleanPassword = CleanPasswordChars(raw)
End Function

Private Sub UpdateReferenceCheckboxes()
    On Error Resume Next
    Dim hasCC As Boolean, hasCDO As Boolean, ref As Object
    For Each ref In Application.VBE.ActiveVBProject.References
        If InStr(ref.name, "MSComctlLib") > 0 Then hasCC = True
        If InStr(ref.name, "CDO") > 0 Then hasCDO = True
    Next ref
    chkCommonControls.value = hasCC
    chkCDO.value = hasCDO
    On Error GoTo 0
End Sub

'========================================
' MAIN BACKUP WIZARD
'========================================
Private Sub btnBackupNow_Click()
    Dim destPath As String
    destPath = Trim(txtBackupLocation.value)
    
    If destPath = "" Then
        MsgBox "Backup location empty! Click 'Select Location' first.", vbExclamation
        Exit Sub
    End If
    
    On Error Resume Next
    If Dir(destPath, vbDirectory) = "" Then CreateFolderRecursive destPath
    On Error GoTo 0
    
    If Dir(destPath, vbDirectory) = "" Then
        MsgBox "Cannot create backup folder: " & destPath, vbCritical
        Exit Sub
    End If
    
    Dim includePhotos As Boolean
    Dim zipFile As String
    Dim t0 As Double
    
    If MsgBox("Include PHOTOS in backup?" & vbCrLf & vbCrLf & _
              "WITHOUT Photos = FAST (database only)" & vbCrLf & _
              "WITH Photos = SLOW (complete backup with all photos)" & vbCrLf & vbCrLf & _
              "For complete restore, choose YES.", vbYesNo + vbQuestion, "Complete Backup?") = vbYes Then
        includePhotos = True
    End If
    
    Application.StatusBar = "Creating complete backup (please wait)..."
    t0 = Timer
    zipFile = CreateCompleteBackup(destPath, includePhotos)
    Application.StatusBar = False
    
    If zipFile = "" Then
        MsgBox "BACKUP FAILED!" & vbCrLf & "Check disk space, permissions & photo folder path.", vbCritical
        Exit Sub
    End If
    
    MsgBox "LOCAL BACKUP SUCCESS!" & vbCrLf & vbCrLf & _
           "File: " & GetFileName(zipFile) & vbCrLf & _
           "Size: " & GetFileSizeText(zipFile) & vbCrLf & _
           "Time: " & Format(Timer - t0, "0.0") & " sec" & vbCrLf & _
           "Contains: Database + " & IIf(includePhotos, "All Photos", "No Photos") & vbCrLf & _
           "Folder: " & destPath, vbInformation, "Backup Complete"
    
    CleanOldBackups destPath
    
    If chkEnableCloud.value Then
        If MsgBox("Send backup to GMAIL?" & vbCrLf & vbCrLf & _
                  "To: " & txtGmailID.value, vbYesNo + vbQuestion, "Gmail Backup?") = vbYes Then
            Call DoGmailBackup(zipFile)
        End If
    Else
        If MsgBox("Cloud backup is DISABLED." & vbCrLf & "Enable in settings or send anyway?" & vbCrLf & vbCrLf & _
                  "Send to Gmail: " & txtGmailID.value, vbYesNo + vbQuestion, "Gmail Backup?") = vbYes Then
            Call DoGmailBackup(zipFile)
        End If
    End If
    
    If MsgBox("Upload to GOOGLE DRIVE?" & vbCrLf & _
              "(Requires Google Drive Desktop installed)", vbYesNo + vbQuestion, "Drive Backup?") = vbYes Then
        Call DoDriveBackup(zipFile)
    End If
    
    MsgBox "ALL BACKUP STEPS COMPLETE!" & vbCrLf & vbCrLf & _
           "Local folder keeps 5 latest backups." & vbCrLf & _
           "Older backups auto-deleted." & vbCrLf & vbCrLf & _
           "Your backup contains EVERYTHING needed for restore.", vbInformation, "Done"
End Sub

'========================================
' COMPLETE BACKUP ZIP (Database + Photos)
'========================================
Private Function CreateCompleteBackup(ByVal destFolder As String, ByVal includePhotos As Boolean) As String
    On Error GoTo Failed
    
    Dim ts As String
    ts = Format(Now, "yyyy-mm-dd_hh-mm-ss")
    
    Dim zipPath As String
    zipPath = destFolder & "\" & BACKUP_PREFIX & "COMPLETE_" & ts & ".zip"
    
    Dim tempRoot As String, tempDB As String, tempPhotos As String
    tempRoot = Environ("TEMP") & "\GITS_BACKUP_" & ts
    tempDB = tempRoot & "\Database"
    tempPhotos = tempRoot & "\Photos"
    
    On Error Resume Next
    MkDir tempRoot
    MkDir tempDB
    If includePhotos Then MkDir tempPhotos
    On Error GoTo 0
    
    ' Copy Excel file to Database folder
    Dim tempExcel As String
    tempExcel = tempDB & "\GLOBAL_SOFT.xlsm"
    
    Application.ScreenUpdating = False
    ThisWorkbook.SaveCopyAs tempExcel
    Application.ScreenUpdating = True
    
    ' Copy photos if requested — USING ROBOCOPY (better than xcopy)
    If includePhotos Then
        Dim photoRoot As String
        photoRoot = FindPhotosFolder
        
        If photoRoot <> "" Then
            Dim wsh As Object
            Set wsh = CreateObject("WScript.Shell")
            
            ' Trim trailing backslash to avoid quote escaping issues
            Dim photoRootClean As String
            photoRootClean = photoRoot
            If Right(photoRootClean, 1) = "\" Then photoRootClean = Left(photoRootClean, Len(photoRootClean) - 1)
            
            ' ROBOCOPY: faster, reliable, unicode support, long path support
            Dim robocopyCmd As String
            robocopyCmd = "robocopy """ & photoRootClean & """ """ & tempPhotos & """ /E /R:1 /W:1"
            Dim rc As Long
rc = wsh.Run(robocopyCmd, 0, True)

If rc >= 8 Then
    MsgBox "Photo backup failed during ROBOCOPY!", vbCritical
End If
            Set wsh = Nothing
        End If
    End If
    
    ' ZIP the entire tempRoot
    Dim wsh2 As Object
    Set wsh2 = CreateObject("WScript.Shell")
    
    Dim psCmd As String
    psCmd = "powershell -NoProfile -ExecutionPolicy Bypass -Command """ & _
            "if(Test-Path '" & zipPath & "'){Remove-Item '" & zipPath & "' -Force}; " & _
            "Compress-Archive -Path '" & tempRoot & "\*' -DestinationPath '" & zipPath & "' -Force"""
    
    wsh2.Run psCmd, 0, True
    
    ' Cleanup temp
    On Error Resume Next
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.DeleteFolder tempRoot, True
    Set fso = Nothing
    Set wsh2 = Nothing
    On Error GoTo 0
    
    If Dir(zipPath) <> "" Then
        CreateCompleteBackup = zipPath
    Else
        GoTo Failed
    End If
    
    Exit Function
    
Failed:
    CreateCompleteBackup = ""
    On Error Resume Next
    Set wsh = Nothing
    Set wsh2 = Nothing
    Set fso = Nothing
    On Error GoTo 0
End Function

'========================================
' FIND PHOTOS FOLDER
'========================================
Private Function FindPhotosFolder() As String
    On Error Resume Next
    Dim basePath As String
    basePath = ThisWorkbook.path
    
    ' PRIMARY: GLOBAL_SOFT_DATA (all subfolders inside)
    If Dir(basePath & "\GLOBAL_SOFT_DATA\", vbDirectory) <> "" Then
        FindPhotosFolder = basePath & "\GLOBAL_SOFT_DATA\"
        Exit Function
    End If
    
    ' FALLBACK: Old GLOBAL_SOFT_Photos structure
    If Dir(basePath & "\GLOBAL_SOFT_Photos\", vbDirectory) <> "" Then
        FindPhotosFolder = basePath & "\GLOBAL_SOFT_Photos\"
        Exit Function
    End If
    
    ' FALLBACK: Other common paths
    Dim paths(1 To 4) As String
    paths(1) = basePath & "\Photos\"
    paths(2) = basePath & "\Images\"
    paths(3) = basePath & "\Data\Photos\"
    paths(4) = basePath & "\GLOBAL_SOFT_DATA\Images\"
    
    Dim i As Integer
    For i = 1 To 4
        If Dir(paths(i), vbDirectory) <> "" Then
            FindPhotosFolder = paths(i)
            Exit Function
        End If
    Next i
    
    FindPhotosFolder = ""
    On Error GoTo 0
End Function

'========================================
' GMAIL BACKUP — WITH SIZE CHECK
'========================================
Private Sub DoGmailBackup(ByVal zipFile As String)
    If Not chkCDO.value Then
        MsgBox "CDO Library missing!" & vbCrLf & "Tick 'Microsoft CDO for Windows 2000 Library' checkbox and install it.", vbCritical
        Exit Sub
    End If
    
    Dim gmail As String, pass As String
    gmail = Trim(txtGmailID.value)
    pass = GetCleanPassword()
    
    If gmail = "" Then
        MsgBox "Gmail ID not set! Configure in Cloud Settings first.", vbExclamation
        Exit Sub
    End If
    
    If Len(pass) <> 16 Then
        MsgBox "App Password must be exactly 16 characters!" & vbCrLf & _
               "Current: " & Len(pass) & " chars" & vbCrLf & vbCrLf & _
               "Detected password: '" & pass & "'" & vbCrLf & vbCrLf & _
               "FIX: Re-enter 16-digit App Password (no spaces) and click Save Settings.", vbExclamation
        Exit Sub
    End If
    
    ' Gmail size limit check
    Dim zipMB As Double
    zipMB = FileLen(zipFile) / 1024 / 1024
    If zipMB > 24 Then
        MsgBox "Backup too large for Gmail!" & vbCrLf & _
               "Size: " & Format(zipMB, "0.0") & " MB" & vbCrLf & vbCrLf & _
               "Gmail limit is 25 MB. Use Google Drive backup instead.", vbExclamation
        Exit Sub
    End If
    
    Application.StatusBar = "Sending backup to Gmail..."
    Dim t0 As Double: t0 = Timer
    Dim ok As Boolean
    ok = SendBackupEmail(gmail, pass, zipFile, "Complete Backup")
    Application.StatusBar = False
    
    If ok Then
        MsgBox "GMAIL BACKUP SUCCESS!" & vbCrLf & vbCrLf & _
               "Sent to: " & gmail & vbCrLf & _
               "File: " & GetFileName(zipFile) & vbCrLf & _
               "Time: " & Format(Timer - t0, "0.0") & " sec", vbInformation
    Else
        MsgBox "GMAIL BACKUP FAILED!" & vbCrLf & vbCrLf & _
               "Check:" & vbCrLf & _
               "1. Internet connection" & vbCrLf & _
               "2. Gmail App Password (16-digit)" & vbCrLf & _
               "3. CDO Library installed" & vbCrLf & _
               "4. Less Secure Apps enabled OR 2-Step + App Password used", vbCritical
    End If
End Sub

'========================================
' GOOGLE DRIVE BACKUP
'========================================
Private Sub DoDriveBackup(ByVal zipFile As String)
    Dim drivePath As String
    drivePath = GetGoogleDriveFolder
    
    If drivePath = "" Then
        MsgBox "Google Drive NOT DETECTED!" & vbCrLf & vbCrLf & _
               "Solutions:" & vbCrLf & _
               "1. Install Google Drive Desktop from google.com/drive" & vbCrLf & _
               "2. Sign in and let it sync" & vbCrLf & _
               "3. OR manually set path in Settings sheet cell B15", vbExclamation
        Exit Sub
    End If
    
    Dim driveBackupFolder As String
    driveBackupFolder = drivePath & "\GLOBAL_SOFT_Backups"
    
    On Error Resume Next
    If Dir(driveBackupFolder, vbDirectory) = "" Then MkDir driveBackupFolder
    On Error GoTo 0
    
    Application.StatusBar = "Copying backup to Drive..."
    Dim t0 As Double: t0 = Timer
    Dim ok As Boolean
    
    ok = CopyToDrive(zipFile, driveBackupFolder)
    Application.StatusBar = False
    
    If ok Then
        MsgBox "DRIVE BACKUP SUCCESS!" & vbCrLf & vbCrLf & _
               "File: " & GetFileName(zipFile) & vbCrLf & _
               "Location: " & driveBackupFolder & vbCrLf & _
               "Time: " & Format(Timer - t0, "0.0") & " sec" & vbCrLf & _
               "Google Drive will sync in 1-2 minutes.", vbInformation
    Else
        MsgBox "DRIVE BACKUP FAILED!", vbCritical
        Exit Sub
    End If
    
    CleanOldBackups driveBackupFolder
End Sub

'========================================
' SEND EMAIL (CDO)
'========================================
Private Function SendBackupEmail(gmail As String, pass As String, _
                                  attachmentPath As String, subjectTag As String) As Boolean
    On Error GoTo Failed
    
    Dim cdoMsg As Object, cdoConf As Object
    Set cdoMsg = CreateObject("CDO.Message")
    Set cdoConf = CreateObject("CDO.Configuration")
    
    With cdoConf.Fields
        .item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.gmail.com"
        .item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 465
        .item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
        .item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
        .item("http://schemas.microsoft.com/cdo/configuration/sendusername") = gmail
        .item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = pass
        .item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
        .Update
    End With
    
    Set cdoMsg.Configuration = cdoConf
    With cdoMsg
        .To = gmail
        .From = "GLOBAL SOFT <" & gmail & ">"
        .subject = "GITS " & subjectTag & " - " & Format(Now, "dd-mm-yyyy hh:mm")
        .TextBody = "GLOBAL SOFT Backup" & vbCrLf & _
                    "Type: " & subjectTag & vbCrLf & _
                    "Date: " & Now & vbCrLf & _
                    "User: " & CurrentUserName & vbCrLf & _
                    "PC: " & Environ("COMPUTERNAME")
        If attachmentPath <> "" And Dir(attachmentPath) <> "" Then
            .AddAttachment attachmentPath
        End If
        .Send
    End With
    
    SendBackupEmail = True
    Exit Function
    
Failed:
    SendBackupEmail = False
End Function

'========================================
' COPY TO DRIVE
'========================================
Private Function CopyToDrive(ByVal srcFile As String, ByVal destFolder As String) As Boolean
    On Error GoTo Failed
    
    If Dir(srcFile) = "" Then GoTo Failed
    
    Dim destFile As String
    destFile = destFolder & "\" & GetFileName(srcFile)
    
    If Dir(destFile) <> "" Then Kill destFile
    
    FileCopy srcFile, destFile
    
    CopyToDrive = (Dir(destFile) <> "")
    Exit Function
    
Failed:
    CopyToDrive = False
End Function

'========================================
' FIND GOOGLE DRIVE
'========================================
Private Function GetGoogleDriveFolder() As String
    On Error Resume Next
    
    Dim manual As String
    manual = CStr(ThisWorkbook.Sheets("Settings").Range("B15").value)
    If manual <> "" Then
        If Right(manual, 1) <> "\" Then manual = manual & "\"
        If Dir(manual, vbDirectory) <> "" Then
            GetGoogleDriveFolder = manual
            Exit Function
        End If
    End If
    
    Dim paths(1 To 10) As String
    Dim userProf As String
    userProf = Environ("USERPROFILE")
    
    paths(1) = userProf & "\Google Drive\My Drive\"
    paths(2) = userProf & "\Google Drive\"
    paths(3) = userProf & "\My Drive\"
    paths(4) = userProf & "\Drive\My Drive\"
    paths(5) = userProf & "\OneDrive\"
    paths(6) = "G:\My Drive\"
    paths(7) = "H:\My Drive\"
    paths(8) = "I:\My Drive\"
    paths(9) = "J:\My Drive\"
    paths(10) = "K:\My Drive\"
    
    Dim i As Integer
    For i = 1 To 10
        If Dir(paths(i), vbDirectory) <> "" Then
            GetGoogleDriveFolder = paths(i)
            Exit Function
        End If
    Next i
    
    Dim drv As Long
    Dim drvLetter As String
    For drv = 68 To 90
        drvLetter = Chr(drv)
        If Dir(drvLetter & ":\My Drive\", vbDirectory) <> "" Then
            GetGoogleDriveFolder = drvLetter & ":\My Drive\"
            Exit Function
        End If
    Next drv
    
    GetGoogleDriveFolder = ""
    On Error GoTo 0
End Function

'========================================
' CLEAN OLD BACKUPS (Keep 5)
'========================================
Private Sub CleanOldBackups(ByVal folderPath As String)
    On Error Resume Next
    
    If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
    If Dir(folderPath, vbDirectory) = "" Then Exit Sub
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim folder As Object
    Set folder = fso.GetFolder(folderPath)
    
    Dim fileList As Object
    Set fileList = CreateObject("System.Collections.ArrayList")
    
    Dim f As Object
    For Each f In folder.files
        If LCase(Right(f.name, 4)) = ".zip" And InStr(1, f.name, BACKUP_PREFIX, vbTextCompare) = 1 Then
            fileList.Add f.path
        End If
    Next f
    
    If fileList.count <= KEEP_BACKUPS Then Exit Sub
    
    Dim i As Integer, j As Integer
    Dim temp As String
    For i = 0 To fileList.count - 2
        For j = i + 1 To fileList.count - 1
            If FileDateTime(fileList(i)) < FileDateTime(fileList(j)) Then
                temp = fileList(i)
                fileList(i) = fileList(j)
                fileList(j) = temp
            End If
        Next j
    Next i
    
    For i = KEEP_BACKUPS To fileList.count - 1
        On Error Resume Next
        Kill fileList(i)
        On Error GoTo 0
    Next i
    
    Set fso = Nothing
    On Error GoTo 0
End Sub

'========================================
' RESTORE BACKUP — WITH FILE LOCK FIX
'========================================
Private Sub btnRestoreBackup_Click()
    Dim fd As FileDialog
    Dim BackupFile As String
    
    If MsgBox("RESTORE WARNING:" & vbCrLf & vbCrLf & _
              "Current data will be REPLACED!" & vbCrLf & _
              "Excel will close and reopen with restored data." & vbCrLf & vbCrLf & _
              "Continue?", vbYesNo + vbCritical, "Confirm Restore") = vbNo Then Exit Sub
    
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .title = "Select Complete Backup ZIP (Database + Photos)"
        .Filters.Clear
        .Filters.Add "ZIP Files", "*.zip"
        If .Show <> -1 Then Exit Sub
        BackupFile = .SelectedItems(1)
    End With
    
    Call ExecuteRestore(BackupFile)
End Sub

Private Sub ExecuteRestore(ByVal zipFile As String)
    On Error GoTo ErrorHandler
    
    Dim fso As Object, wsh As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set wsh = CreateObject("WScript.Shell")
    
    Dim mainPath As String
    mainPath = ThisWorkbook.path
    
    Dim extractPath As String
    extractPath = mainPath & "\TEMP_RESTORE_" & Format(Now, "yyyymmddhhmmss")
    
    MkDir extractPath
    
    ' Extract ZIP
    Application.StatusBar = "Extracting backup..."
    Dim psCmd As String
    psCmd = "powershell -NoProfile -ExecutionPolicy Bypass -Command ""Expand-Archive -Path '" & zipFile & "' -DestinationPath '" & extractPath & "' -Force"""
    wsh.Run psCmd, 0, True
    
    ' Find database file
    Dim dbSource As String
    dbSource = ""
    
    If Dir(extractPath & "\Database\*.xlsm") <> "" Then
        dbSource = extractPath & "\Database\" & Dir(extractPath & "\Database\*.xlsm")
    ElseIf Dir(extractPath & "\GLOBAL_SOFT.xlsm") <> "" Then
        dbSource = extractPath & "\GLOBAL_SOFT.xlsm"
    End If
    
    If dbSource = "" Then
        MsgBox "Invalid backup! Database file not found.", vbCritical
        fso.DeleteFolder extractPath, True
        Exit Sub
    End If
    
    ' Prepare file swap
    Dim currentFile As String, oldFile As String
    currentFile = ThisWorkbook.FullName
    oldFile = currentFile & ".OLD_" & Format(Now, "yyyymmddhhmmss")
    
    Application.ScreenUpdating = False
    
    ' Backup current file (safe copy, not SaveAs)
    FileCopy currentFile, oldFile
    
    ' FIX: Create restored file with NEW name to avoid file lock on open workbook
    Dim tempRestoreFile As String
    tempRestoreFile = mainPath & "\RESTORED_GLOBAL_SOFT_" & _
                  Format(Now, "yyyymmdd_hhnnss") & ".xlsm"
    If Dir(tempRestoreFile) <> "" Then Kill tempRestoreFile
    FileCopy dbSource, tempRestoreFile
    
    ' Restore photos to GLOBAL_SOFT_DATA with overwrite protection
    If Dir(extractPath & "\Photos\", vbDirectory) <> "" Then
        Dim photoDest As String
        photoDest = mainPath & "\GLOBAL_SOFT_DATA\"
        
        ' Delete old photos first to avoid duplicates/mixing
        If Dir(photoDest, vbDirectory) <> "" Then
            On Error Resume Next
            fso.DeleteFolder photoDest, True
            On Error GoTo 0
        End If
        MkDir photoDest
        
        fso.CopyFolder extractPath & "\Photos\", photoDest, True
    End If
    
    ' Cleanup
    fso.DeleteFolder extractPath, True
    
    Application.ScreenUpdating = True
    
    MsgBox "RESTORE SUCCESSFUL!" & vbCrLf & vbCrLf & _
           "Restored file: RESTORED_GLOBAL_SOFT.xlsm" & vbCrLf & _
           "Old file saved as: " & GetFileName(oldFile) & vbCrLf & _
           "Excel will now close. Please reopen.", vbInformation
           
    ' Safe reopen batch — opens the NEW restored file (not the locked old one)
    Dim batFile As String
    batFile = mainPath & "\_restore_reopen.bat"
    
    Dim bNum As Integer
    bNum = FreeFile
    Open batFile For Output As #bNum
    Print #bNum, "@echo off"
    Print #bNum, "timeout /t 2 /nobreak >nul"
    Print #bNum, "start """" """ & tempRestoreFile & """"
    Print #bNum, "del ""%~f0"""
    Close #bNum
    
    wsh.Run """" & batFile & """", 0, False
    
    ThisWorkbook.Saved = True
    Application.DisplayAlerts = False
    Application.Quit
    
    Exit Sub
    
ErrorHandler:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "Restore Failed! " & Err.Description, vbCritical
End Sub

'========================================
' UI BUTTONS
'========================================
Private Sub btnSelectLocation_Click()
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    fd.title = "Select Backup Folder"
    If fd.Show = -1 Then
        txtBackupLocation.value = fd.SelectedItems(1)
        On Error Resume Next
        ThisWorkbook.Sheets("Settings").Range("B10").value = fd.SelectedItems(1)
        On Error GoTo 0
    End If
    Set fd = Nothing
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnSaveCloud_Click()
    On Error GoTo SaveError
    
    If Trim(txtGmailID.value) = "" Then
        MsgBox "Enter Gmail ID!", vbExclamation
        Exit Sub
    End If
    
    On Error Resume Next
    ThisWorkbook.Sheets("Settings").Range("B12").value = Trim(txtGmailID.value)
    On Error GoTo 0
    
    Dim rawPass As String, cleanPass As String
    
    If txtPassword.value <> "" And _
       txtPassword.value <> "****************" And _
       txtPassword.value <> "???????????????" And _
       txtPassword.value <> "*****************" Then
        
        rawPass = txtPassword.value
        cleanPass = CleanPasswordChars(rawPass)
        
        If Len(cleanPass) = 16 Then
            On Error Resume Next
            ThisWorkbook.Sheets("Settings").Range("B13").value = cleanPass
            txtPassword.Tag = "SAVED:" & cleanPass
            On Error GoTo 0
        Else
            MsgBox "Password must be 16 chars (letters/numbers only)!" & vbCrLf & _
                   "After removing spaces/special chars: " & Len(cleanPass) & " chars" & vbCrLf & vbCrLf & _
                   "Google shows: xxxx xxxx xxxx xxxx (copy WITHOUT spaces)", vbExclamation
            Exit Sub
        End If
        
    Else
        On Error Resume Next
        rawPass = CStr(ThisWorkbook.Sheets("Settings").Range("B13").value)
        On Error GoTo 0
        
        If rawPass <> "" And rawPass <> "Empty" Then
            cleanPass = CleanPasswordChars(rawPass)
            If Len(cleanPass) = 16 Then
                txtPassword.Tag = "SAVED:" & cleanPass
            End If
        End If
    End If
    
    On Error Resume Next
    ThisWorkbook.Sheets("Settings").Range("B14").value = IIf(chkEnableCloud.value, "TRUE", "FALSE")
    On Error GoTo 0
    
    lblCloudStatus.caption = "Settings Saved"
    lblCloudStatus.ForeColor = RGB(0, 128, 0)
    MsgBox "Cloud Settings Saved!", vbInformation
    
    txtPassword.value = "****************"
    
    Exit Sub
    
SaveError:
    MsgBox "Error: " & Err.Description, vbCritical
End Sub

Private Sub btnTestConnection_Click()
    If Trim(txtGmailID.value) = "" Then
        MsgBox "Enter Gmail ID!", vbExclamation
        Exit Sub
    End If
    
    Dim tp As String
    tp = GetCleanPassword()
    
    If Len(tp) <> 16 Then
        MsgBox "Password must be exactly 16 letters/numbers!" & vbCrLf & vbCrLf & _
               "After removing spaces & special chars:" & vbCrLf & _
               "Length = " & Len(tp) & " characters" & vbCrLf & vbCrLf & _
               "What we detected: '" & tp & "'" & vbCrLf & vbCrLf & _
               "FIX: Go to Google App Passwords page," & vbCrLf & _
               "copy the 16-digit code WITHOUT spaces," & vbCrLf & _
               "paste it fresh, then click Save Settings.", vbExclamation
        Exit Sub
    End If
    
    Me.MousePointer = fmMousePointerHourGlass
    Application.StatusBar = "Testing Gmail..."
    
    Dim ok As Boolean
    ok = SendBackupEmail(txtGmailID.value, tp, "", "Test")
    
    Me.MousePointer = fmMousePointerDefault
    Application.StatusBar = False
    
    If ok Then
        MsgBox "SUCCESS! Test email sent to " & txtGmailID.value, vbInformation
        lblCloudStatus.caption = "Test Passed"
        lblCloudStatus.ForeColor = RGB(0, 150, 0)
    Else
        MsgBox "FAILED! Check internet & app password.", vbCritical
        lblCloudStatus.caption = "Test Failed"
        lblCloudStatus.ForeColor = RGB(255, 0, 0)
    End If
End Sub

Private Sub btnSetupGuide_Click()
    MsgBox "GMAIL SETUP:" & vbCrLf & String(40, "=") & vbCrLf & _
           "1. myaccount.google.com" & vbCrLf & _
           "2. Security -> 2-Step Verification ON" & vbCrLf & _
           "3. Security -> App Passwords" & vbCrLf & _
           "4. Select: Mail | Windows Computer" & vbCrLf & _
           "5. COPY 16-digit code" & vbCrLf & _
           "6. PASTE here -> Test -> Save" & vbCrLf & vbCrLf & _
           "GOOGLE DRIVE:" & vbCrLf & _
           "Install Google Drive Desktop" & vbCrLf & _
           "OR set path in Settings sheet B15", vbInformation
End Sub

'========================================
' UTILITY HELPERS
'========================================
Private Function GetFileName(ByVal fullPath As String) As String
    Dim p As Integer
    p = InStrRev(fullPath, "\")
    If p > 0 Then GetFileName = Mid(fullPath, p + 1) Else GetFileName = fullPath
End Function

Private Function GetFileSizeText(ByVal filePath As String) As String
    On Error Resume Next
    Dim sz As Double
    sz = FileLen(filePath)
    If sz > 1024 ^ 3 Then
        GetFileSizeText = Format(sz / 1024 ^ 3, "0.00") & " GB"
    ElseIf sz > 1024 ^ 2 Then
        GetFileSizeText = Format(sz / 1024 ^ 2, "0.00") & " MB"
    Else
        GetFileSizeText = Format(sz / 1024, "0.00") & " KB"
    End If
    On Error GoTo 0
End Function

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

Private Sub Frame1_Click()
End Sub

