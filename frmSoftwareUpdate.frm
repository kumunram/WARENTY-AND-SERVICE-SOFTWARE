VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSoftwareUpdate 
   Caption         =   "SOFTWARE UPDATE"
   ClientHeight    =   12750
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17490
   OleObjectBlob   =   "frmSoftwareUpdate.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSoftwareUpdate"
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
' FORM INITIALIZE
'===========================================
Private Sub UserForm_Initialize()
    LoadCurrentVersion
    LoadUpdateSettings
End Sub

'===========================================
' LOAD CURRENT VERSION INFO
'===========================================
Private Sub LoadCurrentVersion()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Software_Config")
    On Error GoTo 0
    
    If ws Is Nothing Then Exit Sub
    
    Me.txtInstalledVersion.value = CStr(ws.Range("B56").value)
    Me.txtReleaseDate.value = CStr(ws.Range("B51").value)
    Me.txtLicenseStatus.value = CStr(ws.Range("B52").value)
End Sub

'===========================================
' LOAD UPDATE SETTINGS
'===========================================
Private Sub LoadUpdateSettings()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Software_Config")
    On Error GoTo 0
    
    If ws Is Nothing Then Exit Sub
    
    Me.txtNewFile.value = CStr(ws.Range("B53").value)
    Me.txtBackupTo.value = CStr(ws.Range("B54").value)
    Me.txtRestoreFile.value = CStr(ws.Range("B55").value)
End Sub

'===========================================
' BROWSE NEW FILE BUTTON
'===========================================
Private Sub btnBrowseNewFile_Click()
    Dim fd As Office.FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    With fd
        .title = "Select New Version File"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xlsm; *.xlsx"
        
        If .Show = -1 Then
            Me.txtNewFile.value = .SelectedItems(1)
        End If
    End With
End Sub

'===========================================
' BROWSE BACKUP LOCATION BUTTON
'===========================================
Private Sub btnBrowseBackup_Click()
    Dim fd As Office.FileDialog
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    
    With fd
        .title = "Select Backup Folder"
        If .Show = -1 Then
            Me.txtBackupTo.value = .SelectedItems(1)
        End If
    End With
End Sub

'===========================================
' UPDATE NOW BUTTON
'===========================================
Private Sub btnUpdateNow_Click()
    Dim newFilePath As String
    Dim backupPath As String
    
    newFilePath = Me.txtNewFile.value
    backupPath = Me.txtBackupTo.value
    
    ' Validation
    If newFilePath = "" Then
        MsgBox "Please select new version file!", vbExclamation
        Exit Sub
    End If
    
    If backupPath = "" Then
        MsgBox "Please select backup location!", vbExclamation
        Exit Sub
    End If
    
    If Dir(newFilePath) = "" Then
        MsgBox "New file not found!", vbExclamation
        Exit Sub
    End If
    
    ' Confirm update
    If MsgBox("Are you sure you want to update?" & vbCrLf & vbCrLf & _
              "Current Version: " & Me.txtInstalledVersion.value & vbCrLf & _
              "New File: " & newFilePath & vbCrLf & vbCrLf & _
              "A backup will be created before update.", vbQuestion + vbYesNo, "Confirm Update") = vbNo Then
        Exit Sub
    End If
    
    ' Perform update
    PerformUpdate newFilePath, backupPath
End Sub

'===========================================
' PERFORM UPDATE
'===========================================
Private Sub PerformUpdate(ByVal newFilePath As String, ByVal backupPath As String)
    On Error GoTo ErrorHandler
    
    Dim currentFile As String
    Dim BackupFile As String
    
    currentFile = ThisWorkbook.FullName
    
    ' Create backup filename with timestamp
    BackupFile = backupPath & "\GLOBAL_SOFT_Backup_" & Format(Now, "yyyymmdd_hhmmss") & ".xlsm"
    
    ' Step 1: Create Backup
    Application.DisplayAlerts = False
    FileCopy currentFile, BackupFile
    Application.DisplayAlerts = True
    
    MsgBox "Backup created successfully!" & vbCrLf & "Location: " & BackupFile, vbInformation
    
    ' Step 2: Transfer data to new file
    TransferDataToNewFile newFilePath
    
    ' Save settings
    SaveUpdateSettings newFilePath, backupPath
    
    MsgBox "Update completed successfully!" & vbCrLf & vbCrLf & _
           "Please close this file and open the updated version.", vbInformation
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error during update: " & Err.Description, vbCritical
End Sub

'===========================================
' TRANSFER DATA TO NEW FILE
'===========================================
Private Sub TransferDataToNewFile(ByVal newFilePath As String)
    On Error GoTo ErrorHandler
    
    Dim newWB As Workbook
    Dim currentWB As Workbook
    Dim i As Integer
    
    Set currentWB = ThisWorkbook
    
    ' Open new version file
    Application.DisplayAlerts = False
    Set newWB = Workbooks.Open(newFilePath, ReadOnly:=False)
    Application.DisplayAlerts = True
    
    ' Transfer data from current to new
    For i = 1 To currentWB.Worksheets.count
        Dim wsName As String
        wsName = currentWB.Worksheets(i).name
        
        ' Skip system sheets (new version ??? already ?????)
        If wsName <> "Software_Config" And wsName <> "Version_Info" Then
            On Error Resume Next
            Dim targetWS As Worksheet
            Set targetWS = newWB.Worksheets(wsName)
            On Error GoTo ErrorHandler
            
            If Not targetWS Is Nothing Then
                ' Copy data
                currentWB.Worksheets(i).UsedRange.Copy
                targetWS.Range("A1").PasteSpecial xlPasteAll
                Application.CutCopyMode = False
            End If
        End If
    Next i
    
    ' Save new file with transferred data
    newWB.Save
    newWB.Close SaveChanges:=True
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error transferring data: " & Err.Description, vbCritical
End Sub

'===========================================
' OPEN BACKUP FOLDER BUTTON
'===========================================
Private Sub btnOpenBackupFolder_Click()
    Dim backupPath As String
    backupPath = Me.txtBackupTo.value
    
    If backupPath = "" Then
        MsgBox "Please select backup location first!", vbExclamation
        Exit Sub
    End If
    
    If Dir(backupPath, vbDirectory) = "" Then
        MsgBox "Backup folder not found!", vbExclamation
        Exit Sub
    End If
    
    ShellExecute 0, "open", backupPath, vbNullString, vbNullString, 1
End Sub

'===========================================
' BROWSE RESTORE FILE BUTTON
'===========================================
Private Sub btnBrowseRestore_Click()
    Dim fd As Office.FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    With fd
        .title = "Select Backup File to Restore"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Excel Backup Files", "*.xlsm; *.xlsx"
        
        If .Show = -1 Then
            Me.txtRestoreFile.value = .SelectedItems(1)
        End If
    End With
End Sub

'===========================================
' RESTORE BUTTON
'===========================================
Private Sub btnRestore_Click()
    Dim restoreFilePath As String
    
    restoreFilePath = Me.txtRestoreFile.value
    
    If restoreFilePath = "" Then
        MsgBox "Please select backup file to restore!", vbExclamation
        Exit Sub
    End If
    
    If Dir(restoreFilePath) = "" Then
        MsgBox "Backup file not found!", vbExclamation
        Exit Sub
    End If
    
    ' Confirm restore
    If MsgBox("WARNING: This will restore the software to previous version!" & vbCrLf & vbCrLf & _
              "Current data may be lost." & vbCrLf & vbCrLf & _
              "Are you sure you want to continue?", vbExclamation + vbYesNo, "Confirm Restore") = vbNo Then
        Exit Sub
    End If
    
    ' Open backup file
    Workbooks.Open restoreFilePath
    
    MsgBox "Restore completed!" & vbCrLf & "Previous version is now active.", vbInformation
    
    ' Close current workbook
    ThisWorkbook.Close SaveChanges:=False
End Sub

'===========================================
' SAVE UPDATE SETTINGS
'===========================================
Private Sub SaveUpdateSettings(ByVal newFile As String, ByVal backupPath As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Software_Config")
    
    ws.Range("B53").value = newFile
    ws.Range("B54").value = backupPath
End Sub

'===========================================
' SAVE BUTTON
'===========================================
Private Sub btnSave_Click()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Software_Config")
    
    ws.Range("B56").value = Me.txtInstalledVersion.value
    ws.Range("B51").value = Me.txtReleaseDate.value
    ws.Range("B52").value = Me.txtLicenseStatus.value
    ws.Range("B53").value = Me.txtNewFile.value
    ws.Range("B54").value = Me.txtBackupTo.value
    ws.Range("B55").value = Me.txtRestoreFile.value
    
    MsgBox "Update settings saved!", vbInformation
End Sub

'===========================================
' RESET BUTTON
'===========================================
Private Sub btnReset_Click()
    LoadCurrentVersion
    LoadUpdateSettings
    MsgBox "Settings reloaded!", vbInformation
End Sub

'===========================================
' CLOSE BUTTON
'===========================================
Private Sub btnClose_Click()
    Unload Me
End Sub

