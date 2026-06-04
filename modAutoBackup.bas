Attribute VB_Name = "modAutoBackup"
Option Explicit

Public Const BACKUP_FOLDER As String = "AutoBackup"
Public Const BACKUP_INTERVAL As String = "06:00:00"   ' 6 hours
Public g_OpenBackupDone As Boolean

'========================================
' MAIN BACKUP PROCEDURE
'========================================
Public Sub DoAutoBackup(Optional showSuccessMsg As Boolean = True)
    On Error GoTo ErrHandler
    
    Dim wb As Workbook
    Set wb = ThisWorkbook
    
    ' Check if workbook is saved first
    If wb.path = "" Then
        If showSuccessMsg Then
            MsgBox "Please save the workbook first before backup!", vbExclamation, "Backup"
        End If
        Exit Sub
    End If
    
    ' Create backup folder
    Dim backupPath As String
    backupPath = wb.path & "\" & BACKUP_FOLDER & "\"
    If Dir(backupPath, vbDirectory) = "" Then MkDir backupPath
    
    ' Create backup filename
    Dim bkpFile As String
    bkpFile = "BACKUP_" & Format(Now, "yyyy-mm-dd_hh-mm-ss") & "_" & wb.name
    
    ' Save backup copy
    Application.DisplayAlerts = False
    wb.SaveCopyAs fileName:=backupPath & bkpFile
    Application.DisplayAlerts = True
    
    ' Save last backup time
    SaveLastBackupTime Now
    
    ' Schedule next backup
    ScheduleNextBackup
    
    ' Show success message only if requested
    If showSuccessMsg Then
        Application.StatusBar = "Backup created: " & bkpFile
    End If
    
    Exit Sub
    
ErrHandler:
    Application.DisplayAlerts = True
    If showSuccessMsg Then
        MsgBox "Auto Backup Failed!" & vbCrLf & "Error: " & Err.Description, vbCritical, "Backup Error"
    End If
End Sub

'========================================
' SAVE LAST BACKUP TIME
'========================================
Private Sub SaveLastBackupTime(dt As Date)
    On Error Resume Next
    ThisWorkbook.Sheets("Software_Config").Range("Z1").value = dt
    On Error GoTo 0
End Sub

Public Function GetLastBackupTime() As Date
    On Error Resume Next
    GetLastBackupTime = ThisWorkbook.Sheets("Software_Config").Range("Z1").value
    If Err.Number <> 0 Or GetLastBackupTime = 0 Then
        GetLastBackupTime = #1/1/2000#
    End If
    On Error GoTo 0
End Function

'========================================
' SCHEDULE NEXT BACKUP
'========================================
Public Sub ScheduleNextBackup()
    On Error Resume Next
    
    ' Cancel old schedule
    Dim oldTime As Date
    oldTime = ThisWorkbook.Sheets("Software_Config").Range("Z2").value
    If oldTime > #1/1/2000# Then
        Application.OnTime EarliestTime:=oldTime, Procedure:="DoTimedBackup", Schedule:=False
    End If
    
    ' Set new schedule
    Dim nextTime As Date
    nextTime = Now + TimeValue(BACKUP_INTERVAL)
    
    ThisWorkbook.Sheets("Software_Config").Range("Z2").value = nextTime
    Application.OnTime EarliestTime:=nextTime, Procedure:="DoTimedBackup"
    
    On Error GoTo 0
End Sub

'========================================
' CANCEL SCHEDULED BACKUP
'========================================
Public Sub CancelScheduledBackup()
    On Error Resume Next
    Dim oldTime As Date
    oldTime = ThisWorkbook.Sheets("Software_Config").Range("Z2").value
    If oldTime > #1/1/2000# Then
        Application.OnTime EarliestTime:=oldTime, Procedure:="DoTimedBackup", Schedule:=False
    End If
    ThisWorkbook.Sheets("Software_Config").Range("Z2").ClearContents
    On Error GoTo 0
End Sub

'========================================
' TIMED BACKUP TRIGGER
'========================================
Public Sub DoTimedBackup()
    DoAutoBackup False  ' Don't show msgbox for auto-triggered backups
End Sub

'========================================
' BACKUP ON OPEN (Once per session)
'========================================
Public Sub DoOpenBackup()
    On Error Resume Next
    If Not g_OpenBackupDone Then
        g_OpenBackupDone = True
        DoAutoBackup False  ' Silent backup on open
    End If
    On Error GoTo 0
End Sub
