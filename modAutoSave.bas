Attribute VB_Name = "modAutoSave"
Option Explicit

Public g_NextAutoSave As Date
Public Const AUTO_SAVE_INTERVAL As String = "00:05:00"  ' 5 minutes

'========================================
' AUTO-SAVE TIMER START
'========================================
Public Sub StartAutoSaveTimer()
    On Error Resume Next
    
    ' Cancel any existing timer first
    If g_NextAutoSave > 0 Then
        Application.OnTime EarliestTime:=g_NextAutoSave, Procedure:="DoAutoSaveNow", Schedule:=False
    End If
    
    ' Set next auto-save time (Now + 5 minutes)
    g_NextAutoSave = Now + TimeValue(AUTO_SAVE_INTERVAL)
    
    ' Schedule the auto-save
    Application.OnTime EarliestTime:=g_NextAutoSave, Procedure:="DoAutoSaveNow"
    
    On Error GoTo 0
End Sub

'========================================
' AUTO-SAVE EXECUTE (Every 5 Minutes)
'========================================
Public Sub DoAutoSaveNow()
    On Error GoTo SaveErr
    
    ' ============================================================
    ' ?? CRITICAL: Check if workbook has unsaved changes
    ' ============================================================
    If Not ThisWorkbook.Saved Then
        Application.DisplayAlerts = False
        ThisWorkbook.Save
        Application.DisplayAlerts = True
        Application.StatusBar = "Auto-saved at " & Format(Now, "hh:mm:ss")
    End If
    
    ' Reschedule for next time
    StartAutoSaveTimer
    Exit Sub

SaveErr:
    Application.DisplayAlerts = True
    Application.StatusBar = "AutoSave Failed: " & Err.Description
    StartAutoSaveTimer  ' Still reschedule
End Sub

'========================================
' TIMER STOP (Workbook Close)
'========================================
Public Sub StopAutoSaveTimer()
    On Error Resume Next
    If g_NextAutoSave > 0 Then
        Application.OnTime EarliestTime:=g_NextAutoSave, Procedure:="DoAutoSaveNow", Schedule:=False
    End If
    g_NextAutoSave = 0
    On Error GoTo 0
End Sub
