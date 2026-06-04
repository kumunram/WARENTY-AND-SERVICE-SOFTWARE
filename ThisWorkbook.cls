VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Workbook_Open()
    On Error GoTo OpenError
    
    ' === INITIALIZE ===
    Call InitializeGlobalConfig
    Call StartAutoSaveTimer
    Call modAutoBackup.ScheduleNextBackup
    Call modAutoBackup.DoOpenBackup

    ' Hide UI
    Application.ScreenUpdating = False
    Application.DisplayFormulaBar = False
    Application.DisplayStatusBar = True
    Application.DisplayAlerts = False

    ' LoginScreen
    Dim wsLogin As Worksheet
    On Error Resume Next
    Set wsLogin = ThisWorkbook.Sheets("LoginScreen")
    On Error GoTo OpenError
    If wsLogin Is Nothing Then
        Set wsLogin = ThisWorkbook.Sheets.Add
        wsLogin.name = "LoginScreen"
    End If
    wsLogin.visible = xlSheetVisible
    wsLogin.Activate

    ' Hide others
    Dim sh As Worksheet
    For Each sh In ThisWorkbook.Sheets
        If sh.name <> "LoginScreen" Then sh.visible = xlVeryHidden
    Next sh

    ActiveWindow.DisplayGridlines = False
    ActiveWindow.DisplayWorkbookTabs = False
    Application.visible = False
    Application.ScreenUpdating = True

    ' Main Loop
    Do
        IsLoggedIn = False
        CurrentUserName = ""
        CurrentUserRole = ""
        gUserName = ""
        gUserRole = ""
        gUserRow = 0

        Load frmLogin
        frmLogin.Show vbModal
        If Not IsLoggedIn Then Exit Do

        For Each sh In ThisWorkbook.Sheets
            sh.visible = xlSheetVisible
        Next sh

        Load frmDashboard
        frmDashboard.Show vbModal
        If IsLoggedIn Then Exit Do

        For Each sh In ThisWorkbook.Sheets
            If sh.name <> "LoginScreen" Then sh.visible = xlVeryHidden
        Next sh
    Loop

    ' Cleanup
    Application.visible = True
    Application.DisplayFormulaBar = True
    Application.DisplayStatusBar = True
    Application.DisplayAlerts = True
    Application.Quit
    Exit Sub

OpenError:
    Application.visible = True
    Application.DisplayFormulaBar = True
    Application.DisplayStatusBar = True
    Application.DisplayAlerts = True
    MsgBox "Startup Error: " & Err.Description, vbCritical, "GLOBAL SOFT"
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error GoTo CloseError
    
    ' Stop timers
    Call StopAutoSaveTimer
    modAutoBackup.CancelScheduledBackup
    
    ' Force save with check
    Application.DisplayAlerts = False
    
    Dim saveSuccess As Boolean
    saveSuccess = True
    
    On Error Resume Next
    If Not ThisWorkbook.Saved Then
        ThisWorkbook.Save
    End If
    If Err.Number <> 0 Then saveSuccess = False
    On Error GoTo CloseError
    
    Application.DisplayAlerts = True
    
    ' Backup only if save OK
    If saveSuccess Then
        modAutoBackup.DoAutoBackup False
    End If
    
    ' Restore UI
    Application.visible = True
    Application.DisplayFormulaBar = True
    Application.DisplayStatusBar = True
    Application.DisplayAlerts = True
    
    Dim sh As Worksheet
    For Each sh In ThisWorkbook.Sheets
        sh.visible = xlSheetVisible
    Next sh
    
    Exit Sub

CloseError:
    Application.DisplayAlerts = True
    MsgBox "Error during close: " & Err.Description & vbCrLf & _
           "Please manually save the file (Ctrl+S) before closing!", vbCritical
End Sub

