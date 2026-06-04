Attribute VB_Name = "modUserSession"
Option Explicit

Public CurrentUserRole As String
Public CurrentUserName As String
Public IsLoggedIn As Boolean

Public Sub LockExcelUI()
    On Error Resume Next
    ' Only hide UI elements, DON'T hide sheets here
    Application.ScreenUpdating = False
    Application.visible = False
    Application.DisplayFormulaBar = False
    Application.DisplayStatusBar = False
    Application.DisplayAlerts = False
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.DisplayWorkbookTabs = False
    Application.ScreenUpdating = True
    On Error GoTo 0
End Sub

Public Sub UnlockExcelUI()
    On Error Resume Next
    Application.ScreenUpdating = False
    Application.visible = True
    Application.DisplayFormulaBar = True
    Application.DisplayStatusBar = True
    Application.DisplayAlerts = True
    ActiveWindow.DisplayGridlines = True
    ActiveWindow.DisplayWorkbookTabs = True
    
    ' Show all sheets
    Dim sh As Worksheet
    For Each sh In ThisWorkbook.Sheets
        sh.visible = xlSheetVisible
    Next sh
    
    Application.ScreenUpdating = True
    On Error GoTo 0
End Sub

Public Function HasPermission(ByVal buttonName As String) As Boolean
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("User_Permission")
    Dim lastRow As Long, i As Long
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    
    For i = 2 To lastRow
        If UCase(Trim(ws.Cells(i, 1).value & "")) = UCase(CurrentUserName) Then
            Select Case buttonName
                Case "Dashboard":     HasPermission = (ws.Cells(i, 4).value & "" = "YES")
                Case "NewEntry":      HasPermission = (ws.Cells(i, 5).value & "" = "YES")
                Case "AssignWork":    HasPermission = (ws.Cells(i, 6).value & "" = "YES")
                Case "ServiceUpdate": HasPermission = (ws.Cells(i, 7).value & "" = "YES")
                Case "CustomerDB":    HasPermission = (ws.Cells(i, 8).value & "" = "YES")
                Case "EngineerReg":   HasPermission = (ws.Cells(i, 9).value & "" = "YES")
                Case "Reports":       HasPermission = (ws.Cells(i, 10).value & "" = "YES")
                Case "Settings":      HasPermission = (ws.Cells(i, 11).value & "" = "YES")
                Case "Backup":        HasPermission = (ws.Cells(i, 12).value & "" = "YES")
                Case Else:            HasPermission = False
            End Select
            Exit Function
        End If
    Next i
    HasPermission = False
    On Error GoTo 0
End Function
