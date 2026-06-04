Attribute VB_Name = "Module1"
Option Explicit

'========================================
' DEVELOPER MODE — ?? ??? ?????
'========================================
Sub DeveloperMode_ON()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Sheets
        ws.visible = xlSheetVisible
    Next ws
    MsgBox "DEVELOPER MODE ON" & vbCrLf & vbCrLf & "All sheets are now visible." & vbCrLf & "You can delete old data.", vbInformation
End Sub

'========================================
' PRODUCTION MODE — ?? hide ??? (????? login/dashboard visible)
'========================================
Sub DeveloperMode_OFF()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Sheets
        ' ?? sheets visible ???? ???, ???? name ???? ????
        Select Case ws.name
            Case "Customer_Master", "Dashboard", "Sheet1", "LoginScreen"
                ws.visible = xlSheetVisible
            Case Else
                ws.visible = xlSheetVeryHidden
        End Select
    Next ws
    MsgBox "PRODUCTION MODE ON" & vbCrLf & vbCrLf & "All sensitive sheets hidden." & vbCrLf & "User login will show now.", vbInformation
End Sub
Public Sub FixAssignMasterSwap()
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim tempD As String, tempE As String
    
    Set ws = ThisWorkbook.Sheets("Assign_Master")
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    
    For i = 2 To lastRow
        tempD = Trim(ws.Cells(i, 4).value & "")
        tempE = Trim(ws.Cells(i, 5).value & "")
        ws.Cells(i, 4).value = tempE   ' Mobile
        ws.Cells(i, 5).value = tempD   ' Customer Name
    Next i
    
    MsgBox "Fixed " & (lastRow - 1) & " rows!", vbInformation
End Sub
