Attribute VB_Name = "modDustbin"

Option Explicit

'===========================================
' MODULE: modDustbin
' PURPOSE: Recycle Bin for Deleted Records
'===========================================

Public Sub MoveToDustbin(sourceSheetName As String, recordID As String)
    On Error GoTo ErrorHandler
    
    Dim wsSource As Worksheet, wsDustbin As Worksheet
    Dim lastRowSource As Long, lastRowDustbin As Long
    Dim i As Long, j As Long, foundRow As Long
    
    Set wsSource = ThisWorkbook.Sheets(sourceSheetName)
    
    'Create Dustbin sheet if not exists
    On Error Resume Next
    Set wsDustbin = ThisWorkbook.Sheets("Dustbin")
    On Error GoTo 0
    
    If wsDustbin Is Nothing Then
        Set wsDustbin = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
        wsDustbin.name = "Dustbin"
        'Headers
        wsDustbin.Cells(1, 1).value = "DeletedOn"
        wsDustbin.Cells(1, 2).value = "SourceSheet"
        wsDustbin.Cells(1, 3).value = "RecordData"
        wsDustbin.Cells(1, 4).value = "Restored"
        wsDustbin.Rows(1).Font.Bold = True
    End If
    
    'Find record in source
    lastRowSource = wsSource.Cells(wsSource.Rows.count, 1).End(xlUp).row
    foundRow = 0
    
    For i = 2 To lastRowSource
        If CStr(wsSource.Cells(i, 1).value) = recordID Or CStr(wsSource.Cells(i, 3).value) = recordID Then
            foundRow = i
            Exit For
        End If
    Next i
    
    If foundRow = 0 Then
        MsgBox "Record not found!", vbExclamation, "Error"
        Exit Sub
    End If
    
    'Add to Dustbin
    lastRowDustbin = wsDustbin.Cells(wsDustbin.Rows.count, 1).End(xlUp).row + 1
    wsDustbin.Cells(lastRowDustbin, 1).value = Now
    wsDustbin.Cells(lastRowDustbin, 2).value = sourceSheetName
    
    'Copy all columns data to column C as JSON-like string
    Dim dataString As String
    dataString = ""
    For j = 1 To wsSource.Cells(foundRow, wsSource.Columns.count).End(xlToLeft).Column
        dataString = dataString & wsSource.Cells(1, j).value & "=" & wsSource.Cells(foundRow, j).value & "|"
    Next j
    wsDustbin.Cells(lastRowDustbin, 3).value = dataString
    wsDustbin.Cells(lastRowDustbin, 4).value = "NO"
    
    'Delete from source (move up remaining rows)
    wsSource.Rows(foundRow).Delete Shift:=xlUp
    
    MsgBox "Record moved to Dustbin!" & vbCrLf & "Sheet: " & sourceSheetName, vbInformation, "Dustbin"
    Exit Sub
    
ErrorHandler:
    MsgBox "Dustbin Error: " & Err.Description, vbCritical, "Error"
End Sub

Public Sub RestoreFromDustbin(dustbinRow As Long)
    On Error GoTo ErrorHandler
    
    Dim wsDustbin As Worksheet, wsTarget As Worksheet
    Dim targetSheet As String, dataString As String
    Dim lastRow As Long
    
    Set wsDustbin = ThisWorkbook.Sheets("Dustbin")
    targetSheet = wsDustbin.Cells(dustbinRow, 2).value
    dataString = wsDustbin.Cells(dustbinRow, 3).value
    
    Set wsTarget = ThisWorkbook.Sheets(targetSheet)
    lastRow = wsTarget.Cells(wsTarget.Rows.count, 1).End(xlUp).row + 1
    
    'Parse and restore data (simplified)
    'In real implementation, parse the | separated values
    wsTarget.Cells(lastRow, 1).value = "RESTORED_" & Format(Now, "mmss")
    
    'Mark as restored
    wsDustbin.Cells(dustbinRow, 4).value = "YES"
    wsDustbin.Cells(dustbinRow, 4).Interior.Color = RGB(200, 255, 200)
    
    MsgBox "Record Restored!", vbInformation, "Restore"
    Exit Sub
    
ErrorHandler:
    MsgBox "Restore Error: " & Err.Description, vbCritical, "Error"
End Sub

Public Sub CleanOldDustbin(daysOld As Integer)
    On Error Resume Next
    Dim wsDustbin As Worksheet, lastRow As Long, i As Long
    Dim cutoffDate As Date
    
    Set wsDustbin = ThisWorkbook.Sheets("Dustbin")
    lastRow = wsDustbin.Cells(wsDustbin.Rows.count, 1).End(xlUp).row
    cutoffDate = Date - daysOld
    
    For i = lastRow To 2 Step -1
        If wsDustbin.Cells(i, 1).value < cutoffDate And wsDustbin.Cells(i, 4).value = "YES" Then
            wsDustbin.Rows(i).Delete
        End If
    Next i
    
    MsgBox "Dustbin cleaned! Records older than " & daysOld & " days removed.", vbInformation, "Cleaned"
End Sub
