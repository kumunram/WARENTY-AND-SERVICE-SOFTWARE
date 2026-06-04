Attribute VB_Name = "modIntegrity"
'=================================
' FULL DATABASE SCAN
'=================================

Sub RunIntegrityScan()

Call CheckBrokenRelations
Call CheckDuplicateCustomers
Call CheckDuplicateEntries

MsgBox "Database Scan Completed", vbInformation

End Sub


'=================================
' BROKEN RELATION CHECK
'=================================

Sub CheckBrokenRelations()

Dim ws As Worksheet
Dim i As Long
Dim entryID As String

Set ws = Sheets("Job_Product")

For i = 2 To ws.Cells(ws.Rows.count, 1).End(xlUp).row

entryID = ws.Cells(i, 1).value

If EntryExists_Master(entryID) = False Then

ws.Cells(i, 1).Interior.Color = vbRed

End If

Next i

End Sub


'=================================
' DUPLICATE CUSTOMER CHECK
'=================================

Sub CheckDuplicateCustomers()

Dim ws As Worksheet
Dim dict As Object
Dim i As Long
Dim mobile As String

Set ws = Sheets("Customer_Master")
Set dict = CreateObject("Scripting.Dictionary")

For i = 2 To ws.Cells(ws.Rows.count, 2).End(xlUp).row

mobile = ws.Cells(i, 2).value

If dict.exists(mobile) Then

ws.Cells(i, 2).Interior.Color = vbYellow

Else

dict.Add mobile, True

End If

Next i

End Sub


'=================================
' DUPLICATE ENTRY CHECK
'=================================

Sub CheckDuplicateEntries()

Dim ws As Worksheet
Dim dict As Object
Dim i As Long
Dim entryID As String

Set ws = Sheets("Job_Master")
Set dict = CreateObject("Scripting.Dictionary")

For i = 2 To ws.Cells(ws.Rows.count, 1).End(xlUp).row

entryID = ws.Cells(i, 1).value

If dict.exists(entryID) Then

ws.Cells(i, 1).Interior.Color = vbYellow

Else

dict.Add entryID, True

End If

Next i

End Sub
