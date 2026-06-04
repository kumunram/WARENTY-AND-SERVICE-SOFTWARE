Attribute VB_Name = "modRelation"
Function CustomerExists(customerID As String) As Boolean

Dim ws As Worksheet
Dim f As Range

Set ws = Sheets("Customer_Master")

Set f = ws.Columns(1).Find(customerID, LookIn:=xlValues, LookAt:=xlWhole)

If f Is Nothing Then
CustomerExists = False
Else
CustomerExists = True
End If

End Function
Function EntryExists(entryID As String) As Boolean

Dim ws As Worksheet
Dim f As Range

Set ws = Sheets("Job_Master")

Set f = ws.Columns(1).Find(entryID, LookIn:=xlValues, LookAt:=xlWhole)

If f Is Nothing Then
EntryExists = False
Else
EntryExists = True
End If

End Function

Sub CheckRelations()

Dim ws As Worksheet
Dim i As Long
Dim entryID As String

Set ws = Sheets("Job_Product")

For i = 2 To ws.Cells(ws.Rows.count, 1).End(xlUp).row

entryID = ws.Cells(i, 1).value

If EntryExists(entryID) = False Then

ws.Cells(i, 1).Interior.Color = vbRed

End If

Next i

End Sub
Function CustomerHasJobs(customerID As String) As Boolean

Dim ws As Worksheet
Dim f As Range

Set ws = Sheets("Job_Master")

Set f = ws.Columns(2).Find(customerID, LookIn:=xlValues, LookAt:=xlWhole)

CustomerHasJobs = Not f Is Nothing

End Function
