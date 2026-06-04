VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmStatusManager 
   Caption         =   "Status Manager"
   ClientHeight    =   10605
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   19995
   OleObjectBlob   =   "frmStatusManager.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmStatusManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    Me.caption = "GLOBAL SOFT - Status Manager"
    SetupListView
    LoadAllEntries
    AddMinMaxButtons Me
    'Initial labels
    lblEntryID.caption = "Entry ID: --"
    lblPresentStatus.caption = "Present Status: --"
    cmbNewStatus.Clear
End Sub

Private Sub SetupListView()
    With lstEntries
        .ColumnHeaders.Clear
        .ListItems.Clear
        .View = lvwReport
        .FullRowSelect = True
        .Gridlines = True
        
        .ColumnHeaders.Add , , "EntryID", 80
        .ColumnHeaders.Add , , "Type", 80
        .ColumnHeaders.Add , , "Current Status", 120
        .ColumnHeaders.Add , , "Customer", 140
        .ColumnHeaders.Add , , "Mobile", 100
        .ColumnHeaders.Add , , "Product", 120
        .ColumnHeaders.Add , , "Company", 100
        .ColumnHeaders.Add , , "Model", 100
        .ColumnHeaders.Add , , "Serial", 120
    End With
End Sub

Private Sub LoadAllEntries()
    On Error Resume Next
    Dim wsEntry As Worksheet, wsCust As Worksheet
    Dim lastRow As Long, lastRowCust As Long
    Dim i As Long, k As Long
    Dim li As Object
    Dim entryID As String, entryType As String, status As String
    Dim custID As String, custName As String, mobile As String
    
    Set wsEntry = ThisWorkbook.Sheets("Job_Product")
    Set wsCust = ThisWorkbook.Sheets("Customer_Master")
    lastRow = wsEntry.Cells(wsEntry.Rows.count, 1).End(xlUp).row
    lastRowCust = wsCust.Cells(wsCust.Rows.count, 1).End(xlUp).row
    
    lstEntries.ListItems.Clear
    
    For i = 2 To lastRow
        entryID = Trim(wsEntry.Cells(i, 1).value)
        entryType = UCase(Trim(wsEntry.Cells(i, 3).value))
        status = UCase(Trim(wsEntry.Cells(i, 4).value))
        
        If entryID = "" Then GoTo NextRow
        
        'Customer Lookup
        custID = Trim(wsEntry.Cells(i, 2).value)
        custName = "": mobile = ""
        For k = 2 To lastRowCust
            If UCase(Trim(wsCust.Cells(k, 1).value)) = UCase(custID) Then
                custName = Trim(wsCust.Cells(k, 3).value)
                mobile = Trim(wsCust.Cells(k, 2).value)
                Exit For
            End If
        Next k
        
        Set li = lstEntries.ListItems.Add(text:=entryID)
        li.SubItems(1) = entryType
        li.SubItems(2) = status
        li.SubItems(3) = custName
        li.SubItems(4) = mobile
        li.SubItems(5) = Trim(wsEntry.Cells(i, 6).value)
        li.SubItems(6) = Trim(wsEntry.Cells(i, 7).value)
        li.SubItems(7) = Trim(wsEntry.Cells(i, 8).value)
        li.SubItems(8) = Trim(wsEntry.Cells(i, 9).value)
        
        'Color by status
        Select Case status
            Case "ACTIVE": li.ForeColor = RGB(0, 150, 0): li.Bold = True
            Case "ASSIGNED": li.ForeColor = RGB(0, 0, 200): li.Bold = True
            Case "IN PROGRESS": li.ForeColor = RGB(200, 100, 0): li.Bold = True
            Case "COMPLETED": li.ForeColor = RGB(150, 0, 150): li.Bold = True
            Case "READY": li.ForeColor = RGB(128, 0, 128): li.Bold = True
            Case "DELIVERED": li.ForeColor = RGB(100, 100, 100): li.Bold = False
            Case "RECEIVE_REPAIRED": li.ForeColor = RGB(0, 102, 153): li.Bold = True
        End Select
NextRow:
    Next i
    
    lblCount.caption = "Total: " & lstEntries.ListItems.count
End Sub

'========================================
' ??? LIST CLICK - UPDATE LABELS + COMBO ???
'========================================
Private Sub lstEntries_Click()
    If lstEntries.selectedItem Is Nothing Then Exit Sub
    
    Dim entryID As String, currentStatus As String
    entryID = lstEntries.selectedItem.text
    currentStatus = UCase(Trim(lstEntries.selectedItem.SubItems(2)))
    
    '? SHOW ENTRY ID & PRESENT STATUS ?
    lblEntryID.caption = "Entry ID: " & entryID
    lblPresentStatus.caption = "Present Status: " & currentStatus
    
    '? POPULATE COMBO WITH VALID NEXT STATUSES ?
    PopulateValidStatuses currentStatus
End Sub

Private Sub PopulateValidStatuses(currentStatus As String)
    cmbNewStatus.Clear
    
    Select Case currentStatus
        Case "ACTIVE"
            cmbNewStatus.AddItem "ASSIGNED"
            cmbNewStatus.AddItem "IN PROGRESS"
            
        Case "ASSIGNED"
            cmbNewStatus.AddItem "IN PROGRESS"
            cmbNewStatus.AddItem "COMPLETED"
            cmbNewStatus.AddItem "ACTIVE"  'Back to active if needed
            
        Case "IN PROGRESS"
            cmbNewStatus.AddItem "COMPLETED"
            cmbNewStatus.AddItem "READY"
            
        Case "COMPLETED"
            cmbNewStatus.AddItem "READY"
            cmbNewStatus.AddItem "IN PROGRESS"  'Back if needed
            
        Case "READY"
            cmbNewStatus.AddItem "DELIVERED"
            cmbNewStatus.AddItem "COMPLETED"  'Back if needed
            
        Case "DELIVERED"
            cmbNewStatus.AddItem "READY"
            cmbNewStatus.AddItem "COMPLETED"
            
        Case "RECEIVE_REPAIRED"
            cmbNewStatus.AddItem "COMPLETED"
            cmbNewStatus.AddItem "READY"
            
        Case Else
            cmbNewStatus.AddItem "ACTIVE"
            cmbNewStatus.AddItem "ASSIGNED"
            cmbNewStatus.AddItem "IN PROGRESS"
            cmbNewStatus.AddItem "COMPLETED"
            cmbNewStatus.AddItem "READY"
    End Select
    
    If cmbNewStatus.ListCount > 0 Then cmbNewStatus.ListIndex = 0
End Sub

'========================================
' UPDATE STATUS BUTTON
'========================================
Private Sub btnUpdateStatus_Click()
    If lstEntries.selectedItem Is Nothing Then
        MsgBox "Please select an entry first!", vbExclamation
        Exit Sub
    End If
    
    Dim entryID As String, currentStatus As String, newStatus As String
    entryID = lstEntries.selectedItem.text
    currentStatus = UCase(Trim(lstEntries.selectedItem.SubItems(2)))
    newStatus = UCase(Trim(cmbNewStatus.value))
    
    If currentStatus = newStatus Then
        MsgBox "New status is same as current status!", vbExclamation
        Exit Sub
    End If
    
    If MsgBox("Change status of " & entryID & "?" & vbCrLf & _
              "From: " & currentStatus & vbCrLf & _
              "To: " & newStatus, vbQuestion + vbYesNo, "Confirm") = vbNo Then Exit Sub
    
    'Update in Job_Product
    Dim ws As Worksheet, lastRow As Long, i As Long
    Set ws = ThisWorkbook.Sheets("Job_Product")
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    
    For i = 2 To lastRow
        If UCase(Trim(ws.Cells(i, 1).value)) = UCase(entryID) Then
            ws.Cells(i, 4).value = newStatus
            ws.Cells(i, 5).value = Format(Now, "dd-mm-yyyy hh:mm")
            Exit For
        End If
    Next i
    
    'Log
    LogStatusChange entryID, currentStatus, newStatus
    
    MsgBox "Status updated successfully!" & vbCrLf & entryID & " is now " & newStatus, vbInformation
    
    'Refresh
    LoadAllEntries
    lblEntryID.caption = "Entry ID: --"
    lblPresentStatus.caption = "Present Status: --"
    cmbNewStatus.Clear
End Sub

Private Sub LogStatusChange(entryID As String, oldStatus As String, newStatus As String)
    On Error Resume Next
    Dim ws As Worksheet, NextRow As Long
    Set ws = ThisWorkbook.Sheets("Status_Log")
    NextRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row + 1
    
    ws.Cells(NextRow, 1).value = Format(Now, "dd-mm-yyyy hh:mm:ss")
    ws.Cells(NextRow, 2).value = entryID
    ws.Cells(NextRow, 3).value = oldStatus
    ws.Cells(NextRow, 4).value = newStatus
    ws.Cells(NextRow, 5).value = Environ("Username")
End Sub

'========================================
' DOUBLE CLICK - OPEN APPROPRIATE FORM
'========================================
Private Sub lstEntries_DblClick()
    If lstEntries.selectedItem Is Nothing Then Exit Sub
    
    Dim entryID As String, entryType As String, status As String
    entryID = lstEntries.selectedItem.text
    entryType = UCase(lstEntries.selectedItem.SubItems(1))
    status = UCase(Trim(lstEntries.selectedItem.SubItems(2)))
    
    Me.Hide
    
    Select Case status
        Case "ACTIVE"
            frmEntryWizard.Tag = entryID
            frmEntryWizard.Show vbModal
            
        Case "ASSIGNED"
            If entryType = "WARRANTY" Then
                frmWarrantyAssignment.Tag = entryID
                frmWarrantyAssignment.Show vbModal
            Else
                frmServiceAssignment.Tag = entryID
                frmServiceAssignment.Show vbModal
            End If
            
        Case "IN PROGRESS"
            If entryType = "WARRANTY" Then
                frmWarrantyReturn.Tag = entryID
                frmWarrantyReturn.Show vbModal
            Else
                frmServiceWorkExpense.Tag = entryID
                frmServiceWorkExpense.Show vbModal
            End If
            
        Case "COMPLETED", "READY"
            frmDelivery.Tag = entryID
            frmDelivery.Show vbModal
            
        Case "DELIVERED"
            MsgBox "This entry is already delivered!", vbInformation
    End Select
    
    Me.Show
    LoadAllEntries
End Sub

Private Sub btnRefresh_Click()
    LoadAllEntries
    lblEntryID.caption = "Entry ID: --"
    lblPresentStatus.caption = "Present Status: --"
    cmbNewStatus.Clear
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

