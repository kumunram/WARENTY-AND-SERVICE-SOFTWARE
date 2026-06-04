VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAssignmentDashboard 
   Caption         =   "Assignment Dashboard"
   ClientHeight    =   13410
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   19365
   OleObjectBlob   =   "frmAssignmentDashboard.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmAssignmentDashboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnSearch_Click()
    LoadAssignmentData
End Sub

Private Sub Frame3_Click()
End Sub

Private Sub lblSummary_Click()
End Sub

'===========================================
' GLOBAL SOFT v1.0 - Assignment Dashboard
'===========================================

Private Sub UserForm_Initialize()
    Me.Width = 980
    Me.Height = 700
    Me.StartUpPosition = 1
    
    AddMinMaxButtons Me
    
    ConfigureListView
    PopulateFilterCombos
    LoadAssignmentData
    UpdateSummary
    
    Dim ctrl As Control
    For Each ctrl In Me.Controls
        If TypeName(ctrl) = "CommandButton" Then
            ctrl.BackColor = RGB(30, 144, 255)
            ctrl.ForeColor = RGB(255, 255, 255)
        End If
    Next ctrl
End Sub

Private Sub ConfigureListView()
    With lstAssignments
        .ListItems.Clear
        .ColumnHeaders.Clear
        .View = lvwReport
        .FullRowSelect = True
        .Gridlines = True

        ' === NEW ORDER: EntryID | Type | Status | Customer | Mobile | Product | Company | Model | Serial | AssignTo | Date | Days ===
        .ColumnHeaders.Add , , "EntryID", 70
        .ColumnHeaders.Add , , "Type", 80
        .ColumnHeaders.Add , , "Status", 120        ' <--- Status ???? ? ??? (3rd)
        .ColumnHeaders.Add , , "Customer Name", 150 ' <--- Customer 4th
        .ColumnHeaders.Add , , "Mobile", 80
        .ColumnHeaders.Add , , "Product", 130
        .ColumnHeaders.Add , , "Company", 100
        .ColumnHeaders.Add , , "Model", 100
        .ColumnHeaders.Add , , "Serial", 150
        .ColumnHeaders.Add , , "Assign To", 150
        .ColumnHeaders.Add , , "Assign Date", 90
        .ColumnHeaders.Add , , "Days", 60
    End With
End Sub

Private Sub lstAssignments_DblClick()
    Dim entryID As String
    Dim entryType As String
    Dim status As String
    Dim frm As Object
    
    If lstAssignments.selectedItem Is Nothing Then Exit Sub
    
    entryID = lstAssignments.selectedItem.text
    entryType = UCase(lstAssignments.selectedItem.SubItems(1))
    
    ' Normalize status: space/underscore ? dash
    status = UCase(Trim(lstAssignments.selectedItem.SubItems(2)))
    status = Replace(status, " ", "-")
    status = Replace(status, "_", "-")
    
    ' Accessory row check
    If Trim(entryID) = "" Then
        MsgBox "Please select a main entry!", vbExclamation
        Exit Sub
    End If
    
    Me.Hide
    
    Select Case status
        
        ' === ACTIVE = Entry Wizard (Edit) ===
        Case "ACTIVE"
            On Error Resume Next
            Set frm = New frmEntryWizard
            If Err.Number <> 0 Then
                MsgBox "frmEntryWizard not found!", vbCritical
                GoTo Cleanup
            End If
            On Error GoTo 0
            frm.Tag = entryID
            frm.Show vbModal
            Unload frm
            
        ' === ASSIGNED = Assignment Form (Edit Assign) ===
        Case "ASSIGNED"
            On Error Resume Next
            If entryType = "WARRANTY" Then
                Set frm = New frmWarrantyAssignment
            ElseIf entryType = "SERVICE" Then
                Set frm = New frmServiceAssignment
            Else
                MsgBox "Unknown Type: " & entryType, vbExclamation
                GoTo Cleanup
            End If
            If Err.Number <> 0 Then
                MsgBox "Assignment form not found!", vbCritical
                GoTo Cleanup
            End If
            On Error GoTo 0
            frm.Tag = entryID
            frm.Show vbModal
            Unload frm
            
        ' === IN-PROGRESS / RECEIVE-REPAIRED = Work/Return Form ===
        Case "IN-PROGRESS", "RECEIVE-REPAIRED"
            On Error Resume Next
            If entryType = "WARRANTY" Then
                Set frm = New frmWarrantyReturn
            ElseIf entryType = "SERVICE" Then
                Set frm = New frmServiceWorkExpense
            Else
                MsgBox "Unknown Type: " & entryType, vbExclamation
                GoTo Cleanup
            End If
            If Err.Number <> 0 Then
                MsgBox "Work form not found!", vbCritical
                GoTo Cleanup
            End If
            On Error GoTo 0
            frm.Tag = entryID
            frm.Show vbModal
            Unload frm
            
        ' === COMPLETED / READY / DELIVERED = Delivery Form ===
        Case "COMPLETED", "READY", "DELIVERED"
            On Error Resume Next
            Set frm = New frmDelivery
            If Err.Number <> 0 Then
                MsgBox "frmDelivery not found!", vbCritical
                GoTo Cleanup
            End If
            On Error GoTo 0
            frm.Tag = entryID
            frm.Show vbModal
            Unload frm
            
        Case Else
            MsgBox "Unknown Status: " & status & vbCrLf & _
                   "Please update Job_Product sheet status to standard values.", vbExclamation
    End Select
    
Cleanup:
    Set frm = Nothing
    Me.Show
    LoadAssignmentData
    UpdateSummary
End Sub

Private Sub PopulateFilterCombos()
    With cmbFilterStatus
        .Clear
        .AddItem "ALL"           ' ? ????? Non-Delivered (ACTIVE, ASSIGNED, IN-PROGRESS, READY)
        .AddItem "ACTIVE"
        .AddItem "ASSIGNED"
        .AddItem "IN-PROGRESS"
        .AddItem "READY"
        .AddItem "DELIVERED"     ' ? ??? ?? DELIVERED ????? ?? ???
        .ListIndex = 0
    End With
    
    With cmbFilterType
        .Clear
        .AddItem "ALL"
        .AddItem "WARRANTY"
        .AddItem "SERVICE"
        .ListIndex = 0
    End With
End Sub

Public Sub LoadAssignmentData()
    Dim wsEntry As Worksheet, wsAssign As Worksheet, wsCust As Worksheet, wsAcc As Worksheet
    Dim lastRowEntry As Long, lastRowAssign As Long, lastRowCust As Long, lastRowAcc As Long
    Dim i As Long, j As Long, k As Long, a As Long
    Dim listItem As Object, subItem As Object
    Dim entryID As String, entryType As String, status As String
    Dim custName As String, mobile As String, custID As String
    Dim product As String, company As String, model As String, serial As String
    Dim assignTo As String, assignDate As String, daysPending As Long
    Dim filterStatus As String, filterType As String, searchText As String
    
    On Error Resume Next
    Set wsEntry = ThisWorkbook.Sheets("Job_Product")
    Set wsAssign = ThisWorkbook.Sheets("Assign_Master")
    On Error GoTo 0
    If wsEntry Is Nothing Or wsAssign Is Nothing Then Exit Sub
    
    lastRowEntry = wsEntry.Cells(wsEntry.Rows.count, 1).End(xlUp).row
    lastRowAssign = wsAssign.Cells(wsAssign.Rows.count, 1).End(xlUp).row
    
    filterStatus = UCase(Trim(cmbFilterStatus.value))
    filterType = UCase(Trim(cmbFilterType.value))
    searchText = UCase(Trim(txtSearch.value))
    
    lstAssignments.ListItems.Clear
    If lastRowEntry < 2 Then Exit Sub
    
    Set wsCust = ThisWorkbook.Sheets("Customer_Master")
    lastRowCust = wsCust.Cells(wsCust.Rows.count, 1).End(xlUp).row
    
    For i = 2 To lastRowEntry
        entryID = Trim(wsEntry.Cells(i, 1).value)
        entryType = UCase(Trim(wsEntry.Cells(i, 3).value))
        status = UCase(Trim(wsEntry.Cells(i, 4).value))
        product = Trim(wsEntry.Cells(i, 6).value)
        company = Trim(wsEntry.Cells(i, 7).value)
        model = Trim(wsEntry.Cells(i, 8).value)
        serial = Trim(wsEntry.Cells(i, 9).value)
        
        ' Customer Lookup
        ' Customer Lookup - USES GLOBAL CONFIG (Auto-detected)
custID = Trim(wsEntry.Cells(i, 2).value)
Call GetCustomerData(custID, custName, mobile)
        
        ' ============================================
        ' FIXED FILTER LOGIC
        ' ============================================
        
        ' Type Filter
        If filterType <> "ALL" And entryType <> filterType Then GoTo NextRow
        
        ' Status Filter - FIXED!
        If filterStatus = "ALL" Then
            ' ALL = ????? Non-Delivered entries (ACTIVE, ASSIGNED, IN-PROGRESS, READY)
            If status = "DELIVERED" Then GoTo NextRow  ' ? DELIVERED skip!
        ElseIf filterStatus <> "ALL" And status <> filterStatus Then
            GoTo NextRow
        End If
        
        ' Search Filter
        If searchText <> "" Then
            If InStr(1, UCase(entryID & custName & product & company & model), searchText, vbTextCompare) = 0 Then GoTo NextRow
        End If
        
        ' Skip blank Entry_ID
        If entryID = "" Then GoTo NextRow
        
        ' ============================================
        ' REST OF CODE SAME AS BEFORE...
        ' ============================================
        
        ' Assign Lookup
        assignTo = ""
        assignDate = ""
        For j = 2 To lastRowAssign
            If UCase(Trim(wsAssign.Cells(j, 2).value)) = UCase(entryID) Then
                assignTo = Trim(wsAssign.Cells(j, 14).value)
                assignDate = wsAssign.Cells(j, 22).value
                Exit For
            End If
        Next j
        
        ' === ADD MAIN ROW ===
        Set listItem = lstAssignments.ListItems.Add(, , entryID)
        
        listItem.SubItems(1) = entryType
        listItem.SubItems(2) = status
        listItem.SubItems(3) = custName
        listItem.SubItems(4) = mobile
        listItem.SubItems(5) = product
        listItem.SubItems(6) = company
        listItem.SubItems(7) = model
        listItem.SubItems(8) = serial
        listItem.SubItems(9) = assignTo
        listItem.SubItems(10) = assignDate
        
        ' Days
        If IsDate(assignDate) Then
            listItem.SubItems(11) = Date - CDate(assignDate)
        Else
            listItem.SubItems(11) = ""
        End If
        
        ' Color coding
        Select Case status
            Case "ACTIVE"
                listItem.ForeColor = RGB(0, 150, 0)
                listItem.Bold = True
            Case "ASSIGNED"
                listItem.ForeColor = RGB(0, 0, 200)
                listItem.Bold = True
            Case "IN-PROGRESS"
                listItem.ForeColor = RGB(200, 100, 0)
                listItem.Bold = True
            Case "READY"
                listItem.ForeColor = RGB(150, 0, 150)
                listItem.Bold = True
            Case "DELIVERED"
                listItem.ForeColor = RGB(100, 100, 100)
                listItem.Bold = False
            Case Else
                listItem.ForeColor = RGB(0, 0, 0)
                listItem.Bold = False
        End Select
        
        ' === ACCESSORY ROWS ===
        On Error Resume Next
        Set wsAcc = ThisWorkbook.Sheets("Job_Accessory")
        lastRowAcc = wsAcc.Cells(wsAcc.Rows.count, 1).End(xlUp).row
        On Error GoTo 0
        
        If Not wsAcc Is Nothing Then
            For a = 2 To lastRowAcc
                If UCase(Trim(wsAcc.Cells(a, 1).value)) = UCase(entryID) Then
                    Set subItem = lstAssignments.ListItems.Add(, , "")
                    
                    subItem.SubItems(1) = ""
                    subItem.SubItems(2) = ""
                    subItem.SubItems(3) = ""
                    subItem.SubItems(4) = ""
                    subItem.SubItems(5) = "> " & wsAcc.Cells(a, 4).value
                    subItem.SubItems(6) = wsAcc.Cells(a, 5).value
                    subItem.SubItems(7) = wsAcc.Cells(a, 7).value
                    subItem.SubItems(8) = wsAcc.Cells(a, 6).value
                    subItem.SubItems(9) = ""
                    subItem.SubItems(10) = ""
                    subItem.SubItems(11) = ""
                    
                    subItem.ForeColor = RGB(0, 102, 153)
                End If
            Next a
        End If

NextRow:
    Next i
End Sub

Private Sub UpdateSummary()
    Dim wsEntry As Worksheet
    Dim lastRow As Long, i As Long
    Dim cntActive As Long, cntAssigned As Long, cntInProgress As Long, cntReady As Long
    Dim cntWarranty As Long, cntService As Long
    
    On Error Resume Next
    Set wsEntry = ThisWorkbook.Sheets("Job_Product")
    On Error GoTo 0
    If wsEntry Is Nothing Then
        lblSummary.caption = "Summary: Data not available"
        Exit Sub
    End If
    
    lastRow = wsEntry.Cells(wsEntry.Rows.count, 1).End(xlUp).row
    For i = 2 To lastRow
        Dim st As String, et As String
        st = UCase(Trim(wsEntry.Cells(i, 4).value))
        et = UCase(Trim(wsEntry.Cells(i, 3).value))
        Select Case st
            Case "ACTIVE": cntActive = cntActive + 1
            Case "ASSIGNED": cntAssigned = cntAssigned + 1
            Case "IN-PROGRESS": cntInProgress = cntInProgress + 1
            Case "READY": cntReady = cntReady + 1
        End Select
        Select Case et
            Case "WARRANTY": cntWarranty = cntWarranty + 1
            Case "SERVICE": cntService = cntService + 1
        End Select
    Next i
    
    lblSummary.caption = "ACTIVE: " & cntActive & _
                     "    ASSIGNED: " & cntAssigned & _
                     "    IN-PROGRESS: " & cntInProgress & _
                     "    READY: " & cntReady & _
                     "    WARRANTY: " & cntWarranty & _
                     "    SERVICE: " & cntService
End Sub

' Event handlers
Private Sub cmbFilterStatus_Change()
    LoadAssignmentData
End Sub

Private Sub cmbFilterType_Change()
    LoadAssignmentData
End Sub

Private Sub txtSearch_Change()
    LoadAssignmentData
End Sub

'===========================================
' BUTTON CLICK EVENTS
'===========================================

Private Sub btnNewWarrantyAssign_Click()
    Dim entryID As String, entryType As String, currentStatus As String
    
    ' === KUCH BHI SELECT NAHI HAI ? Blank form khule ===
    If lstAssignments.selectedItem Is Nothing Then
        Me.Hide
        frmWarrantyAssignment.Show vbModal
        Me.Show
        LoadAssignmentData
        UpdateSummary
        Exit Sub
    End If
    
    entryID = lstAssignments.selectedItem.text
    entryType = UCase(lstAssignments.selectedItem.SubItems(1))
    currentStatus = UCase(Trim(lstAssignments.selectedItem.SubItems(2)))
    
    ' === VALIDATION: Sirf WARRANTY + ACTIVE allowed ===
    If entryType <> "WARRANTY" Then
        MsgBox "Please select a WARRANTY entry!" & vbCrLf & _
               "Current selection: " & entryType, vbExclamation, "GLOBAL SOFT"
        Exit Sub
    End If
    
    If currentStatus <> "ACTIVE" Then
        MsgBox "Entry must be ACTIVE to create Warranty Assignment!" & vbCrLf & _
               "Current Status: " & currentStatus, vbExclamation, "GLOBAL SOFT"
        Exit Sub
    End If
    
    ' === OPEN FORM WITH ENTRYID ===
    Me.Hide
    frmWarrantyAssignment.Tag = entryID     ' ? YE LINE SABSE IMPORTANT HAI
    frmWarrantyAssignment.Show vbModal
    Me.Show
    LoadAssignmentData
    UpdateSummary
End Sub

Private Sub btnNewServiceAssign_Click()
    Dim entryID As String, entryType As String, currentStatus As String
    
    ' Validation
    If lstAssignments.selectedItem Is Nothing Then
        MsgBox "Please select an entry first!", vbExclamation, "GLOBAL SOFT"
        Exit Sub
    End If
    
    entryID = lstAssignments.selectedItem.text
    entryType = UCase(lstAssignments.selectedItem.SubItems(1))
    currentStatus = UCase(Trim(lstAssignments.selectedItem.SubItems(2)))
    
    ' Must be SERVICE
    If entryType <> "SERVICE" Then
        MsgBox "This button is only for SERVICE entries!" & vbCrLf & _
               "Selected Type: " & entryType, vbExclamation, "GLOBAL SOFT"
        Exit Sub
    End If
    
    ' Must be ACTIVE
    If currentStatus <> "ACTIVE" Then
        MsgBox "Entry must be ACTIVE to create new Service Assignment!" & vbCrLf & _
               "Current Status: " & currentStatus, vbExclamation, "GLOBAL SOFT"
        Exit Sub
    End If
    
    Me.Hide
    frmServiceAssignment.Show vbModal
    Me.Show
    LoadAssignmentData
    UpdateSummary
End Sub

Private Sub btnViewDetails_Click()
    Dim entryID As String, entryType As String
    If lstAssignments.selectedItem Is Nothing Then
        MsgBox "Please select an entry first!", vbExclamation
        Exit Sub
    End If
    
    entryID = lstAssignments.selectedItem.text
    entryType = lstAssignments.selectedItem.SubItems(1)
    
    MsgBox "=== ENTRY DETAILS ===" & vbCrLf & vbCrLf & _
           "Entry ID: " & entryID & vbCrLf & _
           "Type: " & entryType & vbCrLf & _
           "Status: " & lstAssignments.selectedItem.SubItems(2) & vbCrLf & _
           "Customer: " & lstAssignments.selectedItem.SubItems(3) & vbCrLf & _
           "Mobile: " & lstAssignments.selectedItem.SubItems(4) & vbCrLf & _
           "Product: " & lstAssignments.selectedItem.SubItems(5) & vbCrLf & _
           "Company: " & lstAssignments.selectedItem.SubItems(6) & vbCrLf & _
           "Model: " & lstAssignments.selectedItem.SubItems(7) & vbCrLf & _
           "Serial: " & lstAssignments.selectedItem.SubItems(8) & vbCrLf & _
           "Assigned To: " & lstAssignments.selectedItem.SubItems(9), _
           vbInformation, "View Detail - " & entryID
End Sub

Private Sub btnUpdateStatus_Click()
    Dim entryID As String, currentStatus As String
    If lstAssignments.selectedItem Is Nothing Then
        MsgBox "Please select an entry first!", vbExclamation
        Exit Sub
    End If
    
    entryID = lstAssignments.selectedItem.text
    currentStatus = UCase(Trim(lstAssignments.selectedItem.SubItems(2))) ' <--- Changed from 8 to 2
    
    If currentStatus <> "IN-PROGRESS" Then
        MsgBox "Only IN-PROGRESS entries can be marked as READY!" & vbCrLf & _
               "Current status: " & currentStatus, vbExclamation
        Exit Sub
    End If
    
    If MsgBox("Mark this entry as READY?" & vbCrLf & "Entry ID: " & entryID, vbQuestion + vbYesNo) = vbNo Then Exit Sub
    
    Dim ws As Worksheet, lastRow As Long, i As Long
    Set ws = ThisWorkbook.Sheets("Job_Product")
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    For i = 2 To lastRow
        If UCase(Trim(ws.Cells(i, 1).value)) = UCase(entryID) Then
            ws.Cells(i, 4).value = "READY"
            ws.Cells(i, 5).value = Format(Now, "dd-mm-yyyy hh:mm")
            Exit For
        End If
    Next i
    
    MsgBox "Status updated to READY successfully!", vbInformation
    LoadAssignmentData
    UpdateSummary
End Sub

Private Sub btnReceiveReturn_Click()
    Dim entryID As String, entryType As String, currentStatus As String
    
    If lstAssignments.selectedItem Is Nothing Then
        MsgBox "Please select an entry first!", vbExclamation
        Exit Sub
    End If
    
    entryID = lstAssignments.selectedItem.text
    entryType = UCase(lstAssignments.selectedItem.SubItems(1))
    currentStatus = UCase(Trim(lstAssignments.selectedItem.SubItems(2)))
    
    ' === ONLY FOR WARRANTY ===
    If entryType <> "WARRANTY" Then
        MsgBox "This button is only for WARRANTY entries!" & vbCrLf & _
               "Selected Type: " & entryType & vbCrLf & vbCrLf & _
               "For SERVICE entries, please use 'Add Expense' button.", vbExclamation, "GLOBAL SOFT - Wrong Button"
        Exit Sub
    End If
    
    ' === ONLY ASSIGNED or IN-PROGRESS ===
    If currentStatus <> "ASSIGNED" And currentStatus <> "IN-PROGRESS" Then
        MsgBox "Warranty Return is only for ASSIGNED or IN-PROGRESS entries!" & vbCrLf & _
               "Current Status: " & currentStatus, vbExclamation, "GLOBAL SOFT"
        Exit Sub
    End If
    
    Me.Hide
    frmWarrantyReturn.Tag = entryID
    frmWarrantyReturn.Show vbModal
    Me.Show
    LoadAssignmentData
    UpdateSummary
End Sub
Private Sub btnAddExpense_Click()
    Dim entryID As String, entryType As String, currentStatus As String
    
    If lstAssignments.selectedItem Is Nothing Then
        MsgBox "Please select an entry first!", vbExclamation
        Exit Sub
    End If
    
    entryID = lstAssignments.selectedItem.text
    entryType = UCase(lstAssignments.selectedItem.SubItems(1))
    currentStatus = UCase(Trim(lstAssignments.selectedItem.SubItems(2)))
    
    ' === ONLY FOR SERVICE ===
    If entryType <> "SERVICE" Then
        MsgBox "This button is only for SERVICE entries!" & vbCrLf & _
               "Selected Type: " & entryType, vbExclamation, "GLOBAL SOFT"
        Exit Sub
    End If
    
    ' === ONLY ASSIGNED or IN-PROGRESS ===
    If currentStatus <> "ASSIGNED" And currentStatus <> "IN-PROGRESS" Then
        MsgBox "Service Work Expense is only for ASSIGNED or IN-PROGRESS entries!" & vbCrLf & _
               "Current Status: " & currentStatus, vbExclamation, "GLOBAL SOFT"
        Exit Sub
    End If
    
    Me.Hide
    frmServiceWorkExpense.Tag = entryID
    frmServiceWorkExpense.Show vbModal
    Me.Show
    LoadAssignmentData
    UpdateSummary
End Sub
Private Sub btnDelivery_Click()
    Dim entryID As String, currentStatus As String
    If lstAssignments.selectedItem Is Nothing Then
        MsgBox "Please select an entry first!", vbExclamation
        Exit Sub
    End If
    
    entryID = lstAssignments.selectedItem.text
    currentStatus = UCase(Trim(lstAssignments.selectedItem.SubItems(2))) ' <--- Changed from 8 to 2
    
    If currentStatus <> "READY" Then
        MsgBox "Only READY entries can be delivered!" & vbCrLf & _
               "Current status: " & currentStatus, vbExclamation
        Exit Sub
    End If
    
    Me.Hide
    frmDelivery.Tag = entryID
    frmDelivery.Show vbModal
    Me.Show
    LoadAssignmentData
    UpdateSummary
End Sub

Private Sub btnDeliver_Click()
    Dim entryID As String, currentStatus As String
    If lstAssignments.selectedItem Is Nothing Then
        MsgBox "Please select an entry first!", vbExclamation
        Exit Sub
    End If
    
    entryID = lstAssignments.selectedItem.text
    currentStatus = UCase(Trim(lstAssignments.selectedItem.SubItems(2))) ' <--- Changed from 8 to 2
    
    If currentStatus <> "READY" Then
        MsgBox "Only READY entries can be delivered!" & vbCrLf & _
               "Current status: " & currentStatus, vbExclamation
        Exit Sub
    End If
    
    If MsgBox("Mark this entry as DELIVERED?" & vbCrLf & "Entry ID: " & entryID, vbQuestion + vbYesNo) = vbNo Then Exit Sub
    
    Dim ws As Worksheet, wsDel As Worksheet, lastRow As Long, i As Long, delRow As Long
    Set ws = ThisWorkbook.Sheets("Job_Product")
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    For i = 2 To lastRow
        If UCase(Trim(ws.Cells(i, 1).value)) = UCase(entryID) Then
            ws.Cells(i, 4).value = "DELIVERED"
            ws.Cells(i, 5).value = Format(Now, "dd-mm-yyyy hh:mm")
            Exit For
        End If
    Next i
    
    On Error Resume Next
    Set wsDel = ThisWorkbook.Sheets("Delivery_Master")
    On Error GoTo 0
    If Not wsDel Is Nothing Then
        delRow = wsDel.Cells(wsDel.Rows.count, 1).End(xlUp).row + 1
        wsDel.Cells(delRow, 1).value = "DEL" & Format(delRow, "00000")
        wsDel.Cells(delRow, 2).value = entryID
        wsDel.Cells(delRow, 3).value = lstAssignments.selectedItem.SubItems(4) ' Mobile now at 4
        wsDel.Cells(delRow, 4).value = Format(Now, "dd-mm-yyyy")
        wsDel.Cells(delRow, 5).value = Format(Now, "hh:mm")
        wsDel.Cells(delRow, 6).value = "COMPLETED"
        wsDel.Cells(delRow, 7).value = "System"
    End If
    
    MsgBox "Entry marked as DELIVERED successfully!", vbInformation
    LoadAssignmentData
    UpdateSummary
End Sub

Private Sub btnPrintReport_Click()
    Dim wsTemp As Worksheet, wsConfig As Worksheet
    Dim i As Long, row As Long
    Dim filterType As String, filterStatus As String, reportTitle As String
    Dim companyName As String, companyAddress As String, companyMobile As String
    Dim totalEntries As Long, serialNum As Long, printChoice As Integer
    Dim statusColor As Long
    
    totalEntries = 0
    On Error Resume Next
    totalEntries = lstAssignments.ListItems.count
    On Error GoTo 0
    If totalEntries = 0 Then
        MsgBox "No data to print!", vbExclamation, "No Data"
        Exit Sub
    End If
    
    filterType = UCase(Trim(cmbFilterType.value))
    filterStatus = UCase(Trim(cmbFilterStatus.value))
    
    If filterType = "ALL" And filterStatus = "ALL" Then
        reportTitle = "ALL ENTRIES REPORT"
    ElseIf filterType = "ALL" Then
        reportTitle = "ALL " & filterStatus & " ENTRIES REPORT"
    ElseIf filterStatus = "ALL" Then
        reportTitle = "ALL " & filterType & " ENTRIES REPORT"
    Else
        reportTitle = filterType & " - " & filterStatus & " REPORT"
    End If
    
    On Error Resume Next
    Set wsConfig = ThisWorkbook.Sheets("Software_Config")
    companyName = Trim(wsConfig.Range("B2").value)
    companyAddress = Trim(wsConfig.Range("B3").value)
    companyMobile = Trim(wsConfig.Range("B4").value)
    On Error GoTo 0
    
    If companyName = "" Then companyName = "GLOBAL IT SOLUTIONS"
    If companyAddress = "" Then companyAddress = "Your Company Address"
    If companyMobile = "" Then companyMobile = "Your Contact Number"
    
    printChoice = MsgBox("Print Report: " & reportTitle & vbCrLf & _
              "Total Entries: " & totalEntries & vbCrLf & vbCrLf & _
              "Click YES to Print Now" & vbCrLf & _
              "Click NO to View in Excel", vbQuestion + vbYesNoCancel, "Print Report")
    
    If printChoice = vbCancel Then Exit Sub
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Sheets("PrintReport").Delete
    On Error GoTo 0
    Set wsTemp = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
    wsTemp.name = "PrintReport"
    Application.DisplayAlerts = True
    
    With wsTemp
        ' Header
        .Range("A1").value = UCase(companyName)
        .Range("A1").Font.name = "Arial": .Range("A1").Font.Size = 16: .Range("A1").Font.Bold = True
        .Range("A1").HorizontalAlignment = xlCenter
        .Range("A2").value = companyAddress
        .Range("A2").Font.name = "Arial": .Range("A2").Font.Size = 10
        .Range("A2").HorizontalAlignment = xlCenter
        .Range("A3").value = "Mobile: " & companyMobile
        .Range("A3").Font.name = "Arial": .Range("A3").Font.Size = 10
        .Range("A3").HorizontalAlignment = xlCenter
        .Range("A1:L1").Merge: .Range("A2:L2").Merge: .Range("A3:L3").Merge
        
        ' Title
        .Range("A5").value = reportTitle
        .Range("A5").Font.name = "Arial": .Range("A5").Font.Size = 14: .Range("A5").Font.Bold = True
        .Range("A5").HorizontalAlignment = xlCenter
        .Range("A5:L5").Merge
        
        ' Date
        .Range("A6").value = "Print Date: " & Format(Now, "dd-mmm-yyyy hh:mm AM/PM")
        .Range("A6").Font.name = "Arial": .Range("A6").Font.Size = 9
        .Range("A6").HorizontalAlignment = xlRight
        .Range("A6:L6").Merge
        
        ' === COLUMN HEADERS - NEW ORDER ===
        row = 8
        .Cells(row, 1).value = "S.No"
        .Cells(row, 2).value = "Entry ID"
        .Cells(row, 3).value = "Type"
        .Cells(row, 4).value = "Status"         ' <--- Moved here
        .Cells(row, 5).value = "Customer Name"
        .Cells(row, 6).value = "Mobile"
        .Cells(row, 7).value = "Product"
        .Cells(row, 8).value = "Company"
        .Cells(row, 9).value = "Model"
        .Cells(row, 10).value = "Serial"
        .Cells(row, 11).value = "Assigned To"
        .Cells(row, 12).value = "Days"
        
        With .Range(.Cells(row, 1), .Cells(row, 12))
            .Font.name = "Arial": .Font.Bold = True: .Font.Size = 10
            .Interior.Color = RGB(220, 220, 220)
            .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous: .Borders.Weight = xlMedium
        End With
        
        ' === DATA ROWS - NEW SUBITEMS MAPPING ===
        serialNum = 0
        For i = 1 To lstAssignments.ListItems.count
            row = row + 1
            
            If Trim(lstAssignments.ListItems(i).text) <> "" Then
                serialNum = serialNum + 1
                .Cells(row, 1).value = serialNum
            Else
                .Cells(row, 1).value = ""
            End If
            
            .Cells(row, 2).value = lstAssignments.ListItems(i).text
            .Cells(row, 2).HorizontalAlignment = xlCenter
            
            .Cells(row, 3).value = lstAssignments.ListItems(i).SubItems(1)   ' Type
            .Cells(row, 4).value = lstAssignments.ListItems(i).SubItems(2)   ' Status (was 8)
            .Cells(row, 5).value = lstAssignments.ListItems(i).SubItems(3)   ' Customer (was 2)
            .Cells(row, 6).value = lstAssignments.ListItems(i).SubItems(4)   ' Mobile (was 3)
            .Cells(row, 7).value = lstAssignments.ListItems(i).SubItems(5)   ' Product (was 4)
            .Cells(row, 8).value = lstAssignments.ListItems(i).SubItems(6)   ' Company (was 5)
            .Cells(row, 9).value = lstAssignments.ListItems(i).SubItems(7)   ' Model (was 6)
            .Cells(row, 10).value = lstAssignments.ListItems(i).SubItems(8)  ' Serial (was 7)
            .Cells(row, 11).value = lstAssignments.ListItems(i).SubItems(9)  ' Assign To
            .Cells(row, 12).value = lstAssignments.ListItems(i).SubItems(11) ' Days
            
            ' Alignments
            .Cells(row, 3).HorizontalAlignment = xlCenter  ' Type
            .Cells(row, 4).HorizontalAlignment = xlCenter  ' Status
            .Cells(row, 6).HorizontalAlignment = xlCenter  ' Mobile
            .Cells(row, 12).HorizontalAlignment = xlCenter ' Days
            
            ' Status Color
            Select Case UCase(Trim(lstAssignments.ListItems(i).SubItems(2))) ' <--- Changed from 8 to 2
                Case "ACTIVE": statusColor = RGB(200, 255, 200)
                Case "ASSIGNED": statusColor = RGB(200, 200, 255)
                Case "IN-PROGRESS": statusColor = RGB(255, 230, 200)
                Case "READY": statusColor = RGB(255, 200, 255)
                Case "DELIVERED": statusColor = RGB(230, 230, 230)
                Case Else: statusColor = RGB(255, 255, 255)
            End Select
            .Cells(row, 4).Interior.Color = statusColor ' <--- Column 4 now
            
            ' Borders
            With .Range(.Cells(row, 1), .Cells(row, 12))
                .Borders.LineStyle = xlContinuous: .Borders.Weight = xlThin
                .Font.name = "Arial": .Font.Size = 9
            End With
            
            ' Accessory formatting
            If Trim(lstAssignments.ListItems(i).text) = "" Then
                .Range(.Cells(row, 1), .Cells(row, 12)).Font.Italic = True
                ' Accessory data now in columns 7-10 (Product, Company, Model, Serial)
                .Range(.Cells(row, 7), .Cells(row, 10)).Font.Color = RGB(0, 102, 153)
            End If
            
            .Rows(row).AutoFit
        Next i
        
        ' Footer
        row = row + 2
        .Cells(row, 1).value = "Total Main Entries: " & serialNum
        .Cells(row, 1).Font.name = "Arial": .Cells(row, 1).Font.Bold = True: .Cells(row, 1).Font.Size = 10
        .Range("A" & row & ":L" & row).Merge
        
        ' Column Widths
        .Columns(1).ColumnWidth = 6
        .Columns(2).ColumnWidth = 12
        .Columns(3).ColumnWidth = 10
        .Columns(4).ColumnWidth = 10   ' Status
        .Columns(5).ColumnWidth = 22   ' Customer
        .Columns(6).ColumnWidth = 12   ' Mobile
        .Columns(7).ColumnWidth = 18   ' Product
        .Columns(8).ColumnWidth = 12   ' Company
        .Columns(9).ColumnWidth = 12   ' Model
        .Columns(10).ColumnWidth = 15  ' Serial
        .Columns(11).ColumnWidth = 15  ' Assign To
        .Columns(12).ColumnWidth = 6   ' Days
        
        ' Page Setup
        On Error Resume Next
        With .PageSetup
            .Orientation = xlLandscape
            .PaperSize = xlPaperA4
            .Zoom = False
            .FitToPagesWide = 1
            .FitToPagesTall = False
            .CenterHorizontally = True
            .LeftMargin = Application.InchesToPoints(0.5)
            .RightMargin = Application.InchesToPoints(0.5)
            .TopMargin = Application.InchesToPoints(0.5)
            .BottomMargin = Application.InchesToPoints(0.5)
            .HeaderMargin = Application.InchesToPoints(0.3)
            .FooterMargin = Application.InchesToPoints(0.3)
        End With
        On Error GoTo 0
    End With
    
    Application.ScreenUpdating = True
    
    If printChoice = vbNo Then
        wsTemp.Activate
        MsgBox "Report ready in 'PrintReport' sheet.", vbInformation
        Exit Sub
    ElseIf printChoice = vbYes Then
        On Error Resume Next
        wsTemp.PrintOut Copies:=1, Collate:=True, IgnorePrintAreas:=False
        If Err.Number <> 0 Then
            MsgBox "Print Error: " & Err.Description, vbExclamation
            wsTemp.Activate
            Exit Sub
        End If
        On Error GoTo 0
        MsgBox "Report printed successfully!", vbInformation
        Application.DisplayAlerts = False
        wsTemp.Delete
        Application.DisplayAlerts = True
    End If
End Sub

Private Sub btnRefresh_Click()
    LoadAssignmentData
    UpdateSummary
    MsgBox "Data refreshed!", vbInformation
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

