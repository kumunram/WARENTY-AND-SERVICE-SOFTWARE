VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmReports 
   Caption         =   "UserForm1"
   ClientHeight    =   12780
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   20265
   OleObjectBlob   =   "frmReports.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'========================================
' API DECLARATIONS
'========================================
Private Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
    ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
    ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private m_CurrentCategory As String
Private m_IsDataLoaded As Boolean

'========================================
' FORM INITIALIZE
'========================================
Private Sub UserForm_Initialize()
    On Error Resume Next
    AddMinMaxButtons Me
    Me.caption = "GLOBAL SOFT - REPORTS & ANALYTICS"
    Me.Width = 1030
    Me.Height = 670
    Me.StartUpPosition = 1
    
    AddMinMaxButtons Me
    
    m_CurrentCategory = "MASTER"
    m_IsDataLoaded = False
    
    Call SafeSetupListView
    Call SafeSetupFilters
    Call SafeLoadCustomers
    
    txtDateFrom.text = "01-04-2026"
    txtDateTo.text = "31-12-2026"
    
    Call HighlightCategory("MASTER")
    Call SetupMasterColumns
    
    lblStatus.caption = "Ready - Select Category & Click SHOW"
    
    On Error GoTo 0
End Sub

'========================================
' SAFE SETUP LISTVIEW
'========================================
Private Sub SafeSetupListView()
    On Error Resume Next
    With lstData
        .View = lvwReport
        .FullRowSelect = True
        .Gridlines = True
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "S.No", 50
        .ColumnHeaders.Add , , "Entry ID", 90
        .ColumnHeaders.Add , , "Date", 90
        .ColumnHeaders.Add , , "Customer", 140
        .ColumnHeaders.Add , , "Mobile", 100
        .ColumnHeaders.Add , , "Product", 120
        .ColumnHeaders.Add , , "Company", 100
        .ColumnHeaders.Add , , "Status", 100
        .ColumnHeaders.Add , , "Amount", 90
    End With
End Sub

'========================================
' SAFE SETUP FILTERS
'========================================
Private Sub SafeSetupFilters()
    On Error Resume Next
    
    cmbEntryType.Clear
    cmbEntryType.AddItem "All"
    cmbEntryType.AddItem "WARRANTY"
    cmbEntryType.AddItem "SERVICE"
    If cmbEntryType.ListCount > 0 Then cmbEntryType.ListIndex = 0
    
    cmbStatus.Clear
    cmbStatus.AddItem "All"
    cmbStatus.AddItem "ACTIVE"
    cmbStatus.AddItem "ASSIGNED"
    cmbStatus.AddItem "PENDING"
    cmbStatus.AddItem "IN PROGRESS"
    cmbStatus.AddItem "COMPLETED"
    cmbStatus.AddItem "DELIVERED"
    cmbStatus.AddItem "DELETED"
    If cmbStatus.ListCount > 0 Then cmbStatus.ListIndex = 0
    
    cmbAssignType.Clear
    cmbAssignType.AddItem "All"
    cmbAssignType.AddItem "VENDOR"
    cmbAssignType.AddItem "ENGINEER"
    cmbAssignType.AddItem "INHOUSE"
    cmbAssignType.AddItem "SERVICE_STATION"
    If cmbAssignType.ListCount > 0 Then cmbAssignType.ListIndex = 0
    
    cmbPayment.Clear
    cmbPayment.AddItem "All"
    cmbPayment.AddItem "PAID"
    cmbPayment.AddItem "PARTIAL"
    cmbPayment.AddItem "PENDING"
    cmbPayment.AddItem "DISCOUNT"
    If cmbPayment.ListCount > 0 Then cmbPayment.ListIndex = 0
    
    On Error GoTo 0
End Sub

'========================================
' SAFE LOAD CUSTOMERS
'========================================
Private Sub SafeLoadCustomers()
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Customer_Master")
    On Error GoTo 0
    
    cmbCustomer.Clear
    cmbCustomer.AddItem "All Customers"
    
    If ws Is Nothing Then
        If cmbCustomer.ListCount > 0 Then cmbCustomer.ListIndex = 0
        Exit Sub
    End If
    
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    If lastRow < 2 Then
        If cmbCustomer.ListCount > 0 Then cmbCustomer.ListIndex = 0
        Exit Sub
    End If
    
    For i = 2 To lastRow
        Dim cName As String, cMobile As String
        cName = Trim(ws.Cells(i, 2).value & "")
        cMobile = Trim(ws.Cells(i, 3).value & "")
        If cName <> "" Then cmbCustomer.AddItem cName & " | " & cMobile
    Next i
    
    If cmbCustomer.ListCount > 0 Then cmbCustomer.ListIndex = 0
End Sub

'========================================
' CATEGORY BUTTONS
'========================================
Private Sub btnMaster_Click()
    Call ClearReportData
    m_CurrentCategory = "MASTER"
    Call HighlightCategory("MASTER")
    Call SetupMasterColumns
    lblStatus.caption = "Master Report - Click SHOW"
End Sub

Private Sub btnJob_Click()
    Call ClearReportData
    m_CurrentCategory = "JOB"
    Call HighlightCategory("JOB")
    Call SetupJobColumns
    lblStatus.caption = "Job/Entry Report - Click SHOW"
End Sub

Private Sub btnAssign_Click()
    Call ClearReportData
    m_CurrentCategory = "ASSIGN"
    Call HighlightCategory("ASSIGN")
    Call SetupAssignColumns
    lblStatus.caption = "Assignment Report - Click SHOW"
End Sub

Private Sub btnFinance_Click()
    Call ClearReportData
    m_CurrentCategory = "FINANCE"
    Call HighlightCategory("FINANCE")
    Call SetupFinanceColumns
    lblStatus.caption = "Finance Report - Click SHOW"
End Sub

Private Sub btnDelivery_Click()
    Call ClearReportData
    m_CurrentCategory = "DELIVERY"
    Call HighlightCategory("DELIVERY")
    Call SetupDeliveryColumns
    lblStatus.caption = "Delivery Report - Click SHOW"
End Sub

Private Sub btnSummary_Click()
    Call ClearReportData
    m_CurrentCategory = "SUMMARY"
    Call HighlightCategory("SUMMARY")
    Call SetupSummaryColumns
    lblStatus.caption = "Summary Report - Click SHOW"
End Sub

Private Sub btnSmart_Click()
    Call ClearReportData
    m_CurrentCategory = "SMART"
    Call HighlightCategory("SMART")
    Call SetupSmartColumns
    lblStatus.caption = "Smart Alerts - Click SHOW"
End Sub

'========================================
' AUTO CLEAR
'========================================
Private Sub ClearReportData()
    On Error Resume Next
    lstData.ListItems.Clear
    Call UpdateSummary(0, 0, 0)
    m_IsDataLoaded = False
    On Error GoTo 0
End Sub

'========================================
' HIGHLIGHT ACTIVE CATEGORY
'========================================
Private Sub HighlightCategory(cat As String)
    On Error Resume Next
    btnMaster.BackColor = RGB(240, 240, 240)
    btnJob.BackColor = RGB(240, 240, 240)
    btnAssign.BackColor = RGB(240, 240, 240)
    btnFinance.BackColor = RGB(240, 240, 240)
    btnDelivery.BackColor = RGB(240, 240, 240)
    btnSummary.BackColor = RGB(240, 240, 240)
    btnSmart.BackColor = RGB(240, 240, 240)
    
    Select Case cat
        Case "MASTER":  btnMaster.BackColor = RGB(26, 188, 156)
        Case "JOB":     btnJob.BackColor = RGB(243, 156, 18)
        Case "ASSIGN":  btnAssign.BackColor = RGB(155, 89, 182)
        Case "FINANCE": btnFinance.BackColor = RGB(52, 152, 219)
        Case "DELIVERY": btnDelivery.BackColor = RGB(231, 76, 60)
        Case "SUMMARY": btnSummary.BackColor = RGB(39, 174, 96)
        Case "SMART":   btnSmart.BackColor = RGB(142, 68, 173)
    End Select
    On Error GoTo 0
End Sub

'========================================
' COLUMN SETUP
'========================================
Private Sub SetupMasterColumns()
    On Error Resume Next
    With lstData
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "S.No", 50
        .ColumnHeaders.Add , , "Customer ID", 90
        .ColumnHeaders.Add , , "Name", 150
        .ColumnHeaders.Add , , "Mobile", 100
        .ColumnHeaders.Add , , "Address", 200
        .ColumnHeaders.Add , , "Email", 150
        .ColumnHeaders.Add , , "GST", 120
        .ColumnHeaders.Add , , "Total Jobs", 90
    End With
End Sub

Private Sub SetupJobColumns()
    On Error Resume Next
    With lstData
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "S.No", 50
        .ColumnHeaders.Add , , "Entry ID", 100
        .ColumnHeaders.Add , , "Date", 90
        .ColumnHeaders.Add , , "Customer", 140
        .ColumnHeaders.Add , , "Mobile", 100
        .ColumnHeaders.Add , , "Type", 80
        .ColumnHeaders.Add , , "Product", 120
        .ColumnHeaders.Add , , "Company", 100
        .ColumnHeaders.Add , , "Model", 100
        .ColumnHeaders.Add , , "Serial", 120
        .ColumnHeaders.Add , , "Assign Type", 100
        .ColumnHeaders.Add , , "Status", 100
    End With
End Sub

Private Sub SetupAssignColumns()
    On Error Resume Next
    With lstData
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "S.No", 50
        .ColumnHeaders.Add , , "Entry ID", 100
        .ColumnHeaders.Add , , "Assign Date", 90
        .ColumnHeaders.Add , , "Customer", 140
        .ColumnHeaders.Add , , "Product", 120
        .ColumnHeaders.Add , , "Assigned To", 120
        .ColumnHeaders.Add , , "Name", 140
        .ColumnHeaders.Add , , "Status", 100
        .ColumnHeaders.Add , , "Exp Return", 90
    End With
End Sub

Private Sub SetupFinanceColumns()
    On Error Resume Next
    With lstData
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "S.No", 50
        .ColumnHeaders.Add , , "Payment ID", 100
        .ColumnHeaders.Add , , "Date", 90
        .ColumnHeaders.Add , , "Entry ID", 100
        .ColumnHeaders.Add , , "Customer", 140
        .ColumnHeaders.Add , , "Category", 100
        .ColumnHeaders.Add , , "Pay Mode", 100
        .ColumnHeaders.Add , , "Amount", 100
        .ColumnHeaders.Add , , "Status", 100
    End With
End Sub

Private Sub SetupDeliveryColumns()
    On Error Resume Next
    With lstData
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "S.No", 50
        .ColumnHeaders.Add , , "Entry ID", 100
        .ColumnHeaders.Add , , "Customer", 140
        .ColumnHeaders.Add , , "Product", 120
        .ColumnHeaders.Add , , "Assigned To", 120
        .ColumnHeaders.Add , , "Send Date", 90
        .ColumnHeaders.Add , , "Exp Return", 90
        .ColumnHeaders.Add , , "Days", 80
        .ColumnHeaders.Add , , "Status", 100
    End With
End Sub

Private Sub SetupSummaryColumns()
    On Error Resume Next
    With lstData
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "S.No", 50
        .ColumnHeaders.Add , , "Date/Period", 120
        .ColumnHeaders.Add , , "Total", 80
        .ColumnHeaders.Add , , "Warranty", 80
        .ColumnHeaders.Add , , "Service", 80
        .ColumnHeaders.Add , , "Completed", 80
        .ColumnHeaders.Add , , "Pending", 80
        .ColumnHeaders.Add , , "Collection", 100
    End With
End Sub

Private Sub SetupSmartColumns()
    On Error Resume Next
    With lstData
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "S.No", 50
        .ColumnHeaders.Add , , "Alert Type", 150
        .ColumnHeaders.Add , , "Entry ID", 100
        .ColumnHeaders.Add , , "Customer", 140
        .ColumnHeaders.Add , , "Product", 120
        .ColumnHeaders.Add , , "Issue", 200
        .ColumnHeaders.Add , , "Days", 60
        .ColumnHeaders.Add , , "Action", 100
    End With
End Sub

'========================================
' SHOW BUTTON
'========================================
Private Sub btnShow_Click()
    On Error Resume Next
    lstData.ListItems.Clear
    m_IsDataLoaded = False
    
    Select Case m_CurrentCategory
        Case "MASTER":  Call LoadMasterReport
        Case "JOB":     Call LoadJobReport
        Case "ASSIGN":  Call LoadAssignReport
        Case "FINANCE": Call LoadFinanceReport
        Case "DELIVERY": Call LoadDeliveryReport
        Case "SUMMARY": Call LoadSummaryReport
        Case "SMART":   Call LoadSmartReport
    End Select
    On Error GoTo 0
End Sub

'========================================
' SEARCH BUTTON
'========================================
Private Sub btnSearch_Click()
    If Len(Trim(txtSearch.text)) < 2 Then
        MsgBox "Please enter at least 2 characters!", vbExclamation
        Exit Sub
    End If
    
    If Not m_IsDataLoaded Then
        MsgBox "Please click SHOW first!", vbExclamation
        Exit Sub
    End If
    
    Call SearchInListView(Trim(txtSearch.text))
End Sub

Private Sub SearchInListView(searchText As String)
    Dim i As Long, j As Long
    Dim li As Object
    Dim matchFound As Boolean
    Dim keepItems As Collection
    Dim allData As Collection
    Dim rowData As Variant
    Dim cols As Long
    
    On Error Resume Next
    
    searchText = UCase(searchText)
    Set keepItems = New Collection
    
    For i = 1 To lstData.ListItems.count
        Set li = lstData.ListItems(i)
        matchFound = False
        
        If InStr(UCase(li.text), searchText) > 0 Then matchFound = True
        
        If Not matchFound Then
            For j = 1 To li.ListSubItems.count
                If InStr(UCase(li.ListSubItems(j).text), searchText) > 0 Then
                    matchFound = True
                    Exit For
                End If
            Next j
        End If
        
        If matchFound Then keepItems.Add i
    Next i
    
    If keepItems.count = 0 Then
        lblStatus.caption = "No records found for '" & Trim(txtSearch.text) & "'"
        Exit Sub
    End If
    
    cols = lstData.ColumnHeaders.count
    Set allData = New Collection
    
    For i = 1 To lstData.ListItems.count
        Set li = lstData.ListItems(i)
        ReDim rowData(1 To cols)
        rowData(1) = li.text
        For j = 1 To li.ListSubItems.count
            If j + 1 <= cols Then rowData(j + 1) = li.ListSubItems(j).text
        Next j
        allData.Add rowData
    Next i
    
    lstData.ListItems.Clear
    
    For i = 1 To allData.count
        Dim shouldKeep As Boolean
        shouldKeep = False
        For j = 1 To keepItems.count
            If keepItems(j) = i Then
                shouldKeep = True
                Exit For
            End If
        Next j
        
        If shouldKeep Then
            rowData = allData(i)
            Set li = lstData.ListItems.Add(, , rowData(1))
            For j = 2 To cols
                If j - 1 <= UBound(rowData) Then li.SubItems(j - 1) = rowData(j)
            Next j
        End If
    Next i
    
    lblStatus.caption = "Search: " & keepItems.count & " records found"
    
    On Error GoTo 0
End Sub

'========================================
' DATE FILTER
'========================================
Private Function IsDateInRange(chkDate As Date) As Boolean
    Dim dFrom As Date, dTo As Date
    
    On Error Resume Next
    dFrom = CDate(txtDateFrom.text)
    dTo = CDate(txtDateTo.text)
    
    If Err.Number <> 0 Then
        IsDateInRange = True
        Exit Function
    End If
    
    If dFrom > dTo Then
        IsDateInRange = True
        Exit Function
    End If
    
    IsDateInRange = (chkDate >= dFrom And chkDate <= dTo)
End Function

'========================================
' ASSIGN TYPE DICTIONARY
'========================================
Private Function GetAssignTypeDict() As Object
    Dim dict As Object
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    
    Set dict = CreateObject("Scripting.Dictionary")
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Assign_Master")
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set GetAssignTypeDict = dict
        Exit Function
    End If
    
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    
    For i = 2 To lastRow
        Dim eID As String, aType As String
        eID = UCase(Trim(ws.Cells(i, 2).value & ""))
        aType = UCase(Trim(ws.Cells(i, 11).value & ""))
        If eID <> "" And Not dict.exists(eID) Then
            dict.Add eID, aType
        End If
    Next i
    
    Set GetAssignTypeDict = dict
End Function

'========================================
' REPORT LOADERS
'========================================
Private Sub LoadMasterReport()
    Dim wsCust As Worksheet, wsJob As Worksheet
    Dim lastRow As Long, i As Long, sr As Long
    Dim custID As String, custName As String, custMobile As String
    Dim totalJobs As Long
    Dim li As Object
    
    On Error Resume Next
    Set wsCust = ThisWorkbook.Sheets("Customer_Master")
    Set wsJob = ThisWorkbook.Sheets("Job_Product")
    On Error GoTo 0
    
    If wsCust Is Nothing Then
        lblStatus.caption = "Customer_Master not found!"
        Exit Sub
    End If
    
    lastRow = wsCust.Cells(wsCust.Rows.count, 1).End(xlUp).row
    sr = 0
    
    For i = 2 To lastRow
        custID = Trim(wsCust.Cells(i, 1).value & "")
        custName = Trim(wsCust.Cells(i, 2).value & "")
        custMobile = Trim(wsCust.Cells(i, 3).value & "")
        
        If custID = "" Then GoTo NextMaster
        
        If cmbCustomer.ListIndex > 0 Then
            If InStr(cmbCustomer.text, custName) = 0 Then GoTo NextMaster
        End If
        
        If Trim(txtSearch.text) <> "" Then
            Dim searchTxt As String
            searchTxt = UCase(Trim(txtSearch.text))
            If InStr(UCase(custID), searchTxt) = 0 And _
               InStr(UCase(custName), searchTxt) = 0 And _
               InStr(UCase(custMobile), searchTxt) = 0 Then
                GoTo NextMaster
            End If
        End If
        
        totalJobs = 0
        If Not wsJob Is Nothing Then
            Dim j As Long, jobLast As Long
            jobLast = wsJob.Cells(wsJob.Rows.count, 1).End(xlUp).row
            For j = 2 To jobLast
                If UCase(Trim(wsJob.Cells(j, 2).value & "")) = UCase(custID) Then
                    totalJobs = totalJobs + 1
                End If
            Next j
        End If
        
        sr = sr + 1
        Set li = lstData.ListItems.Add(, , sr)
        li.SubItems(1) = custID
        li.SubItems(2) = custName
        li.SubItems(3) = custMobile
        li.SubItems(4) = Trim(wsCust.Cells(i, 4).value & "")
        li.SubItems(5) = Trim(wsCust.Cells(i, 5).value & "")
        li.SubItems(6) = Trim(wsCust.Cells(i, 6).value & "")
        li.SubItems(7) = totalJobs
        
NextMaster:
    Next i
    
    Call UpdateSummary(sr, 0, 0)
    m_IsDataLoaded = True
    lblStatus.caption = "Master Report: " & sr & " customers loaded"
End Sub

Private Sub LoadJobReport()
    Dim wsJob As Worksheet, wsCust As Worksheet
    Dim lastRow As Long, i As Long, sr As Long
    Dim entryTypeFilter As String, statusFilter As String, assignTypeFilter As String
    Dim li As Object
    Dim assignDict As Object
    
    On Error Resume Next
    Set wsJob = ThisWorkbook.Sheets("Job_Product")
    Set wsCust = ThisWorkbook.Sheets("Customer_Master")
    On Error GoTo 0
    
    If wsJob Is Nothing Then
        lblStatus.caption = "Job_Product not found!"
        Exit Sub
    End If
    
    Set assignDict = GetAssignTypeDict()
    
    entryTypeFilter = UCase(cmbEntryType.text)
    statusFilter = UCase(cmbStatus.text)
    assignTypeFilter = UCase(cmbAssignType.text)
    
    lastRow = wsJob.Cells(wsJob.Rows.count, 1).End(xlUp).row
    sr = 0
    
    For i = 2 To lastRow
        Dim eType As String, eStatus As String
        eType = UCase(Trim(wsJob.Cells(i, 3).value & ""))
        eStatus = UCase(Trim(wsJob.Cells(i, 4).value & ""))
        
        If entryTypeFilter <> "ALL" And eType <> entryTypeFilter Then GoTo NextJob
        If statusFilter <> "ALL" And eStatus <> statusFilter Then GoTo NextJob
        
        Dim jobDate As String
        jobDate = Trim(wsJob.Cells(i, 5).value & "")
        If jobDate <> "" Then
            On Error Resume Next
            If Not IsDateInRange(CDate(jobDate)) Then GoTo NextJob
            On Error GoTo 0
        End If
        
        Dim entryID As String, assignType As String
        entryID = UCase(Trim(wsJob.Cells(i, 1).value & ""))
        assignType = ""
        If assignDict.exists(entryID) Then assignType = assignDict(entryID)
        
        If assignTypeFilter <> "ALL" And assignType <> assignTypeFilter Then GoTo NextJob
        
        Dim custID As String, custName As String, custMobile As String
        custID = Trim(wsJob.Cells(i, 2).value & "")
        custName = "": custMobile = ""
        If Not wsCust Is Nothing Then
            Dim cRow As Long
            For cRow = 2 To wsCust.Cells(wsCust.Rows.count, 1).End(xlUp).row
                If UCase(Trim(wsCust.Cells(cRow, 1).value & "")) = UCase(custID) Then
                    custName = wsCust.Cells(cRow, 2).value
                    custMobile = wsCust.Cells(cRow, 3).value
                    Exit For
                End If
            Next cRow
        End If
        
        If cmbCustomer.ListIndex > 0 Then
            If InStr(cmbCustomer.text, custName) = 0 Then GoTo NextJob
        End If
        
        sr = sr + 1
        Set li = lstData.ListItems.Add(, , sr)
        li.SubItems(1) = Trim(wsJob.Cells(i, 1).value & "")
        li.SubItems(2) = jobDate
        li.SubItems(3) = custName
        li.SubItems(4) = custMobile
        li.SubItems(5) = eType
        li.SubItems(6) = Trim(wsJob.Cells(i, 6).value & "")
        li.SubItems(7) = Trim(wsJob.Cells(i, 7).value & "")
        li.SubItems(8) = Trim(wsJob.Cells(i, 8).value & "")
        li.SubItems(9) = Trim(wsJob.Cells(i, 9).value & "")
        li.SubItems(10) = assignType
        li.SubItems(11) = eStatus
        
        Select Case eStatus
            Case "ACTIVE": li.ForeColor = RGB(0, 128, 0)
            Case "ASSIGNED": li.ForeColor = RGB(0, 0, 200)
            Case "PENDING": li.ForeColor = RGB(200, 0, 0)
            Case "COMPLETED": li.ForeColor = RGB(0, 128, 0)
            Case "DELIVERED": li.ForeColor = RGB(100, 100, 100)
        End Select
        
NextJob:
    Next i
    
    Call UpdateSummary(sr, 0, 0)
    m_IsDataLoaded = True
    lblStatus.caption = "Job Report: " & sr & " entries loaded"
End Sub

Private Sub LoadAssignReport()
    Dim wsAssign As Worksheet
    Dim lastRow As Long, i As Long, sr As Long
    Dim assignTypeFilter As String, statusFilter As String
    Dim li As Object
    
    On Error Resume Next
    Set wsAssign = ThisWorkbook.Sheets("Assign_Master")
    On Error GoTo 0
    
    If wsAssign Is Nothing Then
        lblStatus.caption = "Assign_Master not found!"
        Exit Sub
    End If
    
    assignTypeFilter = UCase(cmbAssignType.text)
    statusFilter = UCase(cmbStatus.text)
    
    lastRow = wsAssign.Cells(wsAssign.Rows.count, 1).End(xlUp).row
    sr = 0
    
    For i = 2 To lastRow
        Dim aType As String, aStatus As String, aName As String
        aType = UCase(Trim(wsAssign.Cells(i, 11).value & ""))
        aStatus = UCase(Trim(wsAssign.Cells(i, 32).value & ""))
        aName = Trim(wsAssign.Cells(i, 14).value & "")
        
        If assignTypeFilter <> "ALL" And aType <> assignTypeFilter Then GoTo NextAssign
        If statusFilter <> "ALL" And aStatus <> statusFilter Then GoTo NextAssign
        
        Dim assignDate As String
        assignDate = Trim(wsAssign.Cells(i, 22).value & "")
        If assignDate <> "" Then
            On Error Resume Next
            If Not IsDateInRange(CDate(assignDate)) Then GoTo NextAssign
            On Error GoTo 0
        End If
        
        Dim aCustomer As String
        aCustomer = Trim(wsAssign.Cells(i, 5).value & "")
        If cmbCustomer.ListIndex > 0 Then
            If InStr(cmbCustomer.text, aCustomer) = 0 Then GoTo NextAssign
        End If
        
        sr = sr + 1
        Set li = lstData.ListItems.Add(, , sr)
        li.SubItems(1) = Trim(wsAssign.Cells(i, 2).value & "")
        li.SubItems(2) = assignDate
        li.SubItems(3) = aCustomer
        li.SubItems(4) = Trim(wsAssign.Cells(i, 6).value & "")
        li.SubItems(5) = aType
        li.SubItems(6) = aName
        li.SubItems(7) = aStatus
        li.SubItems(8) = Trim(wsAssign.Cells(i, 27).value & "")
        
        Select Case aStatus
            Case "PENDING": li.ForeColor = RGB(200, 0, 0)
            Case "COMPLETED": li.ForeColor = RGB(0, 128, 0)
            Case "IN PROGRESS": li.ForeColor = RGB(255, 165, 0)
        End Select
        
NextAssign:
    Next i
    
    Call UpdateSummary(sr, 0, 0)
    m_IsDataLoaded = True
    lblStatus.caption = "Assignment Report: " & sr & " records loaded"
End Sub

Private Sub LoadFinanceReport()
    Dim wsPay As Worksheet
    Dim lastRow As Long, i As Long, sr As Long
    Dim payStatusFilter As String
    Dim li As Object
    Dim totalAmt As Double
    
    On Error Resume Next
    Set wsPay = ThisWorkbook.Sheets("Payment_Master")
    On Error GoTo 0
    
    If wsPay Is Nothing Then
        lblStatus.caption = "Payment_Master not found!"
        Exit Sub
    End If
    
    payStatusFilter = UCase(cmbPayment.text)
    lastRow = wsPay.Cells(wsPay.Rows.count, 1).End(xlUp).row
    sr = 0
    totalAmt = 0
    
    For i = 2 To lastRow
        If Trim(wsPay.Cells(i, 6).value & "") <> "" Then GoTo NextFinance
        
        Dim pDate As String, pAmt As Double
        pDate = Trim(wsPay.Cells(i, 7).value & "")
        pAmt = val(wsPay.Cells(i, 16).value)
        
        If pDate <> "" Then
            On Error Resume Next
            If Not IsDateInRange(CDate(pDate)) Then GoTo NextFinance
            On Error GoTo 0
        End If
        
        Dim pStatus As String
        pStatus = UCase(Trim(wsPay.Cells(i, 17).value & ""))
        If payStatusFilter <> "ALL" And pStatus <> payStatusFilter Then GoTo NextFinance
        
        Dim pCustomer As String
        pCustomer = Trim(wsPay.Cells(i, 10).value & "")
        If cmbCustomer.ListIndex > 0 Then
            If InStr(cmbCustomer.text, pCustomer) = 0 Then GoTo NextFinance
        End If
        
        sr = sr + 1
        totalAmt = totalAmt + pAmt
        
        Set li = lstData.ListItems.Add(, , sr)
        li.SubItems(1) = Trim(wsPay.Cells(i, 1).value & "")
        li.SubItems(2) = pDate
        li.SubItems(3) = Trim(wsPay.Cells(i, 2).value & "")
        li.SubItems(4) = pCustomer
        li.SubItems(5) = Trim(wsPay.Cells(i, 9).value & "")
        li.SubItems(6) = Trim(wsPay.Cells(i, 15).value & "")
        li.SubItems(7) = Format(pAmt, "0.00")
        li.SubItems(8) = pStatus
        
NextFinance:
    Next i
    
    Call UpdateSummary(sr, totalAmt, 0)
    m_IsDataLoaded = True
    lblStatus.caption = "Finance Report: " & sr & " payments loaded"
End Sub

Private Sub LoadDeliveryReport()
    Dim wsAssign As Worksheet
    Dim lastRow As Long, i As Long, sr As Long
    Dim statusFilter As String
    Dim li As Object
    
    On Error Resume Next
    Set wsAssign = ThisWorkbook.Sheets("Assign_Master")
    On Error GoTo 0
    
    If wsAssign Is Nothing Then
        lblStatus.caption = "Assign_Master not found!"
        Exit Sub
    End If
    
    statusFilter = UCase(cmbStatus.text)
    
    lastRow = wsAssign.Cells(wsAssign.Rows.count, 1).End(xlUp).row
    sr = 0
    
    For i = 2 To lastRow
        Dim dStatus As String, dSendDate As String, dExpDate As String
        dStatus = UCase(Trim(wsAssign.Cells(i, 32).value & ""))
        dSendDate = Trim(wsAssign.Cells(i, 26).value & "")
        dExpDate = Trim(wsAssign.Cells(i, 27).value & "")
        
        If statusFilter <> "ALL" And dStatus <> statusFilter Then GoTo NextDelivery
        
        If dStatus = "COMPLETED" Or dStatus = "DELIVERED" Then GoTo NextDelivery
        
        If dSendDate <> "" Then
            On Error Resume Next
            If Not IsDateInRange(CDate(dSendDate)) Then GoTo NextDelivery
            On Error GoTo 0
        End If
        
        Dim dCustomer As String
        dCustomer = Trim(wsAssign.Cells(i, 5).value & "")
        If cmbCustomer.ListIndex > 0 Then
            If InStr(cmbCustomer.text, dCustomer) = 0 Then GoTo NextDelivery
        End If
        
        Dim daysPending As Long
        daysPending = 0
        If dSendDate <> "" Then
            On Error Resume Next
            daysPending = DateDiff("d", CDate(dSendDate), Date)
            On Error GoTo 0
        End If
        
        sr = sr + 1
        Set li = lstData.ListItems.Add(, , sr)
        li.SubItems(1) = Trim(wsAssign.Cells(i, 2).value & "")
        li.SubItems(2) = dCustomer
        li.SubItems(3) = Trim(wsAssign.Cells(i, 6).value & "")
        li.SubItems(4) = Trim(wsAssign.Cells(i, 11).value & "")
        li.SubItems(5) = dSendDate
        li.SubItems(6) = dExpDate
        li.SubItems(7) = daysPending
        li.SubItems(8) = dStatus
        
        If daysPending > 7 Then li.ForeColor = RGB(200, 0, 0)
        
NextDelivery:
    Next i
    
    Call UpdateSummary(sr, 0, 0)
    m_IsDataLoaded = True
    lblStatus.caption = "Delivery Report: " & sr & " pending deliveries loaded"
End Sub

Private Sub LoadSummaryReport()
    Dim wsJob As Worksheet
    Dim lastRow As Long, i As Long
    Dim dFrom As Date, dTo As Date
    Dim dict As Object
    Dim sDate As String, dt As Date
    Dim eType As String, sStatus As String
    Dim li As Object
    Dim key As String
    
    On Error Resume Next
    Set wsJob = ThisWorkbook.Sheets("Job_Product")
    On Error GoTo 0
    
    If wsJob Is Nothing Then
        lblStatus.caption = "Job_Product not found!"
        Exit Sub
    End If
    
    On Error Resume Next
    dFrom = CDate(txtDateFrom.text)
    dTo = CDate(txtDateTo.text)
    On Error GoTo 0
    
    If dFrom > dTo Then
        lblStatus.caption = "From Date > To Date!"
        Exit Sub
    End If
    
    Set dict = CreateObject("Scripting.Dictionary")
    
    lastRow = wsJob.Cells(wsJob.Rows.count, 1).End(xlUp).row
    
    For i = 2 To lastRow
        sDate = Trim(wsJob.Cells(i, 5).value & "")
        If sDate = "" Then GoTo NextSumRow
        
        On Error Resume Next
        dt = CDate(sDate)
        On Error GoTo 0
        If Err.Number <> 0 Then GoTo NextSumRow
        
        If dt < dFrom Or dt > dTo Then GoTo NextSumRow
        
        eType = UCase(Trim(wsJob.Cells(i, 3).value & ""))
        sStatus = UCase(Trim(wsJob.Cells(i, 4).value & ""))
        
        key = Format(dt, "dd-mm-yyyy")
        
        If Not dict.exists(key) Then
            dict.Add key, Array(0, 0, 0, 0, 0)
        End If
        
        Dim arr As Variant
        arr = dict(key)
        arr(0) = arr(0) + 1
        
        If eType = "WARRANTY" Then arr(1) = arr(1) + 1
        If eType = "SERVICE" Then arr(2) = arr(2) + 1
        
        If sStatus = "COMPLETED" Or sStatus = "DELIVERED" Then
            arr(3) = arr(3) + 1
        Else
            arr(4) = arr(4) + 1
        End If
        
        dict(key) = arr
        
NextSumRow:
    Next i
    
    If dict.count = 0 Then
        lblStatus.caption = "No data for Summary"
        m_IsDataLoaded = True
        Exit Sub
    End If
    
    Dim dates() As Date
    ReDim dates(1 To dict.count)
    i = 1
    Dim k As Variant
    For Each k In dict.keys
        dates(i) = CDate(k)
        i = i + 1
    Next k
    
    Dim j As Long
    For i = 1 To UBound(dates) - 1
        For j = i + 1 To UBound(dates)
            If dates(i) > dates(j) Then
                Dim tmp As Date
                tmp = dates(i)
                dates(i) = dates(j)
                dates(j) = tmp
            End If
        Next j
    Next i
    
    Dim sr As Long
    sr = 0
    
    For i = 1 To UBound(dates)
        key = Format(dates(i), "dd-mm-yyyy")
        arr = dict(key)
        sr = sr + 1
        Set li = lstData.ListItems.Add(, , sr)
        li.SubItems(1) = key
        li.SubItems(2) = arr(0)
        li.SubItems(3) = arr(1)
        li.SubItems(4) = arr(2)
        li.SubItems(5) = arr(3)
        li.SubItems(6) = arr(4)
        li.SubItems(7) = "-"
    Next i
    
    Call UpdateSummary(sr, 0, 0)
    m_IsDataLoaded = True
    lblStatus.caption = "Summary Report: " & sr & " periods loaded"
End Sub

Private Sub LoadSmartReport()
    Dim wsAssign As Worksheet, wsJob As Worksheet
    Dim lastRow As Long, i As Long, sr As Long
    Dim li As Object
    
    On Error Resume Next
    Set wsAssign = ThisWorkbook.Sheets("Assign_Master")
    Set wsJob = ThisWorkbook.Sheets("Job_Product")
    On Error GoTo 0
    
    sr = 0
    
    If Not wsAssign Is Nothing Then
        lastRow = wsAssign.Cells(wsAssign.Rows.count, 1).End(xlUp).row
        
        For i = 2 To lastRow
            Dim expDate As String, aStatus As String
            expDate = Trim(wsAssign.Cells(i, 27).value & "")
            aStatus = UCase(Trim(wsAssign.Cells(i, 32).value & ""))
            
            If aStatus <> "PENDING" And aStatus <> "IN PROGRESS" Then GoTo NextSmart1
            
            If expDate <> "" Then
                On Error Resume Next
                If CDate(expDate) < Date Then
                    Dim daysOverdue As Long
                    daysOverdue = DateDiff("d", CDate(expDate), Date)
                    
                    sr = sr + 1
                    Set li = lstData.ListItems.Add(, , sr)
                    li.SubItems(1) = "OVERDUE RETURN"
                    li.SubItems(2) = Trim(wsAssign.Cells(i, 2).value & "")
                    li.SubItems(3) = Trim(wsAssign.Cells(i, 5).value & "")
                    li.SubItems(4) = Trim(wsAssign.Cells(i, 6).value & "")
                    li.SubItems(5) = "Expected return " & expDate & " crossed by " & daysOverdue & " days"
                    li.SubItems(6) = daysOverdue
                    li.SubItems(7) = "URGENT"
                    li.ForeColor = RGB(200, 0, 0)
                End If
                On Error GoTo 0
            End If
NextSmart1:
        Next i
    End If
    
    If Not wsJob Is Nothing Then
        lastRow = wsJob.Cells(wsJob.Rows.count, 1).End(xlUp).row
        
        For i = 2 To lastRow
            Dim jStatus As String, jDate As String
            jStatus = UCase(Trim(wsJob.Cells(i, 4).value & ""))
            jDate = Trim(wsJob.Cells(i, 5).value & "")
            
            If jStatus = "PENDING" Or jStatus = "ACTIVE" Then
                Dim daysPending As Long
                daysPending = 0
                If jDate <> "" Then
                    On Error Resume Next
                    daysPending = DateDiff("d", CDate(jDate), Date)
                    On Error GoTo 0
                End If
                
                If daysPending > 3 Then
                    sr = sr + 1
                    Set li = lstData.ListItems.Add(, , sr)
                    li.SubItems(1) = "PENDING JOB"
                    li.SubItems(2) = Trim(wsJob.Cells(i, 1).value & "")
                    
                    Dim jCustID As String, jCustName As String
                    jCustID = Trim(wsJob.Cells(i, 2).value & "")
                    jCustName = jCustID
                    On Error Resume Next
                    Dim wsCust As Worksheet
                    Set wsCust = ThisWorkbook.Sheets("Customer_Master")
                    If Not wsCust Is Nothing Then
                        Dim jcRow As Long
                        For jcRow = 2 To wsCust.Cells(wsCust.Rows.count, 1).End(xlUp).row
                            If UCase(Trim(wsCust.Cells(jcRow, 1).value & "")) = UCase(jCustID) Then
                                jCustName = wsCust.Cells(jcRow, 2).value
                                Exit For
                            End If
                        Next jcRow
                    End If
                    On Error GoTo 0
                    
                    li.SubItems(3) = jCustName
                    li.SubItems(4) = Trim(wsJob.Cells(i, 6).value & "")
                    li.SubItems(5) = "Job pending for " & daysPending & " days without assignment"
                    li.SubItems(6) = daysPending
                    li.SubItems(7) = "FOLLOWUP"
                    li.ForeColor = RGB(255, 140, 0)
                End If
            End If
        Next i
    End If
    
    Call UpdateSummary(sr, 0, 0)
    m_IsDataLoaded = True
    lblStatus.caption = "Smart Alerts: " & sr & " alerts"
End Sub

'========================================
' UPDATE SUMMARY BAR
'========================================
Private Sub UpdateSummary(recCount As Long, totalAmt As Double, pendingAmt As Double)
    On Error Resume Next
    lblTotalRecords.caption = recCount
    If totalAmt > 0 Then
        lblTotalAmount.caption = "Rs." & Format(totalAmt, "#,##0.00")
    Else
        lblTotalAmount.caption = "-"
    End If
    If pendingAmt > 0 Then
        lblTotalPending.caption = "Rs." & Format(pendingAmt, "#,##0.00")
    Else
        lblTotalPending.caption = "-"
    End If
    On Error GoTo 0
End Sub

'========================================
' SETTINGS READERS  [READS FROM YOUR SETTING SHEETS]
'========================================

'--- Company Name from Software_Config B2 ---
Private Function GetCompanyName() As String
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Software_Config")
    GetCompanyName = Trim(ws.Range("B2").value & "")
    On Error GoTo 0
    If GetCompanyName = "" Then GetCompanyName = "GLOBAL SOFT"
End Function

'--- Company Email from Software_Config B10 ---
Private Function GetCompanyEmail() As String
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Software_Config")
    GetCompanyEmail = Trim(ws.Range("B10").value & "")
    On Error GoTo 0
End Function

'--- Company Mobile from Software_Config B9 ---
Private Function GetCompanyMobile() As String
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Software_Config")
    GetCompanyMobile = Trim(ws.Range("B9").value & "")
    On Error GoTo 0
End Function

'--- Paper Size from Settings B20 (A4/A5) ---
Private Function GetPaperSize() As XlPaperSize
    Dim ws As Worksheet
    Dim paper As String
    paper = "A4"
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Settings")
    paper = UCase(Trim(ws.Range("B20").value & ""))
    On Error GoTo 0
    
    If paper = "" Then paper = "A4"
    
    Select Case paper
        Case "A4": GetPaperSize = xlPaperA4
        Case "A5": GetPaperSize = xlPaperA5
        Case "LETTER": GetPaperSize = xlPaperLetter
        Case "LEGAL": GetPaperSize = xlPaperLegal
        Case Else: GetPaperSize = xlPaperA4
    End Select
End Function

'========================================
' EXPORT HELPERS
'========================================

Private Function SafeFileName(text As String) As String
    Dim invalid As String, i As Long
    invalid = "\/:*?""<>|"
    For i = 1 To Len(invalid)
        text = Replace(text, Mid(invalid, i, 1), "_")
    Next i
    SafeFileName = text
End Function

Private Function GetColumnIndex(colName As String) As Long
    Dim i As Long
    For i = 1 To lstData.ColumnHeaders.count
        If UCase(lstData.ColumnHeaders(i).text) = UCase(colName) Then
            GetColumnIndex = i
            Exit Function
        End If
    Next i
    GetColumnIndex = 0
End Function

Private Function cleanMobile(mobile As String) As String
    Dim s As String
    s = Trim(mobile)
    s = Replace(s, " ", "")
    s = Replace(s, "-", "")
    s = Replace(s, "(", "")
    s = Replace(s, ")", "")
    s = Replace(s, "+", "")
    s = Replace(s, ".", "")
    
    If Left(s, 1) = "0" Then s = "91" & Mid(s, 2)
    If Len(s) = 10 And Left(s, 1) <> "9" Then s = "91" & s
    If Len(s) = 10 Then s = "91" & s
    
    cleanMobile = s
End Function

Private Function URLEncode(text As String) As String
    Dim i As Long, char As String, result As String
    result = ""
    For i = 1 To Len(text)
        char = Mid(text, i, 1)
        Select Case char
            Case " ": result = result & "%20"
            Case "!": result = result & "%21"
            Case "#": result = result & "%23"
            Case "$": result = result & "%24"
            Case "&": result = result & "%26"
            Case "'": result = result & "%27"
            Case "(": result = result & "%28"
            Case ")": result = result & "%29"
            Case "*": result = result & "%2A"
            Case "+": result = result & "%2B"
            Case ",": result = result & "%2C"
            Case "/": result = result & "%2F"
            Case ":": result = result & "%3A"
            Case ";": result = result & "%3B"
            Case "=": result = result & "%3D"
            Case "?": result = result & "%3F"
            Case "@": result = result & "%40"
            Case "[": result = result & "%5B"
            Case "]": result = result & "%5D"
            Case Chr(13): result = result & "%0D"
            Case Chr(10): result = result & "%0A"
            Case Else: result = result & char
        End Select
    Next i
    URLEncode = result
End Function

Private Function CreateReportSheet() As Worksheet
    Dim wsTemp As Worksheet
    Dim i As Long, j As Long
    Dim li As Object
    Dim colCount As Long
    Dim lastDataRow As Long
    
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets("TempReport").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    Set wsTemp = ThisWorkbook.Sheets.Add
    wsTemp.name = "TempReport"
    wsTemp.Cells.Font.name = "Calibri"
    wsTemp.Cells.Font.Size = 10
    
    With wsTemp.Range("A1")
        .value = GetCompanyName()
        .Font.Size = 16
        .Font.Bold = True
        .Font.Color = RGB(26, 188, 156)
    End With
    
    With wsTemp.Range("A2")
        .value = m_CurrentCategory & " REPORT"
        .Font.Size = 13
        .Font.Bold = True
    End With
    
    wsTemp.Range("A3").value = "Date: " & Format(Date, "dd-mm-yyyy") & "  |  Time: " & Format(Time, "hh:mm AM/PM")
    wsTemp.Range("A4").value = "Period: " & txtDateFrom.text & " to " & txtDateTo.text & "  |  Records: " & lstData.ListItems.count
    
    colCount = lstData.ColumnHeaders.count
    For i = 1 To colCount
        With wsTemp.Cells(6, i)
            .value = lstData.ColumnHeaders(i).text
            .Font.Bold = True
            .Interior.Color = RGB(26, 188, 156)
            .Font.Color = RGB(255, 255, 255)
            .HorizontalAlignment = xlCenter
        End With
    Next i
    
    lastDataRow = 6
    For i = 1 To lstData.ListItems.count
        Set li = lstData.ListItems(i)
        lastDataRow = lastDataRow + 1
        wsTemp.Cells(lastDataRow, 1).value = li.text
        For j = 1 To li.ListSubItems.count
            If j < colCount Then
                wsTemp.Cells(lastDataRow, j + 1).value = li.ListSubItems(j).text
            End If
        Next j
    Next i
    
    With wsTemp.Range(wsTemp.Cells(6, 1), wsTemp.Cells(lastDataRow, colCount))
        .Columns.AutoFit
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders.Color = RGB(200, 200, 200)
    End With
    
    With wsTemp.PageSetup
        .PaperSize = GetPaperSize()
        .Orientation = xlPortrait
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .CenterHorizontally = True
        
        .LeftMargin = Application.InchesToPoints(0.3)
        .RightMargin = Application.InchesToPoints(0.3)
        .TopMargin = Application.InchesToPoints(0.5)
        .BottomMargin = Application.InchesToPoints(0.5)
        .HeaderMargin = Application.InchesToPoints(0.3)
        .FooterMargin = Application.InchesToPoints(0.3)
        
        .LeftHeader = "&""Calibri,Bold""&14" & GetCompanyName()
        .CenterHeader = "&""Calibri,Bold""&12" & m_CurrentCategory & " REPORT"
        .RightHeader = "&10Date: " & Format(Date, "dd-mm-yyyy")
        .LeftFooter = "&8" & GetCompanyName() & " - Reports & Analytics"
        .CenterFooter = "&8Page &P of &N"
        .RightFooter = "&8Confidential"
        
        .PrintArea = wsTemp.Range("A1", wsTemp.Cells(lastDataRow, colCount)).Address
    End With
    
    Set CreateReportSheet = wsTemp
End Function

'========================================
' PRINT BUTTON
'========================================
Private Sub btnPrint_Click()
    If Not m_IsDataLoaded Then
        MsgBox "Load data first!", vbExclamation
        Exit Sub
    End If
    
    Dim wsTemp As Worksheet
    On Error GoTo Print_Error
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Set wsTemp = CreateReportSheet()
    wsTemp.PrintOut Copies:=1, Collate:=True
    
    lblStatus.caption = "Report sent to printer [" & UCase(GetSettingValue("PAPER_SIZE", "A4")) & "]"
    
Print_Error:
    On Error Resume Next
    If Not wsTemp Is Nothing Then wsTemp.Delete
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    On Error GoTo 0
End Sub

'========================================
' PDF BUTTON
'========================================
Private Sub btnPDF_Click()
    If Not m_IsDataLoaded Then
        MsgBox "Load data first!", vbExclamation
        Exit Sub
    End If
    
    Dim fileName As Variant
    Dim suggestedName As String
    
    suggestedName = ThisWorkbook.path & "\" & SafeFileName(m_CurrentCategory) & "_Report_" & Format(Date, "dd-mm-yyyy")
    
    fileName = Application.GetSaveAsFilename( _
        InitialFileName:=suggestedName, _
        FileFilter:="PDF Files (*.pdf), *.pdf")
    
    If fileName = False Then Exit Sub
    
    Dim wsTemp As Worksheet
    On Error GoTo PDF_Error
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Set wsTemp = CreateReportSheet()
    
    wsTemp.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        fileName:=CStr(fileName), _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=True
    
    lblStatus.caption = "PDF saved [" & UCase(GetSettingValue("PAPER_SIZE", "A4")) & "]: " & fileName
    
PDF_Error:
    On Error Resume Next
    If Not wsTemp Is Nothing Then wsTemp.Delete
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    On Error GoTo 0
End Sub

'========================================
' EXCEL BUTTON
'========================================
Private Sub btnExcel_Click()
    If Not m_IsDataLoaded Then
        MsgBox "Load data first!", vbExclamation
        Exit Sub
    End If
    
    Dim fileName As Variant
    Dim suggestedName As String
    
    suggestedName = ThisWorkbook.path & "\" & SafeFileName(m_CurrentCategory) & "_Report_" & Format(Date, "dd-mm-yyyy")
    
    fileName = Application.GetSaveAsFilename( _
        InitialFileName:=suggestedName, _
        FileFilter:="Excel Files (*.xlsx), *.xlsx")
    
    If fileName = False Then Exit Sub
    
    Dim wsTemp As Worksheet
    Dim wb As Workbook
    On Error GoTo Excel_Error
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Set wsTemp = CreateReportSheet()
    wsTemp.Copy
    Set wb = ActiveWorkbook
    
    wb.Sheets(1).name = SafeFileName(m_CurrentCategory) & "_Report"
    
    Dim ws As Worksheet
    For Each ws In wb.Sheets
        If ws.name <> SafeFileName(m_CurrentCategory) & "_Report" Then
            ws.Delete
        End If
    Next ws
    
    wb.SaveAs fileName:=CStr(fileName), FileFormat:=xlOpenXMLWorkbook
    wb.Close SaveChanges:=False
    
    lblStatus.caption = "Excel saved: " & fileName
    
Excel_Error:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    If Not wsTemp Is Nothing Then wsTemp.Delete
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    On Error GoTo 0
End Sub

'========================================
' WHATSAPP BUTTON  [DESKTOP APP -> WEB FALLBACK + INPUTBOX]
'========================================
Private Sub btnWhatsApp_Click()
    If Not m_IsDataLoaded Then
        MsgBox "Load data first!", vbExclamation
        Exit Sub
    End If
    
    ' Build message
    Dim msg As String
    msg = "*" & GetCompanyName() & "*" & vbCrLf
    msg = msg & m_CurrentCategory & " Report" & vbCrLf
    msg = msg & "Date: " & Format(Date, "dd-mm-yyyy") & vbCrLf
    msg = msg & "Total Records: " & lstData.ListItems.count & vbCrLf
    msg = msg & "Period: " & txtDateFrom.text & " to " & txtDateTo.text & vbCrLf & vbCrLf
    msg = msg & "Please check the report details."
    
    ' Find first mobile from report as default
    Dim defaultMobile As String
    defaultMobile = ""
    
    Dim mobileCol As Long
    mobileCol = GetColumnIndex("Mobile")
    If mobileCol = 0 Then mobileCol = GetColumnIndex("MOBILE")
    
    Dim i As Long
    If mobileCol > 0 Then
        For i = 1 To lstData.ListItems.count
            Dim li As Object
            Set li = lstData.ListItems(i)
            Dim mobile As String
            If mobileCol = 1 Then
                mobile = li.text
            Else
                If mobileCol - 1 <= li.ListSubItems.count Then
                    mobile = li.ListSubItems(mobileCol - 1).text
                End If
            End If
            mobile = cleanMobile(mobile)
            If mobile <> "" Then
                defaultMobile = mobile
                Exit For
            End If
        Next i
    End If
    
    ' Ask user for mobile number
    Dim userMobile As String
    userMobile = InputBox("Enter WhatsApp Number (with country code e.g. 919876543210):" & vbCrLf & _
                          "Leave blank to cancel.", "Send WhatsApp", defaultMobile)
    
    If Trim(userMobile) = "" Then
        lblStatus.caption = "WhatsApp cancelled by user."
        Exit Sub
    End If
    
    userMobile = cleanMobile(userMobile)
    If Len(userMobile) < 10 Then
        MsgBox "Invalid mobile number! Please enter at least 10 digits.", vbExclamation
        Exit Sub
    End If
    
    Dim encodedMsg As String
    encodedMsg = URLEncode(msg)
    
    ' TRY 1: WhatsApp Desktop App (whatsapp:// protocol)
    Dim deskUrl As String
    deskUrl = "whatsapp://send?phone=" & userMobile & "&text=" & encodedMsg
    
    Dim ret As Long
    ret = ShellExecute(0, "open", deskUrl, vbNullString, vbNullString, vbNormalFocus)
    
    ' TRY 2: If desktop not available (ret <= 32), open WhatsApp Web
    If ret <= 32 Then
        Dim webUrl As String
        webUrl = "https://wa.me/" & userMobile & "?text=" & encodedMsg
        ret = ShellExecute(0, "open", webUrl, vbNullString, vbNullString, vbNormalFocus)
        lblStatus.caption = "WhatsApp Web opened for: " & userMobile
    Else
        lblStatus.caption = "WhatsApp Desktop opened for: " & userMobile
    End If
    
    ' Create text file for copy-paste backup
    Dim tempFile As String
    tempFile = Environ("TEMP") & "\WhatsApp_Report_" & SafeFileName(m_CurrentCategory) & ".txt"
    
    Dim fNum As Integer
    fNum = FreeFile
    Open tempFile For Output As #fNum
    Print #fNum, msg
    Print #fNum, String(50, "-")
    
    Dim hLine As String
    For i = 1 To lstData.ColumnHeaders.count
        hLine = hLine & lstData.ColumnHeaders(i).text & vbTab
    Next i
    Print #fNum, hLine
    
    For i = 1 To lstData.ListItems.count
        Set li = lstData.ListItems(i)
        Dim dLine As String
        dLine = li.text & vbTab
        Dim j As Long
        For j = 1 To li.ListSubItems.count
            dLine = dLine & li.ListSubItems(j).text & vbTab
        Next j
        Print #fNum, dLine
    Next i
    Close #fNum
    
    shell "notepad.exe " & tempFile, vbNormalFocus
End Sub

'========================================
' EMAIL BUTTON  [MAILTO + INPUTBOX FOR RECIPIENT]
'========================================
Private Sub btnEmail_Click()
    If Not m_IsDataLoaded Then
        MsgBox "Load data first!", vbExclamation
        Exit Sub
    End If
    
    ' Build subject
    Dim subject As String
    subject = GetCompanyName() & " - " & m_CurrentCategory & " Report [" & Format(Date, "dd-mm-yyyy") & "]"
    
    ' Build plain text body
    Dim body As String
    body = "Dear Sir/Madam," & vbCrLf & vbCrLf
    body = body & "Please find the " & m_CurrentCategory & " report details below:" & vbCrLf & vbCrLf
    body = body & "Company: " & GetCompanyName() & vbCrLf
    body = body & "Report Date: " & Format(Date, "dd-mm-yyyy") & vbCrLf
    body = body & "Period: " & txtDateFrom.text & " to " & txtDateTo.text & vbCrLf
    body = body & "Total Records: " & lstData.ListItems.count & vbCrLf & vbCrLf
    
    body = body & "REPORT SUMMARY:" & vbCrLf
    body = body & String(50, "-") & vbCrLf
    
    Dim i As Long, j As Long
    Dim hLine As String
    For i = 1 To lstData.ColumnHeaders.count
        hLine = hLine & lstData.ColumnHeaders(i).text & vbTab
    Next i
    body = body & hLine & vbCrLf
    body = body & String(50, "-") & vbCrLf
    
    For i = 1 To lstData.ListItems.count
        Dim li As Object
        Set li = lstData.ListItems(i)
        Dim dLine As String
        dLine = li.text & vbTab
        For j = 1 To li.ListSubItems.count
            dLine = dLine & li.ListSubItems(j).text & vbTab
        Next j
        body = body & dLine & vbCrLf
    Next i
    
    body = body & vbCrLf & "Regards," & vbCrLf
    body = body & GetCompanyName() & vbCrLf
    body = body & "Email: " & GetCompanyEmail() & vbCrLf
    body = body & "Phone: " & GetCompanyMobile() & vbCrLf
    
    ' Find default email from report (or use company email)
    Dim defaultEmail As String
    defaultEmail = GetCompanyEmail()
    
    Dim emailCol As Long
    emailCol = GetColumnIndex("Email")
    If emailCol = 0 Then emailCol = GetColumnIndex("EMAIL")
    
    If emailCol > 0 Then
        For i = 1 To lstData.ListItems.count
            Set li = lstData.ListItems(i)
            Dim rEmail As String
            If emailCol = 1 Then
                rEmail = li.text
            Else
                If emailCol - 1 <= li.ListSubItems.count Then
                    rEmail = li.ListSubItems(emailCol - 1).text
                End If
            End If
            If InStr(rEmail, "@") > 0 Then
                defaultEmail = rEmail
                Exit For
            End If
        Next i
    End If
    
    ' Ask user for recipient email
    Dim userEmail As String
    userEmail = InputBox("Enter Email Address to send report:" & vbCrLf & _
                         "You can change it or keep as default.", "Send Email", defaultEmail)
    
    If Trim(userEmail) = "" Then
        lblStatus.caption = "Email cancelled by user."
        Exit Sub
    End If
    
    If InStr(userEmail, "@") = 0 Or InStr(userEmail, ".") = 0 Then
        MsgBox "Invalid email address!", vbExclamation
        Exit Sub
    End If
    
    ' Create mailto URL (opens Gmail/Outlook/default client)
    Dim mailtoUrl As String
    mailtoUrl = "mailto:" & userEmail & _
                "?subject=" & URLEncode(subject) & _
                "&body=" & URLEncode(body)
    
    Dim ret As Long
    ret = ShellExecute(0, "open", mailtoUrl, vbNullString, vbNullString, vbNormalFocus)
    
    If ret > 32 Then
        lblStatus.caption = "Email client opened for: " & userEmail
    Else
        MsgBox "Could not open email client. Please set a default email app.", vbExclamation
    End If
    
    ' Backup text file
    Dim tempFile As String
    tempFile = Environ("TEMP") & "\Email_Report_" & SafeFileName(m_CurrentCategory) & ".txt"
    
    Dim fNum As Integer
    fNum = FreeFile
    Open tempFile For Output As #fNum
    Print #fNum, "To: " & userEmail
    Print #fNum, "Subject: " & subject
    Print #fNum, String(50, "=")
    Print #fNum, body
    Close #fNum
    
    shell "notepad.exe " & tempFile, vbNormalFocus
End Sub

'========================================
' CLEAR & CLOSE
'========================================
Private Sub btnClear_Click()
    On Error Resume Next
    lstData.ListItems.Clear
    txtDateFrom.text = "01-04-2026"
    txtDateTo.text = "31-12-2026"
    cmbCustomer.ListIndex = 0
    cmbEntryType.ListIndex = 0
    cmbStatus.ListIndex = 0
    cmbAssignType.ListIndex = 0
    cmbPayment.ListIndex = 0
    txtSearch.text = ""
    Call UpdateSummary(0, 0, 0)
    lblStatus.caption = "Cleared - Click SHOW"
    m_IsDataLoaded = False
    On Error GoTo 0
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

'========================================
' DATE VALIDATION
'========================================
Private Sub txtDateFrom_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    On Error Resume Next
    If txtDateFrom.text <> "" Then txtDateFrom.text = Format(CDate(txtDateFrom.text), "dd-mm-yyyy")
    On Error GoTo 0
End Sub

Private Sub txtDateTo_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    On Error Resume Next
    If txtDateTo.text <> "" Then txtDateTo.text = Format(CDate(txtDateTo.text), "dd-mm-yyyy")
    On Error GoTo 0
End Sub

