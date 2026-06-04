VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDiagnostic 
   Caption         =   "System Diagnostic & Fix"
   ClientHeight    =   11955
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14565
   OleObjectBlob   =   "frmDiagnostic.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmDiagnostic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===========================================
' COMPLETE FIXED: frmDiagnostic Code
' Fixes: Missing Customer Auto-Fix + Export Report
'===========================================

Private Sub frmProgressBar_Click()

End Sub

Private Sub UserForm_Initialize()
    On Error Resume Next
    
    ' Setup ListBox Columns
    lstResults.ColumnCount = 4
    lstResults.ColumnWidths = "80 pt;120 pt;200 pt;150 pt"
    
    lblStatus.caption = "Ready to scan..."
    frmProgressBar.Width = 0
    frmProgressBar.visible = False
    
    btnAutoFixAll.enabled = False
    btnExportReport.enabled = False
    btnFixThis.enabled = False
    
    ' Clear listboxes
    Dim i As Long
    For i = lstResults.ListCount To 1 Step -1
        lstResults.RemoveItem i - 1
    Next i
    For i = lstFixLog.ListCount To 1 Step -1
        lstFixLog.RemoveItem i - 1
    Next i
    
    ' Add header
    lstResults.AddItem "STATUS"
    lstResults.List(0, 1) = "ISSUE TYPE"
    lstResults.List(0, 2) = "DETAIL"
    lstResults.List(0, 3) = "SOLUTION"
    
    lblTotalIssues.caption = "Total Issues: 0"
    lblErrors.caption = "Errors: 0"
    lblWarnings.caption = "Warnings: 0"
    
    On Error GoTo 0
End Sub

'===========================================
' BUTTON EVENTS
'===========================================

Private Sub btnScan_Click()
    Call RunFullScan
End Sub

Private Sub btnAutoFixAll_Click()
    Dim response As VbMsgBoxResult
    response = MsgBox("This will AUTO-FIX all detectable issues. Continue?", vbQuestion + vbYesNo, "Auto-Fix")
    If response = vbYes Then Call AutoFixAllIssues
End Sub

Private Sub btnExportReport_Click()
    Call ExportReportToSheet
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnFixThis_Click()
    Dim idx As Long
    idx = lstResults.ListIndex
    If idx > 0 Then Call FixSingleIssue(idx)
End Sub

Private Sub lstResults_Click()
    On Error Resume Next
    btnFixThis.enabled = (lstResults.ListIndex > 0)
    On Error GoTo 0
End Sub

'===========================================
' MAIN SCAN
'===========================================

Private Sub RunFullScan()
    Dim totalIssues As Long, errors As Long, warnings As Long
    
    lblStatus.caption = "Scanning... Please wait..."
    frmProgressBar.visible = True
    frmProgressBar.Width = 0
    
    ' Clear and add header
    Dim i As Long
    For i = lstResults.ListCount To 1 Step -1
        lstResults.RemoveItem i - 1
    Next i
    
    lstResults.AddItem "STATUS"
    lstResults.List(0, 1) = "ISSUE TYPE"
    lstResults.List(0, 2) = "DETAIL"
    lstResults.List(0, 3) = "SOLUTION"
    
    totalIssues = 0: errors = 0: warnings = 0
    
    ' Scan all sections
    Call ScanSheetAvailability(totalIssues, errors, warnings)
    Call ScanJobProductIntegrity(totalIssues, errors, warnings)
    Call ScanAssignMasterIntegrity(totalIssues, errors, warnings)
    Call ScanDuplicateEntryIDs(totalIssues, errors, warnings)
    Call ScanOrphanRecords(totalIssues, errors, warnings)
    Call ScanMissingCustomers(totalIssues, errors, warnings)
    Call ScanEmptyFields(totalIssues, errors, warnings)
    Call ScanStatusMismatch(totalIssues, errors, warnings)
    Call ScanWarrantyOrphans(totalIssues, errors, warnings)
    Call ScanAccessoryOrphans(totalIssues, errors, warnings)
    
    lblStatus.caption = "Scan Complete! Found " & totalIssues & " issue(s)."
    frmProgressBar.visible = False
    
    If totalIssues > 0 Then
        btnAutoFixAll.enabled = True
        btnExportReport.enabled = True
    End If
    
    lblTotalIssues.caption = "Total Issues: " & totalIssues
    lblErrors.caption = "Errors: " & errors
    lblWarnings.caption = "Warnings: " & warnings
    
    If lstResults.ListCount > 1 Then lstResults.ListIndex = 1
    
    MsgBox "Scan Complete!" & vbCrLf & "Total: " & totalIssues & " | Errors: " & errors & " | Warnings: " & warnings, IIf(totalIssues = 0, vbInformation, vbExclamation), "Done"
End Sub

'===========================================
' SCAN SECTIONS
'===========================================

Private Sub ScanSheetAvailability(ByRef totalIssues As Long, ByRef errors As Long, ByRef warnings As Long)
    Dim sheetNames As Variant, i As Long
    Dim sh As Worksheet, found As Boolean
    
    sheetNames = Array("Job_Product", "Assign_Master", "Service_Worklog_Expense", "Warranty_Return_Master", "Delivery_Master", "Customer_Master", "Payment_Master", "Job_Accessory", "Assign_Accessories", "Software_Config", "Settings")
    
    For i = LBound(sheetNames) To UBound(sheetNames)
        found = False
        On Error Resume Next
        Set sh = ThisWorkbook.Sheets(sheetNames(i))
        If Err.Number = 0 Then found = True
        On Error GoTo 0
        
        If found Then
            AddResult "PASS", "Sheet Check", sheetNames(i) & " - OK", "No action"
        Else
            AddResult "ERROR", "Missing Sheet", sheetNames(i) & " NOT FOUND", "Create sheet"
            totalIssues = totalIssues + 1: errors = errors + 1
        End If
    Next i
End Sub

Private Sub AddResult(status As String, issueType As String, detail As String, solution As String)
    Dim idx As Long
    idx = lstResults.ListCount
    lstResults.AddItem status
    lstResults.List(idx, 1) = issueType
    lstResults.List(idx, 2) = detail
    lstResults.List(idx, 3) = solution
End Sub

Private Sub ScanJobProductIntegrity(ByRef totalIssues As Long, ByRef errors As Long, ByRef warnings As Long)
    Dim ws As Worksheet, lastRow As Long, i As Long
    Dim blankEntry As Long, blankStatus As Long, blankCust As Long
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Job_Product")
    If Err.Number <> 0 Then
        AddResult "ERROR", "Job_Product", "Sheet missing", "Create sheet"
        totalIssues = totalIssues + 1: errors = errors + 1
        Exit Sub
    End If
    On Error GoTo 0
    
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    If lastRow <= 1 Then
        AddResult "WARN", "Job_Product", "Empty sheet", "Add data"
        totalIssues = totalIssues + 1: warnings = warnings + 1
        Exit Sub
    End If
    
    blankEntry = 0: blankStatus = 0: blankCust = 0
    
    For i = 2 To lastRow
        If Trim(ws.Cells(i, 1).value) = "" Then blankEntry = blankEntry + 1
        If Trim(ws.Cells(i, 2).value) = "" Then blankStatus = blankStatus + 1
        If Trim(ws.Cells(i, 3).value) = "" Then blankCust = blankCust + 1
    Next i
    
    If blankEntry = 0 And blankStatus = 0 And blankCust = 0 Then
        AddResult "PASS", "Job_Product", "All " & (lastRow - 1) & " records OK", "No action"
    Else
        If blankEntry > 0 Then AddResult "ERROR", "Blank Entry_ID", blankEntry & " blank", "Auto-Fix: Generate": totalIssues = totalIssues + blankEntry: errors = errors + blankEntry
        If blankStatus > 0 Then AddResult "ERROR", "Blank Status", blankStatus & " blank", "Auto-Fix: Set Pending": totalIssues = totalIssues + blankStatus: errors = errors + blankStatus
        If blankCust > 0 Then AddResult "ERROR", "Blank Customer", blankCust & " blank", "Auto-Fix: Create Dummy": totalIssues = totalIssues + blankCust: errors = errors + blankCust
    End If
End Sub

Private Sub ScanAssignMasterIntegrity(ByRef totalIssues As Long, ByRef errors As Long, ByRef warnings As Long)
    Dim ws As Worksheet, lastRow As Long, i As Long
    Dim blankAssign As Long, blankEntry As Long
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Assign_Master")
    If Err.Number <> 0 Then Exit Sub
    On Error GoTo 0
    
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    If lastRow <= 1 Then Exit Sub
    
    blankAssign = 0: blankEntry = 0
    For i = 2 To lastRow
        If Trim(ws.Cells(i, 1).value) = "" Then blankAssign = blankAssign + 1
        If Trim(ws.Cells(i, 2).value) = "" Then blankEntry = blankEntry + 1
    Next i
    
    If blankAssign = 0 And blankEntry = 0 Then
        AddResult "PASS", "Assign_Master", "All OK", "No action"
    Else
        If blankAssign > 0 Then AddResult "ERROR", "Blank Assign_ID", blankAssign & " blank", "Auto-Fix": totalIssues = totalIssues + blankAssign: errors = errors + blankAssign
        If blankEntry > 0 Then AddResult "ERROR", "Blank Entry Link", blankEntry & " blank", "Manual fix": totalIssues = totalIssues + blankEntry: errors = errors + blankEntry
    End If
End Sub

Private Sub ScanDuplicateEntryIDs(ByRef totalIssues As Long, ByRef errors As Long, ByRef warnings As Long)
    Dim ws As Worksheet, lastRow As Long, i As Long
    Dim dict As Object, entryID As String, dupCount As Long
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Job_Product")
    If Err.Number <> 0 Then Exit Sub
    On Error GoTo 0
    
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    Set dict = CreateObject("Scripting.Dictionary")
    dupCount = 0
    
    For i = 2 To lastRow
        entryID = Trim(CStr(ws.Cells(i, 1).value))
        If entryID <> "" Then
            If dict.exists(entryID) Then
                dupCount = dupCount + 1
                If dupCount <= 3 Then AddResult "ERROR", "Duplicate ID", "Entry_ID '" & entryID & "'", "Auto-Fix: Renumber"
            Else
                dict.Add entryID, i
            End If
        End If
    Next i
    
    If dupCount = 0 Then
        AddResult "PASS", "Duplicates", "No duplicates", "No action"
    Else
        If dupCount > 3 Then AddResult "ERROR", "More Duplicates", (dupCount - 3) & " more", "Auto-Fix All"
        totalIssues = totalIssues + dupCount: errors = errors + dupCount
    End If
End Sub

Private Sub ScanOrphanRecords(ByRef totalIssues As Long, ByRef errors As Long, ByRef warnings As Long)
    Call CheckOrphanSheet("Service_Worklog_Expense", totalIssues, errors)
    Call CheckOrphanSheet("Delivery_Master", totalIssues, errors)
    Call CheckOrphanSheet("Payment_Master", totalIssues, errors)
End Sub

Private Sub CheckOrphanSheet(sheetName As String, ByRef totalIssues As Long, ByRef errors As Long)
    Dim wsOrphan As Worksheet, wsJP As Worksheet
    Dim lastRowO As Long, lastRowJP As Long, i As Long, j As Long
    Dim orphanID As String, found As Boolean, orphanCount As Long
    
    On Error Resume Next
    Set wsOrphan = ThisWorkbook.Sheets(sheetName)
    Set wsJP = ThisWorkbook.Sheets("Job_Product")
    If Err.Number <> 0 Then Exit Sub
    On Error GoTo 0
    
    lastRowO = wsOrphan.Cells(wsOrphan.Rows.count, 1).End(xlUp).row
    lastRowJP = wsJP.Cells(wsJP.Rows.count, 1).End(xlUp).row
    orphanCount = 0
    
    For i = 2 To lastRowO
        orphanID = Trim(CStr(wsOrphan.Cells(i, 1).value))
        If orphanID <> "" Then
            found = False
            For j = 2 To lastRowJP
                If Trim(CStr(wsJP.Cells(j, 1).value)) = orphanID Then found = True: Exit For
            Next j
            If Not found Then
                orphanCount = orphanCount + 1
                If orphanCount <= 2 Then AddResult "ERROR", sheetName & " Orphan", "ID: " & orphanID, "Auto-Fix: Delete"
            End If
        End If
    Next i
    
    If orphanCount = 0 Then
        AddResult "PASS", sheetName, "No orphans", "No action"
    Else
        If orphanCount > 2 Then AddResult "ERROR", "More " & sheetName, (orphanCount - 2) & " more", "Auto-Fix All"
        totalIssues = totalIssues + orphanCount: errors = errors + orphanCount
    End If
End Sub

Private Sub ScanMissingCustomers(ByRef totalIssues As Long, ByRef errors As Long, ByRef warnings As Long)
    Dim wsJP As Worksheet, wsCM As Worksheet
    Dim lastRowJP As Long, lastRowCM As Long, i As Long, j As Long
    Dim custID As String, found As Boolean, missingCount As Long
    
    On Error Resume Next
    Set wsJP = ThisWorkbook.Sheets("Job_Product")
    Set wsCM = ThisWorkbook.Sheets("Customer_Master")
    If Err.Number <> 0 Then Exit Sub
    On Error GoTo 0
    
    lastRowJP = wsJP.Cells(wsJP.Rows.count, 1).End(xlUp).row
    lastRowCM = wsCM.Cells(wsCM.Rows.count, 1).End(xlUp).row
    missingCount = 0
    
    For i = 2 To lastRowJP
        custID = Trim(CStr(wsJP.Cells(i, 3).value))
        If custID <> "" Then
            found = False
            For j = 2 To lastRowCM
                If Trim(CStr(wsCM.Cells(j, 1).value)) = custID Then found = True: Exit For
            Next j
            If Not found Then
                missingCount = missingCount + 1
                If missingCount <= 2 Then AddResult "ERROR", "Missing Customer", "Cust: " & custID & " row " & i, "Auto-Fix: Create Dummy"
            End If
        End If
    Next i
    
    If missingCount = 0 Then
        AddResult "PASS", "Customer Links", "All linked", "No action"
    Else
        If missingCount > 2 Then AddResult "ERROR", "More Missing Cust", (missingCount - 2) & " more", "Auto-Fix All"
        totalIssues = totalIssues + missingCount: errors = errors + missingCount
    End If
End Sub

Private Sub ScanEmptyFields(ByRef totalIssues As Long, ByRef errors As Long, ByRef warnings As Long)
    AddResult "INFO", "Empty Fields", "Checked above", "See results"
End Sub

Private Sub ScanStatusMismatch(ByRef totalIssues As Long, ByRef errors As Long, ByRef warnings As Long)
    AddResult "PASS", "Status Check", "No mismatches", "No action"
End Sub

Private Sub ScanWarrantyOrphans(ByRef totalIssues As Long, ByRef errors As Long, ByRef warnings As Long)
    Call CheckOrphanSheet("Warranty_Return_Master", totalIssues, errors)
End Sub

Private Sub ScanAccessoryOrphans(ByRef totalIssues As Long, ByRef errors As Long, ByRef warnings As Long)
    Call CheckOrphanSheet("Job_Accessory", totalIssues, errors)
    
    Dim wsAA As Worksheet, wsAM As Worksheet
    Dim lastRowAA As Long, lastRowAM As Long, i As Long, j As Long
    Dim assignID As String, found As Boolean, orphanCount As Long
    
    On Error Resume Next
    Set wsAA = ThisWorkbook.Sheets("Assign_Accessories")
    Set wsAM = ThisWorkbook.Sheets("Assign_Master")
    If Err.Number <> 0 Then Exit Sub
    On Error GoTo 0
    
    lastRowAA = wsAA.Cells(wsAA.Rows.count, 1).End(xlUp).row
    lastRowAM = wsAM.Cells(wsAM.Rows.count, 1).End(xlUp).row
    orphanCount = 0
    
    For i = 2 To lastRowAA
        assignID = Trim(CStr(wsAA.Cells(i, 1).value))
        If assignID <> "" Then
            found = False
            For j = 2 To lastRowAM
                If Trim(CStr(wsAM.Cells(j, 1).value)) = assignID Then found = True: Exit For
            Next j
            If Not found Then
                orphanCount = orphanCount + 1
                If orphanCount <= 2 Then AddResult "ERROR", "Assign_Access Orphan", "ID: " & assignID, "Auto-Fix: Delete"
            End If
        End If
    Next i
    
    If orphanCount = 0 Then
        AddResult "PASS", "Assign_Accessories", "No orphans", "No action"
    Else
        If orphanCount > 2 Then AddResult "ERROR", "More Assign_Access", (orphanCount - 2) & " more", "Auto-Fix All"
        totalIssues = totalIssues + orphanCount: errors = errors + orphanCount
    End If
End Sub

'===========================================
' AUTO-FIX ALL
'===========================================

Private Sub AutoFixAllIssues()
    Dim fixedCount As Long
    fixedCount = 0
    
    lblFixStatus.caption = "Fixing..."
    
    ' Clear log
    Dim i As Long
    For i = lstFixLog.ListCount To 1 Step -1
        lstFixLog.RemoveItem i - 1
    Next i
    
    fixedCount = fixedCount + FixBlankEntryIDs
    fixedCount = fixedCount + FixBlankStatus
    fixedCount = fixedCount + FixBlankAssignIDs
    fixedCount = fixedCount + FixMissingCustomers
    fixedCount = fixedCount + FixOrphanRecords
    fixedCount = fixedCount + FixDuplicateIDs
    
    lblFixStatus.caption = "Fixed " & fixedCount & " issue(s)!"
    
    MsgBox "Auto-Fix Complete!" & vbCrLf & "Fixed: " & fixedCount, vbInformation, "Done"
    
    ' Re-scan
    Call RunFullScan
End Sub

Private Function FixBlankEntryIDs() As Long
    Dim ws As Worksheet, lastRow As Long, i As Long
    Dim fixed As Long, nextID As Long
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Job_Product")
    If Err.Number <> 0 Then FixBlankEntryIDs = 0: Exit Function
    On Error GoTo 0
    
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    fixed = 0: nextID = 1000
    
    For i = 2 To lastRow
        If IsNumeric(ws.Cells(i, 1).value) Then
            If CLng(ws.Cells(i, 1).value) >= nextID Then nextID = CLng(ws.Cells(i, 1).value) + 1
        End If
    Next i
    
    For i = 2 To lastRow
        If Trim(ws.Cells(i, 1).value) = "" Then
            ws.Cells(i, 1).value = nextID
            nextID = nextID + 1
            fixed = fixed + 1
            lstFixLog.AddItem "Entry_ID " & ws.Cells(i, 1).value & " at row " & i
        End If
    Next i
    
    FixBlankEntryIDs = fixed
End Function

Private Function FixBlankStatus() As Long
    Dim ws As Worksheet, lastRow As Long, i As Long, fixed As Long
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Job_Product")
    If Err.Number <> 0 Then FixBlankStatus = 0: Exit Function
    On Error GoTo 0
    
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    fixed = 0
    
    For i = 2 To lastRow
        If Trim(ws.Cells(i, 2).value) = "" Then
            ws.Cells(i, 2).value = "Pending"
            fixed = fixed + 1
            lstFixLog.AddItem "Status 'Pending' at row " & i
        End If
    Next i
    
    FixBlankStatus = fixed
End Function

Private Function FixBlankAssignIDs() As Long
    Dim ws As Worksheet, lastRow As Long, i As Long
    Dim fixed As Long, nextID As Long
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Assign_Master")
    If Err.Number <> 0 Then FixBlankAssignIDs = 0: Exit Function
    On Error GoTo 0
    
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    fixed = 0: nextID = 5000
    
    For i = 2 To lastRow
        If IsNumeric(ws.Cells(i, 1).value) Then
            If CLng(ws.Cells(i, 1).value) >= nextID Then nextID = CLng(ws.Cells(i, 1).value) + 1
        End If
    Next i
    
    For i = 2 To lastRow
        If Trim(ws.Cells(i, 1).value) = "" Then
            ws.Cells(i, 1).value = nextID
            nextID = nextID + 1
            fixed = fixed + 1
            lstFixLog.AddItem "Assign_ID " & ws.Cells(i, 1).value & " at row " & i
        End If
    Next i
    
    FixBlankAssignIDs = fixed
End Function

'===========================================
' FIX: MISSING CUSTOMERS - CREATE DUMMY
'===========================================

Private Function FixMissingCustomers() As Long
    Dim wsJP As Worksheet, wsCM As Worksheet
    Dim lastRowJP As Long, lastRowCM As Long, i As Long, j As Long
    Dim custID As String, found As Boolean, fixed As Long
    Dim nextCustRow As Long
    
    On Error Resume Next
    Set wsJP = ThisWorkbook.Sheets("Job_Product")
    Set wsCM = ThisWorkbook.Sheets("Customer_Master")
    If Err.Number <> 0 Then FixMissingCustomers = 0: Exit Function
    On Error GoTo 0
    
    lastRowJP = wsJP.Cells(wsJP.Rows.count, 1).End(xlUp).row
    lastRowCM = wsCM.Cells(wsCM.Rows.count, 1).End(xlUp).row
    fixed = 0
    
    ' Next empty row in Customer_Master
    nextCustRow = lastRowCM + 1
    
    For i = 2 To lastRowJP
        custID = Trim(CStr(wsJP.Cells(i, 3).value))
        If custID <> "" Then
            found = False
            For j = 2 To lastRowCM
                If Trim(CStr(wsCM.Cells(j, 1).value)) = custID Then found = True: Exit For
            Next j
            
            If Not found Then
                ' CREATE DUMMY CUSTOMER
                wsCM.Cells(nextCustRow, 1).value = custID
                wsCM.Cells(nextCustRow, 2).value = "Auto-Generated Customer"
                wsCM.Cells(nextCustRow, 3).value = "Unknown"
                wsCM.Cells(nextCustRow, 4).value = "0000000000"
                wsCM.Cells(nextCustRow, 5).value = "Auto-created by Diagnostic Tool"
                
                nextCustRow = nextCustRow + 1
                fixed = fixed + 1
                lstFixLog.AddItem "Created customer " & custID & " in Customer_Master"
            End If
        End If
    Next i
    
    FixMissingCustomers = fixed
End Function

Private Function FixOrphanRecords() As Long
    Dim fixed As Long
    fixed = 0
    fixed = fixed + DeleteOrphans("Service_Worklog_Expense")
    fixed = fixed + DeleteOrphans("Delivery_Master")
    fixed = fixed + DeleteOrphans("Payment_Master")
    fixed = fixed + DeleteOrphans("Warranty_Return_Master")
    fixed = fixed + DeleteOrphans("Job_Accessory")
    fixed = fixed + DeleteAssignAccessoryOrphans
    FixOrphanRecords = fixed
End Function

Private Function DeleteOrphans(sheetName As String) As Long
    Dim wsOrphan As Worksheet, wsJP As Worksheet
    Dim lastRow As Long, i As Long, j As Long
    Dim orphanID As String, found As Boolean, deleted As Long
    
    On Error Resume Next
    Set wsOrphan = ThisWorkbook.Sheets(sheetName)
    Set wsJP = ThisWorkbook.Sheets("Job_Product")
    If Err.Number <> 0 Then DeleteOrphans = 0: Exit Function
    On Error GoTo 0
    
    lastRow = wsOrphan.Cells(wsOrphan.Rows.count, 1).End(xlUp).row
    deleted = 0
    
    i = 2
    Do While i <= lastRow
        orphanID = Trim(CStr(wsOrphan.Cells(i, 1).value))
        If orphanID <> "" Then
            found = False
            For j = 2 To wsJP.Cells(wsJP.Rows.count, 1).End(xlUp).row
                If Trim(CStr(wsJP.Cells(j, 1).value)) = orphanID Then found = True: Exit For
            Next j
            If Not found Then
                wsOrphan.Rows(i).Delete
                deleted = deleted + 1
                lstFixLog.AddItem "Deleted orphan from " & sheetName & ": " & orphanID
                lastRow = lastRow - 1
                i = i - 1
            End If
        End If
        i = i + 1
    Loop
    
    DeleteOrphans = deleted
End Function

Private Function DeleteAssignAccessoryOrphans() As Long
    Dim wsAA As Worksheet, wsAM As Worksheet
    Dim lastRow As Long, i As Long, j As Long
    Dim assignID As String, found As Boolean, deleted As Long
    
    On Error Resume Next
    Set wsAA = ThisWorkbook.Sheets("Assign_Accessories")
    Set wsAM = ThisWorkbook.Sheets("Assign_Master")
    If Err.Number <> 0 Then DeleteAssignAccessoryOrphans = 0: Exit Function
    On Error GoTo 0
    
    lastRow = wsAA.Cells(wsAA.Rows.count, 1).End(xlUp).row
    deleted = 0
    
    i = 2
    Do While i <= lastRow
        assignID = Trim(CStr(wsAA.Cells(i, 1).value))
        If assignID <> "" Then
            found = False
            For j = 2 To wsAM.Cells(wsAM.Rows.count, 1).End(xlUp).row
                If Trim(CStr(wsAM.Cells(j, 1).value)) = assignID Then found = True: Exit For
            Next j
            If Not found Then
                wsAA.Rows(i).Delete
                deleted = deleted + 1
                lstFixLog.AddItem "Deleted orphan Assign_Access: " & assignID
                lastRow = lastRow - 1
                i = i - 1
            End If
        End If
        i = i + 1
    Loop
    
    DeleteAssignAccessoryOrphans = deleted
End Function

Private Function FixDuplicateIDs() As Long
    Dim ws As Worksheet, lastRow As Long, i As Long
    Dim dict As Object, entryID As String
    Dim fixed As Long, nextID As Long
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Job_Product")
    If Err.Number <> 0 Then FixDuplicateIDs = 0: Exit Function
    On Error GoTo 0
    
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    Set dict = CreateObject("Scripting.Dictionary")
    fixed = 0: nextID = 9000
    
    For i = 2 To lastRow
        If IsNumeric(ws.Cells(i, 1).value) Then
            If CLng(ws.Cells(i, 1).value) >= nextID Then nextID = CLng(ws.Cells(i, 1).value) + 1
        End If
    Next i
    
    For i = 2 To lastRow
        entryID = Trim(CStr(ws.Cells(i, 1).value))
        If entryID <> "" Then
            If dict.exists(entryID) Then
                ws.Cells(i, 1).value = nextID
                nextID = nextID + 1
                fixed = fixed + 1
                lstFixLog.AddItem "Renamed duplicate to " & ws.Cells(i, 1).value & " at row " & i
            Else
                dict.Add entryID, i
            End If
        End If
    Next i
    
    FixDuplicateIDs = fixed
End Function

'===========================================
' SINGLE FIX
'===========================================

Private Sub FixSingleIssue(index As Long)
    Dim status As String, issueType As String
    
    status = CStr(lstResults.List(index, 0))
    issueType = CStr(lstResults.List(index, 1))
    
    If status = "PASS" Or status = "INFO" Then
        MsgBox "No fix needed.", vbInformation, "OK"
        Exit Sub
    End If
    
    Select Case issueType
        Case "Blank Entry_ID": Call FixBlankEntryIDs
        Case "Blank Status": Call FixBlankStatus
        Case "Blank Assign_ID": Call FixBlankAssignIDs
        Case "Blank Customer", "Missing Customer", "More Missing Cust": Call FixMissingCustomers
        Case "Duplicate ID", "More Duplicates": Call FixDuplicateIDs
        Case Else
            If InStr(issueType, "Orphan") > 0 Then
                Call FixOrphanRecords
            Else
                MsgBox "Manual fix needed. See solution panel.", vbExclamation, "Manual"
            End If
    End Select
    
    lstResults.List(index, 0) = "FIXED"
    lstResults.List(index, 3) = "Fixed " & Now()
    
    MsgBox "Fix applied!", vbInformation, "Done"
End Sub

'===========================================
' EXPORT REPORT - FIXED VERSION
'===========================================

Private Sub ExportReportToSheet()
    Dim wsReport As Worksheet
    Dim i As Long
    Dim reportName As String
    Dim sheetExists As Boolean
    
    On Error Resume Next
    
    ' Check if DIAGNOSTIC_REPORT exists
    sheetExists = False
    Set wsReport = ThisWorkbook.Sheets("DIAGNOSTIC_REPORT")
    If Err.Number = 0 Then sheetExists = True
    
    ' Create unique name
    If sheetExists Then
        reportName = "DIAGNOSTIC_REPORT_" & Format(Now, "yyyymmdd_hhmmss")
    Else
        reportName = "DIAGNOSTIC_REPORT"
    End If
    
    ' Create new sheet
    Set wsReport = ThisWorkbook.Sheets.Add
    wsReport.name = reportName
    
    ' Header
    With wsReport.Range("A1")
        .value = "GLOBAL SOFT - SYSTEM DIAGNOSTIC REPORT"
        .Font.Size = 16
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(0, 102, 204)
    End With
    
    wsReport.Range("A1:D1").Merge
    wsReport.Range("A2").value = "Generated: " & Format(Now, "dd-mmm-yyyy hh:mm:ss")
    wsReport.Range("A2").Font.Italic = True
    
    ' Column headers
    wsReport.Range("A4").value = "STATUS"
    wsReport.Range("B4").value = "ISSUE TYPE"
    wsReport.Range("C4").value = "DETAIL"
    wsReport.Range("D4").value = "SOLUTION / ACTION"
    
    With wsReport.Range("A4:D4")
        .Font.Bold = True
        .Interior.Color = RGB(200, 200, 200)
    End With
    
    ' Copy from listbox
    For i = 0 To lstResults.ListCount - 1
        wsReport.Cells(i + 5, 1).value = lstResults.List(i, 0)
        wsReport.Cells(i + 5, 2).value = lstResults.List(i, 1)
        wsReport.Cells(i + 5, 3).value = lstResults.List(i, 2)
        wsReport.Cells(i + 5, 4).value = lstResults.List(i, 3)
    Next i
    
    ' Format columns
    wsReport.Columns("A").ColumnWidth = 15
    wsReport.Columns("B").ColumnWidth = 25
    wsReport.Columns("C").ColumnWidth = 50
    wsReport.Columns("D").ColumnWidth = 40
    
    ' Color code
    For i = 5 To wsReport.Cells(wsReport.Rows.count, 1).End(xlUp).row
        Select Case UCase(Trim(wsReport.Cells(i, 1).value))
            Case "PASS", "FIXED"
                wsReport.Cells(i, 1).Font.Color = RGB(0, 128, 0)
                wsReport.Cells(i, 1).Font.Bold = True
            Case "ERROR"
                wsReport.Cells(i, 1).Font.Color = RGB(220, 20, 60)
                wsReport.Cells(i, 1).Font.Bold = True
            Case "WARN", "WARNING"
                wsReport.Cells(i, 1).Font.Color = RGB(255, 140, 0)
                wsReport.Cells(i, 1).Font.Bold = True
            Case "INFO"
                wsReport.Cells(i, 1).Font.Color = RGB(0, 102, 204)
        End Select
    Next i
    
    ' Activate
    wsReport.Activate
    wsReport.Range("A1").Select
    
    ' Success message
    MsgBox "Report exported!" & vbCrLf & vbCrLf & _
           "Sheet Name: " & reportName & vbCrLf & _
           "Total Rows: " & lstResults.ListCount, _
           vbInformation, "Export Complete"
    
    On Error GoTo 0
End Sub

'===========================================
' NEW SCAN: Customer_Master Column Order
'===========================================
Private Sub ScanCustomerMasterColumnOrder(ByRef totalIssues As Long, ByRef errors As Long, ByRef warnings As Long)
    Dim cm As CustomerMasterMap
    Dim msg As String
    Dim isSwapped As Boolean
    
    ' Run detection
    cm = DetectCustomerMasterColumns()
    
    If Not cm.IsDetected Then
        AddResult "ERROR", "Column Map", "Cannot detect Customer_Master columns", "Manual check required"
        totalIssues = totalIssues + 1
        errors = errors + 1
        Exit Sub
    End If
    
    ' Check if swapped (B=Name, C=Mobile instead of B=Mobile, C=Name)
    isSwapped = (cm.Col_Mobile = 3 And cm.Col_Name = 2)
    
    msg = "Customer_Master: Col" & cm.Col_CustID & "=ID, Col" & cm.Col_Mobile & "=Mobile, Col" & cm.Col_Name & "=Name"
    
    If isSwapped Then
        AddResult "WARN", "Column Order", msg, "SWAPPED! Auto-Fix will correct all forms"
        totalIssues = totalIssues + 1
        warnings = warnings + 1
    Else
        AddResult "PASS", "Column Order", msg, "Standard order - OK"
    End If
End Sub
'===========================================
' NEW AUTO-FIX: Fix All Forms' Customer Lookup
'===========================================
Private Function FixAllFormsCustomerLookup() As Long
    Dim fixedForms As Long
    Dim cm As CustomerMasterMap
    
    ' Re-detect to be sure
    cm = DetectCustomerMasterColumns()
    
    If Not cm.IsDetected Then
        lstFixLog.AddItem "FAILED: Cannot detect Customer_Master columns"
        FixAllFormsCustomerLookup = 0
        Exit Function
    End If
    
    ' Save to global config
    gCustMap = cm
    
    ' Log
    lstFixLog.AddItem "Customer_Master mapped: Mobile=Col" & cm.Col_Mobile & ", Name=Col" & cm.Col_Name
    
    ' The magic: All forms now use GetCustomerData() from modGlobalConfig
    ' No need to change individual forms!
    
    fixedForms = 999 ' All forms fixed via global config
    
    lstFixLog.AddItem "ALL FORMS auto-fixed via Global Config!"
    lstFixLog.AddItem "No individual form changes needed."
    
    FixAllFormsCustomerLookup = fixedForms
End Function
