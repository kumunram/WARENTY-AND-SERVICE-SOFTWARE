Attribute VB_Name = "diagnostic"
Option Explicit

'===========================================
' GLOBAL SOFT - COMPLETE SYSTEM DIAGNOSTIC
' Run this to find ALL problems in your file
' Paste this entire code in a new Module and Run "RunCompleteDiagnostic"
'===========================================

Sub RunCompleteDiagnostic()
    Dim wsReport As Worksheet
    Dim r As Long
    Dim totalIssues As Long

    ' Create/Reset Report Sheet
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Sheets("DIAGNOSTIC_REPORT").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    Set wsReport = ThisWorkbook.Sheets.Add
    wsReport.name = "DIAGNOSTIC_REPORT"

    ' ========== HEADER ==========
    With wsReport.Range("A1")
        .value = "GLOBAL SOFT - SYSTEM DIAGNOSTIC REPORT"
        .Font.Size = 16
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(0, 102, 204)
    End With
    wsReport.Range("A1:D1").Merge

    wsReport.Range("A2").value = "Report Generated: " & Now()
    wsReport.Range("A2").Font.Italic = True

    r = 4
    totalIssues = 0

    '===========================================
    ' SECTION 1: SHEET AVAILABILITY CHECK
    '===========================================
    r = AddSectionHeader(wsReport, r, "1. SHEET AVAILABILITY CHECK")
    r = CheckSheetExists(wsReport, r, "Job_Product", totalIssues)
    r = CheckSheetExists(wsReport, r, "Assign_Master", totalIssues)
    r = CheckSheetExists(wsReport, r, "Service_Worklog_Expense", totalIssues)
    r = CheckSheetExists(wsReport, r, "Warranty_Return_Master", totalIssues)
    r = CheckSheetExists(wsReport, r, "Delivery_Master", totalIssues)
    r = CheckSheetExists(wsReport, r, "Customer_Master", totalIssues)
    r = CheckSheetExists(wsReport, r, "Payment_Master", totalIssues)
    r = CheckSheetExists(wsReport, r, "Job_Accessory", totalIssues)
    r = CheckSheetExists(wsReport, r, "Assign_Accessories", totalIssues)
    r = CheckSheetExists(wsReport, r, "Software_Config", totalIssues)
    r = CheckSheetExists(wsReport, r, "Settings", totalIssues)

    '===========================================
    ' SECTION 2: Job_Product DATA INTEGRITY
    '===========================================
    r = AddSectionHeader(wsReport, r, "2. Job_Product DATA INTEGRITY")
    r = CheckJobProductIntegrity(wsReport, r, totalIssues)

    '===========================================
    ' SECTION 3: Assign_Master DATA INTEGRITY
    '===========================================
    r = AddSectionHeader(wsReport, r, "3. Assign_Master DATA INTEGRITY")
    r = CheckAssignMasterIntegrity(wsReport, r, totalIssues)

    '===========================================
    ' SECTION 4: STATUS MISMATCH BETWEEN SHEETS
    '===========================================
    r = AddSectionHeader(wsReport, r, "4. STATUS MISMATCH (Job_Product vs Assign_Master)")
    r = CheckStatusMismatch(wsReport, r, totalIssues)

    '===========================================
    ' SECTION 5: Service Work Log ORPHAN CHECK
    '===========================================
    r = AddSectionHeader(wsReport, r, "5. Service_Worklog ORPHAN CHECK")
    r = CheckServiceWorklogOrphans(wsReport, r, totalIssues)

    '===========================================
    ' SECTION 6: Delivery ORPHAN CHECK
    '===========================================
    r = AddSectionHeader(wsReport, r, "6. Delivery ORPHAN CHECK")
    r = CheckDeliveryOrphans(wsReport, r, totalIssues)

    '===========================================
    ' SECTION 7: Duplicate Entry ID Check
    '===========================================
    r = AddSectionHeader(wsReport, r, "7. DUPLICATE Entry ID Check")
    r = CheckDuplicateEntryIDs(wsReport, r, totalIssues)

    '===========================================
    ' SECTION 8: Missing Customer Links
    '===========================================
    r = AddSectionHeader(wsReport, r, "8. Missing Customer Links")
    r = CheckMissingCustomers(wsReport, r, totalIssues)

    '===========================================
    ' SECTION 9: Payment ORPHAN Check
    '===========================================
    r = AddSectionHeader(wsReport, r, "9. Payment ORPHAN Check")
    r = CheckPaymentOrphans(wsReport, r, totalIssues)

    '===========================================
    ' SECTION 10: Warranty ORPHAN Check
    '===========================================
    r = AddSectionHeader(wsReport, r, "10. Warranty_Return ORPHAN Check")
    r = CheckWarrantyOrphans(wsReport, r, totalIssues)

    '===========================================
    ' SECTION 11: Job_Accessory ORPHAN Check
    '===========================================
    r = AddSectionHeader(wsReport, r, "11. Job_Accessory ORPHAN Check")
    r = CheckJobAccessoryOrphans(wsReport, r, totalIssues)

    '===========================================
    ' SECTION 12: Assign_Accessories ORPHAN Check
    '===========================================
    r = AddSectionHeader(wsReport, r, "12. Assign_Accessories ORPHAN Check")
    r = CheckAssignAccessoryOrphans(wsReport, r, totalIssues)

    '===========================================
    ' SECTION 13: Empty/Blank Critical Fields
    '===========================================
    r = AddSectionHeader(wsReport, r, "13. EMPTY CRITICAL FIELDS")
    r = CheckEmptyCriticalFields(wsReport, r, totalIssues)

    '===========================================
    ' SECTION 14: Summary & Recommendations
    '===========================================
    r = AddSectionHeader(wsReport, r, "14. SUMMARY & RECOMMENDATIONS")
    r = GenerateSummary(wsReport, r, totalIssues)

    ' Format Report
    wsReport.Columns("A").ColumnWidth = 25
    wsReport.Columns("B").ColumnWidth = 40
    wsReport.Columns("C").ColumnWidth = 50
    wsReport.Columns("D").ColumnWidth = 20

    ' Freeze panes
    wsReport.Range("A5").Select
    ActiveWindow.FreezePanes = True

    ' Auto-fit and activate
    wsReport.Activate

    ' Final Message
    If totalIssues = 0 Then
        MsgBox "Diagnostic Complete!" & vbCrLf & vbCrLf & _
               "? NO ISSUES FOUND!" & vbCrLf & _
               "Your file is clean. Check 'DIAGNOSTIC_REPORT' sheet for details.", _
               vbInformation, "GLOBAL SOFT DIAGNOSTIC - ALL CLEAR"
    Else
        MsgBox "Diagnostic Complete!" & vbCrLf & vbCrLf & _
               "?? TOTAL ISSUES FOUND: " & totalIssues & vbCrLf & _
               "Check 'DIAGNOSTIC_REPORT' sheet for full details and solutions.", _
               vbExclamation, "GLOBAL SOFT DIAGNOSTIC - ISSUES FOUND"
    End If
End Sub

'===========================================
' HELPER: Add Section Header
'===========================================
Private Function AddSectionHeader(ws As Worksheet, rowNum As Long, title As String) As Long
    With ws.Cells(rowNum, 1)
        .value = title
        .Font.Size = 12
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(0, 102, 204)
    End With
    ws.Range(ws.Cells(rowNum, 1), ws.Cells(rowNum, 4)).Merge
    ws.Cells(rowNum, 1).HorizontalAlignment = xlLeft
    AddSectionHeader = rowNum + 2
End Function

'===========================================
' HELPER: Add Result Row
'===========================================
Private Function AddResult(ws As Worksheet, rowNum As Long, status As String, _
                          issue As String, detail As String, solution As String) As Long
    ws.Cells(rowNum, 1).value = status
    ws.Cells(rowNum, 2).value = issue
    ws.Cells(rowNum, 3).value = detail
    ws.Cells(rowNum, 4).value = solution

    ' Color coding
    Select Case status
        Case "? PASS"
            ws.Cells(rowNum, 1).Font.Color = RGB(0, 128, 0)
            ws.Cells(rowNum, 1).Font.Bold = True
        Case "?? WARNING"
            ws.Cells(rowNum, 1).Font.Color = RGB(255, 140, 0)
            ws.Cells(rowNum, 1).Font.Bold = True
        Case "? ERROR"
            ws.Cells(rowNum, 1).Font.Color = RGB(220, 20, 60)
            ws.Cells(rowNum, 1).Font.Bold = True
        Case "?? INFO"
            ws.Cells(rowNum, 1).Font.Color = RGB(0, 102, 204)
    End Select

    AddResult = rowNum + 1
End Function

'===========================================
' SECTION 1: Check Sheet Exists
'===========================================
Private Function CheckSheetExists(ws As Worksheet, rowNum As Long, _
                                  sheetName As String, ByRef issueCount As Long) As Long
    Dim sh As Worksheet
    Dim found As Boolean
    found = False

    On Error Resume Next
    Set sh = ThisWorkbook.Sheets(sheetName)
    If Err.Number = 0 Then found = True
    On Error GoTo 0

    If found Then
        CheckSheetExists = AddResult(ws, rowNum, "? PASS", _
            "Sheet: " & sheetName, "Sheet exists and is accessible.", "No action needed.")
    Else
        issueCount = issueCount + 1
        CheckSheetExists = AddResult(ws, rowNum, "? ERROR", _
            "Sheet: " & sheetName, "Sheet is MISSING!", _
            "Create sheet or restore from backup. Many features will fail without this sheet.")
    End If
End Function

'===========================================
' SECTION 2: Job_Product Data Integrity
'===========================================
Private Function CheckJobProductIntegrity(ws As Worksheet, rowNum As Long, _
                                          ByRef issueCount As Long) As Long
    Dim wsJP As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim blankEntryID As Long
    Dim blankStatus As Long
    Dim blankCustomer As Long
    Dim invalidStatus As Long
    Dim validStatuses As Variant
    Dim r As Long

    r = rowNum
    validStatuses = Array("Pending", "Assigned", "In Progress", "Completed", _
                          "Delivered", "Closed", "Cancelled", "Warranty", "Returned")

    On Error Resume Next
    Set wsJP = ThisWorkbook.Sheets("Job_Product")
    If Err.Number <> 0 Then
        issueCount = issueCount + 1
        CheckJobProductIntegrity = AddResult(ws, r, "? ERROR", _
            "Job_Product Check", "Sheet not found!", "Cannot check integrity without sheet.")
        Exit Function
    End If
    On Error GoTo 0

    lastRow = wsJP.Cells(wsJP.Rows.count, 1).End(xlUp).row

    If lastRow <= 1 Then
        issueCount = issueCount + 1
        CheckJobProductIntegrity = AddResult(ws, r, "?? WARNING", _
            "Job_Product Check", "Sheet is EMPTY (no data rows)!", _
            "Verify data import or check if data starts from correct row.")
        Exit Function
    End If

    blankEntryID = 0
    blankStatus = 0
    blankCustomer = 0
    invalidStatus = 0

    For i = 2 To lastRow
        ' Check Entry_ID (Column A)
        If Trim(wsJP.Cells(i, 1).value) = "" Then blankEntryID = blankEntryID + 1

        ' Check Status (Column B or detect)
        If Trim(wsJP.Cells(i, 2).value) = "" Then
            blankStatus = blankStatus + 1
        Else
            ' Check if status is valid
            If Not IsInArray(Trim(wsJP.Cells(i, 2).value), validStatuses) Then
                invalidStatus = invalidStatus + 1
            End If
        End If

        ' Check Customer_ID (Column C or detect)
        If Trim(wsJP.Cells(i, 3).value) = "" Then blankCustomer = blankCustomer + 1
    Next i

    ' Report findings
    If blankEntryID = 0 And blankStatus = 0 And blankCustomer = 0 And invalidStatus = 0 Then
        r = AddResult(ws, r, "? PASS", "Job_Product Integrity", _
            "All " & (lastRow - 1) & " records checked. No blank critical fields.", _
            "No action needed.")
    Else
        If blankEntryID > 0 Then
            issueCount = issueCount + blankEntryID
            r = AddResult(ws, r, "? ERROR", "Blank Entry_ID", _
                blankEntryID & " rows have blank Entry_ID.", _
                "Fill Entry_ID for these rows. Entry_ID is primary key and cannot be blank.")
        End If
        If blankStatus > 0 Then
            issueCount = issueCount + blankStatus
            r = AddResult(ws, r, "? ERROR", "Blank Status", _
                blankStatus & " rows have blank Status.", _
                "Fill Status for these rows. Valid: Pending, Assigned, In Progress, Completed, Delivered, Closed, Cancelled, Warranty, Returned.")
        End If
        If invalidStatus > 0 Then
            issueCount = issueCount + invalidStatus
            r = AddResult(ws, r, "?? WARNING", "Invalid Status", _
                invalidStatus & " rows have invalid/unrecognized Status values.", _
                "Correct status to one of the valid values listed above.")
        End If
        If blankCustomer > 0 Then
            issueCount = issueCount + blankCustomer
            r = AddResult(ws, r, "? ERROR", "Blank Customer_ID", _
                blankCustomer & " rows have blank Customer_ID.", _
                "Fill Customer_ID. Every job must be linked to a customer.")
        End If
    End If

    CheckJobProductIntegrity = r
End Function

'===========================================
' SECTION 3: Assign_Master Data Integrity
'===========================================
Private Function CheckAssignMasterIntegrity(ws As Worksheet, rowNum As Long, _
                                            ByRef issueCount As Long) As Long
    Dim wsAM As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim blankAssignID As Long
    Dim blankEntryID As Long
    Dim blankTechnician As Long
    Dim r As Long

    r = rowNum

    On Error Resume Next
    Set wsAM = ThisWorkbook.Sheets("Assign_Master")
    If Err.Number <> 0 Then
        issueCount = issueCount + 1
        CheckAssignMasterIntegrity = AddResult(ws, r, "? ERROR", _
            "Assign_Master Check", "Sheet not found!", "Cannot check integrity without sheet.")
        Exit Function
    End If
    On Error GoTo 0

    lastRow = wsAM.Cells(wsAM.Rows.count, 1).End(xlUp).row

    If lastRow <= 1 Then
        issueCount = issueCount + 1
        CheckAssignMasterIntegrity = AddResult(ws, r, "?? WARNING", _
            "Assign_Master Check", "Sheet is EMPTY!", _
            "Verify data import or check starting row.")
        Exit Function
    End If

    blankAssignID = 0
    blankEntryID = 0
    blankTechnician = 0

    For i = 2 To lastRow
        If Trim(wsAM.Cells(i, 1).value) = "" Then blankAssignID = blankAssignID + 1
        If Trim(wsAM.Cells(i, 2).value) = "" Then blankEntryID = blankEntryID + 1
        If Trim(wsAM.Cells(i, 3).value) = "" Then blankTechnician = blankTechnician + 1
    Next i

    If blankAssignID = 0 And blankEntryID = 0 And blankTechnician = 0 Then
        r = AddResult(ws, r, "? PASS", "Assign_Master Integrity", _
            "All " & (lastRow - 1) & " records checked. No blank critical fields.", _
            "No action needed.")
    Else
        If blankAssignID > 0 Then
            issueCount = issueCount + blankAssignID
            r = AddResult(ws, r, "? ERROR", "Blank Assign_ID", _
                blankAssignID & " rows have blank Assign_ID.", _
                "Fill Assign_ID. This is primary key for assignments.")
        End If
        If blankEntryID > 0 Then
            issueCount = issueCount + blankEntryID
            r = AddResult(ws, r, "? ERROR", "Blank Entry_ID (Link)", _
                blankEntryID & " rows have blank Entry_ID.", _
                "Fill Entry_ID to link assignment to Job_Product. Orphan assignments cannot be tracked.")
        End If
        If blankTechnician > 0 Then
            issueCount = issueCount + blankTechnician
            r = AddResult(ws, r, "?? WARNING", "Blank Technician", _
                blankTechnician & " rows have blank Technician name.", _
                "Fill Technician name for accountability.")
        End If
    End If

    CheckAssignMasterIntegrity = r
End Function

'===========================================
' SECTION 4: Status Mismatch (Job_Product vs Assign_Master)
'===========================================
Private Function CheckStatusMismatch(ws As Worksheet, rowNum As Long, _
                                     ByRef issueCount As Long) As Long
    Dim wsJP As Worksheet, wsAM As Worksheet
    Dim lastRowJP As Long, lastRowAM As Long
    Dim i As Long, j As Long
    Dim jpEntryID As String, jpStatus As String
    Dim amEntryID As String, amStatus As String
    Dim mismatchCount As Long
    Dim r As Long
    Dim dictJP As Object

    r = rowNum
    mismatchCount = 0

    On Error Resume Next
    Set wsJP = ThisWorkbook.Sheets("Job_Product")
    Set wsAM = ThisWorkbook.Sheets("Assign_Master")
    If Err.Number <> 0 Then
        issueCount = issueCount + 1
        CheckStatusMismatch = AddResult(ws, r, "? ERROR", _
            "Status Mismatch Check", "Required sheets missing!", _
            "Both Job_Product and Assign_Master must exist.")
        Exit Function
    End If
    On Error GoTo 0

    lastRowJP = wsJP.Cells(wsJP.Rows.count, 1).End(xlUp).row
    lastRowAM = wsAM.Cells(wsAM.Rows.count, 1).End(xlUp).row

    ' Build dictionary of Job_Product statuses
    Set dictJP = CreateObject("Scripting.Dictionary")
    For i = 2 To lastRowJP
        jpEntryID = Trim(CStr(wsJP.Cells(i, 1).value))
        jpStatus = Trim(CStr(wsJP.Cells(i, 2).value))
        If jpEntryID <> "" Then
            dictJP(jpEntryID) = jpStatus
        End If
    Next i

    ' Check Assign_Master against Job_Product
    For i = 2 To lastRowAM
        amEntryID = Trim(CStr(wsAM.Cells(i, 2).value))
        amStatus = Trim(CStr(wsAM.Cells(i, 4).value)) ' Assuming status in col D

        If amEntryID <> "" Then
            If dictJP.exists(amEntryID) Then
                jpStatus = dictJP(amEntryID)
                ' Simple mismatch check - can be customized
                If jpStatus = "Pending" And amStatus = "Completed" Then
                    mismatchCount = mismatchCount + 1
                    If mismatchCount <= 10 Then ' Limit detailed output
                        r = AddResult(ws, r, "?? WARNING", "Status Mismatch", _
                            "Entry_ID: " & amEntryID & " | JP Status: " & jpStatus & " | AM Status: " & amStatus, _
                            "Verify which status is correct. Update the incorrect sheet.")
                    End If
                End If
            End If
        End If
    Next i

    If mismatchCount = 0 Then
        r = AddResult(ws, r, "? PASS", "Status Mismatch", _
            "No obvious status mismatches found.", _
            "No action needed.")
    Else
        issueCount = issueCount + mismatchCount
        If mismatchCount > 10 Then
            r = AddResult(ws, r, "?? INFO", "More Mismatches", _
                "... and " & (mismatchCount - 10) & " more mismatches found.", _
                "Review all records manually or increase the display limit in code.")
        End If
    End If

    CheckStatusMismatch = r
End Function

'===========================================
' SECTION 5: Service_Worklog Orphan Check
'===========================================
Private Function CheckServiceWorklogOrphans(ws As Worksheet, rowNum As Long, _
                                            ByRef issueCount As Long) As Long
    Dim wsSW As Worksheet, wsJP As Worksheet
    Dim lastRowSW As Long, lastRowJP As Long
    Dim i As Long, j As Long
    Dim swEntryID As String
    Dim found As Boolean
    Dim orphanCount As Long
    Dim r As Long

    r = rowNum
    orphanCount = 0

    On Error Resume Next
    Set wsSW = ThisWorkbook.Sheets("Service_Worklog_Expense")
    Set wsJP = ThisWorkbook.Sheets("Job_Product")
    If Err.Number <> 0 Then
        issueCount = issueCount + 1
        CheckServiceWorklogOrphans = AddResult(ws, r, "? ERROR", _
            "Worklog Orphan Check", "Required sheets missing!", _
            "Both Service_Worklog_Expense and Job_Product must exist.")
        Exit Function
    End If
    On Error GoTo 0

    lastRowSW = wsSW.Cells(wsSW.Rows.count, 1).End(xlUp).row
    lastRowJP = wsJP.Cells(wsJP.Rows.count, 1).End(xlUp).row

    For i = 2 To lastRowSW
        swEntryID = Trim(CStr(wsSW.Cells(i, 1).value))
        If swEntryID <> "" Then
            found = False
            For j = 2 To lastRowJP
                If Trim(CStr(wsJP.Cells(j, 1).value)) = swEntryID Then
                    found = True
                    Exit For
                End If
            Next j
            If Not found Then
                orphanCount = orphanCount + 1
                If orphanCount <= 10 Then
                    r = AddResult(ws, r, "? ERROR", "Worklog Orphan", _
                        "Entry_ID: " & swEntryID & " in Service_Worklog_Expense (Row " & i & ")", _
                        "Either add this Entry_ID to Job_Product OR delete this orphan worklog record.")
                End If
            End If
        End If
    Next i

    If orphanCount = 0 Then
        r = AddResult(ws, r, "? PASS", "Worklog Orphans", _
            "All worklog records are linked to valid jobs.", _
            "No action needed.")
    Else
        issueCount = issueCount + orphanCount
        If orphanCount > 10 Then
            r = AddResult(ws, r, "?? INFO", "More Orphans", _
                "... and " & (orphanCount - 10) & " more orphan worklog records.", _
                "Use Find feature to locate all orphan Entry_IDs in Service_Worklog_Expense.")
        End If
    End If

    CheckServiceWorklogOrphans = r
End Function

'===========================================
' SECTION 6: Delivery Orphan Check
'===========================================
Private Function CheckDeliveryOrphans(ws As Worksheet, rowNum As Long, _
                                      ByRef issueCount As Long) As Long
    Dim wsDM As Worksheet, wsJP As Worksheet
    Dim lastRowDM As Long, lastRowJP As Long
    Dim i As Long, j As Long
    Dim dmEntryID As String
    Dim found As Boolean
    Dim orphanCount As Long
    Dim r As Long

    r = rowNum
    orphanCount = 0

    On Error Resume Next
    Set wsDM = ThisWorkbook.Sheets("Delivery_Master")
    Set wsJP = ThisWorkbook.Sheets("Job_Product")
    If Err.Number <> 0 Then
        issueCount = issueCount + 1
        CheckDeliveryOrphans = AddResult(ws, r, "? ERROR", _
            "Delivery Orphan Check", "Required sheets missing!", _
            "Both Delivery_Master and Job_Product must exist.")
        Exit Function
    End If
    On Error GoTo 0

    lastRowDM = wsDM.Cells(wsDM.Rows.count, 1).End(xlUp).row
    lastRowJP = wsJP.Cells(wsJP.Rows.count, 1).End(xlUp).row

    For i = 2 To lastRowDM
        dmEntryID = Trim(CStr(wsDM.Cells(i, 1).value))
        If dmEntryID <> "" Then
            found = False
            For j = 2 To lastRowJP
                If Trim(CStr(wsJP.Cells(j, 1).value)) = dmEntryID Then
                    found = True
                    Exit For
                End If
            Next j
            If Not found Then
                orphanCount = orphanCount + 1
                If orphanCount <= 10 Then
                    r = AddResult(ws, r, "? ERROR", "Delivery Orphan", _
                        "Entry_ID: " & dmEntryID & " in Delivery_Master (Row " & i & ")", _
                        "Add Entry_ID to Job_Product OR delete this orphan delivery record.")
                End If
            End If
        End If
    Next i

    If orphanCount = 0 Then
        r = AddResult(ws, r, "? PASS", "Delivery Orphans", _
            "All delivery records are linked to valid jobs.", _
            "No action needed.")
    Else
        issueCount = issueCount + orphanCount
        If orphanCount > 10 Then
            r = AddResult(ws, r, "?? INFO", "More Orphans", _
                "... and " & (orphanCount - 10) & " more orphan delivery records.", _
                "Use Find feature to locate all orphan Entry_IDs in Delivery_Master.")
        End If
    End If

    CheckDeliveryOrphans = r
End Function

'===========================================
' SECTION 7: Duplicate Entry ID Check
'===========================================
Private Function CheckDuplicateEntryIDs(ws As Worksheet, rowNum As Long, _
                                        ByRef issueCount As Long) As Long
    Dim wsJP As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim entryID As String
    Dim dict As Object
    Dim dupCount As Long
    Dim r As Long

    r = rowNum
    dupCount = 0

    On Error Resume Next
    Set wsJP = ThisWorkbook.Sheets("Job_Product")
    If Err.Number <> 0 Then
        issueCount = issueCount + 1
        CheckDuplicateEntryIDs = AddResult(ws, r, "? ERROR", _
            "Duplicate Check", "Job_Product sheet not found!", _
            "Cannot check duplicates without Job_Product sheet.")
        Exit Function
    End If
    On Error GoTo 0

    lastRow = wsJP.Cells(wsJP.Rows.count, 1).End(xlUp).row
    Set dict = CreateObject("Scripting.Dictionary")

    For i = 2 To lastRow
        entryID = Trim(CStr(wsJP.Cells(i, 1).value))
        If entryID <> "" Then
            If dict.exists(entryID) Then
                dupCount = dupCount + 1
                If dupCount <= 10 Then
                    r = AddResult(ws, r, "? ERROR", "Duplicate Entry_ID", _
                        "Entry_ID '" & entryID & "' appears at rows " & dict(entryID) & " and " & i, _
                        "Make Entry_ID unique. Add suffix or renumber. Duplicates break lookups and reports.")
                End If
            Else
                dict.Add entryID, i
            End If
        End If
    Next i

    If dupCount = 0 Then
        r = AddResult(ws, r, "? PASS", "Duplicate Entry_IDs", _
            "All " & (lastRow - 1) & " Entry_IDs are unique.", _
            "No action needed.")
    Else
        issueCount = issueCount + dupCount
        If dupCount > 10 Then
            r = AddResult(ws, r, "?? INFO", "More Duplicates", _
                "... and " & (dupCount - 10) & " more duplicate Entry_IDs found.", _
                "Use Conditional Formatting > Highlight Duplicates to find all duplicates visually.")
        End If
    End If

    CheckDuplicateEntryIDs = r
End Function

'===========================================
' SECTION 8: Missing Customer Links
'===========================================
Private Function CheckMissingCustomers(ws As Worksheet, rowNum As Long, _
                                       ByRef issueCount As Long) As Long
    Dim wsJP As Worksheet, wsCM As Worksheet
    Dim lastRowJP As Long, lastRowCM As Long
    Dim i As Long, j As Long
    Dim custID As String
    Dim found As Boolean
    Dim missingCount As Long
    Dim r As Long

    r = rowNum
    missingCount = 0

    On Error Resume Next
    Set wsJP = ThisWorkbook.Sheets("Job_Product")
    Set wsCM = ThisWorkbook.Sheets("Customer_Master")
    If Err.Number <> 0 Then
        issueCount = issueCount + 1
        CheckMissingCustomers = AddResult(ws, r, "? ERROR", _
            "Customer Link Check", "Required sheets missing!", _
            "Both Job_Product and Customer_Master must exist.")
        Exit Function
    End If
    On Error GoTo 0

    lastRowJP = wsJP.Cells(wsJP.Rows.count, 1).End(xlUp).row
    lastRowCM = wsCM.Cells(wsCM.Rows.count, 1).End(xlUp).row

    For i = 2 To lastRowJP
        custID = Trim(CStr(wsJP.Cells(i, 3).value))
        If custID <> "" Then
            found = False
            For j = 2 To lastRowCM
                If Trim(CStr(wsCM.Cells(j, 1).value)) = custID Then
                    found = True
                    Exit For
                End If
            Next j
            If Not found Then
                missingCount = missingCount + 1
                If missingCount <= 10 Then
                    r = AddResult(ws, r, "? ERROR", "Missing Customer", _
                        "Customer_ID '" & custID & "' in Job_Product (Row " & i & ") not found in Customer_Master", _
                        "Add Customer_ID to Customer_Master OR correct the Customer_ID in Job_Product.")
                End If
            End If
        End If
    Next i

    If missingCount = 0 Then
        r = AddResult(ws, r, "? PASS", "Customer Links", _
            "All customer links in Job_Product are valid.", _
            "No action needed.")
    Else
        issueCount = issueCount + missingCount
        If missingCount > 10 Then
            r = AddResult(ws, r, "?? INFO", "More Missing", _
                "... and " & (missingCount - 10) & " more missing customer links.", _
                "Use VLOOKUP or XLOOKUP to validate all Customer_IDs against Customer_Master.")
        End If
    End If

    CheckMissingCustomers = r
End Function

'===========================================
' SECTION 9: Payment Orphan Check
'===========================================
Private Function CheckPaymentOrphans(ws As Worksheet, rowNum As Long, _
                                     ByRef issueCount As Long) As Long
    Dim wsPM As Worksheet, wsJP As Worksheet
    Dim lastRowPM As Long, lastRowJP As Long
    Dim i As Long, j As Long
    Dim pmEntryID As String
    Dim found As Boolean
    Dim orphanCount As Long
    Dim r As Long

    r = rowNum
    orphanCount = 0

    On Error Resume Next
    Set wsPM = ThisWorkbook.Sheets("Payment_Master")
    Set wsJP = ThisWorkbook.Sheets("Job_Product")
    If Err.Number <> 0 Then
        issueCount = issueCount + 1
        CheckPaymentOrphans = AddResult(ws, r, "? ERROR", _
            "Payment Orphan Check", "Required sheets missing!", _
            "Both Payment_Master and Job_Product must exist.")
        Exit Function
    End If
    On Error GoTo 0

    lastRowPM = wsPM.Cells(wsPM.Rows.count, 1).End(xlUp).row
    lastRowJP = wsJP.Cells(wsJP.Rows.count, 1).End(xlUp).row

    For i = 2 To lastRowPM
        pmEntryID = Trim(CStr(wsPM.Cells(i, 1).value))
        If pmEntryID <> "" Then
            found = False
            For j = 2 To lastRowJP
                If Trim(CStr(wsJP.Cells(j, 1).value)) = pmEntryID Then
                    found = True
                    Exit For
                End If
            Next j
            If Not found Then
                orphanCount = orphanCount + 1
                If orphanCount <= 10 Then
                    r = AddResult(ws, r, "? ERROR", "Payment Orphan", _
                        "Entry_ID: " & pmEntryID & " in Payment_Master (Row " & i & ")", _
                        "Add Entry_ID to Job_Product OR delete this orphan payment record.")
                End If
            End If
        End If
    Next i

    If orphanCount = 0 Then
        r = AddResult(ws, r, "? PASS", "Payment Orphans", _
            "All payment records are linked to valid jobs.", _
            "No action needed.")
    Else
        issueCount = issueCount + orphanCount
        If orphanCount > 10 Then
            r = AddResult(ws, r, "?? INFO", "More Orphans", _
                "... and " & (orphanCount - 10) & " more orphan payment records.", _
                "Use Find feature to locate all orphan Entry_IDs in Payment_Master.")
        End If
    End If

    CheckPaymentOrphans = r
End Function

'===========================================
' SECTION 10: Warranty_Return Orphan Check
'===========================================
Private Function CheckWarrantyOrphans(ws As Worksheet, rowNum As Long, _
                                      ByRef issueCount As Long) As Long
    Dim wsWR As Worksheet, wsJP As Worksheet
    Dim lastRowWR As Long, lastRowJP As Long
    Dim i As Long, j As Long
    Dim wrEntryID As String
    Dim found As Boolean
    Dim orphanCount As Long
    Dim r As Long

    r = rowNum
    orphanCount = 0

    On Error Resume Next
    Set wsWR = ThisWorkbook.Sheets("Warranty_Return_Master")
    Set wsJP = ThisWorkbook.Sheets("Job_Product")
    If Err.Number <> 0 Then
        issueCount = issueCount + 1
        CheckWarrantyOrphans = AddResult(ws, r, "? ERROR", _
            "Warranty Orphan Check", "Required sheets missing!", _
            "Both Warranty_Return_Master and Job_Product must exist.")
        Exit Function
    End If
    On Error GoTo 0

    lastRowWR = wsWR.Cells(wsWR.Rows.count, 1).End(xlUp).row
    lastRowJP = wsJP.Cells(wsJP.Rows.count, 1).End(xlUp).row

    For i = 2 To lastRowWR
        wrEntryID = Trim(CStr(wsWR.Cells(i, 1).value))
        If wrEntryID <> "" Then
            found = False
            For j = 2 To lastRowJP
                If Trim(CStr(wsJP.Cells(j, 1).value)) = wrEntryID Then
                    found = True
                    Exit For
                End If
            Next j
            If Not found Then
                orphanCount = orphanCount + 1
                If orphanCount <= 10 Then
                    r = AddResult(ws, r, "? ERROR", "Warranty Orphan", _
                        "Entry_ID: " & wrEntryID & " in Warranty_Return_Master (Row " & i & ")", _
                        "Add Entry_ID to Job_Product OR delete this orphan warranty record.")
                End If
            End If
        End If
    Next i

    If orphanCount = 0 Then
        r = AddResult(ws, r, "? PASS", "Warranty Orphans", _
            "All warranty records are linked to valid jobs.", _
            "No action needed.")
    Else
        issueCount = issueCount + orphanCount
        If orphanCount > 10 Then
            r = AddResult(ws, r, "?? INFO", "More Orphans", _
                "... and " & (orphanCount - 10) & " more orphan warranty records.", _
                "Use Find feature to locate all orphan Entry_IDs in Warranty_Return_Master.")
        End If
    End If

    CheckWarrantyOrphans = r
End Function

'===========================================
' SECTION 11: Job_Accessory Orphan Check
'===========================================
Private Function CheckJobAccessoryOrphans(ws As Worksheet, rowNum As Long, _
                                          ByRef issueCount As Long) As Long
    Dim wsJA As Worksheet, wsJP As Worksheet
    Dim lastRowJA As Long, lastRowJP As Long
    Dim i As Long, j As Long
    Dim jaEntryID As String
    Dim found As Boolean
    Dim orphanCount As Long
    Dim r As Long

    r = rowNum
    orphanCount = 0

    On Error Resume Next
    Set wsJA = ThisWorkbook.Sheets("Job_Accessory")
    Set wsJP = ThisWorkbook.Sheets("Job_Product")
    If Err.Number <> 0 Then
        issueCount = issueCount + 1
        CheckJobAccessoryOrphans = AddResult(ws, r, "? ERROR", _
            "Job_Accessory Orphan Check", "Required sheets missing!", _
            "Both Job_Accessory and Job_Product must exist.")
        Exit Function
    End If
    On Error GoTo 0

    lastRowJA = wsJA.Cells(wsJA.Rows.count, 1).End(xlUp).row
    lastRowJP = wsJP.Cells(wsJP.Rows.count, 1).End(xlUp).row

    For i = 2 To lastRowJA
        jaEntryID = Trim(CStr(wsJA.Cells(i, 1).value))
        If jaEntryID <> "" Then
            found = False
            For j = 2 To lastRowJP
                If Trim(CStr(wsJP.Cells(j, 1).value)) = jaEntryID Then
                    found = True
                    Exit For
                End If
            Next j
            If Not found Then
                orphanCount = orphanCount + 1
                If orphanCount <= 10 Then
                    r = AddResult(ws, r, "? ERROR", "Job_Accessory Orphan", _
                        "Entry_ID: " & jaEntryID & " in Job_Accessory (Row " & i & ")", _
                        "Add Entry_ID to Job_Product OR delete this orphan accessory record.")
                End If
            End If
        End If
    Next i

    If orphanCount = 0 Then
        r = AddResult(ws, r, "? PASS", "Job_Accessory Orphans", _
            "All job accessory records are linked to valid jobs.", _
            "No action needed.")
    Else
        issueCount = issueCount + orphanCount
        If orphanCount > 10 Then
            r = AddResult(ws, r, "?? INFO", "More Orphans", _
                "... and " & (orphanCount - 10) & " more orphan job accessory records.", _
                "Use Find feature to locate all orphan Entry_IDs in Job_Accessory.")
        End If
    End If

    CheckJobAccessoryOrphans = r
End Function

'===========================================
' SECTION 12: Assign_Accessories Orphan Check
'===========================================
Private Function CheckAssignAccessoryOrphans(ws As Worksheet, rowNum As Long, _
                                             ByRef issueCount As Long) As Long
    Dim wsAA As Worksheet, wsAM As Worksheet
    Dim lastRowAA As Long, lastRowAM As Long
    Dim i As Long, j As Long
    Dim aaAssignID As String
    Dim found As Boolean
    Dim orphanCount As Long
    Dim r As Long

    r = rowNum
    orphanCount = 0

    On Error Resume Next
    Set wsAA = ThisWorkbook.Sheets("Assign_Accessories")
    Set wsAM = ThisWorkbook.Sheets("Assign_Master")
    If Err.Number <> 0 Then
        issueCount = issueCount + 1
        CheckAssignAccessoryOrphans = AddResult(ws, r, "? ERROR", _
            "Assign_Accessories Orphan Check", "Required sheets missing!", _
            "Both Assign_Accessories and Assign_Master must exist.")
        Exit Function
    End If
    On Error GoTo 0

    lastRowAA = wsAA.Cells(wsAA.Rows.count, 1).End(xlUp).row
    lastRowAM = wsAM.Cells(wsAM.Rows.count, 1).End(xlUp).row

    For i = 2 To lastRowAA
        aaAssignID = Trim(CStr(wsAA.Cells(i, 1).value))
        If aaAssignID <> "" Then
            found = False
            For j = 2 To lastRowAM
                If Trim(CStr(wsAM.Cells(j, 1).value)) = aaAssignID Then
                    found = True
                    Exit For
                End If
            Next j
            If Not found Then
                orphanCount = orphanCount + 1
                If orphanCount <= 10 Then
                    r = AddResult(ws, r, "? ERROR", "Assign_Accessory Orphan", _
                        "Assign_ID: " & aaAssignID & " in Assign_Accessories (Row " & i & ")", _
                        "Add Assign_ID to Assign_Master OR delete this orphan accessory record.")
                End If
            End If
        End If
    Next i

    If orphanCount = 0 Then
        r = AddResult(ws, r, "? PASS", "Assign_Accessory Orphans", _
            "All assign accessory records are linked to valid assignments.", _
            "No action needed.")
    Else
        issueCount = issueCount + orphanCount
        If orphanCount > 10 Then
            r = AddResult(ws, r, "?? INFO", "More Orphans", _
                "... and " & (orphanCount - 10) & " more orphan assign accessory records.", _
                "Use Find feature to locate all orphan Assign_IDs in Assign_Accessories.")
        End If
    End If

    CheckAssignAccessoryOrphans = r
End Function

'===========================================
' SECTION 13: Empty Critical Fields
'===========================================
Private Function CheckEmptyCriticalFields(ws As Worksheet, rowNum As Long, _
                                          ByRef issueCount As Long) As Long
    Dim wsJP As Worksheet, wsAM As Worksheet
    Dim wsSW As Worksheet, wsDM As Worksheet
    Dim wsPM As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim blankCount As Long
    Dim r As Long
    Dim sheetName As String

    r = rowNum

    ' Check Job_Product for blank critical fields
    On Error Resume Next
    Set wsJP = ThisWorkbook.Sheets("Job_Product")
    If Err.Number = 0 Then
        lastRow = wsJP.Cells(wsJP.Rows.count, 1).End(xlUp).row
        blankCount = 0
        For i = 2 To lastRow
            If Trim(wsJP.Cells(i, 1).value) = "" Or _
               Trim(wsJP.Cells(i, 2).value) = "" Or _
               Trim(wsJP.Cells(i, 3).value) = "" Then
                blankCount = blankCount + 1
            End If
        Next i
        If blankCount > 0 Then
            issueCount = issueCount + blankCount
            r = AddResult(ws, r, "? ERROR", "Job_Product Blank Fields", _
                blankCount & " rows have one or more blank critical fields (Entry_ID/Status/Customer_ID).", _
                "Fill all blank fields. Critical fields cannot be empty.")
        Else
            r = AddResult(ws, r, "? PASS", "Job_Product Fields", _
                "No blank critical fields found.", _
                "No action needed.")
        End If
    End If
    On Error GoTo 0

    ' Check Assign_Master
    On Error Resume Next
    Set wsAM = ThisWorkbook.Sheets("Assign_Master")
    If Err.Number = 0 Then
        lastRow = wsAM.Cells(wsAM.Rows.count, 1).End(xlUp).row
        blankCount = 0
        For i = 2 To lastRow
            If Trim(wsAM.Cells(i, 1).value) = "" Or _
               Trim(wsAM.Cells(i, 2).value) = "" Then
                blankCount = blankCount + 1
            End If
        Next i
        If blankCount > 0 Then
            issueCount = issueCount + blankCount
            r = AddResult(ws, r, "? ERROR", "Assign_Master Blank Fields", _
                blankCount & " rows have blank Assign_ID or Entry_ID.", _
                "Fill all blank fields.")
        Else
            r = AddResult(ws, r, "? PASS", "Assign_Master Fields", _
                "No blank critical fields found.", _
                "No action needed.")
        End If
    End If
    On Error GoTo 0

    ' Check Service_Worklog_Expense
    On Error Resume Next
    Set wsSW = ThisWorkbook.Sheets("Service_Worklog_Expense")
    If Err.Number = 0 Then
        lastRow = wsSW.Cells(wsSW.Rows.count, 1).End(xlUp).row
        blankCount = 0
        For i = 2 To lastRow
            If Trim(wsSW.Cells(i, 1).value) = "" Then blankCount = blankCount + 1
        Next i
        If blankCount > 0 Then
            issueCount = issueCount + blankCount
            r = AddResult(ws, r, "? ERROR", "Worklog Blank Entry_ID", _
                blankCount & " rows have blank Entry_ID.", _
                "Fill Entry_ID for all worklog records.")
        Else
            r = AddResult(ws, r, "? PASS", "Worklog Fields", _
                "No blank Entry_IDs found.", _
                "No action needed.")
        End If
    End If
    On Error GoTo 0

    ' Check Delivery_Master
    On Error Resume Next
    Set wsDM = ThisWorkbook.Sheets("Delivery_Master")
    If Err.Number = 0 Then
        lastRow = wsDM.Cells(wsDM.Rows.count, 1).End(xlUp).row
        blankCount = 0
        For i = 2 To lastRow
            If Trim(wsDM.Cells(i, 1).value) = "" Then blankCount = blankCount + 1
        Next i
        If blankCount > 0 Then
            issueCount = issueCount + blankCount
            r = AddResult(ws, r, "? ERROR", "Delivery Blank Entry_ID", _
                blankCount & " rows have blank Entry_ID.", _
                "Fill Entry_ID for all delivery records.")
        Else
            r = AddResult(ws, r, "? PASS", "Delivery Fields", _
                "No blank Entry_IDs found.", _
                "No action needed.")
        End If
    End If
    On Error GoTo 0

    ' Check Payment_Master
    On Error Resume Next
    Set wsPM = ThisWorkbook.Sheets("Payment_Master")
    If Err.Number = 0 Then
        lastRow = wsPM.Cells(wsPM.Rows.count, 1).End(xlUp).row
        blankCount = 0
        For i = 2 To lastRow
            If Trim(wsPM.Cells(i, 1).value) = "" Then blankCount = blankCount + 1
        Next i
        If blankCount > 0 Then
            issueCount = issueCount + blankCount
            r = AddResult(ws, r, "? ERROR", "Payment Blank Entry_ID", _
                blankCount & " rows have blank Entry_ID.", _
                "Fill Entry_ID for all payment records.")
        Else
            r = AddResult(ws, r, "? PASS", "Payment Fields", _
                "No blank Entry_IDs found.", _
                "No action needed.")
        End If
    End If
    On Error GoTo 0

    CheckEmptyCriticalFields = r
End Function

'===========================================
' SECTION 14: Summary & Recommendations
'===========================================
Private Function GenerateSummary(ws As Worksheet, rowNum As Long, _
                                 totalIssues As Long) As Long
    Dim r As Long
    r = rowNum

    If totalIssues = 0 Then
        r = AddResult(ws, r, "? PASS", "OVERALL STATUS", _
            "NO ISSUES FOUND! Your file is clean and healthy.", _
            "Keep up the good data management practices!")
        r = AddResult(ws, r, "?? INFO", "Recommendation", _
            "Run this diagnostic weekly to catch issues early.", _
            "Set a reminder to run this tool every Monday morning.")
    ElseIf totalIssues < 10 Then
        r = AddResult(ws, r, "?? WARNING", "OVERALL STATUS", _
            "Minor issues found (" & totalIssues & " issues).", _
            "Fix issues one by one. Start with ERROR items, then WARNING items.")
        r = AddResult(ws, r, "?? INFO", "Tip", _
            "Use Ctrl+F to find specific Entry_IDs mentioned in the report.", _
            "This will help you locate problem records quickly.")
    ElseIf totalIssues < 50 Then
        r = AddResult(ws, r, "?? WARNING", "OVERALL STATUS", _
            "Moderate issues found (" & totalIssues & " issues).", _
            "Prioritize fixing orphan records and duplicate Entry_IDs first.")
        r = AddResult(ws, r, "?? INFO", "Tip", _
            "Consider creating a backup before making bulk corrections.", _
            "File > Save As > Create a backup copy before fixing.")
    Else
        r = AddResult(ws, r, "? ERROR", "OVERALL STATUS", _
            "CRITICAL: Major issues found (" & totalIssues & " issues)!", _
            "URGENT: Fix immediately. Data integrity is severely compromised.")
        r = AddResult(ws, r, "?? INFO", "Emergency Tip", _
            "Consider restoring from last known good backup if available.", _
            "If no backup exists, fix issues systematically starting with missing sheets and blank primary keys.")
    End If

    r = AddResult(ws, r, "?? INFO", "Quick Fix Order", _
        "1. Missing Sheets ? 2. Blank Primary Keys ? 3. Duplicates ? 4. Orphans ? 5. Status Mismatch", _
        "Follow this order for efficient problem resolution.")

    GenerateSummary = r
End Function

'===========================================
' UTILITY: Check if value exists in array
'===========================================
Private Function IsInArray(val As String, arr As Variant) As Boolean
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        If LCase(Trim(val)) = LCase(Trim(arr(i))) Then
            IsInArray = True
            Exit Function
        End If
    Next i
    IsInArray = False
End Function


