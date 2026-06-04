VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmServiceWorkExpense 
   Caption         =   "Service Work Log & Expense Entry"
   ClientHeight    =   13245
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   18165
   OleObjectBlob   =   "frmServiceWorkExpense.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmServiceWorkExpense"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
    ByVal hwnd As LongPtr, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, _
    ByVal lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As LongPtr

'===========================================
' MODULE VARIABLES
'===========================================
Private IsWorkLogSaved As Boolean
Private CurrentAssignID As String
Private CurrentCustomerID As String
Private CurrentCustomerAddress As String

'=== EDIT MODE VARIABLES ===
Private isEditMode As Boolean
Private EditLogRow As Long
Private IsLoadingData As Boolean




Private Sub lblCustomerValue_Click()

End Sub

Private Sub lblMobileValue_Click()

End Sub

'===========================================
' FORM INITIALIZE
'===========================================
Private Sub UserForm_Initialize()
    Me.caption = "Service Work Log & Expense Entry"
    Me.Width = 920
    Me.Height = 691
    Me.StartUpPosition = 1
    
    AddMinMaxButtons Me
    
    ' Reset variables
    IsWorkLogSaved = False
    CurrentAssignID = ""
    CurrentCustomerID = ""
    CurrentCustomerAddress = ""
    isEditMode = False
    EditLogRow = 0
    IsLoadingData = False
    
    ' Setup Work Status dropdown
    cmbWorkStatus.Clear
    cmbWorkStatus.AddItem "IN PROGRESS"
    cmbWorkStatus.AddItem "COMPLETED"
    cmbWorkStatus.AddItem "PENDING PARTS"
    cmbWorkStatus.AddItem "CANCELLED"
    
    ' ===== PAYMENT LISTVIEW SETUP =====
    On Error Resume Next
    With lstPayments
        .View = lvwReport
        .FullRowSelect = True
        .Gridlines = True
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "Payment ID", 80
        .ColumnHeaders.Add , , "Entry ID", 80
        .ColumnHeaders.Add , , "Type", 80
        .ColumnHeaders.Add , , "Category", 100
        .ColumnHeaders.Add , , "PayMode", 80
        .ColumnHeaders.Add , , "Name", 120
        .ColumnHeaders.Add , , "Company", 100
        .ColumnHeaders.Add , , "Model", 100
        .ColumnHeaders.Add , , "Serial", 100
        .ColumnHeaders.Add , , "Qty", 50
        .ColumnHeaders.Add , , "Amount", 80
        .ColumnHeaders.Add , , "Warranty", 80
    End With
    On Error GoTo 0
    
    Call UpdatePaymentButtons
    
    ' ===== CLEAR ACCESSORIES CHECKBOXES =====
    Dim chk As Control
    For Each chk In Me.Controls
        If TypeName(chk) = "CheckBox" Then
            If chk.name <> "chkCustomerApproved" Then
                chk.value = False
                chk.caption = ""
                chk.enabled = False
                chk.visible = False
            End If
        End If
    Next chk
    
    ' Load Assigned Entries
    LoadAssignedEntries
    
    ' ===== CHECK IF EDIT MODE (FROM DASHBOARD) =====
    Dim dbEntryID As String
    dbEntryID = Trim(Me.Tag)
    
    If dbEntryID <> "" Then
        ' ===== EDIT MODE =====
        isEditMode = True
        
        ' Add to combo if not exists
        Dim found As Boolean, idx As Long
        found = False
        For idx = 0 To ComboBox1.ListCount - 1
            If UCase(Trim(ComboBox1.List(idx))) = UCase(dbEntryID) Then
                found = True
                Exit For
            End If
        Next idx
        If Not found Then ComboBox1.AddItem dbEntryID
        
        ' Set EntryID - triggers ComboBox1_Change
        ComboBox1.value = dbEntryID
        
        Me.caption = "SERVICE WORK LOG - EDIT (" & dbEntryID & ")"
        btnSave.caption = "UPDATE WORK LOG"
        IsWorkLogSaved = True
        btnUpdateToCustomer.enabled = True
        Me.Tag = ""
    Else
        ' ===== NEW MODE =====
        isEditMode = False
        EditLogRow = 0
        btnSave.caption = "SAVE WORK LOG"
        IsWorkLogSaved = False
        btnUpdateToCustomer.enabled = False
        ClearForm
    End If
End Sub

'===========================================
' LOAD ASSIGNED ENTRIES (Only SERVICE - SER)
'===========================================
Private Sub LoadAssignedEntries()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Assign_Master")
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "Assign_Master sheet not found!", vbExclamation
        Exit Sub
    End If
    
    ComboBox1.Clear
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    
    For i = 2 To lastRow
        Dim status As String
        status = UCase(Trim(ws.Cells(i, 32).value))
        Dim entryID As String
        entryID = UCase(Trim(ws.Cells(i, 2).value))
        
        ' SIRF SERVICE (SER) entries jo ASSIGNED ya PENDING hain
        If (status = "ASSIGNED" Or status = "SENT" Or status = "IN PROGRESS" Or status = "PENDING" Or _
            status = "COMPLETED" Or status = "PENDING PARTS" Or status = "CANCELLED") And _
           Left(entryID, 3) = "SER" Then
            ComboBox1.AddItem ws.Cells(i, 2).value
        End If
    Next i
    
    If ComboBox1.ListCount = 0 Then
        MsgBox "No pending SERVICE entries found!" & vbCrLf & _
               "Please assign a SERVICE entry first.", vbInformation
    End If
End Sub

'===========================================
' COMBOBOX1 CHANGE - LOAD ASSIGNMENT + WORK LOG
'===========================================
Private Sub ComboBox1_Change()
    Dim wsAssign As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim entryID As String
    Dim custID As String
    
    entryID = Trim(ComboBox1.value)
    If entryID = "" Then Exit Sub
    
    On Error Resume Next
    Set wsAssign = ThisWorkbook.Sheets("Assign_Master")
    On Error GoTo 0
    
    If wsAssign Is Nothing Then Exit Sub
    
    lastRow = wsAssign.Cells(wsAssign.Rows.count, 1).End(xlUp).row
    
    For i = 2 To lastRow
        If UCase(Trim(wsAssign.Cells(i, 2).value)) = UCase(entryID) Then
            
            ' Store Assign ID
            CurrentAssignID = wsAssign.Cells(i, 1).value
            custID = wsAssign.Cells(i, 3).value
            CurrentCustomerID = custID
            CurrentCustomerAddress = GetCustomerAddress(custID)
            
            ' Load Assignment Details to labels
            lblEntryIDValue.caption = wsAssign.Cells(i, 2).value
            lblProductValue.caption = wsAssign.Cells(i, 6).value
            lblCompanyValue.caption = wsAssign.Cells(i, 7).value
            lblModelValue.caption = wsAssign.Cells(i, 8).value
            lblSerialValue.caption = wsAssign.Cells(i, 9).value
            lblCustomerValue.caption = wsAssign.Cells(i, 5).value
            lblMobileValue.caption = wsAssign.Cells(i, 4).value
            lblAssignedToValue.caption = wsAssign.Cells(i, 14).value
            lblAssignedDateValue.caption = wsAssign.Cells(i, 22).value
            
            ' Load existing problem if any
            If Trim(wsAssign.Cells(i, 29).value) <> "" Then
                txtProblemFound.value = wsAssign.Cells(i, 29).value
            End If
            
            ' Load Customer Photo
            If custID <> "" Then
                LoadCustomerPhotoByID custID
            End If
            
            ' Load Product Photo
            LoadProductPhoto entryID
            
            Exit For
        End If
    Next i
    
    ' ===== LOAD EXISTING WORK LOG =====
    Call LoadServiceWorkLogData(entryID)
    
    ' ===== LOAD PAYMENT LIST =====
    Call LoadPaymentList(entryID)
    
    ' ===== LOAD ACCESSORIES =====
    Call LoadAccessories(entryID)
End Sub

'===========================================
' LOAD EXISTING WORK LOG DATA (EDIT MODE)
'===========================================
Private Sub LoadServiceWorkLogData(entryID As String)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Service_Worklog_Expense")
    On Error GoTo 0
    
    If ws Is Nothing Then Exit Sub
    
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    EditLogRow = 0
    
    ' Find LATEST work log by EntryID (Column 3)
    For i = lastRow To 2 Step -1
        If UCase(Trim(ws.Cells(i, 3).value & "")) = UCase(entryID) Then
            EditLogRow = i
            
            IsLoadingData = True
            
            ' Load Work Log Details
            txtStartDate.value = Trim(ws.Cells(i, 11).value & "")
            txtEndDate.value = Trim(ws.Cells(i, 12).value & "")
            cmbWorkStatus.value = Trim(ws.Cells(i, 13).value & "")
            txtProblemFound.value = Trim(ws.Cells(i, 14).value & "")
            txtSolutionApplied.value = Trim(ws.Cells(i, 15).value & "")
            txtRemarks.value = Trim(ws.Cells(i, 24).value & "")
            
            ' Customer Approval
            Dim approval As String
            approval = UCase(Trim(ws.Cells(i, 21).value & ""))
            If approval = "YES" Or approval = "TRUE" Then
                chkCustomerApproved.value = True
                txtApprovalDate.value = Trim(ws.Cells(i, 22).value & "")
            Else
                chkCustomerApproved.value = False
                txtApprovalDate.value = ""
            End If
            
            IsLoadingData = False
            
            IsWorkLogSaved = True
            btnUpdateToCustomer.enabled = True
            btnSave.caption = "UPDATE WORK LOG"
            
            Exit For
        End If
    Next i
    
    If EditLogRow = 0 Then
        ' No existing work log - new entry
        IsWorkLogSaved = False
        btnUpdateToCustomer.enabled = False
        btnSave.caption = "SAVE WORK LOG"
        
        ' Reset work fields only (keep assignment details)
        IsLoadingData = True
        cmbWorkStatus.value = "IN PROGRESS"
        txtStartDate.value = Format(Date, "dd-mm-yyyy")
        txtEndDate.value = ""
        txtProblemFound.value = ""
        txtSolutionApplied.value = ""
        txtRemarks.value = ""
        chkCustomerApproved.value = False
        txtApprovalDate.value = ""
        IsLoadingData = False
    End If
End Sub

'===========================================
' LOAD CUSTOMER PHOTO BY CUSTOMER ID
'===========================================
Private Sub LoadCustomerPhotoByID(custID As String)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim photoName As String
    Dim folderPath As String
    Dim fullPath As String
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Customer_Master")
    On Error GoTo 0
    
    If ws Is Nothing Then Exit Sub
    
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    photoName = ""
    
    For i = 2 To lastRow
        If UCase(Trim(ws.Cells(i, 1).value)) = UCase(custID) Then
            photoName = Trim(ws.Cells(i, 7).value & "")
            Exit For
        End If
    Next i
    
    If photoName = "" Then
        Set imgCustomer.Picture = Nothing
        Exit Sub
    End If
    
    folderPath = ThisWorkbook.path & "\GLOBAL_SOFT_DATA\Images\Customers\"
    If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
    fullPath = folderPath & photoName
    
    If Dir(fullPath) <> "" Then
        Set imgCustomer.Picture = LoadPicture(fullPath)
    Else
        Set imgCustomer.Picture = Nothing
    End If
End Sub

'===========================================
' LOAD PRODUCT PHOTO
'===========================================
Private Sub LoadProductPhoto(entryID As String)
    Dim photoPath As String
    Dim folderPath As String
    
    On Error Resume Next
    Set imgProduct.Picture = Nothing
    
    If Trim(entryID) = "" Then Exit Sub
    
    folderPath = ThisWorkbook.path & "\GLOBAL_SOFT_DATA\Images\Products\" & entryID & "\"
    photoPath = folderPath & "Photo1.jpg"
    
    If Dir(photoPath) = "" Then
        photoPath = ThisWorkbook.path & "\GLOBAL_SOFT_DATA\Images\Products\" & entryID & ".jpg"
    End If
    
    If Dir(photoPath) <> "" Then
        imgProduct.Picture = LoadPicture(photoPath)
    Else
        Set imgProduct.Picture = Nothing
    End If
    
    On Error GoTo 0
End Sub

'===========================================
' UPDATE TOTAL AMOUNT FROM PAYMENT MASTER
'===========================================
Private Sub UpdateTotalAmount()
    Dim ws As Worksheet
    Dim lastRow As Long, r As Long
    Dim total As Double
    
    total = 0
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Payment_Master")
    On Error GoTo 0
    
    If Not ws Is Nothing Then
        lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
        For r = 2 To lastRow
            If Trim(ws.Cells(r, 2).value & "") = Trim(ComboBox1.value & "") Then
                If ws.Cells(r, 6).value = "" Then
                    total = total + val(ws.Cells(r, 16).value)
                End If
            End If
        Next r
    End If
    
    On Error Resume Next
    lblTotalAmount.caption = "Total Amount: Rs. " & Format(total, "0.00")
    On Error GoTo 0
End Sub

'===========================================
' SAVE / UPDATE WORK LOG
'===========================================
Private Sub btnSave_Click()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim wsAssign As Worksheet
    Dim wsJob As Worksheet
    Dim lastRow As Long
    Dim workLogID As String
    Dim i As Integer
    Dim selectedStatus As String
    Dim custID As String
    Dim engineerID As String
    
    ' === GET CHARGES FROM PAYMENT MASTER ===
    Dim partsNames As String
    Dim laborCharges As Double
    Dim transportCharges As Double
    Dim otherCharges As Double
    Dim actualPartsCost As Double
    Dim totalExpense As Double
    
    ' === CHECK ACCESSORIES DUE ===
    Dim dueItems As String
    Dim chk As Control
    
    dueItems = ""
    For Each chk In Me.Controls
        If TypeName(chk) = "CheckBox" Then
            If chk.visible And chk.caption <> "" And chk.value = False Then
                dueItems = dueItems & "• " & chk.caption & vbCrLf
            End If
        End If
    Next chk
    
    If dueItems <> "" Then
        If MsgBox("WARNING! These accessories are DUE from engineer:" & vbCrLf & vbCrLf & _
                  dueItems & vbCrLf & _
                  "Do you still want to SAVE work log?", vbExclamation + vbYesNo, "ACCESSORIES DUE - GLOBAL SOFT") = vbNo Then
            Exit Sub
        End If
    End If
    
    Call GetPaymentSummary(ComboBox1.value, partsNames, laborCharges, transportCharges, otherCharges, actualPartsCost, totalExpense)
    
    ' === VALIDATION ===
    If ComboBox1.value = "" Then
        MsgBox "Please select an Assigned Entry!", vbExclamation
        ComboBox1.SetFocus
        Exit Sub
    End If
    
    If cmbWorkStatus.value = "" Then
        MsgBox "Please select Work Status!", vbExclamation
        cmbWorkStatus.SetFocus
        Exit Sub
    End If
    
    If Trim(txtProblemFound.value) = "" Then
        MsgBox "Please enter Problem Found!", vbExclamation
        txtProblemFound.SetFocus
        Exit Sub
    End If
    
    If Trim(txtSolutionApplied.value) = "" Then
        MsgBox "Please enter Solution Applied!", vbExclamation
        txtSolutionApplied.SetFocus
        Exit Sub
    End If
    
    selectedStatus = UCase(Trim(cmbWorkStatus.value))
    
    ' === GET CUSTOMER ID & ENGINEER ID ===
    custID = ""
    engineerID = ""
    On Error Resume Next
    Set wsAssign = ThisWorkbook.Sheets("Assign_Master")
    On Error GoTo ErrorHandler
    
    If Not wsAssign Is Nothing Then
        lastRow = wsAssign.Cells(wsAssign.Rows.count, 1).End(xlUp).row
        For i = 2 To lastRow
            If UCase(Trim(wsAssign.Cells(i, 2).value)) = UCase(ComboBox1.value) Then
                custID = Trim(wsAssign.Cells(i, 3).value)
                engineerID = Trim(wsAssign.Cells(i, 13).value)
                Exit For
            End If
        Next i
    End If
    
    ' === OPEN / CREATE Service_Worklog_Expense ===
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Service_Worklog_Expense")
    On Error GoTo ErrorHandler
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add
        ws.name = "Service_Worklog_Expense"
        
        ws.Cells(1, 1).value = "LogID"
        ws.Cells(1, 2).value = "AssignID"
        ws.Cells(1, 3).value = "EntryID"
        ws.Cells(1, 4).value = "CustomerID"
        ws.Cells(1, 5).value = "ProductType"
        ws.Cells(1, 6).value = "Company"
        ws.Cells(1, 7).value = "Model"
        ws.Cells(1, 8).value = "SerialNumber"
        ws.Cells(1, 9).value = "EngineerID"
        ws.Cells(1, 10).value = "EngineerName"
        ws.Cells(1, 11).value = "WorkStartDate"
        ws.Cells(1, 12).value = "WorkEndDate"
        ws.Cells(1, 13).value = "WorkStatus"
        ws.Cells(1, 14).value = "ProblemFound"
        ws.Cells(1, 15).value = "ProblemSolution"
        ws.Cells(1, 16).value = "PartsUsed"
        ws.Cells(1, 17).value = "LaborCharges"
        ws.Cells(1, 18).value = "TransportCharges"
        ws.Cells(1, 19).value = "OtherCharges"
        ws.Cells(1, 20).value = "TotalExpense"
        ws.Cells(1, 21).value = "CustomerApproval"
        ws.Cells(1, 22).value = "ApprovalDate"
        ws.Cells(1, 23).value = "ExpensePhotoPath"
        ws.Cells(1, 24).value = "Remarks"
        ws.Cells(1, 25).value = "CreatedDate"
        ws.Cells(1, 26).value = "EditHistory"
        ws.Rows(1).Font.Bold = True
        ws.Rows(1).Interior.Color = RGB(200, 200, 200)
    End If
    
    ' ===== EDIT MODE = UPDATE EXISTING =====
    If EditLogRow > 0 Then
        With ws.Rows(EditLogRow)
            .Cells(1, 11).value = txtStartDate.value
            .Cells(1, 12).value = txtEndDate.value
            .Cells(1, 13).value = selectedStatus
            .Cells(1, 14).value = txtProblemFound.value
            .Cells(1, 15).value = txtSolutionApplied.value
            .Cells(1, 16).value = partsNames
            .Cells(1, 17).value = laborCharges
            .Cells(1, 18).value = transportCharges
            .Cells(1, 19).value = otherCharges
            .Cells(1, 20).value = totalExpense
            .Cells(1, 21).value = IIf(chkCustomerApproved.value = True, "YES", "NO")
            .Cells(1, 22).value = txtApprovalDate.value
            .Cells(1, 24).value = txtRemarks.value
            .Cells(1, 26).value = Format(Now, "dd-MM-yyyy hh:mm:ss") & " EDITED by " & Environ("Username")
        End With
        
        workLogID = ws.Cells(EditLogRow, 1).value
    Else
        ' ===== NEW MODE = INSERT =====
        lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row + 1
        workLogID = "WLG" & Format(lastRow - 1, "00000")
        
        With ws.Rows(lastRow)
            .Cells(1, 1).value = workLogID
            .Cells(1, 2).value = CurrentAssignID
            .Cells(1, 3).value = ComboBox1.value
            .Cells(1, 4).value = custID
            .Cells(1, 5).value = lblProductValue.caption
            .Cells(1, 6).value = lblCompanyValue.caption
            .Cells(1, 7).value = lblModelValue.caption
            .Cells(1, 8).value = lblSerialValue.caption
            .Cells(1, 9).value = engineerID
            .Cells(1, 10).value = lblAssignedToValue.caption
            .Cells(1, 11).value = txtStartDate.value
            .Cells(1, 12).value = txtEndDate.value
            .Cells(1, 13).value = selectedStatus
            .Cells(1, 14).value = txtProblemFound.value
            .Cells(1, 15).value = txtSolutionApplied.value
            .Cells(1, 16).value = partsNames
            .Cells(1, 17).value = laborCharges
            .Cells(1, 18).value = transportCharges
            .Cells(1, 19).value = otherCharges
            .Cells(1, 20).value = totalExpense
            .Cells(1, 21).value = IIf(chkCustomerApproved.value = True, "YES", "NO")
            .Cells(1, 22).value = txtApprovalDate.value
            .Cells(1, 23).value = ""
            .Cells(1, 24).value = txtRemarks.value
            .Cells(1, 25).value = Format(Now, "dd-MM-yyyy hh:mm:ss")
            .Cells(1, 26).value = GetAccessoriesStatus()
        End With
        
        EditLogRow = lastRow
    End If
    
    ' === SYNC STATUS TO Assign_Master ===
If Not wsAssign Is Nothing Then
    lastRow = wsAssign.Cells(wsAssign.Rows.count, 1).End(xlUp).row
    For i = 2 To lastRow
        If UCase(Trim(wsAssign.Cells(i, 2).value)) = UCase(ComboBox1.value) Then
            wsAssign.Cells(i, 32).value = selectedStatus
            ' Date Update for ALL status changes
            wsAssign.Cells(i, 35).value = Format(Now, "dd-MM-yyyy hh:mm:ss")
            Exit For
        End If
    Next i
End If
    
        ' === SYNC STATUS TO Job_Product ===
    Set wsJob = ThisWorkbook.Sheets("Job_Product")
    lastRow = wsJob.Cells(wsJob.Rows.count, 1).End(xlUp).row
    For i = 2 To lastRow
        If UCase(Trim(wsJob.Cells(i, 1).value)) = UCase(ComboBox1.value) Then
            Select Case selectedStatus
                Case "COMPLETED": wsJob.Cells(i, 4).value = "COMPLETED"
                Case "IN PROGRESS": wsJob.Cells(i, 4).value = "IN PROGRESS"
                Case "PENDING PARTS": wsJob.Cells(i, 4).value = "PENDING PARTS"
                Case "CANCELLED": wsJob.Cells(i, 4).value = "CANCELLED"
            End Select
            wsJob.Cells(i, 5).value = Format(Now, "dd-MM-yyyy hh:mm")   '? Date Update
            Exit For
        End If
    Next i
    
    IsWorkLogSaved = True
    btnUpdateToCustomer.enabled = True
    btnSave.caption = "UPDATE WORK LOG"
    btnSave.enabled = False
    
    MsgBox "Work Log Saved Successfully!" & vbCrLf & "Log ID: " & workLogID, vbInformation
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Save Error: " & Err.Number & " - " & Err.Description, vbCritical
End Sub

'===========================================
' CANCEL BUTTON
'===========================================
Private Sub btnCancel_Click()
    Unload Me
End Sub

'===========================================
' CLEAR FORM
'===========================================
Private Sub ClearForm()
    ComboBox1.value = ""
    
    ' Assignment Details clear
    lblEntryIDValue.caption = ""
    lblProductValue.caption = ""
    lblCompanyValue.caption = ""
    lblModelValue.caption = ""
    lblSerialValue.caption = ""
    lblCustomerValue.caption = ""
    lblMobileValue.caption = ""
    lblAssignedToValue.caption = ""
    lblAssignedDateValue.caption = ""
    
    ' Work Status
    IsLoadingData = True
    cmbWorkStatus.value = "IN PROGRESS"
    txtStartDate.value = Format(Date, "dd-mm-yyyy")
    txtEndDate.value = ""
    IsLoadingData = False
    
    ' Problem & Solution
    txtProblemFound.value = ""
    txtSolutionApplied.value = ""
    
    ' Customer Approval
    chkCustomerApproved.value = False
    txtApprovalDate.value = ""
    
    ' Remarks
    txtRemarks.value = ""
    
    ' Photos
    On Error Resume Next
    Set imgCustomer.Picture = Nothing
    Set imgProduct.Picture = Nothing
    On Error GoTo 0
    
    ' Variables
    CurrentAssignID = ""
    CurrentCustomerID = ""
    CurrentCustomerAddress = ""
    IsWorkLogSaved = False
    isEditMode = False
    EditLogRow = 0
    
    ' Accessories
    Dim chk As Control
    For Each chk In Me.Controls
        If TypeName(chk) = "CheckBox" Then
            If chk.name <> "chkCustomerApproved" Then
                chk.value = False
                chk.caption = ""
                chk.enabled = False
                chk.visible = False
            End If
        End If
    Next chk
    
    ' Buttons
    btnUpdateToCustomer.enabled = False
    btnSave.enabled = True
    btnSave.caption = "SAVE WORK LOG"
    
    ' Payment list
    lstPayments.ListItems.Clear
    Call UpdatePaymentButtons
    
    ComboBox1.SetFocus
End Sub

'===========================================
' PRINT WORK LOG
'===========================================
Private Sub btnPrint_Click()
    MsgBox "Print function - implement as needed", vbInformation
End Sub

'===========================================
' DATE VALIDATION
'===========================================
Private Sub txtStartDate_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    On Error Resume Next
    If txtStartDate.value <> "" Then
        txtStartDate.value = Format(CDate(txtStartDate.value), "dd-mm-yyyy")
    End If
    On Error GoTo 0
End Sub

Private Sub txtEndDate_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    On Error Resume Next
    If txtEndDate.value <> "" Then
        txtEndDate.value = Format(CDate(txtEndDate.value), "dd-mm-yyyy")
    End If
    On Error GoTo 0
End Sub

Private Sub txtApprovalDate_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    On Error Resume Next
    If txtApprovalDate.value <> "" Then
        txtApprovalDate.value = Format(CDate(txtApprovalDate.value), "dd-mm-yyyy")
    End If
    On Error GoTo 0
End Sub

'===========================================
' WORK STATUS CHANGE
'===========================================
Private Sub cmbWorkStatus_Change()
    If IsLoadingData Then Exit Sub
    
    Dim status As String
    status = UCase(Trim(cmbWorkStatus.value))
    
    Select Case status
        Case "IN PROGRESS"
            If txtStartDate.value = "" Then txtStartDate.value = Format(Date, "dd-mm-yyyy")
            txtEndDate.value = ""
            
        Case "COMPLETED"
            If txtStartDate.value = "" Then txtStartDate.value = Format(Date, "dd-mm-yyyy")
            txtEndDate.value = Format(Date, "dd-mm-yyyy")
            
        Case "PENDING PARTS", "CANCELLED"
            If txtStartDate.value = "" Then txtStartDate.value = Format(Date, "dd-mm-yyyy")
            txtEndDate.value = ""
    End Select
End Sub

'===========================================
' CUSTOMER APPROVED CHECKBOX
'===========================================
Private Sub chkCustomerApproved_Click()
    If IsLoadingData Then Exit Sub
    
    If chkCustomerApproved.value = True Then
        txtApprovalDate.value = Format(Date, "dd-mm-yyyy")
    Else
        txtApprovalDate.value = ""
    End If
End Sub

'===========================================
' CLEAR BUTTON
'===========================================
Private Sub btnClear_Click()
    If MsgBox("Are you sure you want to clear the form?" & vbCrLf & "All unsaved data will be lost!", vbQuestion + vbYesNo, "GLOBAL SOFT") = vbYes Then
        ClearForm
        LoadAssignedEntries
    End If
End Sub

'===========================================
' UPDATE TO CUSTOMER - WHATSAPP
'===========================================
Private Sub btnUpdateToCustomer_Click()
    On Error GoTo ErrorHandler
    
    Dim customerMobile As String
    Dim custName As String
    Dim messageText As String
    Dim waURL As String
    Dim companyName As String
    Dim companyAddress As String
    Dim companyPhone As String
    Dim statusText As String
    Dim i As Integer
    
    If ComboBox1.value = "" Then
        MsgBox "Please select an Entry ID first!", vbExclamation, "GLOBAL SOFT"
        ComboBox1.SetFocus
        Exit Sub
    End If
    
    If chkCustomerApproved.value = False Then
        MsgBox "Customer approval is PENDING!" & vbCrLf & _
               "Please get customer approval before sending update.", vbExclamation, "GLOBAL SOFT"
        chkCustomerApproved.SetFocus
        Exit Sub
    End If
    
    If Not IsWorkLogSaved Then
        MsgBox "Please SAVE Work Log first!" & vbCrLf & "Then click UPDATE TO CUSTOMER.", vbExclamation, "GLOBAL SOFT"
        Exit Sub
    End If
    
    '--- GET COMPANY INFO ---
    companyName = "GLOBAL IT SOLUTIONS"
    companyAddress = "Main Road, In front of Aahar, Odisha, India"
    companyPhone = ""
    
    On Error Resume Next
    Dim wsConfig As Worksheet
    Set wsConfig = ThisWorkbook.Sheets("Software_Config")
    On Error GoTo 0
    
    If Not wsConfig Is Nothing Then
        If Trim(wsConfig.Range("B2").value & "") <> "" Then companyName = Trim(wsConfig.Range("B2").value & "")
        
        Dim addr1 As String, addr2 As String, city As String
        Dim pinCode As String, stateName As String, countryName As String
        Dim fullAddress As String
        
        addr1 = Trim(wsConfig.Range("B3").value & "")
        addr2 = Trim(wsConfig.Range("B4").value & "")
        city = Trim(wsConfig.Range("B5").value & "")
        pinCode = Trim(wsConfig.Range("B6").value & "")
        stateName = Trim(wsConfig.Range("B7").value & "")
        countryName = Trim(wsConfig.Range("B8").value & "")
        
        fullAddress = addr1
        If addr2 <> "" Then fullAddress = fullAddress & ", " & addr2
        If city <> "" Then fullAddress = fullAddress & ", " & city
        If pinCode <> "" Then fullAddress = fullAddress & " - " & pinCode
        If stateName <> "" Then fullAddress = fullAddress & ", " & stateName
        If countryName <> "" Then fullAddress = fullAddress & ", " & countryName
        If fullAddress <> "" Then companyAddress = fullAddress
        If Trim(wsConfig.Range("B9").value & "") <> "" Then companyPhone = Trim(wsConfig.Range("B9").value & "")
    End If
    
    '--- CUSTOMER MOBILE ---
    customerMobile = Trim(lblMobileValue.caption & "")
    If customerMobile = "" Then
        MsgBox "Customer mobile number not found!", vbExclamation, "GLOBAL SOFT"
        Exit Sub
    End If
    
    customerMobile = Replace(customerMobile, " ", "")
    customerMobile = Replace(customerMobile, "-", "")
    customerMobile = Replace(customerMobile, "(", "")
    customerMobile = Replace(customerMobile, ")", "")
    customerMobile = Replace(customerMobile, "+", "")
    
    If Len(customerMobile) = 10 Then customerMobile = "91" & customerMobile
    
    If Len(customerMobile) <> 12 Or Left(customerMobile, 2) <> "91" Then
        MsgBox "Invalid mobile number format!" & vbCrLf & "Found: " & customerMobile, vbExclamation, "GLOBAL SOFT"
        Exit Sub
    End If
    
    '--- CUSTOMER DETAILS ---
    custName = Trim(lblCustomerValue.caption & "")
    If custName = "" Then custName = "Customer"
    
    Dim custAddress As String
    custAddress = CurrentCustomerAddress
    If custAddress = "" Then custAddress = "N/A"
    
    Dim custEmail As String
    custEmail = GetCustomerEmail(CurrentCustomerID)
    If custEmail = "" Then custEmail = "N/A"
    
    '--- STATUS ---
    statusText = UCase(Trim(cmbWorkStatus.value & ""))
    If statusText = "" Then statusText = "IN PROGRESS"
    
    '--- READ CHARGES FROM PAYMENT MASTER ---
    Dim wsPay As Worksheet
    Dim lastRowPay As Long
    Dim discussionAmt As Double
    Dim partsCost As Double, chargesCost As Double
    Dim receivedAmt As Double, discountAmt As Double
    Dim totalDue As Double
    Dim partsListText As String
    
    discussionAmt = 0: partsCost = 0: chargesCost = 0
    receivedAmt = 0: discountAmt = 0: totalDue = 0
    partsListText = ""
    
    On Error Resume Next
    Set wsPay = ThisWorkbook.Sheets("Payment_Master")
    On Error GoTo 0
    
    If Not wsPay Is Nothing Then
        lastRowPay = wsPay.Cells(wsPay.Rows.count, 1).End(xlUp).row
        
        For i = 2 To lastRowPay
            If Trim(wsPay.Cells(i, 2).value & "") = Trim(ComboBox1.value & "") Then
                If Trim(wsPay.Cells(i, 6).value & "") = "" Then
                    
                    Dim pName As String, pCategory As String, pAmt As Double
                    
                    pName = Trim(wsPay.Cells(i, 10).value & "")
                    pCategory = UCase(Trim(wsPay.Cells(i, 9).value & ""))
                    pAmt = val(wsPay.Cells(i, 16).value)
                    
                    If pCategory = "DISCUSSION" Then
                        discussionAmt = discussionAmt + pAmt
                    End If
                    
                    If pCategory = "PART_USED" Or pCategory = "PART_REPLACED" Then
                        partsCost = partsCost + pAmt
                        If partsListText <> "" Then partsListText = partsListText & vbCrLf
                        partsListText = partsListText & "• " & pName & " = Rs." & Format(pAmt, "0.00")
                    End If
                    
                    If pCategory = "CHARGE" Or pCategory = "SERVICE_CHARGE" Or _
                       pCategory = "LABOR_CHARGE" Or pCategory = "TRANSPORT_CHARGE" Or _
                       pCategory = "EXPENSE" Or pCategory = "OTHER_CHARGE" Then
                        chargesCost = chargesCost + pAmt
                        If partsListText <> "" Then partsListText = partsListText & vbCrLf
                        partsListText = partsListText & "• " & pName & " = Rs." & Format(pAmt, "0.00") & " (" & pCategory & ")"
                    End If
                    
                    If pCategory = "RECEIVED" Or pCategory = "ADVANCE" Then
                        receivedAmt = receivedAmt + pAmt
                    End If
                    
                    If pCategory = "DISCOUNT" Then
                        discountAmt = discountAmt + pAmt
                    End If
                    
                End If
            End If
        Next i
    End If
    
    totalDue = (partsCost + chargesCost) - (receivedAmt + discountAmt)
    
    '--- BUILD MESSAGE ---
    messageText = "*" & companyName & "*" & vbCrLf
    If companyAddress <> "" Then messageText = messageText & companyAddress & vbCrLf
    If companyPhone <> "" Then messageText = messageText & "Phone: " & companyPhone & vbCrLf
    
    messageText = messageText & vbCrLf & "Dear " & custName & "," & vbCrLf & vbCrLf
    messageText = messageText & "Your product service status has been updated." & vbCrLf & vbCrLf
    
    messageText = messageText & "*Customer Details:*" & vbCrLf
    messageText = messageText & "Name: " & custName & vbCrLf
    messageText = messageText & "Mobile: " & lblMobileValue.caption & vbCrLf
    messageText = messageText & "Address: " & custAddress & vbCrLf
    messageText = messageText & "Email: " & custEmail & vbCrLf & vbCrLf
    
    messageText = messageText & "*Product Details:*" & vbCrLf
    messageText = messageText & "Product Type: " & lblProductValue.caption & vbCrLf
    messageText = messageText & "Company: " & lblCompanyValue.caption & vbCrLf
    messageText = messageText & "Model: " & lblModelValue.caption & vbCrLf
    messageText = messageText & "Serial No: " & lblSerialValue.caption & vbCrLf & vbCrLf
    
    messageText = messageText & "*Status:* " & statusText & vbCrLf & vbCrLf
    
    Dim problemFound As String, solutionApplied As String
    problemFound = Trim(txtProblemFound.value & "")
    solutionApplied = Trim(txtSolutionApplied.value & "")
    
    If problemFound <> "" Then
        messageText = messageText & "*Problem Found:*" & vbCrLf & problemFound & vbCrLf & vbCrLf
    End If
    If solutionApplied <> "" Then
        messageText = messageText & "*Solution Applied:*" & vbCrLf & solutionApplied & vbCrLf & vbCrLf
    End If
    
    If discussionAmt > 0 Then
        messageText = messageText & "*Discussion Amount:* Rs. " & Format(discussionAmt, "0.00") & vbCrLf
        messageText = messageText & "(Initial estimated charges before service)" & vbCrLf & vbCrLf
    End If
    
    If partsListText <> "" Then
        messageText = messageText & "*Parts Used & Charges:*" & vbCrLf
        messageText = messageText & partsListText & vbCrLf
        messageText = messageText & "------------------------" & vbCrLf
    End If
    
    If receivedAmt > 0 Or discountAmt > 0 Then
        messageText = messageText & vbCrLf & "*Payments & Discounts:*" & vbCrLf
        If receivedAmt > 0 Then
            messageText = messageText & "Amount Received: Rs. " & Format(receivedAmt, "0.00") & vbCrLf
        End If
        If discountAmt > 0 Then
            messageText = messageText & "Discount Given: -Rs. " & Format(discountAmt, "0.00") & vbCrLf
        End If
        messageText = messageText & "------------------------" & vbCrLf
    End If
    
    If totalDue > 0 Then
        messageText = messageText & "*TOTAL AMOUNT DUE: Rs. " & Format(totalDue, "0.00") & "*" & vbCrLf & vbCrLf
    ElseIf totalDue < 0 Then
        messageText = messageText & "*REFUND DUE TO CUSTOMER: Rs. " & Format(Abs(totalDue), "0.00") & "*" & vbCrLf & vbCrLf
    Else
        messageText = messageText & "*TOTAL AMOUNT DUE: Rs. 0.00 (ALL CLEAR)*" & vbCrLf & vbCrLf
    End If
    
    messageText = messageText & "*Service Note:*" & vbCrLf
    
    Dim actionText As String
    Select Case statusText
        Case "COMPLETED", "REPAIRED": actionText = "successfully repaired"
        Case "REPLACED": actionText = "successfully replaced"
        Case "PENDING PARTS": actionText = "pending for parts"
        Case "CANCELLED": actionText = "cancelled"
        Case Else: actionText = "in progress"
    End Select
    
    messageText = messageText & "Product has been " & actionText & " by your service engineer." & vbCrLf & vbCrLf
    
    If statusText = "COMPLETED" Or statusText = "REPAIRED" Or statusText = "REPLACED" Then
        messageText = messageText & "Please collect your product from our service center within *7 days*." & vbCrLf
        messageText = messageText & "After 7 days, if product is misplaced, damaged, or lost, we will not be responsible." & vbCrLf
        messageText = messageText & "Storage charges will also apply as per company policy." & vbCrLf & vbCrLf
    End If
    
    messageText = messageText & "Thank you for trusting *" & companyName & "*." & vbCrLf
    If companyPhone <> "" Then messageText = messageText & "For any query, contact us: " & companyPhone
    
    '--- URL ENCODE ---
    messageText = Replace(messageText, " ", "%20")
    messageText = Replace(messageText, vbCrLf, "%0A")
    messageText = Replace(messageText, "&", "%26")
    messageText = Replace(messageText, "#", "%23")
    messageText = Replace(messageText, "=", "%3D")
    messageText = Replace(messageText, "*", "%2A")
    messageText = Replace(messageText, "+", "%2B")
    
    '--- OPEN WHATSAPP ---
    waURL = "whatsapp://send?phone=" & customerMobile & "&text=" & messageText
    ShellExecute 0, "Open", waURL, "", "", 1
    
    MsgBox "WhatsApp opened successfully!" & vbCrLf & vbCrLf & _
           "Customer: " & custName & vbCrLf & _
           "Mobile: +" & customerMobile & vbCrLf & _
           "Status: " & statusText, vbInformation, "GLOBAL SOFT - WhatsApp Ready"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "WhatsApp Error: " & Err.Number & " - " & Err.Description, vbCritical, "GLOBAL SOFT"
End Sub

'===========================================
' GET CUSTOMER ADDRESS
'===========================================
Private Function GetCustomerAddress(custID As String) As String
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim cellVal As String
    
    GetCustomerAddress = ""
    If Trim(custID) = "" Then Exit Function
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Customer_Master")
    On Error GoTo 0
    If ws Is Nothing Then Exit Function
    
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    
    For i = 2 To lastRow
        If UCase(Trim(ws.Cells(i, 1).value)) = UCase(custID) Then
            Dim col As Integer
            For col = 5 To 8
                cellVal = Trim(ws.Cells(i, col).value & "")
                If cellVal <> "" And InStr(cellVal, "@") = 0 And Len(cellVal) > 6 Then
                    If Not IsNumeric(Replace(Replace(cellVal, " ", ""), "-", "")) Then
                        GetCustomerAddress = cellVal
                        Exit For
                    End If
                End If
            Next col
            Exit For
        End If
    Next i
End Function

'===========================================
' GET CUSTOMER EMAIL
'===========================================
Private Function GetCustomerEmail(custID As String) As String
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim cellVal As String
    
    GetCustomerEmail = ""
    If Trim(custID) = "" Then Exit Function
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Customer_Master")
    On Error GoTo 0
    If ws Is Nothing Then Exit Function
    
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    
    For i = 2 To lastRow
        If UCase(Trim(ws.Cells(i, 1).value)) = UCase(custID) Then
            Dim col As Integer
            For col = 5 To 8
                cellVal = Trim(ws.Cells(i, col).value & "")
                If InStr(cellVal, "@") > 0 And InStr(cellVal, ".") > 0 Then
                    GetCustomerEmail = cellVal
                    Exit For
                End If
            Next col
            Exit For
        End If
    Next i
End Function

'===========================================
' LOAD PAYMENT LIST
'===========================================
Public Sub LoadPaymentList(entryID As String)
    Dim ws As Worksheet
    Dim lastRow As Long, r As Long
    Dim itm As listItem

    If Trim(entryID) = "" Then
        Me.lstPayments.ListItems.Clear
        Call UpdatePaymentButtons
        Exit Sub
    End If

    Set ws = ThisWorkbook.Sheets("Payment_Master")
    Me.lstPayments.ListItems.Clear

    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row

    For r = 2 To lastRow
        If Trim(ws.Cells(r, 2).value & "") = Trim(entryID & "") Then
            If ws.Cells(r, 6).value = "" Then
                Set itm = Me.lstPayments.ListItems.Add
                itm.text = ws.Cells(r, 1).value
                itm.SubItems(1) = ws.Cells(r, 2).value
                itm.SubItems(2) = ws.Cells(r, 8).value
                itm.SubItems(3) = ws.Cells(r, 9).value
                itm.SubItems(4) = ws.Cells(r, 15).value
                itm.SubItems(5) = ws.Cells(r, 10).value
                itm.SubItems(6) = ws.Cells(r, 11).value
                itm.SubItems(7) = ws.Cells(r, 12).value
                itm.SubItems(8) = ws.Cells(r, 13).value
                itm.SubItems(9) = ws.Cells(r, 14).value
                itm.SubItems(10) = Format(ws.Cells(r, 16).value, "0.00")
                itm.SubItems(11) = ws.Cells(r, 17).value
            End If
        End If
    Next r
    
    Call UpdatePaymentButtons
End Sub

'===========================================
' GET EXISTING PAYMENT ID
'===========================================
Private Function GetExistingPaymentID(entryID As String) As String
    Dim ws As Worksheet
    Dim r As Long, lastRow As Long
    
    If Trim(entryID) = "" Then Exit Function
    
    Set ws = ThisWorkbook.Sheets("Payment_Master")
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    
    For r = 2 To lastRow
        If Trim(ws.Cells(r, 2).value & "") = Trim(entryID & "") Then
            If ws.Cells(r, 6).value = "" Then
                GetExistingPaymentID = ws.Cells(r, 1).value
                Exit Function
            End If
        End If
    Next r
    
    GetExistingPaymentID = ""
End Function

'===========================================
' UPDATE BUTTON STATES
'===========================================
Private Sub UpdatePaymentButtons()
    If Me.lstPayments.ListItems.count > 0 Then
        btnAddPayment.enabled = False
        btnEditPayment.enabled = True
    Else
        btnAddPayment.enabled = True
        btnEditPayment.enabled = False
    End If
    Me.Repaint
End Sub

'===========================================
' ADD CHARGES BUTTON
'===========================================
Private Sub btnAddPayment_Click()
    Dim payForm As frmPaymentSection
    Dim entryID As String
    Dim existingPayID As String
    
    entryID = Trim(ComboBox1.value)
    If entryID = "" Then
        MsgBox "Please select Entry ID first!", vbExclamation
        ComboBox1.SetFocus
        Exit Sub
    End If
    
    existingPayID = GetExistingPaymentID(entryID)
    
    Set payForm = New frmPaymentSection
    Set payForm.parentForm = Me
    
    payForm.customerID = CurrentCustomerID
    
    If existingPayID <> "" Then
        payForm.cmbEntryID.value = entryID
        payForm.cmbEntryID.enabled = False
        payForm.LoadPaymentForEdit existingPayID
    Else
        payForm.cmbEntryID.value = entryID
        payForm.cmbEntryID.enabled = False
    End If
    
    payForm.Show vbModal
    
    Call LoadPaymentList(entryID)
    Set payForm = Nothing
End Sub

'===========================================
' EDIT CHARGES BUTTON
'===========================================
Private Sub btnEditPayment_Click()
    Dim payID As String
    
    On Error Resume Next
    payID = Me.lstPayments.selectedItem.text
    On Error GoTo 0
    
    If payID = "" Then
        MsgBox "Please select a payment first!", vbExclamation
        Exit Sub
    End If
    
    Dim payForm As New frmPaymentSection
    Set payForm.parentForm = Me
    
    payForm.LoadPaymentForEdit payID
    payForm.cmbEntryID.enabled = False
    
    payForm.Show vbModal
    
    Call LoadPaymentList(ComboBox1.value)
End Sub

'===========================================
' LISTVIEW DOUBLE CLICK
'===========================================
Private Sub lstPayments_DblClick()
    Dim payID As String
    
    On Error Resume Next
    payID = Me.lstPayments.selectedItem.text
    On Error GoTo 0
    
    If payID = "" Then Exit Sub

    Dim payForm As New frmPaymentSection
    Set payForm.parentForm = Me

    payForm.LoadPaymentForEdit payID
    payForm.cmbEntryID.enabled = False
    
    payForm.Show vbModal

    Call LoadPaymentList(ComboBox1.value)
End Sub

'===========================================
' GET CHARGES SUMMARY FROM PAYMENT MASTER
'===========================================
Private Sub GetPaymentSummary(entryID As String, ByRef partsNames As String, _
    ByRef laborCharges As Double, ByRef transportCharges As Double, _
    ByRef otherCharges As Double, ByRef actualPartsCost As Double, ByRef totalExpense As Double)
    
    Dim ws As Worksheet
    Dim lastRow As Long, r As Long
    Dim pName As String, pCategory As String, amt As Double
    
    partsNames = "": laborCharges = 0: transportCharges = 0: otherCharges = 0: actualPartsCost = 0: totalExpense = 0
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Payment_Master")
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    
    For r = 2 To lastRow
        If Trim(ws.Cells(r, 2).value & "") = Trim(entryID & "") Then
            If ws.Cells(r, 6).value = "" Then
                pName = Trim(ws.Cells(r, 10).value & "")
                pCategory = UCase(Trim(ws.Cells(r, 9).value & ""))
                amt = val(ws.Cells(r, 16).value)
                
                If pCategory = "TRANSPORT_CHARGE" Then
                    transportCharges = transportCharges + amt
                ElseIf pCategory = "LABOR_CHARGE" Or pCategory = "SERVICE_CHARGE" Then
                    laborCharges = laborCharges + amt
                ElseIf pCategory = "OTHER_CHARGE" Then
                    otherCharges = otherCharges + amt
                Else
                    actualPartsCost = actualPartsCost + amt
                    If partsNames <> "" Then partsNames = partsNames & ", "
                    partsNames = partsNames & pName
                End If
                
                totalExpense = totalExpense + amt
            End If
        End If
    Next r
End Sub

'===========================================
' LOAD ACCESSORIES FROM Job_Accessory
'===========================================
Private Sub LoadAccessories(entryID As String)
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim accCount As Integer
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Job_Accessory")
    On Error GoTo 0
    
    Dim chk As Control
    For Each chk In Me.Controls
        If TypeName(chk) = "CheckBox" Then
            If chk.name <> "chkCustomerApproved" Then
                chk.value = False
                chk.caption = ""
                chk.enabled = False
                chk.visible = False
            End If
        End If
    Next chk
    
    If ws Is Nothing Then Exit Sub
    
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    accCount = 0
    
    For i = 2 To lastRow
        If UCase(Trim(ws.Cells(i, 1).value & "")) = UCase(entryID) Then
            accCount = accCount + 1
            If accCount <= 6 Then
                Dim chkName As String
                chkName = "CheckBox" & accCount
                On Error Resume Next
                Me.Controls(chkName).caption = ws.Cells(i, 4).value
                Me.Controls(chkName).value = True
                Me.Controls(chkName).enabled = True
                Me.Controls(chkName).visible = True
                On Error GoTo 0
            End If
        End If
    Next i
End Sub

'===========================================
' GET ACCESSORIES STATUS STRING
'===========================================
Private Function GetAccessoriesStatus() As String
    Dim chk As Control
    Dim dueStr As String
    Dim okStr As String
    
    dueStr = ""
    okStr = ""
    
    For Each chk In Me.Controls
        If TypeName(chk) = "CheckBox" Then
            If chk.visible And chk.caption <> "" And chk.name <> "chkCustomerApproved" Then
                If chk.value = False Then
                    If dueStr <> "" Then dueStr = dueStr & ","
                    dueStr = dueStr & chk.caption & "(DUE)"
                Else
                    If okStr <> "" Then okStr = okStr & ","
                    okStr = okStr & chk.caption & "(OK)"
                End If
            End If
        End If
    Next chk
    
    If dueStr = "" Then
        GetAccessoriesStatus = "ALL RETURNED"
    Else
        GetAccessoriesStatus = okStr & "|" & dueStr
    End If
End Function

'===========================================
' USERFORM ACTIVATE (Dashboard double-click backup)
'===========================================
Private Sub UserForm_Activate()
    Dim entryID As String
    Dim i As Long
    Dim found As Boolean
    
    entryID = Trim(Me.Tag)
    
    If entryID <> "" Then
        found = False
        For i = 0 To ComboBox1.ListCount - 1
            If UCase(Trim(ComboBox1.List(i))) = UCase(entryID) Then
                found = True
                Exit For
            End If
        Next i
        
        If Not found Then ComboBox1.AddItem entryID
        
        On Error Resume Next
        ComboBox1.value = entryID
        On Error GoTo 0
        
        ' If not already in edit mode, load work log now
        If Not isEditMode Then
            isEditMode = True
            IsLoadingData = True
            Call LoadServiceWorkLogData(entryID)
            IsLoadingData = False
            Me.caption = "SERVICE WORK LOG - EDIT (" & entryID & ")"
            btnSave.caption = "UPDATE WORK LOG"
            btnUpdateToCustomer.enabled = True
        End If
        
        Me.Tag = ""
    End If
End Sub


