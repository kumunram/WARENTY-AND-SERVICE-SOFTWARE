VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPaymentSection 
   Caption         =   "PAYMENT SECTION"
   ClientHeight    =   13815
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   18495
   OleObjectBlob   =   "frmPaymentSection.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmPaymentSection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public parentForm As Object

Dim isLoading As Boolean
Dim currentPaymentID As String
Dim editMode As Boolean
Dim editingTransactionIndex As Long
Public customerID As String

'=================================
' SAVE BUTTON (A-R Mapping)
'=================================
Private Sub btnSave_Click()
    Call CalculateSummary
    
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim paymentID As String
    Dim r As Long

    Set ws = ThisWorkbook.Sheets("Payment_Master")
    
    ' === EDIT/NEW MODE LOGIC (???? existing code) ===
    If editMode And currentPaymentID <> "" Then
        paymentID = currentPaymentID
        For r = ws.Cells(ws.Rows.count, 1).End(xlUp).row To 2 Step -1
            If Trim(ws.Cells(r, 1).value & "") = paymentID Then
                ws.Rows(r).Delete
            End If
        Next r
    Else
        Dim existingPayID As String
        existingPayID = ""
        
        For r = 2 To ws.Cells(ws.Rows.count, 1).End(xlUp).row
            If Trim(ws.Cells(r, 2).value & "") = Trim(Me.cmbEntryID.value & "") Then
                If ws.Cells(r, 6).value = "" Then
                    existingPayID = ws.Cells(r, 1).value
                    Exit For
                End If
            End If
        Next r
        
        If existingPayID <> "" Then
            paymentID = existingPayID
            For r = ws.Cells(ws.Rows.count, 1).End(xlUp).row To 2 Step -1
                If Trim(ws.Cells(r, 1).value & "") = paymentID Then
                    ws.Rows(r).Delete
                End If
            Next r
        Else
            paymentID = modPaymentMaster.GeneratePaymentID()
        End If
    End If

    ' === A-R Mapping (???? existing code) ===
    For i = 1 To lstAllTransactions.ListItems.count
        lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row + 1
        
        ws.Cells(lastRow, 1).value = paymentID
        ws.Cells(lastRow, 2).value = Me.cmbEntryID.value
        ' ... (baki columns same as your code) ...
        On Error Resume Next
        ws.Cells(lastRow, 3).value = Me.customerID
        If ws.Cells(lastRow, 3).value = "" Then ws.Cells(lastRow, 3).value = parentForm.txtCustomerID.caption
        If ws.Cells(lastRow, 3).value = "" Then ws.Cells(lastRow, 3).value = parentForm.txtCustomerID.value
        If ws.Cells(lastRow, 3).value = "" Then ws.Cells(lastRow, 3).value = parentForm.lblCustomerID.caption
        If ws.Cells(lastRow, 3).value = "" Then ws.Cells(lastRow, 3).value = parentForm.lblCustomerIDValue.caption
        On Error GoTo 0
        If ws.Cells(lastRow, 3).value = "" Then ws.Cells(lastRow, 3).value = "CUST00000"
        ws.Cells(lastRow, 4).value = ""
        ws.Cells(lastRow, 5).value = Date
        ws.Cells(lastRow, 6).value = ""
        ws.Cells(lastRow, 7).value = ""
        ws.Cells(lastRow, 8).value = lstAllTransactions.ListItems(i).SubItems(1)
        ws.Cells(lastRow, 9).value = lstAllTransactions.ListItems(i).SubItems(2)
        ws.Cells(lastRow, 10).value = lstAllTransactions.ListItems(i).SubItems(4)
        ws.Cells(lastRow, 11).value = lstAllTransactions.ListItems(i).SubItems(5)
        ws.Cells(lastRow, 12).value = lstAllTransactions.ListItems(i).SubItems(6)
        ws.Cells(lastRow, 13).value = lstAllTransactions.ListItems(i).SubItems(7)
        ws.Cells(lastRow, 14).value = lstAllTransactions.ListItems(i).SubItems(8)
        ws.Cells(lastRow, 15).value = lstAllTransactions.ListItems(i).SubItems(3)
        ws.Cells(lastRow, 16).value = lstAllTransactions.ListItems(i).SubItems(9)
        ws.Cells(lastRow, 17).value = lstAllTransactions.ListItems(i).SubItems(10)
        ws.Cells(lastRow, 18).value = lstAllTransactions.ListItems(i).SubItems(11)
    Next i
    
    ' ============================================================
    ' ?? CRITICAL FIX: SAVE WITH PROPER ERROR HANDLING
    ' ============================================================
    On Error GoTo SaveFailed
    
    Application.DisplayAlerts = False
    ThisWorkbook.Save
    Application.DisplayAlerts = True
    
    On Error GoTo 0
    ' ============================================================

    ' Refresh parent form
    If Not parentForm Is Nothing Then
        parentForm.LoadPaymentList parentForm.txtCustomerID.caption
    End If
    
    ' Reset variables
    editMode = False
    currentPaymentID = ""
    editingTransactionIndex = 0
    btnDeleteTransaction.enabled = True
    
    MsgBox "Payment Saved Successfully!" & vbCrLf & "Payment ID: " & paymentID, vbInformation
    Call ClearForm
    Unload Me
    Exit Sub

SaveFailed:
    Application.DisplayAlerts = True
    MsgBox "CRITICAL ERROR: Payment data written to sheet but FILE NOT SAVED!" & vbCrLf & _
           "Error: " & Err.Description & vbCrLf & vbCrLf & _
           "Please press Ctrl+S manually to save the file immediately!" & vbCrLf & _
           "Do NOT close Excel until you save!", vbCritical, "SAVE FAILED - ACTION REQUIRED"
    
    ' Still refresh parent so user sees data
    If Not parentForm Is Nothing Then
        On Error Resume Next
        parentForm.LoadPaymentList parentForm.txtCustomerID.caption
        On Error GoTo 0
    End If
End Sub

Private Sub frmPayment_Click()

End Sub

'=================================
' FORM INITIALIZE
'=================================
Private Sub UserForm_Initialize()
    isLoading = True
AddMinMaxButtons Me
    Call LoadEntryIDs

    cmbType.Clear
    cmbType.AddItem "PAYMENT"
    cmbType.AddItem "PART"

    With Me.lstProductDetail
        .View = lvwReport
        .FullRowSelect = True
        .Gridlines = True
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "Entry ID", 80
        .ColumnHeaders.Add , , "Product", 100
        .ColumnHeaders.Add , , "Company", 80
        .ColumnHeaders.Add , , "Model", 80
        .ColumnHeaders.Add , , "Serial", 100
        .ColumnHeaders.Add , , "Warranty", 80
        .ColumnHeaders.Add , , "Status", 80
    End With
    
    cmbCategory.Clear
    cmbCategory.AddItem "DISCUSSION"
    cmbCategory.AddItem "PART_USED"
    cmbCategory.AddItem "PART_REPLACED"
    cmbCategory.AddItem "CHARGE"
    cmbCategory.AddItem "EXPENSE"
    cmbCategory.AddItem "RECEIVED"
    cmbCategory.AddItem "ADVANCE"
    cmbCategory.AddItem "DISCOUNT"
    cmbCategory.AddItem "TRANSPORT_CHARGE"

    cmbWarranty.Clear
    cmbWarranty.AddItem "YES"
    cmbWarranty.AddItem "NO"

    cmbWarrantyDays.Clear
    cmbWarrantyDays.AddItem "30 Days"
    cmbWarrantyDays.AddItem "60 Days"
    cmbWarrantyDays.AddItem "90 Days"
    cmbWarrantyDays.AddItem "180 Days"
    cmbWarrantyDays.AddItem "365 Days"
    cmbWarrantyDays.AddItem "730 Days"
    cmbWarrantyDays.AddItem "+ Add New"

    cmbPaymentMode.Clear
    cmbPaymentMode.AddItem "+ Add New"
    cmbPaymentMode.AddItem "CASH"
    cmbPaymentMode.AddItem "NET BANKING"
    cmbPaymentMode.AddItem "UPI"
    cmbPaymentMode.AddItem "CHEQUE"
    cmbPaymentMode.AddItem "CREDIT CARD"
    
    txtQty.value = "1"

    Call SetupNewListView
    Call ClearSummary
    ' ===== RESET MODULE VARIABLES =====
    editMode = False
    currentPaymentID = ""
    editingTransactionIndex = 0
    
    ' ===== EXPLICIT BLANK =====
    cmbType.ListIndex = -1
    cmbCategory.ListIndex = -1
    txtQty.value = ""
    
            ' ===== DELETE BUTTON: Entry time = Enable, Edit mode = Disable =====
    On Error Resume Next
    btnDeleteTransaction.enabled = True   ' New entry mein Enable
    On Error GoTo 0
    
    isLoading = False
    
    
End Sub

'=================================
' SETUP LISTVIEW (12 COLUMNS)
' Order: Date|Type|Category|PayMode|Name|Company|Model|Serial|Qty|Amount|Warranty|Remark
'=================================
Sub SetupNewListView()
    With Me.lstAllTransactions
        .View = lvwReport
        .Gridlines = True
        .FullRowSelect = True
        .ColumnHeaders.Clear
        
        .ColumnHeaders.Add , , "Date", 80
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
        .ColumnHeaders.Add , , "Remark", 120
    End With
End Sub

Private Sub cmbWarrantyDays_Change()
    If cmbWarrantyDays.value = "+ Add New" Then
        Dim newDay As String
        newDay = InputBox("Enter Warranty Days")
        If newDay <> "" Then
            cmbWarrantyDays.AddItem newDay
            cmbWarrantyDays.value = newDay
        End If
    End If
End Sub

Private Sub cmbCategory_Change()
    If cmbCategory.value = "+ Add New" Then
        Dim newCat As String
        newCat = InputBox("Enter New Category")
        If newCat <> "" Then
            cmbCategory.AddItem newCat
            cmbCategory.value = newCat
        End If
    End If
End Sub

'=================================
' COMPANY "+ Add New" HANDLER
'=================================
Private Sub cmbCompany_Change()
    If cmbCompany.value = "+ Add New" Then
        Dim newComp As String
        newComp = InputBox("Enter New Company Name:")
        If Trim(newComp) <> "" Then
            cmbCompany.AddItem newComp, cmbCompany.ListCount - 1
            cmbCompany.value = newComp
        Else
            cmbCompany.ListIndex = -1
        End If
    End If
End Sub

Sub LoadCustomerDetails(entryID As String)
    Dim wsJob As Worksheet
    Dim wsCust As Worksheet
    Dim r As Long
    Dim custID As String

    Set wsJob = ThisWorkbook.Sheets("Job_Product")
    For r = 2 To wsJob.Cells(wsJob.Rows.count, 1).End(xlUp).row
        If wsJob.Cells(r, 1).value = entryID Then
            custID = wsJob.Cells(r, 2).value
            Exit For
        End If
    Next r

    Set wsCust = ThisWorkbook.Sheets("Customer_Master")
    For r = 2 To wsCust.Cells(wsCust.Rows.count, 1).End(xlUp).row
        If wsCust.Cells(r, 1).value = custID Then
            lblCustomerIDValue.caption = wsCust.Cells(r, 1).value
            lblCustomerValue.caption = wsCust.Cells(r, 3).value
            lblMobileValue.caption = wsCust.Cells(r, 2).value
            Exit For
        End If
    Next r
End Sub

Private Sub cmbEntryID_Change()
    If isLoading Then Exit Sub
    If cmbEntryID.value = "" Then Exit Sub
    Call LoadProductDetails(cmbEntryID.value)
    Call LoadCustomerDetails(cmbEntryID.value)
End Sub

Private Sub cmbName_Change()
    If isLoading Then Exit Sub

    If cmbName.value = "+ Add New" Then
        Dim newPart As String
        newPart = InputBox("Enter New Part Name")
        If newPart <> "" Then
            cmbName.AddItem newPart
            cmbName.value = newPart
            Call SaveNewPart(newPart)
        Else
            cmbName.ListIndex = -1
        End If
        Exit Sub
    End If

    ' AUTO LOAD from Part_Master: C=Company, D=Model, G=WarrantyDays
    If cmbName.value <> "" And cmbType.value = "PART" Then
        Call LoadModelFromPartMaster(cmbName.value)
        Call LoadCompanyFromPartMaster(cmbName.value)
    End If
End Sub

Sub SaveNewPart(partName As String)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim comp As String
    Dim model As String
    Dim wDays As String
    Dim price As String

    Set ws = ThisWorkbook.Sheets("Part_Master")
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row + 1

    ' ===== SAB KUCH POOCHO =====
    comp = InputBox("Enter Company Name for '" & partName & "':" & vbCrLf & "(Example: INTEL, HP, DELL)")
    model = InputBox("Enter Model Number for '" & partName & "':" & vbCrLf & "(Example: H61, i5, etc.)")
    price = InputBox("Enter Price for '" & partName & "':" & vbCrLf & "(Example: 2000, 1500, etc.)")
    wDays = InputBox("Enter Warranty Days for '" & partName & "':" & vbCrLf & "(Example: 365, 180, 90)")

    ' ===== A-G COLUMNS MEIN SAVE =====
    ws.Cells(lastRow, 1).value = "PRT" & Format(lastRow - 1, "0000")  ' A: PartID
    ws.Cells(lastRow, 2).value = partName                              ' B: PartName
    ws.Cells(lastRow, 3).value = comp                                  ' C: Company
    ws.Cells(lastRow, 4).value = model                                 ' D: Model
    ws.Cells(lastRow, 6).value = val(price)                            ' F: Price
    ws.Cells(lastRow, 7).value = val(wDays)                            ' G: WarrantyDays

    ' ===== DROPDOWN REFRESH =====
    Call LoadPartMaster
    
    ' ===== AUTO-SELECT THE NEW PART =====
    cmbName.value = partName
    Call LoadModelFromPartMaster(partName)
    Call LoadCompanyFromPartMaster(partName)

    MsgBox "Part Saved!" & vbCrLf & _
           "ID: " & ws.Cells(lastRow, 1).value & vbCrLf & _
           "Name: " & partName & vbCrLf & _
           "Company: " & comp & vbCrLf & _
           "Model: " & model & vbCrLf & _
           "Price: " & price & vbCrLf & _
           "Warranty: " & wDays & " Days", vbInformation
End Sub

'=================================
' LOAD PART MASTER (A=ID, B=Name, C=Company, D=Model, G=WarrantyDays)
'=================================
Sub LoadPartMaster()
    Dim ws As Worksheet
    Dim r As Long
    Set ws = ThisWorkbook.Sheets("Part_Master")
    
    cmbName.Clear
    cmbCompany.Clear
    cmbModel.Clear
    
    For r = 2 To ws.Cells(ws.Rows.count, 1).End(xlUp).row
        If Trim(ws.Cells(r, 2).value & "") <> "" Then
            If Not IsInCombo(cmbName, ws.Cells(r, 2).value) Then cmbName.AddItem ws.Cells(r, 2).value
        End If
        If Trim(ws.Cells(r, 3).value & "") <> "" Then
            If Not IsInCombo(cmbCompany, ws.Cells(r, 3).value) Then cmbCompany.AddItem ws.Cells(r, 3).value
        End If
        If Trim(ws.Cells(r, 4).value & "") <> "" Then
            If Not IsInCombo(cmbModel, ws.Cells(r, 4).value) Then cmbModel.AddItem ws.Cells(r, 4).value
        End If
    Next r
    
    cmbName.AddItem "+ Add New"
    cmbCompany.AddItem "+ Add New"
    cmbModel.AddItem "+ Add New"
End Sub

Private Sub cmbModel_Change()
    If cmbModel.value = "+ Add New" Then
        Dim newModel As String
        newModel = InputBox("Enter New Model")
        If newModel <> "" Then
            cmbModel.AddItem newModel
            cmbModel.value = newModel
        End If
    End If
End Sub

Private Sub cmbType_Change()
    cmbCategory.Clear

    If cmbType.value = "PAYMENT" Then
        cmbCategory.AddItem "DISCUSSION"
        cmbCategory.AddItem "ADVANCE"
        cmbCategory.AddItem "CHARGE"
        cmbCategory.AddItem "EXPENSE"
        cmbCategory.AddItem "RECEIVED"
        cmbCategory.AddItem "DISCOUNT"

                cmbName.enabled = False
        cmbCompany.enabled = False
        cmbModel.enabled = False
        txtSerial.enabled = False
        txtQty.enabled = False
        cmbWarranty.enabled = False
        cmbWarrantyDays.enabled = False
        cmbCompany.value = ""
        
        ' ===== QTY BLANK IN PAYMENT =====
        txtQty.value = ""

    ElseIf cmbType.value = "PART" Then
        cmbCategory.AddItem "PART_USED"
        cmbCategory.AddItem "PART_REPLACED"
        Call LoadPartMaster

                cmbName.enabled = True
        cmbCompany.enabled = True
        cmbModel.enabled = True
        txtSerial.enabled = True
        txtQty.enabled = True
        cmbWarranty.enabled = True
        cmbWarrantyDays.enabled = True
        
        ' ===== QTY DEFAULT 1 IN PART =====
        txtQty.value = "1"
    End If
End Sub

Public Sub LoadEntryIDs()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim custID As String
    
    ' ===== PEHLE CHECK KARO: Agar customerID directly set hua hai =====
    If Trim(Me.customerID & "") <> "" Then
        custID = Me.customerID
    Else
        ' ===== PARENT FORM se flexible retrieve =====
        On Error Resume Next
        custID = parentForm.txtCustomerID.caption      ' Entry Form
        If custID = "" Then custID = parentForm.txtCustomerID.value
        If custID = "" Then custID = parentForm.lblCustomerID.caption
        If custID = "" Then custID = parentForm.lblCustomerIDValue.caption
        On Error GoTo 0
    End If
    
    If custID = "" Then Exit Sub
    
    Set ws = ThisWorkbook.Sheets("Job_Product")
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    
    Me.cmbEntryID.Clear
    For r = 2 To lastRow
        If ws.Cells(r, 2).value = custID Then
            If Not IsInCombo(Me.cmbEntryID, ws.Cells(r, 1).value) Then
                Me.cmbEntryID.AddItem ws.Cells(r, 1).value
            End If
        End If
    Next r
    
    If Me.cmbEntryID.ListCount = 1 Then
        Me.cmbEntryID.ListIndex = 0
    End If
End Sub

Function IsInCombo(cmb As ComboBox, val As String) As Boolean
    Dim i As Integer
    For i = 0 To cmb.ListCount - 1
        If cmb.List(i) = val Then
            IsInCombo = True
            Exit Function
        End If
    Next i
    IsInCombo = False
End Function

Sub LoadProductDetails(entryID As String)
    Dim ws As Worksheet
    Dim lastRow As Long, r As Long
    Dim itm As listItem
    
    Set ws = ThisWorkbook.Sheets("Job_Product")
    Me.lstProductDetail.ListItems.Clear
    
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    For r = 2 To lastRow
        If Trim(ws.Cells(r, 1).value & "") = entryID Then
            Set itm = Me.lstProductDetail.ListItems.Add
            itm.text = ws.Cells(r, 1).value
            itm.SubItems(1) = ws.Cells(r, 6).value
            itm.SubItems(2) = ws.Cells(r, 7).value
            itm.SubItems(3) = ws.Cells(r, 8).value
            itm.SubItems(4) = ws.Cells(r, 9).value
            itm.SubItems(5) = ws.Cells(r, 11).value
            itm.SubItems(6) = ws.Cells(r, 4).value
            Exit For
        End If
    Next r
End Sub

Sub ClearSummary()
    Me.txtTotalCharge.value = "0.00"
    Me.txtTotalExpense.value = "0.00"
    Me.txtTotalReceived.value = "0.00"
    Me.txtDiscount.value = "0.00"
    Me.txtPending.value = "0.00"
End Sub

Private Sub btnClose_Click()
    editMode = False
    currentPaymentID = ""
    editingTransactionIndex = 0
    Unload Me
End Sub

Private Sub cmbPaymentMode_Change()
    If Me.cmbPaymentMode.value = "+ Add New" Then
        Dim newMode As String
        newMode = InputBox("Enter New Payment Mode:")
        If Trim(newMode) <> "" Then
            Me.cmbPaymentMode.AddItem newMode, Me.cmbPaymentMode.ListCount - 1
            Me.cmbPaymentMode.value = newMode
        Else
            Me.cmbPaymentMode.ListIndex = -1
        End If
    End If
End Sub

Private Sub cmbExpenseType_Change()
    If Me.cmbExpenseType.value = "+ Add New" Then
        Dim newType As String
        newType = InputBox("Enter New Expense Type:")
        If Trim(newType) <> "" Then
            Me.cmbExpenseType.AddItem newType, Me.cmbExpenseType.ListCount - 1
            Me.cmbExpenseType.value = newType
        Else
            Me.cmbExpenseType.ListIndex = -1
        End If
    End If
End Sub

Private Sub cmbReceiptMode_Change()
    If Me.cmbReceiptMode.value = "+ Add New" Then
        Dim newMode As String
        newMode = InputBox("Enter New Receipt Mode:")
        If Trim(newMode) <> "" Then
            Me.cmbReceiptMode.AddItem newMode, Me.cmbReceiptMode.ListCount - 1
            Me.cmbReceiptMode.value = newMode
        Else
            Me.cmbReceiptMode.ListIndex = -1
        End If
    End If
End Sub

'=================================
' CALCULATE SUMMARY (Amount = SubItems(9))
'=================================
'=================================
' CALCULATE SUMMARY — DISCOUNT AUTO FROM LISTVIEW
'=================================
'=================================
' CALCULATE SUMMARY — DISCOUNT AUTO FROM LISTVIEW
'=================================
Sub CalculateSummary()
    Dim i As Long
    Dim totalCharge As Double
    Dim totalExpense As Double
    Dim totalReceived As Double
    Dim discount As Double

    totalCharge = 0
    totalExpense = 0
    totalReceived = 0
    discount = 0

    For i = 1 To lstAllTransactions.ListItems.count
        Dim typ As String
        Dim cat As String
        Dim amt As Double

        typ = lstAllTransactions.ListItems(i).SubItems(1)
        cat = lstAllTransactions.ListItems(i).SubItems(2)
        amt = val(lstAllTransactions.ListItems(i).SubItems(9))

        If typ = "PAYMENT" Then
            If cat = "CHARGE" Then totalCharge = totalCharge + amt
            If cat = "EXPENSE" Then totalExpense = totalExpense + amt
            If cat = "RECEIVED" Then totalReceived = totalReceived + amt
            If cat = "ADVANCE" Then totalReceived = totalReceived + amt
            ' ===== DISCOUNT AUTO DETECT — YE MINUS HOGA =====
            If cat = "DISCOUNT" Then discount = discount + amt
        End If

        If typ = "PART" Then
            totalCharge = totalCharge + amt
        End If
    Next i

    ' ===== DISCOUNT BOX AUTO SHOW =====
    txtDiscount.value = Format(discount, "0.00")

    txtTotalCharge.value = Format(totalCharge, "0.00")
    txtTotalExpense.value = Format(totalExpense, "0.00")
    txtTotalReceived.value = Format(totalReceived, "0.00")
    
    ' ===== FINAL: MINUS DISCOUNT =====
    txtPending.value = Format(totalCharge + totalExpense - totalReceived - discount, "0.00")
End Sub



Private Sub btnAddTransaction_Click()
    Dim transAmount As Double
    Dim useAmount As Double
    
    ' Validation
    If cmbType.value = "" Or cmbCategory.value = "" Then
        MsgBox "Select Type and Category!", vbExclamation
        Exit Sub
    End If

    ' ===== AMOUNT CHECK: Type ke hisaab se =====
    If cmbType.value = "PART" Then
        ' PART: Right side amount (Product Amount) use hoga
        If txtProductAmount.value = "" Or val(txtProductAmount.value) <= 0 Then
            MsgBox "Enter valid Product Amount (Right side)!", vbExclamation
            Exit Sub
        End If
        useAmount = val(txtProductAmount.value)
    Else
        ' PAYMENT: Left side amount (Transaction Amount) use hoga
        If txtAmount.value = "" Or val(txtAmount.value) <= 0 Then
            MsgBox "Enter valid Transaction Amount (Left side)!", vbExclamation
            Exit Sub
        End If
        useAmount = val(txtAmount.value)
    End If

    Dim itm As listItem

    If editingTransactionIndex > 0 Then
        ' UPDATE MODE
        Set itm = lstAllTransactions.ListItems(editingTransactionIndex)
        itm.SubItems(1) = cmbType.value
        itm.SubItems(2) = cmbCategory.value
        itm.SubItems(3) = cmbPaymentMode.value
        itm.SubItems(4) = cmbName.value
        itm.SubItems(5) = cmbCompany.value
        itm.SubItems(6) = cmbModel.value
        itm.SubItems(7) = txtSerial.value
        itm.SubItems(8) = txtQty.value
        itm.SubItems(9) = Format(useAmount, "0.00")   ' ? Correct amount
        itm.SubItems(10) = cmbWarranty.value
        itm.SubItems(11) = txtRemark.value
        
        editingTransactionIndex = 0
        Me.btnAddTransaction.caption = "Add Payment & Product"
    Else
        ' ADD NEW MODE
        Set itm = lstAllTransactions.ListItems.Add
        itm.text = Format(Date, "dd-mm-yyyy")
        itm.SubItems(1) = cmbType.value
        itm.SubItems(2) = cmbCategory.value
        itm.SubItems(3) = cmbPaymentMode.value
        itm.SubItems(4) = cmbName.value
        itm.SubItems(5) = cmbCompany.value
        itm.SubItems(6) = cmbModel.value
        itm.SubItems(7) = txtSerial.value
        itm.SubItems(8) = txtQty.value
        itm.SubItems(9) = Format(useAmount, "0.00")   ' ? Correct amount
        itm.SubItems(10) = cmbWarranty.value
        itm.SubItems(11) = txtRemark.value
    End If

    ' ===== CLEAR ALL UPPER FIELDS =====
    cmbType.value = ""
    cmbCategory.value = ""
    cmbPaymentMode.value = ""
    cmbName.value = ""
    cmbCompany.value = ""
    cmbModel.value = ""
    txtSerial.value = ""
    txtQty.value = ""
    txtAmount.value = ""           ' Left side clear
    txtProductAmount.value = ""    ' Right side clear  ? YE NAAM CHECK KARO
    cmbWarranty.value = ""
    cmbWarrantyDays.value = ""
    txtRemark.value = ""

    Call CalculateSummary
End Sub

'=================================
' CLEAR FORM
'=================================
Sub ClearForm()
    lstAllTransactions.ListItems.Clear
    editMode = False
    currentPaymentID = ""
    editingTransactionIndex = 0
    Me.btnAddTransaction.caption = "Add Payment"
    
    cmbEntryID.value = ""
    cmbType.value = ""
    cmbCategory.value = ""
    cmbName.value = ""
    cmbCompany.value = ""
    cmbModel.value = ""
    
        txtAmount.value = ""           ' Left side
    txtProductAmount.value = ""    ' Right side  ? YE NAAM CHECK KARO
    txtQty.value = ""
    txtSerial.value = ""
    txtRemark.value = ""
    cmbPaymentMode.value = ""
    
    txtTotalCharge.value = "0.00"
    txtTotalExpense.value = "0.00"
    txtTotalReceived.value = "0.00"
    txtDiscount.value = "0.00"
    txtPending.value = "0.00"
End Sub

'=================================
' LOAD FOR EDIT (A-R Mapping)
'=================================
Public Sub LoadPaymentForEdit(payID As String)
    Dim ws As Worksheet
    Dim r As Long, lastRow As Long
    Dim itm As listItem
    Dim found As Boolean
    Dim cleanPayID As String
    
    ' ===== EMPTY CHECK =====
    If Trim(payID & "") = "" Then
        MsgBox "Payment ID is empty!", vbExclamation
        editMode = False
        currentPaymentID = ""
        Exit Sub
    End If
    
    cleanPayID = UCase(Trim(payID & ""))
    
    Set ws = ThisWorkbook.Sheets("Payment_Master")
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    
    isLoading = True
    
    lstAllTransactions.ListItems.Clear
    editMode = True
    currentPaymentID = cleanPayID
    found = False
    
    For r = 2 To lastRow
        ' ===== ROBUST MATCH: UCase + Trim =====
        If UCase(Trim(ws.Cells(r, 1).value & "")) = cleanPayID Then
            If Not found Then
                found = True
                
                Me.cmbEntryID.value = ws.Cells(r, 2).value & ""
                
                ' Header fields CLEAR
                Me.cmbType.value = ""
                Me.cmbCategory.value = ""
                Me.cmbName.value = ""
                Me.cmbCompany.value = ""
                Me.cmbModel.value = ""
                Me.txtSerial.value = ""
                Me.txtQty.value = ""
                Me.txtAmount.value = ""
                Me.cmbWarranty.value = ""
                Me.cmbWarrantyDays.value = ""
                Me.cmbPaymentMode.value = ""
                Me.txtRemark.value = ""
                
                Call LoadCustomerDetails(Me.cmbEntryID.value)
                Call LoadProductDetails(Me.cmbEntryID.value)
            End If
            
            ' ListView load (A-R mapping)
            Set itm = lstAllTransactions.ListItems.Add
            itm.text = Format(ws.Cells(r, 5).value, "dd-mm-yyyy")   ' E: Date
            itm.SubItems(1) = ws.Cells(r, 8).value   ' H: Type
            itm.SubItems(2) = ws.Cells(r, 9).value   ' I: Category
            itm.SubItems(3) = ws.Cells(r, 15).value  ' O: PaymentMode
            itm.SubItems(4) = ws.Cells(r, 10).value  ' J: Name
            itm.SubItems(5) = ws.Cells(r, 11).value  ' K: Company
            itm.SubItems(6) = ws.Cells(r, 12).value  ' L: Model
            itm.SubItems(7) = ws.Cells(r, 13).value  ' M: Serial
            itm.SubItems(8) = ws.Cells(r, 14).value  ' N: Qty
            itm.SubItems(9) = Format(ws.Cells(r, 16).value, "0.00") ' P: Amount
            itm.SubItems(10) = ws.Cells(r, 17).value ' Q: Warranty
            itm.SubItems(11) = ws.Cells(r, 18).value ' R: Remarks
        End If
    Next r
    
    isLoading = False
    
    ' ===== MANUAL AUTO-LOAD: Edit mode mein cmbName_Change block hota hai =====
    If Me.cmbType.value = "PART" And Me.cmbName.value <> "" Then
        Call LoadModelFromPartMaster(Me.cmbName.value)
        Call LoadCompanyFromPartMaster(Me.cmbName.value)
    End If
    
    If found Then
        Call CalculateSummary
        editingTransactionIndex = 0
        Me.btnAddTransaction.caption = "Add Payment"
    Else
        MsgBox "Payment ID '" & payID & "' not found in database!" & vbCrLf & _
               "Please check Payment_Master sheet.", vbExclamation
        editMode = False
        currentPaymentID = ""
    End If
End Sub

'=================================
' DOUBLE CLICK TO EDIT
'=================================
Private Sub lstAllTransactions_DblClick()
    Dim itm As listItem
    
    If lstAllTransactions.selectedItem Is Nothing Then Exit Sub
    Set itm = lstAllTransactions.selectedItem
    
    Me.cmbType.value = itm.SubItems(1)
    Me.cmbCategory.value = itm.SubItems(2)
    Me.cmbPaymentMode.value = itm.SubItems(3)
    Me.cmbName.value = itm.SubItems(4)
    Me.cmbCompany.value = itm.SubItems(5)
    Me.cmbModel.value = itm.SubItems(6)
    Me.txtSerial.value = itm.SubItems(7)
    Me.txtQty.value = itm.SubItems(8)
    
    ' ===== AMOUNT LOAD: Type ke hisaab se =====
    If itm.SubItems(1) = "PART" Then
        Me.txtProductAmount.value = itm.SubItems(9)   ' Right side
        Me.txtAmount.value = ""
    Else
        Me.txtAmount.value = itm.SubItems(9)           ' Left side
        Me.txtProductAmount.value = ""
    End If
    
    
        Me.txtAmount.value = itm.SubItems(9)
    Me.cmbWarranty.value = itm.SubItems(10)
    Me.txtRemark.value = itm.SubItems(11)
    
    ' ===== MANUAL AUTO-LOAD: DblClick edit mein bhi =====
    If itm.SubItems(1) = "PART" And Me.cmbName.value <> "" Then
        Call LoadModelFromPartMaster(Me.cmbName.value)
        Call LoadCompanyFromPartMaster(Me.cmbName.value)
    End If
    
    editingTransactionIndex = itm.index
    
    
    Me.btnAddTransaction.caption = "Update"
    
    If itm.SubItems(1) = "PART" Then
        cmbName.enabled = True
        cmbCompany.enabled = True
        cmbModel.enabled = True
        txtSerial.enabled = True
        txtQty.enabled = True
        cmbWarranty.enabled = True
        cmbWarrantyDays.enabled = True
    End If
End Sub

'=================================
' AUTO LOAD MODEL (Part_Master D=Model, G=WarrantyDays)
'=================================
Sub LoadModelFromPartMaster(partName As String)
    Dim ws As Worksheet
    Dim r As Long
    Dim modelName As String
    Dim wDays As String
    Dim price As String
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Part_Master")
    
    For r = 2 To ws.Cells(ws.Rows.count, 1).End(xlUp).row
        If Trim(ws.Cells(r, 2).value & "") = Trim(partName & "") Then
            
            ' Model = Column D
            modelName = ws.Cells(r, 4).value & ""
            If modelName <> "" Then
                If Not IsInCombo(cmbModel, modelName) Then cmbModel.AddItem modelName
                cmbModel.value = modelName
            End If
            
            ' Warranty Days = Column G
            wDays = ws.Cells(r, 7).value & ""
            If wDays <> "" And IsNumeric(wDays) Then
                Dim dayText As String
                dayText = wDays & " Days"
                If Not IsInCombo(cmbWarrantyDays, dayText) Then cmbWarrantyDays.AddItem dayText
                cmbWarrantyDays.value = dayText
                cmbWarranty.value = "YES"  ' ? Auto YES if warranty exists
            End If
            
            ' Price = Column F ? Right-side Product Amount
            price = ws.Cells(r, 6).value & ""
            If price <> "" And IsNumeric(price) Then
                txtProductAmount.value = Format(val(price), "0.00")
            End If
            
            Exit For
        End If
    Next r
    On Error GoTo 0
End Sub
'=================================
' AUTO LOAD COMPANY (Part_Master C=Company)
'=================================
Sub LoadCompanyFromPartMaster(partName As String)
    Dim ws As Worksheet
    Dim r As Long
    Dim comp As String
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Part_Master")
    
    For r = 2 To ws.Cells(ws.Rows.count, 1).End(xlUp).row
        If Trim(ws.Cells(r, 2).value & "") = Trim(partName & "") Then
            comp = ws.Cells(r, 3).value & ""  ' Column C = Company
            If comp <> "" Then
                If Not IsInCombo(cmbCompany, comp) Then cmbCompany.AddItem comp
                cmbCompany.value = comp
            End If
            Exit For
        End If
    Next r
    On Error GoTo 0
End Sub
Private Sub btnDeleteTransaction_Click()
    Dim itm As listItem
    
    If lstAllTransactions.selectedItem Is Nothing Then
        MsgBox "Please select a transaction first!", vbExclamation
        Exit Sub
    End If
    
    Set itm = lstAllTransactions.selectedItem
    
    If MsgBox("Delete this transaction?" & vbCrLf & _
              "Type: " & itm.SubItems(1) & vbCrLf & _
              "Name: " & itm.SubItems(4), vbQuestion + vbYesNo, "Confirm Delete") = vbYes Then
        
        lstAllTransactions.ListItems.Remove itm.index
        Call CalculateSummary
        
        ' ============================================================
        ' OPTIONAL: Auto-save after delete (if you want immediate save)
        ' ============================================================
         On Error Resume Next
         Application.DisplayAlerts = False
         ThisWorkbook.Save
         Application.DisplayAlerts = True
         On Error GoTo 0
        ' ============================================================
        
        MsgBox "Transaction deleted from list!" & vbCrLf & _
               "Click 'Save Payment' to finalize changes.", vbInformation
    End If
End Sub

Private Sub txtDiscount_Change()
    Call CalculateSummary
End Sub
