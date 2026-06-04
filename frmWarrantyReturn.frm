VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmWarrantyReturn 
   Caption         =   "Warranty Return - Receive Product"
   ClientHeight    =   15015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   20235
   OleObjectBlob   =   "frmWarrantyReturn.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmWarrantyReturn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===========================================================================
' frmWarrantyReturn - WARRANTY RETURN / RECEIVE PRODUCT
' Complete Integrated Code for GLOBAL SOFT v1.0
'===========================================================================

Option Explicit
Private Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
    ByVal hwnd As LongPtr, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, _
    ByVal lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As LongPtr

'=== MODULE VARIABLES ===
Private CurrentAssignID As String
Private CurrentEntryID As String
Private CurrentSerial As String
Private CurrentVendorID As String
Private CurrentVendorName As String
Private IsReturnSaved As Boolean
Private CurrentReturnPhotoPath As String
Private CurrentInvoicePhotoPath As String
Public parentForm As Object
Private CurrentCustomerID As String

'=== EDIT MODE VARIABLES ===
Private isEditMode As Boolean
Private EditReturnRow As Long
Private IsLoadingData As Boolean





'===========================================
' FORM INITIALIZE
'===========================================
Private Sub UserForm_Initialize()
    Me.caption = "Warranty Return - Receive Product"
    Me.Width = 1025
    Me.Height = 780
    Me.StartUpPosition = 1
    AddMinMaxButtons Me
    ' Reset variables
    CurrentAssignID = ""
    CurrentEntryID = ""
    CurrentSerial = ""
    CurrentVendorID = ""
    CurrentVendorName = ""
    IsReturnSaved = False
    CurrentReturnPhotoPath = ""
    CurrentInvoicePhotoPath = ""
    isEditMode = False
    EditReturnRow = 0
    IsLoadingData = False
    
    ' Setup Return Mode
    cmbReturnMode.Clear
    cmbReturnMode.AddItem "COURIER"
    cmbReturnMode.AddItem "BUS"
    cmbReturnMode.AddItem "HAND"
    cmbReturnMode.AddItem "TRANSPORT"
    cmbReturnMode.AddItem "INDIAN POST"
    
    ' Clear dependent fields
    cmbReturnName.Clear
    txtReturnMobile.value = ""
    txtReturnAddress.value = ""
    
    ' ===== PAYMENT LISTVIEW SETUP =====
    On Error Resume Next
    With ListView1
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
            chk.value = False
            chk.caption = ""
            chk.enabled = False
        End If
    Next chk
    
    ' Default settings
    optRepaired.value = True
    txtReturnDate.value = Format(Date, "dd-mm-yyyy")
    txtReturnCharges.value = ""
    UpdateStatusDetailsVisibility
    
    ' Load combo entries first
    Call LoadAssignedEntries
    
    ' ===== CHECK IF EDIT MODE (FROM DASHBOARD) =====
    Dim dbEntryID As String
    dbEntryID = Trim(Me.Tag)
    
    If dbEntryID <> "" Then
        ' ===== EDIT MODE =====
        isEditMode = True
        
        ' Add entryID to combo if not exists
        Dim found As Boolean, idx As Long
        found = False
        For idx = 0 To cmbEntryID.ListCount - 1
            If UCase(Trim(cmbEntryID.List(idx))) = UCase(dbEntryID) Then
                found = True
                Exit For
            End If
        Next idx
        If Not found Then cmbEntryID.AddItem dbEntryID
        
        ' ===== LOAD ALL DATA =====
        IsLoadingData = True
        
        ' Set EntryID - triggers cmbEntryID_Change
        cmbEntryID.value = dbEntryID
        
        ' Load return details
        Call LoadWarrantyReturnData(dbEntryID)
        
        IsLoadingData = False
        
        ' ===== SET EDIT MODE UI =====
        Me.caption = "Receive Warranty Product - EDIT (" & dbEntryID & ")"
        btnReceive.caption = "UPDATE PRODUCT"
        btnUpdateToCustomer.enabled = True
        
        Me.Tag = ""
    Else
        ' ===== NEW MODE =====
        isEditMode = False
        EditReturnRow = 0
        ClearForm
        IsReturnSaved = False
        btnUpdateToCustomer.enabled = False
    End If
    
End Sub

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
    
    '--- VALIDATION ---
    If cmbEntryID.value = "" Then
        MsgBox "Please select an Entry ID first!", vbExclamation, "GLOBAL SOFT"
        cmbEntryID.SetFocus
        Exit Sub
    End If
    
    If Not IsReturnSaved Then
        MsgBox "Please RECEIVE product first!" & vbCrLf & "Then click UPDATE TO CUSTOMER.", vbExclamation, "GLOBAL SOFT"
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
    customerMobile = Trim(lblCustomerMobile.caption & "")
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
    
    '--- STATUS ---
    If optRepaired.value Then
        statusText = "RECEIVED_REPAIRED"
    ElseIf optReplaced.value Then
        statusText = "RECEIVED_REPLACED"
    ElseIf optRejected.value Then
        statusText = "RECEIVED_REJECTED"
    Else
        statusText = "RECEIVED"
    End If
    
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
            If Trim(wsPay.Cells(i, 2).value & "") = Trim(cmbEntryID.value & "") Then
                If Trim(wsPay.Cells(i, 6).value & "") = "" Then  ' Active
                    
                    Dim pName As String, pCategory As String, pAmt As Double
                    
                    pName = Trim(wsPay.Cells(i, 10).value & "")
                    pCategory = UCase(Trim(wsPay.Cells(i, 9).value & ""))
                    pAmt = val(wsPay.Cells(i, 16).value)
                    
                    ' DISCUSSION: SIRF SHOW
                    If pCategory = "DISCUSSION" Then
                        discussionAmt = discussionAmt + pAmt
                    End If
                    
                    ' PLUS: PARTS + CHARGES
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
                    
                    ' MINUS: RECEIVED + DISCOUNT
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
    
    ' TOTAL DUE
    totalDue = (partsCost + chargesCost) - (receivedAmt + discountAmt)
    
    '--- BUILD MESSAGE ---
    messageText = "*" & companyName & "*" & vbCrLf
    If companyAddress <> "" Then messageText = messageText & companyAddress & vbCrLf
    If companyPhone <> "" Then messageText = messageText & "Phone: " & companyPhone & vbCrLf
    
    messageText = messageText & vbCrLf & "Dear " & custName & "," & vbCrLf & vbCrLf
    messageText = messageText & "Your warranty return product has been received." & vbCrLf & vbCrLf
    
    ' Product Details
    messageText = messageText & "*Product Details:*" & vbCrLf
    messageText = messageText & "Entry ID: " & cmbEntryID.value & vbCrLf
    messageText = messageText & "Product: " & lblProductValue.caption & vbCrLf
    messageText = messageText & "Company: " & lblCompanyValue.caption & vbCrLf
    messageText = messageText & "Model: " & lblModelValue.caption & vbCrLf
    messageText = messageText & "Serial: " & lblSerialValue.caption & vbCrLf
    messageText = messageText & vbCrLf
    
    ' Status
    messageText = messageText & "*Return Status:* " & statusText & vbCrLf & vbCrLf
    
    ' DISCUSSION AMOUNT
    If discussionAmt > 0 Then
        messageText = messageText & "*Discussion Amount:* Rs. " & Format(discussionAmt, "0.00") & vbCrLf
        messageText = messageText & "(Initial estimated charges)" & vbCrLf & vbCrLf
    End If
    
    ' PARTS & CHARGES
    If partsListText <> "" Then
        messageText = messageText & "*Parts Used & Charges:*" & vbCrLf
        messageText = messageText & partsListText & vbCrLf
        messageText = messageText & "------------------------" & vbCrLf
    End If
    
    ' PAYMENTS & DISCOUNTS
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
    
    ' TOTAL DUE
    If totalDue > 0 Then
        messageText = messageText & "*TOTAL AMOUNT DUE: Rs. " & Format(totalDue, "0.00") & "*" & vbCrLf & vbCrLf
    ElseIf totalDue < 0 Then
        messageText = messageText & "*REFUND DUE: Rs. " & Format(Abs(totalDue), "0.00") & "*" & vbCrLf & vbCrLf
    Else
        messageText = messageText & "*TOTAL AMOUNT DUE: Rs. 0.00 (ALL CLEAR)*" & vbCrLf & vbCrLf
    End If
    
    ' SERVICE NOTE
    messageText = messageText & "*Service Note:*" & vbCrLf
    messageText = messageText & "Your product is " & statusText & " under warranty." & vbCrLf
    messageText = messageText & "We will update you once the service is completed." & vbCrLf & vbCrLf
    
    ' FOOTER
    messageText = messageText & "Thank you for trusting *" & companyName & "*." & vbCrLf
    If companyPhone <> "" Then messageText = messageText & "For any query, contact us: " & companyPhone
    
    ' URL ENCODE
    messageText = Replace(messageText, " ", "%20")
    messageText = Replace(messageText, vbCrLf, "%0A")
    messageText = Replace(messageText, "&", "%26")
    messageText = Replace(messageText, "#", "%23")
    messageText = Replace(messageText, "=", "%3D")
    messageText = Replace(messageText, "*", "%2A")
    messageText = Replace(messageText, "+", "%2B")
    
    ' OPEN WHATSAPP
    waURL = "whatsapp://send?phone=91" & customerMobile & "&text=" & messageText
    ShellExecute 0, "Open", waURL, "", "", 1
    
    MsgBox "WhatsApp opened successfully!", vbInformation, "GLOBAL SOFT"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "WhatsApp Error: " & Err.Number & " - " & Err.Description, vbCritical, "GLOBAL SOFT"
End Sub

'===========================================
' LOAD ONLY WARRANTY (WAR) ENTRIES
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
    
    cmbEntryID.Clear
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    
    For i = 2 To lastRow
        Dim status As String
        status = UCase(Trim(ws.Cells(i, 32).value))
        Dim entryID As String
        entryID = UCase(Trim(ws.Cells(i, 2).value))
        
        ' SIRF WARRANTY entries jo vendor ko gaye hain ya received ho chuke hain
        If (status = "ASSIGNED" Or status = "SENT" Or status = "IN PROGRESS" Or _
            status = "PENDING" Or status = "RECEIVE_REPAIRED" Or status = "RECEIVE_REPLACED" Or _
            status = "RECEIVE_REJECTED") And Left(entryID, 3) = "WAR" Then
            cmbEntryID.AddItem ws.Cells(i, 2).value
        End If
    Next i
    
    If cmbEntryID.ListCount = 0 Then
        lblEntryHint.caption = "No pending warranty entries found!"
    Else
        lblEntryHint.caption = UCase(" Only SENT TO VENDOR entries show here")
    End If
End Sub

'===========================================
' LOAD WARRANTY RETURN DATA (EDIT MODE)
'===========================================
Private Sub LoadWarrantyReturnData(entryID As String)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Warranty_Return_Master")
    On Error GoTo 0
    
    If ws Is Nothing Then Exit Sub
    
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    EditReturnRow = 0
    
    ' Find LATEST return record by EntryID (Column 3)
    For i = lastRow To 2 Step -1
        If UCase(Trim(ws.Cells(i, 3).value & "")) = UCase(entryID) Then
            EditReturnRow = i
            
            ' Load Return Details (Columns 12-24)
            Dim retMode As String
            retMode = Trim(ws.Cells(i, 12).value & "")
            If retMode <> "" Then
                cmbReturnMode.value = retMode
                ' Names load karo pehle
                Call LoadReturnNames(retMode)
            End If
            
            ' Ab name set karo (list mein available hoga)
            cmbReturnName.value = Trim(ws.Cells(i, 13).value & "")
            txtReturnMobile.value = Trim(ws.Cells(i, 14).value & "")
            txtReturnAddress.value = Trim(ws.Cells(i, 15).value & "")
            txtReturnDocket.value = Trim(ws.Cells(i, 16).value & "")
            txtReturnDate.value = Trim(ws.Cells(i, 17).value & "")
            txtReturnCharges.value = Trim(ws.Cells(i, 18).value & "")
            
            ' Status (Column 19)
            Dim retStatus As String
            retStatus = UCase(Trim(ws.Cells(i, 19).value & ""))
            If retStatus = "RECEIVE_REPAIRED" Or retStatus = "REPAIRED" Then
                optRepaired.value = True
            ElseIf retStatus = "RECEIVE_REPLACED" Or retStatus = "REPLACED" Then
                optReplaced.value = True
            ElseIf retStatus = "RECEIVE_REJECTED" Or retStatus = "REJECTED" Then
                optRejected.value = True
            End If
            Call UpdateStatusDetailsVisibility
            
            ' Details
            txtVerifiedBy.value = Trim(ws.Cells(i, 20).value & "")
            txtNewSerial.value = Trim(ws.Cells(i, 21).value & "")
            txtRejectReason.value = Trim(ws.Cells(i, 22).value & "")
            txtInvoiceNumber.value = Trim(ws.Cells(i, 23).value & "")
            txtInvoiceAmount.value = Trim(ws.Cells(i, 24).value & "")
            
            ' Photos
            CurrentReturnPhotoPath = Trim(ws.Cells(i, 25).value & "")
            CurrentInvoicePhotoPath = Trim(ws.Cells(i, 26).value & "")
            
            If CurrentReturnPhotoPath <> "" Then
                On Error Resume Next
                Set imgReturnPhoto.Picture = LoadPicture(CurrentReturnPhotoPath)
                On Error GoTo 0
            End If
            
            IsReturnSaved = True
            
            Exit For
        End If
    Next i
    
    If EditReturnRow = 0 Then
        ' No return record found - treat as new return for this assignment
        IsReturnSaved = False
    End If
End Sub

'===========================================
' ENTRY ID CHANGE - LOAD ALL DETAILS
'===========================================
Private Sub cmbEntryID_Change()
    Dim wsAssign As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim entryID As String
    
    entryID = Trim(cmbEntryID.value)
    If entryID = "" Then
        ClearDetails
        Exit Sub
    End If
    
    On Error Resume Next
    Set wsAssign = ThisWorkbook.Sheets("Assign_Master")
    On Error GoTo 0
    
    If wsAssign Is Nothing Then Exit Sub
    
    lastRow = wsAssign.Cells(wsAssign.Rows.count, 1).End(xlUp).row
    
    For i = 2 To lastRow
        If UCase(Trim(wsAssign.Cells(i, 2).value)) = UCase(entryID) Then
            
            ' Store values
            CurrentAssignID = wsAssign.Cells(i, 1).value
            CurrentEntryID = wsAssign.Cells(i, 2).value
            CurrentSerial = wsAssign.Cells(i, 9).value
            CurrentVendorName = wsAssign.Cells(i, 14).value
            CurrentCustomerID = wsAssign.Cells(i, 3).value
            
            ' Assignment Details
            lblEntryIDValue.caption = wsAssign.Cells(i, 2).value
            lblProductValue.caption = wsAssign.Cells(i, 6).value
            lblCompanyValue.caption = wsAssign.Cells(i, 7).value
            lblModelValue.caption = wsAssign.Cells(i, 8).value
            lblSerialValue.caption = wsAssign.Cells(i, 9).value
            lblCustomerValue.caption = wsAssign.Cells(i, 5).value
            lblCustomerMobile.caption = wsAssign.Cells(i, 4).value
            lblVendorValue.caption = wsAssign.Cells(i, 14).value
            lblVendorMobile.caption = wsAssign.Cells(i, 15).value
            
            ' Dates
            If IsDate(wsAssign.Cells(i, 26).value) Then
                lblSentDateValue.caption = Format(wsAssign.Cells(i, 26).value, "dd-mm-yyyy")
            Else
                lblSentDateValue.caption = wsAssign.Cells(i, 26).value
            End If
            
            If IsDate(wsAssign.Cells(i, 27).value) Then
                lblExpectedValue.caption = Format(wsAssign.Cells(i, 27).value, "dd-mm-yyyy")
            Else
                lblExpectedValue.caption = wsAssign.Cells(i, 27).value
            End If
            
            ' Sent Via
            Dim sendMode As String, sendDetail As String
            sendMode = Trim(wsAssign.Cells(i, 23).value)
            sendDetail = Trim(wsAssign.Cells(i, 24).value)
            If sendMode <> "" And sendDetail <> "" Then
                lblSentVia.caption = "Sent via: " & sendMode & " - " & sendDetail
            ElseIf sendMode <> "" Then
                lblSentVia.caption = "Sent via: " & sendMode
            Else
                lblSentVia.caption = ""
            End If
            
            ' Serial Labels
            lblSameSerialValue.caption = wsAssign.Cells(i, 9).value
            lblOldSerialValue.caption = wsAssign.Cells(i, 9).value
            
            ' Load Photos
            Me.Repaint
            DoEvents
            LoadCustomerPhotoByID wsAssign.Cells(i, 3).value
            Me.Repaint
            DoEvents
            LoadProductPhoto entryID
            Me.Repaint
            DoEvents
            
            Exit For
        End If
    Next i
    
    ' Only clear return details if NOT in edit mode data load
    If Not isEditMode Then
        ' Clear Return Details on every new entry
        cmbReturnMode.value = ""
        cmbReturnName.Clear
        txtReturnMobile.value = ""
        txtReturnAddress.value = ""
        txtReturnDocket.value = ""
        txtReturnCharges.value = ""
        
        ' Reset Status
        optRepaired.value = True
        UpdateStatusDetailsVisibility
        
        IsReturnSaved = False
        btnUpdateToCustomer.enabled = False
    End If
    
    ' ===== LOAD PAYMENT LIST =====
    Call LoadPaymentList(cmbEntryID.value)
    
    ' ===== LOAD ACCESSORIES =====
    Call LoadAccessories(entryID)
End Sub

'===========================================
' LOAD CUSTOMER PHOTO
'===========================================
Private Sub LoadCustomerPhotoByID(customerID As String)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim photoName As String
    Dim folderPath As String
    Dim fullPath As String
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Customer_Master")
    On Error GoTo 0
    
    If ws Is Nothing Then
        SafeLoadPicture imgCustomer, ""
        Exit Sub
    End If
    
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    
    For i = 2 To lastRow
        If UCase(Trim(ws.Cells(i, 1).value)) = UCase(customerID) Then
            photoName = Trim(ws.Cells(i, 7).value)
            
            If photoName <> "" Then
                On Error Resume Next
                folderPath = ThisWorkbook.Sheets("Settings").Range("B3").value
                On Error GoTo 0
                
                If folderPath = "" Then
                    folderPath = ThisWorkbook.path & "\GLOBAL_SOFT_DATA\Images\Customers\"
                End If
                If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
                
                fullPath = folderPath & photoName
                SafeLoadPicture imgCustomer, fullPath
            Else
                SafeLoadPicture imgCustomer, ""
            End If
            Exit For
        End If
    Next i
End Sub

'===========================================
' LOAD PRODUCT PHOTO
'===========================================
Private Sub LoadProductPhoto(entryID As String)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim photoPath As String
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Job_Product")
    On Error GoTo 0
    
    If ws Is Nothing Then
        SafeLoadPicture imgProductBefore, ""
        Exit Sub
    End If
    
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    
    For i = 2 To lastRow
        If UCase(Trim(ws.Cells(i, 1).value)) = UCase(entryID) Then
            photoPath = Trim(ws.Cells(i, 19).value)
            
            If photoPath = "" Then
                photoPath = ThisWorkbook.path & "\GLOBAL_SOFT_DATA\Images\Products\" & entryID & "\Photo1.jpg"
            End If
            
            If Dir(photoPath) <> "" Then
                SafeLoadPicture imgProductBefore, photoPath
            Else
                photoPath = ThisWorkbook.path & "\GLOBAL_SOFT_DATA\Images\Products\" & entryID & ".jpg"
                If Dir(photoPath) <> "" Then
                    SafeLoadPicture imgProductBefore, photoPath
                Else
                    SafeLoadPicture imgProductBefore, ""
                End If
            End If
            Exit For
        End If
    Next i
End Sub

'===========================================
' SAFE LOAD PICTURE
'===========================================
Private Sub SafeLoadPicture(imgCtrl As Image, filePath As String)
    On Error Resume Next
    Set imgCtrl.Picture = Nothing
    DoEvents
    
    If Trim(filePath) = "" Then Exit Sub
    If Dir(filePath) = "" Then
        Set imgCtrl.Picture = Nothing
        Exit Sub
    End If
    
    imgCtrl.Picture = LoadPicture(filePath)
    
    If Err.Number <> 0 Then
        Set imgCtrl.Picture = Nothing
    End If
    On Error GoTo 0
End Sub

'===========================================
' OPTION BUTTONS CLICK
'===========================================
Private Sub optRepaired_Click()
    UpdateStatusDetailsVisibility
End Sub

Private Sub optReplaced_Click()
    UpdateStatusDetailsVisibility
End Sub

Private Sub optRejected_Click()
    UpdateStatusDetailsVisibility
End Sub

'===========================================
' UPDATE STATUS DETAILS - ENABLE/DISABLE
'===========================================
Private Sub UpdateStatusDetailsVisibility()
    
    ' Sabse pehle sab disable karo
    On Error Resume Next
    lblSameSerialLabel.enabled = False
    lblSameSerialValue.enabled = False
    txtVerifiedBy.enabled = False
    
    lblOldSerialLabel.enabled = False
    lblOldSerialValue.enabled = False
    lblNewSerialLabel.enabled = False
    txtNewSerial.enabled = False
    
    lblRejectReasonLabel.enabled = False
    txtRejectReason.enabled = False
    On Error GoTo 0
    
    ' Selected ke hisaab se enable karo
    If optRepaired.value Then
        lblSameSerialLabel.enabled = True
        lblSameSerialValue.enabled = True
        txtVerifiedBy.enabled = True
        lblSameSerialValue.caption = CurrentSerial
        txtVerifiedBy.value = ""
        txtNewSerial.value = ""
        txtRejectReason.value = ""
        
    ElseIf optReplaced.value Then
        lblOldSerialLabel.enabled = True
        lblOldSerialValue.enabled = True
        lblNewSerialLabel.enabled = True
        txtNewSerial.enabled = True
        lblOldSerialValue.caption = CurrentSerial
        txtNewSerial.value = ""
        txtVerifiedBy.value = ""
        txtRejectReason.value = ""
        
    ElseIf optRejected.value Then
        lblRejectReasonLabel.enabled = True
        txtRejectReason.enabled = True
        txtRejectReason.value = ""
        txtVerifiedBy.value = ""
        txtNewSerial.value = ""
    End If
End Sub

'===========================================
' SEARCH BUTTON
'===========================================
Private Sub btnSearch_Click()
    Dim searchText As String
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim found As Boolean
    
    searchText = Trim(txtSearch.value)
    If searchText = "" Then
        MsgBox "Please enter search text!", vbExclamation
        txtSearch.SetFocus
        Exit Sub
    End If
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Assign_Master")
    On Error GoTo 0
    
    If ws Is Nothing Then Exit Sub
    
    found = False
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    
    For i = 2 To lastRow
        Dim status As String
        status = UCase(Trim(ws.Cells(i, 32).value))
        Dim entryID As String
        entryID = UCase(Trim(ws.Cells(i, 2).value))
        
        If (status = "ASSIGNED" Or status = "SENT" Or status = "IN PROGRESS" Or _
            status = "PENDING" Or status = "RECEIVE_REPAIRED" Or status = "RECEIVE_REPLACED" Or _
            status = "RECEIVE_REJECTED") And Left(entryID, 3) = "WAR" Then
            
            If InStr(1, UCase(Trim(ws.Cells(i, 2).value)), UCase(searchText), vbTextCompare) > 0 Or _
               InStr(1, UCase(Trim(ws.Cells(i, 5).value)), UCase(searchText), vbTextCompare) > 0 Or _
               InStr(1, UCase(Trim(ws.Cells(i, 9).value)), UCase(searchText), vbTextCompare) > 0 Or _
               InStr(1, UCase(Trim(ws.Cells(i, 14).value)), UCase(searchText), vbTextCompare) > 0 Then
                
                cmbEntryID.value = ws.Cells(i, 2).value
                DoEvents
                cmbEntryID_Change
                found = True
                Exit For
            End If
        End If
    Next i
    
    If Not found Then
        MsgBox "No pending warranty entry found!", vbInformation
    End If
End Sub

'===========================================
' RETURN MODE CHANGE - LOAD NAMES
'===========================================
Private Sub cmbReturnMode_Change()
    ' ===== AGAR DATA LOAD HORAH EDIT MODE MEIN, TOH NAHI =====
    If IsLoadingData Then Exit Sub
    
    Dim mode As String
    mode = Trim(cmbReturnMode.value)
    
    If mode = "" Then
        cmbReturnName.Clear
        txtReturnMobile.value = ""
        txtReturnAddress.value = ""
        Exit Sub
    End If
    
    LoadReturnNames mode
End Sub

'===========================================
' LOAD RETURN NAMES
'===========================================
Private Sub LoadReturnNames(mode As String)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Courier_Data")
    On Error GoTo 0
    
    If ws Is Nothing Then Exit Sub
    
    cmbReturnName.Clear
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    
    For i = 2 To lastRow
        If UCase(Trim(ws.Cells(i, 1).value)) = UCase(mode) Then
            On Error Resume Next
            cmbReturnName.AddItem ws.Cells(i, 2).value
            On Error GoTo 0
        End If
    Next i
End Sub

'===========================================
' RETURN NAME CHANGE - LOAD MOBILE & ADDRESS
'===========================================
Private Sub cmbReturnName_Change()
    ' ===== AGAR DATA LOAD HORAH EDIT MODE MEIN, TOH NAHI =====
    If IsLoadingData Then Exit Sub
    
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim mode As String
    Dim name As String
    
    mode = Trim(cmbReturnMode.value)
    name = Trim(cmbReturnName.value)
    
    If name = "" Then
        txtReturnMobile.value = ""
        txtReturnAddress.value = ""
        Exit Sub
    End If
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Courier_Data")
    On Error GoTo 0
    
    If ws Is Nothing Then Exit Sub
    
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    
    For i = 2 To lastRow
        If UCase(Trim(ws.Cells(i, 1).value)) = UCase(mode) And _
           UCase(Trim(ws.Cells(i, 2).value)) = UCase(name) Then
            txtReturnMobile.value = ws.Cells(i, 3).value
            txtReturnAddress.value = ws.Cells(i, 4).value
            Exit For
        End If
    Next i
End Sub

'===========================================
' ADD RETURN MODE
'===========================================
Private Sub btnAddReturnMode_Click()
    Dim newMode As String
    newMode = InputBox("Enter New Return Mode:", "Add Return Mode")
    If Trim(newMode) = "" Then Exit Sub
    
    Dim i As Long
    For i = 0 To cmbReturnMode.ListCount - 1
        If UCase(cmbReturnMode.List(i)) = UCase(newMode) Then
            MsgBox "Mode '" & newMode & "' already exists!", vbExclamation
            Exit Sub
        End If
    Next i
    
    cmbReturnMode.AddItem UCase(newMode)
    cmbReturnMode.value = UCase(newMode)
    MsgBox "Return Mode Added: " & newMode, vbInformation
End Sub

'===========================================
' ADD RETURN NAME
'===========================================
Private Sub btnAddReturnName_Click()
    Dim mode As String
    Dim newName As String
    Dim newMobile As String
    Dim newAddress As String
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    mode = Trim(cmbReturnMode.value)
    If mode = "" Then
        MsgBox "Please select Return Mode first!", vbExclamation
        Exit Sub
    End If
    
    newName = InputBox("Enter " & mode & " Name:", "Add " & mode)
    If Trim(newName) = "" Then Exit Sub
    
    newMobile = InputBox("Enter Mobile Number:", "Add " & mode)
    If Trim(newMobile) = "" Then Exit Sub
    
    newAddress = InputBox("Enter Address:", "Add " & mode)
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Courier_Data")
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
        ws.name = "Courier_Data"
        ws.Cells(1, 1).value = "Mode"
        ws.Cells(1, 2).value = "Name"
        ws.Cells(1, 3).value = "Mobile"
        ws.Cells(1, 4).value = "Address"
        With ws.Rows(1)
            .Font.Bold = True
            .Interior.Color = RGB(200, 200, 200)
        End With
    End If
    
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    
    For i = 2 To lastRow
        If UCase(Trim(ws.Cells(i, 1).value)) = UCase(mode) And _
           UCase(Trim(ws.Cells(i, 2).value)) = UCase(newName) Then
            MsgBox "'" & newName & "' already exists!", vbExclamation
            Exit Sub
        End If
    Next i
    
    lastRow = lastRow + 1
    ws.Cells(lastRow, 1).value = UCase(mode)
    ws.Cells(lastRow, 2).value = UCase(newName)
    ws.Cells(lastRow, 3).value = newMobile
    ws.Cells(lastRow, 4).value = newAddress
    
    MsgBox UCase(mode) & " '" & UCase(newName) & "' Added!", vbInformation
    
    LoadReturnNames mode
    cmbReturnName.value = UCase(newName)
End Sub

'===========================================
' RETURN PHOTO CAPTURE
'===========================================
Private Sub btnReturnPhoto_Click()
    On Error GoTo ErrorHandler
    
    If cmbEntryID.value = "" Then
        MsgBox "Please select Entry ID first!", vbExclamation
        Exit Sub
    End If
    
    Dim currentStatus As String
    If optRepaired.value Then
        currentStatus = "REPAIRED"
    ElseIf optReplaced.value Then
        currentStatus = "REPLACED"
    ElseIf optRejected.value Then
        currentStatus = "REJECTED"
    Else
        MsgBox "Please select Product Status first!", vbExclamation
        Exit Sub
    End If
    
    Dim basePath As String
    Dim folderPath As String
    Dim photoPath As String
    Dim fileName As String
    
    basePath = ThisWorkbook.path & "\GLOBAL_SOFT_DATA\Images\"
    
    If Dir(basePath, vbDirectory) = "" Then MkDir basePath
    If Dir(basePath & "Return\", vbDirectory) = "" Then MkDir basePath & "Return\"
    If Dir(basePath & "Return\Receive\", vbDirectory) = "" Then MkDir basePath & "Return\Receive\"
    If Dir(basePath & "Return\Receive\" & CurrentEntryID & "\", vbDirectory) = "" Then
        MkDir basePath & "Return\Receive\" & CurrentEntryID & "\"
    End If
    
    folderPath = basePath & "Return\Receive\" & CurrentEntryID & "\"
    fileName = CurrentEntryID & "_" & currentStatus & "_" & Format(Now, "yyyymmdd_hhmmss") & ".jpg"
    photoPath = folderPath & fileName
    
    Dim lastFile As String
    lastFile = GetLatestCameraPhoto()
    
    shell "explorer.exe shell:AppsFolder\Microsoft.WindowsCamera_8wekyb3d8bbwe!App", vbNormalFocus
    
    If MsgBox("Photo capture karke Camera close karo, phir OK dabao.", vbOKCancel + vbInformation) <> vbOK Then Exit Sub
    
    Dim newFile As String
    newFile = GetLatestCameraPhoto()
    
    If newFile = "" Or newFile = lastFile Then
        MsgBox "Photo detect nahi hua!", vbExclamation
        Exit Sub
    End If
    
    If Dir(photoPath) <> "" Then Kill photoPath
    FileCopy newFile, photoPath
    
    On Error Resume Next
    Set imgReturnPhoto.Picture = LoadPicture(photoPath)
    On Error GoTo ErrorHandler
    
    CurrentReturnPhotoPath = photoPath
    
    Exit Sub
ErrorHandler:
    MsgBox "Camera Error: " & Err.Description, vbCritical
End Sub

'===========================================
' GET LATEST CAMERA PHOTO
'===========================================
Private Function GetLatestCameraPhoto() As String
    On Error Resume Next
    
    Dim fso As Object
    Dim folder As Object
    Dim file As Object
    Dim latestDate As Date
    Dim camPath As String
    
    GetLatestCameraPhoto = ""
    latestDate = #1/1/1900#
    
    camPath = Environ("USERPROFILE") & "\Pictures\Camera Roll\"
    If Dir(camPath, vbDirectory) = "" Then
        camPath = Environ("USERPROFILE") & "\OneDrive\Pictures\Camera Roll\"
    End If
    If Dir(camPath, vbDirectory) = "" Then Exit Function
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(camPath)
    
    For Each file In folder.files
        If LCase(Right(file.name, 4)) = ".jpg" Then
            If file.DateLastModified > latestDate Then
                latestDate = file.DateLastModified
                GetLatestCameraPhoto = file.path
            End If
        End If
    Next
    
    Set fso = Nothing
    On Error GoTo 0
End Function

'===========================================
' RECEIVE PRODUCT - MAIN SAVE / UPDATE
'===========================================
Private Sub btnReceive_Click()
    On Error GoTo ErrorHandler
    
    ' Validation
    If cmbEntryID.value = "" Then
        MsgBox "Please select Entry ID!", vbExclamation
        cmbEntryID.SetFocus
        Exit Sub
    End If
    
    If cmbReturnMode.value = "" Then
        MsgBox "Please select Return Mode!", vbExclamation
        cmbReturnMode.SetFocus
        Exit Sub
    End If
    
    If cmbReturnName.value = "" Then
        MsgBox "Please select Return Name!", vbExclamation
        cmbReturnName.SetFocus
        Exit Sub
    End If
    
    If Trim(txtReturnDate.value) = "" Then
        MsgBox "Please enter Return Date!", vbExclamation
        txtReturnDate.SetFocus
        Exit Sub
    End If
    
    ' Status based validation
    If optRepaired.value And Trim(txtVerifiedBy.value) = "" Then
        MsgBox "Please enter Verified By name!", vbExclamation
        txtVerifiedBy.SetFocus
        Exit Sub
    End If
    
    If optReplaced.value And Trim(txtNewSerial.value) = "" Then
        MsgBox "Please enter New Serial Number!", vbExclamation
        txtNewSerial.SetFocus
        Exit Sub
    End If
    
    If optRejected.value And Trim(txtRejectReason.value) = "" Then
        MsgBox "Please enter Rejection Reason!", vbExclamation
        txtRejectReason.SetFocus
        Exit Sub
    End If
    
    ' Determine Status
    Dim returnStatus As String
    If optRepaired.value Then
        returnStatus = "RECEIVE_REPAIRED"
    ElseIf optReplaced.value Then
        returnStatus = "RECEIVE_REPLACED"
    ElseIf optRejected.value Then
        returnStatus = "RECEIVE_REJECTED"
    End If
    
    '===========================================
    ' CHECK ACCESSORIES DUE - BEFORE SAVE
    '===========================================
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
        If MsgBox("WARNING! These accessories are DUE from vendor:" & vbCrLf & vbCrLf & _
                  dueItems & vbCrLf & _
                  "Do you still want to RECEIVE product?", vbExclamation + vbYesNo, "ACCESSORIES DUE - GLOBAL SOFT") = vbNo Then
            Exit Sub
        End If
    End If
    
    ' ===== EDIT MODE = UPDATE EXISTING =====
    If isEditMode And EditReturnRow > 0 Then
        Call UpdateExistingReturn(returnStatus)
        Exit Sub
    End If
    
    ' ===== NEW MODE = SAVE NEW =====
    
    ' Save to Warranty_Return_Master
    Dim wsReturn As Worksheet
    Dim lastRow As Long
    Dim returnID As String
    
    On Error Resume Next
    Set wsReturn = ThisWorkbook.Sheets("Warranty_Return_Master")
    On Error GoTo 0
    
    If wsReturn Is Nothing Then
        Set wsReturn = ThisWorkbook.Sheets.Add
        wsReturn.name = "Warranty_Return_Master"
        With wsReturn.Rows(1)
            .Cells(1, 1).value = "ReturnID"
            .Cells(1, 2).value = "AssignID"
            .Cells(1, 3).value = "EntryID"
            .Cells(1, 4).value = "CustomerID"
            .Cells(1, 5).value = "CustomerName"
            .Cells(1, 6).value = "ProductType"
            .Cells(1, 7).value = "Company"
            .Cells(1, 8).value = "Model"
            .Cells(1, 9).value = "OriginalSerial"
            .Cells(1, 10).value = "VendorName"
            .Cells(1, 11).value = "VendorMobile"
            .Cells(1, 12).value = "ReturnMode"
            .Cells(1, 13).value = "ReturnName"
            .Cells(1, 14).value = "ReturnMobile"
            .Cells(1, 15).value = "ReturnAddress"
            .Cells(1, 16).value = "DocketNumber"
            .Cells(1, 17).value = "ReturnDate"
            .Cells(1, 18).value = "ReturnCharges"
            .Cells(1, 19).value = "ReturnStatus"
            .Cells(1, 20).value = "VerifiedBy"
            .Cells(1, 21).value = "NewSerial"
            .Cells(1, 22).value = "RejectReason"
            .Cells(1, 23).value = "InvoiceNumber"
            .Cells(1, 24).value = "InvoiceAmount"
            .Cells(1, 25).value = "ReturnPhotoPath"
            .Cells(1, 26).value = "InvoicePhotoPath"
            .Cells(1, 27).value = "CreatedDate"
            .Cells(1, 28).value = "AccessoriesStatus"
            .Cells(1, 29).value = "EditHistory"
            .Font.Bold = True
            .Interior.Color = RGB(200, 200, 200)
        End With
    End If
    
    lastRow = wsReturn.Cells(wsReturn.Rows.count, 1).End(xlUp).row + 1
    returnID = "WRT" & Format(lastRow - 1, "00000")
    
    ' Get CustomerID
    Dim wsAssign As Worksheet
    Dim assignLastRow As Long
    Dim j As Long
    Dim customerID As String
    
    Set wsAssign = ThisWorkbook.Sheets("Assign_Master")
    assignLastRow = wsAssign.Cells(wsAssign.Rows.count, 1).End(xlUp).row
    
    For j = 2 To assignLastRow
        If UCase(Trim(wsAssign.Cells(j, 2).value)) = UCase(CurrentEntryID) Then
            customerID = wsAssign.Cells(j, 3).value
            Exit For
        End If
    Next j
    
    ' Write Data
    With wsReturn.Rows(lastRow)
        .Cells(1, 1).value = returnID
        .Cells(1, 2).value = CurrentAssignID
        .Cells(1, 3).value = CurrentEntryID
        .Cells(1, 4).value = customerID
        .Cells(1, 5).value = lblCustomerValue.caption
        .Cells(1, 6).value = lblProductValue.caption
        .Cells(1, 7).value = lblCompanyValue.caption
        .Cells(1, 8).value = lblModelValue.caption
        .Cells(1, 9).value = lblSerialValue.caption
        .Cells(1, 10).value = lblVendorValue.caption
        .Cells(1, 11).value = lblVendorMobile.caption
        .Cells(1, 12).value = cmbReturnMode.value
        .Cells(1, 13).value = cmbReturnName.value
        .Cells(1, 14).value = txtReturnMobile.value
        .Cells(1, 15).value = txtReturnAddress.value
        .Cells(1, 16).value = txtReturnDocket.value
        .Cells(1, 17).value = txtReturnDate.value
        .Cells(1, 18).value = val(txtReturnCharges.value)
        .Cells(1, 19).value = returnStatus
        .Cells(1, 20).value = txtVerifiedBy.value
        .Cells(1, 21).value = txtNewSerial.value
        .Cells(1, 22).value = txtRejectReason.value
        .Cells(1, 23).value = txtInvoiceNumber.value
        .Cells(1, 24).value = val(txtInvoiceAmount.value)
        .Cells(1, 25).value = CurrentReturnPhotoPath
        .Cells(1, 26).value = CurrentInvoicePhotoPath
        .Cells(1, 27).value = Format(Now, "dd-MM-yyyy hh:mm:ss")
        .Cells(1, 28).value = GetAccessoriesStatus()
    End With
    
    ' Update Assign_Master
    For j = 2 To assignLastRow
        If UCase(Trim(wsAssign.Cells(j, 2).value)) = UCase(CurrentEntryID) Then
            wsAssign.Cells(j, 32).value = returnStatus
            wsAssign.Cells(j, 35).value = Format(Now, "dd-MM-yyyy hh:mm:ss")
            Exit For
        End If
    Next j
    
        ' Update Job_Product
    Dim wsProduct As Worksheet
    Dim productLastRow As Long
    
    On Error Resume Next
    Set wsProduct = ThisWorkbook.Sheets("Job_Product")
    On Error GoTo ErrorHandler
    
    If Not wsProduct Is Nothing Then
        productLastRow = wsProduct.Cells(wsProduct.Rows.count, 1).End(xlUp).row
        For j = 2 To productLastRow
            If UCase(Trim(wsProduct.Cells(j, 1).value)) = UCase(CurrentEntryID) Then
                If optReplaced.value And Trim(txtNewSerial.value) <> "" Then
                    wsProduct.Cells(j, 9).value = txtNewSerial.value
                End If
                wsProduct.Cells(j, 4).value = returnStatus   '? ???? Status Update ?? ??? ??
                Exit For
            End If
        Next j
    End If
    
    ' Success
    IsReturnSaved = True
    btnUpdateToCustomer.enabled = True
    btnReceive.enabled = False
    
    MsgBox "Product Received Successfully!" & vbCrLf & vbCrLf & _
           "Return ID: " & returnID & vbCrLf & _
           "Entry ID: " & CurrentEntryID & vbCrLf & _
           "Status: " & returnStatus & vbCrLf & vbCrLf & _
           "Now click 'UPDATE TO CUSTOMER' to send WhatsApp.", vbInformation, "GLOBAL SOFT"
    
    Exit Sub
ErrorHandler:
    MsgBox "Error: " & Err.Number & " - " & Err.Description, vbCritical
End Sub

'===========================================
' UPDATE EXISTING RETURN (EDIT MODE)
'===========================================
Private Sub UpdateExistingReturn(returnStatus As String)
    Dim wsReturn As Worksheet
    Dim wsAssign As Worksheet
    Dim wsProduct As Worksheet
    Dim assignLastRow As Long
    Dim j As Long
    
    On Error Resume Next
    Set wsReturn = ThisWorkbook.Sheets("Warranty_Return_Master")
    On Error GoTo 0
    
    If wsReturn Is Nothing Or EditReturnRow = 0 Then
        MsgBox "Cannot update - no existing return record found!", vbExclamation
        Exit Sub
    End If
    
    ' Update existing row
    With wsReturn.Rows(EditReturnRow)
        .Cells(1, 12).value = cmbReturnMode.value
        .Cells(1, 13).value = cmbReturnName.value
        .Cells(1, 14).value = txtReturnMobile.value
        .Cells(1, 15).value = txtReturnAddress.value
        .Cells(1, 16).value = txtReturnDocket.value
        .Cells(1, 17).value = txtReturnDate.value
        .Cells(1, 18).value = val(txtReturnCharges.value)
        .Cells(1, 19).value = returnStatus
        .Cells(1, 20).value = txtVerifiedBy.value
        .Cells(1, 21).value = txtNewSerial.value
        .Cells(1, 22).value = txtRejectReason.value
        .Cells(1, 23).value = txtInvoiceNumber.value
        .Cells(1, 24).value = val(txtInvoiceAmount.value)
        .Cells(1, 25).value = CurrentReturnPhotoPath
        .Cells(1, 26).value = CurrentInvoicePhotoPath
        .Cells(1, 28).value = GetAccessoriesStatus()
        .Cells(1, 29).value = Format(Now, "dd-MM-yyyy hh:mm:ss") & " EDITED by " & Environ("Username")
    End With
    
    ' Update Assign_Master
    Set wsAssign = ThisWorkbook.Sheets("Assign_Master")
    assignLastRow = wsAssign.Cells(wsAssign.Rows.count, 1).End(xlUp).row
    For j = 2 To assignLastRow
        If UCase(Trim(wsAssign.Cells(j, 2).value)) = UCase(CurrentEntryID) Then
            wsAssign.Cells(j, 32).value = returnStatus
            wsAssign.Cells(j, 35).value = Format(Now, "dd-MM-yyyy hh:mm:ss") & " (EDIT)"
            Exit For
        End If
    Next j
    
        ' ===== FIX: Job_Product ??? ?? Status Update ??? =====
    Dim wsProduct As Worksheet
    Dim productLastRow As Long
    
    On Error Resume Next
    Set wsProduct = ThisWorkbook.Sheets("Job_Product")
    On Error GoTo 0
    
    If Not wsProduct Is Nothing Then
        productLastRow = wsProduct.Cells(wsProduct.Rows.count, 1).End(xlUp).row
        For j = 2 To productLastRow
            If UCase(Trim(wsProduct.Cells(j, 1).value)) = UCase(CurrentEntryID) Then
                If optReplaced.value And Trim(txtNewSerial.value) <> "" Then
                    wsProduct.Cells(j, 9).value = txtNewSerial.value
                End If
                wsProduct.Cells(j, 4).value = returnStatus
                wsProduct.Cells(j, 5).value = Format(Now, "dd-MM-yyyy hh:mm:ss")  'Date ?? Update
                Exit For
            End If
        Next j
    End If
    
    
    IsReturnSaved = True
    btnUpdateToCustomer.enabled = True
    
    MsgBox "Return Updated Successfully!" & vbCrLf & _
           "Entry ID: " & CurrentEntryID & vbCrLf & _
           "Status: " & returnStatus, vbInformation, "UPDATE COMPLETE"
End Sub

'===========================================
' KEEP PENDING
'===========================================
Private Sub btnKeepPending_Click()
    On Error GoTo ErrorHandler
    
    If cmbEntryID.value = "" Then
        MsgBox "Please select an Entry!", vbExclamation
        cmbEntryID.SetFocus
        Exit Sub
    End If
    
    Dim wsAssign As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    On Error Resume Next
    Set wsAssign = ThisWorkbook.Sheets("Assign_Master")
    On Error GoTo ErrorHandler
    
    If wsAssign Is Nothing Then
        MsgBox "Assign_Master sheet not found!", vbExclamation
        Exit Sub
    End If
    
    lastRow = wsAssign.Cells(wsAssign.Rows.count, 1).End(xlUp).row
    
    For i = 2 To lastRow
        If UCase(Trim(wsAssign.Cells(i, 2).value)) = UCase(CurrentEntryID) Then
            wsAssign.Cells(i, 32).value = "PENDING"
            wsAssign.Cells(i, 31).value = "Keep Pending - " & Format(Now, "dd-MM-yyyy")
            Exit For
        End If
    Next i
    
    ' Job_Product mein bhi PENDING
    Dim wsProduct As Worksheet
    On Error Resume Next
    Set wsProduct = ThisWorkbook.Sheets("Job_Product")
    On Error GoTo 0
    
    If Not wsProduct Is Nothing Then
        For i = 2 To wsProduct.Cells(wsProduct.Rows.count, 1).End(xlUp).row
            If UCase(Trim(wsProduct.Cells(i, 1).value)) = UCase(CurrentEntryID) Then
                wsProduct.Cells(i, 4).value = "PENDING"
                wsProduct.Cells(i, 5).value = Format(Now, "dd-MM-yyyy hh:mm:ss")
                Exit For
            End If
        Next i
    End If
    
    MsgBox "Entry kept as PENDING!" & vbCrLf & "Entry ID: " & CurrentEntryID, vbInformation
    
    ClearForm
    LoadAssignedEntries
    
    Exit Sub
ErrorHandler:
    MsgBox "Error: " & Err.Number & " - " & Err.Description, vbCritical
End Sub

'===========================================
' CANCEL
'===========================================
Private Sub btnCancel_Click()
    Unload Me
End Sub

'===========================================
' CLEAR FORM
'===========================================
Private Sub ClearForm()
    cmbEntryID.value = ""
    txtSearch.value = ""
    
    ' Assignment Details
    lblEntryIDValue.caption = ""
    lblProductValue.caption = ""
    lblCompanyValue.caption = ""
    lblModelValue.caption = ""
    lblSerialValue.caption = ""
    lblCustomerValue.caption = ""
    lblCustomerMobile.caption = ""
    lblVendorValue.caption = ""
    lblVendorMobile.caption = ""
    lblSentDateValue.caption = ""
    lblSentVia.caption = ""
    lblExpectedValue.caption = ""
    lblSameSerialValue.caption = ""
    lblOldSerialValue.caption = ""
    
    ' Return Details
    cmbReturnMode.value = ""
    cmbReturnName.Clear
    txtReturnMobile.value = ""
    txtReturnAddress.value = ""
    txtReturnDocket.value = ""
    txtReturnDate.value = Format(Date, "dd-mm-yyyy")
    txtReturnCharges.value = ""
    
    ' Product Status
    optRepaired.value = True
    UpdateStatusDetailsVisibility
    
    ' Status Details
    txtVerifiedBy.value = ""
    txtNewSerial.value = ""
    txtRejectReason.value = ""
    
    ' Bill Details
    txtInvoiceNumber.value = ""
    txtInvoiceAmount.value = ""
    
    ' Photos
    Set imgReturnPhoto.Picture = Nothing
    CurrentReturnPhotoPath = ""
    On Error Resume Next
    CurrentInvoicePhotoPath = ""
    On Error GoTo 0
    
    SafeLoadPicture imgCustomer, ""
    SafeLoadPicture imgProductBefore, ""
    
    ' Variables
    CurrentAssignID = ""
    CurrentEntryID = ""
    CurrentSerial = ""
    CurrentVendorID = ""
    CurrentVendorName = ""
    IsReturnSaved = False
    isEditMode = False
    EditReturnRow = 0
    
    ' ===== CLEAR PAYMENT LIST =====
    ListView1.ListItems.Clear
    Call UpdatePaymentButtons
    
    ' ===== CLEAR ACCESSORIES =====
    Dim chk As Control
    For Each chk In Me.Controls
        If TypeName(chk) = "CheckBox" Then
            chk.value = False
            chk.caption = ""
            chk.enabled = False
            chk.visible = False
        End If
    Next chk
    
    btnReceive.caption = "RECEIVE PRODUCT"
    btnReceive.enabled = True
    btnUpdateToCustomer.enabled = False
End Sub

'===========================================
' CLEAR DETAILS
'===========================================
Private Sub ClearDetails()
    lblEntryIDValue.caption = ""
    lblProductValue.caption = ""
    lblCompanyValue.caption = ""
    lblModelValue.caption = ""
    lblSerialValue.caption = ""
    lblCustomerValue.caption = ""
    lblCustomerMobile.caption = ""
    lblVendorValue.caption = ""
    lblVendorMobile.caption = ""
    lblSentDateValue.caption = ""
    lblSentVia.caption = ""
    lblExpectedValue.caption = ""
    lblSameSerialValue.caption = ""
    lblOldSerialValue.caption = ""
    
    SafeLoadPicture imgCustomer, ""
    SafeLoadPicture imgProductBefore, ""
    
    CurrentAssignID = ""
    CurrentEntryID = ""
    CurrentSerial = ""
End Sub

'===========================================
' DATE VALIDATION
'===========================================
Private Sub txtReturnDate_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    On Error Resume Next
    If txtReturnDate.value <> "" Then
        txtReturnDate.value = Format(CDate(txtReturnDate.value), "dd-mm-yyyy")
    End If
    On Error GoTo 0
End Sub

'===========================================
' NUMBERS ONLY VALIDATION
'===========================================
Private Sub txtReturnCharges_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If Not (KeyAscii >= 48 And KeyAscii <= 57) And KeyAscii <> 8 And KeyAscii <> 46 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtInvoiceAmount_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If Not (KeyAscii >= 48 And KeyAscii <= 57) And KeyAscii <> 8 And KeyAscii <> 46 Then
        KeyAscii = 0
    End If
End Sub

'===========================================
' SEARCH ON ENTER KEY
'===========================================
Private Sub txtSearch_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then
        KeyAscii = 0
        btnSearch_Click
    End If
End Sub

'===========================================
' FORM CLOSE CONFIRMATION
'===========================================
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        If cmbEntryID.value <> "" And Not IsReturnSaved Then
            If MsgBox("Return entry not saved! Close anyway?", vbQuestion + vbYesNo) = vbNo Then
                Cancel = 1
            End If
        End If
    End If
End Sub

'===========================================
' WHATSAPP MESSAGE BUTTON
'===========================================
Private Sub btnWhatsApp_Click()
    Dim customerMobile As String
    Dim statusText As String
    Dim messageText As String
    Dim waURL As String
    Dim custName As String
    
    On Error GoTo ErrorHandler
    
    '--- VALIDATION ---
    If cmbEntryID.value = "" Then
        MsgBox "Please select Entry ID first!", vbExclamation, "GLOBAL SOFT"
        cmbEntryID.SetFocus
        Exit Sub
    End If
    
    ' Customer mobile check
    customerMobile = Trim(lblCustomerMobile.caption)
    If customerMobile = "" Then
        MsgBox "Customer mobile number not found!" & vbCrLf & "Please check Customer Master.", vbExclamation, "GLOBAL SOFT"
        Exit Sub
    End If
    
    '--- CLEAN MOBILE NUMBER ---
    customerMobile = Replace(customerMobile, " ", "")
    customerMobile = Replace(customerMobile, "-", "")
    customerMobile = Replace(customerMobile, "(", "")
    customerMobile = Replace(customerMobile, ")", "")
    customerMobile = Replace(customerMobile, "+", "")
    
    If Len(customerMobile) = 10 Then
        customerMobile = "91" & customerMobile
    ElseIf Left(customerMobile, 1) = "0" And Len(customerMobile) = 11 Then
        customerMobile = "91" & Mid(customerMobile, 2)
    End If
    
    If Len(customerMobile) <> 12 Or Left(customerMobile, 2) <> "91" Then
        MsgBox "Invalid mobile number format!" & vbCrLf & "Found: " & customerMobile, vbExclamation, "GLOBAL SOFT"
        Exit Sub
    End If
    
    '--- DETERMINE STATUS ---
    If optRepaired.value Then
        statusText = "REPAIRED "
    ElseIf optReplaced.value Then
        statusText = "REPLACED "
    ElseIf optRejected.value Then
        statusText = "REJECTED "
    Else
        statusText = "COMPLETED "
    End If
    
    '--- BUILD MESSAGE ---
    custName = Trim(lblCustomerValue.caption)
    If custName = "" Then custName = "Customer"
    
    messageText = "Dear " & custName & "," & vbCrLf & vbCrLf & _
        "Your product *" & lblProductValue.caption & "* (Entry: " & CurrentEntryID & ") has been *" & statusText & "*." & vbCrLf & vbCrLf & _
        "Please pick up your product within *7 days* from our service center." & vbCrLf & vbCrLf & _
        "Product: " & lblProductValue.caption & vbCrLf & _
        "Company: " & lblCompanyValue.caption & vbCrLf & _
        "Model: " & lblModelValue.caption & vbCrLf & vbCrLf & _
        "Thank you," & vbCrLf & _
        "*GLOBAL IT SOLUTIONS*"
    
    '--- URL ENCODE ---
    messageText = Replace(messageText, " ", "%20")
    messageText = Replace(messageText, vbCrLf, "%0A")
    messageText = Replace(messageText, "&", "%26")
    messageText = Replace(messageText, "#", "%23")
    messageText = Replace(messageText, "=", "%3D")
    messageText = Replace(messageText, "*", "%2A")
    
    '--- OPEN WHATSAPP ---
    waURL = "https://api.whatsapp.com/send?phone= " & customerMobile & "&text=" & messageText
    
    ShellExecute 0, "Open", waURL, "", "", 1
    
    MsgBox "WhatsApp Web/App opened!" & vbCrLf & vbCrLf & _
           "Customer: " & custName & vbCrLf & _
           "Mobile: +" & customerMobile & vbCrLf & _
           "Status: " & statusText, vbInformation, "GLOBAL SOFT - WhatsApp Ready"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "WhatsApp Error: " & Err.Number & " - " & Err.Description, vbCritical, "GLOBAL SOFT"
End Sub

'===========================================
' LOAD PAYMENT LIST
'===========================================
Public Sub LoadPaymentList(entryID As String)
    Dim ws As Worksheet
    Dim lastRow As Long, r As Long
    Dim itm As listItem

    If Trim(entryID) = "" Then
        Me.ListView1.ListItems.Clear
        Call UpdatePaymentButtons
        Exit Sub
    End If

    Set ws = ThisWorkbook.Sheets("Payment_Master")
    Me.ListView1.ListItems.Clear

    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row

    For r = 2 To lastRow
        If Trim(ws.Cells(r, 2).value & "") = Trim(entryID & "") Then
            If ws.Cells(r, 6).value = "" Then   ' Active
            
                Set itm = Me.ListView1.ListItems.Add
                itm.text = ws.Cells(r, 1).value              ' A: PaymentID
                itm.SubItems(1) = ws.Cells(r, 2).value       ' B: EntryID
                itm.SubItems(2) = ws.Cells(r, 8).value       ' H: Type
                itm.SubItems(3) = ws.Cells(r, 9).value       ' I: Category
                itm.SubItems(4) = ws.Cells(r, 15).value      ' O: PaymentMode
                itm.SubItems(5) = ws.Cells(r, 10).value      ' J: Name
                itm.SubItems(6) = ws.Cells(r, 11).value      ' K: Company
                itm.SubItems(7) = ws.Cells(r, 12).value      ' L: Model
                itm.SubItems(8) = ws.Cells(r, 13).value      ' M: Serial
                itm.SubItems(9) = ws.Cells(r, 14).value      ' N: Qty
                itm.SubItems(10) = Format(ws.Cells(r, 16).value, "0.00") ' P: Amount
                itm.SubItems(11) = ws.Cells(r, 17).value     ' Q: Warranty
                
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
    If Me.ListView1.ListItems.count > 0 Then
        CommandButton1.enabled = False      ' Add Charges
        CommandButton2.enabled = True       ' Edit Charges
    Else
        CommandButton1.enabled = True       ' Add Charges
        CommandButton2.enabled = False      ' Edit Charges
    End If
    Me.Repaint
End Sub

'===========================================
' ADD CHARGES — CommandButton1
'===========================================
Private Sub CommandButton1_Click()
    Dim payForm As frmPaymentSection
    Dim entryID As String
    Dim existingPayID As String
    
    entryID = Trim(cmbEntryID.value)
    If entryID = "" Then
        MsgBox "Please select Entry ID first!", vbExclamation
        cmbEntryID.SetFocus
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
' EDIT CHARGES — CommandButton2
'===========================================
Private Sub CommandButton2_Click()
    Dim payID As String
    
    On Error Resume Next
    payID = Me.ListView1.selectedItem.text
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
    
    Call LoadPaymentList(cmbEntryID.value)
End Sub

'===========================================
' LISTVIEW DOUBLE CLICK
'===========================================
Private Sub ListView1_DblClick()
    Dim payID As String
    
    On Error Resume Next
    payID = Me.ListView1.selectedItem.text
    On Error GoTo 0
    
    If payID = "" Then Exit Sub

    Dim payForm As New frmPaymentSection
    Set payForm.parentForm = Me

    payForm.LoadPaymentForEdit payID
    payForm.cmbEntryID.enabled = False
    
    payForm.Show vbModal

    Call LoadPaymentList(cmbEntryID.value)
End Sub

'===========================================
' LOAD ACCESSORIES FROM Job_Accessory
' Default ALL CHECKED
'===========================================
Private Sub LoadAccessories(entryID As String)
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim accCount As Integer
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Job_Accessory")
    On Error GoTo 0
    
    ' Reset all checkboxes
    Dim chk As Control
    For Each chk In Me.Controls
        If TypeName(chk) = "CheckBox" Then
            chk.value = False
            chk.caption = ""
            chk.enabled = False
            chk.visible = False
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
                Me.Controls(chkName).value = True      ' Default CHECKED
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
            If chk.visible And chk.caption <> "" Then
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

Private Sub UserForm_Activate()
    Dim entryID As String
    Dim i As Long
    Dim found As Boolean
    
    entryID = Trim(Me.Tag)
    
    If entryID <> "" Then
        found = False
        For i = 0 To cmbEntryID.ListCount - 1
            If UCase(Trim(cmbEntryID.List(i))) = UCase(entryID) Then
                found = True
                Exit For
            End If
        Next i
        
        If Not found Then cmbEntryID.AddItem entryID
        
        On Error Resume Next
        cmbEntryID.value = entryID
        On Error GoTo 0
        
        ' If not already in edit mode, load return data now
        If Not isEditMode Then
            isEditMode = True
            IsLoadingData = True
            Call LoadWarrantyReturnData(entryID)
            IsLoadingData = False
            Me.caption = "Receive Warranty Product - EDIT (" & entryID & ")"
            btnReceive.caption = "UPDATE PRODUCT"
            btnUpdateToCustomer.enabled = True
        End If
        
        Me.Tag = ""
    End If
End Sub

