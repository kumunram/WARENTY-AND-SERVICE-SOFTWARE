VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEntryWizard 
   Caption         =   "JOB ENTRY"
   ClientHeight    =   14235
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   18765
   OleObjectBlob   =   "frmEntryWizard.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmEntryWizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
        ByVal hwnd As LongPtr, ByVal lpOperation As String, ByVal lpFile As String, _
        ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As LongPtr
#Else
    Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
        ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
        ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
#End If

'=== Output Variables (frmOutputOptions ???? same ???) ===
Private m_CustomerID As String
Private m_EntryID As String
Private m_CustomerName As String
Private m_Mobile As String
Public selectedEntryID As String

' OTP ?? ??? ?????? ???????
Private mCurrentOTP As String
Private mOTPTimestamp As Date
Private mPendingMobile As String

' ===== NEW: YE 2 VARIABLES ADD KARO =====
Private mBusy As Boolean
Private m_EntrySaved As Boolean
' ===== END =====

'=================================
' ADD PAYMENT BUTTON CLICK
'=================================
Private Sub btnAddPayment_Click()
    Dim payForm As frmPaymentSection
    Dim selectedEntryID As String
    
    ' Check if product selected
    If Me.lstProducts.selectedItem Is Nothing Then
        MsgBox "Please select a product first!", vbExclamation
        Exit Sub
    End If
    
    selectedEntryID = Me.lstProducts.selectedItem.SubItems(1)  ' Entry ID
    
    ' Open Payment Form
    Set payForm = New frmPaymentSection
    Set payForm.parentForm = Me
    
    payForm.LoadEntryIDs
    payForm.cmbEntryID.value = selectedEntryID
    
    payForm.Show vbModal
    
    ' Refresh after closing
    Call LoadPaymentList(Me.txtCustomerID.caption)
    Set payForm = Nothing
End Sub

'==============================
' ADD PRODUCT BUTTON (?? ?? ???)
'==============================
Private Sub btnAddProduct_Click()
    If Trim(Me.txtCustomerID.caption) = "" Then
        MsgBox "Please enter customer mobile number first.", vbExclamation, "Required"
        Exit Sub
    End If
    
    If Trim(txtMobile.value) = "" Then
        MsgBox "Enter Mobile Number first", vbExclamation
        txtMobile.SetFocus
        Exit Sub
    End If
    
    If Trim(txtCustomerName.caption) = "" Then
        MsgBox "Customer Name missing", vbExclamation
        Exit Sub
    End If
    
    frmAddProduct.Show
End Sub


'=================================
' DELETE PAYMENT BUTTON CLICK
'=================================
Private Sub btnDeletePayment_Click()

    If Me.lstPayments.selectedItem Is Nothing Then
        MsgBox "Please select a payment first!", vbExclamation
        Exit Sub
    End If

    Dim payID As String
    payID = Trim(Me.lstPayments.selectedItem.text & "")

    If payID = "" Then
        MsgBox "Invalid Payment ID!", vbExclamation
        Exit Sub
    End If

    Dim ws As Worksheet
    Dim r As Long, lastRow As Long
    Dim deletedCount As Long

    Set ws = ThisWorkbook.Sheets("Payment_Master")
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    deletedCount = 0

    ' ===== SOFT DELETE ALL ROWS WITH THIS PaymentID =====
    For r = 2 To lastRow
        If UCase(Trim(ws.Cells(r, 1).value & "")) = UCase(payID) Then
            ws.Cells(r, 4).value = "DELETED"                    ' D: Status
            ws.Cells(r, 6).value = Environ("Username")          ' F: DeletedBy
            ws.Cells(r, 7).value = Now                          ' G: DeletedTime
            deletedCount = deletedCount + 1
        End If
    Next r

    If deletedCount > 0 Then
        MsgBox deletedCount & " record(s) marked as DELETED!", vbInformation
        Call LoadPaymentList(Me.txtCustomerID.caption)
    Else
        MsgBox "Payment ID '" & payID & "' not found in database!", vbExclamation
    End If

End Sub
Private Sub btnEditPayment_Click()

    If Me.lstPayments.selectedItem Is Nothing Then
        MsgBox "Select payment first", vbExclamation
        Exit Sub
    End If

    Call lstPayments_DblClick

End Sub

'==============================
' FORM INITIALIZE
'==============================
Private Sub UserForm_Initialize()

Dim editEntryID As String
    editEntryID = Trim(Me.Tag)
    
     lblDate.caption = Format(Date, "dd-mmm-yyyy")
    lblTime.caption = Format(Time, "hh:mm AM/PM")
    
    On Error Resume Next
    Me.lblCompanyName.caption = ThisWorkbook.Sheets("Software_Config").Range("B2").value
    On Error GoTo 0
    
    AddMinMaxButtons Me
     
    lstNameResults.visible = False

    ' Products ListView Setup
    With lstProducts
        .View = lvwReport
        .FullRowSelect = True
        .Gridlines = True
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "Entry Type", 90
        .ColumnHeaders.Add , , "Entry ID", 100
        .ColumnHeaders.Add , , "Product", 120
        .ColumnHeaders.Add , , "Company", 120
        .ColumnHeaders.Add , , "Model", 100
        .ColumnHeaders.Add , , "Serial", 120
        .ColumnHeaders.Add , , "Warranty", 100
        .ColumnHeaders.Add , , "Status", 80
    End With

            With Me.lstPayments
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
    
    Call LoadVerifyType
    
    If Trim(txtMobile.value) <> "" Then
    Call SearchCustomer
End If
    
    ' ===== START: 5 BUTTON DISABLE =====
btnPrintReceipt.enabled = False
btnSendWhatsApp.enabled = False
btnSendEmail.enabled = False
btnExportPDF.enabled = False
btnAssignment.enabled = False
' ===== END =====
    
End Sub


'==============================
' MOBILE CHANGE
'==============================
Private Sub txtMobile_Change()
    If mBusy Then Exit Sub          ' ? Agar already busy hai, toh exit
    
    Dim CleanText As String
    Dim i As Long
    Dim currentText As String
    
    currentText = txtMobile.value
    
    ' Sirf 0-9 allow karo
    For i = 1 To Len(currentText)
        If Mid(currentText, i, 1) >= "0" And Mid(currentText, i, 1) <= "9" Then
            CleanText = CleanText & Mid(currentText, i, 1)
        End If
    Next i
    
    If CleanText <> currentText Then
        txtMobile.value = CleanText
        Exit Sub
    End If
    
    If Len(txtMobile.value) > 10 Then
        txtMobile.value = Left(txtMobile.value, 10)
        Exit Sub
    End If
    
        If Len(txtMobile.value) = 10 Then
        mBusy = True
        On Error GoTo SafeExit      ' ? Error aaye toh bhi unlock ho
        SearchCustomer
SafeExit:
        mBusy = False
        On Error GoTo 0
    Else
        ClearCustomerFields
        lstProducts.ListItems.Clear
        lstPayments.ListItems.Clear
    End If
End Sub

'==============================
' SEARCH CUSTOMER (OTP ?? ???)
'==============================
Public Sub SearchCustomer()
    Dim ws As Worksheet
    Dim f As Range
    Dim enteredOTP As String
    Dim timeDiff As Double
    Dim searchMobile As String
    
    searchMobile = Trim(txtMobile.value)
    
    If Len(searchMobile) = 0 Then Exit Sub
    
    Set ws = Sheets("Customer_Master")
    
    '??? FIX: Columns(3) = Mobile Column (C) ???
    Set f = ws.Columns(3).Find(What:=searchMobile, LookAt:=xlWhole, MatchCase:=False)
    
    ' Numeric fallback
    If f Is Nothing And IsNumeric(searchMobile) Then
        Set f = ws.Columns(3).Find(What:=val(searchMobile), LookAt:=xlWhole)
    End If

        If Not f Is Nothing Then
        '========== CUSTOMER FOUND ==========
        txtCustomerID.caption = CStr(ws.Cells(f.row, 1).value)      'A: ID
        txtCustomerName.caption = CStr(ws.Cells(f.row, 2).value)    'B: Name
        
        ' ? Bas tabhi set karo jab alag ho — loop rokne ke liye
        If txtMobile.value <> CStr(ws.Cells(f.row, 3).value) Then
            txtMobile.value = CStr(ws.Cells(f.row, 3).value)        'C: Mobile
        End If
        
        txtAddress.caption = CStr(ws.Cells(f.row, 4).value)         'D: Address
        txtEmail.caption = LCase(CStr(ws.Cells(f.row, 5).value))    'E: Email
        txtGST.caption = CStr(ws.Cells(f.row, 6).value)           'F: GST
        
        Call LoadCustomerPhoto(CStr(ws.Cells(f.row, 7).value))
        Call LoadProductList(txtCustomerID.caption)
        Call LoadPaymentList(txtCustomerID.caption)
        
    Else
        '========== NOT FOUND ? OTP ==========
        ClearCustomerFields
        
        If MsgBox("Customer Not Found!" & vbCrLf & "Create New Customer?", vbYesNo + vbQuestion, "New Customer") = vbYes Then
            
            Randomize
            mCurrentOTP = Format(Int(Rnd() * 900000) + 100000, "000000")
            mOTPTimestamp = Now
            mPendingMobile = txtMobile.value
            
            Call SendWhatsAppOTP(txtMobile.value, mCurrentOTP)
            
            enteredOTP = InputBox("OTP sent to your WhatsApp" & vbCrLf & _
                                "Mobile: " & txtMobile.value & vbCrLf & vbCrLf & _
                                "Enter 6-digit OTP:", "Verify OTP")
            
            If Trim(enteredOTP) = "" Then
                MsgBox "OTP Required!", vbExclamation
                mCurrentOTP = ""
                Exit Sub
            End If
            
            timeDiff = DateDiff("n", mOTPTimestamp, Now)
            If timeDiff > 10 Then
                MsgBox "OTP Expired!", vbCritical
                mCurrentOTP = ""
                Exit Sub
            End If
            
            If enteredOTP = mCurrentOTP Then
                MsgBox "Verified Successfully!", vbInformation
                mCurrentOTP = ""
                
                ' ===== CUSTOMER CREATION =====
Dim custForm As frmAddCustomer
Dim formTag As String

Set custForm = New frmAddCustomer
custForm.PrepareNewCustomer txtMobile.value
custForm.Show vbModal

' ===== PEHLE TAG LE LO =====
On Error Resume Next
formTag = custForm.Tag
If Err.Number <> 0 Then formTag = ""
On Error GoTo 0

' ===== AB UNLOAD KARO =====
Unload custForm
Set custForm = Nothing

' ===== CHECK SAVED =====
If formTag = "SAVED" Then
    DoEvents
    Application.Wait Now + TimeValue("00:00:01")
    Call SearchCustomer
    Exit Sub
End If
                
            Else
                MsgBox "Invalid OTP!", vbCritical
                mCurrentOTP = ""
            End If
        End If
    End If
End Sub
Private Sub SendWhatsAppOTP(mobile As String, otp As String)
    Dim url As String
    Dim message As String
    Dim wsh As Object
    
    If Trim(mobile) = "" Then Exit Sub
    
    message = "GLOBAL IT SOLUTIONS" & vbCrLf & vbCrLf & _
              "New Customer Verification Code: " & otp & vbCrLf & vbCrLf & _
              "This Code is Valid for 10 minutes" & vbCrLf & _
              "Do not share this code with anyone"
    
    ' === ????? ???? ?? ===
    ' wa.me ?? ??? whatsapp:// protocol
    url = "whatsapp://send?phone=91" & Trim(mobile) & "&text=" & URLEncode(message)
    
    Set wsh = CreateObject("WScript.Shell")
    wsh.Run url, 1, False
    Set wsh = Nothing
End Sub
'==============================
' URL Encode
'==============================
Private Function URLEncode(text As String) As String
    Dim i As Integer, char As String, result As String, CharCode As Long
    
    If Len(Trim(text)) = 0 Then
        URLEncode = ""
        Exit Function
    End If
    
    For i = 1 To Len(text)
        char = Mid(text, i, 1)
        CharCode = Asc(char)
        
        If (CharCode >= 65 And CharCode <= 90) Or _
           (CharCode >= 97 And CharCode <= 122) Or _
           (CharCode >= 48 And CharCode <= 57) Then
            result = result & char
        ElseIf char = " " Then
            result = result & "%20"
        ElseIf char = vbLf Then
            result = result & "%0A"
        ElseIf char = vbCr Then
            ' ?????
        ElseIf CharCode < 16 Then
            result = result & "%0" & Hex(CharCode)
        ElseIf CharCode <= 255 Then
            result = result & "%" & Hex(CharCode)
        Else
            result = result & "%3F"
        End If
    Next i
    
    URLEncode = result
End Function

'==============================
' EDIT CUSTOMER
'==============================
Private Sub btnEditCustomer_Click()
    If Trim(txtMobile.value) = "" Then
        MsgBox "Enter Mobile Number first", vbExclamation
        txtMobile.SetFocus
        Exit Sub
    End If
    
    If txtCustomerID.caption <> "" Then
        frmAddCustomer.LoadCustomerForEdit txtCustomerID.caption  ' ? Customer ID bhejo
        frmAddCustomer.Show vbModal
        Call SearchCustomer
    Else
        MsgBox "No customer loaded to edit!", vbExclamation
    End If
End Sub

'==============================
' CLEAR FIELDS
'==============================
Sub ClearCustomerFields()
    txtCustomerID.caption = ""
    txtCustomerName.caption = ""
    txtAddress.caption = ""
    txtEmail.caption = ""
    txtGST.caption = ""
    On Error Resume Next
    imgCustomerPhoto.Picture = LoadPicture()
    On Error GoTo 0
End Sub

'==============================
' LOAD CUSTOMER PHOTO - ROBUST
'==============================
Public Sub LoadCustomerPhoto(ByVal photoName As String)
    Dim fullPath As String
    Dim folderPath As String
    Dim baseName As String
    Dim ext As String
    Dim i As Integer
    Dim extensions(1 To 4) As String
    
    extensions(1) = ".jpg"
    extensions(2) = ".jpeg"
    extensions(3) = ".png"
    extensions(4) = ".bmp"
    
    On Error Resume Next
    
    ' === GET FOLDER PATH ===
    folderPath = GetFolderSetting("B3")
    If folderPath = "" Then
        folderPath = ThisWorkbook.path & "\GLOBAL_SOFT_DATA\Images\Customers\"
    End If
    If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
    
    ' === TRY WITH ORIGINAL FILENAME ===
    If Trim(photoName) <> "" Then
        fullPath = folderPath & photoName
        If Dir(fullPath) <> "" Then
            imgCustomerPhoto.Picture = LoadPicture(fullPath)
            On Error GoTo 0
            Exit Sub
        End If
        
        ' === TRY DIFFERENT EXTENSIONS ===
        ' Extract basename without extension
        baseName = photoName
        For i = LBound(extensions) To UBound(extensions)
            If LCase(Right(baseName, Len(extensions(i)))) = LCase(extensions(i)) Then
                baseName = Left(baseName, Len(baseName) - Len(extensions(i)))
                Exit For
            End If
        Next i
        
        ' Try each extension
        For i = LBound(extensions) To UBound(extensions)
            fullPath = folderPath & baseName & extensions(i)
            If Dir(fullPath) <> "" Then
                imgCustomerPhoto.Picture = LoadPicture(fullPath)
                On Error GoTo 0
                Exit Sub
            End If
        Next i
    End If
    
    ' === NO PHOTO FOUND - CLEAR IMAGE ===
    imgCustomerPhoto.Picture = LoadPicture()
    On Error GoTo 0
End Sub

'==============================
' GET FOLDER SETTING
'==============================
Function GetFolderSetting(CellAddress As String) As String
    Dim p As String
    On Error Resume Next
    p = Sheets("Settings").Range(CellAddress).value
    On Error GoTo 0
    If p <> "" And Right(p, 1) <> "\" Then p = p & "\"
    GetFolderSetting = p
End Function

'==============================
' DELETE PRODUCT
'==============================
Private Sub btnDeleteProduct_Click()
    Dim ws As Worksheet
    Dim entryID As String
    Dim r As Long
    
    If Me.lstProducts.selectedItem Is Nothing Then
        MsgBox "Please select product first!", vbExclamation
        Exit Sub
    End If
    
    entryID = Me.lstProducts.selectedItem.SubItems(1)
    
    If MsgBox("Delete this product?" & vbCrLf & "Entry: " & entryID, vbQuestion + vbYesNo) = vbNo Then Exit Sub
    
    Set ws = ThisWorkbook.Sheets("Job_Product")
    
    For r = 2 To ws.Cells(ws.Rows.count, 1).End(xlUp).row
        If UCase(Trim(ws.Cells(r, 1).value)) = UCase(entryID) Then
            ws.Cells(r, 13).value = "DELETED"
            MsgBox "Product Deleted!", vbInformation
            Exit For
        End If
    Next r
    
    Call LoadProductList(Me.txtCustomerID.caption)
End Sub

'==============================
' EDIT PRODUCT
'==============================
Private Sub btnEditProduct_Click()
    Dim entryID As String
    
    If Me.lstProducts.selectedItem Is Nothing Then
        MsgBox "Please select product first!", vbExclamation
        Exit Sub
    End If
    
    entryID = Me.lstProducts.selectedItem.SubItems(1)
    
    Dim prodForm As frmAddProduct
    Set prodForm = New frmAddProduct
    
    prodForm.Tag = entryID
    prodForm.LoadProductForEdit entryID
    prodForm.Show vbModal
    
    Set prodForm = Nothing
    Call LoadProductList(Me.txtCustomerID.caption)
End Sub

'==============================
' LOAD PRODUCT LIST - FINAL UPDATE
'==============================
Public Sub LoadProductList(customerID As String)
    Dim ws As Worksheet, wsAcc As Worksheet
    Dim r As Long, accRow As Long
    Dim item As listItem
    Dim prodStatus As String
    Dim entryID As String
    Dim hasAccessories As Boolean
    
    If Trim(customerID) = "" Then Exit Sub
    
    Set ws = ThisWorkbook.Sheets("Job_Product")
    Set wsAcc = ThisWorkbook.Sheets("Job_Accessory")
    
    Me.lstProducts.ListItems.Clear
    
    For r = 2 To ws.Cells(ws.Rows.count, 1).End(xlUp).row
        If UCase(Trim(ws.Cells(r, 2).value)) = UCase(Trim(customerID)) Then
            prodStatus = UCase(Trim(ws.Cells(r, 4).value))
            
            If prodStatus <> "DELETED" And prodStatus <> "DELIVERED" Then
                entryID = ws.Cells(r, 1).value
                
                '=== PRODUCT (????? ????) ===
                Set item = Me.lstProducts.ListItems.Add(, , ws.Cells(r, 3).value) ' Entry Type
                item.SubItems(1) = entryID
                item.SubItems(2) = ws.Cells(r, 6).value ' Product
                item.SubItems(3) = ws.Cells(r, 7).value ' Company
                item.SubItems(4) = ws.Cells(r, 8).value ' Model
                item.SubItems(5) = ws.Cells(r, 9).value ' Serial
                item.SubItems(6) = ws.Cells(r, 11).value ' Warranty
                item.SubItems(7) = ws.Cells(r, 4).value ' Status
                
                ' Product = Bold + Black
                item.Bold = True
                item.ForeColor = RGB(0, 0, 0)
                
                                                '=== ACCESSORIES (??? Columns) ===
                hasAccessories = False
                
                For accRow = 2 To wsAcc.Cells(wsAcc.Rows.count, 1).End(xlUp).row
                    If UCase(Trim(wsAcc.Cells(accRow, 1).value)) = UCase(entryID) Then
                        hasAccessories = True
                        
                        Set item = Me.lstProducts.ListItems.Add(, , "") ' Entry Type blank
                        item.SubItems(1) = "" ' Entry ID blank
                        
                        ' Accessory Name (Column D = 4)
                        item.SubItems(2) = "    › " & wsAcc.Cells(accRow, 4).value
                        
                        ' Status (Column E = 5) - RECEIVED / NOT RECEIVED
                        item.SubItems(3) = wsAcc.Cells(accRow, 5).value
                        
                        ' Model: blank for accessories (???? ??? ????)
                        item.SubItems(4) = ""
                        
                        ' Serial Number (Column F = 6) - Accessory ?? Serial
                        item.SubItems(5) = wsAcc.Cells(accRow, 6).value
                        
                        ' Working Status (Column G = 7) - WORKING / NOT WORKING
                        item.SubItems(6) = wsAcc.Cells(accRow, 7).value
                        
                        ' Status column: blank
                        item.SubItems(7) = ""
                        
                        ' Accessories: Gray + Normal
                        item.Bold = False
                        item.ForeColor = RGB(100, 100, 100)
                        
                    End If
                Next accRow
                
                ' Spacing row (??? accessories ??? ??)
                If hasAccessories Then
                    Set item = Me.lstProducts.ListItems.Add(, , "")
                    item.SubItems(2) = "    "
                End If
                
            End If
        End If
    Next r
    
    ' Column widths ??? ?????? ???? ???? ?? ??? ???-??? ????
    Call AdjustListViewColumns(Me.lstProducts)
    
End Sub

'==============================
' COLUMN WIDTHS ?? ????????? ???? ?? ???
'==============================
Private Sub AdjustListViewColumns(lv As ListView)
    Dim i As Integer
    ' Column 0 (Entry Type): ????
    lv.ColumnHeaders(1).Width = 80
    ' Column 1 (Entry ID): ??????
    lv.ColumnHeaders(2).Width = 90
    ' Column 2 (Product): ???? (Accessory name ???? ??????)
    lv.ColumnHeaders(3).Width = 180
    ' Column 3 (Company): ??????
    lv.ColumnHeaders(4).Width = 100
    ' Column 4 (Model/Serial): ??????
    lv.ColumnHeaders(5).Width = 120
    ' Column 5 (Serial/Received): ??????
    lv.ColumnHeaders(6).Width = 100
    ' Column 6 (Warranty): ????
    lv.ColumnHeaders(7).Width = 80
    ' Column 7 (Status): ??????
    lv.ColumnHeaders(8).Width = 100
End Sub
Private Sub lstProducts_DblClick()
    Dim selectedItem As listItem
    Dim entryID As String
    Dim parentEntryID As String
    Dim i As Integer
    
    If lstProducts.selectedItem Is Nothing Then Exit Sub
    
    Set selectedItem = lstProducts.selectedItem
    
    ' === ??? ???? ?? Product ?? ?? Accessory ===
    ' Product: Entry Type (Column 0) ??? Warranty/Service ???? ?? (???? ????)
    ' Accessory: Entry Type (Column 0) ???? ??
    
    If Trim(selectedItem.text) <> "" Then
        ' === ?? PRODUCT ?? ===
        entryID = selectedItem.SubItems(1)
        
        If entryID <> "" Then
            Dim prodForm As frmAddProduct
            Set prodForm = New frmAddProduct
            
            prodForm.Tag = entryID
            prodForm.LoadProductForEdit entryID
            prodForm.Show vbModal
            
            Set prodForm = Nothing
            Call LoadProductList(Me.txtCustomerID.caption)
        End If
        
    Else
        ' === ?? ACCESSORY ?? ===
        ' Parent Entry ID ?????? (??? ?? ??? ???? Product ???? Entry Type ???? ? ??)
        For i = selectedItem.index To 1 Step -1
            If Trim(lstProducts.ListItems(i).text) <> "" Then
                ' ?? Parent Product ??? ???
                parentEntryID = lstProducts.ListItems(i).SubItems(1)
                Exit For
            End If
        Next i
        
        If parentEntryID <> "" Then
            Dim accForm As frmAccessories
            Set accForm = New frmAccessories
            
            accForm.customerID = Me.txtCustomerID.caption
            accForm.txtEntryID.value = parentEntryID
            accForm.Tag = parentEntryID
            accForm.Show vbModal
            
            Set accForm = Nothing
            
            ' List ??????? ????
            Call LoadProductList(Me.txtCustomerID.caption)
        Else
            MsgBox "?? Accessory ?? Parent Product ???? ????!", vbExclamation
        End If
        
    End If
End Sub
Private Sub btnSaveEntry_Click()
    Dim ws As Worksheet
    Dim r As Long
    Dim entryID As String
    Dim prefix As String

    ' === VALIDATIONS ===
    If Me.txtCustomerID.caption = "" Then
        MsgBox "Please Add Customer First", vbExclamation
        Exit Sub
    End If

    If Me.lstProducts.ListItems.count = 0 Then
        MsgBox "Please Add Product First", vbExclamation
        Exit Sub
    End If

    If Me.cmbVerifyType.value = "" Then
        MsgBox "Please Select Verify Type", vbExclamation
        Exit Sub
    End If

    If Me.cmbVerifyName.value = "" Then
        MsgBox "Please Select Verify Name", vbExclamation
        Exit Sub
    End If

    ' === GENERATE ENTRY ID ===
    If Me.lstProducts.ListItems(1).text = "Warranty" Then
        prefix = "WAR"
    Else
        prefix = "SER"
    End If

    Set ws = Sheets("Job_Master")
    r = ws.Cells(ws.Rows.count, 1).End(xlUp).row + 1
    entryID = prefix & Format(r, "00000")

    ' === SAVE TO JOB_MASTER ===
    ws.Cells(r, 1).value = entryID
    ws.Cells(r, 2).value = Me.txtCustomerID.caption
    ws.Cells(r, 3).value = Me.txtMobile.value
    ws.Cells(r, 4).value = Me.txtCustomerName.caption
    ws.Cells(r, 5).value = Me.txtAddress.caption
    ws.Cells(r, 6).value = Me.cmbVerifyType.value
    ws.Cells(r, 7).value = Me.cmbVerifyName.value
    ws.Cells(r, 8).value = Format(Now, "dd-MMM-yyyy hh:mm AM/PM")  ' Entry Date/Time

    ' === UPDATE CUSTOMER MASTER ===
    Dim wsCust As Worksheet
    Dim f As Range
    Set wsCust = Sheets("Customer_Master")
    Set f = wsCust.Columns(2).Find(Me.txtMobile.value, LookAt:=xlWhole)

    If Not f Is Nothing Then
        wsCust.Cells(f.row, 12).value = Me.cmbVerifyType.value
        wsCust.Cells(f.row, 13).value = Me.cmbVerifyName.value
    End If

    ' ============================================================
    ' ?? CRITICAL FIX #1: SAVE WORKBOOK AFTER DATA WRITE
    ' ============================================================
    On Error GoTo SaveError
    Application.DisplayAlerts = False
    ThisWorkbook.Save
    Application.DisplayAlerts = True
    ' ============================================================

    ' === REFRESH LISTS ===
    Call LoadProductList(Me.txtCustomerID.caption)
    Call LoadPaymentList(Me.txtCustomerID.caption)

    ' === ENABLE OUTPUT BUTTONS ===
    m_EntrySaved = True
    btnPrintReceipt.enabled = True
    btnSendWhatsApp.enabled = True
    btnSendEmail.enabled = True
    btnExportPDF.enabled = True
    btnAssignment.enabled = True

    MsgBox "Entry Saved Successfully!" & vbCrLf & vbCrLf & _
           "Entry ID: " & entryID & vbCrLf & _
           "File Auto-Saved", vbInformation, "Saved"

    Exit Sub

SaveError:
    Application.DisplayAlerts = True
    MsgBox "CRITICAL ERROR: Data written but FILE NOT SAVED!" & vbCrLf & _
           "Error: " & Err.Description & vbCrLf & vbCrLf & _
           "Please manually save the file (Ctrl+S) immediately!", vbCritical, "SAVE FAILED"
End Sub





'==============================
' RESET FORM
'==============================
Sub ResetEntryForm()
    txtCustomerID.caption = ""
    txtCustomerName.caption = ""
    txtAddress.caption = ""
    txtEmail.caption = ""
    txtGST.caption = ""
    txtMobile.value = ""
    lstProducts.ListItems.Clear
    lstPayments.ListItems.Clear
    cmbVerifyType.value = ""
    cmbVerifyName.value = ""
    On Error Resume Next
    imgCustomerPhoto.Picture = LoadPicture()
    On Error GoTo 0
    txtMobile.SetFocus
    
    ' ===== RESET KE BAAD 5 BUTTONS DISABLE =====
    m_EntrySaved = False
    btnPrintReceipt.enabled = False
    btnSendWhatsApp.enabled = False
    btnSendEmail.enabled = False
    btnExportPDF.enabled = False
    btnAssignment.enabled = False
    
End Sub

'===========================================
' SEARCH BY NAME / MOBILE / ADDRESS / ANYTHING
'===========================================
Private Sub txtSearchName_Change()
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim searchText As String
    Dim matchCount As Long
    Dim custID As String, custName As String, custMobile As String
    Dim custAddress As String, custEmail As String
    
    searchText = Trim(txtSearchName.text)
    
    ' 2 ????? ?? ?? ?? ?? List Hide
    If Len(searchText) < 2 Then
        lstNameResults.visible = False
        Exit Sub
    End If
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Customer_Master")
    On Error GoTo 0
    
    If ws Is Nothing Then
        lstNameResults.visible = False
        Exit Sub
    End If
    
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    If lastRow < 2 Then
        lstNameResults.visible = False
        Exit Sub
    End If
    
    lstNameResults.Clear
    matchCount = 0
    
    ' ==========================================
    ' MULTI-FIELD SEARCH: Name, Mobile, Address, Email, GST, ID
    ' ==========================================
    For i = 2 To lastRow
        
        ' Read all fields (Auto-detect column order)
        custID = Trim(ws.Cells(i, 1).value & "")
        custName = Trim(ws.Cells(i, 2).value & "")      ' Col B = Name (current)
        custMobile = Trim(ws.Cells(i, 3).value & "")    ' Col C = Mobile (current)
        custAddress = Trim(ws.Cells(i, 4).value & "")    ' Col D = Address
        custEmail = Trim(ws.Cells(i, 5).value & "")       ' Col E = Email
        
        ' Search in ALL fields
        Dim found As Boolean
        found = False
        
        ' 1. Search in Name
        If InStr(1, LCase(custName), LCase(searchText), vbTextCompare) > 0 Then found = True
        
        ' 2. Search in Mobile
        If Not found Then
            If InStr(1, LCase(custMobile), LCase(searchText), vbTextCompare) > 0 Then found = True
        End If
        
        ' 3. Search in Address
        If Not found Then
            If InStr(1, LCase(custAddress), LCase(searchText), vbTextCompare) > 0 Then found = True
        End If
        
        ' 4. Search in Email
        If Not found Then
            If InStr(1, LCase(custEmail), LCase(searchText), vbTextCompare) > 0 Then found = True
        End If
        
        ' 5. Search in Customer ID
        If Not found Then
            If InStr(1, LCase(custID), LCase(searchText), vbTextCompare) > 0 Then found = True
        End If
        
        ' If found, add to list
        If found Then
            ' Format: "Name | Mobile | Address"
            Dim displayText As String
            displayText = custName & " | " & custMobile
            If custAddress <> "" Then
                displayText = displayText & " | " & Left(custAddress, 20)
            End If
            
            lstNameResults.AddItem displayText
            
            ' Store row number in hidden data (for retrieval)
            ' We'll use ItemData if available, or just search again on click
            
            matchCount = matchCount + 1
            If matchCount >= 10 Then Exit For  ' Max 10 results
        End If
        
    Next i
    
    lstNameResults.visible = (matchCount > 0)
    
    ' Auto-size listbox height based on results
    If matchCount > 0 Then
        lstNameResults.Height = Application.Min(matchCount * 15 + 10, 150)
    End If
End Sub
'===========================================
' LIST CLICK - LOAD SELECTED CUSTOMER
'===========================================
Private Sub lstNameResults_Click()
    Dim selectedItem As String
    Dim mobile As String
    Dim pos As Long
    
    If lstNameResults.ListIndex < 0 Then Exit Sub
    
    selectedItem = lstNameResults.List(lstNameResults.ListIndex)
    
    ' Extract Mobile from format: "Name | Mobile | Address"
    ' Mobile is between first " | " and second " | " (or end)
    pos = InStr(selectedItem, " | ")
    If pos > 0 Then
        ' Get text after first " | "
        mobile = Trim(Mid(selectedItem, pos + 3))
        
        ' If there's another " | ", trim it
        pos = InStr(mobile, " | ")
        If pos > 0 Then
            mobile = Trim(Left(mobile, pos - 1))
        End If
        
        ' Clean mobile (remove non-numeric)
        Dim cleanMobile As String
        Dim i As Long
        cleanMobile = ""
        For i = 1 To Len(mobile)
            If Mid(mobile, i, 1) >= "0" And Mid(mobile, i, 1) <= "9" Then
                cleanMobile = cleanMobile & Mid(mobile, i, 1)
            End If
        Next i
        
        ' Set mobile and trigger search
        txtMobile.value = cleanMobile
        lstNameResults.visible = False
        txtSearchName.text = "" ' Clear search box
        
        ' Trigger customer search
        Call SearchCustomer
    End If
End Sub
'==============================
' List ?? ???? Click ???? ?? Hide ??
'==============================
Private Sub UserForm_Click()
    lstNameResults.visible = False
End Sub
'=================================
' ADD PAYMENT BUTTON CLICK
'=================================
Private Sub btnPayment_Click()  ' ???????? ??? ?? ??? ?? ?? ??
    Dim payForm As frmPaymentSection
    Dim selectedEntryID As String
    
    ' Check if any product selected
    If Me.lstProducts.ListItems.count = 0 Then
        MsgBox "Please add product first!", vbExclamation
        Exit Sub
    End If
    
    ' Get selected EntryID from list (first column)
    If Me.lstProducts.selectedItem Is Nothing Then
        MsgBox "Please select a product from list!", vbExclamation
        Exit Sub
    End If
    
    selectedEntryID = Me.lstProducts.selectedItem.SubItems(1)
    
    ' Open Payment Form
    Set payForm = New frmPaymentSection
    Set payForm.parentForm = Me  ' Reference pass ???
    
    ' Load Entry IDs for this customer
    payForm.LoadEntryIDs
    
    ' Auto-select the current EntryID
    payForm.cmbEntryID.value = selectedEntryID
    
    ' If payment already exists for this entry, load it for edit
    If PaymentExists(selectedEntryID) Then
        Dim payID As String
        payID = GetPaymentIDByEntry(selectedEntryID)
        payForm.LoadPaymentForEdit payID
    End If
    
    ' Show form
    payForm.Show vbModal
    
    ' Refresh payment list after closing
    Call LoadPaymentList(Me.txtCustomerID.caption)
    
    Set payForm = Nothing
End Sub
Public Sub LoadPaymentList(customerID As String)
    Dim ws As Worksheet
    Dim lastRow As Long, r As Long
    Dim itm As listItem

    Set ws = ThisWorkbook.Sheets("Payment_Master")
    Me.lstPayments.ListItems.Clear

    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row

    For r = 2 To lastRow
        ' CustomerID match (Col C) AND Active (Col F = DeletedBy blank)
        If Trim(ws.Cells(r, 3).value & "") = Trim(customerID & "") Then
            If ws.Cells(r, 6).value = "" Then   ' F: DeletedBy = blank = Active
            
                Set itm = Me.lstPayments.ListItems.Add
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
End Sub
'=================================
' CHECK IF PAYMENT EXISTS
'=================================
Private Function PaymentExists(entryID As String) As Boolean
    Dim ws As Worksheet
    Dim f As Range
    
    Set ws = ThisWorkbook.Sheets("Payment_Master")
    Set f = ws.Columns(2).Find(entryID, LookIn:=xlValues, LookAt:=xlWhole)
    
    PaymentExists = Not f Is Nothing
End Function

'==============================
' LOAD VERIFY TYPE (????? ?? ???)
'==============================
'==============================
' LOAD VERIFY TYPE (FIXED)
'==============================
Public Sub LoadVerifyType()
    Dim ws As Worksheet
    Dim r As Long
    
    ' PEHLE CLEAR KARO — Error ke baad bhi clean rahe
    Me.cmbVerifyType.Clear
    
    On Error GoTo SafeExit
    Set ws = ThisWorkbook.Sheets("Verify_Master")
    
    For r = 2 To ws.Cells(ws.Rows.count, 1).End(xlUp).row
        If ws.Cells(r, 2).value <> "" Then
            If Not IsInCombo(Me.cmbVerifyType, ws.Cells(r, 2).value) Then
                Me.cmbVerifyType.AddItem ws.Cells(r, 2).value
            End If
        End If
    Next r
    
SafeExit:
    On Error Resume Next
    ' Duplicate Add Type se bachao
    If Not IsInCombo(Me.cmbVerifyType, "+ Add Type") Then
        Me.cmbVerifyType.AddItem "+ Add Type"
    End If
    On Error GoTo 0
End Sub

'==============================
' LOAD VERIFY NAME (FINAL FIX)
'==============================
Public Sub LoadVerifyName()
    Dim ws As Worksheet
    Dim r As Long
    Dim t As String
    
    ' PEHLE CLEAR KARO — Safe tareeke se
    On Error Resume Next
    cmbVerifyName.Clear
    On Error GoTo 0
    
    On Error GoTo SafeExit
    Set ws = ThisWorkbook.Sheets("Verify_Master")
    
    t = Trim(cmbVerifyType.value & "")
    If t = "" Then GoTo SafeExit
    
    ' Type ke hisaab se names load karo
    For r = 2 To ws.Cells(ws.Rows.count, 1).End(xlUp).row
        If UCase(Trim(ws.Cells(r, 2).value & "")) = UCase(t) Then
            If Trim(ws.Cells(r, 3).value & "") <> "" Then
                cmbVerifyName.AddItem ws.Cells(r, 3).value
            End If
        End If
    Next r
    
SafeExit:
    ' + Add Name HAMESHA last me add karo — chahe koi data mile ya na mile
    On Error Resume Next
    cmbVerifyName.AddItem "+ Add Name"
    On Error GoTo 0
End Sub
'==============================
' VERIFY TYPE CHANGE (FINAL FIX)
'==============================
Private Sub cmbVerifyType_Change()
    If mBusy Then Exit Sub
    mBusy = True
    
    If cmbVerifyType.value = "+ Add Type" Then
        Dim val As String
        val = InputBox("Enter New Type")
        If Trim(val) <> "" Then
            Call SaveVerifyType(val)
            Call LoadVerifyType
            On Error Resume Next
            cmbVerifyType.value = val
            On Error GoTo 0
        End If
    ElseIf Trim(cmbVerifyType.value & "") <> "" Then
        Call LoadVerifyName
    End If
    
    mBusy = False
End Sub

'==============================
' VERIFY NAME CHANGE (FINAL FIX)
'==============================
Private Sub cmbVerifyName_Change()
    If mBusy Then Exit Sub
    mBusy = True
    
    If cmbVerifyName.value = "+ Add Name" Then
        Dim val As String
        Dim selectedType As String
        
        selectedType = cmbVerifyType.value
        
        If selectedType = "" Then
            MsgBox "Please Select Verify Type First!", vbExclamation
            cmbVerifyName.value = ""
            mBusy = False
            Exit Sub
        End If
        
        val = InputBox(selectedType & " - Enter New Name:", "Add New Name")
        
        If Trim(val) <> "" Then
            Call SaveVerifyName(selectedType, val)
            Call LoadVerifyName
            On Error Resume Next
            cmbVerifyName.value = val
            On Error GoTo 0
        Else
            cmbVerifyName.value = ""
        End If
    End If
    
    mBusy = False
End Sub

'==============================
' SAVE VERIFY TYPE (FIXED)
'==============================
Private Sub SaveVerifyType(val As String)
    Dim ws As Worksheet
    Dim r As Long
    
    Set ws = GetVerifyMasterSheet()  ' Auto-create agar nahi hai
    
    r = ws.Cells(ws.Rows.count, 1).End(xlUp).row + 1
    ws.Cells(r, 1) = r - 1
    ws.Cells(r, 2) = val
    ws.Cells(r, 3) = ""
End Sub


'==============================
' SAVE VERIFY NAME (FIXED)
'==============================
Private Sub SaveVerifyName(verifyType As String, val As String)
    Dim ws As Worksheet
    Dim r As Long
    
    Set ws = GetVerifyMasterSheet()  ' Auto-create agar nahi hai
    
    r = ws.Cells(ws.Rows.count, 1).End(xlUp).row + 1
    ws.Cells(r, 1) = r - 1
    ws.Cells(r, 2) = verifyType
    ws.Cells(r, 3) = val
End Sub
'==============================
' IsInCombo Helper (????? ?? ???)
'==============================
Private Function IsInCombo(cmb As ComboBox, val As String) As Boolean
    Dim i As Long
    For i = 0 To cmb.ListCount - 1
        If cmb.List(i) = val Then
            IsInCombo = True
            Exit Function
        End If
    Next i
    IsInCombo = False
End Function
Private Sub lstPayments_DblClick()

    If Me.lstPayments.selectedItem Is Nothing Then Exit Sub

    Dim payID As String
    payID = Me.lstPayments.selectedItem.text

    Dim payForm As New frmPaymentSection
    Set payForm.parentForm = Me

    payForm.LoadPaymentForEdit payID

    payForm.Show vbModal

    ' Refresh
    Call LoadPaymentList(Me.txtCustomerID.caption)

End Sub
'=================================
' EDIT SELECTED PAYMENT (Common Function)
'=================================
Private Sub EditSelectedPayment()
    Dim payForm As frmPaymentSection
    Dim selectedPayID As String
    Dim selectedEntryID As String
    
    
    ' Check if payment selected
    If Me.lstPayments.selectedItem Is Nothing Then
        MsgBox "Please select a payment from list!", vbExclamation
        Exit Sub
    End If
    
    ' Get selected values
    selectedPayID = Me.lstPayments.selectedItem.text  ' Payment ID
    selectedEntryID = Me.lstPayments.selectedItem.SubItems(1)  ' Entry ID
    
    If Trim(selectedPayID) = "" Then
        MsgBox "Invalid Payment ID!", vbExclamation
        Exit Sub
    End If
    
    ' Open Payment Form
    Set payForm = New frmPaymentSection
    Set payForm.parentForm = Me
    
    ' Load data
    payForm.LoadEntryIDs
    payForm.cmbEntryID.value = selectedEntryID
    payForm.LoadPaymentForEdit selectedPayID  ' Edit Mode
    
    payForm.Show vbModal
    
    ' Refresh after closing
    Call LoadPaymentList(Me.txtCustomerID.caption)
    Set payForm = Nothing
End Sub
Private Sub txtMobile_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    ' ????? 0-9 allowed (ASCII 48-57)
    If Not (KeyAscii >= 48 And KeyAscii <= 57) Then
        KeyAscii = 0  ' ????? ?? ??
        Exit Sub
    End If
    
    ' 10 ????? ???? ?? ??? ?? ???? ????? ??
    If Len(txtMobile.value) >= 10 And KeyAscii <> 8 Then ' 8 = Backspace
        KeyAscii = 0
    End If
End Sub


Public Sub LoadEntryForEdit(entryID As String)
    Dim wsProduct As Worksheet
    Dim wsCustomer As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim customerID As String
    Dim customerMobile As String
    
    If Trim(entryID) = "" Then Exit Sub
    
    On Error Resume Next
    Set wsProduct = ThisWorkbook.Sheets("Job_Product")
    Set wsCustomer = ThisWorkbook.Sheets("Customer_Master")
    On Error GoTo 0
    
    If wsProduct Is Nothing Then
        MsgBox "Job_Product sheet not found!", vbCritical
        Exit Sub
    End If
    
    ' === FIND ENTRY IN Job_Product ===
    lastRow = wsProduct.Cells(wsProduct.Rows.count, 1).End(xlUp).row
    customerID = ""
    customerMobile = ""
    
    For r = 2 To lastRow
        If UCase(Trim(wsProduct.Cells(r, 1).value)) = UCase(Trim(entryID)) Then
            
            ' === GET CUSTOMER ID ===
            customerID = Trim(wsProduct.Cells(r, 2).value)
            
            ' === FIND CUSTOMER MOBILE FROM Customer_Master ===
            If customerID <> "" And Not wsCustomer Is Nothing Then
                Dim custRow As Long
                For custRow = 2 To wsCustomer.Cells(wsCustomer.Rows.count, 1).End(xlUp).row
                    If UCase(Trim(wsCustomer.Cells(custRow, 1).value)) = UCase(customerID) Then
                        customerMobile = Trim(wsCustomer.Cells(custRow, 3).value) ' ? Col C = Mobile
                        Exit For
                    End If
                Next custRow
            End If
            
            Exit For
        End If
    Next r
    
    ' === LOAD MOBILE NUMBER AND TRIGGER SEARCH ===
    If customerMobile <> "" Then
        Me.txtMobile.value = customerMobile
        ' Trigger the search which will load everything automatically
        Call SearchCustomer
    Else
        MsgBox "Customer mobile number not found for Entry ID: " & entryID, vbExclamation
    End If
    
        ' ===== EXISTING ENTRY LOAD HUA, 5 BUTTONS ENABLE =====
    m_EntrySaved = True
    btnPrintReceipt.enabled = True
    btnSendWhatsApp.enabled = True
    btnSendEmail.enabled = True
    btnExportPDF.enabled = True
    btnAssignment.enabled = True
    
End Sub

' === HELPER: Load Customer by ID ===
Private Sub LoadCustomerByID(custID As String)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    Set ws = ThisWorkbook.Sheets("Customer_Master")
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    
    For i = 2 To lastRow
        If UCase(Trim(ws.Cells(i, 1).value)) = UCase(custID) Then
            
            '??? FIX: Col 2 = Name, Col 3 = Mobile ???
            Me.txtCustomerID.caption = custID
            Me.txtCustomerName.caption = Trim(ws.Cells(i, 2).value)   'B: Name
            Me.txtMobile.value = Trim(ws.Cells(i, 3).value)           'C: Mobile
            Me.txtAddress.caption = Trim(ws.Cells(i, 4).value)        'D: Address
            Me.txtEmail.caption = Trim(ws.Cells(i, 5).value)          'E: Email
            Me.txtGST.caption = Trim(ws.Cells(i, 6).value)            'F: GST
            
            Exit For
        End If
    Next i
End Sub


'========================================
' REFRESH OUTPUT VARIABLES FROM FORM
'========================================
Private Sub RefreshOutputVars()
    m_CustomerID = Trim(Me.txtCustomerID.caption & "")
    m_CustomerName = Trim(Me.txtCustomerName.caption & "")
    m_Mobile = Trim(Me.txtMobile.value & "")
    
    ' EntryID = selected product ka EntryID (agar koi select hai)
    If Not Me.lstProducts.selectedItem Is Nothing Then
        m_EntryID = Trim(Me.lstProducts.selectedItem.SubItems(1) & "")
    Else
        m_EntryID = ""
    End If
End Sub

Private Sub btnPrintReceipt_Click()
    On Error GoTo ErrorHandler
    
    Call RefreshOutputVars
    
    If m_CustomerID = "" Then
        MsgBox "Customer ID not found! Please load customer first.", vbExclamation
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    Dim wsReceipt As Worksheet
    Set wsReceipt = GenerateReceiptSheet
    
    If wsReceipt Is Nothing Then
        MsgBox "Receipt generation failed!", vbCritical
        GoTo Cleanup
    End If
    
    wsReceipt.PrintOut Copies:=1, Preview:=False
    
Cleanup:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Exit Sub
    
ErrorHandler:
    MsgBox "Print Error: " & Err.Description, vbCritical
    Resume Cleanup
End Sub
Private Sub btnSendWhatsApp_Click()
    On Error GoTo ErrorHandler
    
    Call RefreshOutputVars
    
    If Trim(m_Mobile) = "" Then
        MsgBox "Mobile number not found!", vbExclamation
        Exit Sub
    End If
    
    ' 10 digit clean
    Dim cleanMobile As String
    cleanMobile = FormatMobileNumber(m_Mobile)
    
    If Left(cleanMobile, 2) = "91" And Len(cleanMobile) > 10 Then
        cleanMobile = Right(cleanMobile, 10)
    End If
    
    Dim msg As String
    msg = BuildWhatsAppMessage()
    msg = URLEncodeOutput(msg)
    
    Dim url As String
    url = "whatsapp://send?phone=91" & cleanMobile & "&text=" & msg
    
    ShellExecute 0, "open", url, vbNullString, vbNullString, 1
    
    Exit Sub
    
ErrorHandler:
    MsgBox "WhatsApp Error: " & Err.Description, vbExclamation
End Sub
Private Sub btnSendEmail_Click()
    On Error GoTo ErrorHandler
    
    Call RefreshOutputVars
    
    If m_CustomerID = "" Then
        MsgBox "No Customer Selected!", vbExclamation
        Exit Sub
    End If
    
    ' Email Settings from Software_Config B17-B22
    Dim wsConfig As Worksheet
    Dim senderEmail As String, emailPassword As String, appPassword As String
    Dim smtpServer As String, smtpPort As String, sslEnable As String
    
    Set wsConfig = ThisWorkbook.Sheets("Software_Config")
    senderEmail = Trim(wsConfig.Range("B17").value & "")
    smtpServer = Trim(wsConfig.Range("B19").value & "")
    appPassword = Trim(wsConfig.Range("B20").value & "")
    emailPassword = Trim(wsConfig.Range("B18").value & "")
    smtpPort = Trim(wsConfig.Range("B21").value & "")
    sslEnable = UCase(Trim(wsConfig.Range("B22").value & ""))
    
    If senderEmail = "" Or smtpServer = "" Then
        MsgBox "Email settings not configured in Communication Settings!", vbExclamation
        Exit Sub
    End If
    
    Dim usePassword As String
    If appPassword <> "" Then usePassword = appPassword Else usePassword = emailPassword
    If usePassword = "" Then MsgBox "Email password missing!", vbExclamation: Exit Sub
    If smtpPort = "" Then smtpPort = "587"
    
    ' Customer Email
    Dim wsCust As Worksheet
    Dim custEmail As String
    Dim lastRow As Long, i As Long
    custEmail = ""
    
    Set wsCust = ThisWorkbook.Sheets("Customer_Master")
    lastRow = wsCust.Cells(wsCust.Rows.count, 1).End(xlUp).row
    For i = 2 To lastRow
        If UCase(Trim(wsCust.Cells(i, 1).value & "")) = UCase(m_CustomerID) Then
            custEmail = Trim(wsCust.Cells(i, 5).value & "")
            Exit For
        End If
    Next i
    
    If custEmail = "" Then
        MsgBox "Customer email not found!", vbExclamation
        Exit Sub
    End If
    
    ' Generate HTML
    Dim htmlReceipt As String
    htmlReceipt = GenerateReceiptHTML()
    If htmlReceipt = "" Then MsgBox "Failed to generate receipt!", vbCritical: Exit Sub
    
    ' Send via CDO
    Dim cdoMsg As Object, cdoConf As Object
    Set cdoMsg = CreateObject("CDO.Message")
    Set cdoConf = CreateObject("CDO.Configuration")
    
    With cdoConf.Fields
        .item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
        .item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = GetSMTPServerAddress(smtpServer)
        .item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = CInt(smtpPort)
        .item("http://schemas.microsoft.com/cdo/configuration/sendusername") = senderEmail
        .item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = usePassword
        .item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
        If sslEnable = "YES" Then
            .item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
        Else
            .item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = False
        End If
        .item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
        .Update
    End With
    
    Dim compName As String
    compName = Trim(wsConfig.Range("B2").value & "")
    If compName = "" Then compName = "GLOBAL IT SOLUTIONS"
    
    With cdoMsg
        Set .Configuration = cdoConf
        .From = senderEmail
        .To = custEmail
        .subject = "Service Receipt - " & m_CustomerName & " (" & m_CustomerID & ")"
        .htmlBody = htmlReceipt
        .Send
    End With
    
    Set cdoMsg = Nothing
    Set cdoConf = Nothing
    
    MsgBox "Email sent successfully!" & vbCrLf & "To: " & custEmail, vbInformation
    Exit Sub
    
ErrorHandler:
    MsgBox "Email Error: " & Err.Description, vbCritical
End Sub
Private Sub btnExportPDF_Click()
    On Error GoTo ErrorHandler
    
    Call RefreshOutputVars
    
    If m_CustomerID = "" Then
        MsgBox "No Customer Selected!", vbExclamation
        Exit Sub
    End If
    
    ' Folder
    Dim configFolder As String
    configFolder = ThisWorkbook.path & "\Configuration\"
    If Dir(configFolder, vbDirectory) = "" Then MkDir configFolder
    
    ' File Name
    Dim entryIDForFile As String
    Dim safeCustomerName As String
    Dim fileName As String
    Dim pdfPath As String
    
    If Trim(m_EntryID) <> "" Then
        entryIDForFile = m_EntryID
    Else
        entryIDForFile = m_CustomerID
    End If
    
    safeCustomerName = CleanFileName(m_CustomerName)
    If safeCustomerName = "" Then safeCustomerName = "Customer"
    
    fileName = entryIDForFile & "_" & safeCustomerName & "_" & Format(Now, "ddmmyyyy_hhmmss") & ".pdf"
    pdfPath = configFolder & fileName
    
    ' Generate & Export
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    Dim wsReceipt As Worksheet
    Set wsReceipt = GenerateReceiptSheet
    
    If wsReceipt Is Nothing Then
        MsgBox "Failed to create receipt!", vbCritical
        GoTo Cleanup
    End If
    
    wsReceipt.ExportAsFixedFormat Type:=xlTypePDF, fileName:=pdfPath, Quality:=xlQualityStandard
    
    Application.DisplayAlerts = False
    wsReceipt.Delete
    Application.DisplayAlerts = True
    
    ShellExecute 0, "open", pdfPath, vbNullString, vbNullString, 1
    
    MsgBox "PDF Opened!" & vbCrLf & "Saved at: " & pdfPath, vbInformation
    
Cleanup:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Exit Sub
    
ErrorHandler:
    MsgBox "PDF Error: " & Err.Description, vbCritical
    Resume Cleanup
End Sub
Private Sub btnAssignment_Click()
    Dim assignForm As frmAssignmentDashboard
    Set assignForm = New frmAssignmentDashboard
    assignForm.Show vbModal
    Set assignForm = Nothing
End Sub
Private Sub btnClear_Click()

    ' Form clear
    Call ClearForm

    ' ???? disable state
    Call ControlButtons(False)

End Sub
Private Sub btnCancel_Click()
    Unload Me
End Sub
'========================================
' GET SMTP SERVER
'========================================
Private Function GetSMTPServerAddress(ByVal serverName As String) As String
    Select Case Trim(LCase(serverName))
        Case "gmail": GetSMTPServerAddress = "smtp.gmail.com"
        Case "outlook": GetSMTPServerAddress = "smtp.office365.com"
        Case "yahoo": GetSMTPServerAddress = "smtp.mail.yahoo.com"
        Case "hotmail": GetSMTPServerAddress = "smtp.live.com"
        Case Else: GetSMTPServerAddress = serverName
    End Select
End Function

'========================================
' URL ENCODE
'========================================
'========================================
' URL ENCODE FOR OUTPUT (RENAME KARO)
'========================================
Private Function URLEncodeOutput(ByVal StringVal As String) As String
    Dim i As Integer, CharCode As Integer, char As String, OutString As String
    For i = 1 To Len(StringVal)
        char = Mid(StringVal, i, 1)
        CharCode = Asc(char)
        If (CharCode >= 48 And CharCode <= 57) Or (CharCode >= 65 And CharCode <= 90) Or _
           (CharCode >= 97 And CharCode <= 122) Or char = "-" Or char = "_" Or char = "." Or char = "~" Then
            OutString = OutString & char
        ElseIf char = " " Then
            OutString = OutString & "%20"
        ElseIf char = vbLf Or char = vbCr Then
            OutString = OutString & "%0A"
        Else
            OutString = OutString & "%" & Right("0" & Hex(CharCode), 2)
        End If
    Next i
    URLEncodeOutput = OutString
End Function

'========================================
' CLEAN FILE NAME
'========================================
Private Function CleanFileName(ByVal strName As String) As String
    Dim invalidChars As String, i As Integer
    invalidChars = "\/:*?""<>|"
    For i = 1 To Len(invalidChars)
        strName = Replace(strName, Mid(invalidChars, i, 1), "_")
    Next i
    CleanFileName = Trim(strName)
End Function

'========================================
' CLEAN TEXT
'========================================
Function CleanText(txt As String) As String
    txt = Replace(txt, vbCrLf, " ")
    txt = Replace(txt, vbLf, " ")
    txt = Replace(txt, vbCr, " ")
    CleanText = Trim(txt)
End Function

'========================================
' FORMAT MOBILE (10 digit clean)
'========================================
Private Function FormatMobileNumber(mobile As String) As String
    Dim i As Integer, ch As String, result As String
    For i = 1 To Len(mobile)
        ch = Mid(mobile, i, 1)
        If ch >= "0" And ch <= "9" Then result = result & ch
    Next i
    FormatMobileNumber = result
End Function
Private Function GenerateReceiptSheet() As Worksheet
    On Error GoTo ErrorHandler
    
    Dim wsCust As Worksheet, wsProd As Worksheet, wsAcc As Worksheet, wsPay As Worksheet, wsJob As Worksheet
    Dim wsSettings As Worksheet, wsSet As Worksheet, ws As Worksheet, wsConfig As Worksheet
    Dim r As Long, i As Long, termRow As Long, t As Integer
    Dim compAddr As String, compMob As String, compName As String, logoPath As String, compEmail As String
    Dim custEmail As String, custGST As String, pProblem As String
    Dim eID As String, payID As String, entryDate As String
    Dim entryRows As Collection
    Dim lastRowProd As Long, prodRow As Long, accRow As Long, lastRowAcc As Long, accCount As Long
    Dim lastRowPay As Long, payRow As Long, jobRow As Long
    Dim foundPayment As Boolean, termsFound As Boolean
    Dim payCount As Long, totalAmt As Double, pAmt As Double
    Dim totalCharge As Double, totalExpense As Double, totalReceived As Double, totalDiscount As Double
    Dim cat As String
    Dim verifyType As String, verifyName As String
    Dim pageSize As String, terms(1 To 8) As String
    Dim shpLogo As Shape, shpPhoto As Shape
    Dim showCustPhoto As String, showProdPhoto As String, showAccPhoto As String
    Dim custPhotoPath As String, prodPhotoPath As String, accPhotoPath As String
    
    ' PAGINATION VARIABLES
    Dim totalEntries As Long, entryNum As Long
    Dim entriesOnPage As Long
    Dim totalPages As Long, currentPage As Long
    Const MAX_ENTRIES = 2
    
    ' Delete old sheet
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    On Error Resume Next
    ThisWorkbook.Sheets("Temp_Receipt").Delete
    On Error GoTo 0
    
    ' Create new sheet
    Set ws = ThisWorkbook.Sheets.Add
    ws.name = "Temp_Receipt"
    
    ' Page Setup
    On Error Resume Next
    Set wsSettings = ThisWorkbook.Sheets("Settings")
    pageSize = UCase(wsSettings.Range("B20").value & "")
    showCustPhoto = UCase(wsSettings.Range("B21").value & "")
    showProdPhoto = UCase(wsSettings.Range("B22").value & "")
    showAccPhoto = UCase(wsSettings.Range("B23").value & "")
    custPhotoPath = wsSettings.Range("B3").value & ""
    prodPhotoPath = wsSettings.Range("B2").value & ""
    accPhotoPath = wsSettings.Range("B6").value & ""
    logoPath = wsSettings.Range("B11").value & ""
    On Error GoTo 0
    
    ' Setup paths
    If custPhotoPath <> "" And Right(custPhotoPath, 1) <> "\" Then custPhotoPath = custPhotoPath & "\"
    If prodPhotoPath <> "" And Right(prodPhotoPath, 1) <> "\" Then prodPhotoPath = prodPhotoPath & "\"
    If accPhotoPath <> "" And Right(accPhotoPath, 1) <> "\" Then accPhotoPath = accPhotoPath & "\"
    If prodPhotoPath = "" Then prodPhotoPath = ThisWorkbook.path & "\GLOBAL_SOFT_DATA\Images\Products\"
    If accPhotoPath = "" Then accPhotoPath = ThisWorkbook.path & "\GLOBAL_SOFT_DATA\Images\Accessories\"
    If logoPath = "" Then logoPath = ThisWorkbook.path & "\GLOBAL_SOFT_DATA\Images\Logo\"
    If Right(logoPath, 1) <> "\" Then logoPath = logoPath & "\"
    
    ' Check logo file
    Dim finalLogoPath As String
    finalLogoPath = logoPath & "LOGO.png"
    If Dir(finalLogoPath) = "" Then finalLogoPath = logoPath & "LOGO.jpg"
    If Dir(finalLogoPath) = "" Then finalLogoPath = logoPath & "logo.png"
    If Dir(finalLogoPath) = "" Then finalLogoPath = logoPath & "logo.jpg"
    
    If pageSize = "" Then pageSize = "A4"
    
    ' Calculate total pages
    Set entryRows = New Collection
    Set wsProd = ThisWorkbook.Sheets("Job_Product")
    lastRowProd = wsProd.Cells(wsProd.Rows.count, 1).End(xlUp).row
    For i = 2 To lastRowProd
        If Trim(UCase(wsProd.Cells(i, 2).value & "")) = Trim(UCase(m_CustomerID)) Then
            entryRows.Add i
        End If
    Next i
    
    totalEntries = entryRows.count
    totalPages = ((totalEntries + MAX_ENTRIES - 1) \ MAX_ENTRIES)
    If totalPages < 1 Then totalPages = 1
    
    ' Page Setup
    With ws.PageSetup
        If pageSize = "A5" Then
            .PaperSize = xlPaperA5
            .LeftMargin = Application.InchesToPoints(0.25)
            .RightMargin = Application.InchesToPoints(0.25)
            .TopMargin = Application.InchesToPoints(0.25)
            .BottomMargin = Application.InchesToPoints(0.25)
        Else
            .PaperSize = xlPaperA4
            .LeftMargin = Application.InchesToPoints(0.5)
            .RightMargin = Application.InchesToPoints(0.5)
            .TopMargin = Application.InchesToPoints(0.5)
            .BottomMargin = Application.InchesToPoints(0.5)
        End If
        .Orientation = xlPortrait
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
    End With
    
    ' Column widths
    If pageSize = "A5" Then
        ws.Columns("A").ColumnWidth = 2
        ws.Columns("B").ColumnWidth = 8
        ws.Columns("C").ColumnWidth = 20
        ws.Columns("D").ColumnWidth = 15
        ws.Columns("E").ColumnWidth = 15
        ws.Columns("F").ColumnWidth = 15
        ws.Columns("G").ColumnWidth = 15
        ws.Columns("H").ColumnWidth = 2
    Else
        ws.Columns("A").ColumnWidth = 2
        ws.Columns("B").ColumnWidth = 10
        ws.Columns("C").ColumnWidth = 22
        ws.Columns("D").ColumnWidth = 18
        ws.Columns("E").ColumnWidth = 18
        ws.Columns("F").ColumnWidth = 18
        ws.Columns("G").ColumnWidth = 18
        ws.Columns("H").ColumnWidth = 2
    End If
    
    ' Set worksheets
    Set wsCust = ThisWorkbook.Sheets("Customer_Master")
    Set wsAcc = ThisWorkbook.Sheets("Job_Accessory")
    Set wsPay = ThisWorkbook.Sheets("Payment_Master")
    Set wsJob = ThisWorkbook.Sheets("Job_Master")
    Set wsConfig = ThisWorkbook.Sheets("Software_Config")
    
    ' Load Terms & Conditions
    On Error Resume Next
    Set wsSet = ThisWorkbook.Sheets("Settings")
    termsFound = False
    If Not wsSet Is Nothing Then
        For termRow = 25 To 32
            If Trim(wsSet.Cells(termRow, 2).value & "") <> "" Then
                termsFound = True
                Exit For
            End If
        Next termRow
    End If
    On Error GoTo 0
    
    If Not termsFound Then
        terms(1) = "1. WARRANTY: Coverage as per manufacturer's policy only."
        terms(2) = "2. TIMELINE: Repair confirmation within 30 days."
        terms(3) = "3. RECEIPT: Original receipt mandatory for collection."
        terms(4) = "4. DATA: Not responsible for data loss during repair."
        terms(5) = "5. DAMAGE: Not liable for pre-existing damage."
        terms(6) = "6. REPLACEMENT: No refund if condition remains same."
        terms(7) = "7. VOID: Warranty void if seal removed/tampered."
        terms(8) = "8. LEGAL: All disputes subject to Nayagarh jurisdiction."
    End If
    
    ' ========== READ COMPANY DATA FROM Software_Config ==========
    compName = Trim(wsConfig.Range("B2").value & "")
    Dim addr1 As String, addr2 As String
    addr1 = Trim(wsConfig.Range("B3").value & "")
    addr2 = Trim(wsConfig.Range("B4").value & "")
    compMob = Trim(wsConfig.Range("B9").value & "")
    compEmail = Trim(wsConfig.Range("B10").value & "")
    
    ' Main loop
    r = 1
    currentPage = 0
    entriesOnPage = 0
    
    For entryNum = 1 To totalEntries
        
        ' NEW PAGE CHECK
                If entryNum > 1 And entriesOnPage >= MAX_ENTRIES Then
            Call AddPageFooter(ws, r, termsFound, wsSet, terms, verifyType, verifyName)
            r = r + 1
            ws.Rows(r).PageBreak = xlPageBreakManual
            entriesOnPage = 0
            r = r + 1
        End If
        
        currentPage = currentPage + 1
        
        ' PAGE HEADER
        If entriesOnPage = 0 Then
            
            ' ========== ROW 1: COMPANY NAME ==========
            With ws.Range("B" & r & ":G" & r)
                .Merge
                .value = compName
                .Font.Size = 26
                .Font.Bold = True
                .Font.name = "Arial"
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
            End With
            ws.Rows(r).RowHeight = 40
            
            ' LOGO - LEFT SIDE
            If Dir(finalLogoPath) <> "" Then
                On Error Resume Next
                Set shpLogo = ws.Shapes.AddPicture(finalLogoPath, msoFalse, msoTrue, 0, 0, -1, -1)
                If Not shpLogo Is Nothing Then
                    With shpLogo
                        .LockAspectRatio = msoTrue
                        .Height = 100
                        .Top = ws.Range("B" & r).Top + 5
                        .Left = ws.Range("A" & r).Left + 5
                        If .Width > 130 Then .Width = 130
                    End With
                End If
                On Error GoTo ErrorHandler
            End If
            r = r + 1

            ' ========== ROW 2: ADDRESS LINE 1 ==========
            With ws.Range("B" & r & ":G" & r)
                .Merge
                .value = addr1
                .Font.Size = 13
                .Font.name = "Arial"
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
            End With
            ws.Rows(r).RowHeight = 22
            r = r + 1
            
            ' ========== ROW 3: ADDRESS LINE 2 ==========
            If addr2 <> "" Then
                With ws.Range("B" & r & ":G" & r)
                    .Merge
                    .value = addr2
                    .Font.Size = 13
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlCenter
                End With
                ws.Rows(r).RowHeight = 22
                r = r + 1
            End If

            ' ========== ROW 4: MOBILE & EMAIL ==========
            With ws.Range("B" & r & ":G" & r)
                .Merge
                .value = "Mobile: " & compMob & "  |  Email: " & compEmail
                .Font.Size = 13
                .Font.Bold = True
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
            End With
            ws.Rows(r).RowHeight = 26
            r = r + 1
            
            ' ========== ROW 5: WARRANTY HEADER - NO GAP ==========
            With ws.Range("B" & r & ":G" & r)
                .Merge
                .value = "WARRANTY AND SERVICE REPORT"
                .Font.Bold = True
                .Font.Size = 16
                .Font.Color = RGB(255, 255, 255)
                .Interior.Color = RGB(0, 128, 0)
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
            End With
            ws.Rows(r).RowHeight = 32
            r = r + 1
            
            ' ========== ROW 6: CUSTOMER ID HEADER - NO GAP ==========
            With ws.Range("B" & r & ":G" & r)
                .Merge
                .value = "CUSTOMER ID: " & m_CustomerID
                .Font.Bold = True
                .Font.Size = 13
                .Font.Color = RGB(255, 255, 255)
                .Interior.Color = RGB(0, 0, 128)
                .HorizontalAlignment = xlLeft
                .VerticalAlignment = xlCenter
            End With
            ws.Rows(r).RowHeight = 28
            
            ' CUSTOMER PHOTO - Positioned properly (moved down)
            If showCustPhoto = "YES" And custPhotoPath <> "" Then
                If Dir(custPhotoPath & m_Mobile & ".jpg") <> "" Then
                    On Error Resume Next
                    Set shpPhoto = ws.Shapes.AddPicture(custPhotoPath & m_Mobile & ".jpg", msoFalse, msoTrue, 0, 0, 100, 80)
                    If Not shpPhoto Is Nothing Then
                        With shpPhoto
                            .LockAspectRatio = msoTrue
                            .Width = 100
                            .Height = 80
                            .Top = ws.Range("F" & r).Top + 30
                            .Left = ws.Range("F" & r).Left + 40
                        End With
                    End If
                    On Error GoTo ErrorHandler
                End If
            End If
            r = r + 1

            ' ========== ROW 7: CUSTOMER NAME ==========
            ws.Range("B" & r).value = "Name:"
            ws.Range("B" & r).Font.Bold = True
            ws.Range("B" & r).Font.Size = 12
            ws.Range("C" & r & ":E" & r).Merge
            ws.Range("C" & r).value = m_CustomerName
            ws.Range("C" & r).Font.Size = 12
            r = r + 1

            ' ========== ROW 8: MOBILE (LEFT ALIGNED) ==========
            ws.Range("B" & r).value = "Mobile:"
            ws.Range("B" & r).Font.Bold = True
            ws.Range("B" & r).Font.Size = 12
            ws.Range("C" & r).value = m_Mobile
            ws.Range("C" & r).Font.Size = 12
            ws.Range("C" & r).HorizontalAlignment = xlLeft
            r = r + 1
            
            ' Get customer details
            custEmail = ""
            custGST = ""
            Dim custAddress As String
            custAddress = ""
            For i = 2 To wsCust.Cells(wsCust.Rows.count, 1).End(xlUp).row
                If UCase(Trim(wsCust.Cells(i, 1).value & "")) = UCase(m_CustomerID) Then
                    custEmail = wsCust.Cells(i, 5).value & ""
                    custGST = wsCust.Cells(i, 6).value & ""
                    custAddress = wsCust.Cells(i, 4).value & ""
                    Exit For
                End If
            Next i
            
            ' ========== ROW 9: EMAIL ==========
            If custEmail <> "" Then
                ws.Range("B" & r).value = "Email:"
                ws.Range("B" & r).Font.Bold = True
                ws.Range("B" & r).Font.Size = 12
                ws.Range("C" & r & ":E" & r).Merge
                ws.Range("C" & r).value = custEmail
                ws.Range("C" & r).Font.Size = 12
                r = r + 1
            End If

            ' ========== ROW 10: ADDRESS (CENTERED) ==========
            If custAddress <> "" Then
    
    ws.Range("B" & r).value = "Address:"
    ws.Range("B" & r).Font.Bold = True
    ws.Range("B" & r).Font.Size = 12
    
    ws.Range("C" & r & ":G" & r).Merge
    ws.Range("C" & r).value = Replace(Replace(custAddress, vbCrLf, ", "), vbLf, ", ")
    ws.Range("C" & r).Font.Size = 12
    ws.Range("C" & r).HorizontalAlignment = xlLeft
    
    ws.Rows(r).RowHeight = 22
    
    r = r + 1
End If
            
            ' ========== ROW 11: GST ==========
            If custGST <> "" Then
                ws.Range("B" & r).value = "GST:"
                ws.Range("B" & r).Font.Bold = True
                ws.Range("B" & r).Font.Size = 12
                ws.Range("C" & r).value = custGST
                ws.Range("C" & r).Font.Size = 12
                r = r + 1
            End If
            
            r = r + 1
        End If
        
        ' ========== ENTRY DATA ==========
        entriesOnPage = entriesOnPage + 1
        prodRow = entryRows(entryNum)
        eID = wsProd.Cells(prodRow, 1).value & ""
        
        ' Get Entry Date and Verify Info from Job_Master
        entryDate = ""
        verifyType = ""
        verifyName = ""
        For i = 2 To wsJob.Cells(wsJob.Rows.count, 1).End(xlUp).row
            If Trim(UCase(wsJob.Cells(i, 1).value & "")) = Trim(UCase(eID)) Then
                entryDate = wsJob.Cells(i, 8).value & ""
                verifyType = wsJob.Cells(i, 6).value & ""
                verifyName = wsJob.Cells(i, 7).value & ""
                If entryDate = "" Then entryDate = Format(Date, "dd-MMM-yyyy")
                Exit For
            End If
        Next i
        If entryDate = "" Then entryDate = Format(Date, "dd-MMM-yyyy")

        ' Entry Header - LEFT ALIGNED
        With ws.Range("B" & r & ":G" & r)
            .Merge
            .value = "ENTRY ID: " & eID & "  |  Date: " & entryDate
            .Font.Bold = True
            .Font.Size = 13
            .Font.Color = RGB(255, 255, 255)
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter
            If UCase(Left(eID, 3)) = "WAR" Then
                .Interior.Color = RGB(200, 0, 0)
            ElseIf UCase(Left(eID, 3)) = "SER" Then
                .Interior.Color = RGB(255, 165, 0)
            Else
                .Interior.Color = RGB(0, 100, 150)
            End If
        End With
        ws.Rows(r).RowHeight = 28
        r = r + 1

        ' Product Details
        ws.Range("B" & r).value = "Product:"
        ws.Range("B" & r).Font.Bold = True
        ws.Range("B" & r).Font.Size = 11
        ws.Range("C" & r & ":D" & r).Merge
        ws.Range("C" & r).value = wsProd.Cells(prodRow, 6).value
        ws.Range("C" & r).Font.Size = 11
        ws.Range("E" & r).value = "Company:"
        ws.Range("E" & r).Font.Bold = True
        ws.Range("E" & r).Font.Size = 11
        ws.Range("F" & r & ":G" & r).Merge
        ws.Range("F" & r).value = wsProd.Cells(prodRow, 7).value
        ws.Range("F" & r).Font.Size = 11
        r = r + 1
        
        ws.Range("B" & r).value = "Model:"
        ws.Range("B" & r).Font.Bold = True
        ws.Range("B" & r).Font.Size = 11
        ws.Range("C" & r).value = wsProd.Cells(prodRow, 8).value
        ws.Range("C" & r).Font.Size = 11
        ws.Range("D" & r).value = "Serial:"
        ws.Range("D" & r).Font.Bold = True
        ws.Range("D" & r).Font.Size = 11
        ws.Range("E" & r & ":F" & r).Merge
        ws.Range("E" & r).value = wsProd.Cells(prodRow, 9).value
        ws.Range("E" & r).Font.Size = 11
        r = r + 1

        ' Problem
        pProblem = wsProd.Cells(prodRow, 13).value & ""
        If pProblem <> "" Then
            With ws.Range("B" & r & ":G" & r)
                .Merge
                .value = "Problem: " & pProblem
                .Font.Color = RGB(200, 0, 0)
                .Font.Size = 11
            End With
            r = r + 1
        End If

        ' PRODUCT PHOTO
        If showProdPhoto = "YES" And prodPhotoPath <> "" Then
            For i = 1 To 4
                If Dir(prodPhotoPath & eID & "\Photo" & i & ".jpg") <> "" Then
                    On Error Resume Next
                    Set shpPhoto = ws.Shapes.AddPicture(prodPhotoPath & eID & "\Photo" & i & ".jpg", msoFalse, msoTrue, 0, 0, 90, 90)
                    If Not shpPhoto Is Nothing Then
                        With shpPhoto
                            .LockAspectRatio = msoTrue
                            .Width = 90
                            .Height = 90
                            .Top = ws.Range("G" & (r - 3)).Top + 6
                            .Left = ws.Range("G" & (r - 3)).Left - 5
                        End With
                    End If
                    On Error GoTo ErrorHandler
                    Exit For
                End If
            Next i
        End If
        
                ' Accessories
        lastRowAcc = wsAcc.Cells(wsAcc.Rows.count, 1).End(xlUp).row
        accCount = 0
        For accRow = 2 To lastRowAcc
            If Trim(UCase(wsAcc.Cells(accRow, 1).value & "")) = Trim(UCase(eID)) Then accCount = accCount + 1
        Next accRow
        
        If accCount > 0 Then
            With ws.Range("B" & r & ":G" & r)
                .Merge
                .value = "ACCESSORIES:"
                .Font.Color = RGB(0, 0, 150)
                .Font.Bold = True
                .Font.Size = 11
            End With
            r = r + 1
            
                        ' Header row
            ws.Range("B" & r).value = "S.No"
            ws.Range("B" & r).Font.Bold = True
            ws.Range("B" & r).Font.Size = 10
            ws.Range("B" & r).HorizontalAlignment = xlCenter
            
            ws.Range("C" & r).value = "Accessory Name"
            ws.Range("C" & r).Font.Bold = True
            ws.Range("C" & r).Font.Size = 10
            
            ws.Range("D" & r).value = "Status"          ' <=== "Brand" ?? ???
            ws.Range("D" & r).Font.Bold = True
            ws.Range("D" & r).Font.Size = 10
            
            ws.Range("E" & r).value = "Serial No"        ' <=== Serial
            ws.Range("E" & r).Font.Bold = True
            ws.Range("E" & r).Font.Size = 10
            
            ws.Range("F" & r & ":G" & r).Merge
            ws.Range("F" & r).value = "Working"         ' <=== "Received" ?? ???
            ws.Range("F" & r).Font.Bold = True
            ws.Range("F" & r).Font.Size = 10
            ws.Range("F" & r).HorizontalAlignment = xlLeft
            r = r + 1
            
            Dim accIndex As Integer
            accIndex = 1
            For accRow = 2 To lastRowAcc
                If Trim(UCase(wsAcc.Cells(accRow, 1).value & "")) = Trim(UCase(eID)) Then
                    ws.Rows(r).RowHeight = 26
                    
                    ' Accessory Photo - BIGGER
                    If showAccPhoto = "YES" And accPhotoPath <> "" Then
                        If Dir(accPhotoPath & eID & "\Acc" & accIndex & ".jpg") <> "" Then
                            On Error Resume Next
                            Set shpPhoto = ws.Shapes.AddPicture(accPhotoPath & eID & "\Acc" & accIndex & ".jpg", msoFalse, msoTrue, 0, 0, 40, 30)
                            If Not shpPhoto Is Nothing Then
                                With shpPhoto
                                    .LockAspectRatio = msoTrue
                                    .Width = 40
                                    .Height = 30
                                    .Top = ws.Range("B" & r).Top - 15
                                    .Left = ws.Range("B" & r).Left + 5
                                End With
                            End If
                            On Error GoTo ErrorHandler
                        End If
                    End If
                                        ' Accessory Details (??? Columns)
                    ws.Range("B" & r).value = accIndex
                    ws.Range("B" & r).Font.Size = 10
                    ws.Range("B" & r).HorizontalAlignment = xlCenter
                    
                    ws.Range("C" & r).value = wsAcc.Cells(accRow, 4).value   ' D: Name
                    ws.Range("C" & r).Font.Size = 10
                    
                    ws.Range("D" & r).value = wsAcc.Cells(accRow, 5).value   ' E: Status
                    ws.Range("D" & r).Font.Size = 10
                    
                    ws.Range("E" & r).value = wsAcc.Cells(accRow, 6).value   ' F: Serial
                    ws.Range("E" & r).Font.Size = 10
                    
                    ws.Range("F" & r & ":G" & r).Merge
                    ws.Range("F" & r).value = wsAcc.Cells(accRow, 7).value   ' G: Working
                    ws.Range("F" & r).Font.Size = 10
                    ws.Range("F" & r).HorizontalAlignment = xlLeft
                    
                    r = r + 1
                    accIndex = accIndex + 1
                    If accIndex > 4 Then Exit For
                End If
            Next accRow
        Else
            With ws.Range("B" & r & ":G" & r)
                .Merge
                .value = "ACCESSORIES: NO ACCESSORIES RECEIVED"
                .Font.Color = RGB(128, 128, 128)
                .Font.Italic = True
                .Font.Size = 10
            End With
            r = r + 1
        End If

       
    
        r = r + 1
        
    Next entryNum
    
        ' ===== PAYMENT DETAILS - ALL ENTRIES COMBINED =====
    lastRowPay = wsPay.Cells(wsPay.Rows.count, 1).End(xlUp).row
    payCount = 0
    totalCharge = 0
    totalExpense = 0
    totalReceived = 0
    totalDiscount = 0
    
    ' Count total active payments across ALL entries
    For entryNum = 1 To totalEntries
        prodRow = entryRows(entryNum)
        eID = wsProd.Cells(prodRow, 1).value & ""
        For payRow = 2 To lastRowPay
            If Trim(UCase(wsPay.Cells(payRow, 2).value & "")) = Trim(UCase(eID)) Then
                If UCase(Trim(wsPay.Cells(payRow, 4).value & "")) <> "DELETED" And Trim(wsPay.Cells(payRow, 6).value & "") = "" Then
                    payCount = payCount + 1
                End If
            End If
        Next
    Next
    
    If payCount > 0 Then
        ' Payment Header
        With ws.Range("B" & r & ":G" & r)
            .Merge
            .value = "PAYMENT DETAILS"
            .Font.Color = RGB(255, 255, 255)
            .Font.Bold = True
            .Font.Size = 11
            .Interior.Color = RGB(0, 100, 0)
            .HorizontalAlignment = xlLeft
        End With
        r = r + 1
        
        ' Column Headers
        ws.Range("B" & r).value = "PayID"
        ws.Range("B" & r).Font.Bold = True
        ws.Range("B" & r).Font.Size = 9
        ws.Range("C" & r).value = "Type"
        ws.Range("C" & r).Font.Bold = True
        ws.Range("C" & r).Font.Size = 9
        ws.Range("D" & r).value = "Category"
        ws.Range("D" & r).Font.Bold = True
        ws.Range("D" & r).Font.Size = 9
        ws.Range("E" & r).value = "Mode"
        ws.Range("E" & r).Font.Bold = True
        ws.Range("E" & r).Font.Size = 9
        ws.Range("F" & r).value = "Qty"
        ws.Range("F" & r).Font.Bold = True
        ws.Range("F" & r).Font.Size = 9
        ws.Range("F" & r).HorizontalAlignment = xlCenter
        ws.Range("G" & r).value = "Amount"
        ws.Range("G" & r).Font.Bold = True
        ws.Range("G" & r).Font.Size = 9
        ws.Range("G" & r).HorizontalAlignment = xlRight
        r = r + 1
        
        ' Payment Rows for ALL entries
        For entryNum = 1 To totalEntries
            prodRow = entryRows(entryNum)
            eID = wsProd.Cells(prodRow, 1).value & ""
            
            For payRow = 2 To lastRowPay
                If Trim(UCase(wsPay.Cells(payRow, 2).value & "")) = Trim(UCase(eID)) Then
                    If UCase(Trim(wsPay.Cells(payRow, 4).value & "")) <> "DELETED" And Trim(wsPay.Cells(payRow, 6).value & "") = "" Then
                        ws.Range("B" & r).value = wsPay.Cells(payRow, 1).value
                        ws.Range("B" & r).Font.Size = 9
                        ws.Range("C" & r).value = wsPay.Cells(payRow, 8).value
                        ws.Range("C" & r).Font.Size = 9
                        ws.Range("D" & r).value = wsPay.Cells(payRow, 9).value
                        ws.Range("D" & r).Font.Size = 9
                        ws.Range("E" & r).value = wsPay.Cells(payRow, 15).value
                        ws.Range("E" & r).Font.Size = 9
                        ws.Range("F" & r).value = wsPay.Cells(payRow, 14).value
                        ws.Range("F" & r).Font.Size = 9
                        ws.Range("F" & r).HorizontalAlignment = xlCenter
                        
                        If IsNumeric(wsPay.Cells(payRow, 16).value) Then
                            pAmt = CDbl(wsPay.Cells(payRow, 16).value)
                        Else
                            pAmt = 0
                        End If
                        
                        cat = UCase(Trim(wsPay.Cells(payRow, 9).value & ""))
                        If cat = "CHARGE" Or cat = "PART_USED" Or cat = "PART_REPLACED" Then totalCharge = totalCharge + pAmt
                        If cat = "EXPENSE" Then totalExpense = totalExpense + pAmt
                        If cat = "RECEIVED" Or cat = "ADVANCE" Then totalReceived = totalReceived + pAmt
                        If cat = "DISCOUNT" Then totalDiscount = totalDiscount + pAmt
                        
                        ws.Range("G" & r).value = Format(pAmt, "0.00")
                        ws.Range("G" & r).Font.Size = 9
                        ws.Range("G" & r).HorizontalAlignment = xlRight
                        r = r + 1
                    End If
                End If
            Next
        Next
        
        ' Grand Summary
        r = r + 1
        Dim summaryText As String
        summaryText = "CHARGE: Rs." & Format(totalCharge, "0.00")
        If totalExpense > 0 Then summaryText = summaryText & "  |  EXPENSE: Rs." & Format(totalExpense, "0.00")
        summaryText = summaryText & "  |  RECEIVED: Rs." & Format(totalReceived, "0.00")
        If totalDiscount > 0 Then summaryText = summaryText & "  |  DISCOUNT: Rs." & Format(totalDiscount, "0.00")
        summaryText = summaryText & "  |  PENDING: Rs." & Format((totalCharge + totalExpense) - totalReceived - totalDiscount, "0.00")
        
        With ws.Range("B" & r & ":G" & r)
            .Merge
            .value = summaryText
            .Font.Bold = True
            .Font.Size = 10
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Interior.Color = RGB(240, 248, 255)
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
            .Borders.Color = RGB(0, 0, 128)
        End With
        ws.Rows(r).RowHeight = 24
        r = r + 1
    Else
        With ws.Range("B" & r & ":G" & r)
            .Merge
            .value = "Payment: No transactions found"
            .Font.Color = RGB(128, 128, 128)
            .Font.Size = 11
        End With
        r = r + 1
    End If
    
    r = r + 1
    
    ' Add FOOTER to last page
    Call AddPageFooter(ws, r, termsFound, wsSet, terms, verifyType, verifyName)
    
    
   
    ws.Activate
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Set GenerateReceiptSheet = ws
    Exit Function
    
ErrorHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "Error: " & Err.Description, vbCritical
    Set GenerateReceiptSheet = Nothing

End Function

Private Sub AddPageFooter(ws As Worksheet, ByRef r As Long, termsFound As Boolean, wsSet As Worksheet, terms() As String, Optional vType As String = "", Optional vName As String = "")
    Dim termRow As Long, t As Integer
    
    r = r + 1
    
        ' ===== VERIFIED BY SECTION (Above Terms & Conditions) =====
    If Trim(vType) <> "" Or Trim(vName) <> "" Then
        With ws.Range("B" & r & ":G" & r)
            .Merge
            .value = "Verified By: " & vName & " (" & vType & ")"
            .Font.Bold = True
            .Font.Size = 11
            .Font.Color = RGB(0, 0, 128)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Interior.Color = RGB(230, 240, 255)
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
        End With
        ws.Rows(r).RowHeight = 24
        r = r + 2
    End If
    
    ' TERMS & CONDITIONS
    With ws.Range("B" & r & ":G" & r)
        .Merge
        .value = "TERMS & CONDITIONS"
        .Font.Bold = True
        .Font.Size = 12
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(128, 0, 0)
        .HorizontalAlignment = xlCenter
    End With
    ws.Rows(r).RowHeight = 24
    r = r + 1
    
    ' Terms Content
    If termsFound Then
        For termRow = 25 To 32
            If Trim(wsSet.Cells(termRow, 2).value & "") <> "" Then
                With ws.Range("B" & r & ":G" & r)
                    .Merge
                    .value = wsSet.Cells(termRow, 2).value
                    .Font.Size = 9
                    .WrapText = True
                End With
                ws.Rows(r).RowHeight = 16
                r = r + 1
            End If
        Next termRow
    Else
        For t = 1 To 8
            With ws.Range("B" & r & ":G" & r)
                .Merge
                .value = terms(t)
                .Font.Size = 9
                .WrapText = True
            End With
            ws.Rows(r).RowHeight = 16
            r = r + 1
        Next t
    End If
    
    ' ===== SIGNATURES SECTION =====
    r = r + 2
    
    ' Customer Signature Line (Left Side)
    With ws.Range("B" & r & ":C" & r)
        .Merge
        .value = ""
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeTop).Weight = xlThin
    End With
    
    ' Authorized Signature Line (Right Side)
    With ws.Range("F" & r & ":G" & r)
        .Merge
        .value = ""
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeTop).Weight = xlThin
    End With
    
    r = r + 1
    
    ' Customer Signature Text
    With ws.Range("B" & r & ":C" & r)
        .Merge
        .value = "Customer Signature"
        .Font.Bold = True
        .Font.Size = 11
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlTop
    End With
    
    ' Authorized Signature Text
    With ws.Range("F" & r & ":G" & r)
        .Merge
        .value = "Authorized Signature"
        .Font.Bold = True
        .Font.Size = 11
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlTop
    End With
    
    ' For GLOBAL IT SOLUTIONS
    r = r + 1
    
    With ws.Range("F" & r & ":G" & r)
        .Merge
        .value = "For GLOBAL IT SOLUTIONS"
        .Font.Italic = True
        .Font.Size = 11
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
    
    ' THANK YOU
    r = r + 2
    With ws.Range("B" & r & ":G" & r)
        .Merge
        .value = "*** Thank You for choosing US ***"
        .Font.Bold = True
        .Font.Size = 12
        .Font.name = "Arial"
        .Font.Color = RGB(0, 100, 0)
        .HorizontalAlignment = xlCenter
        .Interior.Color = RGB(255, 250, 205)
    End With
    ws.Rows(r).RowHeight = 26
    
    r = r + 2
End Sub

'========================================
' GENERATE RECEIPT AS HTML FOR EMAIL
'========================================
Private Function GenerateReceiptHTML() As String
    On Error GoTo ErrorHandler
    
    Dim wsCust As Worksheet, wsProd As Worksheet, wsAcc As Worksheet
    Dim wsPay As Worksheet, wsJob As Worksheet, wsConfig As Worksheet
    Dim html As String, i As Long, entryNum As Long
    Dim compName As String, addr1 As String, addr2 As String
    Dim compMob As String, compEmail As String
    Dim entryRows As Collection, prodRow As Long, eID As String
    Dim custAddress As String, custGST As String, custEmail As String
    Dim totalEntries As Long
    
    ' Set worksheets
    Set wsCust = ThisWorkbook.Sheets("Customer_Master")
    Set wsProd = ThisWorkbook.Sheets("Job_Product")
    Set wsAcc = ThisWorkbook.Sheets("Job_Accessory")
    Set wsPay = ThisWorkbook.Sheets("Payment_Master")
    Set wsJob = ThisWorkbook.Sheets("Job_Master")
    Set wsConfig = ThisWorkbook.Sheets("Software_Config")
    
    ' Company Info from Software_Config
    compName = Trim(wsConfig.Range("B2").value & "")
    addr1 = Trim(wsConfig.Range("B3").value & "")
    addr2 = Trim(wsConfig.Range("B4").value & "")
    compMob = Trim(wsConfig.Range("B9").value & "")
    compEmail = Trim(wsConfig.Range("B10").value & "")
    
    If compName = "" Then compName = "GLOBAL IT SOLUTIONS"
    
    ' Customer Details
    custAddress = ""
    custGST = ""
    custEmail = ""
    For i = 2 To wsCust.Cells(wsCust.Rows.count, 1).End(xlUp).row
        If UCase(Trim(wsCust.Cells(i, 1).value & "")) = UCase(m_CustomerID) Then
            custAddress = wsCust.Cells(i, 4).value & ""
            custEmail = wsCust.Cells(i, 5).value & ""
            custGST = wsCust.Cells(i, 6).value & ""
            Exit For
        End If
    Next i
    
    ' Get all entries for this customer
    Set entryRows = New Collection
    For i = 2 To wsProd.Cells(wsProd.Rows.count, 1).End(xlUp).row
        If Trim(UCase(wsProd.Cells(i, 2).value & "")) = Trim(UCase(m_CustomerID)) Then
            entryRows.Add i
        End If
    Next i
    totalEntries = entryRows.count
    
    ' ===== BUILD HTML =====
    html = "<!DOCTYPE html>"
    html = html & "<html><head>"
    html = html & "<style>"
    html = html & "body { font-family: Arial, sans-serif; margin: 20px; background: #f5f5f5; }"
    html = html & ".receipt { max-width: 800px; margin: 0 auto; background: white; border: 2px solid #333; }"
    html = html & ".header { background: #006400; color: white; text-align: center; padding: 15px; }"
    html = html & ".header h1 { margin: 0; font-size: 24px; }"
    html = html & ".company-info { text-align: center; padding: 10px; border-bottom: 1px solid #ccc; }"
    html = html & ".customer-header { background: #000080; color: white; padding: 10px; font-weight: bold; }"
    html = html & ".section { padding: 10px; border-bottom: 1px solid #ddd; }"
    html = html & ".entry-header { background: #c80000; color: white; padding: 8px; font-weight: bold; margin-top: 10px; }"
    html = html & ".entry-header.service { background: #ffa500; }"
    html = html & "table { width: 100%; border-collapse: collapse; margin: 10px 0; }"
    html = html & "th, td { padding: 8px; text-align: left; border: 1px solid #ddd; }"
    html = html & "th { background: #f0f0f0; font-weight: bold; }"
    html = html & ".payment { background: #006400; color: white; padding: 8px; }"
    html = html & ".terms { background: #800000; color: white; padding: 10px; font-weight: bold; }"
    html = html & ".footer { text-align: center; padding: 15px; background: #ffffcd; color: #006400; font-weight: bold; }"
    html = html & ".label { font-weight: bold; color: #333; }"
    html = html & "</style></head><body>"
    
    ' Receipt Container
    html = html & "<div class='receipt'>"
    
    ' Company Header
    html = html & "<div class='header'>"
    html = html & "<h1>" & compName & "</h1>"
    html = html & "</div>"
   ' Next entryNum
    
    ' Company Info
    html = html & "<div class='company-info'>"
    If addr1 <> "" Then html = html & "<p>" & addr1 & "</p>"
    If addr2 <> "" Then html = html & "<p>" & addr2 & "</p>"
    html = html & "<p><strong>Mobile: " & compMob & " | Email: " & compEmail & "</strong></p>"
    html = html & "</div>"
    
    ' Title
    html = html & "<div class='header'>"
    html = html & "<h2>WARRANTY AND SERVICE REPORT</h2>"
    html = html & "</div>"
    
    ' Customer Header
    html = html & "<div class='customer-header'>"
    html = html & "CUSTOMER ID: " & m_CustomerID
    html = html & "</div>"
    
    ' Customer Details
    html = html & "<div class='section'>"
    html = html & "<table>"
    html = html & "<tr><td class='label'>Name:</td><td>" & m_CustomerName & "</td></tr>"
    html = html & "<tr><td class='label'>Mobile:</td><td>" & m_Mobile & "</td></tr>"
    If custEmail <> "" Then html = html & "<tr><td class='label'>Email:</td><td>" & custEmail & "</td></tr>"
    If custAddress <> "" Then html = html & "<tr><td class='label'>Address:</td><td>" & Replace(custAddress, vbCrLf, ", ") & "</td></tr>"
    If custGST <> "" Then html = html & "<tr><td class='label'>GST:</td><td>" & custGST & "</td></tr>"
    html = html & "</table>"
    html = html & "</div>"
    Next entryNum
    
    ' Entries Loop
    For entryNum = 1 To totalEntries
        prodRow = entryRows(entryNum)
        eID = wsProd.Cells(prodRow, 1).value & ""
        
        ' Entry Type & Color
        Dim entryType As String, entryClass As String
        If UCase(Left(eID, 3)) = "WAR" Then
            entryType = "WARRANTY"
            entryClass = "entry-header"
        ElseIf UCase(Left(eID, 3)) = "SER" Then
            entryType = "SERVICE"
            entryClass = "entry-header service"
        Else
            entryType = "ENTRY"
            entryClass = "entry-header"
        End If
        
        ' Entry Date from Job_Master
        Dim entryDate As String, verifyType As String, verifyName As String
        entryDate = ""
        verifyType = ""
        verifyName = ""
        For i = 2 To wsJob.Cells(wsJob.Rows.count, 1).End(xlUp).row
            If Trim(UCase(wsJob.Cells(i, 1).value & "")) = Trim(UCase(eID)) Then
                entryDate = wsJob.Cells(i, 8).value & ""
                verifyType = wsJob.Cells(i, 6).value & ""
                verifyName = wsJob.Cells(i, 7).value & ""
                Exit For
            End If
        Next i
        If entryDate = "" Then entryDate = Format(Date, "dd-MMM-yyyy")
        
        ' Entry Header
        html = html & "<div class='" & entryClass & "'>"
        html = html & entryType & " ID: " & eID & " | Date: " & entryDate
        html = html & "</div>"
        
        ' Product Details
        html = html & "<div class='section'>"
        html = html & "<table>"
        html = html & "<tr><td class='label'>Product:</td><td>" & wsProd.Cells(prodRow, 6).value & "</td>"
        html = html & "<td class='label'>Company:</td><td>" & wsProd.Cells(prodRow, 7).value & "</td></tr>"
        html = html & "<tr><td class='label'>Model:</td><td>" & wsProd.Cells(prodRow, 8).value & "</td>"
        html = html & "<td class='label'>Serial:</td><td>" & wsProd.Cells(prodRow, 9).value & "</td></tr>"
        
        ' Problem
        Dim prob As String
        prob = wsProd.Cells(prodRow, 13).value & ""
        If prob <> "" Then
            html = html & "<tr><td class='label'>Problem:</td><td colspan='3' style='color: #c80000;'>" & prob & "</td></tr>"
        End If
        html = html & "</table>"
        
                ' Accessories
        lastRowAcc = wsAcc.Cells(wsAcc.Rows.count, 1).End(xlUp).row
        accCount = 0
        For accRow = 2 To lastRowAcc
            If Trim(UCase(wsAcc.Cells(accRow, 1).value & "")) = Trim(UCase(eID)) Then accCount = accCount + 1
        Next accRow
        
        If accCount > 0 Then
            With ws.Range("B" & r & ":G" & r)
                .Merge
                .value = "ACCESSORIES:"
                .Font.Color = RGB(0, 0, 150)
                .Font.Bold = True
                .Font.Size = 11
            End With
            r = r + 1
            
            ' Header row
            ws.Range("B" & r).value = "S.No"
            ws.Range("B" & r).Font.Bold = True
            ws.Range("B" & r).Font.Size = 10
            ws.Range("B" & r).HorizontalAlignment = xlCenter
            
            ws.Range("C" & r).value = "Accessory Name"
            ws.Range("C" & r).Font.Bold = True
            ws.Range("C" & r).Font.Size = 10
            
            ws.Range("D" & r).value = "Brand"
            ws.Range("D" & r).Font.Bold = True
            ws.Range("D" & r).Font.Size = 10
            
            ws.Range("E" & r).value = "Serial"
            ws.Range("E" & r).Font.Bold = True
            ws.Range("E" & r).Font.Size = 10
            
            ws.Range("F" & r & ":G" & r).Merge
            ws.Range("F" & r).value = "Received"
            ws.Range("F" & r).Font.Bold = True
            ws.Range("F" & r).Font.Size = 10
            ws.Range("F" & r).HorizontalAlignment = xlLeft
            r = r + 1
            
            Dim accIndex As Integer
            accIndex = 1
            For accRow = 2 To lastRowAcc
                If Trim(UCase(wsAcc.Cells(accRow, 1).value & "")) = Trim(UCase(eID)) Then
                    ws.Rows(r).RowHeight = 26
                    
                    ' Accessory Photo
                    If showAccPhoto = "YES" And accPhotoPath <> "" Then
                        If Dir(accPhotoPath & eID & "\Acc" & accIndex & ".jpg") <> "" Then
                            On Error Resume Next
                            Set shpPhoto = ws.Shapes.AddPicture(accPhotoPath & eID & "\Acc" & accIndex & ".jpg", msoFalse, msoTrue, 0, 0, 40, 30)
                            If Not shpPhoto Is Nothing Then
                                With shpPhoto
                                    .LockAspectRatio = msoTrue
                                    .Width = 40
                                    .Height = 30
                                    .Top = ws.Range("B" & r).Top - 15
                                    .Left = ws.Range("B" & r).Left + 5
                                End With
                            End If
                            On Error GoTo ErrorHandler
                        End If
                    End If
                    
                    ' Accessory Details
                    ws.Range("B" & r).value = accIndex
                    ws.Range("B" & r).Font.Size = 10
                    ws.Range("B" & r).HorizontalAlignment = xlCenter
                    
                    ws.Range("C" & r).value = wsAcc.Cells(accRow, 4).value
                    ws.Range("C" & r).Font.Size = 10
                    
                    ws.Range("D" & r).value = wsAcc.Cells(accRow, 5).value
                    ws.Range("D" & r).Font.Size = 10
                    
                    ws.Range("E" & r).value = wsAcc.Cells(accRow, 6).value
                    ws.Range("E" & r).Font.Size = 10
                    
                    ws.Range("F" & r & ":G" & r).Merge
                    ws.Range("F" & r).value = wsAcc.Cells(accRow, 7).value
                    ws.Range("F" & r).Font.Size = 10
                    ws.Range("F" & r).HorizontalAlignment = xlLeft
                    
                    r = r + 1
                    accIndex = accIndex + 1
                    If accIndex > 4 Then Exit For
                End If
            Next accRow
        Else
            With ws.Range("B" & r & ":G" & r)
                .Merge
                .value = "ACCESSORIES: NO ACCESSORIES RECEIVED"
                .Font.Color = RGB(128, 128, 128)
                .Font.Italic = True
                .Font.Size = 10
            End With
            r = r + 1
        End If
        
        ' Count valid payments
        For payRow = 2 To lastRowPay
            If Trim(UCase(wsPay.Cells(payRow, 2).value & "")) = Trim(UCase(eID)) Then
                If UCase(Trim(wsPay.Cells(payRow, 4).value & "")) <> "DELETED" And Trim(wsPay.Cells(payRow, 6).value & "") = "" Then
                    payCount = payCount + 1
                End If
            End If
        Next
        
        If payCount > 0 Then
                        html = html & "<div class='payment'>"
            html = html & "PAYMENT DETAILS"
            html = html & "</div>"
            
            html = html & "<table>"
            html = html & "<tr>"
            html = html & "<th>PayID</th>"
            html = html & "<th>Type</th>"
            html = html & "<th>Category</th>"
            html = html & "<th>Mode</th>"
            html = html & "<th>Qty</th>"
            html = html & "<th>Amount</th>"
            html = html & "</tr>"
            
            For payRow = 2 To lastRowPay
                If Trim(UCase(wsPay.Cells(payRow, 2).value & "")) = Trim(UCase(eID)) Then
                    If UCase(Trim(wsPay.Cells(payRow, 4).value & "")) <> "DELETED" And Trim(wsPay.Cells(payRow, 6).value & "") = "" Then
                        
                        If IsNumeric(wsPay.Cells(payRow, 16).value) Then
                            pAmt = CDbl(wsPay.Cells(payRow, 16).value)
                        Else
                            pAmt = 0
                        End If
                        cat = UCase(Trim(wsPay.Cells(payRow, 9).value & ""))
                        If cat <> "DISCUSSION" Then
                            totalAmt = totalAmt + pAmt
                        End If
                        
                        If cat = "CHARGE" Or cat = "PART_USED" Or cat = "PART_REPLACED" Then totalCharge = totalCharge + pAmt
                        If cat = "EXPENSE" Then totalExpense = totalExpense + pAmt
                        If cat = "RECEIVED" Or cat = "ADVANCE" Then totalReceived = totalReceived + pAmt
                        If cat = "DISCOUNT" Then totalDiscount = totalDiscount + pAmt
                        
                        html = html & "<tr>"
                        html = html & "<td>" & wsPay.Cells(payRow, 1).value & "</td>"
                        html = html & "<td>" & wsPay.Cells(payRow, 8).value & "</td>"
                        html = html & "<td>" & wsPay.Cells(payRow, 9).value & "</td>"
                        html = html & "<td>" & wsPay.Cells(payRow, 15).value & "</td>"
                        html = html & "<td>" & wsPay.Cells(payRow, 14).value & "</td>"
                        html = html & "<td><strong>Rs." & Format(pAmt, "0.00") & "</strong></td>"
                        html = html & "</tr>"
                    End If
                End If
            Next
            
                        html = html & "<tr style='background:#f0f0f0;'>"
            html = html & "<td colspan='5' style='text-align:right;'><strong>TOTAL AMOUNT</strong></td>"
            html = html & "<td><strong>Rs." & Format(totalAmt, "0.00") & "</strong></td>"
            html = html & "</tr>"
            html = html & "</table>"
            
            ' Payment Summary
            html = html & "<table style='margin-top:5px;'>"
            html = html & "<tr style='background:#000080; color:white;'><td colspan='2'><strong>PAYMENT SUMMARY</strong></td></tr>"
            html = html & "<tr><td>Total Charge:</td><td><strong>Rs." & Format(totalCharge, "0.00") & "</strong></td></tr>"
            html = html & "<tr><td>Total Expense:</td><td><strong>Rs." & Format(totalExpense, "0.00") & "</strong></td></tr>"
            html = html & "<tr><td>Total Received:</td><td><strong>Rs." & Format(totalReceived, "0.00") & "</strong></td></tr>"
            If totalDiscount > 0 Then html = html & "<tr><td>Total Discount:</td><td><strong>Rs." & Format(totalDiscount, "0.00") & "</strong></td></tr>"
            html = html & "<tr style='background:#ffe0e0; color:#c80000;'><td><strong>TOTAL PENDING:</strong></td><td><strong>Rs." & Format(totalCharge + totalExpense - totalReceived - totalDiscount, "0.00") & "</strong></td></tr>"
            html = html & "</table>"
        End If
        
        ' Verification
        If verifyType <> "" Or verifyName <> "" Then
            html = html & "<p style='background: #f0f8ff; padding: 8px; border: 2px solid #000080; text-align: center;'>"
            html = html & "<strong>Verified Type:</strong> " & verifyType & " | <strong>Verified By:</strong> " & verifyName
            html = html & "</p>"
        End If
        
        html = html & "</div>"
    Next entryNum
    
    ' Terms & Conditions
    html = html & "<div class='terms'>TERMS & CONDITIONS</div>"
    html = html & "<div class='section'>"
    html = html & "<ol>"
    html = html & "<li>WARRANTY: Coverage as per manufacturer's policy only.</li>"
    html = html & "<li>TIMELINE: Repair confirmation within 30 days.</li>"
    html = html & "<li>RECEIPT: Original receipt mandatory for collection.</li>"
    html = html & "<li>DATA: Not responsible for data loss during repair.</li>"
    html = html & "<li>DAMAGE: Not liable for pre-existing damage.</li>"
    html = html & "<li>REPLACEMENT: No refund if condition remains same.</li>"
    html = html & "<li>VOID: Warranty void if seal removed/tampered.</li>"
    html = html & "<li>LEGAL: All disputes subject to local jurisdiction.</li>"
    html = html & "</ol>"
    html = html & "</div>"
    
    ' Footer
    html = html & "<div class='footer'>"
    html = html & "*** Thank You for choosing " & compName & "! ***"
    html = html & "</div>"
    
    ' Close containers
    html = html & "</div>"  ' receipt
    html = html & "</body></html>"
    
    GenerateReceiptHTML = html
    Exit Function
    
ErrorHandler:
    GenerateReceiptHTML = ""
End Function

Private Function BuildWhatsAppMessage() As String
    On Error Resume Next
    
    Dim msg As String
    Dim wsCust As Worksheet, wsProd As Worksheet, wsAcc As Worksheet, wsJob As Worksheet
    Dim wsConfig As Worksheet
    Dim wsPay As Worksheet
    Dim i As Long, entryNum As Long
    Dim compName As String, compAddr1 As String, compAddr2 As String, compMob As String
    Dim entryRows As Collection
    Dim prodRow As Long, eID As String
    Dim custAddress As String, custEmail As String
    Dim lastRowPay As Long, payRow As Long
    Dim payCount As Integer
        Dim pAmt As Double, totalAmt As Double
    Dim totalCharge As Double, totalExpense As Double, totalReceived As Double, totalDiscount As Double
    Dim cat As String
    
    ' === SHEETS SET ===
    Set wsCust = ThisWorkbook.Sheets("Customer_Master")
    Set wsProd = ThisWorkbook.Sheets("Job_Product")
    Set wsAcc = ThisWorkbook.Sheets("Job_Accessory")
    Set wsJob = ThisWorkbook.Sheets("Job_Master")
    Set wsConfig = ThisWorkbook.Sheets("Software_Config")
    Set wsPay = ThisWorkbook.Sheets("Payment_Master")
    
    ' === COMPANY INFO ===
    compName = Trim(wsConfig.Range("B2").value & "")
    If compName = "" Then compName = "GLOBAL IT SOLUTIONS"
    
    compAddr1 = Trim(wsConfig.Range("B3").value & "")
    compAddr2 = Trim(wsConfig.Range("B4").value & "")
    If compAddr2 = "" Then compAddr2 = "Nayagarh-752069, Odisha"
    
    compMob = Trim(wsConfig.Range("B9").value & "")
    If compMob = "" Then compMob = "9777971045"
    
    ' === CUSTOMER DETAILS ===
    custAddress = ""
    custEmail = ""
    For i = 2 To wsCust.Cells(wsCust.Rows.count, 1).End(xlUp).row
        If UCase(Trim(wsCust.Cells(i, 1).value & "")) = UCase(m_CustomerID) Then
            custAddress = wsCust.Cells(i, 4).value & ""
            custEmail = wsCust.Cells(i, 5).value & ""
            Exit For
        End If
    Next i
    
    ' === BUILD MESSAGE (????????? ???????) ===
    ' NOTE: vbLf ???? ??? (????? ????), vbCrLf ?? ???? ??? (??? ??? ???? ??)
    
    msg = "*" & compName & "*" & vbLf & _
          compAddr1 & vbLf & _
          compAddr2 & vbLf & _
          "WhatsApp & Call: " & compMob & vbLf & _
          "------------------------------" & vbLf & _
          "*YOUR PRODUCT RECEIVED*" & vbLf & _
          "------------------------------" & vbLf & _
          "*ID:* " & m_CustomerID & vbLf & _
          "*Name:* " & m_CustomerName & vbLf & _
          "*Mobile:* " & m_Mobile
    
    ' Address (??? ?? ??)
    If custAddress <> "" Then
        msg = msg & vbLf & "*Address:* " & Replace(custAddress, vbCrLf, ", ")
    End If
    
    ' Email (??? ?? ??)
    If custEmail <> "" Then
        msg = msg & vbLf & "*Email:* " & custEmail
    End If
    
    ' === ENTRIES ===
    Set entryRows = New Collection
    For i = 2 To wsProd.Cells(wsProd.Rows.count, 1).End(xlUp).row
        If Trim(UCase(wsProd.Cells(i, 2).value & "")) = Trim(UCase(m_CustomerID)) Then
            entryRows.Add i
        End If
    Next i
    
    For entryNum = 1 To entryRows.count
        prodRow = entryRows(entryNum)
        eID = wsProd.Cells(prodRow, 1).value & ""
        
        ' Entry Type
        Dim entryType As String
        If UCase(Left(eID, 3)) = "WAR" Then
            entryType = "WARRANTY"
        ElseIf UCase(Left(eID, 3)) = "SER" Then
            entryType = "SERVICE"
        Else
            entryType = "ENTRY"
        End If
        
        ' Date from Job_Master
        Dim entryDate As String
        entryDate = ""
        For i = 2 To wsJob.Cells(wsJob.Rows.count, 1).End(xlUp).row
            If Trim(UCase(wsJob.Cells(i, 1).value & "")) = Trim(UCase(eID)) Then
                entryDate = wsJob.Cells(i, 8).value & ""
                Exit For
            End If
        Next i
        If entryDate = "" Then entryDate = Format(Date, "dd-MMM-yyyy")
        
        ' Entry Details (?????????)
        msg = msg & vbLf & "------------------------------" & vbLf & _
              "*" & entryType & " ID: " & eID & "*" & vbLf & _
              "Date: " & entryDate & vbLf & _
              "Product: " & wsProd.Cells(prodRow, 6).value & " | " & wsProd.Cells(prodRow, 7).value & vbLf & _
              "Model: " & wsProd.Cells(prodRow, 8).value & " | S/N: " & wsProd.Cells(prodRow, 9).value
        
        ' Problem (?? ?? ???? ???)
        Dim prob As String
        prob = wsProd.Cells(prodRow, 13).value & ""
        If prob <> "" Then
            msg = msg & vbLf & "Issue: " & prob
        End If
        
        ' Accessories
        Dim accRow As Long, lastRowAcc As Long, accCount As Integer
        lastRowAcc = wsAcc.Cells(wsAcc.Rows.count, 1).End(xlUp).row
        accCount = 0
        
        For accRow = 2 To lastRowAcc
            If Trim(UCase(wsAcc.Cells(accRow, 1).value & "")) = Trim(UCase(eID)) Then
                accCount = accCount + 1
                If accCount = 1 Then msg = msg & vbLf & "Accessories:"
                
                Dim accName As String, accReceived As String
                accName = wsAcc.Cells(accRow, 4).value & ""
                accReceived = wsAcc.Cells(accRow, 7).value & ""
                
                msg = msg & vbLf & accCount & ". " & accName
                If accReceived <> "" Then msg = msg & " (" & accReceived & ")"
                
                If accCount >= 4 Then Exit For
            End If
        Next accRow
                If accCount = 0 Then
            msg = msg & vbLf & "Accessories: None"
        End If
       
        
        
    Next entryNum
    
        ' ===== COMBINED PAYMENT DETAILS - ALL ENTRIES =====
    lastRowPay = wsPay.Cells(wsPay.Rows.count, 1).End(xlUp).row
    payCount = 0
    totalAmt = 0
    totalCharge = 0
    totalExpense = 0
    totalReceived = 0
    totalDiscount = 0
    
    ' Count all valid payments across ALL entries
    For entryNum = 1 To entryRows.count
        prodRow = entryRows(entryNum)
        eID = wsProd.Cells(prodRow, 1).value & ""
        For payRow = 2 To lastRowPay
            If Trim(UCase(wsPay.Cells(payRow, 2).value & "")) = Trim(UCase(eID)) Then
                If UCase(Trim(wsPay.Cells(payRow, 4).value & "")) <> "DELETED" And Trim(wsPay.Cells(payRow, 6).value & "") = "" Then
                    payCount = payCount + 1
                End If
            End If
        Next
    Next
    
    If payCount > 0 Then
        msg = msg & vbLf & "------------------------------"
        msg = msg & vbLf & "*Payment Details:*"
        
        ' Reset totals for actual calculation
        totalAmt = 0
        totalCharge = 0
        totalExpense = 0
        totalReceived = 0
        totalDiscount = 0
        
        For entryNum = 1 To entryRows.count
            prodRow = entryRows(entryNum)
            eID = wsProd.Cells(prodRow, 1).value & ""
            
            For payRow = 2 To lastRowPay
                If Trim(UCase(wsPay.Cells(payRow, 2).value & "")) = Trim(UCase(eID)) Then
                    If UCase(Trim(wsPay.Cells(payRow, 4).value & "")) <> "DELETED" And Trim(wsPay.Cells(payRow, 6).value & "") = "" Then
                        
                        If IsNumeric(wsPay.Cells(payRow, 16).value) Then
                            pAmt = CDbl(wsPay.Cells(payRow, 16).value)
                        Else
                            pAmt = 0
                        End If
                        
                        cat = UCase(Trim(wsPay.Cells(payRow, 9).value & ""))
                        If cat <> "DISCUSSION" Then totalAmt = totalAmt + pAmt
                        If cat = "CHARGE" Or cat = "PART_USED" Or cat = "PART_REPLACED" Then totalCharge = totalCharge + pAmt
                        If cat = "EXPENSE" Then totalExpense = totalExpense + pAmt
                        If cat = "RECEIVED" Or cat = "ADVANCE" Then totalReceived = totalReceived + pAmt
                        If cat = "DISCOUNT" Then totalDiscount = totalDiscount + pAmt
                        
                        msg = msg & vbLf & "› " & wsPay.Cells(payRow, 1).value & " | " & _
                              wsPay.Cells(payRow, 8).value & " | " & _
                              wsPay.Cells(payRow, 15).value & " | Rs." & Format(pAmt, "0.00")
                    End If
                End If
            Next
        Next
        
        msg = msg & vbLf & "------------------------------"
        msg = msg & vbLf & "*Payment Summary:*"
        msg = msg & vbLf & "Total Charge: Rs." & Format(totalCharge, "0.00")
        msg = msg & vbLf & "Total Expense: Rs." & Format(totalExpense, "0.00")
        msg = msg & vbLf & "Total Received: Rs." & Format(totalReceived, "0.00")
        If totalDiscount > 0 Then msg = msg & vbLf & "Total Discount: Rs." & Format(totalDiscount, "0.00")
        msg = msg & vbLf & "*TOTAL PENDING: Rs." & Format(totalCharge + totalExpense - totalReceived - totalDiscount, "0.00") & "*"
    End If
    
    ' ===== FOOTER =====
    msg = msg & vbLf & "------------------------------" & vbLf & _
          "*Thank You for choosing " & compName & "!*"
          
    
    
    BuildWhatsAppMessage = msg
End Function
'===========================================
' LOAD PRODUCTS TO GRID
'===========================================
Private Sub LoadProductsToGrid(entryID As String)
    Dim wsJob As Worksheet
    Dim wsAcc As Worksheet
    Dim lastRow As Long, lastRowAcc As Long
    Dim i As Long, a As Long
    Dim listItem As Object
    
    On Error Resume Next
    Set wsJob = ThisWorkbook.Sheets("Job_Product")
    Set wsAcc = ThisWorkbook.Sheets("Job_Accessory")
    On Error GoTo 0
    
    If wsJob Is Nothing Then Exit Sub
    
    ' Clear product grid
    On Error Resume Next
    Me.Controls("lstProducts").ListItems.Clear
    On Error GoTo 0
    
    lastRow = wsJob.Cells(wsJob.Rows.count, 1).End(xlUp).row
    
    For i = 2 To lastRow
        If UCase(Trim(wsJob.Cells(i, 1).value)) = UCase(entryID) Then
            
            ' Main Product
            On Error Resume Next
            Set listItem = Me.Controls("lstProducts").ListItems.Add(, , Trim(wsJob.Cells(i, 3).value)) ' EntryType
            listItem.SubItems(1) = Trim(wsJob.Cells(i, 1).value)   ' EntryID
            listItem.SubItems(2) = Trim(wsJob.Cells(i, 6).value)   ' ProductType
            listItem.SubItems(3) = Trim(wsJob.Cells(i, 7).value)   ' Company
            listItem.SubItems(4) = Trim(wsJob.Cells(i, 8).value)   ' Model
            listItem.SubItems(5) = Trim(wsJob.Cells(i, 9).value)   ' Serial
            listItem.SubItems(6) = Trim(wsJob.Cells(i, 12).value)  ' WarrantyStatus
            listItem.SubItems(7) = Trim(wsJob.Cells(i, 4).value)   ' Status
            On Error GoTo 0
            
            ' Accessories
            If Not wsAcc Is Nothing Then
                lastRowAcc = wsAcc.Cells(wsAcc.Rows.count, 1).End(xlUp).row
                For a = 2 To lastRowAcc
                    If UCase(Trim(wsAcc.Cells(a, 1).value)) = UCase(entryID) Then
                        On Error Resume Next
                        Set listItem = Me.Controls("lstProducts").ListItems.Add(, , "")
                        listItem.SubItems(1) = ""
                        listItem.SubItems(2) = "> " & Trim(wsAcc.Cells(a, 4).value)
                        listItem.SubItems(3) = Trim(wsAcc.Cells(a, 5).value)
                        listItem.SubItems(4) = Trim(wsAcc.Cells(a, 7).value)
                        listItem.SubItems(5) = Trim(wsAcc.Cells(a, 6).value)
                        listItem.SubItems(6) = ""
                        listItem.SubItems(7) = ""
                        listItem.ForeColor = RGB(0, 102, 153)
                        On Error GoTo 0
                    End If
                Next a
            End If
            
            Exit For
        End If
    Next i
End Sub

'===========================================
' LOAD PAYMENTS TO GRID
'===========================================
Private Sub LoadPaymentsToGrid(entryID As String)
    Dim wsPay As Worksheet
    Dim lastRow As Long, i As Long
    Dim listItem As Object
    
    On Error Resume Next
    Set wsPay = ThisWorkbook.Sheets("Payment_Master")
    On Error GoTo 0
    
    If wsPay Is Nothing Then Exit Sub
    
    ' Clear payment grid
    On Error Resume Next
    Me.Controls("lstPayments").ListItems.Clear
    On Error GoTo 0
    
    lastRow = wsPay.Cells(wsPay.Rows.count, 1).End(xlUp).row
    
    For i = 2 To lastRow
        If UCase(Trim(wsPay.Cells(i, 2).value)) = UCase(entryID) Then
            On Error Resume Next
            Set listItem = Me.Controls("lstPayments").ListItems.Add(, , Trim(wsPay.Cells(i, 1).value)) ' PaymentID
            listItem.SubItems(1) = Trim(wsPay.Cells(i, 2).value)   ' EntryID
            listItem.SubItems(2) = Trim(wsPay.Cells(i, 3).value)   ' Type
            listItem.SubItems(3) = Trim(wsPay.Cells(i, 4).value)   ' Category
            listItem.SubItems(4) = Trim(wsPay.Cells(i, 5).value)   ' PayMode
            listItem.SubItems(5) = Trim(wsPay.Cells(i, 6).value)   ' Name
            listItem.SubItems(6) = Trim(wsPay.Cells(i, 7).value)   ' Company
            listItem.SubItems(7) = Trim(wsPay.Cells(i, 8).value)   ' Model
            listItem.SubItems(8) = Trim(wsPay.Cells(i, 9).value)   ' Serial
            listItem.SubItems(9) = Trim(wsPay.Cells(i, 10).value)  ' Qty
            On Error GoTo 0
        End If
    Next i
End Sub

'===========================================
' LOAD VERIFY INFO
'===========================================
Private Sub LoadVerifyInfo(entryID As String)
    Dim wsJob As Worksheet
    Dim lastRow As Long, i As Long
    
    On Error Resume Next
    Set wsJob = ThisWorkbook.Sheets("Job_Product")
    On Error GoTo 0
    
    If wsJob Is Nothing Then Exit Sub
    
    lastRow = wsJob.Cells(wsJob.Rows.count, 1).End(xlUp).row
    
    For i = 2 To lastRow
        If UCase(Trim(wsJob.Cells(i, 1).value)) = UCase(entryID) Then
            
            ' Verify Type (combo)
            On Error Resume Next
            Me.Controls("cmbVerifyType").value = Trim(wsJob.Cells(i, 14).value)
            On Error GoTo 0
            
            ' Verify Name (combo)
            On Error Resume Next
            Me.Controls("cmbVerifyName").value = Trim(wsJob.Cells(i, 15).value)
            On Error GoTo 0
            
            Exit For
        End If
    Next i
End Sub
Private Sub UserForm_Activate()
    Dim entryID As String
    entryID = Trim(Me.Tag)
    
    If entryID <> "" Then
        Call LoadEntryData(entryID)
        Me.Tag = ""  ' Clear so it doesn't reload
     End If
     
     On Error Resume Next
    lblTime.caption = Format(Time, "hh:mm AM/PM")
    On Error GoTo 0
End Sub
'===========================================
' LOAD EXISTING ENTRY DATA - EDIT MODE
'===========================================
Private Sub LoadEntryData(entryID As String)
    Dim wsJob As Worksheet, wsCust As Worksheet, wsJobMaster As Worksheet
    Dim lastRowJob As Long, lastRowCust As Long, lastRowJobM As Long
    Dim i As Long, k As Long, jm As Long
    Dim custID As String
    Dim found As Boolean
    Dim photoFile As String
    
    On Error Resume Next
    Set wsJob = ThisWorkbook.Sheets("Job_Product")
    Set wsCust = ThisWorkbook.Sheets("Customer_Master")
    Set wsJobMaster = ThisWorkbook.Sheets("Job_Master")
    On Error GoTo 0
    
    If wsJob Is Nothing Then
        MsgBox "Job_Product sheet not found!", vbCritical
        Exit Sub
    End If
    
    found = False
    lastRowJob = wsJob.Cells(wsJob.Rows.count, 1).End(xlUp).row
    
    ' ===== STEP 1: FIND ENTRY & GET CUSTOMER ID =====
    For i = 2 To lastRowJob
        If UCase(Trim(wsJob.Cells(i, 1).value)) = UCase(entryID) Then
            found = True
            custID = Trim(wsJob.Cells(i, 2).value)
            Exit For
        End If
    Next i
    
    If Not found Then
        MsgBox "Entry not found: " & entryID, vbExclamation
        Exit Sub
    End If
    
    ' ===== STEP 2: LOAD CUSTOMER DETAILS =====
    If wsCust Is Nothing Then
        MsgBox "Customer_Master sheet not found!", vbCritical
        Exit Sub
    End If
    
    lastRowCust = wsCust.Cells(wsCust.Rows.count, 1).End(xlUp).row
    
    For k = 2 To lastRowCust
        If UCase(Trim(wsCust.Cells(k, 1).value)) = UCase(custID) Then
            
            '??? FIX: Col 2 = Name, Col 3 = Mobile ???
            txtCustomerID.caption = custID
            txtMobile.value = Trim(wsCust.Cells(k, 3).value)            'C: Mobile
            txtCustomerName.caption = Trim(wsCust.Cells(k, 2).value)  'B: Name
            txtAddress.caption = Trim(wsCust.Cells(k, 4).value)       'D: Address
            txtEmail.caption = Trim(wsCust.Cells(k, 5).value)         'E: Email
            txtGST.caption = Trim(wsCust.Cells(k, 6).value)          'F: GST
            
            photoFile = Trim(wsCust.Cells(k, 7).value)
            If photoFile <> "" Then
                Call LoadCustomerPhoto(photoFile)
            End If
            
            Exit For
        End If
    Next k
    
    ' ===== STEP 3: LOAD VERIFY TYPE/NAME FROM Job_Master =====
If Not wsJobMaster Is Nothing Then
    lastRowJobM = wsJobMaster.Cells(wsJobMaster.Rows.count, 1).End(xlUp).row
    For jm = 2 To lastRowJobM
        If UCase(Trim(wsJobMaster.Cells(jm, 1).value)) = UCase(entryID) Then
            
            Dim vType As String, vName As String
            vType = Trim(wsJobMaster.Cells(jm, 6).value & "")
            vName = Trim(wsJobMaster.Cells(jm, 7).value & "")
            
            ' === VERIFY TYPE SAFELY ===
            If vType <> "" Then
                On Error Resume Next
                cmbVerifyType.value = vType
                If Err.Number <> 0 Then
                    cmbVerifyType.AddItem vType
                    cmbVerifyType.value = vType
                    Err.Clear
                End If
                On Error GoTo 0
                
                ' Load names for this type
                On Error Resume Next      ' <-- ADD
                Call LoadVerifyName
                On Error GoTo 0           ' <-- ADD
                
                ' === VERIFY NAME SAFELY ===
                If vName <> "" Then
                    On Error Resume Next
                    cmbVerifyName.value = vName
                    If Err.Number <> 0 Then
                        cmbVerifyName.AddItem vName
                        cmbVerifyName.value = vName
                        Err.Clear
                    End If
                    On Error GoTo 0
                End If
            End If
            
            Exit For
        End If
    Next jm
End If
    
    ' ===== STEP 4: LOAD PRODUCTS & PAYMENTS =====
    Call LoadProductList(custID)
    Call LoadPaymentList(custID)
    
    ' ===== STEP 5: FORM CAPTION =====
    Me.caption = "JOB ENTRY - EDIT MODE (" & entryID & ")"
    
End Sub
'===========================================
' DEBUG - LIST ALL CONTROL NAMES
'===========================================
Private Sub ListAllControls()
    Dim ctrl As Control
    Dim msg As String
    msg = "CONTROLS IN frmEntryWizard:" & vbCrLf & vbCrLf
    
    For Each ctrl In Me.Controls
        msg = msg & TypeName(ctrl) & " = " & ctrl.name & vbCrLf
    Next ctrl
    
    MsgBox msg, vbInformation, "Control Names"
End Sub
'==============================
' VERIFY MASTER SHEET - AUTO CREATE
'==============================
Private Function GetVerifyMasterSheet() As Worksheet
    On Error Resume Next
    Set GetVerifyMasterSheet = ThisWorkbook.Sheets("Verify_Master")
    On Error GoTo 0
    
    If GetVerifyMasterSheet Is Nothing Then
        Set GetVerifyMasterSheet = ThisWorkbook.Sheets.Add( _
            After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
        GetVerifyMasterSheet.name = "Verify_Master"
        GetVerifyMasterSheet.Range("A1").value = "ID"
        GetVerifyMasterSheet.Range("B1").value = "Type"
        GetVerifyMasterSheet.Range("C1").value = "Name"
        GetVerifyMasterSheet.visible = xlSheetHidden
    End If
End Function

'==============================
' SAMPLE DATA INSERT (One Time)
'==============================
Public Sub InsertSampleVerifyData()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Verify_Master")
    On Error GoTo 0
    
    ' Agar sheet nahi hai to bana do
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
        ws.name = "Verify_Master"
        ws.Range("A1").value = "ID"
        ws.Range("B1").value = "Type"
        ws.Range("C1").value = "Name"
        ws.visible = xlSheetHidden
    End If
    
    ' Purana data hatao (header bachake)
    If ws.Cells(ws.Rows.count, 1).End(xlUp).row > 1 Then
        ws.Range("A2:C" & ws.Cells(ws.Rows.count, 1).End(xlUp).row).ClearContents
    End If
    
    ' ===== SAMPLE DATA =====
    ws.Cells(2, 1).value = 1:  ws.Cells(2, 2).value = "MANAGER":   ws.Cells(2, 3).value = "RAM"
    ws.Cells(3, 1).value = 2:  ws.Cells(3, 2).value = "MANAGER":   ws.Cells(3, 3).value = "SHYAM"
    ws.Cells(4, 1).value = 3:  ws.Cells(4, 2).value = "ENGINEER":  ws.Cells(4, 3).value = "MOHAN"
    ws.Cells(5, 1).value = 4:  ws.Cells(5, 2).value = "ENGINEER":  ws.Cells(5, 3).value = "SOHAN"
    ws.Cells(6, 1).value = 5:  ws.Cells(6, 2).value = "ADMIN":      ws.Cells(6, 3).value = "RAJU"
    
    MsgBox "Sample Data Inserted!" & vbCrLf & vbCrLf & _
           "MANAGER ? RAM, SHYAM" & vbCrLf & _
           "ENGINEER ? MOHAN, SOHAN" & vbCrLf & _
           "ADMIN ? RAJU" & vbCrLf & vbCrLf & _
           "Ab form band karke dubara kholo!", vbInformation
End Sub

