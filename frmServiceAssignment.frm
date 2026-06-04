VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmServiceAssignment 
   Caption         =   "Service Product Assignment"
   ClientHeight    =   13875
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   19935
   OleObjectBlob   =   "frmServiceAssignment.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmServiceAssignment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private IsDataSaved As Boolean
Private LastAssignID As String
Private isEditMode As Boolean      ' <-- NEW
Private EditAssignRow As Long      ' <-- NEW


Private Sub UserForm_Initialize()
    Me.Width = 1012
    Me.Height = 750
    Me.StartUpPosition = 1
    
    AddMinMaxButtons Me
    
    Dim entryID As String
    entryID = Trim(Me.Tag)
    
    ' Default = NEW mode
    isEditMode = False
    EditAssignRow = 0
    
    ' Load Entry IDs
    LoadEntryIDs
    
    IsDataSaved = False
    LastAssignID = ""
    
    ' Load Assign Mode
    cmbAssignTo.Clear
    cmbAssignTo.AddItem "VENDOR"
    cmbAssignTo.AddItem "ENGINEER"
    cmbAssignTo.AddItem "INHOUSE"
    cmbAssignTo.AddItem "SERVICE_STATION"
    
    ' Setup Courier Mode
    SetupCourierMode
    
    ' Setup Pending ListView
    With lstPending
        .View = lvwReport
        .FullRowSelect = True
        .Gridlines = True
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "Entry ID", 90
        .ColumnHeaders.Add , , "Customer", 130
        .ColumnHeaders.Add , , "Mobile", 100
        .ColumnHeaders.Add , , "Product", 120
        .ColumnHeaders.Add , , "Company", 100
        .ColumnHeaders.Add , , "Model", 100
        .ColumnHeaders.Add , , "Serial", 120
        .ColumnHeaders.Add , , "Warranty", 90
    End With
    
    lblPendingCount.caption = "0"
    lblPendingCount.visible = True
    
    ' Payment ListView
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
    
    ' Set default dates
    txtSendDate.value = Format(Date, "dd-mm-yyyy")
    txtExpectedReturn.value = Format(Date + 7, "dd-mm-yyyy")
    
        ' ===== CHECK IF EDIT MODE =====
    If entryID <> "" Then
        ' ===== EDIT MODE =====
        isEditMode = True
        Me.caption = "SERVICE PRODUCT ASSIGNMENT - EDIT (" & entryID & ")"
        
        ' Add entryID to combo (it's ASSIGNED, not in ACTIVE list)
        Dim found As Boolean, i As Long
        found = False
        For i = 0 To cmbEntryID.ListCount - 1
            If UCase(Trim(cmbEntryID.List(i))) = UCase(entryID) Then
                found = True
                Exit For
            End If
        Next i
        If Not found Then cmbEntryID.AddItem entryID
        
        ' Set EntryID - triggers cmbEntryID_Change
        cmbEntryID.value = entryID
        
        ' Load saved assignment data
        Call LoadServiceAssignData(entryID)
        
        ' Load payment list
        Call LoadPaymentList(entryID)
        
        ' ===== EDIT MODE BUTTON STATES =====
        btnAssign.caption = "UPDATE"          ' "ASSIGN" ?? ??? "UPDATE" ????
        btnAssign.enabled = True              ' ?????? ????? ???? ?? ???
        
        ' BackColor HATAYA - Error 91 fix
        
        btnSendWhatsApp.enabled = False       ' EDIT mode ??? ???
        btnSendEmail.enabled = False          ' EDIT mode ??? ???
        btnPrintAssign.enabled = False        ' EDIT mode ??? ???
        btnExportPDF.enabled = False          ' EDIT mode ??? ???
        
        btnClear.enabled = True               ' ????? ????
        
        Me.Tag = ""
    Else
        ' ===== NEW MODE =====
        isEditMode = False
        EditAssignRow = 0
        btnAssign.caption = "ASSIGN"          ' ?? mode ??? "ASSIGN"
        ' BackColor HATAYA - Error 91 fix
        
        ' NEW MODE: ????? Assign ????, ???? ???
        btnAssign.enabled = True
        btnSendWhatsApp.enabled = False
        btnSendEmail.enabled = False
        btnPrintAssign.enabled = False
        btnClear.enabled = False
        btnExportPDF.enabled = False
        
        ClearForm
    End If
    
    Call UpdatePaymentButtons
End Sub

'===========================================
' LOAD ENTRY IDs
'===========================================
Private Sub LoadEntryIDs()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    Set ws = ThisWorkbook.Sheets("Job_Product")
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    
    cmbEntryID.Clear
    
    For i = 2 To lastRow
        If UCase(Trim(ws.Cells(i, 3).value)) = "SERVICE" And _
   UCase(Trim(ws.Cells(i, 4).value)) = "ACTIVE" Then
    cmbEntryID.AddItem ws.Cells(i, 1).value
End If
    Next i
End Sub

Private Sub cmbEntryID_Change()
    Dim wsProduct As Worksheet
    Dim wsCustomer As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim entryID As String
    Dim custID As String
    
    entryID = Trim(cmbEntryID.value)
    If entryID = "" Then Exit Sub
    
    On Error Resume Next
    Set wsProduct = ThisWorkbook.Sheets("Job_Product")
    Set wsCustomer = ThisWorkbook.Sheets("Customer_Master")
    On Error GoTo 0
    
    If wsProduct Is Nothing Then
        MsgBox "Job_Product sheet not found!", vbCritical
        Exit Sub
    End If
    
    lastRow = wsProduct.Cells(wsProduct.Rows.count, 1).End(xlUp).row
    
    For i = 2 To lastRow
        If UCase(Trim(wsProduct.Cells(i, 1).value)) = UCase(entryID) Then
            
            custID = Trim(wsProduct.Cells(i, 2).value)
            
            ' Load all details
            lblEntryIDValue.caption = entryID
            lblProductValue.caption = wsProduct.Cells(i, 6).value
            lblCompanyValue.caption = wsProduct.Cells(i, 7).value
            lblModelValue.caption = wsProduct.Cells(i, 8).value
            lblSerialValue.caption = wsProduct.Cells(i, 9).value
            txtProblemDescription.value = wsProduct.Cells(i, 13).value
            lblWarrantyStatusValue.caption = wsProduct.Cells(i, 12).value
            
            ' Load Customer Details
            If custID <> "" Then
                LoadCustomerDetails custID
            End If
            
            ' Load Product Photo
            On Error Resume Next
            LoadProductPhoto entryID
            On Error GoTo 0
            
            ' Load Accessories
            LoadAccessories entryID
            
            Exit For
        End If
    Next i
    
        ' ===== LOAD PAYMENT LIST =====
    Call LoadPaymentList(cmbEntryID.value)
End Sub
Private Sub LoadCustomerDetails(custID As String)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    If custID = "" Then Exit Sub
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Customer_Master")
    On Error GoTo 0
    
    If ws Is Nothing Then Exit Sub
    
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    
    For i = 2 To lastRow
        If UCase(Trim(ws.Cells(i, 1).value)) = UCase(custID) Then
            ' ? FIX: Column B = Name, Column C = Mobile
            lblCustomerValue.caption = ws.Cells(i, 2).value   ' Name
            lblMobileValue.caption = ws.Cells(i, 3).value     ' Mobile
            
            ' Load Customer Photo
            On Error Resume Next
            If Trim(ws.Cells(i, 7).value) <> "" Then
                LoadCustomerPhoto Trim(ws.Cells(i, 7).value)
            End If
            On Error GoTo 0
            
            Exit For
        End If
    Next i
End Sub
'===========================================
' LOAD CUSTOMER PHOTO
'===========================================
Private Sub LoadCustomerPhoto(photoName As String)
    Dim folderPath As String
    Dim fullPath As String
    
    On Error Resume Next
    
    folderPath = ThisWorkbook.Sheets("Settings").Range("B3").value
    If folderPath = "" Then
        folderPath = ThisWorkbook.path & "\GLOBAL_SOFT_DATA\Images\Customers\"
    End If
    If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
    
    fullPath = folderPath & photoName
    
    If Dir(fullPath) <> "" Then
        imgCustomer.Picture = LoadPicture(fullPath)
    Else
        Set imgCustomer.Picture = Nothing
    End If
    On Error GoTo 0
End Sub

'===========================================
' LOAD PRODUCT PHOTO - CORRECTED NAME
'===========================================
Private Sub LoadProductPhoto(entryID As String)
    Dim photoPath As String
    
    On Error Resume Next
    
    ' Use imgProductPhoto (as per your form)
    Set imgProductPhoto.Picture = Nothing
    
    If Trim(entryID) = "" Then Exit Sub
    
    ' Build photo path
    photoPath = ThisWorkbook.path & "\GLOBAL_SOFT_DATA\Images\Products\" & entryID & "\Photo1.jpg"
    
    ' Check if file exists
    If Dir(photoPath) = "" Then Exit Sub
    
    ' Load photo
    imgProductPhoto.Picture = LoadPicture(photoPath)
    
    On Error GoTo 0
End Sub

'===========================================
' LOAD ACCESSORIES CHECKBOXES
'===========================================
Private Sub LoadAccessories(entryID As String)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim accCount As Integer
    
    Set ws = ThisWorkbook.Sheets("Job_Accessory")
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    
    ' Reset checkboxes
    For i = 1 To 6
        Me.Controls("CheckBox" & i).value = False
        Me.Controls("CheckBox" & i).caption = ""
        Me.Controls("CheckBox" & i).visible = False
    Next i
    
    accCount = 0
    
    For i = 2 To lastRow
        If UCase(Trim(ws.Cells(i, 1).value)) = UCase(entryID) Then
            accCount = accCount + 1
            If accCount <= 6 Then
                Me.Controls("CheckBox" & accCount).caption = ws.Cells(i, 4).value
                Me.Controls("CheckBox" & accCount).visible = True
            End If
        End If
    Next i
End Sub
'===========================================
' ASSIGN TO CHANGE
'===========================================
Private Sub cmbAssignTo_Change()
    Dim mode As String
    
    mode = Trim(cmbAssignTo.value)
    If mode = "" Then Exit Sub
    
    ' Clear labels
    lblVendorEmail.caption = ""
    lblVendorMobile.caption = ""
    lblVendorAddress.caption = ""
    
    ' Refresh dropdown
    Call RefreshAssignNameDropdown(mode)
End Sub
Private Sub cmbAssignName_Change()
    Dim mode As String, name As String
    Dim ws As Worksheet, lastRow As Long, i As Long
    
    mode = UCase(Trim(cmbAssignTo.value))
    name = Trim(cmbAssignName.value)
    
    ' Clear
    
    lblVendorMobile.caption = ""
    lblVendorAddress.caption = ""
    lblVendorEmail.caption = ""
    
    If name = "" Then
        lstPending.ListItems.Clear
        lblPendingCount.caption = "0"
        Exit Sub
    End If
    
    ' Load vendor details from register
    Select Case mode
        Case "VENDOR": Set ws = ThisWorkbook.Sheets("Vendor_Register")
        Case "ENGINEER": Set ws = ThisWorkbook.Sheets("Engineer_Register")
        Case "INHOUSE": Set ws = ThisWorkbook.Sheets("Inhouse_Register")
        Case "SERVICE_STATION": Set ws = ThisWorkbook.Sheets("ServiceStation_Register")
        Case Else: Exit Sub
    End Select
    
    If ws Is Nothing Then Exit Sub
    
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    For i = 2 To lastRow
        If UCase(Trim(ws.Cells(i, 2).value)) = UCase(name) Then
            
            lblVendorMobile.caption = ws.Cells(i, 3).value
            lblVendorEmail.caption = ws.Cells(i, 7).value
            lblVendorAddress.caption = ws.Cells(i, 4).value
            Exit For
        End If
    Next i
    
    ' Load Pending List using our new subroutine
    Call RefreshPendingList
    
End Sub
Private Sub RefreshPendingList()
    Dim wsAssign As Worksheet, wsProduct As Worksheet, wsCustomer As Worksheet
    Dim lastRow As Long, i As Long, prodRow As Long, custRow As Long
    Dim pendingCount As Long
    Dim assignToName As String, statusVal As String, entryID As String, custID As String
    Dim custName As String, custMobile As String
    Dim vendName As String, vendNameShort As String
    Dim listItem As Object
    Dim assignType As String
    
    ' ===== GET SELECTED VENDOR NAME =====
    vendName = Trim(cmbAssignName.value)
    If vendName = "" Then
        lblPendingCount.caption = "0"
        lstPending.ListItems.Clear
        Exit Sub
    End If
    
    ' Get short name
    On Error Resume Next
    vendNameShort = Trim(Split(vendName, "(")(0))
    If Err.Number <> 0 Then vendNameShort = vendName
    On Error GoTo 0
    
    ' Clear list
    lstPending.ListItems.Clear
    pendingCount = 0
    
    ' Get sheets
    On Error Resume Next
    Set wsAssign = ThisWorkbook.Sheets("Assign_Master")
    Set wsProduct = ThisWorkbook.Sheets("Job_Product")
    Set wsCustomer = ThisWorkbook.Sheets("Customer_Master")
    On Error GoTo 0
    
    If wsAssign Is Nothing Then Exit Sub
    
    lastRow = wsAssign.Cells(wsAssign.Rows.count, 1).End(xlUp).row
    
    ' ===== LOOP THROUGH DATA =====
    For i = 2 To lastRow
        assignToName = Trim(wsAssign.Cells(i, "N").value)
        statusVal = UCase(Trim(wsAssign.Cells(i, "AF").value))
        assignType = UCase(Trim(wsAssign.Cells(i, "K").value))
        
        If statusVal = "PENDING" And assignType = "SERVICE" And Len(assignToName) > 0 Then
            ' Match check
            If InStr(1, UCase(assignToName), UCase(vendNameShort)) > 0 Then
                
                pendingCount = pendingCount + 1
                entryID = Trim(wsAssign.Cells(i, "B").value)
                
                ' Find product details
                For prodRow = 2 To wsProduct.Cells(wsProduct.Rows.count, 1).End(xlUp).row
                    If UCase(Trim(wsProduct.Cells(prodRow, 1).value)) = UCase(entryID) Then
                        
                        custID = Trim(wsProduct.Cells(prodRow, 2).value)
                        custName = "": custMobile = ""
                        
                        For custRow = 2 To wsCustomer.Cells(wsCustomer.Rows.count, 1).End(xlUp).row
                            If UCase(Trim(wsCustomer.Cells(custRow, 1).value)) = UCase(custID) Then
                                custName = wsCustomer.Cells(custRow, 3).value
                                custMobile = wsCustomer.Cells(custRow, 2).value
                                Exit For
                            End If
                        Next custRow
                        
                        Set listItem = lstPending.ListItems.Add(, , entryID)
                        listItem.SubItems(1) = custName
                        listItem.SubItems(2) = custMobile
                        listItem.SubItems(3) = wsProduct.Cells(prodRow, 6).value
                        listItem.SubItems(4) = wsProduct.Cells(prodRow, 7).value
                        listItem.SubItems(5) = wsProduct.Cells(prodRow, 8).value
                        listItem.SubItems(6) = wsProduct.Cells(prodRow, 9).value
                        listItem.SubItems(7) = wsProduct.Cells(prodRow, 12).value
                        
                        Exit For
                    End If
                Next prodRow
                
            End If
        End If
    Next i
    
        ' ===== FORCE UPDATE LABEL =====
    With lblPendingCount
        .caption = CStr(pendingCount)
        .visible = True
        .AutoSize = False
        .BackStyle = fmBackStyleTransparent
        .BorderStyle = fmBorderStyleNone
        
        ' Bright colors for visibility
        If pendingCount = 0 Then
            .ForeColor = RGB(0, 128, 0)     ' Green
            .BackColor = RGB(255, 255, 255) ' White
        Else
            .ForeColor = RGB(255, 0, 0)     ' Red
            .BackColor = RGB(255, 255, 200) ' Light Yellow background
        End If
        
        ' Ensure it's on top
        .ZOrder 0
    End With
    
    ' Refresh form
    Me.Repaint
    
End Sub
' LOAD VENDOR NAMES
'===========================================
Private Sub LoadVendorNames()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Vendor_Register")
    On Error GoTo 0
    
    If ws Is Nothing Then Exit Sub
    
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    
    For i = 2 To lastRow
        If UCase(Trim(ws.Cells(i, 6).value)) = "ACTIVE" Then
            cmbCourierMode.AddItem ws.Cells(i, 2).value
        End If
    Next i
End Sub

'===========================================
' LOAD ENGINEER NAMES
'===========================================
Private Sub LoadEngineerNames()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Engineer_Register")
    On Error GoTo 0
    
    If ws Is Nothing Then Exit Sub
    
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    
    For i = 2 To lastRow
        If UCase(Trim(ws.Cells(i, 9).value)) = "ACTIVE" Then
            cmbCourierMode.AddItem ws.Cells(i, 2).value
        End If
    Next i
End Sub

'===========================================
' LOAD INHOUSE NAMES
'===========================================
Private Sub LoadInhouseNames()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Inhouse_Register")
    On Error GoTo 0
    
    ' If sheet doesn't exist, show default
    If ws Is Nothing Then
        cmbCourierMode.AddItem "INHOUSE TEAM"
        Exit Sub
    End If
    
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    
    For i = 2 To lastRow
        If UCase(Trim(ws.Cells(i, 6).value)) = "ACTIVE" Then
            cmbCourierMode.AddItem ws.Cells(i, 2).value
        End If
    Next i
End Sub

'===========================================
' LOAD SERVICE STATION NAMES
'===========================================
Private Sub LoadServiceStationNames()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("ServiceStation_Register")
    On Error GoTo 0
    
    If ws Is Nothing Then Exit Sub
    
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    
    For i = 2 To lastRow
        If UCase(Trim(ws.Cells(i, 8).value)) = "ACTIVE" Then
            cmbCourierMode.AddItem ws.Cells(i, 2).value
        End If
    Next i
End Sub



'===========================================
' LOAD VENDOR DETAILS
'===========================================
Private Sub LoadVendorDetails(vName As String)
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    
    Set ws = ThisWorkbook.Sheets("Vendor_Register")
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    
    For i = 2 To lastRow
        If UCase(Trim(ws.Cells(i, 2).value)) = UCase(vName) Then
            lblVendorName.caption = ws.Cells(i, 2).value
            lblVendorMobile.caption = ws.Cells(i, 3).value
            lblVendorAddress.caption = ws.Cells(i, 4).value
            lblVendorEmail.caption = ws.Cells(i, 7).value    ' G Column - Email (NEW)
            Exit For
        End If
    Next i
End Sub

'===========================================
' LOAD ENGINEER DETAILS
'===========================================
Private Sub LoadEngineerDetails(eName As String)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    Set ws = ThisWorkbook.Sheets("Engineer_Register")
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    
    For i = 2 To lastRow
        If UCase(Trim(ws.Cells(i, 2).value)) = UCase(eName) Then
            lblVendorName.caption = ws.Cells(i, 2).value
            lblVendorMobile.caption = ws.Cells(i, 3).value
            lblVendorAddress.caption = ws.Cells(i, 4).value
            Exit For
        End If
    Next i
End Sub

'===========================================
' LOAD INHOUSE DETAILS
'===========================================
Private Sub LoadInhouseDetails(iName As String)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Inhouse_Register")
    On Error GoTo 0
    
    If ws Is Nothing Then
        ' Default values
        lblVendorName.caption = iName
        lblVendorMobile.caption = ThisWorkbook.Sheets("Software_Config").Range("B4").value
        lblVendorAddress.caption = ThisWorkbook.Sheets("Software_Config").Range("B3").value
        Exit Sub
    End If
    
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    
    For i = 2 To lastRow
        If UCase(Trim(ws.Cells(i, 2).value)) = UCase(iName) Then
            lblVendorName.caption = ws.Cells(i, 2).value
            lblVendorMobile.caption = ws.Cells(i, 3).value
            lblVendorAddress.caption = ws.Cells(i, 4).value
            Exit For
        End If
    Next i
End Sub

'===========================================
' LOAD SERVICE STATION DETAILS
'===========================================
Private Sub LoadServiceStationDetails(sName As String)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    Set ws = ThisWorkbook.Sheets("ServiceStation_Register")
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    
    For i = 2 To lastRow
        If UCase(Trim(ws.Cells(i, 2).value)) = UCase(sName) Then
            lblVendorName.caption = ws.Cells(i, 2).value
            lblVendorMobile.caption = ws.Cells(i, 3).value
            lblVendorAddress.caption = ws.Cells(i, 4).value
            Exit For
        End If
    Next i
End Sub

'===========================================
' ADD BUTTON CLICK - COMPLETE ERROR-PROOF
'===========================================
Private Sub btnAdd_Click()
    On Error GoTo ErrorHandler
    
    Dim mode As String
    Dim newName As String
    Dim newMobile As String
    Dim newAddress As String
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim sheetName As String
    Dim nextID As String
    
    ' Get selected mode
    mode = UCase(Trim(cmbAssignTo.value))
    
    If mode = "" Then
        MsgBox "Please select Assign To first!", vbExclamation, "Required"
        cmbAssignTo.SetFocus
        Exit Sub
    End If
    
    ' Input Name
    newName = InputBox("Enter " & mode & " Name:", "Add New " & mode)
    If Trim(newName) = "" Then Exit Sub
    
    ' Input Mobile
    newMobile = InputBox("Enter Mobile Number:", "Add New " & mode)
    If Trim(newMobile) = "" Then Exit Sub
    
    ' Input Email ID (NEW)
    Dim newEmail As String
    newEmail = InputBox("Enter Email ID:", "Add New " & mode)
    ' Email optional ?? ???? ??, ??? ???? ?? ?? blank ?????
    
    ' Input Address
    newAddress = InputBox("Enter Address:", "Add New " & mode)
    If Trim(newAddress) = "" Then Exit Sub
    
    ' Determine which sheet to use
    Select Case mode
        Case "VENDOR"
            sheetName = "Vendor_Register"
        Case "ENGINEER"
            sheetName = "Engineer_Register"
        Case "INHOUSE"
            sheetName = "Inhouse_Register"
        Case "SERVICE_STATION"
            sheetName = "ServiceStation_Register"
        Case Else
            MsgBox "Invalid selection!", vbExclamation
            Exit Sub
    End Select
    
    ' Check if sheet exists
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo ErrorHandler
    
    ' If sheet doesn't exist, create it
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
        ws.name = sheetName
        
        ' Add headers
        ws.Cells(1, 1).value = "ID"
        ws.Cells(1, 2).value = "Name"
        ws.Cells(1, 3).value = "Mobile"
        ws.Cells(1, 4).value = "Address"
        ws.Cells(1, 5).value = "CreatedDate"
        ws.Cells(1, 6).value = "Status"
        
        ' Format headers
        With ws.Rows(1)
            .Font.Bold = True
            .Interior.Color = RGB(200, 200, 200)
        End With
    End If
    
    ' Check if name already exists
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    
    For i = 2 To lastRow
        If UCase(Trim(ws.Cells(i, 2).value)) = UCase(Trim(newName)) Then
            MsgBox mode & " '" & newName & "' already exists!", vbExclamation
            Exit Sub
        End If
    Next i
    
    ' Generate ID
    lastRow = lastRow + 1
    Select Case mode
        Case "VENDOR": nextID = "VEN" & Format(lastRow - 1, "00000")
        Case "ENGINEER": nextID = "ENG" & Format(lastRow - 1, "00000")
        Case "INHOUSE": nextID = "INH" & Format(lastRow - 1, "00000")
        Case "SERVICE_STATION": nextID = "STN" & Format(lastRow - 1, "00000")
    End Select
    
    ' Save data
    ws.Cells(lastRow, 1).value = nextID
    ws.Cells(lastRow, 2).value = UCase(newName)
    ws.Cells(lastRow, 3).value = newMobile
    ws.Cells(lastRow, 4).value = newAddress
    ws.Cells(lastRow, 5).value = Format(Now, "dd-mm-yyyy")
    ws.Cells(lastRow, 6).value = "ACTIVE"
    ws.Cells(lastRow, 7).value = LCase(newEmail)
    
    MsgBox mode & " Added Successfully!" & vbCrLf & _
           "ID: " & nextID & vbCrLf & _
           "Name: " & newName, vbInformation, "Success"
    
    ' Refresh dropdown - USE cmbAssignName (your combo box name)
    Call RefreshAssignNameDropdown(mode)
    
    ' Select the new name
    On Error Resume Next
    cmbAssignName.value = UCase(newName)
    On Error GoTo 0
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error: " & Err.Description & vbCrLf & _
           "Number: " & Err.Number, vbCritical, "Error"
End Sub

'===========================================
' REFRESH DROPDOWN BASED ON MODE
'===========================================
Private Sub RefreshAssignNameDropdown(mode As String)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    ' Clear dropdown
    On Error Resume Next
    cmbAssignName.Clear
    On Error GoTo 0
    
    ' Determine sheet
    Select Case UCase(mode)
        Case "VENDOR": Set ws = ThisWorkbook.Sheets("Vendor_Register")
        Case "ENGINEER": Set ws = ThisWorkbook.Sheets("Engineer_Register")
        Case "INHOUSE": Set ws = ThisWorkbook.Sheets("Inhouse_Register")
        Case "SERVICE_STATION": Set ws = ThisWorkbook.Sheets("ServiceStation_Register")
        Case Else: Exit Sub
    End Select
    
    If ws Is Nothing Then Exit Sub
    
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    
    ' Load names
    For i = 2 To lastRow
        If UCase(Trim(ws.Cells(i, 6).value)) = "ACTIVE" Or _
           UCase(Trim(ws.Cells(i, 6).value)) = "YES" Then
            On Error Resume Next
            cmbAssignName.AddItem ws.Cells(i, 2).value
            On Error GoTo 0
        End If
    Next i
End Sub


'===========================================
' UPDATE PRODUCT STATUS
'===========================================
Private Sub UpdateProductStatus(entryID As String, newStatus As String)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    Set ws = ThisWorkbook.Sheets("Job_Product")
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    
    For i = 2 To lastRow
        If UCase(Trim(ws.Cells(i, 1).value)) = UCase(entryID) Then
            ws.Cells(i, 4).value = newStatus
            ws.Cells(i, 5).value = Format(Now, "dd-mm-yyyy hh:mm")
            Exit For
        End If
    Next i
End Sub
 
Private Sub btnSendWhatsApp_Click()
    On Error GoTo ErrorHandler
    
    ' ===== ASSIGN SAVE CHECK =====
    If Not IsDataSaved Then
        MsgBox "Please save assignment first!", vbExclamation, "Error"
        Exit Sub
    End If
    
    ' ===== GET VENDOR MOBILE =====
    Dim mobile As String
    mobile = Trim(lblVendorMobile.caption & "")
    If mobile = "" Then
        MsgBox "Vendor mobile number not found!", vbExclamation
        Exit Sub
    End If
    
    ' Clean mobile
    mobile = Replace(mobile, " ", "")
    mobile = Replace(mobile, "-", "")
    mobile = Replace(mobile, "+", "")
    mobile = Replace(mobile, "(", "")
    mobile = Replace(mobile, ")", "")
    
    ' ===== COMPANY DETAILS =====
    Dim wsConfig As Worksheet
    Dim companyName As String, companyMobile As String, companyEmail As String
    Dim companyAddress As String
    Dim addr1 As String, addr2 As String, city As String, pin As String, state As String
    
    Set wsConfig = ThisWorkbook.Sheets("Software_Config")
    companyName = Trim(wsConfig.Range("B2").value & "")
    companyMobile = Trim(wsConfig.Range("B9").value & "")
    companyEmail = Trim(wsConfig.Range("B10").value & "")
    
    addr1 = Trim(wsConfig.Range("B3").value & "")
    addr2 = Trim(wsConfig.Range("B4").value & "")
    city = Trim(wsConfig.Range("B5").value & "")
    pin = Trim(wsConfig.Range("B6").value & "")
    state = Trim(wsConfig.Range("B7").value & "")
    
    companyAddress = ""
    If addr1 <> "" Then companyAddress = companyAddress & addr1
    If addr2 <> "" Then companyAddress = companyAddress & ", " & addr2
    If city <> "" Then companyAddress = companyAddress & ", " & city
    If pin <> "" Then companyAddress = companyAddress & " - " & pin
    If state <> "" Then companyAddress = companyAddress & ", " & state
    
    If companyName = "" Then companyName = "GLOBAL IT SOLUTIONS"
    
    ' ===== ACCESSORIES =====
    Dim accessoriesList As String
    accessoriesList = ""
    If CheckBox1.value = True Then accessoriesList = accessoriesList & "- " & CheckBox1.caption & vbCrLf
    If CheckBox2.value = True Then accessoriesList = accessoriesList & "- " & CheckBox2.caption & vbCrLf
    If CheckBox3.value = True Then accessoriesList = accessoriesList & "- " & CheckBox3.caption & vbCrLf
    If CheckBox4.value = True Then accessoriesList = accessoriesList & "- " & CheckBox4.caption & vbCrLf
    If CheckBox5.value = True Then accessoriesList = accessoriesList & "- " & CheckBox5.caption & vbCrLf
    If CheckBox6.value = True Then accessoriesList = accessoriesList & "- " & CheckBox6.caption & vbCrLf
    
    ' ===== BUILD MESSAGE =====
    Dim msg As String
    msg = "*" & UCase(companyName) & "*" & vbCrLf & _
          companyAddress & vbCrLf
    
    If companyMobile <> "" Then msg = msg & "Mobile: " & companyMobile & vbCrLf
    If companyEmail <> "" Then msg = msg & "Email: " & companyEmail & vbCrLf
    
    msg = msg & "------------------------" & vbCrLf & vbCrLf & _
          "*SERVICE ASSIGNMENT*" & vbCrLf & vbCrLf & _
          "*Entry ID:* " & cmbEntryID.value & vbCrLf & _
          "*Product:* " & lblProductValue.caption & vbCrLf & _
          "*Company:* " & lblCompanyValue.caption & vbCrLf & _
          "*Model:* " & lblModelValue.caption & vbCrLf & _
          "*Serial:* " & lblSerialValue.caption & vbCrLf & vbCrLf
    
    ' Problem
    If Trim(txtProblemDescription.value) <> "" Then
        msg = msg & "*Problem:* " & txtProblemDescription.value & vbCrLf & vbCrLf
    End If
    
    ' Accessories
    If accessoriesList <> "" Then
        msg = msg & "*Accessories Added:*" & vbCrLf & accessoriesList & vbCrLf
    Else
        msg = msg & "*Accessories:* None" & vbCrLf & vbCrLf
    End If
    
    ' ===== ASSIGNMENT DETAILS =====
   msg = msg & "------------------------" & vbCrLf & _
      "*Assigned To:* " & cmbAssignName.value & vbCrLf
    
    ' ===== COURIER DETAILS =====
    msg = msg & "*Courier:* " & cmbCourierMode.value & " - " & cmbCourierCompany.value & vbCrLf
    
    If Trim(txtDocketNumber.value) <> "" Then
        msg = msg & "*Docket No:* " & txtDocketNumber.value & vbCrLf
    End If
    
    If Trim(lblCourierMobile.caption & "") <> "" Then
        msg = msg & "*Courier Mobile:* " & lblCourierMobile.caption & vbCrLf
    End If
    
    ' ===== CHARGES & DATES =====
    msg = msg & "*Courier Charges:* Rs. " & txtCourierCharges.value & vbCrLf  ' <-- courierCharges ki jagah
    msg = msg & "*Send Date:* " & txtSendDate.value & vbCrLf
    
    If Trim(txtExpectedReturn.value) <> "" Then
        msg = msg & "*Expected Return:* " & txtExpectedReturn.value & vbCrLf
    End If
    
    ' Remarks
    If Trim(txtRemarks.value) <> "" Then
        msg = msg & "------------------------" & vbCrLf & _
              "*Remarks:* " & txtRemarks.value & vbCrLf
    End If
    
    msg = msg & "------------------------" & vbCrLf & _
          "Please acknowledge receipt." & vbCrLf & _
          "Thank You!"
    
    ' ===== SEND WHATSAPP =====
    Dim wsh As Object
    Dim url As String
    url = "whatsapp://send?phone=91" & mobile & "&text=" & URLEncode(msg)
    
    Set wsh = CreateObject("WScript.Shell")
    wsh.Run url, 1, False
    Set wsh = Nothing
    
    MsgBox "WhatsApp message sent!", vbInformation
    Exit Sub
    
ErrorHandler:
    MsgBox "WhatsApp Error: " & Err.Description, vbCritical
End Sub
' URL Encode Function
Private Function URLEncode(ByVal StringVal As String) As String
    On Error Resume Next
    
    Dim i As Long
    Dim CharCode As Integer
    Dim char As String
    Dim result As String
    
    result = ""
    
    For i = 1 To Len(StringVal)
        char = Mid(StringVal, i, 1)
        CharCode = Asc(char)
        
        ' ASCII Alphanumeric
        If (CharCode >= 48 And CharCode <= 57) Or _
           (CharCode >= 65 And CharCode <= 90) Or _
           (CharCode >= 97 And CharCode <= 122) Then
            result = result & char
        
        ' Space to %20
        ElseIf CharCode = 32 Then
            result = result & "%20"
        
        ' Line break (CR) - skip
        ElseIf CharCode = 13 Then
            ' Skip
        
        ' Line feed to %0A
        ElseIf CharCode = 10 Then
            result = result & "%0A"
        
        ' Common special characters
        ElseIf CharCode = 42 Then  ' *
            result = result & "%2A"
        ElseIf CharCode = 58 Then  ' :
            result = result & "%3A"
        ElseIf CharCode = 45 Then  ' -
            result = result & "%2D"
        ElseIf CharCode = 95 Then  ' _
            result = result & "_"
        ElseIf CharCode = 46 Then  ' .
            result = result & "."
        ElseIf CharCode = 64 Then  ' @
            result = result & "%40"
        ElseIf CharCode = 40 Then  ' (
            result = result & "%28"
        ElseIf CharCode = 41 Then  ' )
            result = result & "%29"
        ElseIf CharCode = 43 Then  ' +
            result = result & "%2B"
        ElseIf CharCode = 47 Then  ' /
            result = result & "%2F"
        ElseIf CharCode = 61 Then  ' =
            result = result & "%3D"
        ElseIf CharCode = 63 Then  ' ?
            result = result & "%3F"
        ElseIf CharCode = 38 Then  ' &
            result = result & "%26"
        ElseIf CharCode = 37 Then  ' %
            result = result & "%25"
        ElseIf CharCode = 35 Then  ' #
            result = result & "%23"
        
        ' Extended ASCII
        ElseIf CharCode > 127 Then
            result = result & "%" & Hex(CharCode)
        
        ' All other
        Else
            result = result & "%" & Right("0" & Hex(CharCode), 2)
        End If
    Next i
    
    URLEncode = result
End Function


'===========================================
' SEARCH BUTTON
'===========================================
Private Sub btnSearch_Click()
    Dim wsProduct As Worksheet
    Dim wsCustomer As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim searchText As String
    
    searchText = UCase(Trim(txtSearch.value))
    If searchText = "" Then Exit Sub
    
    Set wsProduct = ThisWorkbook.Sheets("Job_Product")
    Set wsCustomer = ThisWorkbook.Sheets("Customer_Master")
    lastRow = wsProduct.Cells(wsProduct.Rows.count, 1).End(xlUp).row
    
    ' Search in Job_Product
    For i = 2 To lastRow
        If InStr(UCase(Trim(wsProduct.Cells(i, 1).value)), searchText) > 0 Or _
           InStr(UCase(Trim(wsProduct.Cells(i, 6).value)), searchText) > 0 Or _
           InStr(UCase(Trim(wsProduct.Cells(i, 7).value)), searchText) > 0 Or _
           InStr(UCase(Trim(wsProduct.Cells(i, 8).value)), searchText) > 0 Or _
           InStr(UCase(Trim(wsProduct.Cells(i, 9).value)), searchText) > 0 Then
           
            If UCase(Trim(wsProduct.Cells(i, 3).value)) = "SERVICE" And _
               UCase(Trim(wsProduct.Cells(i, 4).value)) = "ACTIVE" Then
                cmbEntryID.value = wsProduct.Cells(i, 1).value
                Exit Sub
            End If
        End If
    Next i
    
    ' Search by Mobile/Name in Customer_Master
    Dim custRow As Long
    Dim custID As String
    
    For custRow = 2 To wsCustomer.Cells(wsCustomer.Rows.count, 1).End(xlUp).row
        If InStr(UCase(Trim(wsCustomer.Cells(custRow, 2).value)), searchText) > 0 Or _
           InStr(UCase(Trim(wsCustomer.Cells(custRow, 3).value)), searchText) > 0 Then
           
            custID = wsCustomer.Cells(custRow, 1).value
            
            For i = 2 To lastRow
                If UCase(Trim(wsProduct.Cells(i, 2).value)) = UCase(custID) And _
                   UCase(Trim(wsProduct.Cells(i, 3).value)) = "WARRANTY" And _
                   UCase(Trim(wsProduct.Cells(i, 4).value)) = "ACTIVE" Then
                    cmbEntryID.value = wsProduct.Cells(i, 1).value
                    Exit Sub
                End If
            Next i
        End If
    Next custRow
    
    MsgBox "No matching entry found!", vbExclamation
End Sub
Private Sub ClearForm()
    On Error Resume Next
    
    cmbEntryID.value = ""
    lblProductValue.caption = ""
    lblCompanyValue.caption = ""
    lblModelValue.caption = ""
    lblSerialValue.caption = ""
    lblCustomerValue.caption = ""
    lblMobileValue.caption = ""           ' FIXED HERE
    
    txtProblemDescription.value = ""
    
    cmbAssignTo.value = ""
    cmbAssignName.Clear
    
    lblVendorMobile.caption = ""
    lblVendorAddress.caption = ""
    lblWarrantyStatusValue.caption = ""
    
    cmbCourierMode.value = ""
    cmbCourierCompany.Clear
    txtDocketNumber.value = ""
    txtSendDate.value = Format(Now, "dd-MM-yyyy")
    txtExpectedReturn.value = ""
    txtCourierCharges.value = ""          ' Clear charges
    txtRemarks.value = ""                 ' Clear remarks
    
    CheckBox1.value = False
    CheckBox2.value = False
    CheckBox3.value = False
    CheckBox4.value = False
    CheckBox5.value = False
    CheckBox6.value = False
    
    imgCustomer.Picture = Nothing
    imgProductPhoto.Picture = Nothing
    
    lstPending.ListItems.Clear
    lblPendingCount.caption = "0"
    
    txtSendDate.value = Format(Now, "dd-MM-yyyy")
txtExpectedReturn.value = Format(Date + 7, "dd-mm-yyyy")

End Sub

'===========================================
' CANCEL BUTTON
'===========================================
Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub SetupCourierMode()
    cmbCourierMode.Clear
    cmbCourierMode.AddItem "BY SELF"
    cmbCourierMode.AddItem "COURIER"
    cmbCourierMode.AddItem "INDIAN POST"
    cmbCourierMode.AddItem "TRANSPORT"
End Sub

Private Sub btnAddCourierMode_Click()
    Dim newMode As String
    newMode = InputBox("Enter New Courier Mode:", "Add Courier Mode")
    If Trim(newMode) = "" Then Exit Sub
    
    Dim i As Long
    For i = 0 To cmbCourierMode.ListCount - 1
        If UCase(cmbCourierMode.List(i)) = UCase(newMode) Then
            MsgBox "Mode '" & newMode & "' already exists!", vbExclamation
            Exit Sub
        End If
    Next i
    
    cmbCourierMode.AddItem UCase(newMode)
    cmbCourierMode.value = UCase(newMode)
    MsgBox "Courier Mode Added: " & newMode, vbInformation
End Sub





'===========================================
' DATE VALIDATION
'===========================================
Private Sub txtSendDate_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    On Error Resume Next

If txtSendDate.value <> "" Then
    txtSendDate.value = Format(CDate(txtSendDate.value), "dd-mm-yyyy")
    
    ' ? AUTO Expected Return (+7 days)
    txtExpectedReturn.value = Format(CDate(txtSendDate.value) + 7, "dd-mm-yyyy")
End If

On Error GoTo 0
End Sub

Private Sub txtExpectedReturn_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    On Error Resume Next
    If txtExpectedReturn.value <> "" Then
        txtExpectedReturn.value = Format(CDate(txtExpectedReturn.value), "dd-mm-yyyy")
    End If
    On Error GoTo 0
End Sub
'===========================================
' ASSIGN BUTTON - ONLY SAVE DATA
'===========================================
Private Sub btnAssign_Click()
    On Error GoTo ErrorHandler
    
    ' ==========================================
    ' MANDATORY VALIDATION
    ' ==========================================
    
    If cmbEntryID.value = "" Or cmbEntryID.text = "" Then
        MsgBox "Please select Entry ID first!", vbExclamation, "Required Field"
        cmbEntryID.SetFocus
        Exit Sub
    End If
    
    If cmbAssignTo.value = "" Or cmbAssignTo.text = "" Then
        MsgBox "Please select Assign To!", vbExclamation, "Required Field"
        cmbAssignTo.SetFocus
        Exit Sub
    End If
    
    If cmbAssignName.value = "" Or cmbAssignName.text = "" Then
        MsgBox "Please select Assign Name!", vbExclamation, "Required Field"
        cmbAssignName.SetFocus
        Exit Sub
    End If
    
    If cmbCourierMode.value = "" Or cmbCourierMode.text = "" Then
        MsgBox "Please select Transport Mode!", vbExclamation, "Required Field"
        cmbCourierMode.SetFocus
        Exit Sub
    End If
    
    ' ==========================================
    ' EDIT MODE = UPDATE EXISTING RECORD
    ' ==========================================
    If isEditMode And EditAssignRow > 0 Then
        Call UpdateExistingAssignment
        Exit Sub
    End If
    
    ' ==========================================
    ' NEW MODE = SAVE NEW RECORD
    ' ==========================================
    Call SaveNewAssignment
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error: " & Err.Number & " - " & Err.Description, vbCritical
End Sub

'===========================================
' UPDATE EXISTING ASSIGNMENT (EDIT MODE)
'===========================================
Private Sub UpdateExistingAssignment()
    Dim ws As Worksheet
    Dim wsJob As Worksheet
    Dim accList As String
    Dim i As Integer
    
    Set ws = ThisWorkbook.Sheets("Assign_Master")
    Set wsJob = ThisWorkbook.Sheets("Job_Product")
    
    ' Build Accessories List
    accList = ""
    If CheckBox1.value = True Then accList = accList & CheckBox1.caption & ", "
    If CheckBox2.value = True Then accList = accList & CheckBox2.caption & ", "
    If CheckBox3.value = True Then accList = accList & CheckBox3.caption & ", "
    If CheckBox4.value = True Then accList = accList & CheckBox4.caption & ", "
    If CheckBox5.value = True Then accList = accList & CheckBox5.caption & ", "
    If CheckBox6.value = True Then accList = accList & CheckBox6.caption & ", "
    If Len(accList) > 2 Then accList = Left(accList, Len(accList) - 2)
    
    ' Get Company Details
    Dim wsConfig As Worksheet
    Dim compName As String, compAddr As String, compMobile As String, compEmail As String
    Dim addr1 As String, addr2 As String, city As String, pin As String, state As String
    
    Set wsConfig = ThisWorkbook.Sheets("Software_Config")
    compName = Trim(wsConfig.Range("B2").value)
    compMobile = Trim(wsConfig.Range("B9").value)
    compEmail = Trim(wsConfig.Range("B10").value)
    addr1 = Trim(wsConfig.Range("B3").value)
    addr2 = Trim(wsConfig.Range("B4").value)
    city = Trim(wsConfig.Range("B5").value)
    pin = Trim(wsConfig.Range("B6").value)
    state = Trim(wsConfig.Range("B7").value)
    
    compAddr = ""
    If addr1 <> "" Then compAddr = addr1
    If addr2 <> "" Then compAddr = compAddr & IIf(compAddr <> "", ", ", "") & addr2
    If city <> "" Then compAddr = compAddr & IIf(compAddr <> "", ", ", "") & city
    If pin <> "" Then compAddr = compAddr & " - " & pin
    If state <> "" Then compAddr = compAddr & ", " & state
    
    ' Get AssignToID
    Dim wsReg As Worksheet, regRow As Long, assignToID As String
    assignToID = ""
    Select Case UCase(Trim(cmbAssignTo.value))
        Case "VENDOR": Set wsReg = ThisWorkbook.Sheets("Vendor_Register")
        Case "ENGINEER": Set wsReg = ThisWorkbook.Sheets("Engineer_Register")
        Case "INHOUSE": Set wsReg = ThisWorkbook.Sheets("Inhouse_Register")
        Case "SERVICE_STATION": Set wsReg = ThisWorkbook.Sheets("ServiceStation_Register")
    End Select
    
    If Not wsReg Is Nothing Then
        For regRow = 2 To wsReg.Cells(wsReg.Rows.count, 1).End(xlUp).row
            If UCase(Trim(wsReg.Cells(regRow, 2).value)) = UCase(Trim(cmbAssignName.value)) Then
                assignToID = Trim(wsReg.Cells(regRow, 1).value)
                Exit For
            End If
        Next regRow
    End If
    
    ' ===== UPDATE EXISTING ROW =====
    With ws.Rows(EditAssignRow)
        .Cells(1, 11).value = cmbAssignTo.value                ' K - AssignType
        .Cells(1, 12).value = cmbCourierMode.value             ' L - AssignMode
        .Cells(1, 13).value = assignToID                       ' M - AssignToID
        .Cells(1, 14).value = cmbAssignName.value              ' N - AssignToName
        .Cells(1, 15).value = lblVendorMobile.caption          ' O - AssignToMobile
        .Cells(1, 16).value = lblVendorEmail.caption           ' P - AssignToEmail
        .Cells(1, 17).value = lblVendorAddress.caption         ' Q - AssignToAddress
        .Cells(1, 18).value = compName                         ' R - OurCompanyName
        .Cells(1, 19).value = compAddr                         ' S - OurCompanyAddress
        .Cells(1, 20).value = compMobile                       ' T - OurCompanyMobile
        .Cells(1, 21).value = compEmail                        ' U - OurCompanyEmail
        .Cells(1, 23).value = cmbCourierMode.value             ' W - CourierMode
        .Cells(1, 24).value = cmbCourierCompany.value          ' X - CourierName
        .Cells(1, 25).value = txtDocketNumber.value            ' Y - DocketNumber
        .Cells(1, 26).value = txtSendDate.value                ' Z - SendDate
        .Cells(1, 27).value = txtExpectedReturn.value          ' AA - ExpectedReturnDate
        .Cells(1, 28).value = txtCourierCharges.value          ' AB - CourierCharges
        .Cells(1, 29).value = txtProblemDescription.value      ' AC - ProblemDescription
        .Cells(1, 30).value = accList                          ' AD - AccessoriesSent
        .Cells(1, 31).value = txtRemarks.value                 ' AE - Remarks
        .Cells(1, 34).value = Format(Now, "dd-MM-yyyy hh:mm:ss") ' AH - ModifiedDate
        ' AG - ModifiedBy
        .Cells(1, 33).value = Environ("Username") & " (EDIT)"  ' AG - CreatedBy (Edited)
    End With
    
    ' ===== UPDATE ACCESSORIES =====
    ' ?????? accessories ????? ????
    Dim wsAcc As Worksheet, accRow As Long, accLastRow As Long
    On Error Resume Next
    Set wsAcc = ThisWorkbook.Sheets("Assign_Accessories")
    On Error GoTo 0
    
    If Not wsAcc Is Nothing Then
        accLastRow = wsAcc.Cells(wsAcc.Rows.count, 1).End(xlUp).row
        ' EntryID ?? ??? ???? ?????
        For accRow = accLastRow To 2 Step -1
            If UCase(Trim(wsAcc.Cells(accRow, 2).value)) = UCase(cmbEntryID.value) Then
                wsAcc.Rows(accRow).Delete
            End If
        Next accRow
        
        ' ?? accessories ??? ????
        Dim chk As Control
        For Each chk In Me.Controls
            If TypeName(chk) = "CheckBox" Then
                If chk.value = True And Trim(chk.caption) <> "" Then
                    accRow = wsAcc.Cells(wsAcc.Rows.count, 1).End(xlUp).row + 1
                    wsAcc.Cells(accRow, 1).value = LastAssignID
                    wsAcc.Cells(accRow, 2).value = cmbEntryID.value
                    wsAcc.Cells(accRow, 3).value = chk.caption
                    wsAcc.Cells(accRow, 4).value = "ADDED"
                End If
            End If
        Next chk
    End If
    
        ' ===== UPDATE Job_Product STATUS =====
    Dim wsJob As Worksheet
    On Error Resume Next
    Set wsJob = ThisWorkbook.Sheets("Job_Product")
    On Error GoTo 0
    
    If Not wsJob Is Nothing Then
        Dim jobLastRow As Long
        jobLastRow = wsJob.Cells(wsJob.Rows.count, 1).End(xlUp).row
        For i = 2 To jobLastRow
            If UCase(Trim(wsJob.Cells(i, 1).value)) = UCase(Trim(cmbEntryID.value)) Then
                wsJob.Cells(i, 4).value = "ASSIGNED"  ' ?? ?? ?? Status ??
                wsJob.Cells(i, 5).value = Format(Now, "dd-MM-yyyy hh:mm:ss")
                Exit For
            End If
        Next i
    End If
    
    ' ===== NOW ENABLE ACTION BUTTONS (After UPDATE) =====
    IsDataSaved = True
    btnSendWhatsApp.enabled = True
    btnSendEmail.enabled = True
    btnPrintAssign.enabled = True
    btnExportPDF.enabled = True
    
    MsgBox "Assignment Updated Successfully!", vbInformation, "UPDATE COMPLETE"
    
    Exit Sub
    
End Sub

'===========================================
' SAVE NEW ASSIGNMENT (NEW MODE)
'===========================================
Private Sub SaveNewAssignment()
    Dim ws As Worksheet, wsJob As Worksheet
    Dim lastRow As Long, assignID As String
    Dim accList As String
    Dim i As Integer
    
    Set ws = ThisWorkbook.Sheets("Assign_Master")
    Set wsJob = ThisWorkbook.Sheets("Job_Product")
    
    ' Build Accessories List
    accList = ""
    If CheckBox1.value = True Then accList = accList & CheckBox1.caption & ", "
    If CheckBox2.value = True Then accList = accList & CheckBox2.caption & ", "
    If CheckBox3.value = True Then accList = accList & CheckBox3.caption & ", "
    If CheckBox4.value = True Then accList = accList & CheckBox4.caption & ", "
    If CheckBox5.value = True Then accList = accList & CheckBox5.caption & ", "
    If CheckBox6.value = True Then accList = accList & CheckBox6.caption & ", "
    If Len(accList) > 2 Then accList = Left(accList, Len(accList) - 2)
    
    ' Company Details
    Dim wsConfig As Worksheet
    Dim compName As String, compAddr As String, compMobile As String, compEmail As String
    Dim addr1 As String, addr2 As String, city As String, pin As String, state As String
    
    Set wsConfig = ThisWorkbook.Sheets("Software_Config")
    compName = Trim(wsConfig.Range("B2").value)
    compMobile = Trim(wsConfig.Range("B9").value)
    compEmail = Trim(wsConfig.Range("B10").value)
    addr1 = Trim(wsConfig.Range("B3").value)
    addr2 = Trim(wsConfig.Range("B4").value)
    city = Trim(wsConfig.Range("B5").value)
    pin = Trim(wsConfig.Range("B6").value)
    state = Trim(wsConfig.Range("B7").value)
    
    compAddr = ""
    If addr1 <> "" Then compAddr = addr1
    If addr2 <> "" Then compAddr = compAddr & IIf(compAddr <> "", ", ", "") & addr2
    If city <> "" Then compAddr = compAddr & IIf(compAddr <> "", ", ", "") & city
    If pin <> "" Then compAddr = compAddr & " - " & pin
    If state <> "" Then compAddr = compAddr & ", " & state
    
    ' Customer ID
    Dim custID As String, jobRow As Long
    custID = ""
    For jobRow = 2 To wsJob.Cells(wsJob.Rows.count, 1).End(xlUp).row
        If UCase(Trim(wsJob.Cells(jobRow, 1).value)) = UCase(Trim(cmbEntryID.value)) Then
            custID = Trim(wsJob.Cells(jobRow, 2).value)
            Exit For
        End If
    Next jobRow
    
    ' AssignToID
    Dim wsReg As Worksheet, regRow As Long, assignToID As String
    assignToID = ""
    Select Case UCase(Trim(cmbAssignTo.value))
        Case "VENDOR": Set wsReg = ThisWorkbook.Sheets("Vendor_Register")
        Case "ENGINEER": Set wsReg = ThisWorkbook.Sheets("Engineer_Register")
        Case "INHOUSE": Set wsReg = ThisWorkbook.Sheets("Inhouse_Register")
        Case "SERVICE_STATION": Set wsReg = ThisWorkbook.Sheets("ServiceStation_Register")
    End Select
    
    If Not wsReg Is Nothing Then
        For regRow = 2 To wsReg.Cells(wsReg.Rows.count, 1).End(xlUp).row
            If UCase(Trim(wsReg.Cells(regRow, 2).value)) = UCase(Trim(cmbAssignName.value)) Then
                assignToID = Trim(wsReg.Cells(regRow, 1).value)
                Exit For
            End If
        Next regRow
    End If
    
    ' Save to Assign_Master
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row + 1
    assignID = "ASN" & Format(lastRow - 1, "00000")
    
    With ws.Rows(lastRow)
        .Cells(1, 1).value = assignID                          ' A
        .Cells(1, 2).value = cmbEntryID.value                  ' B
        .Cells(1, 3).value = custID                            ' C
        .Cells(1, 4).value = lblMobileValue.caption            ' D
        .Cells(1, 5).value = lblCustomerValue.caption          ' E
        .Cells(1, 6).value = lblProductValue.caption           ' F
        .Cells(1, 7).value = lblCompanyValue.caption           ' G
        .Cells(1, 8).value = lblModelValue.caption             ' H
        .Cells(1, 9).value = lblSerialValue.caption            ' I
        .Cells(1, 10).value = lblWarrantyStatusValue.caption   ' J
        .Cells(1, 11).value = cmbAssignTo.value                ' K
        .Cells(1, 12).value = cmbCourierMode.value             ' L
        .Cells(1, 13).value = assignToID                       ' M
        .Cells(1, 14).value = cmbAssignName.value              ' N
        .Cells(1, 15).value = lblVendorMobile.caption          ' O
        .Cells(1, 16).value = lblVendorEmail.caption           ' P
        .Cells(1, 17).value = lblVendorAddress.caption         ' Q
        .Cells(1, 18).value = compName                         ' R
        .Cells(1, 19).value = compAddr                         ' S
        .Cells(1, 20).value = compMobile                       ' T
        .Cells(1, 21).value = compEmail                        ' U
        .Cells(1, 22).value = Format(Date, "dd-mm-yyyy")       ' V
        .Cells(1, 23).value = cmbCourierMode.value             ' W
        .Cells(1, 24).value = cmbCourierCompany.value          ' X
        .Cells(1, 25).value = txtDocketNumber.value            ' Y
        .Cells(1, 26).value = txtSendDate.value                ' Z
        .Cells(1, 27).value = txtExpectedReturn.value          ' AA
        .Cells(1, 28).value = txtCourierCharges.value          ' AB
        .Cells(1, 29).value = txtProblemDescription.value      ' AC
        .Cells(1, 30).value = accList                          ' AD
        .Cells(1, 31).value = txtRemarks.value                 ' AE
        .Cells(1, 32).value = "PENDING"                        ' AF
        .Cells(1, 33).value = Environ("Username")              ' AG
        .Cells(1, 34).value = Format(Now, "dd-MM-yyyy hh:mm:ss") ' AH
    End With
    ' Update Job_Product Status
For jobRow = 2 To wsJob.Cells(wsJob.Rows.count, 1).End(xlUp).row
    If UCase(Trim(wsJob.Cells(jobRow, 1).value)) = UCase(Trim(cmbEntryID.value)) Then
        wsJob.Cells(jobRow, 4).value = "ASSIGNED"
        wsJob.Cells(jobRow, 5).value = Format(Now, "dd-MM-yyyy hh:mm:ss")  ' ? ADD Date
        Exit For
    End If
Next jobRow
    
    ' Save Accessories
    If accList <> "" Then
        Dim wsAcc As Worksheet, accRow As Long
        Set wsAcc = ThisWorkbook.Sheets("Assign_Accessories")
        Dim chk As Control
        For Each chk In Me.Controls
            If TypeName(chk) = "CheckBox" Then
                If chk.value = True And Trim(chk.caption) <> "" Then
                    accRow = wsAcc.Cells(wsAcc.Rows.count, 1).End(xlUp).row + 1
                    wsAcc.Cells(accRow, 1).value = assignID
                    wsAcc.Cells(accRow, 2).value = cmbEntryID.value
                    wsAcc.Cells(accRow, 3).value = chk.caption
                    wsAcc.Cells(accRow, 4).value = "ADDED"
                End If
            End If
        Next chk
    End If
    
    ' Set Flags & Enable Buttons
    IsDataSaved = True
    LastAssignID = assignID
    
    btnSendWhatsApp.enabled = True
    btnSendEmail.enabled = True
    btnPrintAssign.enabled = True
    btnClear.enabled = True
    btnExportPDF.enabled = True
    
    MsgBox "Assignment Saved Successfully!" & vbCrLf & "Assign ID: " & assignID, vbInformation
End Sub



Private Sub SaveAssignment()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim assignID As String
    Dim selectedAcc As String
    Dim i As Integer
    
    ' Get selected accessories
    selectedAcc = ""
    For i = 1 To 6
        If Me.Controls("CheckBox" & i).value = True Then
            If selectedAcc <> "" Then selectedAcc = selectedAcc & ", "
            selectedAcc = selectedAcc & Me.Controls("CheckBox" & i).caption
        End If
    Next i
    
    Set ws = ThisWorkbook.Sheets("Assign_Master")
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    assignID = "ASN" & Format(lastRow, "00000")
    
    With ws.Rows(lastRow + 1)
        .Cells(1, 1).value = assignID
        .Cells(1, 2).value = cmbEntryID.value
        .Cells(1, 3).value = "SERVICE"
        .Cells(1, 4).value = Format(Now, "dd-mm-yyyy")
        .Cells(1, 5).value = lblVendorName.caption
        .Cells(1, 6).value = cmbAssignTo.value
        .Cells(1, 7).value = lblVendorMobile.caption
        .Cells(1, 8).value = txtExpectedReturn.value
        .Cells(1, 9).value = txtSendDate.value
        .Cells(1, 10).value = cmbCourierMode.value
        .Cells(1, 11).value = cmbCourierCompany.value
        .Cells(1, 12).value = txtDocketNumber.value
        .Cells(1, 13).value = val(txtCourierCharges.value)
        .Cells(1, 14).value = selectedAcc
        .Cells(1, 15).value = "PENDING"
        .Cells(1, 16).value = lblCourierMobile.caption
        .Cells(1, 17).value = Format(Now, "dd-mm-yyyy hh:mm")
    End With
    
    Dim wsJob As Worksheet
Set wsJob = ThisWorkbook.Sheets("Job_Product")

Dim i As Long

For i = 2 To wsJob.Cells(wsJob.Rows.count, 1).End(xlUp).row
    
    If wsJob.Cells(i, 2).value = cmbEntryID.value Then
        
        wsJob.Cells(i, 4).value = "ASSIGNED"
        
        Exit For
        
    End If
    
Next i
End Sub

Private Sub btnAddCourierCompany_Click()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim newCompany As String
    Dim newMobile As String

    If Trim(cmbCourierMode.value) = "" Then
        MsgBox "Select Mode first!", vbExclamation
        Exit Sub
    End If

    newCompany = InputBox("Enter Name:")
    If Trim(newCompany) = "" Then Exit Sub

    newMobile = InputBox("Enter Mobile:")
    If Trim(newMobile) = "" Then Exit Sub

    Set ws = ThisWorkbook.Sheets("Courier_Data")
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row + 1

    ws.Cells(lastRow, 1).value = cmbCourierMode.value
    ws.Cells(lastRow, 2).value = newCompany
    ws.Cells(lastRow, 3).value = newMobile

    MsgBox "Saved Successfully!", vbInformation

    Call LoadCourierData

End Sub
Sub LoadCourierData()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long

    cmbCourierCompany.Clear

    Set ws = ThisWorkbook.Sheets("Courier_Data")
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row

    For i = 2 To lastRow
        If ws.Cells(i, 1).value = cmbCourierMode.value Then
            cmbCourierCompany.AddItem ws.Cells(i, 2).value
            cmbCourierCompany.List(cmbCourierCompany.ListCount - 1, 1) = ws.Cells(i, 3).value
        End If
    Next i

End Sub
Private Sub cmbCourierMode_Change()

    lblCourierMobile.caption = ""
    Call LoadCourierData

End Sub
Private Sub cmbCourierCompany_Change()

    If cmbCourierCompany.ListIndex <> -1 Then
        lblCourierMobile.caption = cmbCourierCompany.Column(1)
    End If

End Sub
'===========================================
' CLEAR - RESET EVERYTHING
'===========================================
Private Sub btnClear_Click()
    ' Clear all fields
    ClearForm
    
    ' Reset flags
    IsDataSaved = False
    LastAssignID = ""
    
    ' Reset buttons
    btnAssign.enabled = True
    btnSendWhatsApp.enabled = False
    btnSendEmail.enabled = False
    btnPrintAssign.enabled = False
    btnClear.enabled = False
    
    ' Refresh Entry IDs (load new ones)
    LoadEntryIDs
End Sub






'===========================================
' PRINT ASSIGN - ONLY IF SAVED
'===========================================
Private Sub btnPrintAssign_Click()
    If Not IsDataSaved Then
        MsgBox "Please save assignment first!", vbExclamation
        Exit Sub
    End If
    
    Call PrintAssignmentPro
    
    MsgBox "Print Complete!", vbInformation
End Sub
Private Sub PrintAssignmentPro()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet, shp As Shape, r As Long, i As Integer, idx As Integer
    Dim compName$, addr1$, addr2$, city$, state$, pin$, email$, mobile$, logoPath$, logoFullPath$
    Dim entryID$, assignDate$, vendName$, vendMobile$, vendAddress$
    Dim probText$, remText$, courierMode$, courierName$, docketNo$
    Dim sendDate$, courierMobile$, charges$, accText$
    Dim prodPhotoPath$, logoFolder$, fileName$, csp As String, chk As Control
    Dim prodName$, compProd$, model$, serial$, warranty$, warrantyRem$

    ' === DATA GATHERING ===
    On Error Resume Next
    
    entryID = cmbEntryID.value
    assignDate = Format(Date, "dd-mmm-yyyy")
    vendName = lblVendorName.caption
    vendMobile = lblVendorMobile.caption
    vendAddress = lblVendorAddress.caption
    vendEmail = lblVendorEmail.caption
    prodName = lblProductValue.caption
    compProd = lblCompanyValue.caption
    model = lblModelValue.caption
    serial = lblSerialValue.caption
    warranty = lblWarrantyStatusValue.caption
    
    If Not Me.Controls("lblWarrantyRemaining") Is Nothing Then
        warrantyRem = Me.Controls("lblWarrantyRemaining").caption
    Else
        warrantyRem = "N/A"
    End If
    
    If Trim(prodName) = "" Then prodName = "Not Available"
    If Trim(compProd) = "" Then compProd = "Not Available"
    If Trim(model) = "" Then model = "Not Available"
    If Trim(serial) = "" Then serial = "Not Available"
    If Trim(warranty) = "" Then warranty = "Not Available"
    
    probText = txtProblemDescription.value
    If Trim(probText) = "" Then probText = "Not Specified"
    
    accText = ""
    idx = 1
    For Each chk In Me.Controls
        If TypeName(chk) = "CheckBox" Then
            If chk.value = True Then
                If accText <> "" Then accText = accText & ", "
                accText = accText & idx & ". " & chk.caption
                idx = idx + 1
            End If
        End If
    Next chk
    If accText = "" Then accText = "No Accessories"
    
    remText = txtRemarks.value
    courierMode = cmbCourierMode.value
    courierName = cmbCourierCompany.value
    docketNo = txtDocketNumber.value
    sendDate = txtSendDate.value
    charges = txtCourierCharges.value
    courierMobile = lblCourierMobile.caption
    
    On Error GoTo ErrorHandler
    
    ' === WORKSHEET SETUP ===
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Assignment_Print")
    On Error GoTo ErrorHandler
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add
        ws.name = "Assignment_Print"
    End If
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    ws.Cells.ClearContents
    ws.Cells.Borders.LineStyle = xlNone
    ws.Cells.Interior.ColorIndex = xlNone
    
    On Error Resume Next
    For Each shp In ws.Shapes: shp.Delete: Next
    On Error GoTo ErrorHandler
    
    ' === PAGE SETUP ===
    With ws.PageSetup
        .PaperSize = xlPaperA4
        .Orientation = xlPortrait
        .LeftMargin = Application.InchesToPoints(0.5)
        .RightMargin = Application.InchesToPoints(0.5)
        .TopMargin = Application.InchesToPoints(0.5)
        .BottomMargin = Application.InchesToPoints(0.5)
        .HeaderMargin = Application.InchesToPoints(0.2)
        .FooterMargin = Application.InchesToPoints(0.2)
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .Zoom = False
    End With
    
    ' Column Widths
    ws.Columns("A").ColumnWidth = 1
    ws.Columns("B").ColumnWidth = 16
    ws.Columns("C").ColumnWidth = 20
    ws.Columns("D").ColumnWidth = 15
    ws.Columns("E").ColumnWidth = 15
    ws.Columns("F").ColumnWidth = 13
    ws.Columns("G").ColumnWidth = 13
    ws.Columns("H").ColumnWidth = 1

    ' === COMPANY DETAILS FROM CONFIG ===
    With ThisWorkbook.Sheets("Software_Config")
        compName = .Range("B2").value
        addr1 = .Range("B3").value: addr2 = .Range("B4").value
        city = .Range("B5").value: state = .Range("B7").value: pin = .Range("B6").value
        mobile = .Range("B9").value: email = .Range("B10").value: logoPath = .Range("B11").value
    End With
    
    '=== HEADER SECTION ===
    r = 1
    ws.Range("B" & r & ":E" & r).Merge
    ws.Range("B" & r).value = compName
    ws.Range("B" & r).Font.Size = 18
    ws.Range("B" & r).Font.Bold = True
    ws.Range("B" & r).Font.name = "Arial"
    ws.Rows(r).RowHeight = 28
    
    r = 2: ws.Range("B" & r & ":E" & r).Merge: ws.Range("B" & r).value = addr1: ws.Rows(r).RowHeight = 16
    
    If Trim(addr2) <> "" Then
        r = 3: ws.Range("B" & r & ":E" & r).Merge: ws.Range("B" & r).value = addr2: ws.Rows(r).RowHeight = 16
        r = 4
    Else
        r = 3
    End If
    
    csp = city
    If Trim(state) <> "" Then csp = csp & ", " & state
    If Trim(pin) <> "" Then csp = csp & " - " & pin
    
    ws.Range("B" & r & ":E" & r).Merge: ws.Range("B" & r).value = csp
    ws.Range("B" & r).Font.Size = 10: ws.Rows(r).RowHeight = 16
    
    r = r + 1: ws.Range("B" & r & ":E" & r).Merge: ws.Range("B" & r).value = "Email: " & email: ws.Range("B" & r).Font.Size = 11: ws.Rows(r).RowHeight = 14
    r = r + 1: ws.Range("B" & r & ":E" & r).Merge: ws.Range("B" & r).value = "Mobile: " & mobile: ws.Range("B" & r).Font.Size = 11: ws.Rows(r).RowHeight = 14
    
    '=== LOGO ===
    logoFullPath = ""
    If Trim(logoPath) <> "" Then
        If Dir(logoPath) <> "" Then logoFullPath = logoPath
    End If
    
    If logoFullPath = "" Then
        logoFolder = ThisWorkbook.path & "\GLOBAL_SOFT_DATA\Images\Logo\"
        If Dir(logoFolder, vbDirectory) = "" Then logoFolder = ThisWorkbook.path & "\LOGO\"
        fileName = Dir(logoFolder & "*.png")
        If fileName = "" Then fileName = Dir(logoFolder & "*.jpg")
        If fileName <> "" Then logoFullPath = logoFolder & fileName
    End If
    
        If logoFullPath <> "" Then
        On Error Resume Next
        ' Logo size ?? ?? ?? ?? position ??? ???
        ws.Shapes.AddPicture logoFullPath, msoFalse, msoTrue, _
            ws.Range("F1").Left, ws.Range("B1").Top, 120, 80  ' 150,100 ?? 120,80
        On Error GoTo ErrorHandler
    End If
    
    '=== TITLE ===
    r = 7
    ws.Range("B" & r & ":G" & r).Merge
    ws.Range("B" & r).value = "ASSIGNMENT REPORT"
    ws.Range("B" & r).Font.Size = 18
    ws.Range("B" & r).Font.Bold = True
    ws.Range("B" & r).HorizontalAlignment = xlCenter
    ws.Range("B" & r).Interior.Color = RGB(0, 102, 204)
    ws.Range("B" & r).Font.Color = RGB(255, 255, 255)
    ws.Rows(r).RowHeight = 22
    
    '=== ID & DATE ===
       r = 8: ws.Rows(r).RowHeight = 22
    ' Left side - Assignment ID
    ws.Range("B" & r).value = "Assignment ID: " & entryID
    ws.Range("B" & r).Font.Bold = True
    ws.Range("B" & r).Font.Size = 11
    ws.Range("B" & r).HorizontalAlignment = xlLeft
    ' Right side - Date
    ws.Range("G" & r).value = "Date: " & assignDate
    ws.Range("G" & r).Font.Bold = True
    ws.Range("G" & r).Font.Size = 11
    ws.Range("G" & r).HorizontalAlignment = xlRight
    
    '=== ASSIGN TO DETAILS ===
    r = 10
    ws.Range("B" & r & ":G" & r).Merge
    ws.Range("B" & r).value = "ASSIGN TO DETAILS"
    ws.Range("B" & r).Font.Size = 12
    ws.Range("B" & r).Font.Bold = True
    ws.Range("B" & r).Interior.Color = RGB(0, 0, 128)
    ws.Range("B" & r).Font.Color = RGB(255, 255, 255)
    ws.Range("B" & r).HorizontalAlignment = xlLeft
    ws.Rows(r).RowHeight = 18
    
    r = 11: ws.Rows(r).RowHeight = 16
    ws.Range("B" & r).value = "Name:": ws.Range("B" & r).Font.Bold = True: ws.Range("B" & r).Font.Size = 11
    ws.Range("C" & r & ":E" & r).Merge: ws.Range("C" & r).value = vendName: ws.Range("C" & r).Font.Size = 11
    ws.Range("C" & r).HorizontalAlignment = xlLeft
    
    r = 12: ws.Rows(r).RowHeight = 16
    ws.Range("B" & r).value = "Mobile:": ws.Range("B" & r).Font.Bold = True: ws.Range("B" & r).Font.Size = 11
    ws.Range("C" & r & ":D" & r).Merge: ws.Range("C" & r).value = vendMobile: ws.Range("C" & r).Font.Size = 11
    ws.Range("C" & r).HorizontalAlignment = xlLeft
    
    r = 13: ws.Rows(r).RowHeight = 16
    ws.Range("B" & r).value = "Address:": ws.Range("B" & r).Font.Bold = True: ws.Range("B" & r).Font.Size = 11
    ws.Range("C" & r & ":G" & r).Merge: ws.Range("C" & r).value = vendAddress: ws.Range("C" & r).Font.Size = 11
    ws.Range("C" & r).HorizontalAlignment = xlLeft
    
    ' === PRODUCT DETAILS ===
    On Error Resume Next
    ws.Range("B15:G15").Merge
    ws.Range("B15").value = "PRODUCT DETAILS (SERVICE ID: " & entryID & ")"
    ws.Range("B15").Font.Size = 12
    ws.Range("B15").Font.Bold = True
    ws.Range("B15").Interior.Color = RGB(255, 20, 147)
    ws.Range("B15").Font.Color = RGB(255, 255, 255)
    ws.Rows(15).RowHeight = 18
    
    ws.Range("B16").value = "Product:": ws.Range("B16").Font.Bold = True: ws.Range("B16").Font.Size = 11
    ws.Range("C16:E16").Merge: ws.Range("C16").value = lblProductValue.caption: ws.Range("C16").Font.Size = 11
    ws.Rows(16).RowHeight = 16
    
    ws.Range("B17").value = "Company:": ws.Range("B17").Font.Bold = True: ws.Range("B17").Font.Size = 11
    ws.Range("C17:E17").Merge: ws.Range("C17").value = lblCompanyValue.caption: ws.Range("C17").Font.Size = 11
    ws.Rows(17).RowHeight = 16
    
    ws.Range("B18").value = "Model:": ws.Range("B18").Font.Bold = True: ws.Range("B18").Font.Size = 11
    ws.Range("C18").value = lblModelValue.caption: ws.Range("C18").Font.Size = 11
    ws.Rows(18).RowHeight = 16
    
    ws.Range("B19").value = "Serial:": ws.Range("B19").Font.Bold = True: ws.Range("B19").Font.Size = 11
    ws.Range("C19").value = lblSerialValue.caption: ws.Range("C19").Font.Size = 11
    ws.Rows(19).RowHeight = 16
    
    ws.Range("B20").value = "Warranty:": ws.Range("B20").Font.Bold = True: ws.Range("B20").Font.Size = 11
    ws.Range("C20:F20").Merge: ws.Range("C20").value = lblWarrantyStatusValue.caption & IIf(warrantyRem <> "N/A", " (" & warrantyRem & ")", ""): ws.Range("C20").Font.Size = 11
    ws.Rows(20).RowHeight = 16
    
    Dim photoFile As String
    photoFile = ThisWorkbook.path & "\GLOBAL_SOFT_DATA\Images\Products\" & entryID & "\Photo1.jpg"
    If Dir(photoFile) = "" Then photoFile = ThisWorkbook.path & "\GLOBAL_SOFT_DATA\Images\Products\" & entryID & ".jpg"
    If Dir(photoFile) <> "" Then
        On Error Resume Next
        ws.Shapes.AddPicture photoFile, msoFalse, msoTrue, ws.Range("F16").Left, ws.Range("F16").Top, 120, 90
        On Error GoTo ErrorHandler
    End If
    On Error GoTo ErrorHandler
    
    '=== ACCESSORIES ===
    r = 22
    ws.Range("B" & r & ":G" & r).Merge
    ws.Range("B" & r).value = "SEND ACCESSORIES"
    ws.Range("B" & r).Font.Size = 12
    ws.Range("B" & r).Font.Bold = True
    ws.Range("B" & r).Interior.Color = RGB(0, 153, 76)
    ws.Range("B" & r).Font.Color = RGB(255, 255, 255)
    ws.Rows(r).RowHeight = 18
    
    r = 23
    ws.Rows(r).RowHeight = 18
    ws.Range("B" & r & ":G" & r).Merge
    ws.Range("B" & r).value = accText
    ws.Range("B" & r).Font.Size = 11
    ws.Range("B" & r).WrapText = True
    
    '=== PROBLEM ===
    r = 25
    ws.Range("B" & r & ":G" & r).Merge
    ws.Range("B" & r).value = "PROBLEM"
    ws.Range("B" & r).Font.Size = 12
    ws.Range("B" & r).Font.Bold = True
    ws.Range("B" & r).Interior.Color = RGB(204, 102, 0)
    ws.Range("B" & r).Font.Color = RGB(255, 255, 255)
    ws.Rows(r).RowHeight = 18
    
    r = 26
    ws.Rows(r).RowHeight = 20
    ws.Range("B" & r & ":G" & r).Merge
    ws.Range("B" & r).value = probText
    ws.Range("B" & r).Font.Size = 11
    ws.Range("B" & r).WrapText = True
    
    '=== REMARKS ===
    r = 28
    ws.Range("B" & r & ":G" & r).Merge
    ws.Range("B" & r).value = "REMARKS"
    ws.Range("B" & r).Font.Size = 12
    ws.Range("B" & r).Font.Bold = True
    ws.Range("B" & r).Interior.Color = RGB(178, 134, 0)
    ws.Range("B" & r).Font.Color = RGB(255, 255, 255)
    ws.Rows(r).RowHeight = 18
    
    r = 29
    ws.Rows(r).RowHeight = 20
    ws.Range("B" & r & ":G" & r).Merge
    ws.Range("B" & r).value = remText
    ws.Range("B" & r).Font.Size = 12
    ws.Range("B" & r).WrapText = True
        '=== TRANSPORT DETAILS ===
    r = 31
    ws.Range("B" & r & ":G" & r).Merge
    ws.Range("B" & r).value = "TRANSPORT DETAILS"
    ws.Range("B" & r).Font.Size = 12
    ws.Range("B" & r).Font.Bold = True
    ws.Range("B" & r).Interior.Color = RGB(0, 153, 153)
    ws.Range("B" & r).Font.Color = RGB(255, 255, 255)
    ws.Rows(r).RowHeight = 18
    
    ' Line 1: Transport Mode (Left) + Transport Name (Right)
    r = 32: ws.Rows(r).RowHeight = 18
    ws.Range("B" & r).value = "Transport Mode:"
    ws.Range("B" & r).Font.Bold = True: ws.Range("B" & r).Font.Size = 11
    ws.Range("C" & r & ":D" & r).Merge
    ws.Range("C" & r).value = courierMode
    ws.Range("C" & r).Font.Size = 11
    ws.Range("C" & r).HorizontalAlignment = xlLeft
    
    ws.Range("E" & r).value = "Transport Name:"
    ws.Range("E" & r).Font.Bold = True: ws.Range("E" & r).Font.Size = 11
    ws.Range("F" & r & ":G" & r).Merge
    ws.Range("F" & r).value = courierName
    ws.Range("F" & r).Font.Size = 11
    ws.Range("F" & r).HorizontalAlignment = xlLeft
    
    ' Line 2: Mobile (Left) + Docket Number (Right)
    r = 33: ws.Rows(r).RowHeight = 18
    ws.Range("B" & r).value = "Mobile:"
    ws.Range("B" & r).Font.Bold = True: ws.Range("B" & r).Font.Size = 11
    ws.Range("C" & r & ":D" & r).Merge
    ws.Range("C" & r).value = courierMobile
    ws.Range("C" & r).Font.Size = 11
    ws.Range("C" & r).HorizontalAlignment = xlLeft
    
    ws.Range("E" & r).value = "Docket Number:"
    ws.Range("E" & r).Font.Bold = True: ws.Range("E" & r).Font.Size = 11
    ws.Range("F" & r & ":G" & r).Merge
    ws.Range("F" & r).value = docketNo
    ws.Range("F" & r).Font.Size = 11
    ws.Range("F" & r).HorizontalAlignment = xlLeft
    
    ' Line 3: Send Date (Left) + Charges (Right)
    r = 34: ws.Rows(r).RowHeight = 18
    ws.Range("B" & r).value = "Send Date:"
    ws.Range("B" & r).Font.Bold = True: ws.Range("B" & r).Font.Size = 11
    ws.Range("C" & r & ":D" & r).Merge
    ws.Range("C" & r).value = sendDate
    ws.Range("C" & r).Font.Size = 11
    ws.Range("C" & r).HorizontalAlignment = xlLeft
    
    ws.Range("E" & r).value = "Charges:"
    ws.Range("E" & r).Font.Bold = True: ws.Range("E" & r).Font.Size = 11
    ws.Range("F" & r & ":G" & r).Merge
    ws.Range("F" & r).value = charges
    ws.Range("F" & r).Font.Size = 11
    ws.Range("F" & r).HorizontalAlignment = xlLeft
    
    '=== PENDING SERVICES SECTION ===
    Dim wsAssign As Worksheet
    Dim lastRow As Long, j As Long, dataRow As Long
    Dim pendingCount As Integer
    Dim daysCalc As Long, sendDt As Date
    Dim assignToName As String, statusVal As String
    Dim entryID_Pen As String, prodPen As String, compPen As String, modelPen As String, serialPen As String
    
    pendingCount = 0
    
    r = 36
    ws.Range("B" & r & ":G" & r).Merge
    ws.Range("B" & r).value = "PENDING SERVICES"
    ws.Range("B" & r).Font.Size = 12
    ws.Range("B" & r).Font.Bold = True
    ws.Range("B" & r).Interior.Color = RGB(128, 0, 0)
    ws.Range("B" & r).Font.Color = RGB(255, 255, 255)
    ws.Range("B" & r).HorizontalAlignment = xlCenter
    ws.Rows(r).RowHeight = 18
    
    r = 37
    ws.Rows(r).RowHeight = 16
    ws.Range("B" & r).value = "Entry ID": ws.Range("B" & r).Font.Bold = True: ws.Range("B" & r).Font.Size = 8
    ws.Range("C" & r).value = "Product": ws.Range("C" & r).Font.Bold = True: ws.Range("C" & r).Font.Size = 8
    ws.Range("D" & r).value = "Company": ws.Range("D" & r).Font.Bold = True: ws.Range("D" & r).Font.Size = 8
    ws.Range("E" & r).value = "Model": ws.Range("E" & r).Font.Bold = True: ws.Range("E" & r).Font.Size = 8
    ws.Range("F" & r).value = "Serial": ws.Range("F" & r).Font.Bold = True: ws.Range("F" & r).Font.Size = 8
    ws.Range("G" & r).value = "Send Date (Days)": ws.Range("G" & r).Font.Bold = True: ws.Range("G" & r).Font.Size = 8
    ws.Range("B" & r & ":G" & r).Interior.Color = RGB(255, 228, 225)
    
    On Error Resume Next
    Set wsAssign = ThisWorkbook.Sheets("Assign_Master")
    On Error GoTo ErrorHandler
    
    If Not wsAssign Is Nothing Then
        lastRow = wsAssign.Cells(wsAssign.Rows.count, "A").End(xlUp).row
        
        For j = 2 To lastRow
            assignToName = Trim(wsAssign.Cells(j, "N").value)
            statusVal = Trim(LCase(wsAssign.Cells(j, "AF").value))
            
            If statusVal = "pending" And Len(assignToName) > 0 Then
                Dim nameMatch As Boolean
                nameMatch = False
                
                If LCase(assignToName) = LCase(vendName) Then
                    nameMatch = True
                ElseIf InStr(1, LCase(vendName), LCase(assignToName)) > 0 Then
                    nameMatch = True
                ElseIf InStr(1, LCase(assignToName), LCase(vendName)) > 0 Then
                    nameMatch = True
                End If
                
                 If nameMatch Then
                    pendingCount = pendingCount + 1
                    dataRow = 37 + pendingCount  ' 37 + 1 = 38 (???? row), 37+2=39 (?????)
                    ws.Rows(dataRow).RowHeight = 16
                    
                    entryID_Pen = wsAssign.Cells(j, "B").value
                    prodPen = wsAssign.Cells(j, "F").value
                    compPen = wsAssign.Cells(j, "G").value
                    modelPen = wsAssign.Cells(j, "H").value
                    serialPen = wsAssign.Cells(j, "I").value
                    
                    ws.Range("B" & dataRow).value = entryID_Pen
                    ws.Range("B" & dataRow).Font.Size = 9
                    ws.Range("C" & dataRow).value = IIf(prodPen = "", "-", prodPen)
                    ws.Range("C" & dataRow).Font.Size = 9
                    ws.Range("D" & dataRow).value = IIf(compPen = "", "-", compPen)
                    ws.Range("D" & dataRow).Font.Size = 9
                    ws.Range("E" & dataRow).value = IIf(modelPen = "", "-", modelPen)
                    ws.Range("E" & dataRow).Font.Size = 9
                    ws.Range("F" & dataRow).value = IIf(serialPen = "", "-", serialPen)
                    ws.Range("F" & dataRow).Font.Size = 9
                    
                    On Error Resume Next
                    sendDt = CDate(wsAssign.Cells(j, "Z").value)
                    If Err.Number = 0 And sendDt > 0 Then
                        daysCalc = DateDiff("d", sendDt, Date)
                        ws.Range("G" & dataRow).value = Format(sendDt, "dd-mm-yyyy") & " (" & daysCalc & ")"
                    Else
                        ws.Range("G" & dataRow).value = "-"
                    End If
                    On Error GoTo ErrorHandler
                    ws.Range("G" & dataRow).Font.Size = 9
                    
                    ' dataRow = dataRow + 1  <-- ?? ???? ?? ???
                End If
            End If        ' <-- YEH ADD KARO (statusVal wale If ka End If)
        Next j
    End If
    
    If pendingCount = 0 Then
        dataRow = 38
        ws.Rows(dataRow).RowHeight = 16
        ws.Range("B" & dataRow & ":G" & dataRow).Merge
        ws.Range("B" & dataRow).value = "No Pending Services"
        ws.Range("B" & dataRow).HorizontalAlignment = xlCenter
        ws.Range("B" & dataRow).Font.Italic = True
        ws.Range("B" & dataRow).Font.Size = 11
        dataRow = 38
    Else
        dataRow = 38 + pendingCount
    End If
    
    '=== SIGNATURE ===
    r = dataRow + 1
    ws.Rows(r).RowHeight = 12
    
    r = r + 1
    ws.Range("E" & r & ":G" & r).Merge
    ws.Range("E" & r).value = "_________________________"
    ws.Range("E" & r).HorizontalAlignment = xlCenter
    
    r = r + 1
    ws.Range("E" & r & ":G" & r).Merge
    ws.Range("E" & r).value = "Authorized Signature"
    ws.Range("E" & r).Font.Bold = True
    ws.Range("E" & r).HorizontalAlignment = xlCenter
    
    r = r + 1
    ws.Range("E" & r & ":G" & r).Merge
    ws.Range("E" & r).value = "For " & compName
    ws.Range("E" & r).HorizontalAlignment = xlCenter
    ws.Range("E" & r).Font.Size = 12
    ws.Range("E" & r).Font.Italic = True
    
    ws.DisplayPageBreaks = False
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    ws.PrintOut Copies:=1
    MsgBox "Assignment Report Printed Successfully!", vbInformation, "Print Complete"
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "Error: " & Err.Number & " - " & Err.Description, vbCritical, "Print Error"
End Sub


'===========================================
' BUILD COMPANY ADDRESS HELPER FUNCTION
'===========================================
Private Function BuildCompanyAddress(ws As Worksheet) As String
    Dim addr As String
    Dim addr1 As String, addr2 As String, city As String
    Dim pin As String, state As String, country As String
    
    addr1 = Trim(ws.Range("B3").value & "")
    addr2 = Trim(ws.Range("B4").value & "")
    city = Trim(ws.Range("B5").value & "")
    pin = Trim(ws.Range("B6").value & "")
    state = Trim(ws.Range("B7").value & "")
    country = Trim(ws.Range("B8").value & "")
    
    addr = ""
    If addr1 <> "" Then addr = addr & addr1
    If addr2 <> "" Then addr = addr & ", " & addr2
    If city <> "" Then addr = addr & ", " & city
    If pin <> "" Then addr = addr & " - " & pin
    If state <> "" Then addr = addr & ", " & state
    If country <> "" Then addr = addr & ", " & country
    
    BuildCompanyAddress = addr
End Function

'===========================================
' EXPORT TO PDF BUTTON CLICK
'===========================================
Private Sub btnExportPDF_Click()
    If Not IsDataSaved Then
        MsgBox "Please save assignment first!", vbExclamation
        Exit Sub
    End If
    
    Call ExportAssignmentToPDF
    
    MsgBox "Export Complete!", vbInformation
    
    
End Sub

'===========================================
' EXPORT TO PDF - MAIN FUNCTION
'===========================================
Private Sub ExportAssignmentToPDF()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet, shp As Shape, r As Long, i As Integer, idx As Integer
    Dim compName$, addr1$, addr2$, city$, state$, pin$, email$, mobile$, logoPath$, logoFullPath$
    Dim entryID$, assignDate$, vendName$, vendMobile$, vendAddress$
    Dim probText$, remText$, courierMode$, courierName$, docketNo$
    Dim sendDate$, courierMobile$, charges$, accText$
    Dim prodPhotoPath$, logoFolder$, fileName$, csp As String, chk As Control
    Dim prodName$, compProd$, model$, serial$, warranty$, warrantyRem$
    Dim pdfPath As String, pdfFolder As String, assignID As String
    Dim wsSettings As Worksheet
    Dim fso As Object
    Dim wsConfig As Worksheet
    
    ' === GET ASSIGN ID FOR FILENAME ===
    If LastAssignID = "" Then
        assignID = "ASN" & Format(Now, "ddmmyyyyhhmmss")
    Else
        assignID = LastAssignID
    End If
    
    ' === PDF FOLDER PATH FROM SETTINGS ===
    On Error Resume Next
    Set wsSettings = ThisWorkbook.Sheets("Settings")
    ' B12 cell ??? PDF folder path ???? (????: D:\GLOBAL_SOFT_DATA\PDF\)
    pdfFolder = Trim(wsSettings.Range("B12").value)
    If pdfFolder = "" Then
        pdfFolder = ThisWorkbook.path & "\GLOBAL_SOFT_DATA\PDF\"
    End If
    On Error GoTo 0
    
    If Right(pdfFolder, 1) <> "\" Then pdfFolder = pdfFolder & "\"
    
    ' === CREATE FOLDER IF NOT EXISTS ===
    On Error Resume Next
    If Dir(pdfFolder, vbDirectory) = "" Then
        MkDir pdfFolder
    End If
    On Error GoTo 0
    
    pdfPath = pdfFolder & "Assignment_" & assignID & ".pdf"
    
    ' === DATA GATHERING (Same as Print) ===
    On Error Resume Next
    entryID = cmbEntryID.value
    assignDate = Format(Date, "dd-mmm-yyyy")
    vendName = lblVendorName.caption
    vendMobile = lblVendorMobile.caption
    vendAddress = lblVendorAddress.caption
    vendEmail = lblVendorEmail.caption
    prodName = lblProductValue.caption
    compProd = lblCompanyValue.caption
    model = lblModelValue.caption
    serial = lblSerialValue.caption
    warranty = lblWarrantyStatusValue.caption
    
    If Not Me.Controls("lblWarrantyRemaining") Is Nothing Then
        warrantyRem = Me.Controls("lblWarrantyRemaining").caption
    Else
        warrantyRem = "N/A"
    End If
    
    If Trim(prodName) = "" Then prodName = "Not Available"
    If Trim(compProd) = "" Then compProd = "Not Available"
    If Trim(model) = "" Then model = "Not Available"
    If Trim(serial) = "" Then serial = "Not Available"
    If Trim(warranty) = "" Then warranty = "Not Available"
    
    probText = txtProblemDescription.value
    If Trim(probText) = "" Then probText = "Not Specified"
    
    accText = ""
    idx = 1
    For Each chk In Me.Controls
        If TypeName(chk) = "CheckBox" Then
            If chk.value = True Then
                If accText <> "" Then accText = accText & ", "
                accText = accText & idx & ". " & chk.caption
                idx = idx + 1
            End If
        End If
    Next chk
    If accText = "" Then accText = "No Accessories"
    
    remText = txtRemarks.value
    courierMode = cmbCourierMode.value
    courierName = cmbCourierCompany.value
    docketNo = txtDocketNumber.value
    sendDate = txtSendDate.value
    charges = txtCourierCharges.value
    courierMobile = lblCourierMobile.caption
    
    On Error GoTo ErrorHandler
    
    ' === WORKSHEET SETUP ===
    On Error Resume Next
    ' ?????? temp sheet delete ??? ??? ??
    Application.DisplayAlerts = False
    Set ws = ThisWorkbook.Sheets("Assignment_PDF_Temp")
    If Not ws Is Nothing Then ws.Delete
    Application.DisplayAlerts = True
    
    Set ws = ThisWorkbook.Sheets.Add
    ws.name = "Assignment_PDF_Temp"
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    ws.Cells.ClearContents
    ws.Cells.Borders.LineStyle = xlNone
    ws.Cells.Interior.ColorIndex = xlNone
    
    On Error Resume Next
    For Each shp In ws.Shapes: shp.Delete: Next
    On Error GoTo ErrorHandler
    
    ' === PAGE SETUP FOR PDF ===
    With ws.PageSetup
        .PaperSize = xlPaperA4
        .Orientation = xlPortrait
        .LeftMargin = Application.InchesToPoints(0.25)
        .RightMargin = Application.InchesToPoints(0.25)
        .TopMargin = Application.InchesToPoints(0.25)
        .BottomMargin = Application.InchesToPoints(0.25)
        .HeaderMargin = Application.InchesToPoints(0.2)
        .FooterMargin = Application.InchesToPoints(0.2)
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .Zoom = 90
    End With
    
    ' Column Widths
    ws.Columns("A").ColumnWidth = 1
    ws.Columns("B").ColumnWidth = 16
    ws.Columns("C").ColumnWidth = 20
    ws.Columns("D").ColumnWidth = 15
    ws.Columns("E").ColumnWidth = 15
    ws.Columns("F").ColumnWidth = 13
    ws.Columns("G").ColumnWidth = 13
    ws.Columns("H").ColumnWidth = 1

    ' === COMPANY DETAILS ===
    With ThisWorkbook.Sheets("Software_Config")
        compName = .Range("B2").value
        addr1 = .Range("B3").value: addr2 = .Range("B4").value
        city = .Range("B5").value: state = .Range("B7").value: pin = .Range("B6").value
        mobile = .Range("B9").value: email = .Range("B10").value: logoPath = .Range("B11").value
    End With
    
    '=== HEADER ===
    r = 1
    ws.Range("B" & r & ":E" & r).Merge
    ws.Range("B" & r).value = compName
    ws.Range("B" & r).Font.Size = 18
    ws.Range("B" & r).Font.Bold = True
    ws.Range("B" & r).Font.name = "Arial"
    ws.Rows(r).RowHeight = 28
    
    r = 2: ws.Range("B" & r & ":E" & r).Merge: ws.Range("B" & r).value = addr1: ws.Rows(r).RowHeight = 16
    
    If Trim(addr2) <> "" Then
        r = 3: ws.Range("B" & r & ":E" & r).Merge: ws.Range("B" & r).value = addr2: ws.Rows(r).RowHeight = 16
        r = 4
    Else
        r = 3
    End If
    
    csp = city
    If Trim(state) <> "" Then csp = csp & ", " & state
    If Trim(pin) <> "" Then csp = csp & " - " & pin
    
    ws.Range("B" & r & ":E" & r).Merge: ws.Range("B" & r).value = csp
    ws.Range("B" & r).Font.Size = 10: ws.Rows(r).RowHeight = 16
    
    r = r + 1: ws.Range("B" & r & ":E" & r).Merge: ws.Range("B" & r).value = "Email: " & email: ws.Range("B" & r).Font.Size = 9: ws.Rows(r).RowHeight = 14
    r = r + 1: ws.Range("B" & r & ":E" & r).Merge: ws.Range("B" & r).value = "Mobile: " & mobile: ws.Range("B" & r).Font.Size = 9: ws.Rows(r).RowHeight = 14
    
    '=== LOGO ===
    logoFullPath = ""
    If Trim(logoPath) <> "" Then
        If Dir(logoPath) <> "" Then logoFullPath = logoPath
    End If
    
    If logoFullPath = "" Then
        logoFolder = ThisWorkbook.path & "\GLOBAL_SOFT_DATA\Images\Logo\"
        If Dir(logoFolder, vbDirectory) = "" Then logoFolder = ThisWorkbook.path & "\LOGO\"
        fileName = Dir(logoFolder & "*.png")
        If fileName = "" Then fileName = Dir(logoFolder & "*.jpg")
        If fileName <> "" Then logoFullPath = logoFolder & fileName
    End If
    
    If logoFullPath <> "" Then
        On Error Resume Next
                ws.Shapes.AddPicture logoFullPath, msoFalse, msoTrue, ws.Range("F2").Left + 5, ws.Range("B1").Top + 4, 120, 80
        On Error GoTo ErrorHandler
    End If
    
    '=== TITLE ===
    r = 7
    ws.Range("B" & r & ":G" & r).Merge
    ws.Range("B" & r).value = "ASSIGNMENT REPORT"
    ws.Range("B" & r).Font.Size = 18
    ws.Range("B" & r).Font.Bold = True
    ws.Range("B" & r).HorizontalAlignment = xlCenter
    ws.Range("B" & r).Interior.Color = RGB(0, 102, 204)
    ws.Range("B" & r).Font.Color = RGB(255, 255, 255)
    ws.Rows(r).RowHeight = 22
    
    '=== ID & DATE ===
    r = 8: ws.Rows(r).RowHeight = 18
    ws.Range("B" & r & ":D" & r).Merge
    ws.Range("B" & r).value = "Assignment ID: " & entryID
    ws.Range("B" & r).Font.Bold = True: ws.Range("B" & r).Font.Size = 12
    ws.Range("E" & r & ":G" & r).Merge
    ws.Range("E" & r).value = "Date: " & assignDate
    ws.Range("E" & r).Font.Bold = True: ws.Range("E" & r).Font.Size = 12
       ws.Range("E" & r).HorizontalAlignment = xlLeft
    
    '=== ASSIGN TO DETAILS (3 Lines ???? ???? ???) ===
    r = 10
    ws.Range("B" & r & ":G" & r).Merge
    ws.Range("B" & r).value = "ASSIGN TO DETAILS"
    ws.Range("B" & r).Font.Size = 12
    ws.Range("B" & r).Font.Bold = True
    ws.Range("B" & r).Interior.Color = RGB(0, 0, 128)
    ws.Range("B" & r).Font.Color = RGB(255, 255, 255)
    ws.Rows(r).RowHeight = 18
    
        '=== ASSIGN TO DETAILS ===
    r = 10
    ws.Range("B" & r & ":G" & r).Merge
    ws.Range("B" & r).value = "ASSIGN TO DETAILS"
    ws.Range("B" & r).Font.Size = 12
    ws.Range("B" & r).Font.Bold = True
    ws.Range("B" & r).Interior.Color = RGB(0, 0, 128)
    ws.Range("B" & r).Font.Color = RGB(255, 255, 255)
    ws.Rows(r).RowHeight = 20
    
    r = 11: ws.Rows(r).RowHeight = 22
    ws.Range("B" & r).value = "Name:"
    ws.Range("B" & r).Font.Bold = True
    ws.Range("B" & r).Font.Size = 11
    ws.Range("C" & r & ":G" & r).Merge
    ws.Range("C" & r).value = vendName
    ws.Range("C" & r).Font.Size = 11
    ws.Range("C" & r).HorizontalAlignment = xlLeft
    
    r = 12: ws.Rows(r).RowHeight = 22
    ws.Range("B" & r).value = "Mobile:"
    ws.Range("B" & r).Font.Bold = True
    ws.Range("B" & r).Font.Size = 11
    ws.Range("C" & r & ":D" & r).Merge
    ws.Range("C" & r).value = vendMobile
    ws.Range("C" & r).Font.Size = 11
    ws.Range("C" & r).HorizontalAlignment = xlLeft
    
    r = 13: ws.Rows(r).RowHeight = 22
    ws.Range("B" & r).value = "Address:"
    ws.Range("B" & r).Font.Bold = True
    ws.Range("B" & r).Font.Size = 11
    ws.Range("C" & r & ":G" & r).Merge
    ws.Range("C" & r).value = vendAddress
    ws.Range("C" & r).Font.Size = 11
    ws.Range("C" & r).HorizontalAlignment = xlLeft
    ws.Range("C" & r).WrapText = True  ' ??? address ???? ?? ??
    
    ' === PRODUCT DETAILS ===
    On Error Resume Next
    ws.Range("B15:G15").Merge
    ws.Range("B15").value = "PRODUCT DETAILS (SERVICE ID: " & entryID & ")"
    ws.Range("B15").Font.Size = 12
    ws.Range("B15").Font.Bold = True
    ws.Range("B15").Interior.Color = RGB(255, 20, 147)
    ws.Range("B15").Font.Color = RGB(255, 255, 255)
    ws.Rows(15).RowHeight = 18
    
    ws.Range("B16").value = "Product:": ws.Range("B16").Font.Bold = True: ws.Range("B16").Font.Size = 11
    ws.Range("C16:E16").Merge: ws.Range("C16").value = lblProductValue.caption: ws.Range("C16").Font.Size = 11
    ws.Rows(16).RowHeight = 16
    
    ws.Range("B17").value = "Company:": ws.Range("B17").Font.Bold = True: ws.Range("B17").Font.Size = 11
    ws.Range("C17:E17").Merge: ws.Range("C17").value = lblCompanyValue.caption: ws.Range("C17").Font.Size = 11
    ws.Rows(17).RowHeight = 16
    
    ws.Range("B18").value = "Model:": ws.Range("B18").Font.Bold = True: ws.Range("B18").Font.Size = 11
    ws.Range("C18").value = lblModelValue.caption: ws.Range("C18").Font.Size = 11
    ws.Rows(18).RowHeight = 16
    
    ws.Range("B19").value = "Serial:": ws.Range("B19").Font.Bold = True: ws.Range("B19").Font.Size = 11
    ws.Range("C19").value = lblSerialValue.caption: ws.Range("C19").Font.Size = 11
    ws.Rows(19).RowHeight = 16
    
    ws.Range("B20").value = "Warranty:": ws.Range("B20").Font.Bold = True: ws.Range("B20").Font.Size = 11
    ws.Range("C20:F20").Merge: ws.Range("C20").value = lblWarrantyStatusValue.caption & IIf(warrantyRem <> "N/A", " (" & warrantyRem & ")", ""): ws.Range("C20").Font.Size = 11
    ws.Rows(20).RowHeight = 16
    
    '=== PRODUCT PHOTO (???? ???? - 160x120) ===
    Dim photoFile As String
    photoFile = ThisWorkbook.path & "\GLOBAL_SOFT_DATA\Images\Products\" & entryID & "\Photo1.jpg"
    If Dir(photoFile) = "" Then photoFile = ThisWorkbook.path & "\GLOBAL_SOFT_DATA\Images\Products\" & entryID & ".jpg"
    If Dir(photoFile) <> "" Then
        On Error Resume Next
        ws.Shapes.AddPicture photoFile, msoFalse, msoTrue, ws.Range("F16").Left, ws.Range("F16").Top, 120, 80
        On Error GoTo ErrorHandler
    End If
    On Error GoTo ErrorHandler
    
    '=== ACCESSORIES (????) ===
    r = 22
    ws.Range("B" & r & ":G" & r).Merge
    ws.Range("B" & r).value = "SEND ACCESSORIES"
    ws.Range("B" & r).Font.Size = 13
    ws.Range("B" & r).Font.Bold = True
    ws.Range("B" & r).Interior.Color = RGB(0, 153, 76)
    ws.Range("B" & r).Font.Color = RGB(255, 255, 255)
    ws.Rows(r).RowHeight = 20
    
    r = 23
    ws.Rows(r).RowHeight = 20
    ws.Range("B" & r & ":G" & r).Merge
    ws.Range("B" & r).value = accText
    ws.Range("B" & r).Font.Size = 12
    ws.Range("B" & r).WrapText = True
    
    '=== PROBLEM (????) ===
    r = 25
    ws.Range("B" & r & ":G" & r).Merge
    ws.Range("B" & r).value = "PROBLEM"
    ws.Range("B" & r).Font.Size = 13
    ws.Range("B" & r).Font.Bold = True
    ws.Range("B" & r).Interior.Color = RGB(204, 102, 0)
    ws.Range("B" & r).Font.Color = RGB(255, 255, 255)
    ws.Rows(r).RowHeight = 20
    
    r = 26
    ws.Rows(r).RowHeight = 24
    ws.Range("B" & r & ":G" & r).Merge
    ws.Range("B" & r).value = probText
    ws.Range("B" & r).Font.Size = 11
    ws.Range("B" & r).WrapText = True
    
    '=== REMARKS (????) ===
    r = 28
    ws.Range("B" & r & ":G" & r).Merge
    ws.Range("B" & r).value = "REMARKS"
    ws.Range("B" & r).Font.Size = 13
    ws.Range("B" & r).Font.Bold = True
    ws.Range("B" & r).Interior.Color = RGB(178, 134, 0)
    ws.Range("B" & r).Font.Color = RGB(255, 255, 255)
    ws.Rows(r).RowHeight = 20
    
    r = 29
    ws.Rows(r).RowHeight = 24
    ws.Range("B" & r & ":G" & r).Merge
    ws.Range("B" & r).value = remText
    ws.Range("B" & r).Font.Size = 11
    ws.Range("B" & r).WrapText = True
    
        '=== TRANSPORT DETAILS (3 ????) ===
    r = 31
    ws.Range("B" & r & ":G" & r).Merge
    ws.Range("B" & r).value = "TRANSPORT DETAILS"
    ws.Range("B" & r).Font.Size = 12
    ws.Range("B" & r).Font.Bold = True
    ws.Range("B" & r).Interior.Color = RGB(0, 153, 153)
    ws.Range("B" & r).Font.Color = RGB(255, 255, 255)
    ws.Rows(r).RowHeight = 18
    
    ' Line 1: Mode + Name (?? Left Align)
    r = 32: ws.Rows(r).RowHeight = 18
    ws.Range("B" & r).value = "Transport Mode:"
    ws.Range("B" & r).Font.Bold = True: ws.Range("B" & r).Font.Size = 11
    ws.Range("C" & r & ":D" & r).Merge
    ws.Range("C" & r).value = courierMode
    ws.Range("C" & r).Font.Size = 11
    ws.Range("C" & r).HorizontalAlignment = xlLeft
    
    ws.Range("E" & r).value = "Transport Name:"
    ws.Range("E" & r).Font.Bold = True: ws.Range("E" & r).Font.Size = 11
    ws.Range("E" & r).HorizontalAlignment = xlLeft
    ws.Range("F" & r & ":G" & r).Merge
    ws.Range("F" & r).value = courierName
    ws.Range("F" & r).Font.Size = 11
    ws.Range("F" & r).HorizontalAlignment = xlLeft
    
    ' Line 2: Mobile + Docket (?? Left Align)
    r = 33: ws.Rows(r).RowHeight = 18
    ws.Range("B" & r).value = "Mobile:"
    ws.Range("B" & r).Font.Bold = True: ws.Range("B" & r).Font.Size = 11
    ws.Range("C" & r & ":D" & r).Merge
    ws.Range("C" & r).value = courierMobile
    ws.Range("C" & r).Font.Size = 11
    ws.Range("C" & r).HorizontalAlignment = xlLeft
    
    ws.Range("E" & r).value = "Docket Number:"
    ws.Range("E" & r).Font.Bold = True: ws.Range("E" & r).Font.Size = 11
    ws.Range("E" & r).HorizontalAlignment = xlLeft
    ws.Range("F" & r & ":G" & r).Merge
    ws.Range("F" & r).value = docketNo
    ws.Range("F" & r).Font.Size = 11
    ws.Range("F" & r).HorizontalAlignment = xlLeft
    
    ' Line 3: Date + Charges (?? Left Align)
    r = 34: ws.Rows(r).RowHeight = 18
    ws.Range("B" & r).value = "Send Date:"
    ws.Range("B" & r).Font.Bold = True: ws.Range("B" & r).Font.Size = 11
    ws.Range("C" & r & ":D" & r).Merge
    ws.Range("C" & r).value = sendDate
    ws.Range("C" & r).Font.Size = 11
    ws.Range("C" & r).HorizontalAlignment = xlLeft
    
    ws.Range("E" & r).value = "Charges:"
    ws.Range("E" & r).Font.Bold = True: ws.Range("E" & r).Font.Size = 11
    ws.Range("E" & r).HorizontalAlignment = xlLeft
    ws.Range("F" & r & ":G" & r).Merge
    ws.Range("F" & r).value = charges
    ws.Range("F" & r).Font.Size = 11
    ws.Range("F" & r).HorizontalAlignment = xlLeft
    
    '=== PENDING SERVICES ===
    Dim wsAssign As Worksheet
    Dim lastRow As Long, j As Long, dataRow As Long
    Dim pendingCount As Integer
    Dim daysCalc As Long, sendDt As Date
    Dim assignToName As String, statusVal As String
    Dim entryID_Pen As String, prodPen As String, compPen As String, modelPen As String, serialPen As String
    
    pendingCount = 0
    
        '=== PENDING SERVICES ===
    r = 36
    ws.Range("B" & r & ":G" & r).Merge
    ws.Range("B" & r).value = "PENDING SERVICES"
    ws.Range("B" & r).Font.Size = 12
    ws.Range("B" & r).Font.Bold = True
    ws.Range("B" & r).Interior.Color = RGB(128, 0, 0)
    ws.Range("B" & r).Font.Color = RGB(255, 255, 255)
    ws.Range("B" & r).HorizontalAlignment = xlCenter
    ws.Rows(r).RowHeight = 20
    
    ' Table Header
    r = 37
    ws.Rows(r).RowHeight = 18
    ws.Range("B" & r).value = "Entry ID": ws.Range("B" & r).Font.Bold = True: ws.Range("B" & r).Font.Size = 9
    ws.Range("C" & r).value = "Product": ws.Range("C" & r).Font.Bold = True: ws.Range("C" & r).Font.Size = 9
    ws.Range("D" & r).value = "Company": ws.Range("D" & r).Font.Bold = True: ws.Range("D" & r).Font.Size = 9
    ws.Range("E" & r).value = "Model": ws.Range("E" & r).Font.Bold = True: ws.Range("E" & r).Font.Size = 9
    ws.Range("F" & r).value = "Serial": ws.Range("F" & r).Font.Bold = True: ws.Range("F" & r).Font.Size = 9
    ws.Range("G" & r).value = "Send Date (Days)": ws.Range("G" & r).Font.Bold = True: ws.Range("G" & r).Font.Size = 9
    ws.Range("B" & r & ":G" & r).Interior.Color = RGB(255, 228, 225)
    
        On Error Resume Next
    Set wsAssign = ThisWorkbook.Sheets("Assign_Master")
    On Error GoTo ErrorHandler
    
    If Not wsAssign Is Nothing Then
        lastRow = wsAssign.Cells(wsAssign.Rows.count, "A").End(xlUp).row
        dataRow = 38
        
        For j = 2 To lastRow
            assignToName = Trim(wsAssign.Cells(j, "N").value)
            statusVal = Trim(LCase(wsAssign.Cells(j, "AF").value))
            
            If statusVal = "pending" And Len(assignToName) > 0 Then
                Dim nameMatch As Boolean
                nameMatch = False
                
                If LCase(assignToName) = LCase(vendName) Then
                    nameMatch = True
                ElseIf InStr(1, LCase(vendName), LCase(assignToName)) > 0 Then
                    nameMatch = True
                ElseIf InStr(1, LCase(assignToName), LCase(vendName)) > 0 Then
                    nameMatch = True
                End If
                
                If nameMatch Then
                    
                    
                    dataRow = 38 + pendingCount
                    ws.Rows(dataRow).RowHeight = 16
                    
                    entryID_Pen = wsAssign.Cells(j, "B").value
                    prodPen = wsAssign.Cells(j, "F").value
                    compPen = wsAssign.Cells(j, "G").value
                    modelPen = wsAssign.Cells(j, "H").value
                    serialPen = wsAssign.Cells(j, "I").value
                    
                    ws.Range("B" & dataRow).value = entryID_Pen: ws.Range("B" & dataRow).Font.Size = 9
                    ws.Range("C" & dataRow).value = IIf(prodPen = "", "-", prodPen): ws.Range("C" & dataRow).Font.Size = 9
                    ws.Range("D" & dataRow).value = IIf(compPen = "", "-", compPen): ws.Range("D" & dataRow).Font.Size = 9
                    ws.Range("E" & dataRow).value = IIf(modelPen = "", "-", modelPen): ws.Range("E" & dataRow).Font.Size = 9
                    ws.Range("F" & dataRow).value = IIf(serialPen = "", "-", serialPen): ws.Range("F" & dataRow).Font.Size = 9
                    
                    On Error Resume Next
                    sendDt = CDate(wsAssign.Cells(j, "Z").value)
                    If Err.Number = 0 And sendDt > 0 Then
                        daysCalc = DateDiff("d", sendDt, Date)
                        ws.Range("G" & dataRow).value = Format(sendDt, "dd-mm-yyyy") & " (" & daysCalc & ")"
                    Else
                        ws.Range("G" & dataRow).value = "-"
                    End If
                    On Error GoTo ErrorHandler
                    ws.Range("G" & dataRow).Font.Size = 9
                    
                    dataRow = dataRow + 1
                    pendingCount = pendingCount + 1
                End If
            End If
        Next j
    End If
    
    
    ' If no pending
    If pendingCount = 0 Then
        dataRow = 38
        ws.Rows(dataRow).RowHeight = 18
        ws.Range("B" & dataRow & ":G" & dataRow).Merge
        ws.Range("B" & dataRow).value = "No Pending Services"
        ws.Range("B" & dataRow).HorizontalAlignment = xlCenter
        ws.Range("B" & dataRow).Font.Italic = True
        dataRow = 39
    End If
    
    '=== SIGNATURE ===
       '=== SIGNATURE ===
    r = dataRow + 2  ' 1 row gap
    ws.Rows(r).RowHeight = 30  ' ???? height for signature space
    
    r = r + 1
    ws.Range("E" & r & ":G" & r).Merge
    ws.Range("E" & r).value = "_________________________"
    ws.Range("E" & r).HorizontalAlignment = xlCenter
    ws.Range("E" & r).Font.Size = 10
    
    r = r + 1
    ws.Rows(r).RowHeight = 18
    ws.Range("E" & r & ":G" & r).Merge
    ws.Range("E" & r).value = "Authorized Signature"
    ws.Range("E" & r).Font.Bold = True
    ws.Range("E" & r).Font.Size = 10
    ws.Range("E" & r).HorizontalAlignment = xlCenter
    
    r = r + 1
    ws.Rows(r).RowHeight = 16
    ws.Range("E" & r & ":G" & r).Merge
    ws.Range("E" & r).value = "For " & compName
    ws.Range("E" & r).HorizontalAlignment = xlCenter
    ws.Range("E" & r).Font.Size = 9
    ws.Range("E" & r).Font.Italic = True
    
    ' === EXPORT TO PDF ===
    ws.DisplayPageBreaks = False
    Application.ScreenUpdating = True
    
    ' Export as PDF
    ws.ExportAsFixedFormat Type:=xlTypePDF, fileName:=pdfPath, Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=True
    
    ' === CLEANUP ===
    Application.DisplayAlerts = False
    ws.Delete
    Application.DisplayAlerts = True
    
    MsgBox "PDF Exported Successfully!" & vbCrLf & "Saved to: " & pdfPath, vbInformation, "Export Complete"
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "Export Error: " & Err.Number & " - " & Err.Description, vbCritical, "Error"
End Sub

Private Sub btnSendEmail_Click()
    On Error GoTo ErrorHandler
    
     ' ===== ???? VARIABLES ???? DECLARE ??? =====
    Dim wsConfig As Worksheet
    Dim senderEmail As String, emailPassword As String, appPassword As String
    Dim smtpServer As String, smtpPort As String, sslEnable As String
    Dim usePassword As String, assignToEmail As String
    Dim cdoMsg As Object, cdoConf As Object
    Dim compName As String, htmlBody As String
    Dim wsVendor As Worksheet
    Dim mode As String, vendName As String
    Dim lastRow As Long, i As Long
    Dim emailFound As Boolean
    Dim sendDt As Date, daysCalc As Long  ' ??? ?? ?? use ?? ??? ??
    
   ' ===== ?? ???? ?? actual code ???? =====
    If Not IsDataSaved Then
        MsgBox "Please save assignment first!", vbExclamation
        Exit Sub
    End If
    
    ' ===== GET EMAIL SETTINGS FROM Software_Config =====
    Set wsConfig = ThisWorkbook.Sheets("Software_Config")
    
    senderEmail = Trim(wsConfig.Range("B17").value & "")
    emailPassword = Trim(wsConfig.Range("B18").value & "")
    smtpServer = Trim(wsConfig.Range("B19").value & "")
    appPassword = Trim(wsConfig.Range("B20").value & "")
    smtpPort = Trim(wsConfig.Range("B21").value & "")
    sslEnable = UCase(Trim(wsConfig.Range("B22").value & ""))
    
    If Not IsDataSaved Then
        MsgBox "Please save assignment first!", vbExclamation
        Exit Sub
    End If
    
    ' ===== GET EMAIL SETTINGS FROM Software_Config =====
    
    
    
    Set wsConfig = ThisWorkbook.Sheets("Software_Config")
    
    senderEmail = Trim(wsConfig.Range("B17").value & "")
    emailPassword = Trim(wsConfig.Range("B18").value & "")
    smtpServer = Trim(wsConfig.Range("B19").value & "")
    appPassword = Trim(wsConfig.Range("B20").value & "")
    smtpPort = Trim(wsConfig.Range("B21").value & "")
    sslEnable = UCase(Trim(wsConfig.Range("B22").value & ""))
    
    If senderEmail = "" Or smtpServer = "" Then
        MsgBox "Email settings not configured!", vbExclamation
        Exit Sub
    End If
    
    
    If appPassword <> "" Then
        usePassword = appPassword
    Else
        usePassword = emailPassword
    End If
    
    ' ===== GET ASSIGN TO EMAIL (Direct from Register) =====




mode = UCase(Trim(cmbAssignTo.value))
vendName = Trim(cmbAssignName.value)
assignToEmail = ""

' Select correct register sheet
Select Case mode
    Case "VENDOR": Set wsVendor = ThisWorkbook.Sheets("Vendor_Register")
    Case "ENGINEER": Set wsVendor = ThisWorkbook.Sheets("Engineer_Register")
    Case "INHOUSE": Set wsVendor = ThisWorkbook.Sheets("Inhouse_Register")
    Case "SERVICE_STATION": Set wsVendor = ThisWorkbook.Sheets("ServiceStation_Register")
    Case Else: Set wsVendor = Nothing
End Select

' Find email in Column G (7th column)
If Not wsVendor Is Nothing Then
    lastRow = wsVendor.Cells(wsVendor.Rows.count, 1).End(xlUp).row
    For i = 2 To lastRow
        If UCase(Trim(wsVendor.Cells(i, 2).value)) = UCase(vendName) Then
            assignToEmail = Trim(wsVendor.Cells(i, 7).value & "")  ' Column G = Email
            Exit For
        End If
    Next i
End If
    
    ' ===== SMTP Configuration (Reference Code ????) =====
    
    
    
    Set cdoMsg = CreateObject("CDO.Message")
    Set cdoConf = CreateObject("CDO.Configuration")
    
    With cdoConf.Fields
        .item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
        .item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = GetSMTPServerAddress(smtpServer)
        .item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = CInt(IIf(smtpPort = "", "587", smtpPort))
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
    
    ' ===== EMAIL BODY (Assignment Details) =====
    
    htmlBody = GenerateAssignmentEmailHTML()  ' ???? ???? ???
    
    ' ===== SEND EMAIL =====
    
    compName = Trim(wsConfig.Range("B2").value & "")
    If compName = "" Then compName = "GLOBAL IT SOLUTIONS"
    
    With cdoMsg
        Set .Configuration = cdoConf
        .From = senderEmail
        .To = assignToEmail
        .subject = "Service Assignment - " & cmbEntryID.value & " - " & lblProductValue.caption
        .htmlBody = htmlBody
        .Send
    End With
    
    MsgBox "Email sent successfully!" & vbCrLf & "To: " & assignToEmail, vbInformation
    
    Exit Sub

ErrorHandler:
    MsgBox "Email Error: " & Err.Description, vbCritical
End Sub
Private Function GetSMTPServerAddress(ByVal serverName As String) As String
    Select Case Trim(LCase(serverName))
        Case "gmail": GetSMTPServerAddress = "smtp.gmail.com"
        Case "outlook": GetSMTPServerAddress = "smtp.office365.com"
        Case "yahoo": GetSMTPServerAddress = "smtp.mail.yahoo.com"
        Case "hotmail": GetSMTPServerAddress = "smtp.live.com"
        Case Else: GetSMTPServerAddress = serverName
    End Select
End Function
Private Function GenerateAssignmentEmailHTML() As String
    On Error GoTo ErrorHandler
    
    Dim html As String
    Dim wsConfig As Worksheet, wsProd As Worksheet, wsAcc As Worksheet, wsAssign As Worksheet
    Dim entryID As String, assignID As String, custID As String
    Dim compName As String, addr1 As String, addr2 As String, city As String, state As String, pin As String
    Dim compMob As String, compEmail As String
    Dim custName As String, custMobile As String, custAddress As String
    Dim prodName As String, compProd As String, model As String, serial As String, warranty As String, warrantyRem As String
    Dim prob As String, remText As String
    Dim courierMode As String, courierName As String, docketNo As String, sendDate As String, charges As String, courierMobile As String
    Dim vendName As String, vendMobile As String, vendAddress As String
    Dim lastRow As Long, i As Long, j As Long, pendingCount As Integer
    Dim assignDate As String, css As String
    
    Set wsConfig = ThisWorkbook.Sheets("Software_Config")
    Set wsProd = ThisWorkbook.Sheets("Job_Product")
    Set wsAcc = ThisWorkbook.Sheets("Job_Accessory")
    Set wsAssign = ThisWorkbook.Sheets("Assign_Master")
    
    compName = Trim(wsConfig.Range("B2").value & "")
    addr1 = Trim(wsConfig.Range("B3").value & "")
    addr2 = Trim(wsConfig.Range("B4").value & "")
    city = Trim(wsConfig.Range("B5").value & "")
    pin = Trim(wsConfig.Range("B6").value & "")
    state = Trim(wsConfig.Range("B7").value & "")
    compMob = Trim(wsConfig.Range("B9").value & "")
    compEmail = Trim(wsConfig.Range("B10").value & "")
    
    If compName = "" Then compName = "GLOBAL IT SOLUTIONS"
    
    Dim compAddr As String
    compAddr = addr1
    If addr2 <> "" Then compAddr = compAddr & ", " & addr2
    If city <> "" Then compAddr = compAddr & ", " & city
    If pin <> "" Then compAddr = compAddr & " - " & pin
    If state <> "" Then compAddr = compAddr & ", " & state
    
    entryID = cmbEntryID.value
    assignID = LastAssignID
    assignDate = Format(Date, "dd-mmm-yyyy")
    prodName = lblProductValue.caption
    compProd = lblCompanyValue.caption
    model = lblModelValue.caption
    serial = lblSerialValue.caption
    warranty = lblWarrantyStatusValue.caption
    
    On Error Resume Next
    warrantyRem = Me.Controls("lblWarrantyRemaining").caption
    On Error GoTo 0
    If warrantyRem = "" Then warrantyRem = "N/A"
    
    custName = lblCustomerValue.caption
    custMobile = lblMobileValue.caption
    prob = txtProblemDescription.value
    remText = txtRemarks.value
    courierMode = cmbCourierMode.value
    courierName = cmbCourierCompany.value
    docketNo = txtDocketNumber.value
    sendDate = txtSendDate.value
    charges = txtCourierCharges.value
    courierMobile = lblCourierMobile.caption
    vendName = lblVendorName.caption
    vendMobile = lblVendorMobile.caption
    vendAddress = lblVendorAddress.caption
    vendEmail = lblVendorEmail.caption
    
    custID = ""
    For i = 2 To wsProd.Cells(wsProd.Rows.count, 1).End(xlUp).row
        If UCase(Trim(wsProd.Cells(i, 1).value & "")) = UCase(entryID) Then
            custID = Trim(wsProd.Cells(i, 2).value & "")
            Exit For
        End If
    Next i
    
    custAddress = ""
    If custID <> "" Then
        Dim wsCust As Worksheet
        Set wsCust = ThisWorkbook.Sheets("Customer_Master")
        For i = 2 To wsCust.Cells(wsCust.Rows.count, 1).End(xlUp).row
            If UCase(Trim(wsCust.Cells(i, 1).value & "")) = UCase(custID) Then
                custAddress = Trim(wsCust.Cells(i, 4).value & "")
                Exit For
            End If
        Next i
    End If
    
    css = "<style>body{font-family:Arial;margin:0;padding:0;background:#fff;}"
    css = css & ".container{width:100%;max-width:800px;margin:0 auto;border:1px solid #000;}"
    css = css & ".header{background:#fff;padding:10px;text-align:center;border-bottom:2px solid #000;}"
    css = css & ".header h1{margin:0;font-size:24px;font-weight:bold;color:#000;}"
    css = css & ".header p{margin:3px 0;font-size:12px;color:#000;}"
    css = css & ".title-bar{background:#0066CC;color:white;text-align:center;padding:8px;font-size:18px;font-weight:bold;}"
    css = css & ".id-date{padding:8px;font-size:12px;font-weight:bold;overflow:hidden;}"
    css = css & ".id-left{float:left;}.date-right{float:right;}"
    css = css & ".section-header{color:white;padding:6px;font-weight:bold;font-size:13px;margin-top:5px;}"
    css = css & ".assign-header{background:#000080;}.product-header{background:#FF1493;}"
    css = css & ".accessories-header{background:#009933;}.problem-header{background:#CC6600;}"
    css = css & ".remarks-header{background:#B8860B;}.transport-header{background:#008080;}"
    css = css & ".pending-header{background:#800000;text-align:center;}"
    css = css & ".content{padding:8px;font-size:12px;}"
    css = css & "table{width:100%;border-collapse:collapse;}"
    css = css & ".data-table td{border:1px solid #666;padding:5px;vertical-align:top;}"
    css = css & ".label{font-weight:bold;width:20%;background:#f0f0f0;}"
    css = css & ".value{width:30%;}.full-width{width:80%;}"
    css = css & ".acc-table th{background:#E8E8E8;border:1px solid #666;padding:5px;font-size:11px;font-weight:bold;}"
    css = css & ".acc-table td{border:1px solid #666;padding:5px;font-size:11px;text-align:center;}"
    css = css & ".pending-table th{background:#FFE4E1;border:1px solid #800000;padding:4px;font-size:10px;font-weight:bold;color:#800000;}"
    css = css & ".pending-table td{border:1px solid #800000;padding:4px;font-size:10px;text-align:center;}"
    css = css & ".signature{margin-top:30px;text-align:right;padding-right:50px;}"
    css = css & ".sig-line{border-top:1px solid #000;width:200px;display:inline-block;margin-bottom:5px;}"
    css = css & ".footer{background:#FFFFCD;color:#006400;text-align:center;padding:8px;font-weight:bold;border-top:2px solid #006400;margin-top:10px;font-size:12px;}"
    css = css & "</style>"
    
    html = "<!DOCTYPE html><html><head>" & css & "</head><body>"
    html = html & "<div class='container'>"
    html = html & "<div class='header'><h1>" & compName & "</h1>"
    If compAddr <> "" Then html = html & "<p>" & compAddr & "</p>"
    If compEmail <> "" Then html = html & "<p>Email: " & compEmail & "</p>"
    If compMob <> "" Then html = html & "<p>Mobile: " & compMob & "</p>"
    html = html & "</div>"
    
    html = html & "<div class='title-bar'>ASSIGNMENT REPORT</div>"
    html = html & "<div class='id-date'><span class='id-left'>Assignment ID: " & entryID & "</span><span class='date-right'>Date: " & assignDate & "</span></div><div style='clear:both;'></div>"
    
    html = html & "<div class='section-header assign-header'>ASSIGN TO DETAILS</div>"
    html = html & "<div class='content'><table class='data-table'>"
    html = html & "<tr><td class='label'>Mobile:</td><td class='value'>" & vendMobile & "</td><td class='label'>Mail ID:</td><td class='value'>" & vendEmail & "</td></tr>"
    html = html & "<tr><td class='label'>Address:</td><td colspan='3'>" & vendAddress & "</td></tr>"
    html = html & "</table></div>"
    
    html = html & "<div class='section-header product-header'>PRODUCT DETAILS (SERVICE ID: " & entryID & ")</div>"
    html = html & "<div class='content'><table class='data-table'>"
    html = html & "<tr><td class='label'>Product:</td><td class='value'>" & prodName & "</td><td class='label'>Company:</td><td class='value'>" & compProd & "</td></tr>"
    html = html & "<tr><td class='label'>Model:</td><td class='value'>" & model & "</td><td class='label'>Serial:</td><td class='value'>" & serial & "</td></tr>"
    html = html & "<tr><td class='label'>Warranty:</td><td colspan='3'>" & warranty & IIf(warrantyRem <> "N/A", " (" & warrantyRem & ")", "") & "</td></tr>"
    html = html & "</table></div>"
    
    html = html & "<div class='section-header accessories-header'>SEND ACCESSORIES</div>"
    html = html & "<div class='content'>"
    
    Dim accCount As Integer, lastRowAcc As Long
    accCount = 0
    lastRowAcc = wsAcc.Cells(wsAcc.Rows.count, 1).End(xlUp).row
    
    For i = 2 To lastRowAcc
        If UCase(Trim(wsAcc.Cells(i, 1).value & "")) = UCase(entryID) Then accCount = accCount + 1
    Next i
    
    If accCount > 0 Then
        html = html & "<table class='acc-table'><tr><th style='width:10%'>#</th><th style='width:45%'>Accessory</th><th style='width:45%'>Brand/Serial</th></tr>"
        j = 1
        For i = 2 To lastRowAcc
            If UCase(Trim(wsAcc.Cells(i, 1).value & "")) = UCase(entryID) Then
                Dim accName As String, accBrand As String
                accName = wsAcc.Cells(i, 4).value & ""
                accBrand = wsAcc.Cells(i, 5).value & ""
                If accBrand = "" Then accBrand = wsAcc.Cells(i, 6).value & ""
                html = html & "<tr><td>" & j & "</td><td>" & accName & "</td><td>" & accBrand & "</td></tr>"
                j = j + 1
                If j > 6 Then Exit For
            End If
        Next i
        html = html & "</table>"
    Else
        html = html & "<p>No Accessories</p>"
    End If
    html = html & "</div>"
    
    html = html & "<div class='section-header problem-header'>PROBLEM</div>"
    html = html & "<div class='content' style='min-height:30px;border:1px solid #CC6600;padding:8px;'>" & IIf(Trim(prob) = "", "Not Specified", prob) & "</div>"
    
    html = html & "<div class='section-header remarks-header'>REMARKS</div>"
    html = html & "<div class='content' style='min-height:30px;border:1px solid #B8860B;padding:8px;'>" & IIf(Trim(remText) = "", "-", remText) & "</div>"
    
    html = html & "<div class='section-header transport-header'>TRANSPORT DETAILS</div>"
    html = html & "<div class='content'><table class='data-table'>"
    html = html & "<tr><td class='label'>Transport Mode:</td><td class='value'>" & courierMode & "</td><td class='label'>Transport Name:</td><td class='value'>" & courierName & "</td></tr>"
    html = html & "<tr><td class='label'>Mobile:</td><td class='value'>" & courierMobile & "</td><td class='label'>Docket Number:</td><td class='value'>" & docketNo & "</td></tr>"
    html = html & "<tr><td class='label'>Send Date:</td><td class='value'>" & sendDate & "</td><td class='label'>Charges:</td><td class='value'>" & charges & "</td></tr>"
    html = html & "</table></div>"
    
    html = html & "<div class='section-header pending-header'>PENDING SERVICES</div>"
    html = html & "<div class='content'>"
    
    Dim lastRowAssign As Long, penEntry As String, penProd As String, penComp As String
    Dim penModel As String, penSerial As String, penSendDate As String
    Dim daysCalc As Long, sendDt As Date, penRows As String
    
    pendingCount = 0
    penRows = ""
    
    If Not wsAssign Is Nothing Then
        lastRowAssign = wsAssign.Cells(wsAssign.Rows.count, 1).End(xlUp).row
        For i = 2 To lastRowAssign
            Dim assignToName As String, statusVal As String
            assignToName = Trim(wsAssign.Cells(i, "N").value & "")
            statusVal = Trim(LCase(wsAssign.Cells(i, "AF").value & ""))
            
            If statusVal = "pending" And Len(assignToName) > 0 Then
                Dim nameMatch As Boolean
                nameMatch = False
                If LCase(assignToName) = LCase(vendName) Then nameMatch = True
                If InStr(1, LCase(vendName), LCase(assignToName)) > 0 Then nameMatch = True
                If InStr(1, LCase(assignToName), LCase(vendName)) > 0 Then nameMatch = True
                
                If nameMatch Then
                    pendingCount = pendingCount + 1
                    penEntry = wsAssign.Cells(i, "B").value & ""
                    penProd = wsAssign.Cells(i, "F").value & ""
                    penComp = wsAssign.Cells(i, "G").value & ""
                    penModel = wsAssign.Cells(i, "H").value & ""
                    penSerial = wsAssign.Cells(i, "I").value & ""
                    
                    On Error Resume Next
                    sendDt = CDate(wsAssign.Cells(i, "Z").value)
                    If Err.Number = 0 Then
                        daysCalc = DateDiff("d", sendDt, Date)
                        penSendDate = Format(sendDt, "dd-mm-yyyy") & " (" & daysCalc & ")"
                    Else
                        penSendDate = "-"
                    End If
                    On Error GoTo ErrorHandler
                    
                    penRows = penRows & "<tr><td>" & penEntry & "</td><td>" & IIf(penProd = "", "-", penProd) & "</td><td>" & IIf(penComp = "", "-", penComp) & "</td><td>" & IIf(penModel = "", "-", penModel) & "</td><td>" & IIf(penSerial = "", "-", penSerial) & "</td><td>" & penSendDate & "</td></tr>"
                End If
            End If
        Next i
    End If
    
    If pendingCount > 0 Then
        html = html & "<table class='pending-table'><tr><th>Entry ID</th><th>Product</th><th>Company</th><th>Model</th><th>Serial</th><th>Send Date (Days)</th></tr>" & penRows & "</table>"
    Else
        html = html & "<p style='text-align:center;font-style:italic;'>No Pending Services</p>"
    End If
    html = html & "</div>"
    
    html = html & "<div class='signature'><div class='sig-line'></div><br><strong>Authorized Signature</strong><br><em>For " & compName & "</em></div>"
    html = html & "<div class='footer'>*** Thank You for choosing " & compName & "! ***</div>"
    html = html & "</div></body></html>"
    
    GenerateAssignmentEmailHTML = html
    Exit Function
    
ErrorHandler:
    GenerateAssignmentEmailHTML = "<html><body><h3>Error: " & Err.Description & "</h3></body></html>"
End Function


'===========================================
' LOAD PAYMENT LIST - ENTRY FORM STYLE
' Filters by EntryID (1 entry per assignment)
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
        ' Match EntryID (Col B) AND Active (Col F blank)
        If Trim(ws.Cells(r, 2).value & "") = Trim(entryID & "") Then
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
    
    ' ===== UPDATE BUTTONS =====
    Call UpdatePaymentButtons
End Sub

'===========================================
' CHECK IF PAYMENT EXISTS FOR THIS ENTRY
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
' If payment exists: Edit=Enabled, Add=Disabled
' If no payment: Add=Enabled, Edit=Disabled
'===========================================
Private Sub UpdatePaymentButtons()
    If Me.lstPayments.ListItems.count > 0 Then
        ' Payment EXISTS
        btnAddPayment.enabled = False
        btnEditPayment.enabled = True
    Else
        ' NO Payment
        btnAddPayment.enabled = True
        btnEditPayment.enabled = False
    End If
    Me.Repaint
End Sub
Private Sub btnAddPayment_Click()
    Dim payForm As frmPaymentSection
    Dim entryID As String
    Dim existingPayID As String
    
    entryID = Trim(cmbEntryID.value)
    If entryID = "" Then
        MsgBox "Please select Entry ID first!", vbExclamation
        cmbEntryID.SetFocus
        Exit Sub
    End If
    
    ' ===== CHECK EXISTING PAYMENT =====
    existingPayID = GetExistingPaymentID(entryID)
    
    Set payForm = New frmPaymentSection
    Set payForm.parentForm = Me
    
    If existingPayID <> "" Then
        ' EXISTING: Open in EDIT mode
        payForm.cmbEntryID.value = entryID
        payForm.cmbEntryID.enabled = False
        payForm.LoadPaymentForEdit existingPayID
    Else
        ' NEW: Open blank
        payForm.cmbEntryID.value = entryID
        payForm.cmbEntryID.enabled = False
    End If
    
    payForm.Show vbModal
    
    Call LoadPaymentList(entryID)
    Set payForm = Nothing
End Sub
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
    
    Call LoadPaymentList(cmbEntryID.value)
End Sub
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

    Call LoadPaymentList(cmbEntryID.value)
End Sub
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
        
        ' Only if data not already loaded in Initialize
        If cmbAssignTo.value = "" Then
            Call LoadServiceAssignData(entryID)
            
            ' Set EDIT mode flags
            isEditMode = True
            ' LastAssignID already set inside LoadServiceAssignData
            
            ' Button states for EDIT
            btnAssign.caption = "UPDATE"
            btnAssign.enabled = True
            
            btnSendWhatsApp.enabled = False
            btnSendEmail.enabled = False
            btnPrintAssign.enabled = False
            btnExportPDF.enabled = False
        End If
        
        Me.Tag = ""
    End If
End Sub
'===========================================
' LOAD SAVED SERVICE ASSIGNMENT DATA (EDIT MODE)
'===========================================
'===========================================
' LOAD SAVED SERVICE ASSIGNMENT DATA (EDIT MODE)
'===========================================
Private Sub LoadServiceAssignData(entryID As String)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim found As Boolean
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Assign_Master")
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "Assign_Master sheet not found!", vbCritical
        Exit Sub
    End If
    
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    found = False
    EditAssignRow = 0
    
    ' Find assignment record by EntryID (Column B) - LATEST record
    For i = lastRow To 2 Step -1
        If UCase(Trim(ws.Cells(i, 2).value)) = UCase(entryID) Then
            found = True
            EditAssignRow = i
            LastAssignID = Trim(ws.Cells(i, 1).value & "")
            
            ' K - AssignTo
            cmbAssignTo.value = Trim(ws.Cells(i, 11).value & "")
            
            ' Refresh dropdown
            Call RefreshAssignNameDropdown(cmbAssignTo.value)
            
            ' N - AssignToName
            cmbAssignName.value = Trim(ws.Cells(i, 14).value & "")
            
            ' Vendor details
            lblVendorMobile.caption = Trim(ws.Cells(i, 15).value & "")
            lblVendorEmail.caption = Trim(ws.Cells(i, 16).value & "")
            lblVendorAddress.caption = Trim(ws.Cells(i, 17).value & "")
            
            ' Courier Mode
            cmbCourierMode.value = Trim(ws.Cells(i, 12).value & "")
            
            ' Load courier companies then set
Call LoadCourierData
On Error Resume Next
cmbCourierCompany.value = Trim(ws.Cells(i, 24).value & "")
On Error GoTo 0
            
            If cmbCourierCompany.ListIndex <> -1 Then
                lblCourierMobile.caption = cmbCourierCompany.Column(1)
            End If
            
            ' Transport Details
            txtDocketNumber.value = Trim(ws.Cells(i, 25).value & "")
            txtSendDate.value = Trim(ws.Cells(i, 26).value & "")
            txtExpectedReturn.value = Trim(ws.Cells(i, 27).value & "")
            txtCourierCharges.value = Trim(ws.Cells(i, 28).value & "")
            txtRemarks.value = Trim(ws.Cells(i, 31).value & "")
            
            ' Load Accessories
            Call LoadAssignAccessories(entryID)
            
            ' Refresh pending list
            Call RefreshPendingList
            
            Exit For
        End If
    Next i
    
    If Not found Then
        MsgBox "Saved Assignment data not found for Entry ID: " & entryID, vbExclamation
        EditAssignRow = 0
    End If
End Sub
'===========================================
' LOAD ASSIGN ACCESSORIES (EDIT MODE)
'===========================================
Private Sub LoadAssignAccessories(entryID As String)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim accCount As Integer
    
    ' Reset all checkboxes first
    For i = 1 To 6
        Me.Controls("CheckBox" & i).value = False
        Me.Controls("CheckBox" & i).caption = ""
        Me.Controls("CheckBox" & i).visible = False
    Next i
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Assign_Accessories")
    On Error GoTo 0
    
    If ws Is Nothing Then Exit Sub
    
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    accCount = 0
    
    ' Column B = EntryID, Column C = Accessory Name, Column D = Status
    For i = 2 To lastRow
        If UCase(Trim(ws.Cells(i, 2).value)) = UCase(entryID) Then
            accCount = accCount + 1
            If accCount <= 6 Then
                Me.Controls("CheckBox" & accCount).caption = ws.Cells(i, 3).value
                Me.Controls("CheckBox" & accCount).visible = True
                Me.Controls("CheckBox" & accCount).value = True
            End If
        End If
    Next i
End Sub

'===========================================
' GET ASSIGN ID FOR GIVEN ENTRY ID
'===========================================
Private Function GetAssignIDForEntry(entryID As String) As String
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Assign_Master")
    On Error GoTo 0
    
    If ws Is Nothing Then Exit Function
    
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    
    For i = lastRow To 2 Step -1
        If UCase(Trim(ws.Cells(i, 2).value)) = UCase(entryID) Then
            GetAssignIDForEntry = Trim(ws.Cells(i, 1).value & "")
            Exit Function
        End If
    Next i
End Function
