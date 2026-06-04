VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCustomerList 
   Caption         =   "Customer List"
   ClientHeight    =   8775.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   19350
   OleObjectBlob   =   "frmCustomerList.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmCustomerList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'=== OTP Variables (frmEntryWizard ???? ??) ===
Private mCurrentOTP As String
Private mOTPTimestamp As Date
Private mPendingMobile As String

'========================================
' FORM INITIALIZE
'========================================
Private Sub UserForm_Initialize()
   
    Me.caption = "GLOBAL SOFT - Customer Database"
    SetupListView
    LoadCustomerData
    'Initialize ??? count LoadCustomerData ?? set ?? ????
    lblEntryID.caption = "Customer ID: --"
    lblPresentMobile.caption = "Mobile: --"
     AddMinMaxButtons Me
End Sub

Private Sub SetupListView()
    With lstCustomers
        .ColumnHeaders.Clear
        .ListItems.Clear
        .View = lvwReport
        .FullRowSelect = True
        .Gridlines = True
        
        .ColumnHeaders.Add , , "Customer ID", 90
        .ColumnHeaders.Add , , "Customer Name", 160
        .ColumnHeaders.Add , , "Mobile", 110
        .ColumnHeaders.Add , , "Address", 220
        .ColumnHeaders.Add , , "Email", 160
        .ColumnHeaders.Add , , "GST", 130
    End With
End Sub

Private Sub LoadCustomerData()
    On Error Resume Next
    Dim ws As Worksheet, lastRow As Long, i As Long
    Dim li As Object
    
    Set ws = ThisWorkbook.Sheets("Customer_Master")
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    lstCustomers.ListItems.Clear
    
    For i = 2 To lastRow
        If Trim(ws.Cells(i, 1).value) <> "" Then
            Set li = lstCustomers.ListItems.Add(text:=Trim(ws.Cells(i, 1).value))
            li.SubItems(1) = Trim(ws.Cells(i, 2).value)  'B: Name
            li.SubItems(2) = Trim(ws.Cells(i, 3).value)  'C: Mobile
            li.SubItems(3) = Trim(ws.Cells(i, 4).value)  'D: Address
            li.SubItems(4) = Trim(ws.Cells(i, 5).value)  'E: Email
            li.SubItems(5) = Trim(ws.Cells(i, 6).value)  'F: GST
        End If
    Next i
    
    '=== TOTAL CUSTOMERS ===
    lblCount.caption = "Total: " & lstCustomers.ListItems.count
End Sub

'??? CLICK = SHOW PHOTO + DETAILS ???
Private Sub lstCustomers_Click()
    On Error Resume Next
    Dim ws As Worksheet, custID As String, mobile As String
    Dim photoName As String, targetFolder As String, fullPath As String
    Dim i As Long, lastRow As Long
    
    If lstCustomers.selectedItem Is Nothing Then
        imgPhoto.Picture = Nothing
        lblEntryID.caption = "Customer ID: --"
        lblPresentMobile.caption = "Mobile: --"
        Exit Sub
    End If
    
    custID = lstCustomers.selectedItem.text
    mobile = lstCustomers.selectedItem.SubItems(2)
    
    lblEntryID.caption = "Customer ID: " & custID
    lblPresentMobile.caption = "Mobile: " & mobile
    
    Set ws = ThisWorkbook.Sheets("Customer_Master")
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    photoName = ""
    
    For i = 2 To lastRow
        If Trim(ws.Cells(i, 1).value) = custID Then
            photoName = Trim(ws.Cells(i, 7).value)
            Exit For
        End If
    Next i
    
    targetFolder = GetPhotoFolder()
    If photoName <> "" Then
        fullPath = targetFolder & photoName
        If Dir(fullPath) <> "" Then
            imgPhoto.Picture = LoadPicture(fullPath)
        Else
            fullPath = targetFolder & mobile & ".jpg"
            If Dir(fullPath) <> "" Then
                imgPhoto.Picture = LoadPicture(fullPath)
            Else
                imgPhoto.Picture = Nothing
            End If
        End If
    Else
        imgPhoto.Picture = Nothing
    End If
    On Error GoTo 0
End Sub

'========================================
' ADD NEW WITH OTP VERIFICATION
'========================================
Private Sub btnAddNew_Click()
    Dim inputMobile As String
    Dim cleanMobile As String
    Dim i As Long
    Dim ws As Worksheet
    Dim f As Range
    Dim enteredOTP As String
    Dim timeDiff As Double
    
    '========== MOBILE INPUT ==========
    inputMobile = InputBox("Enter Customer Mobile Number (10 digits):", "New Customer")
    inputMobile = Trim(inputMobile)
    If inputMobile = "" Then Exit Sub
    
    'Numbers only
    For i = 1 To Len(inputMobile)
        If Mid(inputMobile, i, 1) >= "0" And Mid(inputMobile, i, 1) <= "9" Then
            cleanMobile = cleanMobile & Mid(inputMobile, i, 1)
        End If
    Next i
    
    If Len(cleanMobile) <> 10 Then
        MsgBox "Enter valid 10-digit mobile!", vbExclamation
        Exit Sub
    End If
    
    '========== DUPLICATE CHECK ==========
    Set ws = ThisWorkbook.Sheets("Customer_Master")
    Set f = ws.Columns(3).Find(What:=cleanMobile, LookAt:=xlWhole, MatchCase:=False)
    
    If Not f Is Nothing Then
        MsgBox "Customer already exists!" & vbCrLf & "Name: " & ws.Cells(f.row, 2).value, vbExclamation
        'List me select karo
        Dim li As Object
        For Each li In lstCustomers.ListItems
            If Trim(li.SubItems(2)) = cleanMobile Then
                li.Selected = True
                li.EnsureVisible
                Call lstCustomers_Click
                Exit For
            End If
        Next li
        Exit Sub
    End If
    
    '========== OTP ==========
    If MsgBox("Send OTP to +91 " & cleanMobile & "?", vbYesNo + vbQuestion) = vbNo Then Exit Sub
    
    Randomize
    mCurrentOTP = Format(Int(Rnd() * 900000) + 100000, "000000")
    mOTPTimestamp = Now
    Call SendWhatsAppOTP(cleanMobile, mCurrentOTP)
    
    enteredOTP = InputBox("Enter 6-digit OTP sent to WhatsApp:", "Verify OTP")
    If Trim(enteredOTP) = "" Then
        MsgBox "OTP Required!", vbExclamation
        Exit Sub
    End If
    
    timeDiff = DateDiff("n", mOTPTimestamp, Now)
    If timeDiff > 10 Then
        MsgBox "OTP Expired!", vbCritical
        Exit Sub
    End If
    
    If enteredOTP <> mCurrentOTP Then
        MsgBox "Invalid OTP!", vbCritical
        Exit Sub
    End If
    
    MsgBox "OTP Verified Successfully!", vbInformation
    
    '========== OPEN ADD CUSTOMER FORM ==========
    Dim custForm As frmAddCustomer
    Set custForm = New frmAddCustomer
    custForm.PrepareNewCustomer cleanMobile   'ID auto generate hogi
    
    Me.Hide
    custForm.Show vbModal
    Me.Show
    
    Unload custForm
    Set custForm = Nothing
    
    'Refresh list
    Call LoadCustomerData
    imgPhoto.Picture = Nothing
    lblEntryID.caption = "Customer ID: --"
    lblPresentMobile.caption = "Mobile: --"
End Sub

'========================================
' SEND WHATSAPP OTP
'========================================
Private Sub SendWhatsAppOTP(mobile As String, otp As String)
    Dim url As String
    Dim message As String
    Dim wsh As Object
    
    If Trim(mobile) = "" Then Exit Sub
    
    message = "GLOBAL IT SOLUTIONS" & vbCrLf & vbCrLf & _
              "New Customer Verification Code: " & otp & vbCrLf & vbCrLf & _
              "This Code is Valid for 10 minutes" & vbCrLf & _
              "Do not share this code with anyone"
    
    url = "whatsapp://send?phone=91" & Trim(mobile) & "&text=" & URLEncode(message)
    
    Set wsh = CreateObject("WScript.Shell")
    wsh.Run url, 1, False
    Set wsh = Nothing
End Sub

'========================================
' URL ENCODE
'========================================
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
            ' ignore
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

Private Sub btnEdit_Click()
    If lstCustomers.selectedItem Is Nothing Then
        MsgBox "Please select a customer first!", vbExclamation
        Exit Sub
    End If
    
    Dim custID As String
    custID = lstCustomers.selectedItem.text  'Column 0 = Customer ID (Unique)
    
    Dim editForm As frmAddCustomer
    Set editForm = New frmAddCustomer
    editForm.Tag = custID        'ID pass ???
    
    Me.Hide
    editForm.Show vbModal
    Me.Show
    
    Unload editForm
    Set editForm = Nothing
    
    Call LoadCustomerData
    imgPhoto.Picture = Nothing
    lblEntryID.caption = "Customer ID: --"
    lblPresentMobile.caption = "Mobile: --"
End Sub
Private Sub btnSearch_Click()
    Dim searchText As String, ws As Worksheet
    Dim lastRow As Long, i As Long, li As Object
    
    searchText = UCase(Trim(txtSearch.value))
    If searchText = "" Then
        LoadCustomerData
        Exit Sub
    End If
    
    Set ws = ThisWorkbook.Sheets("Customer_Master")
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    lstCustomers.ListItems.Clear
    
    For i = 2 To lastRow
        If InStr(1, UCase(Trim(ws.Cells(i, 1).value) & " " & Trim(ws.Cells(i, 2).value) & " " & Trim(ws.Cells(i, 3).value)), searchText, vbTextCompare) > 0 Then
            Set li = lstCustomers.ListItems.Add(text:=Trim(ws.Cells(i, 1).value))
            li.SubItems(1) = Trim(ws.Cells(i, 2).value)
            li.SubItems(2) = Trim(ws.Cells(i, 3).value)
            li.SubItems(3) = Trim(ws.Cells(i, 4).value)
            li.SubItems(4) = Trim(ws.Cells(i, 5).value)
            li.SubItems(5) = Trim(ws.Cells(i, 6).value)
        End If
    Next i
    
    '=== SEARCH RESULT ===
    lblCount.caption = "Found: " & lstCustomers.ListItems.count
End Sub

Private Sub btnRefresh_Click()
    txtSearch.value = ""
    LoadCustomerData        '?? Total: X set ?? ????
    imgPhoto.Picture = Nothing
    lblEntryID.caption = "Customer ID: --"
    lblPresentMobile.caption = "Mobile: --"
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

'Helper - Photo folder path
Private Function GetPhotoFolder() As String
    Dim p As String
    On Error Resume Next
    p = ThisWorkbook.Sheets("Settings").Range("B3").value
    On Error GoTo 0
    If p = "" Then
        p = ThisWorkbook.path & "\GLOBAL_SOFT_DATA\Images\Customers\"
    End If
    If Right(p, 1) <> "\" Then p = p & "\"
    GetPhotoFolder = p
End Function

