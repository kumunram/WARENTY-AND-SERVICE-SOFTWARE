VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAddCustomer 
   Caption         =   "Add New Customer"
   ClientHeight    =   7410
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13455
   OleObjectBlob   =   "frmAddCustomer.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmAddCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mEditRow As Long
Private mIsLoading As Boolean

'========================================
' FORM INITIALIZE
'========================================
Private Sub UserForm_Initialize()
    
    mIsLoading = True
    mEditRow = 0
    
    
    
    'Sab khali
    txtCustomerID.value = ""
    txtMobile.value = ""
    txtCustomerName.value = ""
    txtAddress.value = ""
    txtEmail.value = ""
    txtGST.value = ""
    txtPhotoPath.value = ""
    
    imgCustomerPhoto.PictureSizeMode = fmPictureSizeModeZoom
    Me.caption = "Add New Customer"
    
    mIsLoading = False
    AddMinMaxButtons Me
End Sub

'========================================
' PREPARE NEW CUSTOMER (OTP ke baad call hota hai)
'========================================
Public Sub PrepareNewCustomer(ByVal mobileNo As String)
    mIsLoading = True
    mEditRow = 0          '0 = New Mode
    
    txtMobile.value = mobileNo
    txtCustomerID.value = GetNextCustomerID()   'Auto ID generate
    
    'Baki sab khali
    txtCustomerName.value = ""
    txtAddress.value = ""
    txtEmail.value = ""
    txtGST.value = ""
    txtPhotoPath.value = ""
    
    On Error Resume Next
    imgCustomerPhoto.Picture = LoadPicture()
    On Error GoTo 0
    
    mIsLoading = False
End Sub



'========================================
' GET NEXT CUSTOMER ID
'========================================
Function GetNextCustomerID() As String
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim maxNum As Long
    Dim i As Long
    Dim idVal As String
    Dim numPart As Long
    
    Set ws = Sheets("Customer_Master")
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    
    maxNum = 0
    
    For i = 2 To lastRow
        idVal = CStr(ws.Cells(i, 1).value)
        If Left(idVal, 4) = "CUST" And Len(idVal) >= 5 Then
            On Error Resume Next
            numPart = CLng(Mid(idVal, 5))
            If numPart > maxNum Then maxNum = numPart
            On Error GoTo 0
        End If
    Next i
    
    GetNextCustomerID = "CUST" & Format(maxNum + 1, "0000")
End Function

'========================================
' SAVE CUSTOMER
'========================================
Private Sub btnSaveCustomer_Click()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim newCustID As String
    Dim nextNum As Long
    Dim i As Long
    Dim maxNum As Long
    
    'Validation
    If Trim(txtMobile.value) = "" Then
        MsgBox "Mobile Number required!", vbExclamation
        txtMobile.SetFocus
        Exit Sub
    End If
    
    If Trim(txtCustomerName.value) = "" Then
        MsgBox "Customer Name required!", vbExclamation
        txtCustomerName.SetFocus
        Exit Sub
    End If
    
    Set ws = ThisWorkbook.Sheets("Customer_Master")
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    
    '========== EDIT MODE ==========
    If mEditRow > 0 Then
        ws.Cells(mEditRow, 1).value = Trim(txtCustomerID.value)   'A: ID
        ws.Cells(mEditRow, 2).value = Trim(txtCustomerName.value) 'B: Name
        ws.Cells(mEditRow, 3).value = Trim(txtMobile.value)       'C: Mobile
        ws.Cells(mEditRow, 4).value = Trim(txtAddress.value)      'D: Address
        ws.Cells(mEditRow, 5).value = Trim(txtEmail.value)        'E: Email
        ws.Cells(mEditRow, 6).value = Trim(txtGST.value)          'F: GST
        
        If Trim(txtPhotoPath.value) <> "" Then
            ws.Cells(mEditRow, 7).value = Trim(txtPhotoPath.value) 'G: Photo
        End If
        
        MsgBox "Customer Updated!" & vbCrLf & "ID: " & txtCustomerID.value, vbInformation
        Me.Tag = "SAVED"
        Unload Me
        Exit Sub
    End If
    
    '========== NEW MODE ==========
    maxNum = 0
    For i = 2 To lastRow
        If Left(UCase(Trim(ws.Cells(i, 1).value)), 4) = "CUST" Then
            Dim numPart As String
            numPart = Mid(Trim(ws.Cells(i, 1).value), 5)
            If IsNumeric(numPart) Then
                If CLng(numPart) > maxNum Then maxNum = CLng(numPart)
            End If
        End If
    Next i
    
    nextNum = maxNum + 1
    newCustID = "CUST" & Format(nextNum, "0000")
    
    'ID dikhado form pe
    txtCustomerID.value = newCustID
    
    Dim newRow As Long
    newRow = lastRow + 1
    
    ws.Cells(newRow, 1).value = newCustID                      'A: ID
    ws.Cells(newRow, 2).value = Trim(txtCustomerName.value)    'B: Name
    ws.Cells(newRow, 3).value = Trim(txtMobile.value)          'C: Mobile
    ws.Cells(newRow, 4).value = Trim(txtAddress.value)         'D: Address
    ws.Cells(newRow, 5).value = Trim(txtEmail.value)           'E: Email
    ws.Cells(newRow, 6).value = Trim(txtGST.value)             'F: GST
    
    If Trim(txtPhotoPath.value) <> "" Then
        ws.Cells(newRow, 7).value = Trim(txtPhotoPath.value)   'G: Photo
    End If
    
    ws.Cells(newRow, 10).value = Format(Now, "dd-mm-yyyy hh:mm") 'J: CreatedDate
    
    MsgBox "Customer Saved!" & vbCrLf & "Customer ID: " & newCustID, vbInformation
    
    ' ============================================================
    ' ?? CRITICAL FIX #3: SAVE AFTER CUSTOMER CREATE
    ' ============================================================
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Save
    Application.DisplayAlerts = True
    On Error GoTo 0
    ' ============================================================
    
    Me.Tag = "SAVED"  ' Already exists in your code
    MsgBox "Customer Saved!", vbInformation
    Unload Me
End Sub

'========================================
' PHOTO BUTTONS (same as before)
'========================================
Private Sub btnCapturePhoto_Click()
    shell "explorer.exe shell:AppsFolder\Microsoft.WindowsCamera_8wekyb3d8bbwe!App", vbNormalFocus
    If MsgBox("Photo capture karke Camera close karo, phir OK dabao.", vbOKCancel + vbInformation) = vbOK Then
        ProcessCameraPhoto txtMobile.value
    End If
End Sub

Private Sub btnUpdatePhoto_Click()
    Dim fd As FileDialog
    Dim sourcePath As String
    Dim savedPath As String

    If Trim(txtMobile.value) = "" Then
        MsgBox "Pehle Mobile Number enter karo", vbExclamation
        Exit Sub
    End If

    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    fd.title = "Select Customer Photo"
    fd.Filters.Clear
    fd.Filters.Add "Images", "*.jpg;*.jpeg;*.bmp;*.gif"

    If fd.Show <> -1 Then Exit Sub

    sourcePath = fd.SelectedItems(1)
    savedPath = SavePhotoCopy(sourcePath, Trim(txtMobile.value))

    If savedPath <> "" Then
        txtPhotoPath.value = savedPath
        ShowCustomerPhoto savedPath
    Else
        MsgBox "Photo save nahi ho paya", vbExclamation
    End If
End Sub

Private Function SavePhotoCopy(ByVal sourcePath As String, ByVal mobileNo As String) As String
    Dim fso As Object
    Dim targetFolder As String
    Dim finalPath As String
    Dim basePath As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    
    On Error Resume Next
    targetFolder = GetFolderSetting("B3")
    If targetFolder = "" Then
        basePath = ThisWorkbook.path & "\GLOBAL_SOFT_DATA"
        If Dir(basePath, vbDirectory) = "" Then MkDir basePath
        targetFolder = basePath & "\Images\Customers\"
    End If
    On Error GoTo 0
    
    If Dir(targetFolder, vbDirectory) = "" Then MkDir targetFolder

    If Dir(sourcePath) = "" Then
        MsgBox "Source photo not found", vbCritical
        Exit Function
    End If

    finalPath = targetFolder & mobileNo & ".jpg"

    On Error Resume Next
    fso.CopyFile sourcePath, finalPath, True
    If Err.Number <> 0 Then
        MsgBox "Copy error: " & Err.Description
        Exit Function
    End If
    On Error GoTo 0

    SavePhotoCopy = mobileNo & ".jpg"
End Function

Private Sub ShowCustomerPhoto(ByVal photoName As String)
    Dim fullPath As String
    Dim targetFolder As String
    
    On Error Resume Next
    targetFolder = GetFolderSetting("B3")
    If targetFolder = "" Then
        targetFolder = ThisWorkbook.path & "\GLOBAL_SOFT_DATA\Images\Customers\"
    End If
    On Error GoTo 0

    If Trim(photoName) = "" Then
        imgCustomerPhoto.Picture = LoadPicture()
        Exit Sub
    End If

    fullPath = targetFolder & photoName
    If Dir(fullPath) <> "" Then
        imgCustomerPhoto.Picture = LoadPicture(fullPath)
    Else
        imgCustomerPhoto.Picture = LoadPicture()
    End If
End Sub

Private Sub ProcessCameraPhoto(ByVal mobileNo As String)
    Dim fso As Object, folder As Object, file As Object
    Dim LatestFile As String, latestDate As Date
    Dim cameraPath As String
    Dim targetFolder As String
    Dim finalPath As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    
    cameraPath = Environ("USERPROFILE") & "\Pictures\Camera Roll\"
    If Not fso.FolderExists(cameraPath) Then
        cameraPath = Environ("USERPROFILE") & "\OneDrive\Pictures\Camera Roll\"
    End If

    If Not fso.FolderExists(cameraPath) Then
        MsgBox "Camera folder not found"
        Exit Sub
    End If

    On Error Resume Next
    targetFolder = GetFolderSetting("B3")
    If targetFolder = "" Then
        targetFolder = ThisWorkbook.path & "\GLOBAL_SOFT_DATA\Images\Customers\"
    End If
    On Error GoTo 0
    
    If Dir(targetFolder, vbDirectory) = "" Then MkDir targetFolder

    Set folder = fso.GetFolder(cameraPath)
    For Each file In folder.files
        If file.DateLastModified > latestDate Then
            latestDate = file.DateLastModified
            LatestFile = file.path
        End If
    Next file

    If LatestFile <> "" Then
        finalPath = targetFolder & mobileNo & ".jpg"
        fso.CopyFile LatestFile, finalPath, True
        txtPhotoPath.value = mobileNo & ".jpg"
        imgCustomerPhoto.Picture = LoadPicture(finalPath)
    Else
        MsgBox "Photo detect nahi hua", vbExclamation
    End If
End Sub

Function GetFolderSetting(CellAddress As String) As String
    Dim p As String
    On Error Resume Next
    p = Sheets("Settings").Range(CellAddress).value
    On Error GoTo 0
    If p <> "" And Right(p, 1) <> "\" Then p = p & "\"
    GetFolderSetting = p
End Function

'========================================
' KEYBOARD EVENTS (Uppercase/Lowercase)
'========================================
Private Sub txtCustomerName_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If mIsLoading Then Exit Sub
    Dim pos As Long
    pos = txtCustomerName.SelStart
    txtCustomerName.value = UCase(txtCustomerName.value)
    txtCustomerName.SelStart = pos
End Sub

Private Sub txtAddress_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If mIsLoading Then Exit Sub
    Dim pos As Long
    pos = txtAddress.SelStart
    txtAddress.value = UCase(txtAddress.value)
    txtAddress.SelStart = pos
End Sub

Private Sub txtGST_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If mIsLoading Then Exit Sub
    Dim pos As Long
    pos = txtGST.SelStart
    txtGST.value = UCase(txtGST.value)
    txtGST.SelStart = pos
End Sub

Private Sub txtEmail_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If mIsLoading Then Exit Sub
    Dim pos As Long
    pos = txtEmail.SelStart
    txtEmail.value = LCase(txtEmail.value)
    txtEmail.SelStart = pos
End Sub

Private Sub txtMobile_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If Not (KeyAscii >= 48 And KeyAscii <= 57) Then KeyAscii = 0
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub UserForm_Activate()
    '=== TAG YAHAN CHECK KARO ===
    'Activate .Show ke baad chalta hai, isliye Tag pehle se set hota hai
    If Trim(Me.Tag) <> "" Then
        Call LoadCustomerForEdit(Me.Tag)
        Me.caption = "Edit Customer - " & Me.Tag
    End If
End Sub
Public Sub LoadCustomerForEdit(ByVal custID As String)
    Dim ws As Worksheet
    Dim f As Range
    Dim i As Long

    Set ws = Sheets("Customer_Master")
    
    '=== CUSTOMER ID ?? ????? (Column A = 1) ===
    Set f = ws.Columns(1).Find(What:=Trim(custID), LookIn:=xlValues, LookAt:=xlWhole)

    If f Is Nothing Then
        MsgBox "Customer not found: " & custID, vbExclamation
        Exit Sub
    End If

    mIsLoading = True
    mEditRow = f.row

    txtCustomerID.value = CStr(ws.Cells(f.row, 1).value)     'A: ID
    txtCustomerName.value = CStr(ws.Cells(f.row, 2).value)   'B: Name
    txtMobile.value = CStr(ws.Cells(f.row, 3).value)         'C: Mobile
    txtAddress.value = CStr(ws.Cells(f.row, 4).value)        'D: Address
    txtEmail.value = LCase(ws.Cells(f.row, 5).value)         'E: Email
    txtGST.value = CStr(ws.Cells(f.row, 6).value)            'F: GST
    txtPhotoPath.value = CStr(ws.Cells(f.row, 7).value)      'G: Photo

    '=== PHOTO LOAD ===
    If Trim(txtPhotoPath.value) <> "" Then
        Call ShowCustomerPhoto(txtPhotoPath.value)
    End If
    
    mIsLoading = False
End Sub
