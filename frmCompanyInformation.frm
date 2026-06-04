VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCompanyInformation 
   Caption         =   "COMPANY INFORMATION"
   ClientHeight    =   12465
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12735
   OleObjectBlob   =   "frmCompanyInformation.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmCompanyInformation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
    ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
    ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub UserForm_Initialize()
    LoadComboBoxData
    LoadCompanyInfo
End Sub

Private Sub LoadComboBoxData()
    With Me.cmbCountry
        .Clear
        .AddItem "Add New..."
        .AddItem "India"
        .AddItem "USA"
        .AddItem "UK"
        .AddItem "Australia"
        .AddItem "Canada"
    End With
    
    With Me.cmbState
        .Clear
        .AddItem "Add New..."
        .AddItem "Odisha"
        .AddItem "West Bengal"
        .AddItem "Maharashtra"
        .AddItem "Karnataka"
        .AddItem "Tamil Nadu"
        .AddItem "Delhi"
        .AddItem "Uttar Pradesh"
        .AddItem "Gujarat"
        .AddItem "Rajasthan"
        .AddItem "Punjab"
    End With
    
    With Me.cmbCity
        .Clear
        .AddItem "Add New..."
        .AddItem "Nayagarh"
        .AddItem "Bhubaneswar"
        .AddItem "Cuttack"
        .AddItem "Mumbai"
        .AddItem "Pune"
        .AddItem "Bangalore"
        .AddItem "Chennai"
        .AddItem "Delhi"
        .AddItem "Kolkata"
    End With
    
    With Me.cmbPIN
        .Clear
        .AddItem "Add New..."
        .AddItem "752069"
        .AddItem "751001"
        .AddItem "751003"
        .AddItem "400001"
        .AddItem "411001"
        .AddItem "560001"
    End With
End Sub

Private Sub cmbCity_Click()
    If Me.cmbCity.value = "Add New..." Then
        Dim newCity As String
        newCity = InputBox("Enter New City Name:", "Add New City")
        If newCity <> "" Then
            Me.cmbCity.RemoveItem 0
            Me.cmbCity.AddItem newCity, 0
            Me.cmbCity.AddItem "Add New...", 1
            Me.cmbCity.value = newCity
        Else
            Me.cmbCity.value = ""
        End If
    End If
End Sub

Private Sub cmbPIN_Click()
    If Me.cmbPIN.value = "Add New..." Then
        Dim newPIN As String
        newPIN = InputBox("Enter New PIN Code:", "Add New PIN Code")
        If newPIN <> "" Then
            If Len(newPIN) = 6 And IsNumeric(newPIN) Then
                Me.cmbPIN.RemoveItem 0
                Me.cmbPIN.AddItem newPIN, 0
                Me.cmbPIN.AddItem "Add New...", 1
                Me.cmbPIN.value = newPIN
            Else
                MsgBox "Please enter valid 6-digit PIN Code!", vbExclamation
                Me.cmbPIN.value = ""
            End If
        Else
            Me.cmbPIN.value = ""
        End If
    End If
End Sub

Private Sub cmbState_Click()
    If Me.cmbState.value = "Add New..." Then
        Dim newState As String
        newState = InputBox("Enter New State Name:", "Add New State")
        If newState <> "" Then
            Me.cmbState.RemoveItem 0
            Me.cmbState.AddItem newState, 0
            Me.cmbState.AddItem "Add New...", 1
            Me.cmbState.value = newState
        Else
            Me.cmbState.value = ""
        End If
    End If
End Sub

Private Sub cmbCountry_Click()
    If Me.cmbCountry.value = "Add New..." Then
        Dim newCountry As String
        newCountry = InputBox("Enter New Country Name:", "Add New Country")
        If newCountry <> "" Then
            Me.cmbCountry.RemoveItem 0
            Me.cmbCountry.AddItem newCountry, 0
            Me.cmbCountry.AddItem "Add New...", 1
            Me.cmbCountry.value = newCountry
        Else
            Me.cmbCountry.value = ""
        End If
    End If
End Sub

Private Sub LoadCompanyInfo()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Software_Config")
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    
    Me.txtCompanyName.value = ws.Range("B2").value
    Me.txtAddress1.value = ws.Range("B3").value
    Me.txtAddress2.value = ws.Range("B4").value
    SetComboBoxValue Me.cmbCity, ws.Range("B5").value
    SetComboBoxValue Me.cmbPIN, ws.Range("B6").value
    SetComboBoxValue Me.cmbState, ws.Range("B7").value
    SetComboBoxValue Me.cmbCountry, ws.Range("B8").value
    Me.txtMobile.value = ws.Range("B9").value
    Me.txtEmail.value = ws.Range("B10").value
    Me.txtWebsite.value = ws.Range("B11").value
    Me.txtLogoPath.value = ws.Range("B12").value
    Me.txtReceiptHeader.value = ws.Range("B13").value
    
    If Me.txtLogoPath.value <> "" Then ShowLogoPreview Me.txtLogoPath.value
End Sub

Private Sub SetComboBoxValue(ByRef cmb As ComboBox, ByVal val As String)
    Dim i As Long, found As Boolean
    found = False
    For i = 0 To cmb.ListCount - 1
        If cmb.List(i) = val Then found = True: Exit For
    Next i
    If Not found And val <> "" Then
        cmb.RemoveItem 0
        cmb.AddItem val, 0
        cmb.AddItem "Add New...", 1
    End If
    cmb.value = val
End Sub

Private Sub btnBrowseLogo_Click()
    Dim fd As Office.FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .title = "Select Logo"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "JPG Images", "*.jpg;*.jpeg"
        .Filters.Add "PNG Images", "*.png"
        If .Show = -1 Then
            Me.txtLogoPath.value = .SelectedItems(1)
            ShowLogoPreview .SelectedItems(1)
        End If
    End With
End Sub

Private Sub ShowLogoPreview(ByVal logoPath As String)
    On Error Resume Next
    Me.Image.Picture = LoadPicture(logoPath)
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
End Sub

Private Sub btnSave_Click()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Software_Config")
    ws.Range("B2").value = Me.txtCompanyName.value
    ws.Range("B3").value = Me.txtAddress1.value
    ws.Range("B4").value = Me.txtAddress2.value
    ws.Range("B5").value = Me.cmbCity.value
    ws.Range("B6").value = Me.cmbPIN.value
    ws.Range("B7").value = Me.cmbState.value
    ws.Range("B8").value = Me.cmbCountry.value
    ws.Range("B9").value = Me.txtMobile.value
    ws.Range("B10").value = Me.txtEmail.value
    ws.Range("B11").value = Me.txtWebsite.value
    ws.Range("B12").value = Me.txtLogoPath.value
    ws.Range("B13").value = Me.txtReceiptHeader.value
    MsgBox "Saved Successfully!", vbInformation
End Sub

'===========================================
' RESET TO DEFAULT
'===========================================
Private Sub btnReset_Click()
    Dim result As VbMsgBoxResult
    result = MsgBox("Are you sure you want to reset to default?" & vbCrLf & _
                    "All data will be cleared except Logo and Receipt Header.", _
                    vbQuestion + vbYesNo, "Reset to Default")
    
    If result = vbNo Then Exit Sub
    
    ' Company Details Clear
    Me.txtCompanyName.value = ""
    Me.txtAddress1.value = ""
    Me.txtAddress2.value = ""
    
    ' ComboBox Reset (First item select)
    Me.cmbCity.value = ""
    Me.cmbPIN.value = ""
    Me.cmbState.value = ""
    Me.cmbCountry.value = ""
    
    ' Contact Info Clear
    Me.txtMobile.value = ""
    Me.txtEmail.value = ""
    Me.txtWebsite.value = ""
    
    ' Logo Path - Default ?????
    ' (???? ??? ?? Default path ????)
    Me.txtLogoPath.value = ""
    Me.Image.Picture = Nothing  ' Logo Preview Clear
    
    ' Receipt Header - Default
    Me.txtReceiptHeader.value = "WARRANTY AND SERVICE REPORT"
    
    MsgBox "Reset to Default completed!", vbInformation
End Sub
Private Sub btnClose_Click()
    Unload Me
End Sub

