VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAddProduct 
   Caption         =   "AddProduct"
   ClientHeight    =   12810
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13755
   OleObjectBlob   =   "frmAddProduct.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmAddProduct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public selectedEntryID As String
Public SelectedPhotoBox As Integer
Public CurrentEntryID As String

Function GetProductPhotoFolder() As String
    Dim basePath As String
    Dim entryID As String
    Dim folderPath As String

    basePath = Trim(Sheets("Settings").Range("B2").value)
    entryID = Trim(Me.txtEntryID.caption)

    ' STRICT CHECK
    If basePath = "" Then
        MsgBox "Product Photo Path not set in Settings (B2)!", vbCritical, "Settings Error"
        GetProductPhotoFolder = ""
        Exit Function
    End If
    
    If entryID = "" Then
        MsgBox "Entry ID is required! Please select Entry Type first.", vbCritical, "Entry ID Missing"
        GetProductPhotoFolder = ""
        Exit Function
    End If

    ' Create main folder
    If Dir(basePath, vbDirectory) = "" Then MkDir basePath

    folderPath = basePath & "\" & entryID
    If Dir(folderPath, vbDirectory) = "" Then MkDir folderPath

    GetProductPhotoFolder = folderPath
End Function

Private Function GetPhotoFilePath(ByVal photoNo As Integer) As String

Dim folderPath As String

folderPath = GetProductPhotoFolder()

If folderPath = "" Then Exit Function

GetPhotoFilePath = folderPath & "\Photo" & photoNo & ".jpg"

End Function
Private Sub btnAddPhoto_Click()

Dim fd As FileDialog
Dim filePath As String

Set fd = Application.FileDialog(msoFileDialogFilePicker)

fd.Filters.Clear
fd.Filters.Add "Image Files", "*.jpg;*.jpeg"

If fd.Show = -1 Then

filePath = fd.SelectedItems(1)

If Dir(filePath) = "" Then Exit Sub

On Error Resume Next

Select Case SelectedPhotoBox

Case 1
Call CreateProductPhotoFolder
FileCopy filePath, GetPhotoFilePath(1)
img1.Picture = LoadPicture(GetPhotoFilePath(1))

Case 2
Call CreateProductPhotoFolder
FileCopy filePath, GetPhotoFilePath(2)
img2.Picture = LoadPicture(GetPhotoFilePath(2))

Case 3
Call CreateProductPhotoFolder
FileCopy filePath, GetPhotoFilePath(3)
img3.Picture = LoadPicture(GetPhotoFilePath(3))

Case 4
Call CreateProductPhotoFolder
FileCopy filePath, GetPhotoFilePath(4)
img4.Picture = LoadPicture(GetPhotoFilePath(4))
End Select

On Error GoTo 0

DoEvents
Me.Repaint

End If

End Sub
Private Sub btnCapture_Click()
    On Error Resume Next
    
    If Trim(Me.txtEntryID.caption) = "" Then MsgBox "Save product first", vbExclamation: Exit Sub
    If SelectedPhotoBox = 0 Then MsgBox "Select photo box first", vbExclamation: Exit Sub
    
    Dim lastFile As String, newFile As String, wsh As Object, targetPath As String
    
    ' Remember current photo
    lastFile = GetLatestCameraPhoto()
    
    ' Open Camera (No MsgBox)
    shell "explorer.exe shell:AppsFolder\Microsoft.WindowsCamera_8wekyb3d8bbwe!App", vbNormalFocus
    
    ' Fast check loop every 0.2 seconds
    Do
        DoEvents
        newFile = GetLatestCameraPhoto()
        
        ' New photo detected!
        If newFile <> "" And newFile <> lastFile Then
            Application.Wait Now + TimeValue("00:00:00.5")
            
            ' Close camera
            Set wsh = CreateObject("WScript.Shell")
            wsh.Run "taskkill /f /im WindowsCamera.exe", 0, False
            
            ' Save path
            Select Case SelectedPhotoBox
                Case 1: targetPath = GetPhotoFilePath(1)
                Case 2: targetPath = GetPhotoFilePath(2)
                Case 3: targetPath = GetPhotoFilePath(3)
                Case 4: targetPath = GetPhotoFilePath(4)
            End Select
            
            If Dir(targetPath) <> "" Then Kill targetPath
            FileCopy newFile, targetPath
            
            ' Load fast
            Select Case SelectedPhotoBox
                Case 1: Set Me.img1.Picture = LoadPicture(targetPath)
                Case 2: Set Me.img2.Picture = LoadPicture(targetPath)
                Case 3: Set Me.img3.Picture = LoadPicture(targetPath)
                Case 4: Set Me.img4.Picture = LoadPicture(targetPath)
            End Select
            
            AutoNextPhotoBox
            Exit Do
        End If
        
        Application.Wait Now + TimeValue("00:00:00.2")
    Loop While Timer < Timer + 60 ' 60 sec timeout
    
    On Error GoTo 0
End Sub

Function GetLatestCameraPhoto() As String
    On Error Resume Next ' ???? ???? error suppression
    
    Dim fso As Object, folder As Object, file As Object, latestDate As Date
    GetLatestCameraPhoto = ""
    latestDate = #1/1/1900#
    
    Dim camPath As String
    camPath = Environ("USERPROFILE") & "\Pictures\Camera Roll\"
    If Dir(camPath, vbDirectory) = "" Then camPath = Environ("USERPROFILE") & "\OneDrive\Pictures\Camera Roll\"
    If Dir(camPath, vbDirectory) = "" Then Exit Function
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(camPath)
    
    For Each file In folder.files
        If LCase(Right(file.name, 4)) = ".jpg" Then
            If file.DateLastModified > latestDate Then
                latestDate = file.DateLastModified
                GetLatestCameraPhoto = file.path ' Direct string
            End If
        End If
    Next
    
    Set fso = Nothing
    On Error GoTo 0
End Function

'=================================
' PROCESS PHOTO (UNCHANGED)
'=================================
Private Sub ProcessCapturedPhoto(photoPath As String)
    On Error GoTo ErrorHandler
    
    Call CreateProductPhotoFolder
    
    Dim targetPath As String
    Select Case SelectedPhotoBox
        Case 1: targetPath = GetPhotoFilePath(1)
        Case 2: targetPath = GetPhotoFilePath(2)
        Case 3: targetPath = GetPhotoFilePath(3)
        Case 4: targetPath = GetPhotoFilePath(4)
    End Select
    
    If Dir(targetPath) <> "" Then Kill targetPath
    Application.Wait Now + TimeValue("00:00:01")
    
    FileCopy photoPath, targetPath
    
    Select Case SelectedPhotoBox
        Case 1: Set Me.img1.Picture = LoadPicture(targetPath)
        Case 2: Set Me.img2.Picture = LoadPicture(targetPath)
        Case 3: Set Me.img3.Picture = LoadPicture(targetPath)
        Case 4: Set Me.img4.Picture = LoadPicture(targetPath)
    End Select
    
    AutoNextPhotoBox
    
    Exit Sub
ErrorHandler:
    MsgBox "Photo load error: " & Err.Description, vbExclamation
End Sub


Sub LoadCapturedPhotoAuto(ByVal photoPath As String)

    ' Check folder
    If GetProductPhotoFolder() = "" Then Exit Sub

    ' Wait for file release
    Dim waitCount As Integer
    For waitCount = 1 To 10
        On Error Resume Next
        Open photoPath For Input As #1
        Close #1
        If Err.Number = 0 Then Exit For
        Application.Wait Now + TimeValue("00:00:01")
        DoEvents
    Next

    Dim targetPath As String
    
    Select Case SelectedPhotoBox
        Case 1: targetPath = GetPhotoFilePath(1)
        Case 2: targetPath = GetPhotoFilePath(2)
        Case 3: targetPath = GetPhotoFilePath(3)
        Case 4: targetPath = GetPhotoFilePath(4)
        Case Else: Exit Sub
    End Select

    ' Copy with error handling
    On Error Resume Next
    FileCopy photoPath, targetPath
    If Err.Number = 70 Then
        MsgBox "Camera is still using this photo. Please close Camera app and try again.", vbExclamation
        Exit Sub
    End If
    On Error GoTo 0

    ' Load picture
    On Error Resume Next
    Select Case SelectedPhotoBox
        Case 1: Me.img1.Picture = LoadPicture(targetPath)
        Case 2: Me.img2.Picture = LoadPicture(targetPath)
        Case 3: Me.img3.Picture = LoadPicture(targetPath)
        Case 4: Me.img4.Picture = LoadPicture(targetPath)
    End Select
    On Error GoTo 0

    AutoNextPhotoBox
End Sub
Private Sub btnEditAccessories_Click()
    Dim entryID As String
    
    If frmEntryWizard.lstProducts.selectedItem Is Nothing Then
        MsgBox "Select product first", vbExclamation
        Exit Sub
    End If
    
    entryID = frmEntryWizard.lstProducts.selectedItem.SubItems(1)
    
    Dim accForm As frmAccessories
    Set accForm = New frmAccessories
    
    ' ?? ?????? ?????
    accForm.customerID = frmEntryWizard.txtCustomerID.caption
    accForm.txtEntryID.value = entryID
    accForm.cmbProduct.value = frmEntryWizard.lstProducts.selectedItem.SubItems(2)
    accForm.Tag = entryID  ' ???? ???
    
    accForm.Show vbModal
    Set accForm = Nothing
End Sub
'=================================
' SAVE BUTTON CLICK (PERFECT)
'=================================
Private Sub btnSaveProduct_Click()
    Dim entryID As String
    Dim isEditMode As Boolean
    
    ' 1. VALIDATION
    If Not ValidateAllFields() Then Exit Sub
    
    ' 2. ENTRY ID GENERATION
    If Me.Tag = "" Then
        entryID = GeneratePerfectEntryID(Me.cmbEntryType.value)
        isEditMode = False
    Else
        entryID = Me.Tag
        isEditMode = True
    End If
    
    ' 3. SAVE DATA (DIRECT TO SHEET - NO MODULE)
    Call SaveProductToSheet(entryID)
    
    ' 4. SAVE PHOTOS
    Call SaveProductPhotos(entryID)
    
    If isEditMode Then
    MsgBox "Product Updated: " & entryID, vbInformation
Else

' ============================================================
    ' ?? CRITICAL FIX #2: SAVE AFTER PRODUCT ADD
    ' ============================================================
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Save
    Application.DisplayAlerts = True
    On Error GoTo 0
    ' ============================================================
    
    MsgBox "Product Saved: " & entryID, vbInformation
End If
    
    frmEntryWizard.LoadProductList frmEntryWizard.txtCustomerID.caption
    
    Unload Me
End Sub
'=================================
' DIRECT SAVE TO SHEET (SAFE)
'=================================
Private Sub SaveProductToSheet(entryID As String)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim targetRow As Long
    Dim i As Long
    
    Set ws = ThisWorkbook.Sheets("Job_Product")
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    
    ' Find if EntryID already exists (Edit Mode)
    targetRow = 0
    For i = 2 To lastRow
        If UCase(Trim(ws.Cells(i, 1).value)) = UCase(Trim(entryID)) Then
            targetRow = i
            Exit For
        End If
    Next i
    
    ' If not found, add new row
    If targetRow = 0 Then
        targetRow = lastRow + 1
    End If
    
    ' CORRECT COLUMN MAPPING (Sheet structure ?? ????? ??)
    With ws.Rows(targetRow)
        .Cells(1, 1).value = entryID                    ' A: EntryID
        .Cells(1, 2).value = frmEntryWizard.txtCustomerID.caption  ' B: CustomerID
        .Cells(1, 3).value = Me.cmbEntryType.value      ' C: EntryType
        .Cells(1, 4).value = "ACTIVE"                   ' D: Status (Default)
        .Cells(1, 5).value = Format(Now, "dd-mm-yyyy hh:mm") ' E: Last Status Date/Time
        .Cells(1, 6).value = Me.cmbProductType.value    ' F: ProductType (SSD)
        .Cells(1, 7).value = Me.cmbCompany.value        ' G: Company (GEONIX)
        .Cells(1, 8).value = Me.cmbModel.value          ' H: Model (256GB)
        .Cells(1, 9).value = Me.txtSerialNumber.value   ' I: SerialNumber (12345ASDFG)
        .Cells(1, 10).value = Me.txtPurchaseDate.value  ' J: PurchaseDate
        .Cells(1, 11).value = Me.cmbWarrantyYear.value  ' K: WarrantyYear (1 YEAR)
        .Cells(1, 12).value = Me.lblWarrantyStatus.caption ' L: WarrantyStatus
        .Cells(1, 13).value = Me.txtProblem.value       ' M: Problem (NOT WORKING)
        .Cells(1, 14).value = Me.txtRemark.value        ' N: Remark
    End With
End Sub



'=================================
' VALIDATION FUNCTION (TOP TO BOTTOM)
'=================================
Private Function ValidateAllFields() As Boolean
    ValidateAllFields = False
    
    ' 1. ENTRY TYPE
    If Trim(Me.cmbEntryType.value) = "" Then
        MsgBox "Please select Entry Type!" & vbCrLf & "Entry Type is required.", vbExclamation, "Validation Error"
        Me.cmbEntryType.SetFocus
        Exit Function
    End If
    
    ' 2. PRODUCT TYPE
    If Trim(Me.cmbProductType.value) = "" Then
        MsgBox "Please select Product Type!" & vbCrLf & "Product Type is required.", vbExclamation, "Validation Error"
        Me.cmbProductType.SetFocus
        Exit Function
    End If
    
    ' 3. COMPANY
    If Trim(Me.cmbCompany.value) = "" Then
        MsgBox "Please select Company!" & vbCrLf & "Company is required.", vbExclamation, "Validation Error"
        Me.cmbCompany.SetFocus
        Exit Function
    End If
    
    ' 4. MODEL
    If Trim(Me.cmbModel.value) = "" Then
        MsgBox "Please select Model!" & vbCrLf & "Model is required.", vbExclamation, "Validation Error"
        Me.cmbModel.SetFocus
        Exit Function
    End If
    
    ' 5. SERIAL NUMBER
    If Trim(Me.txtSerialNumber.value) = "" Then
        MsgBox "Please enter Serial Number!" & vbCrLf & "Serial Number is required.", vbExclamation, "Validation Error"
        Me.txtSerialNumber.SetFocus
        Exit Function
    End If
    
    ' 6. WARRANTY FIELDS (Only for Warranty)
    If UCase(Trim(Me.cmbEntryType.value)) = "WARRANTY" Then
        If Trim(Me.txtPurchaseDate.value) = "" Then
            MsgBox "Please enter Purchase Date!" & vbCrLf & "Purchase Date is required for Warranty.", vbExclamation, "Validation Error"
            Me.txtPurchaseDate.SetFocus
            Exit Function
        End If
        
        If Trim(Me.cmbWarrantyYear.value) = "" Then
            MsgBox "Please select Warranty Year!" & vbCrLf & "Warranty Year is required for Warranty.", vbExclamation, "Validation Error"
            Me.cmbWarrantyYear.SetFocus
            Exit Function
        End If
    End If
    
    ' 7. PROBLEM (Mandatory)
    If Trim(Me.txtProblem.value) = "" Then
        MsgBox "Please enter Problem Description!" & vbCrLf & "Problem is required.", vbExclamation, "Validation Error"
        Me.txtProblem.SetFocus
        Exit Function
    End If
    
    ' 8. PHOTO VALIDATION (At least 1 out of 4)
    Dim hasPhoto As Boolean
    hasPhoto = False
    
    On Error Resume Next
    If Not Me.img1.Picture Is Nothing Then hasPhoto = True
    If Not Me.img2.Picture Is Nothing Then hasPhoto = True
    If Not Me.img3.Picture Is Nothing Then hasPhoto = True
    If Not Me.img4.Picture Is Nothing Then hasPhoto = True
    On Error GoTo 0
    
    If Not hasPhoto Then
        MsgBox "Please capture at least 1 Photo!" & vbCrLf & "Minimum 1 photo is required.", vbExclamation, "Photo Required"
        Call HighlightPhotoBoxesRed
        Me.img1.SetFocus
        Exit Function
    Else
        Call ResetPhotoBorders
    End If
    
    ' All passed
    ValidateAllFields = True
End Function

'=================================
' PHOTO HELPERS
'=================================
Private Sub HighlightPhotoBoxesRed()
    On Error Resume Next
    If Me.img1.Picture Is Nothing Then Me.img1.BorderStyle = fmBorderStyleSingle: Me.img1.BorderColor = RGB(255, 0, 0)
    If Me.img2.Picture Is Nothing Then Me.img2.BorderStyle = fmBorderStyleSingle: Me.img2.BorderColor = RGB(255, 0, 0)
    If Me.img3.Picture Is Nothing Then Me.img3.BorderStyle = fmBorderStyleSingle: Me.img3.BorderColor = RGB(255, 0, 0)
    If Me.img4.Picture Is Nothing Then Me.img4.BorderStyle = fmBorderStyleSingle: Me.img4.BorderColor = RGB(255, 0, 0)
    Me.Repaint
    On Error GoTo 0
End Sub

Private Sub ResetPhotoBorders()
    On Error Resume Next
    Me.img1.BorderStyle = fmBorderStyleNone
    Me.img2.BorderStyle = fmBorderStyleNone
    Me.img3.BorderStyle = fmBorderStyleNone
    Me.img4.BorderStyle = fmBorderStyleNone
    Me.Repaint
    On Error GoTo 0
End Sub


'=================================
' CHECK AT LEAST ONE PHOTO EXISTS
'=================================
Private Function AtLeastOnePhotoExists() As Boolean
    On Error Resume Next
    
    If Not Me.img1.Picture Is Nothing Then AtLeastOnePhotoExists = True: Exit Function
    If Not Me.img2.Picture Is Nothing Then AtLeastOnePhotoExists = True: Exit Function
    If Not Me.img3.Picture Is Nothing Then AtLeastOnePhotoExists = True: Exit Function
    If Not Me.img4.Picture Is Nothing Then AtLeastOnePhotoExists = True: Exit Function
    
    AtLeastOnePhotoExists = False
    On Error GoTo 0
End Function

'=================================
' HIGHLIGHT PHOTO BOXES (RED BORDER)
'=================================
Private Sub HighlightPhotoBoxes(highlight As Boolean)
    On Error Resume Next
    
    If highlight Then
        ' Red border for all empty photo boxes
        If Me.img1.Picture Is Nothing Then
            Me.img1.BorderStyle = fmBorderStyleSingle
            Me.img1.BorderColor = RGB(255, 0, 0)  ' Red
        End If
        If Me.img2.Picture Is Nothing Then
            Me.img2.BorderStyle = fmBorderStyleSingle
            Me.img2.BorderColor = RGB(255, 0, 0)
        End If
        If Me.img3.Picture Is Nothing Then
            Me.img3.BorderStyle = fmBorderStyleSingle
            Me.img3.BorderColor = RGB(255, 0, 0)
        End If
        If Me.img4.Picture Is Nothing Then
            Me.img4.BorderStyle = fmBorderStyleSingle
            Me.img4.BorderColor = RGB(255, 0, 0)
        End If
    Else
        ' Remove red border (back to normal)
        Me.img1.BorderStyle = fmBorderStyleNone
        Me.img2.BorderStyle = fmBorderStyleNone
        Me.img3.BorderStyle = fmBorderStyleNone
        Me.img4.BorderStyle = fmBorderStyleNone
    End If
    
    Me.Repaint
    On Error GoTo 0
End Sub

'=================================
' SAVE PHOTOS (Settings ?? ????? ??)
'=================================
Private Sub SaveProductPhotos(entryID As String)
    Dim folderPath As String
    Dim fso As Object
    Dim basePath As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Settings ?? ??? ??? (B2), ??? ???? ?? ?? Default ???
    basePath = Trim(Sheets("Settings").Range("B2").value)
    If basePath = "" Then
        basePath = ThisWorkbook.path & "\GLOBAL_SOFT_DATA\Images\Products"
    End If
    
    ' \ ???? ?? ?? ??????
    If Right(basePath, 1) <> "\" Then basePath = basePath & "\"
    
    folderPath = basePath & entryID & "\"
    
    ' ?????? ?????
    If Not fso.FolderExists(folderPath) Then
        fso.CreateFolder folderPath
    End If
    
    ' ???? ??? ????
    On Error Resume Next
    
    If Not Me.img1.Picture Is Nothing Then
        SavePicture Me.img1.Picture, folderPath & "Photo1.jpg"
    End If
    
    If Not Me.img2.Picture Is Nothing Then
        SavePicture Me.img2.Picture, folderPath & "Photo2.jpg"
    End If
    
    If Not Me.img3.Picture Is Nothing Then
        SavePicture Me.img3.Picture, folderPath & "Photo3.jpg"
    End If
    
    If Not Me.img4.Picture Is Nothing Then
        SavePicture Me.img4.Picture, folderPath & "Photo4.jpg"
    End If
    
    On Error GoTo 0
End Sub
Private Sub CommandButton1_Click()

End Sub

Private Sub img1_Click()

SelectedPhotoBox = 1
ResetBorder

img1.BorderStyle = fmBorderStyleSingle
img1.BorderColor = vbRed

Me.Repaint

End Sub
Private Sub img2_Click()

SelectedPhotoBox = 2
ResetBorder

img2.BorderStyle = fmBorderStyleSingle
img2.BorderColor = vbRed

Me.Repaint

End Sub
Private Sub img3_Click()

SelectedPhotoBox = 3
ResetBorder

img3.BorderStyle = fmBorderStyleSingle
img3.BorderColor = vbRed

Me.Repaint

End Sub
Private Sub img4_Click()

SelectedPhotoBox = 4
ResetBorder

img4.BorderStyle = fmBorderStyleSingle
img4.BorderColor = vbRed

Me.Repaint

End Sub
Sub ResetBorder()

img1.BorderStyle = fmBorderStyleNone
img2.BorderStyle = fmBorderStyleNone
img3.BorderStyle = fmBorderStyleNone
img4.BorderStyle = fmBorderStyleNone

DoEvents
Me.Repaint

End Sub
Private Sub btnUpdate_Click()

If SelectedPhotoBox = 0 Then
    MsgBox "Please select photo box first", vbExclamation
    Exit Sub
End If

Dim selectedBox As Integer
selectedBox = SelectedPhotoBox   ' freeze selection

' Open Camera
shell "explorer.exe shell:AppsFolder\Microsoft.WindowsCamera_8wekyb3d8bbwe!App", vbNormalFocus

' Message
If MsgBox("New photo capture karke Camera close karo, phir OK dabao.", vbOKCancel + vbInformation) <> vbOK Then Exit Sub

' Replace photo
Call ReplaceCapturedPhoto(selectedBox)

AutoNextPhotoBox

End Sub
Private Sub btnDeletePhoto_Click()

If SelectedPhotoBox = 0 Then
MsgBox "Please select photo first"
Exit Sub
End If

Select Case SelectedPhotoBox

Case 1
Set img1.Picture = Nothing

Case 2
Set img2.Picture = Nothing

Case 3
Set img3.Picture = Nothing

Case 4
Set img4.Picture = Nothing

End Select

DoEvents
Me.Repaint

End Sub

Private Sub btnAccessories_Click()
    On Error GoTo ErrorHandler
    
    If Trim(txtEntryID.caption) = "" Then
        MsgBox "Please save product first! Entry ID is required.", vbExclamation, "Required"
        Exit Sub
    End If

    If Trim(cmbProductType.value) = "" Then
        MsgBox "Please select product first", vbExclamation
        Exit Sub
    End If

    Dim accForm As frmAccessories
    Set accForm = New frmAccessories
    
    ' ?? ?????? ????? - CustomerID ??? ???
    accForm.customerID = frmEntryWizard.txtCustomerID.caption
    accForm.txtEntryID.value = txtEntryID.caption
    accForm.cmbProduct.value = cmbProductType.value
    accForm.Tag = ""  ' ??? ???
    
    accForm.Show vbModal
    Set accForm = Nothing
    Exit Sub
    
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
End Sub
Private Sub UserForm_Initialize()

AddMinMaxButtons Me

 Dim photosEnabled As Boolean
    photosEnabled = IsFeatureEnabled("PHOTOS")
    
     '===== PHOTO BUTTONS SETUP FOR ADD MODE =====
    If Me.Tag = "" Then
        ' ADD MODE (New Product)
        Me.btnCapture.visible = True        ' Capture button dikhna chahiye
        Me.btnDeletePhoto.visible = True    ' Delete Photo button dikhna chahiye
        Me.btnUpdate.visible = False        ' Update button hide
        Me.btnAddPhoto.visible = False      ' Add Photo button hide
        
        ' Accessories buttons
        Me.btnAccessories.visible = True
        Me.btnEditAccessories.visible = False
    Else
        ' EDIT MODE (Existing Product)
        Call SetupEditModeButtons
    End If
   
    AddMinMaxButtons Me
    
    If Me.Tag <> "" Then
        Call LoadProductPhotos(Me.Tag)  ' ?? ??? ?? ??
    End If
    

    
cmbEntryType.Clear
cmbEntryType.AddItem "Warranty"
cmbEntryType.AddItem "Service"
cmbEntryType = ""

SelectedPhotoBox = 1

LoadProductType
LoadCompany
LoadModel
LoadWarrantyYear

lblWarrantyStatus.caption = ""
lblPurchaseDate.visible = False
txtPurchaseDate.visible = False

lblWarrantyYear.visible = False
cmbWarrantyYear.visible = False

lblWarrantyStatus.visible = False
btnAddWarranty.visible = False

If Me.Tag <> "" Then
    Call LoadProductPhotos(Me.Tag)    ' <-- Product form ??? Product ???? sub ???? ?????
End If



End Sub

Private Sub cmbEntryType_Change()

' =========================
' EDIT MODE LOCK (IMPORTANT)
' =========================
If Me.Tag <> "" Then Exit Sub

Dim ws As Worksheet
Dim lastRow As Long
Dim lastID As String
Dim num As Long

Set ws = Sheets("Job_Product")

lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row

If lastRow < 2 Then
    num = 1
Else
    lastID = ws.Cells(lastRow, 1).value
    num = CLng(Mid(lastID, 4)) + 1
End If

' =========================
' ENTRY ID GENERATION
' =========================
If cmbEntryType.value = "Warranty" Then
    txtEntryID.caption = "WAR" & Format(num, "00000")
ElseIf cmbEntryType.value = "Service" Then
    txtEntryID.caption = "SER" & Format(num, "00000")
End If

' =========================
' UI SHOW / HIDE
' =========================
If cmbEntryType.value = "Warranty" Then

    lblPurchaseDate.visible = True
    txtPurchaseDate.visible = True
    
    lblWarrantyYear.visible = True
    cmbWarrantyYear.visible = True
    
    lblWarrantyStatus.visible = True
    btnAddWarranty.visible = True

Else

    lblPurchaseDate.visible = False
    txtPurchaseDate.visible = False
    
    lblWarrantyYear.visible = False
    cmbWarrantyYear.visible = False
    
    lblWarrantyStatus.visible = False
    btnAddWarranty.visible = False

End If

' =========================
' PASS ENTRY ID
' =========================
frmEntryWizard.selectedEntryID = txtEntryID.caption

End Sub




'========================
' ADD PRODUCT TYPE
'========================
Private Sub btnAddProductType_Click()

Dim ws As Worksheet
Dim val As String
Dim r As Long
Dim f As Range

val = InputBox("Enter Product Type")

val = UCase(val)

If val = "" Then Exit Sub

val = UCase(val) ' ? CAPITAL FIX

Set ws = Sheets("Product_Master")

Set f = ws.Columns(2).Find(val, LookAt:=xlWhole)

If Not f Is Nothing Then
MsgBox "Product already exists"
Exit Sub
End If

r = ws.Cells(ws.Rows.count, 2).End(xlUp).row + 1

ws.Cells(r, 2) = val

cmbProductType.AddItem val

cmbProductType.value = val

MsgBox "Product Added"

End Sub


'========================
' ADD COMPANY
'========================
Private Sub btnAddCompany_Click()
    Dim ws As Worksheet
    Dim val As String
    Dim productType As String
    Dim r As Long
    Dim lastRow As Long
    Dim i As Long
    Dim exists As Boolean

    val = InputBox("Enter Company Name")
    val = UCase(Trim(val))
    If val = "" Then Exit Sub

    productType = UCase(Trim(cmbProductType.value))
    If productType = "" Then
        MsgBox "???? Product Type ??????? ????!", vbExclamation
        Exit Sub
    End If

    Set ws = Sheets("Product_Master")
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    
    ' ??? ??? ?? ?? ????? ??? Product Type ??? ???? ?? ?? ?? ????
    exists = False
    For i = 2 To lastRow
        If UCase(Trim(ws.Cells(i, 2).value)) = productType And _
           UCase(Trim(ws.Cells(i, 4).value)) = val Then
            exists = True
            Exit For
        End If
    Next i

    If exists Then
        MsgBox "?? ????? ?? Product Type ??? ???? ?? ????? ??!", vbExclamation
        Exit Sub
    End If

    ' ??? ??????? ?? ??? (Product Type ?? ???)
    r = lastRow + 1
    ws.Cells(r, 1).value = r - 1  ' ID
    ws.Cells(r, 2).value = cmbProductType.value  ' Product Type
    ws.Cells(r, 4).value = val   ' Company
    
    ' Combo ??? Add ??? ?? ????? ??????? ?? ???
    cmbCompany.AddItem val
    cmbCompany.value = val  ' <-- ?? ????? (????? ??????? ????)
    
    MsgBox "Company Added: " & val, vbInformation
End Sub

'========================
' ADD MODEL
'========================
Private Sub btnAddModel_Click()
    Dim ws As Worksheet
    Dim val As String
    Dim productType As String
    Dim company As String
    Dim r As Long
    Dim lastRow As Long
    Dim i As Long
    Dim exists As Boolean

    val = InputBox("Enter Model Name")
    val = UCase(Trim(val))
    If val = "" Then Exit Sub

    productType = UCase(Trim(cmbProductType.value))
    company = UCase(Trim(cmbCompany.value))
    
    If productType = "" Then
        MsgBox "???? Product Type ??????? ????!", vbExclamation
        Exit Sub
    End If
    
    If company = "" Then
        MsgBox "???? Company ??????? ????!", vbExclamation
        Exit Sub
    End If

    Set ws = Sheets("Product_Master")
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    
    ' ??? ??? ?? ?? ???? ?? Product Type ?? Company ??? ???? ?? ?? ?? ????
    exists = False
    For i = 2 To lastRow
        If UCase(Trim(ws.Cells(i, 2).value)) = productType And _
           UCase(Trim(ws.Cells(i, 4).value)) = company And _
           UCase(Trim(ws.Cells(i, 5).value)) = val Then
            exists = True
            Exit For
        End If
    Next i

    If exists Then
        MsgBox "?? ???? ?? ????? ??? ???? ?? ????? ??!", vbExclamation
        Exit Sub
    End If

    ' ??? ??????? ?? ???
    r = lastRow + 1
    ws.Cells(r, 1).value = r - 1  ' ID
    ws.Cells(r, 2).value = cmbProductType.value  ' Product Type
    ws.Cells(r, 4).value = cmbCompany.value      ' Company
    ws.Cells(r, 5).value = val                   ' Model
    
    ' Combo ??? Add ??? ?? ????? ??????? ?? ???
    cmbModel.AddItem val
    cmbModel.value = val  ' <-- ?? ?????
    
    MsgBox "Model Added: " & val, vbInformation
End Sub

'========================
' LOAD PRODUCT TYPE (NO DUPLICATE)
'========================
Sub LoadProductType()
    Dim ws As Worksheet
    Dim r As Long
    Dim item As String
    Dim i As Integer
    Dim found As Boolean
    
    Set ws = Sheets("Product_Master")
    cmbProductType.Clear
    
    For r = 2 To ws.Cells(ws.Rows.count, 2).End(xlUp).row
        item = UCase(Trim(ws.Cells(r, 2).value))
        
        If item <> "" Then
            ' Check if already in combo
            found = False
            For i = 0 To cmbProductType.ListCount - 1
                If UCase(Trim(cmbProductType.List(i))) = item Then
                    found = True
                    Exit For
                End If
            Next i
            
            If Not found Then
                cmbProductType.AddItem item
            End If
        End If
    Next r
End Sub

'========================
' LOAD COMPANY (NO DUPLICATE)
'========================
Sub LoadCompany()
    Dim ws As Worksheet
    Dim r As Long
    Dim item As String
    Dim i As Integer
    Dim found As Boolean
    
    Set ws = Sheets("Product_Master")
    cmbCompany.Clear
    
    For r = 2 To ws.Cells(ws.Rows.count, 4).End(xlUp).row
        item = UCase(Trim(ws.Cells(r, 4).value))
        
        If item <> "" Then
            found = False
            For i = 0 To cmbCompany.ListCount - 1
                If UCase(Trim(cmbCompany.List(i))) = item Then
                    found = True
                    Exit For
                End If
            Next i
            
            If Not found Then
                cmbCompany.AddItem item
            End If
        End If
    Next r
End Sub

'========================
' LOAD MODEL (NO DUPLICATE)
'========================
Sub LoadModel()
    Dim ws As Worksheet
    Dim r As Long
    Dim item As String
    Dim i As Integer
    Dim found As Boolean
    
    Set ws = Sheets("Product_Master")
    cmbModel.Clear
    
    For r = 2 To ws.Cells(ws.Rows.count, 5).End(xlUp).row
        item = UCase(Trim(ws.Cells(r, 5).value))
        
        If item <> "" Then
            found = False
            For i = 0 To cmbModel.ListCount - 1
                If UCase(Trim(cmbModel.List(i))) = item Then
                    found = True
                    Exit For
                End If
            Next i
            
            If Not found Then
                cmbModel.AddItem item
            End If
        End If
    Next r
End Sub

'========================
' WARRANTY YEAR (NO DUPLICATE)
'========================
Sub LoadWarrantyYear()
    Dim ws As Worksheet
    Dim r As Long
    Dim item As String
    Dim i As Integer
    Dim found As Boolean
    
    Set ws = Sheets("Warranty_Master")
    cmbWarrantyYear.Clear
    
    For r = 2 To ws.Cells(ws.Rows.count, 1).End(xlUp).row
        item = Trim(ws.Cells(r, 1).value)
        
        If item <> "" Then
            found = False
            For i = 0 To cmbWarrantyYear.ListCount - 1
                If Trim(cmbWarrantyYear.List(i)) = item Then
                    found = True
                    Exit For
                End If
            Next i
            
            If Not found Then
                cmbWarrantyYear.AddItem item
            End If
        End If
    Next r
End Sub

Private Sub btnCancel_Click()

Unload Me

End Sub
Private Sub btnAddWarranty_Click()

Dim ws As Worksheet
Dim val As String
Dim r As Long
Dim f As Range

val = InputBox("Enter Warranty (Example: 3 Months / 1 Year)")

If val = "" Then Exit Sub

Set ws = Sheets("Warranty_Master")

Set f = ws.Columns(1).Find(val, LookAt:=xlWhole)

If Not f Is Nothing Then

MsgBox "Warranty already exists"
Exit Sub

End If

r = ws.Cells(ws.Rows.count, 1).End(xlUp).row + 1

ws.Cells(r, 1) = val

cmbWarrantyYear.AddItem val

MsgBox "Warranty Added"

End Sub
'========================
' AUTO FORMAT DATE INPUT
'========================
Private Sub txtPurchaseDate_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Trim(txtPurchaseDate.value) = "" Then Exit Sub
    
    Dim cleanDate As String
    Dim dt As Date
    
    cleanDate = txtPurchaseDate.value
    
    ' Remove all separators, keep only numbers
    Dim i As Long, ch As String, numOnly As String
    numOnly = ""
    For i = 1 To Len(cleanDate)
        ch = Mid(cleanDate, i, 1)
        If ch >= "0" And ch <= "9" Then numOnly = numOnly & ch
    Next i
    
    ' Try to parse based on length
    On Error Resume Next
    
    If Len(numOnly) = 8 Then
        ' DDMMYYYY format ? 01012024
        dt = DateSerial(CInt(Right(numOnly, 4)), CInt(Mid(numOnly, 3, 2)), CInt(Left(numOnly, 2)))
    ElseIf Len(numOnly) = 6 Then
        ' DDMMYY format ? 010124
        Dim yy As Integer
        yy = CInt(Right(numOnly, 2))
        If yy < 50 Then yy = 2000 + yy Else yy = 1900 + yy
        dt = DateSerial(yy, CInt(Mid(numOnly, 3, 2)), CInt(Left(numOnly, 2)))
    Else
        ' Try direct conversion
        dt = CDate(cleanDate)
    End If
    
    If Err.Number <> 0 Then
        MsgBox "Invalid Date! Please enter date as DD-MM-YYYY or DD.MM.YYYY", vbExclamation
        Cancel = True
        txtPurchaseDate.SetFocus
        On Error GoTo 0
        Exit Sub
    End If
    
    On Error GoTo 0
    
    ' Format to standard: dd-mm-yyyy
    txtPurchaseDate.value = Format(dt, "dd-mm-yyyy")
    
    ' Trigger warranty calculation
    CalculateWarranty
End Sub

' Alternative: KeyPress ?? Change event ?? add ????
Private Sub txtPurchaseDate_Change()
    ' Real-time calculation ?? ??? (optional)
    ' CalculateWarranty
End Sub

'========================
' WARRANTY YEAR COMBO
'========================
Private Sub cmbWarrantyYear_Change()
    ' ?? ????????? ???? ?? value select ???? ?? calculate ??
    If cmbWarrantyYear.value <> "" Then
        CalculateWarranty
    End If
End Sub

'========================
' WARRANTY CALCULATION (FIXED)
'========================
Sub CalculateWarranty()

    Dim pDate As Date
    Dim expDate As Date
    Dim remainDays As Long
    Dim warrantyText As String
    Dim warrantyMonths As Long
    Dim warrantyYears As Long
    
    ' Debug
    Debug.Print "Purchase Date: '" & txtPurchaseDate.value & "'"
    Debug.Print "Warranty Year: '" & cmbWarrantyYear.value & "'"

    If txtPurchaseDate.value = "" Or cmbWarrantyYear.value = "" Then
        lblWarrantyStatus.caption = "Please enter Purchase Date and select Warranty"
        lblWarrantyStatus.ForeColor = RGB(255, 0, 0)
        Exit Sub
    End If
    
    On Error GoTo stopCalc
    
    pDate = CDate(txtPurchaseDate.value)
    
    ' CLEAN: Remove spaces and convert to UPPER
    warrantyText = UCase(Trim(Replace(cmbWarrantyYear.value, " ", "")))
    
    ' ==========================================
    ' PARSE WARRANTY TEXT - ROBUST METHOD
    ' ==========================================
    warrantyMonths = 0
    
    ' Check for MONTHS first
    If InStr(warrantyText, "MONTH") > 0 Then
        ' Extract number before MONTH
        Dim monthVal As String
        monthVal = ""
        Dim i As Long
        For i = 1 To InStr(warrantyText, "MONTH") - 1
            If IsNumeric(Mid(warrantyText, i, 1)) Then
                monthVal = monthVal & Mid(warrantyText, i, 1)
            End If
        Next i
        
        If monthVal <> "" Then
            warrantyMonths = CLng(monthVal)
        End If
        
    ' Check for YEARS
    ElseIf InStr(warrantyText, "YEAR") > 0 Then
        Dim yearVal As String
        yearVal = ""
        For i = 1 To InStr(warrantyText, "YEAR") - 1
            If IsNumeric(Mid(warrantyText, i, 1)) Then
                yearVal = yearVal & Mid(warrantyText, i, 1)
            End If
        Next i
        
        If yearVal <> "" Then
            warrantyMonths = CLng(yearVal) * 12  ' Convert years to months
        End If
        
    ' Check for numeric only (just a number)
    ElseIf IsNumeric(warrantyText) Then
        warrantyMonths = CLng(warrantyText) * 12  ' Assume years if just number
        
    End If
    
    ' ==========================================
    ' CALCULATE EXPIRY DATE
    ' ==========================================
    If warrantyMonths > 0 Then
        expDate = DateAdd("m", warrantyMonths, pDate)
    Else
        lblWarrantyStatus.caption = "Invalid Warranty Format"
        lblWarrantyStatus.ForeColor = RGB(255, 0, 0)
        Exit Sub
    End If
    
    ' ==========================================
    ' DISPLAY STATUS
    ' ==========================================
    remainDays = DateDiff("d", Date, expDate)
    
    If remainDays >= 0 Then
        ' UNDER WARRANTY
        Dim totalDays As Long
        totalDays = DateDiff("d", pDate, expDate)
        
        Dim statusMsg As String
        statusMsg = remainDays & " Days Remaining"
        
        ' Add years/months info
        If warrantyMonths >= 12 Then
            Dim wy As Long, wm As Long
            wy = warrantyMonths \ 12
            wm = warrantyMonths Mod 12
            If wm = 0 Then
                statusMsg = statusMsg & " | " & wy & " Year" & IIf(wy > 1, "s", "") & " Warranty"
            Else
                statusMsg = statusMsg & " | " & wy & "Y " & wm & "M Warranty"
            End If
        Else
            statusMsg = statusMsg & " | " & warrantyMonths & " Month" & IIf(warrantyMonths > 1, "s", "") & " Warranty"
        End If
        
        statusMsg = statusMsg & " | Expiry: " & Format(expDate, "dd-mm-yyyy")
        
        lblWarrantyStatus.caption = statusMsg
        lblWarrantyStatus.ForeColor = RGB(0, 128, 0)  ' Green = Under Warranty
        
    Else
        ' EXPIRED
        lblWarrantyStatus.caption = "EXPIRED on " & Format(expDate, "dd-mm-yyyy") & " | " & Abs(remainDays) & " Days Over"
        lblWarrantyStatus.ForeColor = RGB(255, 0, 0)  ' Red = Expired
        
    End If
    
    Exit Sub

stopCalc:
    lblWarrantyStatus.caption = "Error: " & Err.Description
    lblWarrantyStatus.ForeColor = RGB(255, 0, 0)
    
End Sub
Sub CreateProductPhotoFolder()

Dim f As String

f = GetProductPhotoFolder()

If f = "" Then Exit Sub

If Dir(f, vbDirectory) = "" Then
MkDir f
End If

End Sub
Sub LoadAccessories()

Dim ws As Worksheet
Dim r As Long

Set ws = Sheets("Accessory_Master")

cmbAccessoryName.Clear

For r = 2 To ws.Cells(ws.Rows.count, 1).End(xlUp).row

If ws.Cells(r, 1).value = cmbProductType.value Then

cmbAccessoryName.AddItem ws.Cells(r, 2).value

End If

Next r

End Sub
Private Sub txtSerialNumber_Change()
    txtSerialNumber.value = UCase(txtSerialNumber.value)
    txtSerialNumber.SelStart = Len(txtSerialNumber.value)
End Sub

Private Sub txtProblem_Change()
    txtProblem.value = UCase(txtProblem.value)
    txtProblem.SelStart = Len(txtProblem.value)
End Sub

Private Sub txtRemark_Change()
    txtRemark.value = UCase(txtRemark.value)
    txtRemark.SelStart = Len(txtRemark.value)
End Sub

Sub LoadCapturedPhoto(ByVal selectedBox As Integer)

Dim latestPhoto As String

latestPhoto = GetLatestCameraPhoto()

If latestPhoto = "" Then
    MsgBox "Photo detect nahi hua", vbExclamation
    Exit Sub
End If

Call CreateProductPhotoFolder

Select Case selectedBox

Case 1
    FileCopy latestPhoto, GetPhotoFilePath(1)
    img1.Picture = LoadPicture(GetPhotoFilePath(1))

Case 2
    FileCopy latestPhoto, GetPhotoFilePath(2)
    img2.Picture = LoadPicture(GetPhotoFilePath(2))

Case 3
    FileCopy latestPhoto, GetPhotoFilePath(3)
    img3.Picture = LoadPicture(GetPhotoFilePath(3))

Case 4
    FileCopy latestPhoto, GetPhotoFilePath(4)
    img4.Picture = LoadPicture(GetPhotoFilePath(4))

End Select

' ? ADD THIS
AutoNextPhotoBox

End Sub
Sub AutoNextPhotoBox()

Select Case SelectedPhotoBox

    Case 1
        SelectedPhotoBox = 2
        img2_Click
        
    Case 2
        SelectedPhotoBox = 3
        img3_Click
        
    Case 3
        SelectedPhotoBox = 4
        img4_Click
        
    Case 4
        SelectedPhotoBox = 1
        img1_Click

End Select

End Sub
Sub ReplaceCapturedPhoto(ByVal selectedBox As Integer)

Dim latestPhoto As String
Dim targetPath As String

latestPhoto = GetLatestCameraPhoto()

If latestPhoto = "" Then
    MsgBox "Photo detect nahi hua", vbExclamation
    Exit Sub
End If

Call CreateProductPhotoFolder

Select Case selectedBox

Case 1
    targetPath = GetPhotoFilePath(1)
    FileCopy latestPhoto, targetPath
    img1.Picture = LoadPicture(targetPath)

Case 2
    targetPath = GetPhotoFilePath(2)
    FileCopy latestPhoto, targetPath
    img2.Picture = LoadPicture(targetPath)

Case 3
    targetPath = GetPhotoFilePath(3)
    FileCopy latestPhoto, targetPath
    img3.Picture = LoadPicture(targetPath)

Case 4
    targetPath = GetPhotoFilePath(4)
    FileCopy latestPhoto, targetPath
    img4.Picture = LoadPicture(targetPath)

End Select

DoEvents
Me.Repaint

End Sub
'=================================
' LOAD PRODUCT FOR EDIT (COMPLETE FIX)
'=================================
Public Sub LoadProductForEdit(ByVal entryID As String)
    Dim ws As Worksheet
    Dim r As Long
    Dim foundRow As Long
    Call SetupEditModeButtons
    Set ws = ThisWorkbook.Sheets("Job_Product")
    foundRow = 0
    
    ' Find EntryID
    For r = 2 To ws.Cells(ws.Rows.count, 1).End(xlUp).row
        If UCase(Trim(ws.Cells(r, 1).value)) = UCase(Trim(entryID)) Then
            foundRow = r
            Exit For
        End If
    Next r
    
    If foundRow = 0 Then
        MsgBox "Product not found: " & entryID, vbExclamation
        Exit Sub
    End If
    
    ' Set Edit Mode
    Me.Tag = entryID
    
    ' Load Entry Type ???? (???? Warranty UI ??? ??)
    Me.cmbEntryType.value = ws.Cells(foundRow, 3).value      ' C: EntryType
    
    ' === WARRANTY FIELDS SHOW/HIDE (IMPORTANT) ===
    If UCase(Trim(Me.cmbEntryType.value)) = "WARRANTY" Then
        Me.lblPurchaseDate.visible = True
        Me.txtPurchaseDate.visible = True
        Me.lblWarrantyYear.visible = True
        Me.cmbWarrantyYear.visible = True
        Me.lblWarrantyStatus.visible = True
        Me.btnAddWarranty.visible = True
    Else
        Me.lblPurchaseDate.visible = False
        Me.txtPurchaseDate.visible = False
        Me.lblWarrantyYear.visible = False
        Me.cmbWarrantyYear.visible = False
        Me.lblWarrantyStatus.visible = False
        Me.btnAddWarranty.visible = False
    End If
    ' ============================================
    
    ' ?? ???? Fields Load ??? (Sahi Columns ??)
    With ws.Rows(foundRow)
        Me.txtEntryID.caption = .Cells(1, 1).value      ' A: EntryID
        Me.cmbProductType.value = .Cells(1, 6).value    ' F: ProductType (???? 4 ??)
        Me.cmbCompany.value = .Cells(1, 7).value        ' G: Company (???? 5 ??)
        Me.cmbModel.value = .Cells(1, 8).value          ' H: Model (???? 6 ??)
        Me.txtSerialNumber.value = .Cells(1, 9).value   ' I: Serial (???? 7 ??)
        Me.txtPurchaseDate.value = .Cells(1, 10).value  ' J: PurchaseDate (???? 8 ??)
        Me.cmbWarrantyYear.value = .Cells(1, 11).value  ' K: WarrantyYear (???? 9 ??)
        Me.lblWarrantyStatus.caption = .Cells(1, 12).value ' L: WarrantyStatus (???? 10 ??)
        Me.txtProblem.value = .Cells(1, 13).value       ' M: Problem (???? 11 ??)
        Me.txtRemark.value = .Cells(1, 14).value        ' N: Remark (???? 12 ??)
    End With
    
    ' Load Photos
    Call LoadProductPhotos(entryID)
    
    ' === ACCESSORIES BUTTONS FIX ===
    Me.btnAccessories.visible = False      ' Add button Hide (??????? Edit Mode ??? ???? ?? Saved ??)
    Me.btnEditAccessories.visible = True   ' Edit button Show
    ' ================================
    
    ' Lock Entry Type
    Me.cmbEntryType.enabled = False
    
    ' Warranty Calculate ??? (???? Status ????? ??)
    If UCase(Trim(Me.cmbEntryType.value)) = "WARRANTY" Then
        Call CalculateWarranty
    End If
End Sub
'=================================
' LOAD PRODUCT PHOTOS (PERFECT - SINGLE COPY)
'=================================
Public Sub LoadProductPhotos(ByVal entryID As String)
    Dim basePath As String
    Dim folderPath As String
    Dim fso As Object
    
    On Error GoTo PhotoError
    
    ' Validate entryID
    If Trim(entryID) = "" Then Exit Sub
    
    ' Get path from settings or default
    basePath = Trim(Sheets("Settings").Range("B2").value)
    If basePath = "" Then
        basePath = ThisWorkbook.path & "\GLOBAL_SOFT_DATA\Images"
    End If
    
    folderPath = basePath & "\" & entryID & "\"
    
    ' Check if folder exists
    If Dir(folderPath, vbDirectory) = "" Then
        ' Try alternative path
        folderPath = ThisWorkbook.path & "\GLOBAL_SOFT_DATA\Images\" & entryID & "\"
        If Dir(folderPath, vbDirectory) = "" Then Exit Sub
    End If
    
    ' Load photos with error handling
    On Error Resume Next
    
    If Dir(folderPath & "Photo1.jpg") <> "" Then
        Set Me.img1.Picture = LoadPicture(folderPath & "Photo1.jpg")
    End If
    
    If Dir(folderPath & "Photo2.jpg") <> "" Then
        Set Me.img2.Picture = LoadPicture(folderPath & "Photo2.jpg")
    End If
    
    If Dir(folderPath & "Photo3.jpg") <> "" Then
        Set Me.img3.Picture = LoadPicture(folderPath & "Photo3.jpg")
    End If
    
    If Dir(folderPath & "Photo4.jpg") <> "" Then
        Set Me.img4.Picture = LoadPicture(folderPath & "Photo4.jpg")
    End If
    
    On Error GoTo 0
    Exit Sub
    
PhotoError:
    LogError "LoadProductPhotos Error: " & Err.Description & " EntryID: " & entryID
End Sub

Private Sub cmbProductType_Change()
    Dim ws As Worksheet
    Dim r As Long
    Dim i As Integer
    Dim found As Boolean
    Dim productType As String

    Set ws = Sheets("Product_Master")
    
    cmbCompany.Clear
    cmbModel.Clear
    
    productType = UCase(Trim(cmbProductType.value))
    If productType = "" Then Exit Sub

    ' ????? ??? Product Type ?? ???????? ??? ???
    For r = 2 To ws.Cells(ws.Rows.count, 2).End(xlUp).row
        If UCase(Trim(ws.Cells(r, 2).value)) = productType Then
            ' ??? ??? ?? ?? ????? ???? ?? Combo ??? ?? ?? ????
            found = False
            For i = 0 To cmbCompany.ListCount - 1
                If UCase(cmbCompany.List(i)) = UCase(Trim(ws.Cells(r, 4).value)) Then
                    found = True
                    Exit For
                End If
            Next i
            
            ' ??? ???? ?? ?? ???? ???? ?? ?? Add ???
            If Not found And Trim(ws.Cells(r, 4).value) <> "" Then
                cmbCompany.AddItem ws.Cells(r, 4).value
            End If
        End If
    Next r
End Sub
Private Sub cmbCompany_Change()
    Dim ws As Worksheet
    Dim r As Long
    Dim productType As String
    Dim company As String

    Set ws = Sheets("Product_Master")
    
    cmbModel.Clear
    
    productType = UCase(Trim(cmbProductType.value))
    company = UCase(Trim(cmbCompany.value))
    
    If productType = "" Or company = "" Then Exit Sub

    ' ????? ??? Product Type ?? Company ?? ???? ??? ???
    For r = 2 To ws.Cells(ws.Rows.count, 2).End(xlUp).row
        If UCase(Trim(ws.Cells(r, 2).value)) = productType And _
           UCase(Trim(ws.Cells(r, 4).value)) = company Then
            If Trim(ws.Cells(r, 5).value) <> "" Then
                cmbModel.AddItem ws.Cells(r, 5).value
            End If
        End If
    Next r
End Sub


Private Function EntryExists(checkID As String) As Boolean
    Dim ws As Worksheet
    Dim r As Long
    
    Set ws = ThisWorkbook.Sheets("Job_Product")
    
    For r = 2 To ws.Cells(ws.Rows.count, 1).End(xlUp).row
        If UCase(Trim(ws.Cells(r, 1).value)) = UCase(Trim(checkID)) Then
            EntryExists = True
            Exit Function
        End If
    Next r
    
    EntryExists = False
End Function
'=================================
' GENERATE ENTRY ID (FIXED)
'=================================
Public Function GeneratePerfectEntryID(entryType As String) As String
    Dim ws As Worksheet
    Dim lastRow As Long, r As Long
    Dim maxNum As Long, currentNum As Long
    Dim prefix As String
    Dim newID As String
    
    ' Set prefix based on entry type
    If UCase(Trim(entryType)) = "WARRANTY" Then
        prefix = "WAR"
    Else
        prefix = "SER"
    End If
    
    Set ws = ThisWorkbook.Sheets("Job_Product")
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    
    maxNum = 0
    
    ' Find max number for this prefix
    For r = 2 To lastRow
        If Left(ws.Cells(r, 1).value, 3) = prefix Then
            currentNum = val(Mid(ws.Cells(r, 1).value, 4))
            If currentNum > maxNum Then maxNum = currentNum
        End If
    Next r
    
    ' Generate new ID
    newID = prefix & Format(maxNum + 1, "00000")
    
    ' Double check if already exists (add timestamp if needed)
    If EntryExists(newID) Then
        newID = prefix & Format(maxNum + 1, "00000") & "_" & Format(Now, "hhmmss")
    End If
    
    GeneratePerfectEntryID = newID
End Function

Private Sub SetupEditModeButtons()
    ' EDIT MODE - Existing Product ke liye
    Me.btnCapture.visible = False           ' Capture hide (kyunki pehle se photo hai)
    Me.btnDeletePhoto.visible = False       ' Delete Photo hide (Update mode mein alag logic hoga)
    Me.btnUpdate.visible = True             ' Update dikhna chahiye
    Me.btnAddPhoto.visible = True           ' Add Photo dikhna chahiye
    
    ' Accessories buttons
    Me.btnAccessories.visible = False       ' Add Accessories hide
    Me.btnEditAccessories.visible = True    ' Edit Accessories show
End Sub
