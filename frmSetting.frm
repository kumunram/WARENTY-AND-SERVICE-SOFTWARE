VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSetting 
   Caption         =   "Setting"
   ClientHeight    =   9135.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   19980
   OleObjectBlob   =   "frmSetting.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit




Private Sub btnRepairSystem_Click()
If MsgBox("This will repair database and restart software." & vbCrLf & _
              "Continue?", vbYesNo + vbQuestion) = vbNo Then Exit Sub
    
    ' modSystemValidator ?? function call ????
    Call modSetupValidator.EmergencyRepairWithRestart
End Sub

'=================================
' LOAD SETTINGS
'=================================
Private Sub UserForm_Initialize()

Dim ws As Worksheet
Set ws = Sheets("Settings")

txtImagePath.value = ws.Range("B1").value
txtProductPhotoPath.value = ws.Range("B2").value
txtCustomerPhotoPath.value = ws.Range("B3").value
txtReportPath.value = ws.Range("B4").value
txtBackupPath.value = ws.Range("B5").value
txtAccessoryPhotoPath.value = ws.Range("B6").value
txtLogoPath.value = ws.Range("B11").value
' =========================================
    ' NEW CODE - PAGE SETUP (?? ??????)
    ' =========================================
    
    ' Page Size (B10)
    Me.cmbPageSize.Clear
    Me.cmbPageSize.AddItem "A4"
    Me.cmbPageSize.AddItem "A5"
    Me.cmbPageSize.value = ws.Range("B20").value & ""
    If Me.cmbPageSize.value = "" Then Me.cmbPageSize.value = "A4"
    
    ' Show Pictures (B11, B12, B13)
    Me.cmbShowCustomerPic.Clear
    Me.cmbShowCustomerPic.AddItem "YES"
    Me.cmbShowCustomerPic.AddItem "NO"
    Me.cmbShowCustomerPic.value = ws.Range("B21").value & ""
    If Me.cmbShowCustomerPic.value = "" Then Me.cmbShowCustomerPic.value = "NO"
    
    Me.cmbShowProductPic.Clear
    Me.cmbShowProductPic.AddItem "YES"
    Me.cmbShowProductPic.AddItem "NO"
    Me.cmbShowProductPic.value = ws.Range("B22").value & ""
    If Me.cmbShowProductPic.value = "" Then Me.cmbShowProductPic.value = "NO"
    
    Me.cmbShowAccessoryPic.Clear
    Me.cmbShowAccessoryPic.AddItem "YES"
    Me.cmbShowAccessoryPic.AddItem "NO"
    Me.cmbShowAccessoryPic.value = ws.Range("B23").value & ""
    If Me.cmbShowAccessoryPic.value = "" Then Me.cmbShowAccessoryPic.value = "NO"
    
    ' Terms & Conditions Load (B15:B22 ?? TextBox ???)
    Dim i As Long, termText As String
    termText = ""
    For i = 25 To 35
        If Trim(ws.Cells(i, 2).value & "") <> "" Then
            termText = termText & ws.Cells(i, 2).value & vbCrLf
        End If
    Next i
    Me.txtTermsConditions.value = termText

End Sub


Private Sub btnCreateAllPaths_Click()
    
    On Error Resume Next
    Call btnCreateImagePath_Click
    Call btnCreateProductPhotoPath_Click
    Call btnCreateCustomerPhotoPath_Click
    Call btnCreateAccessoryPhotoPath_Click
    Call btnCreateReportPath_Click
    Call btnCreateBackupPath_Click
    Call btnCreateLogoPath_Click
    On Error GoTo 0
    MsgBox "All folders created successfully", vbInformation
End Sub


'=================================
' SAVE SETTINGS
'=================================
Private Sub btnSaveSettings_Click()

Dim ws As Worksheet
Set ws = Sheets("Settings")

ws.Range("B1").value = txtImagePath.value
ws.Range("B2").value = txtProductPhotoPath.value
ws.Range("B3").value = txtCustomerPhotoPath.value
ws.Range("B4").value = txtReportPath.value
ws.Range("B5").value = txtBackupPath.value
ws.Range("B6").value = txtAccessoryPhotoPath.value
ws.Range("B11").value = txtLogoPath.value
' =========================================
    ' NEW CODE - SAVE PAGE SETUP (?? ??????)
    ' =========================================
    
    ' Save Page Size (B10)
    ws.Range("B20").value = Me.cmbPageSize.value
    
    ' Save Picture Options (B11, B12, B13)
    ws.Range("B21").value = Me.cmbShowCustomerPic.value
    ws.Range("B22").value = Me.cmbShowProductPic.value
    ws.Range("B23").value = Me.cmbShowAccessoryPic.value
    
    ' Save Terms & Conditions (TextBox ?? B15:B22 ??? Split ???? ???)
    Dim lines() As String
    Dim i As Long
    lines = Split(Me.txtTermsConditions.value, vbCrLf)
    
    ' ???? ?????? ???? ?????
    For i = 25 To 35
        ws.Cells(i, 2).ClearContents
    Next i
    
    ' ??? ???? ????? (Max 8 ????)
    For i = 0 To UBound(lines)
        If i < 8 Then  ' B15 ?? B22 ?? (8 ????)
            ws.Cells(25 + i, 2).value = Trim(lines(i))
        End If
    Next i

MsgBox "Settings Saved Successfully", vbInformation

End Sub


'=================================
' CREATE IMAGE PATH
'=================================
Private Sub btnCreateImagePath_Click()

Dim p As String

p = ThisWorkbook.path & "\GLOBAL_SOFT_DATA\Images"

If Dir(p, vbDirectory) = "" Then MkDir p

txtImagePath.value = p
Sheets("Settings").Range("B1").value = p

MsgBox "Image folder created", vbInformation

End Sub


'=================================
' PRODUCT PHOTO PATH
'=================================
Private Sub btnCreateProductPhotoPath_Click()

Dim p As String

p = txtImagePath.value & "\Products"

If Dir(p, vbDirectory) = "" Then MkDir p

txtProductPhotoPath.value = p
Sheets("Settings").Range("B2").value = p

MsgBox "Product photo folder created", vbInformation

End Sub


'=================================
' CUSTOMER PHOTO PATH
'=================================
Private Sub btnCreateCustomerPhotoPath_Click()

Dim p As String

p = txtImagePath.value & "\Customers"

If Dir(p, vbDirectory) = "" Then MkDir p

txtCustomerPhotoPath.value = p
Sheets("Settings").Range("B3").value = p

MsgBox "Customer photo folder created", vbInformation

End Sub


'=================================
' ACCESSORY PHOTO PATH
'=================================
Private Sub btnCreateAccessoryPhotoPath_Click()

Dim p As String

p = txtImagePath.value & "\Accessories"

If Dir(p, vbDirectory) = "" Then MkDir p

txtAccessoryPhotoPath.value = p
Sheets("Settings").Range("B6").value = p

MsgBox "Accessory photo folder created", vbInformation

End Sub


'=================================
' REPORT PATH
'=================================
Private Sub btnCreateReportPath_Click()

Dim p As String

p = ThisWorkbook.path & "\GLOBAL_SOFT_DATA\Reports"

If Dir(p, vbDirectory) = "" Then MkDir p

txtReportPath.value = p
Sheets("Settings").Range("B4").value = p

MsgBox "Report folder created", vbInformation

End Sub


'=================================
' BACKUP PATH
'=================================
Private Sub btnCreateBackupPath_Click()

Dim p As String

p = ThisWorkbook.path & "\GLOBAL_SOFT_DATA\Backup"

If Dir(p, vbDirectory) = "" Then MkDir p

txtBackupPath.value = p
Sheets("Settings").Range("B5").value = p

MsgBox "Backup folder created", vbInformation

End Sub


'=================================
' TEST CAMERA
'=================================
Private Sub btnTestCamera_Click()

On Error GoTo camErr

shell "cmd /c start microsoft.windows.camera:", vbHide

Sheets("Settings").Range("B7").value = "OK"

MsgBox "Camera Working", vbInformation

Exit Sub

camErr:

Sheets("Settings").Range("B7").value = "NOT FOUND"

MsgBox "Camera Not Found", vbCritical

End Sub


'=================================
' TEST DATABASE
'=================================
Private Sub btnTestDatabase_Click()

On Error GoTo dbErr

Sheets("Customer_Master").Activate
Sheets("Product_Master").Activate
Sheets("Accessory_Master").Activate
Sheets("Job_Master").Activate
Sheets("Job_Product").Activate
Sheets("Job_Accessory").Activate

Sheets("Settings").Range("B8").value = "OK"

MsgBox "Database OK", vbInformation

Exit Sub

dbErr:

Sheets("Settings").Range("B8").value = "ERROR"

MsgBox "Database Error Found", vbCritical

End Sub


'=================================
' TEST BACKUP PATH
'=================================
Private Sub btnTestBackup_Click()

Dim path As String

path = Trim(txtBackupPath.value)

If path = "" Then

MsgBox "Backup Path Not Set", vbExclamation
Sheets("Settings").Range("B10").value = "FAIL"

Exit Sub

End If


If Dir(path, vbDirectory) <> "" Then

MsgBox "Backup Path OK", vbInformation
Sheets("Settings").Range("B10").value = "OK"

Else

MsgBox "Backup Folder Not Found", vbCritical
Sheets("Settings").Range("B10").value = "FAIL"

End If

End Sub


'=================================
' TEST VERSION
'=================================
Private Sub btnTestVersion_Click()

Dim ver As String

ver = Sheets("Settings").Range("B9").value

If ver = "v1.0" Then

MsgBox "Software Version OK", vbInformation

Else

MsgBox "Version Mismatch", vbCritical

End If

End Sub
'=================================
' CHECK DATABASE (Repair Function)
'=================================
Private Sub btnRepairDatabase_Click()
    On Error Resume Next
    
    Dim requiredSheets As Variant
    Dim s As Variant
    Dim ws As Worksheet
    Dim found As Boolean
    Dim msg As String
    
    requiredSheets = Array("Customer_Master", "Product_Master", "Accessory_Master", _
                          "Job_Master", "Job_Product", "Job_Accessory", _
                          "Payment_Master", "Engineer_Register", "Vendor_Register", _
                          "Settings", "Warranty_Master", "Verify_Master", "Status_Log")
    
    msg = "Database Repair Report:" & vbCrLf & vbCrLf
    
    For Each s In requiredSheets
        found = False
        For Each ws In ThisWorkbook.Sheets
            If ws.name = s Then
                found = True
                Exit For
            End If
        Next ws
        
        If Not found Then
            Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
            ws.name = s
            msg = msg & "? Created: " & s & vbCrLf
        Else
            msg = msg & "? OK: " & s & vbCrLf
        End If
    Next s
    
    ' Ensure Folders exist
    Call EnsureAllFolders
    
    MsgBox msg, vbInformation, "Database Repair Complete"
    
    On Error GoTo 0
End Sub

'=================================
' ENSURE ALL FOLDERS
'=================================
Private Sub EnsureAllFolders()
    Dim fso As Object
    Dim basePath As String
    Dim folders As Variant
    Dim f As Variant
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    basePath = ThisWorkbook.path & "\GLOBAL_SOFT_DATA\"
    
    If Not fso.FolderExists(basePath) Then
        fso.CreateFolder basePath
    End If
    
    folders = Array("Images", "Images\Customers", "Images\Products", "Images\Accessories", _
                   "Reports", "AutoBackup", "Temp")
    
    For Each f In folders
        If Not fso.FolderExists(basePath & f) Then
            fso.CreateFolder basePath & f
        End If
    Next f
    
    Set fso = Nothing
End Sub


'=================================
' CLOSE FORM
'=================================
Private Sub btnClose_Click()

Unload Me

End Sub
Private Sub btnCreateLogoPath_Click()

Dim p As String

' Main Image Path ?? Logo folder ?????
p = txtImagePath.value & "\Logo"

If Dir(p, vbDirectory) = "" Then MkDir p

txtLogoPath.value = p
Sheets("Settings").Range("B11").value = p

MsgBox "Logo folder created successfully", vbInformation

End Sub
