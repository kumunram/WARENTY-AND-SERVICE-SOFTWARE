VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAccessories 
   Caption         =   "Product Accessories"
   ClientHeight    =   12375
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15495
   OleObjectBlob   =   "frmAccessories.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmAccessories"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public customerID As String
Dim RowIndex As Long
Dim isUpdating As Boolean


'=================================
' FORM LOAD
'=================================
Private Sub UserForm_Initialize()
    On Error Resume Next
    
    AddMinMaxButtons Me
    
    Dim i As Integer
    RowIndex = 1
    isUpdating = False
    
    ' Row 1 ?? ??? dropdowns initialize ???
    With cmbStatus1
        .Clear
        .AddItem "RECEIVED"
        .AddItem "NOT RECEIVED"
        .value = "NOT RECEIVED"
    End With
    
    With cmbWorking1
        .Clear
        .AddItem "WORKING"
        .AddItem "NOT WORKING"
        .value = ""
        .enabled = False
    End With
    
    txtSerial1.enabled = False
    btnCapture1.enabled = False
    
    ' ???? rows hide ???
    For i = 2 To 10
        On Error Resume Next
        Me.Controls("cmbAccessory" & i).visible = False
        Me.Controls("cmbStatus" & i).visible = False
        Me.Controls("txtSerial" & i).visible = False
        Me.Controls("cmbWorking" & i).visible = False
        Me.Controls("imgPhoto" & i).visible = False
        Me.Controls("btnCapture" & i).visible = False
        If ControlExists("lblAcc" & i) Then Me.Controls("lblAcc" & i).visible = False
        On Error GoTo 0
    Next i
    
    ' ???? ??? accessory names load ???
    Call LoadAccessoryNamesForRow(1)
    
    On Error GoTo 0
End Sub

' Helper Function: Check if control exists
Private Function ControlExists(ctrlName As String) As Boolean
    Dim ctrl As Control
    On Error Resume Next
    Set ctrl = Me.Controls(ctrlName)
    ControlExists = Not ctrl Is Nothing
    On Error GoTo 0
End Function

'=================================
' ADD NEW ACCESSORY ROW
'=================================
Private Sub btnAddNewAccessory_Click()
    On Error GoTo ErrorHandler
    
    RowIndex = RowIndex + 1
    
    If RowIndex >= 2 And RowIndex <= 10 Then
        Dim r As Integer
        r = RowIndex
        
        On Error Resume Next
        
        ' 1. Accessory Combo ?????
        If ControlExists("cmbAccessory" & r) Then
            Me.Controls("cmbAccessory" & r).visible = True
            Me.Controls("cmbAccessory" & r).value = ""
            Call LoadAccessoryNamesForRow(r)
        End If
        
        ' 2. Status Combo ????? (RECEIVED/NOT RECEIVED)
        If ControlExists("cmbStatus" & r) Then
            With Me.Controls("cmbStatus" & r)
                .visible = True
                .Clear
                .AddItem "RECEIVED"
                .AddItem "NOT RECEIVED"
                .value = "NOT RECEIVED"
            End With
        End If
        
        ' 3. Serial TextBox ????? (?????? ?????, Status = RECEIVED ???? ?? ????? ????)
        If ControlExists("txtSerial" & r) Then
            Me.Controls("txtSerial" & r).visible = True
            Me.Controls("txtSerial" & r).value = ""
            Me.Controls("txtSerial" & r).enabled = False
        End If
        
        ' 4. Working Combo ????? (WORKING/NOT WORKING)
        If ControlExists("cmbWorking" & r) Then
            With Me.Controls("cmbWorking" & r)
                .visible = True
                .Clear
                .AddItem "WORKING"
                .AddItem "NOT WORKING"
                .value = ""
                .enabled = False  ' Status = RECEIVED ???? ?? ????? ????
            End With
        End If
        
        ' 5. Photo Image ?????
        If ControlExists("imgPhoto" & r) Then
            Me.Controls("imgPhoto" & r).visible = True
            Set Me.Controls("imgPhoto" & r).Picture = Nothing
        End If
        
        ' 6. Capture Button ?????
        If ControlExists("btnCapture" & r) Then
            Me.Controls("btnCapture" & r).visible = True
            Me.Controls("btnCapture" & r).enabled = False  ' Status = RECEIVED ???? ?? ????? ????
        End If
        
        ' 7. Label ??? ?? ?? ?????
        If ControlExists("lblAcc" & r) Then
            Me.Controls("lblAcc" & r).visible = True
        End If
        
        On Error GoTo 0
        
    Else
        MsgBox "Maximum 10 accessories allowed!", vbExclamation
        RowIndex = 10  ' 10 ?? ?????? ???? ???? ??
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error adding row: " & Err.Description, vbCritical
End Sub

'=================================
' LOAD ACCESSORY NAMES (ALL ROWS)
'=================================
Sub LoadAccessoryNames()
    Dim i As Integer
    For i = 1 To 10
        ' ????? ?????? rows ?? ??? ??? ??? ?? visible ??? ?? ???? row ??
        If i = 1 Or (ControlExists("cmbAccessory" & i) And Me.Controls("cmbAccessory" & i).visible = True) Then
            Call LoadAccessoryNamesForRow(i)
        End If
    Next i
End Sub

'=================================
' LOAD ACCESSORY NAMES FOR ROW
'=================================
Sub LoadAccessoryNamesForRow(rowNum As Integer)
    Dim ws As Worksheet
    Dim r As Long
    Dim productType As String
    Dim cmb As ComboBox
    
    On Error Resume Next
    
    Set ws = ThisWorkbook.Sheets("Accessory_Master")
    Set cmb = Me.Controls("cmbAccessory" & rowNum)
    
    If cmb Is Nothing Then Exit Sub
    
    ' ???? Clear ???
    cmb.Clear
    
    ' Product Type Check
    If Me.cmbProduct.value = "" Then
        ' ??? Product ???? ?? ?? ??? Accessory ?? ?????
        Exit Sub
    End If
    
    productType = UCase(Trim(Me.cmbProduct.value))
    
    ' Accessory_Master ?? Load ??? (Column B = ProductType, Column C = AccessoryName)
    For r = 2 To ws.Cells(ws.Rows.count, 2).End(xlUp).row
        If UCase(Trim(ws.Cells(r, 2).value)) = productType Then
            Dim accName As String
            accName = Trim(ws.Cells(r, 3).value)
            
            If accName <> "" Then
                ' Duplicate Check
                Dim alreadyExists As Boolean
                alreadyExists = False
                
                Dim i As Integer
                For i = 1 To 10
                    If i <> rowNum Then
                        If Me.Controls("cmbAccessory" & i).value = accName Then
                            alreadyExists = True
                            Exit For
                        End If
                    End If
                Next i
                
                If Not alreadyExists Then
                    cmb.AddItem accName
                End If
            End If
        End If
    Next r
    
    On Error GoTo 0
End Sub

'=================================
' CHECK IF ACCESSORY ALREADY USED
'=================================
Function IsAccessoryUsed(accName As String, excludeRow As Integer) As Boolean
    Dim i As Integer
    
    For i = 1 To 10
        If i <> excludeRow Then
            If Me.Controls("cmbAccessory" & i).value = accName Then
                IsAccessoryUsed = True
                Exit Function
            End If
        End If
    Next i
    
    IsAccessoryUsed = False
End Function

'=================================
' ADD NEW ACCESSORY TO MASTER
'=================================
Private Sub btnAddAccessoryName_Click()
    Dim ws As Worksheet
    Dim val As String
    Dim r As Long
    
    val = InputBox("Enter Accessory Name")
    
    If Trim(val) = "" Then Exit Sub
    
    val = UCase(Trim(val))  ' Trim ?? ???
    
    ' Check if product selected
    If Trim(Me.cmbProduct.value) = "" Then
        MsgBox "Please select Product first!", vbExclamation
        Exit Sub
    End If
    
    Set ws = Sheets("Accessory_Master")
    
    ' Check duplicate - Case insensitive search
    Dim f As Range
    Set f = ws.Columns(3).Find(What:=val, LookAt:=xlWhole, MatchCase:=False)  ' MatchCase:=False ?????
    
    If Not f Is Nothing Then
        If UCase(Trim(ws.Cells(f.row, 2).value)) = UCase(Trim(Me.cmbProduct.value)) Then
            MsgBox "Accessory already exists for this product!", vbExclamation
            Exit Sub
        End If
    End If
    
    ' Add new
    r = ws.Cells(ws.Rows.count, 1).End(xlUp).row + 1
    ws.Cells(r, 1).value = r - 1
    ws.Cells(r, 2).value = Me.cmbProduct.value
    ws.Cells(r, 3).value = val
    ws.Cells(r, 4).value = 1
    ws.Cells(r, 5).value = Now
    
    ' Reload - ??? visible rows ?? ???
    Call LoadAccessoryNames
    
    MsgBox "Accessory Added: " & val, vbInformation
End Sub

'=================================
' STATUS CHANGE HANDLER (ALL ROWS)
'=================================
Private Sub cmbStatus1_Change(): HandleStatus 1: End Sub
Private Sub cmbStatus2_Change(): HandleStatus 2: End Sub
Private Sub cmbStatus3_Change(): HandleStatus 3: End Sub
Private Sub cmbStatus4_Change(): HandleStatus 4: End Sub
Private Sub cmbStatus5_Change(): HandleStatus 5: End Sub
Private Sub cmbStatus6_Change(): HandleStatus 6: End Sub
Private Sub cmbStatus7_Change(): HandleStatus 7: End Sub
Private Sub cmbStatus8_Change(): HandleStatus 8: End Sub
Private Sub cmbStatus9_Change(): HandleStatus 9: End Sub
Private Sub cmbStatus10_Change(): HandleStatus 10: End Sub

Sub HandleStatus(i As Integer)
    Dim status As String
    status = Me.Controls("cmbStatus" & i).value
    
    If status = "RECEIVED" Then
        Me.Controls("txtSerial" & i).enabled = True
        Me.Controls("cmbWorking" & i).enabled = True
        Me.Controls("btnCapture" & i).enabled = True
    Else
        Me.Controls("txtSerial" & i).value = ""
        Me.Controls("cmbWorking" & i).value = ""
        Set Me.Controls("imgPhoto" & i).Picture = Nothing
        
        Me.Controls("txtSerial" & i).enabled = False
        Me.Controls("cmbWorking" & i).enabled = False
        Me.Controls("btnCapture" & i).enabled = False
    End If
End Sub
Private Sub btnSave_Click()
    Dim ws As Worksheet
    Dim i As Integer
    Dim r As Long
    Dim savedCount As Integer
    
    ' Validation
    If Trim(Me.txtEntryID.value) = "" Then
        MsgBox "Entry ID Missing!", vbCritical
        Exit Sub
    End If
    
    If Trim(Me.customerID) = "" Then
        MsgBox "Customer ID Missing!", vbCritical
        Exit Sub
    End If
    
    Set ws = Sheets("Job_Accessory")
    
    ' DELETE OLD DATA (if edit mode)
    If Me.Tag <> "" Then
        Call DeleteOldAccessories(Me.txtEntryID.value)
    End If
    
    ' SAVE NEW DATA
    savedCount = 0
    
    For i = 1 To 10
        Dim acc As String
        Dim ser As String
        Dim work As String
        Dim status As String
        
        acc = Me.Controls("cmbAccessory" & i).value
        ser = Me.Controls("txtSerial" & i).value
        work = Me.Controls("cmbWorking" & i).value
        status = Me.Controls("cmbStatus" & i).value
        
        If acc <> "" Then
            r = ws.Cells(ws.Rows.count, 1).End(xlUp).row + 1
            
            ws.Cells(r, 1).value = Me.txtEntryID.value
            ws.Cells(r, 2).value = Me.customerID
            ws.Cells(r, 3).value = Me.cmbProduct.value
            ws.Cells(r, 4).value = acc
            ws.Cells(r, 5).value = status
            ws.Cells(r, 6).value = ser
            ws.Cells(r, 7).value = work
            ws.Cells(r, 8).value = ""
            ws.Cells(r, 9).value = Now
            
            savedCount = savedCount + 1
        End If
    Next i
    
    ' Save photos
    Call SaveAccessoryPhotos(Me.txtEntryID.value)
    
    ' ============================================================
    ' ?? CRITICAL FIX: SAVE WORKBOOK AFTER ACCESSORIES
    ' ============================================================
    On Error GoTo SaveFailed
    
    Application.DisplayAlerts = False
    ThisWorkbook.Save
    Application.DisplayAlerts = True
    
    On Error GoTo 0
    ' ============================================================
    
    ' Success message
    If savedCount > 0 Then
        MsgBox savedCount & " Accessories Saved!" & vbCrLf & "File Auto-Saved", vbInformation
    Else
        MsgBox "No accessories saved (select RECEIVED status to save)", vbInformation
    End If
    
    Unload Me
    Exit Sub

SaveFailed:
    Application.DisplayAlerts = True
    MsgBox "CRITICAL ERROR: Accessories data written but FILE NOT SAVED!" & vbCrLf & _
           "Error: " & Err.Description & vbCrLf & vbCrLf & _
           "Please press Ctrl+S manually immediately!", vbCritical, "SAVE FAILED"
End Sub

'=================================
' DELETE OLD ACCESSORIES
'=================================
Sub DeleteOldAccessories(entryID As String)
    Dim ws As Worksheet
    Dim r As Long
    
    Set ws = Sheets("Job_Accessory")
    
    For r = ws.Cells(ws.Rows.count, 1).End(xlUp).row To 2 Step -1
        If ws.Cells(r, 1).value = entryID Then
            ws.Rows(r).Delete
        End If
    Next r
    
    ' ? YAHAN SAVE NAHI CHAHIYE — btnSave_Click mein hoga
End Sub

'=================================
' SAVE ACCESSORY PHOTOS (Settings ?? ????? ??)
'=================================
Sub SaveAccessoryPhotos(entryID As String)
    Dim basePath As String
    Dim folderPath As String
    Dim fso As Object
    Dim i As Integer
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Settings ?? ??? ??? (B6), ??? ???? ?? ?? Default ???
    basePath = Trim(Sheets("Settings").Range("B6").value)
    If basePath = "" Then
        basePath = ThisWorkbook.path & "\GLOBAL_SOFT_DATA\Images\Accessories"
    End If
    
    ' \ ???? ?? ?? ??????
    If Right(basePath, 1) <> "\" Then basePath = basePath & "\"
    
    folderPath = basePath & entryID & "\"
    
    ' ?????? ?????
    If Not fso.FolderExists(basePath) Then
        fso.CreateFolder basePath
    End If
    
    If Not fso.FolderExists(folderPath) Then
        fso.CreateFolder folderPath
    End If
    
    ' ???? ??? ????
    On Error Resume Next
    
    For i = 1 To 10
        Dim img As MSForms.Image
        Set img = Me.Controls("imgPhoto" & i)
        
        If Not img.Picture Is Nothing Then
            Dim accName As String
            accName = Me.Controls("cmbAccessory" & i).value
            
            If accName <> "" Then
                Dim fileName As String
                fileName = "Acc" & i & "_" & CleanFileName(accName) & ".jpg"
                SavePicture img.Picture, folderPath & fileName
            End If
        End If
    Next i
    
    On Error GoTo 0
End Sub

'=================================
' CLEAN FILENAME (remove special chars)
'=================================
Function CleanFileName(fileName As String) As String
    Dim result As String
    Dim invalidChars As String
    Dim i As Integer
    
    invalidChars = "\/:*?""<>|"
    result = fileName
    
    For i = 1 To Len(invalidChars)
        result = Replace(result, Mid(invalidChars, i, 1), "_")
    Next i
    
    CleanFileName = Left(result, 50) ' Max 50 chars
End Function

'=================================
' LOAD ACCESSORIES FOR EDIT (CORRECTED)
'=================================
Sub LoadAccessoriesForEdit(entryID As String)
    Dim ws As Worksheet
    Dim r As Long
    Dim i As Integer
    
    On Error Resume Next
    
    Set ws = Sheets("Job_Accessory")
    
    ' Clear existing
    Call ClearAllRows
    
    ' ===== ROW 1 DROPDOWNS REPOPULATE (ClearAllRows ?? clear ?? ???? ??) =====
    With cmbStatus1
        .Clear
        .AddItem "RECEIVED"
        .AddItem "NOT RECEIVED"
    End With
    With cmbWorking1
        .Clear
        .AddItem "WORKING"
        .AddItem "NOT WORKING"
    End With
    ' ===================================================================
    
    
    ' Load data
    i = 1
    For r = 2 To ws.Cells(ws.Rows.count, 1).End(xlUp).row
        If ws.Cells(r, 1).value = entryID Then
            ' ???? ???? (i=1) ????? ???? ???? ??, ???? (i>1) ?? ?????? ??????
                  ' ??? rows ?? ??? visible ?? dropdown populate ??? (i=1 ?? include)
            RowIndex = i
            
            If ControlExists("cmbAccessory" & i) Then Me.Controls("cmbAccessory" & i).visible = True
            If ControlExists("cmbStatus" & i) Then Me.Controls("cmbStatus" & i).visible = True
            If ControlExists("txtSerial" & i) Then Me.Controls("txtSerial" & i).visible = True
            If ControlExists("cmbWorking" & i) Then Me.Controls("cmbWorking" & i).visible = True
            If ControlExists("imgPhoto" & i) Then Me.Controls("imgPhoto" & i).visible = True
            If ControlExists("btnCapture" & i) Then Me.Controls("btnCapture" & i).visible = True
            If ControlExists("lblAcc" & i) Then Me.Controls("lblAcc" & i).visible = True
            
            ' ??? rows ?? ??? Status ?? Working Combo populate ???
            If ControlExists("cmbStatus" & i) Then
                With Me.Controls("cmbStatus" & i)
                    .Clear
                    .AddItem "RECEIVED"
                    .AddItem "NOT RECEIVED"
                End With
            End If
            
            If ControlExists("cmbWorking" & i) Then
                With Me.Controls("cmbWorking" & i)
                    .Clear
                    .AddItem "WORKING"
                    .AddItem "NOT WORKING"
                End With
            End If
            
            ' ???? Accessories ??? ??? ??? Value ??? ???
            Call LoadAccessoryNamesForRow(i)
            
            ' Values ??? ???
            If ControlExists("cmbAccessory" & i) Then Me.Controls("cmbAccessory" & i).value = ws.Cells(r, 4).value
            If ControlExists("cmbStatus" & i) Then Me.Controls("cmbStatus" & i).value = ws.Cells(r, 5).value
            If ControlExists("txtSerial" & i) Then Me.Controls("txtSerial" & i).value = ws.Cells(r, 6).value
            If ControlExists("cmbWorking" & i) Then Me.Controls("cmbWorking" & i).value = ws.Cells(r, 7).value
            
            ' ??? RECEIVED ?? ?? ????? ???
            If ws.Cells(r, 5).value = "RECEIVED" Then
                If ControlExists("txtSerial" & i) Then Me.Controls("txtSerial" & i).enabled = True
                If ControlExists("cmbWorking" & i) Then Me.Controls("cmbWorking" & i).enabled = True
                If ControlExists("btnCapture" & i) Then Me.Controls("btnCapture" & i).enabled = True
            Else
                If ControlExists("txtSerial" & i) Then Me.Controls("txtSerial" & i).enabled = False
                If ControlExists("cmbWorking" & i) Then Me.Controls("cmbWorking" & i).enabled = False
                If ControlExists("btnCapture" & i) Then Me.Controls("btnCapture" & i).enabled = False
            End If
            
            i = i + 1
            If i > 10 Then Exit For
        End If
    Next r
    
    RowIndex = i - 1  ' ???? ??? RowIndex ??? ????
    Call LoadAccessoryPhotos(entryID)
    On Error GoTo 0
End Sub
'=================================
' LOAD ACCESSORY PHOTOS
'=================================
Sub LoadAccessoryPhotos(entryID As String)
    Dim folderPath As String
    Dim fso As Object
    Dim folder As Object
    Dim file As Object
    Dim i As Integer
    
    folderPath = ThisWorkbook.path & "\GLOBAL_SOFT_DATA\Images\Accessories\" & entryID & "\"
    
    If Dir(folderPath, vbDirectory) = "" Then Exit Sub
    
    On Error Resume Next
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(folderPath)
    
    i = 1
    For Each file In folder.files
        If i > 10 Then Exit For
        
        If LCase(Right(file.name, 4)) = ".jpg" Then
            Me.Controls("imgPhoto" & i).Picture = LoadPicture(file.path)
            i = i + 1
        End If
    Next file
    
    On Error GoTo 0
End Sub
'=================================
' CAMERA CAPTURE - MANUAL CLICK + AUTO CLOSE (FIXED)
'=================================
Sub CapturePhoto(accIndex As Integer)  ' ????? Index ??? ???, Image Object ????
    On Error Resume Next
    
    If Trim(Me.txtEntryID.value) = "" Then Exit Sub
    
    Dim entryID As String, basePath As String, folderPath As String, targetPath As String
    Dim lastFile As String, newFile As String, wsh As Object
    
    entryID = Trim(Me.txtEntryID.value)
    basePath = ThisWorkbook.path & "\GLOBAL_SOFT_DATA\Images\Accessories\"
    folderPath = basePath & entryID & "\"
    targetPath = folderPath & "Acc" & accIndex & ".jpg"
    
    ' Create folders
    If Dir(basePath, vbDirectory) = "" Then MkDir basePath
    If Dir(folderPath, vbDirectory) = "" Then MkDir folderPath
    
    ' === REMEMBER LAST PHOTO (Product ???? ????) ===
    lastFile = GetLatestCameraPhoto()
    
    ' === OPEN CAMERA ===
    shell "explorer.exe shell:AppsFolder\Microsoft.WindowsCamera_8wekyb3d8bbwe!App", vbNormalFocus
    
    ' === LOOP - WAIT FOR NEW PHOTO (Product ???? ????) ===
    Do
        DoEvents
        newFile = GetLatestCameraPhoto()
        
        ' New photo detected!
        If newFile <> "" And newFile <> lastFile Then
            Application.Wait Now + TimeValue("00:00:00.5")
            
            ' Close camera
            Set wsh = CreateObject("WScript.Shell")
            wsh.Run "taskkill /f /im WindowsCamera.exe", 0, False
            Set wsh = Nothing
            
            ' Copy file
            If Dir(targetPath) <> "" Then Kill targetPath
            FileCopy newFile, targetPath
            Application.Wait Now + TimeValue("00:00:01")
            
            ' === LOAD TO CORRECT IMAGE BOX (Select Case ??) ===
            Select Case accIndex
                Case 1: Set Me.imgPhoto1.Picture = LoadPicture(targetPath)
                Case 2: Set Me.imgPhoto2.Picture = LoadPicture(targetPath)
                Case 3: Set Me.imgPhoto3.Picture = LoadPicture(targetPath)
                Case 4: Set Me.imgPhoto4.Picture = LoadPicture(targetPath)
                Case 5: Set Me.imgPhoto5.Picture = LoadPicture(targetPath)
                Case 6: Set Me.imgPhoto6.Picture = LoadPicture(targetPath)
                Case 7: Set Me.imgPhoto7.Picture = LoadPicture(targetPath)
                Case 8: Set Me.imgPhoto8.Picture = LoadPicture(targetPath)
                Case 9: Set Me.imgPhoto9.Picture = LoadPicture(targetPath)
                Case 10: Set Me.imgPhoto10.Picture = LoadPicture(targetPath)
            End Select
            
            Me.Repaint
            Exit Do
        End If
        
        Application.Wait Now + TimeValue("00:00:00.2")
    Loop While Timer < Timer + 60  ' 60 sec timeout
    
    On Error GoTo 0
End Sub

'=================================
' GET LATEST PHOTO (Product ???? ?? Function)
'=================================
Function GetLatestCameraPhoto() As String
    Dim fso As Object, folder As Object, file As Object, latestDate As Date
    Dim camPath As String
    
    GetLatestCameraPhoto = ""
    latestDate = #1/1/1900#
    
    camPath = Environ("USERPROFILE") & "\Pictures\Camera Roll\"
    If Dir(camPath, vbDirectory) = "" Then
        camPath = Environ("USERPROFILE") & "\OneDrive\Pictures\Camera Roll\"
    End If
    If Dir(camPath, vbDirectory) = "" Then Exit Function
    
    On Error Resume Next
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


'=================================
' BUTTON CLICKS (????? Index ??? ???)
'=================================
Private Sub btnCapture1_Click()
    CapturePhoto 1
End Sub

Private Sub btnCapture2_Click()
    CapturePhoto 2
End Sub

Private Sub btnCapture3_Click()
    CapturePhoto 3
End Sub

Private Sub btnCapture4_Click()
    CapturePhoto 4
End Sub

Private Sub btnCapture5_Click()
    CapturePhoto 5
End Sub

Private Sub btnCapture6_Click()
    CapturePhoto 6
End Sub

Private Sub btnCapture7_Click()
    CapturePhoto 7
End Sub

Private Sub btnCapture8_Click()
    CapturePhoto 8
End Sub

Private Sub btnCapture9_Click()
    CapturePhoto 9
End Sub

Private Sub btnCapture10_Click()
    CapturePhoto 10
End Sub


'=================================
' FORM ACTIVATE (LOAD DATA)
'=================================
Private Sub UserForm_Activate()
    On Error Resume Next
    
    If Me.Tag <> "" And Me.txtEntryID.value <> "" Then
        ' ???? ???
        Call LoadAccessoriesForEdit(Me.Tag)
        ' LoadAccessoriesForEdit ??? RowIndex ??? ?????
    Else
        ' ??? ??? - ????? 1 ?? ????
        RowIndex = 1
        Call ClearAllRows
        Call SetupRowOneOnly
        Call LoadAccessoryNamesForRow(1)
    End If
    
    On Error GoTo 0
End Sub
Sub ClearAllRows()
    Dim i As Integer
    
    For i = 1 To 10
        On Error Resume Next
        
        ' Clear values
        If ControlExists("cmbAccessory" & i) Then Me.Controls("cmbAccessory" & i).value = ""
        If ControlExists("cmbStatus" & i) Then
            Me.Controls("cmbStatus" & i).Clear
            Me.Controls("cmbStatus" & i).value = ""
        End If
        If ControlExists("txtSerial" & i) Then Me.Controls("txtSerial" & i).value = ""
        If ControlExists("cmbWorking" & i) Then
            Me.Controls("cmbWorking" & i).Clear
            Me.Controls("cmbWorking" & i).value = ""
        End If
        If ControlExists("imgPhoto" & i) Then Set Me.Controls("imgPhoto" & i).Picture = Nothing
        
        ' Disable fields
        If ControlExists("txtSerial" & i) Then Me.Controls("txtSerial" & i).enabled = False
        If ControlExists("cmbWorking" & i) Then Me.Controls("cmbWorking" & i).enabled = False
        If ControlExists("btnCapture" & i) Then Me.Controls("btnCapture" & i).enabled = False
        
        ' Hide rows 2-10 (???? ?? ????? ??????)
        If i > 1 Then
            If ControlExists("cmbAccessory" & i) Then Me.Controls("cmbAccessory" & i).visible = False
            If ControlExists("cmbStatus" & i) Then Me.Controls("cmbStatus" & i).visible = False
            If ControlExists("txtSerial" & i) Then Me.Controls("txtSerial" & i).visible = False
            If ControlExists("cmbWorking" & i) Then Me.Controls("cmbWorking" & i).visible = False
            If ControlExists("imgPhoto" & i) Then Me.Controls("imgPhoto" & i).visible = False
            If ControlExists("btnCapture" & i) Then Me.Controls("btnCapture" & i).visible = False
            If ControlExists("lblAcc" & i) Then Me.Controls("lblAcc" & i).visible = False
        End If
        On Error GoTo 0
    Next i
End Sub
Sub SetupRowOneOnly()
    On Error Resume Next
    
    ' Row 1 ????? ???? ?? ??? ?? ?????????? ??
    cmbStatus1.Clear
    cmbStatus1.AddItem "RECEIVED"
    cmbStatus1.AddItem "NOT RECEIVED"
    cmbStatus1.value = "NOT RECEIVED"
    
    cmbWorking1.Clear
    cmbWorking1.AddItem "WORKING"
    cmbWorking1.AddItem "NOT WORKING"
    cmbWorking1.value = ""
    cmbWorking1.enabled = False
    
    txtSerial1.value = ""
    txtSerial1.enabled = False
    btnCapture1.enabled = False
    Set imgPhoto1.Picture = Nothing
    
    ' ??? ??????? Visible ??? (Row 1 ??)
    cmbAccessory1.visible = True
    cmbStatus1.visible = True
    txtSerial1.visible = True
    cmbWorking1.visible = True
    imgPhoto1.visible = True
    btnCapture1.visible = True
    
    On Error GoTo 0
End Sub

'=================================
' CANCEL BUTTON
'=================================
Private Sub btnCancel_Click()
    Unload Me
End Sub
Private Sub cmbProduct_Change()
    ' ??? ???? ?? ??? accessory ??????? ?? ?? ???????? ??
    Dim hasAccessories As Boolean
    hasAccessories = False
    
    Dim i As Integer
    For i = 1 To 10
        If ControlExists("cmbAccessory" & i) Then
            If Trim(Me.Controls("cmbAccessory" & i).value) <> "" Then
                hasAccessories = True
                Exit For
            End If
        End If
    Next i
    
    If hasAccessories Then
        If MsgBox("Product ????? ?? ???? Accessories ?? ??????!" & vbCrLf & "???? ???? ?????", vbYesNo + vbExclamation, "Warning") = vbNo Then
            Exit Sub
        End If
        ' ?? ?????? ??? ?? ?????? ??? ???
        Call ClearAllRows
        RowIndex = 1
        Call SetupRowOneOnly
    Else
        ' ??? ???? ?? ?? ????? ??? ???
        Call LoadAccessoryNames
    End If
End Sub
