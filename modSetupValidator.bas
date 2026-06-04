Attribute VB_Name = "modSetupValidator"
'===========================================================
' MODULE: modSetupValidator
' PURPOSE: Complete System Repair & Validation
'===========================================================

Option Explicit

'===========================================================
' MAIN REPAIR FUNCTIONS (Ye previously missing the!)
'===========================================================

Public Sub AutoRepairSheets()
    ' Missing sheets check aur auto-create
    Dim requiredSheets As Variant
    Dim i As Integer
    Dim ws As Worksheet
    Dim sheetExists As Boolean
    
    requiredSheets = Array("Customer_Master", "Job_Master", "Settings", _
                          "Product_Master", "Engineer_Register")
    
    For i = LBound(requiredSheets) To UBound(requiredSheets)
        sheetExists = False
        For Each ws In ThisWorkbook.Worksheets
            If ws.name = requiredSheets(i) Then
                sheetExists = True
                Exit For
            End If
        Next ws
        
        If Not sheetExists Then
            Set ws = ThisWorkbook.Worksheets.Add
            ws.name = requiredSheets(i)
            ' Add default headers based on sheet type
            Call SetupDefaultSheetHeaders(ws, requiredSheets(i))
        End If
    Next i
End Sub

Public Sub AutoRepairPaths()
    ' Missing folders auto-create
    Dim paths As Variant
    Dim i As Integer
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    paths = Array("D:\BACKUP\", "D:\GLOBAL_SOFT_DATA\Images\Customers\", _
                  "D:\GLOBAL_SOFT_DATA\Exports\", "D:\GLOBAL_SOFT_DATA\Logs\")
    
    On Error Resume Next
    For i = LBound(paths) To UBound(paths)
        If Not fso.FolderExists(paths(i)) Then
            fso.CreateFolder (paths(i))
        End If
    Next i
    On Error GoTo 0
    
    Set fso = Nothing
End Sub

Public Sub AutoRepairDatabase()
    ' Database headers validation and fix
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Customer_Master")
    If Not ws Is Nothing Then
        ' Check and fix headers
        If ws.Range("A1").value <> "Customer_ID" Then ws.Range("A1").value = "Customer_ID"
        If ws.Range("B1").value <> "Customer_Name" Then ws.Range("B1").value = "Customer_Name"
        If ws.Range("C1").value <> "Mobile" Then ws.Range("C1").value = "Mobile"
        ' Add more headers as needed
    End If
    
    Set ws = ThisWorkbook.Worksheets("Job_Master")
    If Not ws Is Nothing Then
        If ws.Range("A1").value <> "Job_ID" Then ws.Range("A1").value = "Job_ID"
        ' Add more job fields
    End If
    On Error GoTo 0
End Sub

Public Sub AutoRepairSettings()
    ' Settings integrity check
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Settings")
    If Not ws Is Nothing Then
        ' Check B10 (Backup Path)
        If Trim(ws.Range("B10").value) = "" Then ws.Range("B10").value = "D:\BACKUP\"
        ' Check B12 (Gmail)
        If Trim(ws.Range("B12").value) = "" Then ws.Range("B12").value = ""
        ' Check B14 (Cloud Enabled)
        If Trim(ws.Range("B14").value) = "" Then ws.Range("B14").value = "No"
    End If
    On Error GoTo 0
End Sub

Private Sub SetupDefaultSheetHeaders(ByVal ws As Worksheet, ByVal sheetName As String)
    ' Helper sub to setup default headers for new sheets
    Select Case sheetName
        Case "Customer_Master"
            ws.Range("A1:F1").value = Array("Customer_ID", "Customer_Name", "Mobile", "Email", "Address", "Created_Date")
        Case "Job_Master"
            ws.Range("A1:E1").value = Array("Job_ID", "Customer_ID", "Product", "Issue", "Status")
        Case "Settings"
            ws.Range("A1:A3").value = Array("Backup_Path", "Gmail_ID", "Cloud_Backup")
        Case "Product_Master"
            ws.Range("A1:D1").value = Array("Product_ID", "Product_Name", "Category", "Price")
        Case "Engineer_Register"
            ws.Range("A1:D1").value = Array("Engineer_ID", "Name", "Mobile", "Specialization")
        ' Add more cases as needed
    End Select
End Sub

'===========================================================
' REFERENCE CHECK FUNCTIONS
'===========================================================

Public Sub CheckAllReferences()
    Dim ref As Object
    Dim hasCommonControls As Boolean
    Dim hasCDO As Boolean
    
    On Error Resume Next
    
    For Each ref In Application.VBE.ActiveVBProject.References
        If ref.name = "MSComctlLib" Then hasCommonControls = True
        If ref.name = "CDO" Then hasCDO = True
    Next ref
    
    Dim msg As String
    
    If hasCommonControls And hasCDO Then
        MsgBox "All References are CORRECTLY configured!" & vbCrLf & _
               "You can proceed with Gmail Setup.", vbInformation, "System Check"
    Else
        msg = "MISSING REFERENCES DETECTED!" & vbCrLf & vbCrLf
        
        If Not hasCommonControls Then
            msg = msg & "Missing: Microsoft Windows Common Controls 6.0 (SP6)" & vbCrLf
        End If
        
        If Not hasCDO Then
            msg = msg & "Missing: Microsoft CDO for Windows 2000 Library" & vbCrLf
        End If
        
        msg = msg & vbCrLf & "How to Fix:" & vbCrLf & _
              "1. Press Alt+F11" & vbCrLf & _
              "2. Click Tools -> References" & vbCrLf & _
              "3. Check (tick) the missing items" & vbCrLf & _
              "4. Click OK"
        
        MsgBox msg, vbCritical, "Setup Incomplete"
    End If
End Sub

'===========================================================
' EMERGENCY REPAIR WITH RESTART (FIXED VERSION)
'===========================================================

Public Sub EmergencyRepairWithRestart()
    On Error GoTo ErrorHandler
    
    ' Repairs
    Call AutoRepairSheets
    Call AutoRepairPaths
    Call AutoRepairDatabase
    Call AutoRepairSettings
    ThisWorkbook.Save
    
    Dim filePath As String
    filePath = ThisWorkbook.FullName
    
    ' Batch file create karo temp folder mein
    Dim batchPath As String
    batchPath = Environ("TEMP") & "\restart_gits.bat"
    
    Dim fileNum As Integer
    fileNum = FreeFile
    
    Open batchPath For Output As #fileNum
    Print #fileNum, "@echo off"
    Print #fileNum, "echo Repair Complete!"
    Print #fileNum, "timeout /t 2 /nobreak >nul" ' 2 second wait
    Print #fileNum, "start """" """ & filePath & """"
    Print #fileNum, "del """ & batchPath & """" ' Batch file delete kar do
    Close #fileNum
    
    ' Batch file run karo hidden mode mein
    shell batchPath, vbHide
    
    ' Turant quit karo
    Application.Quit
    Exit Sub
    
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical, "Repair Failed"
End Sub

Public Sub EmergencyRepairSilent()
    ' Silent version without restart confirmation
    Application.ScreenUpdating = False
    
    Call AutoRepairSheets
    Call AutoRepairPaths
    Call AutoRepairDatabase
    Call AutoRepairSettings
    
    ThisWorkbook.Save
    
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    wsh.Run "cmd /c echo Repair Complete && timeout /t 2 && start excel """ & ThisWorkbook.FullName & """", vbHide
    
    Set wsh = Nothing
    Application.Quit
End Sub

