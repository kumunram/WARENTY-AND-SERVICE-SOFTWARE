Attribute VB_Name = "modSystemValidator"
Option Explicit

'===========================================
' SYSTEM VALIDATOR - A to Z Checking
'===========================================
Public Sub FullSystemCheck()
    Dim checkResult As Boolean
    Dim errorCount As Long
    Dim repairCount As Long
    
    errorCount = 0
    repairCount = 0
    
    On Error GoTo ErrorHandler
    
    '1. REQUIRED SHEETS CHECK
    If Not CheckAllSheets Then
        errorCount = errorCount + 1
        If AutoRepairSheets Then repairCount = repairCount + 1
    End If
    
    '2. FOLDER PATHS CHECK
    If Not CheckAllPaths Then
        errorCount = errorCount + 1
        If AutoRepairPaths Then repairCount = repairCount + 1
    End If
    
    '3. DATABASE STRUCTURE CHECK
    If Not CheckDatabaseStructure Then
        errorCount = errorCount + 1
        If AutoRepairDatabase Then repairCount = repairCount + 1
    End If
    
    '4. SETTINGS VALIDATION
    If Not CheckSettings Then
        errorCount = errorCount + 1
        If AutoRepairSettings Then repairCount = repairCount + 1
    End If
    
    '5. HEADERS CHECK
    If Not CheckAllHeaders Then
        errorCount = errorCount + 1
        If AutoRepairHeaders Then repairCount = repairCount + 1
    End If
    
    '6. CORRUPTION CHECK
    If Not CheckCorruption Then
        errorCount = errorCount + 1
        MsgBox "Database may be corrupted!" & vbCrLf & _
               "Please restore from backup.", vbCritical
    End If
    
    'Final Report
    If errorCount = 0 Then
        MsgBox "? System Check Passed!" & vbCrLf & _
               "All systems operational.", vbInformation, "System Ready"
    Else
        MsgBox "?? System Check Complete!" & vbCrLf & _
               "Errors Found: " & errorCount & vbCrLf & _
               "Auto-Repaired: " & repairCount, vbExclamation, "Check Complete"
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Critical Error during system check: " & Err.Description, vbCritical
End Sub

'===========================================
' 1. SHEET VALIDATION
'===========================================
Private Function CheckAllSheets() As Boolean
    Dim requiredSheets As Variant
    Dim i As Integer
    Dim missing As String
    
    requiredSheets = Array("Customer_Master", "Job_Master", "Job_Product", _
                          "Payment_Master", "Product_Master", "Settings", _
                          "Software_Config", "Staff_Master")
    
    missing = ""
    For i = LBound(requiredSheets) To UBound(requiredSheets)
        If Not sheetExists(requiredSheets(i)) Then
            missing = missing & requiredSheets(i) & vbCrLf
        End If
    Next i
    
    If missing <> "" Then
        MsgBox "Missing Sheets:" & vbCrLf & missing, vbExclamation
        CheckAllSheets = False
    Else
        CheckAllSheets = True
    End If
End Function

Private Function sheetExists(sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    sheetExists = (Err.Number = 0)
    On Error GoTo 0
End Function

Private Function AutoRepairSheets() As Boolean
    On Error Resume Next
    
    'Create missing sheets with basic structure
    If Not sheetExists("Settings") Then
        Dim ws As Worksheet
        Set ws = ThisWorkbook.Sheets.Add
        ws.name = "Settings"
        ws.Range("A1").value = "Setting_Name"
        ws.Range("B1").value = "Value"
        'Add default settings
        ws.Range("A2").value = "Backup_Location"
        ws.Range("A3").value = "Gmail_User"
        ws.Range("A4").value = "Gmail_Password"
        ws.Range("A5").value = "Cloud_Enabled"
        ws.Range("B5").value = "FALSE"
    End If
    
    AutoRepairSheets = True
    On Error GoTo 0
End Function

'===========================================
' 2. PATH VALIDATION
'===========================================
Private Function CheckAllPaths() As Boolean
    Dim paths As Variant
    Dim i As Integer
    Dim missing As String
    
    paths = Array(ThisWorkbook.path & "\GLOBAL_SOFT_DATA\Images\Customers\", _
                  ThisWorkbook.path & "\BACKUP\", _
                  ThisWorkbook.path & "\EXPORT\")
    
    missing = ""
    For i = LBound(paths) To UBound(paths)
        If Dir(paths(i), vbDirectory) = "" Then
            missing = missing & paths(i) & vbCrLf
        End If
    Next i
    
    If missing <> "" Then
        MsgBox "Missing Folders:" & vbCrLf & missing, vbExclamation
        CheckAllPaths = False
    Else
        CheckAllPaths = True
    End If
End Function

Private Function AutoRepairPaths() As Boolean
    On Error Resume Next
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    'Create main data folder
    If Dir(ThisWorkbook.path & "\GLOBAL_SOFT_DATA", vbDirectory) = "" Then
        MkDir ThisWorkbook.path & "\GLOBAL_SOFT_DATA"
    End If
    
    'Create Images folder
    If Dir(ThisWorkbook.path & "\GLOBAL_SOFT_DATA\Images", vbDirectory) = "" Then
        MkDir ThisWorkbook.path & "\GLOBAL_SOFT_DATA\Images"
    End If
    
    'Create Customers folder
    If Dir(ThisWorkbook.path & "\GLOBAL_SOFT_DATA\Images\Customers", vbDirectory) = "" Then
        MkDir ThisWorkbook.path & "\GLOBAL_SOFT_DATA\Images\Customers"
    End If
    
    'Create Backup folder
    If Dir(ThisWorkbook.path & "\BACKUP", vbDirectory) = "" Then
        MkDir ThisWorkbook.path & "\BACKUP"
    End If
    
    'Create Export folder
    If Dir(ThisWorkbook.path & "\EXPORT", vbDirectory) = "" Then
        MkDir ThisWorkbook.path & "\EXPORT"
    End If
    
    AutoRepairPaths = True
    On Error GoTo 0
End Function

'===========================================
' 3. DATABASE STRUCTURE CHECK
'===========================================
Private Function CheckDatabaseStructure() As Boolean
    Dim ws As Worksheet
    Dim lastRow As Long
    
    On Error Resume Next
    
    'Check Customer_Master
    Set ws = ThisWorkbook.Sheets("Customer_Master")
    If ws.Range("A1").value <> "CustomerID" Then
        CheckDatabaseStructure = False
        Exit Function
    End If
    
    'Check Job_Master
    Set ws = ThisWorkbook.Sheets("Job_Master")
    If ws.Range("A1").value <> "EntryID" Then
        CheckDatabaseStructure = False
        Exit Function
    End If
    
    'Check for blank rows corruption
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    If lastRow > 100000 Then
        'Possible corruption - too many rows
        CheckDatabaseStructure = False
        Exit Function
    End If
    
    CheckDatabaseStructure = True
    On Error GoTo 0
End Function

Private Function AutoRepairDatabase() As Boolean
    On Error Resume Next
    Dim ws As Worksheet
    Dim lastRow As Long  ' <-- ?? ADD ????!
    
    'Repair Customer_Master headers if blank
    Set ws = ThisWorkbook.Sheets("Customer_Master")
    If ws.Range("A1").value = "" Then
        ws.Range("A1").value = "CustomerID"
        ws.Range("B1").value = "Mobile"
        ws.Range("C1").value = "Name"
        ws.Range("D1").value = "Address"
        ws.Range("E1").value = "Email"
        ws.Range("F1").value = "GST"
        ws.Range("G1").value = "Photo_Name"
    End If
    
    'Repair Job_Master headers
    Set ws = ThisWorkbook.Sheets("Job_Master")
    If ws.Range("A1").value = "" Then
        ws.Range("A1").value = "EntryID"
        ws.Range("B1").value = "CustomerID"
        ws.Range("C1").value = "Mobile"
        ws.Range("D1").value = "CustomerName"
    End If
    
    'Clear excess blank rows (corruption fix)
    Application.ScreenUpdating = False
    Set ws = ThisWorkbook.Sheets("Customer_Master")
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    If lastRow < ws.Rows.count Then
        ws.Range(lastRow + 1 & ":" & ws.Rows.count).ClearContents
    End If
    
    Application.ScreenUpdating = True
    AutoRepairDatabase = True
    On Error GoTo 0
End Function

'===========================================
' 4. SETTINGS VALIDATION
'===========================================
Private Function CheckSettings() As Boolean
    Dim ws As Worksheet
    Dim backupLoc As String
    Dim cloudEnabled As String
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Settings")
    
    backupLoc = ws.Range("B10").value
    cloudEnabled = UCase(ws.Range("B14").value)
    
    If backupLoc = "" Or backupLoc = "OK" Then
        CheckSettings = False
    ElseIf cloudEnabled <> "TRUE" And cloudEnabled <> "FALSE" Then
        CheckSettings = False
    Else
        CheckSettings = True
    End If
    On Error GoTo 0
End Function

Private Function AutoRepairSettings() As Boolean
    On Error Resume Next
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Settings")
    
    'Fix Backup Location
    If ws.Range("B10").value = "" Or ws.Range("B10").value = "OK" Then
        ws.Range("B10").value = ThisWorkbook.path & "\BACKUP\"
        ws.Range("A10").value = "Backup_Location"
    End If
    
    'Fix Cloud Settings
    If ws.Range("B14").value = "" Then
        ws.Range("A14").value = "Cloud_Enabled"
        ws.Range("B14").value = "FALSE"
    End If
    
    'Ensure Gmail settings exist
    If ws.Range("B12").value = "" Then
        ws.Range("A12").value = "Gmail_User"
        ws.Range("B12").value = ""
    End If
    
    AutoRepairSettings = True
    On Error GoTo 0
End Function

'===========================================
' 5. HEADERS CHECK
'===========================================
Private Function CheckAllHeaders() As Boolean
    Dim customerHeaders As Variant
    Dim i As Integer
    Dim ws As Worksheet
    Dim mismatch As String
    
    customerHeaders = Array("CustomerID", "Mobile", "Name", "Address", "Email", "GST", "Photo_Name")
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Customer_Master")
    mismatch = ""
    
    For i = 0 To UBound(customerHeaders)
        If ws.Cells(1, i + 1).value <> customerHeaders(i) Then
            mismatch = mismatch & "Col " & (i + 1) & ": Expected '" & customerHeaders(i) & "' Found '" & ws.Cells(1, i + 1).value & "'" & vbCrLf
        End If
    Next i
    
    If mismatch <> "" Then
        MsgBox "Header Mismatch:" & vbCrLf & mismatch, vbExclamation
        CheckAllHeaders = False
    Else
        CheckAllHeaders = True
    End If
    On Error GoTo 0
End Function

Private Function AutoRepairHeaders() As Boolean
    'Already handled in AutoRepairDatabase
    AutoRepairHeaders = True
End Function

'===========================================
' 6. CORRUPTION CHECK
'===========================================
Private Function CheckCorruption() As Boolean
    Dim ws As Worksheet
    Dim cell As Range
    Dim corrupt As Boolean
    
    corrupt = False
    
    On Error Resume Next
    
    'Check for #REF! errors
    For Each ws In ThisWorkbook.Sheets
        For Each cell In ws.UsedRange
            If InStr(cell.text, "#REF!") > 0 Or _
               InStr(cell.text, "#VALUE!") > 0 Or _
               InStr(cell.text, "#N/A") > 0 Then
                corrupt = True
                cell.Interior.Color = RGB(255, 200, 200) 'Mark red
            End If
        Next cell
    Next ws
    
    CheckCorruption = Not corrupt
    On Error GoTo 0
End Function

'===========================================
' UTILITY: Quick Repair Button
'===========================================
Public Sub EmergencyRepair()
    'Single click repair everything
    MsgBox "Starting Emergency Repair...", vbInformation
    
    Call AutoRepairSheets
    Call AutoRepairPaths
    Call AutoRepairDatabase
    Call AutoRepairSettings
    
    MsgBox "Emergency Repair Complete!" & vbCrLf & _
           "Please restart the software.", vbInformation
End Sub

'===========================================
' VALIDATION BEFORE EXIT
'===========================================
Public Sub ValidateBeforeExit()
    'Check if any data is unsaved
    If ThisWorkbook.Saved = False Then
        Dim res As VbMsgBoxResult
        res = MsgBox("Unsaved changes detected!" & vbCrLf & _
                     "Save before exit?", vbYesNoCancel + vbExclamation)
        
        If res = vbYes Then
            ThisWorkbook.Save
        ElseIf res = vbCancel Then
            'Cancel exit - but how to stop?
            'This needs to be called from BeforeClose
        End If
    End If
    
    'Create exit backup
    CreateExitBackup
End Sub

Private Sub CreateExitBackup()
    'Quick backup on exit
    On Error Resume Next
    Dim backupPath As String
    backupPath = ThisWorkbook.path & "\BACKUP\EXIT_BACKUP_" & Format(Now, "yyyymmdd") & ".xlsm"
    ThisWorkbook.SaveCopyAs backupPath
    On Error GoTo 0
End Sub

