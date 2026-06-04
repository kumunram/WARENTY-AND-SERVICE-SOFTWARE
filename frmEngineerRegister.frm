VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEngineerRegister 
   Caption         =   "Engineer Register"
   ClientHeight    =   10515
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   19935
   OleObjectBlob   =   "frmEngineerRegister.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmEngineerRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mCurrentType As String      ' ENGINEER / VENDOR / INHOUSE / SERVICE_STATION
Private mCurrentName As String      ' Selected Name

'========================================
' FORM INITIALIZE
'========================================
Private Sub UserForm_Initialize()
    
    On Error GoTo ErrorHandler
    
    
    Me.caption = "ENGINEER / VENDOR / INHOUSE / SERVICE STATION REGISTER"
    AddMinMaxButtons Me
    '=== ListView Setup ===
    With lstWork
        .View = lvwReport
        .FullRowSelect = True
        .Gridlines = True
        .ColumnHeaders.Clear
        
        .ColumnHeaders.Add , , "EntryID", 90
        .ColumnHeaders.Add , , "Type", 70
        .ColumnHeaders.Add , , "Status", 100
        .ColumnHeaders.Add , , "Customer", 130
        .ColumnHeaders.Add , , "Mobile", 100
        .ColumnHeaders.Add , , "Product", 110
        .ColumnHeaders.Add , , "Company", 90
        .ColumnHeaders.Add , , "Model", 100
        .ColumnHeaders.Add , , "Serial", 100
        .ColumnHeaders.Add , , "Assign Date", 90
    End With
    
    '=== TYPE DROPDOWN (????) — ?? 4 ???? ===
    cmbType.Clear
    cmbType.AddItem "VENDOR"
    cmbType.AddItem "ENGINEER"
    cmbType.AddItem "INHOUSE"
    cmbType.AddItem "SERVICE_STATION"
    
    '=== STATUS FILTER DROPDOWN (?????) ===
    cmbFilter.Clear
    cmbFilter.AddItem "All"
    cmbFilter.AddItem "COMPLETED"
    cmbFilter.AddItem "PENDING"
    cmbFilter.AddItem "IN PROGRESS"
    cmbFilter.AddItem "RECEIVE_REPAIRE"
    cmbFilter.value = "All"
    
    '=== NAME DROPDOWN ???? ===
    cmbEngineer.Clear
    
    lblCount.caption = "Select Type & Name"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Initialize Error: " & Err.Description, vbCritical
End Sub

'========================================
' TYPE CHANGE ? ??? ??? ???
'========================================
Private Sub cmbType_Change()
    On Error Resume Next
    
    mCurrentType = UCase(Trim(cmbType.value & ""))
    
    If mCurrentType = "" Then
        cmbEngineer.Clear
        lstWork.ListItems.Clear
        lblCount.caption = "Select Type"
        Exit Sub
    End If
    
    ' ??? ??? ???
    Call LoadNamesByType(mCurrentType)
    
    ' ??? ??? ??? ?? ?? ???? auto-select
    If cmbEngineer.ListCount > 0 Then
        cmbEngineer.ListIndex = 0
    Else
        lstWork.ListItems.Clear
        lblCount.caption = "No " & mCurrentType & " found!"
    End If
End Sub

'========================================
' LOAD NAMES BY TYPE — ?? 4 Sheets
'========================================
Private Sub LoadNamesByType(sType As String)
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim sheetName As String
    
    cmbEngineer.Clear
    
    ' ??? ?? Sheet ?? ??? ??? ???? — ?? 4
    Select Case sType
        Case "VENDOR":          sheetName = "Vendor_Register"
        Case "ENGINEER":        sheetName = "Engineer_Register"
        Case "INHOUSE":         sheetName = "Inhouse_Register"
        Case "SERVICE_STATION": sheetName = "ServiceStation_Register"
        Case Else
            lblCount.caption = "Invalid Type!"
            Exit Sub
    End Select
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "Sheet '" & sheetName & "' not found!" & vbCrLf & _
               "Please check sheet name spelling.", vbExclamation
        lblCount.caption = "Sheet missing!"
        Exit Sub
    End If
    
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    
    If lastRow < 2 Then
        lblCount.caption = "No data in " & sheetName
        Exit Sub
    End If
    
    ' Column B (2) ?? Name ??? ???, Column F (6) = ACTIVE
    For i = 2 To lastRow
        If UCase(Trim(ws.Cells(i, 6).value & "")) = "ACTIVE" Or _
           UCase(Trim(ws.Cells(i, 6).value & "")) = "YES" Then
            If Trim(ws.Cells(i, 2).value & "") <> "" Then
                cmbEngineer.AddItem Trim(ws.Cells(i, 2).value)
            End If
        End If
    Next i
End Sub

'========================================
' ENGINEER/NAME CHANGE ? ???? ???
'========================================
Private Sub cmbEngineer_Change()
    On Error Resume Next
    
    mCurrentName = Trim(cmbEngineer.value & "")
    
    If mCurrentName = "" Then
        lstWork.ListItems.Clear
        lblCount.caption = "Select Name"
        Exit Sub
    End If
    
    Call LoadWorkData(cmbFilter.value)
End Sub

'========================================
' FILTER CHANGE ? ???? ?????
'========================================
Private Sub cmbFilter_Change()
    On Error Resume Next
    If Trim(mCurrentName) = "" Then Exit Sub
    Call LoadWorkData(cmbFilter.value)
End Sub

'========================================
' LOAD DATA — Selected Type + Name ??
'========================================
Private Sub LoadWorkData(filterStatus As String)
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim li As Object
    Dim status As String
    Dim assignDate As String
    Dim entryID As String
    Dim engineerName As String
    Dim assignType As String
    Dim dataFound As Boolean
    
    dataFound = False
    
    If Trim(mCurrentName) = "" Then
        lstWork.ListItems.Clear
        lblCount.caption = "Select Name"
        Exit Sub
    End If
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Assign_Master")
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "Assign_Master sheet not found!", vbExclamation
        Exit Sub
    End If
    
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    
    If lastRow < 2 Then
        lstWork.ListItems.Clear
        lblCount.caption = "No data in Assign_Master!"
        Exit Sub
    End If
    
    lstWork.ListItems.Clear
    
    For i = 2 To lastRow
        '=== STATUS = Column AF (32) ===
        status = UCase(Trim(ws.Cells(i, 32).value & ""))
        
        '=== NAME = Column N (14) ===
        engineerName = Trim(ws.Cells(i, 14).value & "")
        
        '=== TYPE = Column K (11) ===
        assignType = UCase(Trim(ws.Cells(i, 11).value & ""))
        
        '=== FILTER 1: SIRF SELECTED TYPE ===
        If assignType <> mCurrentType Then GoTo NextRow
        
        '=== FILTER 2: SIRF SELECTED NAME ===
        If engineerName <> mCurrentName Then GoTo NextRow
        
        '=== FILTER 3: STATUS ===
        If filterStatus <> "All" Then
            If status <> UCase(filterStatus) Then GoTo NextRow
        End If
        
        dataFound = True
        entryID = Trim(ws.Cells(i, 2).value & "")
        
        '=== DATE = Column V (22) ===
        assignDate = Trim(ws.Cells(i, 22).value & "")
        If assignDate <> "" Then
            On Error Resume Next
            assignDate = Format(CDate(assignDate), "dd-mm-yyyy")
            If Err.Number <> 0 Then assignDate = ""
            On Error GoTo 0
        End If
        
        '=== ADD TO LIST ===
        Set li = lstWork.ListItems.Add(, , entryID)
        li.SubItems(1) = Left(entryID, 3)
        li.SubItems(2) = status
        li.SubItems(3) = Trim(ws.Cells(i, 5).value & "")   'Customer
        li.SubItems(4) = Trim(ws.Cells(i, 4).value & "")   'Mobile
        li.SubItems(5) = Trim(ws.Cells(i, 6).value & "")   'Product
        li.SubItems(6) = Trim(ws.Cells(i, 7).value & "")   'Company
        li.SubItems(7) = Trim(ws.Cells(i, 8).value & "")   'Model
        li.SubItems(8) = Trim(ws.Cells(i, 9).value & "")   'Serial
        li.SubItems(9) = assignDate                         'AssignDate
        
        '=== COLOR ===
        Select Case status
            Case "COMPLETED":       li.ForeColor = RGB(0, 128, 0)      'Green
            Case "PENDING":         li.ForeColor = RGB(200, 0, 0)      'Red
            Case "IN PROGRESS":     li.ForeColor = RGB(255, 165, 0)    'Orange
            Case "RECEIVE_REPAIRE": li.ForeColor = RGB(0, 0, 200)      'Blue
        End Select
        
NextRow:
    Next i
    
    If dataFound Then
        lblCount.caption = "Total: " & lstWork.ListItems.count & " | " & mCurrentType & ": " & mCurrentName
    Else
        lblCount.caption = "No records for '" & mCurrentName & "' (" & mCurrentType & ") with status '" & filterStatus & "'"
    End If
End Sub

'========================================
' SORT BY ENTRYID
'========================================
Private Sub btnSort_Click()
    Dim i As Long, j As Long, k As Integer
    Dim tempText As String
    Dim tempSub(1 To 9) As String
    
    If lstWork.ListItems.count < 2 Then
        MsgBox "Sort karne ke liye minimum 2 records chahiye!", vbInformation
        Exit Sub
    End If
    
    For i = 1 To lstWork.ListItems.count - 1
        For j = i + 1 To lstWork.ListItems.count
            If lstWork.ListItems(i).text > lstWork.ListItems(j).text Then
                tempText = lstWork.ListItems(i).text
                For k = 1 To 9
                    tempSub(k) = lstWork.ListItems(i).SubItems(k)
                Next k
                
                lstWork.ListItems(i).text = lstWork.ListItems(j).text
                For k = 1 To 9
                    lstWork.ListItems(i).SubItems(k) = lstWork.ListItems(j).SubItems(k)
                Next k
                
                lstWork.ListItems(j).text = tempText
                For k = 1 To 9
                    lstWork.ListItems(j).SubItems(k) = tempSub(k)
                Next k
            End If
        Next j
    Next i
    
    lblCount.caption = "Sorted by EntryID | Total: " & lstWork.ListItems.count
End Sub

'========================================
' DOUBLE CLICK ? Record Complete ???
'========================================
Private Sub lstWork_DblClick()
    Dim selectedID As String
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim foundRow As Long
    Dim currentStatus As String
    
    If lstWork.selectedItem Is Nothing Then Exit Sub
    
    selectedID = lstWork.selectedItem.text
    foundRow = 0
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Assign_Master")
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "Assign_Master sheet not found!", vbExclamation
        Exit Sub
    End If
    
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    
    For i = 2 To lastRow
        If Trim(ws.Cells(i, 2).value & "") = selectedID Then
            foundRow = i
            Exit For
        End If
    Next i
    
    If foundRow = 0 Then
        MsgBox "Record not found!", vbExclamation
        Exit Sub
    End If
    
    currentStatus = UCase(Trim(ws.Cells(foundRow, 32).value & ""))
    
    Dim msg As String
    msg = "Entry ID: " & selectedID & vbCrLf & _
          "Customer: " & Trim(ws.Cells(foundRow, 5).value & "") & vbCrLf & _
          "Product: " & Trim(ws.Cells(foundRow, 6).value & "") & vbCrLf & _
          "Company: " & Trim(ws.Cells(foundRow, 7).value & "") & vbCrLf & _
          "Current Status: " & currentStatus & vbCrLf & vbCrLf & _
          "Kya aap is record ko COMPLETED mark karna chahte hain?"
    
    If MsgBox(msg, vbQuestion + vbYesNo, "Record Details") = vbYes Then
        ws.Cells(foundRow, 32).value = "COMPLETED"
        ws.Cells(foundRow, 32).Interior.Color = RGB(200, 255, 200)
        Call LoadWorkData(cmbFilter.value)
        MsgBox "Status COMPLETED ho gaya!", vbInformation
    End If
End Sub

'========================================
' CLOSE
'========================================
Private Sub btnClose_Click()
    Unload Me
End Sub

