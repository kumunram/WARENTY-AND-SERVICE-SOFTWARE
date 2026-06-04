VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmfinancialandlegal 
   Caption         =   "FINANCIAL & LEGAL"
   ClientHeight    =   9945.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8535.001
   OleObjectBlob   =   "frmfinancialandlegal.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmfinancialandlegal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub txtGSTNumber_Change()

End Sub

Private Sub UserForm_Initialize()
    LoadCountryCombo
    LoadStateCombo
    LoadFinancialSettings
End Sub

Private Sub LoadCountryCombo()
    With Me.cmbCountry
        .Clear
        .AddItem "Add New..."
        .AddItem "India"
        .AddItem "USA"
        .AddItem "UK"
        .AddItem "Australia"
        .AddItem "Canada"
    End With
End Sub

Private Sub LoadStateCombo()
    With Me.cmbState
        .Clear
        .AddItem "Add New..."
        .AddItem "Odisha"
        .AddItem "West Bengal"
        .AddItem "Maharashtra"
        .AddItem "Karnataka"
        .AddItem "Delhi"
    End With
End Sub

Private Sub cmbCountry_Click()
    If Me.cmbCountry.value = "Add New..." Then
        Dim newCountry As String
        newCountry = InputBox("Enter New Country:", "Add Country")
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

Private Sub cmbState_Click()
    If Me.cmbState.value = "Add New..." Then
        Dim newState As String
        newState = InputBox("Enter New State:", "Add State")
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
'===========================================
' FINANCIAL START DATE - SIMPLE FORMAT
'===========================================
Private Sub txtFinancialStart_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Dim dt As String
    dt = Trim(Me.txtFinancialStart.value)
    
    ' Empty check
    If dt = "" Then Exit Sub
    
    ' Remove all non-numeric characters
    dt = Replace(dt, "/", "")
    dt = Replace(dt, "-", "")
    dt = Replace(dt, ".", "")
    dt = Replace(dt, " ", "")
    
    ' Must be 8 digits (DDMMYYYY)
    If Len(dt) <> 8 Or Not IsNumeric(dt) Then
        MsgBox "Please enter date as DDMMYYYY" & vbCrLf & _
               "Example: 01042025 for 01/04/2025", vbExclamation
        Me.txtFinancialStart.value = ""
        Cancel = True
        Exit Sub
    End If
    
    ' Extract parts
    Dim d As String, m As String, Y As String
    d = Left(dt, 2)
    m = Mid(dt, 3, 2)
    Y = Right(dt, 4)
    
    ' Validate
    If CInt(d) < 1 Or CInt(d) > 31 Then
        MsgBox "Invalid Day!", vbExclamation
        Me.txtFinancialStart.value = ""
        Cancel = True
        Exit Sub
    End If
    
    If CInt(m) < 1 Or CInt(m) > 12 Then
        MsgBox "Invalid Month!", vbExclamation
        Me.txtFinancialStart.value = ""
        Cancel = True
        Exit Sub
    End If
    
    ' Format and display
    Me.txtFinancialStart.value = d & "/" & m & "/" & Y
    
    ' Calculate year
    CalculateFinYear
End Sub

'===========================================
' FINANCIAL END DATE - SIMPLE FORMAT
'===========================================
Private Sub txtFinancialEnd_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Dim dt As String
    dt = Trim(Me.txtFinancialEnd.value)
    
    ' Empty check
    If dt = "" Then Exit Sub
    
    ' Remove all non-numeric characters
    dt = Replace(dt, "/", "")
    dt = Replace(dt, "-", "")
    dt = Replace(dt, ".", "")
    dt = Replace(dt, " ", "")
    
    ' Must be 8 digits (DDMMYYYY)
    If Len(dt) <> 8 Or Not IsNumeric(dt) Then
        MsgBox "Please enter date as DDMMYYYY" & vbCrLf & _
               "Example: 31032025 for 31/03/2025", vbExclamation
        Me.txtFinancialEnd.value = ""
        Cancel = True
        Exit Sub
    End If
    
    ' Extract parts
    Dim d As String, m As String, Y As String
    d = Left(dt, 2)
    m = Mid(dt, 3, 2)
    Y = Right(dt, 4)
    
    ' Validate
    If CInt(d) < 1 Or CInt(d) > 31 Then
        MsgBox "Invalid Day!", vbExclamation
        Me.txtFinancialEnd.value = ""
        Cancel = True
        Exit Sub
    End If
    
    If CInt(m) < 1 Or CInt(m) > 12 Then
        MsgBox "Invalid Month!", vbExclamation
        Me.txtFinancialEnd.value = ""
        Cancel = True
        Exit Sub
    End If
    
    ' Format and display
    Me.txtFinancialEnd.value = d & "/" & m & "/" & Y
    
    ' Calculate year
    CalculateFinYear
End Sub

'===========================================
' CALCULATE FINANCIAL YEAR
'===========================================
Private Sub CalculateFinYear()
    Dim startDate As String, endDate As String
    Dim startParts() As String, endParts() As String
    Dim sy As Integer, ey As Integer
    
    startDate = Me.txtFinancialStart.value
    endDate = Me.txtFinancialEnd.value
    
    If startDate = "" Or endDate = "" Then Exit Sub
    
    ' Parse dates
    startParts = Split(startDate, "/")
    endParts = Split(endDate, "/")
    
    If UBound(startParts) <> 2 Or UBound(endParts) <> 2 Then Exit Sub
    
    sy = CInt(startParts(2))
    ey = CInt(endParts(2))
    
    ' Display Financial Year
    If ey = sy + 1 Then
        Me.txtCurrentYear.value = sy & "-" & Right(CStr(ey), 2)
    Else
        Me.txtCurrentYear.value = sy & "-" & ey
    End If
End Sub
Private Sub LoadFinancialSettings()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Software_Config")
    On Error GoTo 0
    
    If ws Is Nothing Then Exit Sub
    
    ' Country
    If ws.Range("B23").value <> "" Then
        Me.cmbCountry.value = CStr(ws.Range("B23").value)
    End If
    
    ' State
    If ws.Range("B24").value <> "" Then
        Me.cmbState.value = CStr(ws.Range("B24").value)
    End If
    
    ' Text fields
    Me.txtGSTNumber.value = CStr(ws.Range("B25").value)
    Me.txtServiceTax.value = CStr(ws.Range("B26").value)
    Me.txtJudicialArea.value = CStr(ws.Range("B27").value)
    
    ' Financial Start Date - Handle Date type properly
    If ws.Range("B28").value <> "" Then
        If IsDate(ws.Range("B28").value) Then
            ' If it's a Date type, format it as DD/MM/YYYY
            Me.txtFinancialStart.value = Format(ws.Range("B28").value, "dd/mm/yyyy")
        Else
            ' If it's text, use as is
            Me.txtFinancialStart.value = CStr(ws.Range("B28").value)
        End If
    End If
    
    ' Financial End Date
    If ws.Range("B29").value <> "" Then
        If IsDate(ws.Range("B29").value) Then
            Me.txtFinancialEnd.value = Format(ws.Range("B29").value, "dd/mm/yyyy")
        Else
            Me.txtFinancialEnd.value = CStr(ws.Range("B29").value)
        End If
    End If
    
    ' Current Year
    Me.txtCurrentYear.value = CStr(ws.Range("B30").value)
End Sub
Sub CheckConfigSheet()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Software_Config")
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "Software_Config sheet not found!", vbExclamation
        Exit Sub
    End If
    
    ' Check values
    MsgBox "B23 (Country): " & ws.Range("B23").value & vbCrLf & _
           "B28 (Fin Start): " & ws.Range("B28").value & vbCrLf & _
           "B29 (Fin End): " & ws.Range("B29").value, vbInformation
End Sub
'===========================================
' SAVE ALL SETTINGS - PROFESSIONAL
'===========================================
Private Sub btnSave_Click()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Software_Config")
    
    ws.Range("B23").value = Me.cmbCountry.value
    ws.Range("B24").value = Me.cmbState.value
    ws.Range("B25").value = Me.txtGSTNumber.value
    ws.Range("B26").value = Me.txtServiceTax.value
    ws.Range("B27").value = Me.txtJudicialArea.value
    ws.Range("B28").value = Me.txtFinancialStart.value
    ws.Range("B29").value = Me.txtFinancialEnd.value
    ws.Range("B30").value = Me.txtCurrentYear.value
    
    MsgBox "Financial & Legal Settings saved successfully!" & vbCrLf & vbCrLf & _
           "Country: " & Me.cmbCountry.value & vbCrLf & _
           "State: " & Me.cmbState.value & vbCrLf & _
           "Financial Year: " & Me.txtCurrentYear.value, _
           vbInformation, "GLOBAL SOFT - Settings Saved"
End Sub

'===========================================
' RESET TO DEFAULT - FIXED
'===========================================
Private Sub btnReset_Click()
    If MsgBox("Reset all fields to default?" & vbCrLf & _
              "This will clear all entered data.", _
              vbQuestion + vbYesNo, "Reset to Default") = vbNo Then Exit Sub
    
    ' Clear all fields
    Me.cmbCountry.value = ""
    Me.cmbState.value = ""
    Me.txtGSTNumber.value = ""
    Me.txtServiceTax.value = ""
    Me.txtJudicialArea.value = ""
    Me.txtFinancialStart.value = ""
    Me.txtFinancialEnd.value = ""
    Me.txtCurrentYear.value = ""
    
    ' Reload ComboBoxes
    LoadCountryCombo
    LoadStateCombo
    
    MsgBox "All fields reset to default!", vbInformation, "Reset Complete"
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub
