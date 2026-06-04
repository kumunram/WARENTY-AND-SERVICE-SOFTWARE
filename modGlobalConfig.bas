Attribute VB_Name = "modGlobalConfig"
'===========================================
' modGlobalConfig - GLOBAL SOFT v2.0
' Permanent Column Mapping Solution
'===========================================

Public Const CONFIG_SHEET As String = "_GLOBAL_CONFIG"
Public Const CONFIG_CUST_MAP As String = "CUSTOMER_MASTER_MAP"

' Structure to hold detected column positions
Public Type CustomerMasterMap
    Col_CustID As Long      ' Column A usually
    Col_Mobile As Long      ' Detected dynamically
    Col_Name As Long        ' Detected dynamically
    Col_Address As Long     ' Detected dynamically
    Col_Email As Long       ' Detected dynamically
    IsDetected As Boolean   ' Was auto-detection successful?
End Type

' Global variable - accessible to ALL forms
Public gCustMap As CustomerMasterMap

'===========================================
' AUTO-DETECT: Customer_Master Column Order
'===========================================
Public Function DetectCustomerMasterColumns() As CustomerMasterMap
    Dim ws As Worksheet
    Dim cm As CustomerMasterMap
    Dim sampleRow As Long
    Dim valA As String, valB As String, valC As String, valD As String, valE As String
    Dim isB_Mobile As Boolean, isC_Mobile As Boolean
    Dim isB_Name As Boolean, isC_Name As Boolean
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Customer_Master")
    On Error GoTo 0
    
    If ws Is Nothing Then
        cm.IsDetected = False
        DetectCustomerMasterColumns = cm
        Exit Function
    End If
    
    ' Find first data row
    sampleRow = 2
    Do While sampleRow < 100
        valA = Trim(CStr(ws.Cells(sampleRow, 1).value))
        If valA <> "" Then Exit Do
        sampleRow = sampleRow + 1
    Loop
    
    If sampleRow >= 100 Then
        cm.IsDetected = False
        DetectCustomerMasterColumns = cm
        Exit Function
    End If
    
    ' Read sample values
    valB = Trim(CStr(ws.Cells(sampleRow, 2).value))
    valC = Trim(CStr(ws.Cells(sampleRow, 3).value))
    valD = Trim(CStr(ws.Cells(sampleRow, 4).value))
    valE = Trim(CStr(ws.Cells(sampleRow, 5).value))
    
    ' Heuristic detection
    isB_Mobile = IsValidMobile(valB)
    isC_Mobile = IsValidMobile(valC)
    isB_Name = IsValidName(valB)
    isC_Name = IsValidName(valC)
    
    ' Column A is always CustID
    cm.Col_CustID = 1
    
    ' Determine Mobile and Name columns
    If isB_Mobile And isC_Name Then
        ' Standard: B=Mobile, C=Name
        cm.Col_Mobile = 2
        cm.Col_Name = 3
        cm.Col_Address = 4
        cm.Col_Email = 5
        cm.IsDetected = True
        
    ElseIf isC_Mobile And isB_Name Then
        ' SWAPPED: B=Name, C=Mobile  ? YOUR CURRENT PROBLEM!
        cm.Col_Mobile = 3
        cm.Col_Name = 2
        cm.Col_Address = 4
        cm.Col_Email = 5
        cm.IsDetected = True
        
    ElseIf isB_Mobile And Not isC_Mobile Then
        ' B is mobile, C might be name (uncertain but likely)
        cm.Col_Mobile = 2
        cm.Col_Name = 3
        cm.Col_Address = 4
        cm.Col_Email = 5
        cm.IsDetected = True
        
    ElseIf isC_Mobile And Not isB_Mobile Then
        ' C is mobile, B might be name
        cm.Col_Mobile = 3
        cm.Col_Name = 2
        cm.Col_Address = 4
        cm.Col_Email = 5
        cm.IsDetected = True
        
    Else
        ' Cannot determine - use default
        cm.Col_Mobile = 2
        cm.Col_Name = 3
        cm.Col_Address = 4
        cm.Col_Email = 5
        cm.IsDetected = False
    End If
    
    ' Save to config sheet
    SaveColumnConfig cm
    
    DetectCustomerMasterColumns = cm
End Function

'===========================================
' HELPER: Validate Mobile Number
'===========================================
Private Function IsValidMobile(val As String) As Boolean
    Dim cleanVal As String
    Dim i As Long
    
    ' Remove spaces, dashes, +91
    cleanVal = val
    cleanVal = Replace(cleanVal, " ", "")
    cleanVal = Replace(cleanVal, "-", "")
    cleanVal = Replace(cleanVal, "+91", "")
    cleanVal = Replace(cleanVal, "+", "")
    
    ' Check if numeric and 10 digits
    If Len(cleanVal) = 10 Then
        IsValidMobile = True
        For i = 1 To 10
            If Not IsNumeric(Mid(cleanVal, i, 1)) Then
                IsValidMobile = False
                Exit Function
            End If
        Next i
    ElseIf Len(cleanVal) = 12 And Left(cleanVal, 2) = "91" Then
        ' 91 prefix
        IsValidMobile = True
        For i = 3 To 12
            If Not IsNumeric(Mid(cleanVal, i, 1)) Then
                IsValidMobile = False
                Exit Function
            End If
        Next i
    Else
        IsValidMobile = False
    End If
End Function

'===========================================
' HELPER: Validate Name (heuristic)
'===========================================
Private Function IsValidName(val As String) As Boolean
    Dim i As Long
    Dim hasAlpha As Boolean
    
    If Len(val) < 3 Then
        IsValidName = False
        Exit Function
    End If
    
    hasAlpha = False
    For i = 1 To Len(val)
        If UCase(Mid(val, i, 1)) >= "A" And UCase(Mid(val, i, 1)) <= "Z" Then
            hasAlpha = True
            Exit For
        End If
    Next i
    
    ' Name should have alphabets and not be all numeric
    IsValidName = hasAlpha And Not IsNumeric(val)
End Function

'===========================================
' SAVE: Column Config to Sheet
'===========================================
Private Sub SaveColumnConfig(cm As CustomerMasterMap)
    Dim ws As Worksheet
    Dim found As Boolean
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(CONFIG_SHEET)
    found = (Err.Number = 0)
    On Error GoTo 0
    
    If Not found Then
        Set ws = ThisWorkbook.Sheets.Add
        ws.name = CONFIG_SHEET
        ws.visible = xlSheetVeryHidden  ' Hide from users
    End If
    
    With ws
        .Cells.Clear
        .Range("A1").value = "CONFIG_KEY"
        .Range("B1").value = "VALUE"
        .Range("A1:B1").Font.Bold = True
        
        .Range("A2").value = "CUST_COL_CUSTID"
        .Range("B2").value = cm.Col_CustID
        
        .Range("A3").value = "CUST_COL_MOBILE"
        .Range("B3").value = cm.Col_Mobile
        
        .Range("A4").value = "CUST_COL_NAME"
        .Range("B4").value = cm.Col_Name
        
        .Range("A5").value = "CUST_COL_ADDRESS"
        .Range("B5").value = cm.Col_Address
        
        .Range("A6").value = "CUST_COL_EMAIL"
        .Range("B6").value = cm.Col_Email
        
        .Range("A7").value = "CUST_IS_DETECTED"
        .Range("B7").value = IIf(cm.IsDetected, "YES", "NO")
        
        .Range("A8").value = "LAST_UPDATED"
        .Range("B8").value = Now()
    End With
End Sub

'===========================================
' LOAD: Column Config from Sheet
'===========================================
Public Function LoadColumnConfig() As CustomerMasterMap
    Dim ws As Worksheet
    Dim cm As CustomerMasterMap
    Dim found As Boolean
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(CONFIG_SHEET)
    found = (Err.Number = 0)
    On Error GoTo 0
    
    If Not found Then
        ' No config exists - detect now
        cm = DetectCustomerMasterColumns()
        LoadColumnConfig = cm
        Exit Function
    End If
    
    With ws
        cm.Col_CustID = CLng(.Range("B2").value)
        cm.Col_Mobile = CLng(.Range("B3").value)
        cm.Col_Name = CLng(.Range("B4").value)
        cm.Col_Address = CLng(.Range("B5").value)
        cm.Col_Email = CLng(.Range("B6").value)
        cm.IsDetected = (UCase(Trim(.Range("B7").value)) = "YES")
    End With
    
    LoadColumnConfig = cm
End Function

'===========================================
' GLOBAL: Get Customer Data (ALL FORMS USE THIS)
'===========================================
Public Function GetCustomerData(custID As String, ByRef outName As String, ByRef outMobile As String) As Boolean
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim searchID As String
    
    ' Ensure config is loaded
    If Not gCustMap.IsDetected Then
        gCustMap = LoadColumnConfig()
    End If
    
    If Not gCustMap.IsDetected Then
        outName = "Unknown"
        outMobile = "0000000000"
        GetCustomerData = False
        Exit Function
    End If
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Customer_Master")
    On Error GoTo 0
    
    If ws Is Nothing Then
        outName = "Unknown"
        outMobile = "0000000000"
        GetCustomerData = False
        Exit Function
    End If
    
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    searchID = UCase(Trim(custID))
    
    For i = 2 To lastRow
        If UCase(Trim(CStr(ws.Cells(i, gCustMap.Col_CustID).value))) = searchID Then
            outName = Trim(CStr(ws.Cells(i, gCustMap.Col_Name).value))
            outMobile = Trim(CStr(ws.Cells(i, gCustMap.Col_Mobile).value))
            GetCustomerData = True
            Exit Function
        End If
    Next i
    
    ' Not found
    outName = "Unknown"
    outMobile = "0000000000"
    GetCustomerData = False
End Function

'===========================================
' INITIALIZE: Call this in Workbook_Open
'===========================================
Public Sub InitializeGlobalConfig()
    gCustMap = LoadColumnConfig()
    
    If Not gCustMap.IsDetected Then
        gCustMap = DetectCustomerMasterColumns()
    End If
    
    ' Validate
    If gCustMap.IsDetected Then
        Debug.Print "GLOBAL CONFIG: Customer_Master mapped successfully"
        Debug.Print "  Mobile Col: " & gCustMap.Col_Mobile
        Debug.Print "  Name Col: " & gCustMap.Col_Name
    Else
        Debug.Print "GLOBAL CONFIG: FAILED to detect Customer_Master columns!"
    End If
End Sub
