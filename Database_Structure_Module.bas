Attribute VB_Name = "Database_Structure_Module"
Option Explicit

'=================================
' JOB_PRODUCT TABLE STRUCTURE (22 COLUMNS)
'=================================
Public Enum JobProductColumns
    jpEntryID = 1           ' A - EntryID
    jpCustomerID = 2        ' B - CustomerID
    jpEntryType = 3         ' C - EntryType
    jpStatus = 4            ' D - Status
    jpLastStatusDate = 5    ' E - Last Status Date/Time
    jpProductType = 6       ' F - ProductType
    jpCompany = 7           ' G - Company
    jpModel = 8             ' H - Model
    jpSerialNumber = 9      ' I - SerialNumber
    jpPurchaseDate = 10     ' J - PurchaseDate
    jpWarrantyYear = 11     ' K - WarrantyYear
    jpWarrantyStatus = 12   ' L - WarrantyStatus
    jpProblem = 13          ' M - Problem
    jpRemark = 14           ' N - Remark
    jpPicture = 15          ' O - PICTURE
    jpCreateDate = 16       ' P - Creat Date and time
    jpVerifyType = 17       ' Q - Verify Type
    jpVerifyName = 18       ' R - Verify Name
    jpTotalAmount = 19      ' S - TotalAmount
    jpReceivedAmount = 20   ' T - ReceivedAmount
    jpPendingAmount = 21    ' U - PendingAmount
    jpPaymentMode = 22      ' V - PaymentMode
    jpReceivedBy = 23       ' W - ReceivedBy
End Enum

'=================================
' SAFE SAVE - Job_Product (PERFECT)
'=================================
Public Sub SaveJobProduct_Safe( _
    ByVal entryID As String, _
    ByVal customerID As String, _
    ByVal entryType As String, _
    ByVal productType As String, _
    ByVal company As String, _
    ByVal model As String, _
    ByVal serialNumber As String, _
    ByVal purchaseDate As String, _
    ByVal warrantyYear As String, _
    ByVal warrantyStatus As String, _
    ByVal problem As String, _
    ByVal remark As String, _
    Optional ByVal verifyType As String = "", _
    Optional ByVal verifyName As String = "", _
    Optional ByVal totalAmount As Double = 0, _
    Optional ByVal receivedAmount As Double = 0, _
    Optional ByVal pendingAmount As Double = 0, _
    Optional ByVal paymentMode As String = "", _
    Optional ByVal receivedBy As String = "")
    
    Dim ws As Worksheet
    Dim targetRow As Long
    Dim isNewRecord As Boolean
    
    On Error GoTo SaveError
    
    ' Worksheet set
    Set ws = ThisWorkbook.Sheets("Job_Product")
    
    ' Find or create row
    targetRow = FindJobProductRow(ws, entryID)
    isNewRecord = (targetRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row + 1)
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    ' === ALL 22 COLUMNS DATA ===
    With ws
        ' Basic Info (1-4)
        .Cells(targetRow, jpEntryID).value = entryID
        .Cells(targetRow, jpCustomerID).value = customerID
        .Cells(targetRow, jpEntryType).value = entryType
        .Cells(targetRow, jpStatus).value = IIf(isNewRecord, "ACTIVE", .Cells(targetRow, jpStatus).value)
        
        ' Status Date (5)
        .Cells(targetRow, jpLastStatusDate).value = Now
        
        ' Product Info (6-9)
        .Cells(targetRow, jpProductType).value = UCase(productType)
        .Cells(targetRow, jpCompany).value = UCase(company)
        .Cells(targetRow, jpModel).value = UCase(model)
        .Cells(targetRow, jpSerialNumber).value = UCase(serialNumber)
        
        ' Warranty Info (10-12)
        .Cells(targetRow, jpPurchaseDate).value = purchaseDate
        .Cells(targetRow, jpWarrantyYear).value = warrantyYear
        .Cells(targetRow, jpWarrantyStatus).value = warrantyStatus
        
        ' Problem/Remark (13-14)
        .Cells(targetRow, jpProblem).value = UCase(problem)
        .Cells(targetRow, jpRemark).value = UCase(remark)
        
        ' Picture placeholder (15)
        .Cells(targetRow, jpPicture).value = "Images\" & entryID
        
        ' Create Date (16) - ????? new record ???
        If isNewRecord Then
            .Cells(targetRow, jpCreateDate).value = Now
        End If
        
        ' Verify Info (17-18)
        .Cells(targetRow, jpVerifyType).value = verifyType
        .Cells(targetRow, jpVerifyName).value = verifyName
        
        ' Payment Info (19-23)
        .Cells(targetRow, jpTotalAmount).value = totalAmount
        .Cells(targetRow, jpReceivedAmount).value = receivedAmount
        .Cells(targetRow, jpPendingAmount).value = pendingAmount
        .Cells(targetRow, jpPaymentMode).value = paymentMode
        .Cells(targetRow, jpReceivedBy).value = receivedBy
    End With
    
    ' Formatting
    FormatJobProductRow ws, targetRow
    
    ' Save immediately
    ThisWorkbook.Save
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    LogAction "Job_Product Saved: " & entryID & " Row: " & targetRow
    Exit Sub
    
SaveError:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    LogError "SaveJobProduct Error: " & Err.Description & " EntryID: " & entryID
    MsgBox "SAVE FAILED: " & Err.Description, vbCritical
End Sub

'=================================
' FIND ROW (Edit mode ?? ???)
'=================================
Private Function FindJobProductRow(ws As Worksheet, entryID As String) As Long
    Dim f As Range
    
    If entryID = "" Then
        FindJobProductRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row + 1
        Exit Function
    End If
    
    Set f = ws.Columns(1).Find(entryID, LookAt:=xlWhole)
    
    If f Is Nothing Then
        FindJobProductRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row + 1
    Else
        FindJobProductRow = f.row
    End If
End Function

'=================================
' ROW FORMATTING
'=================================
Private Sub FormatJobProductRow(ws As Worksheet, r As Long)
    With ws.Rows(r)
        .VerticalAlignment = xlCenter
        .HorizontalAlignment = xlLeft
    End With
    
    ' Date columns format
    ws.Cells(r, jpLastStatusDate).NumberFormat = "dd-mm-yyyy hh:mm AM/PM"
    ws.Cells(r, jpCreateDate).NumberFormat = "dd-mm-yyyy hh:mm AM/PM"
    ws.Cells(r, jpPurchaseDate).NumberFormat = "dd-mm-yyyy"
    
    ' Amount columns format
    ws.Cells(r, jpTotalAmount).NumberFormat = "0.00"
    ws.Cells(r, jpReceivedAmount).NumberFormat = "0.00"
    ws.Cells(r, jpPendingAmount).NumberFormat = "0.00"
End Sub

'=================================
' LOAD DATA (Edit mode ?? ???)
'=================================
Public Function LoadJobProduct(entryID As String) As Variant
    Dim ws As Worksheet
    Dim f As Range
    Dim data(1 To 23) As Variant
    
    Set ws = ThisWorkbook.Sheets("Job_Product")
    Set f = ws.Columns(1).Find(entryID, LookAt:=xlWhole)
    
    If f Is Nothing Then
        LoadJobProduct = Empty
        Exit Function
    End If
    
    Dim i As Integer
    For i = 1 To 23
        data(i) = ws.Cells(f.row, i).value
    Next i
    
    LoadJobProduct = data
End Function





'=================================
' PAYMENT_MASTER TABLE STRUCTURE (8 COLUMNS)
'=================================
Public Enum PaymentMasterColumns
    pmPaymentID = 1         ' A - PaymentID
    pmEntryID = 2           ' B - EntryID
    pmCustomerID = 3        ' C - CustomerID
    pmTransType = 4         ' D - TransType (PAYMENT/EXPENSE/RECEIPT)
    pmTransMode = 5         ' E - TransMode
    pmAmount = 6            ' F - Amount
    pmTransDate = 7         ' G - TransDate
    pmRemark = 8            ' H - Remark
End Enum




'=================================
' LOAD PAYMENT FOR EDIT
'=================================
Public Function LoadPaymentData(paymentID As String) As Collection
    Dim ws As Worksheet
    Dim r As Long
    Dim coll As New Collection
    Dim data(1 To 6) As Variant
    
    ' Column Constants - Payment_Master Sheet
    Const pmPaymentID As Integer = 1   ' Column A
    Const pmEntryID As Integer = 2     ' Column B
    Const pmCustomerID As Integer = 3  ' Column C
    Const pmTransType As Integer = 4   ' Column D
    Const pmTransMode As Integer = 5   ' Column E
    Const pmAmount As Integer = 6      ' Column F
    Const pmRemark As Integer = 7      ' Column G

    Set ws = ThisWorkbook.Sheets("Payment_Master")

    For r = 2 To ws.Cells(ws.Rows.count, 1).End(xlUp).row
        If ws.Cells(r, pmPaymentID).value = paymentID Then
            data(1) = ws.Cells(r, pmEntryID).value
            data(2) = ws.Cells(r, pmTransType).value
            data(3) = ws.Cells(r, pmTransMode).value
            data(4) = ws.Cells(r, pmAmount).value
            data(5) = ws.Cells(r, pmRemark).value
            data(6) = ws.Cells(r, pmCustomerID).value

            coll.Add data
        End If
    Next r

    Set LoadPaymentData = coll
End Function

'=================================
' GENERATE PERFECT ENTRY ID
'=================================
Public Function GeneratePerfectEntryID(entryType As String) As String
    Dim ws As Worksheet
    Dim lastRow As Long, r As Long
    Dim maxNum As Long, currentNum As Long
    Dim prefix As String
    Dim newID As String
    
    ' Prefix set karo
    If UCase(entryType) = "WARRANTY" Then
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
    
    ' Double check unique (if duplicate, add timestamp)
    If EntryExists_Check(newID) Then
        newID = prefix & Format(maxNum + 1, "00000") & "_" & Format(Now, "hhmmss")
    End If
    
    GeneratePerfectEntryID = newID
End Function

'=================================
' ENTRY EXISTS CHECK
'=================================
Private Function EntryExists_Check(newID As String) As Boolean
    Dim ws As Worksheet
    Dim r As Long
    
    Set ws = ThisWorkbook.Sheets("Job_Product")
    
    For r = 2 To ws.Cells(ws.Rows.count, 1).End(xlUp).row
        If UCase(ws.Cells(r, 1).value) = UCase(newID) Then
            EntryExists_Check = True
            Exit Function
        End If
    Next r
    
    EntryExists_Check = False
End Function
'=================================
' GENERATE PAYMENT ID
'=================================
Public Function GeneratePaymentID() As String
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim maxNum As Long
    Dim r As Long
    
    Set ws = ThisWorkbook.Sheets("Payment_Master")
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    
    maxNum = 0
    
    ' Find max number
    For r = 2 To lastRow
        If Left(ws.Cells(r, 1).value, 3) = "PAY" Then
            Dim currentNum As Long
            currentNum = val(Mid(ws.Cells(r, 1).value, 4))
            If currentNum > maxNum Then maxNum = currentNum
        End If
    Next r
    
    GeneratePaymentID = "PAY" & Format(maxNum + 1, "00000")
End Function

'=================================
' DELETE OLD PAYMENTS (For Edit Mode)
'=================================
Public Sub DeleteOldPayments(paymentID As String)
    Dim ws As Worksheet
    Dim r As Long
    
    Set ws = ThisWorkbook.Sheets("Payment_Master")
    
    ' Column A (1) ??? PaymentID ?? ??? ???? Delete ???
    For r = ws.Cells(ws.Rows.count, 1).End(xlUp).row To 2 Step -1
        If ws.Cells(r, 1).value = paymentID Then
            ws.Rows(r).Delete
        End If
    Next r
    
    LogAction "Old payment deleted: " & paymentID
End Sub
'=================================
' SAVE CONSOLIDATED PAYMENT (ONE LINE)
'=================================
Public Sub SavePayment_Consolidated( _
    ByVal paymentID As String, _
    ByVal entryID As String, _
    ByVal customerID As String, _
    ByVal chargeNames As String, _
    ByVal totalReceipt As Double, _
    ByVal payModes As String, _
    ByVal totalPayment As Double, _
    ByVal expTypes As String, _
    ByVal totalExpense As Double, _
    ByVal remark As String)
    
    Dim ws As Worksheet
    Dim r As Long
    Dim dueAmt As Double
    
    On Error GoTo SaveError
    
    Set ws = ThisWorkbook.Sheets("Payment_Master")
    r = ws.Cells(ws.Rows.count, 1).End(xlUp).row + 1
    
    dueAmt = (totalPayment + totalExpense) - totalReceipt
    
    With ws
        .Cells(r, 1).value = paymentID      ' A: PAY00001
        .Cells(r, 2).value = entryID        ' B: WAR00001
        .Cells(r, 3).value = customerID     ' C: CUST0001
        .Cells(r, 4).value = chargeNames    ' D: Service Charge
        
        ' === ???? Swap ??? ===
        .Cells(r, 5).value = totalPayment   ' E: 1000 (Charges Total)
        .Cells(r, 6).value = payModes       ' F: Cash, PhonePe
        .Cells(r, 7).value = totalReceipt   ' G: 300 (Receipt Total)
        ' ======================
        
        .Cells(r, 8).value = expTypes       ' H: Transport
        .Cells(r, 9).value = totalExpense   ' I: 100
        .Cells(r, 10).value = dueAmt        ' J: 800
        .Cells(r, 11).value = Now           ' K: Date
        .Cells(r, 11).NumberFormat = "dd-mm-yyyy hh:mm AM/PM"
        .Cells(r, 12).value = remark        ' L: Remarks
        
        ' Formatting
        .Cells(r, 5).NumberFormat = "0.00"
        .Cells(r, 7).NumberFormat = "0.00"
        .Cells(r, 9).NumberFormat = "0.00"
        .Cells(r, 10).NumberFormat = "0.00"
    End With
    
    ThisWorkbook.Save
    
    LogAction "Payment Saved: " & paymentID
    Exit Sub
    
SaveError:
    LogError "SavePayment Error: " & Err.Description
    MsgBox "Payment Save Failed: " & Err.Description, vbCritical
End Sub
'=================================
' GET PAYMENT ID BY ENTRY ID (?? ADD ????)
'=================================
Public Function GetPaymentIDByEntry(entryID As String) As String
    Dim ws As Worksheet
    Dim f As Range
    
    Set ws = ThisWorkbook.Sheets("Payment_Master")
    Set f = ws.Columns(2).Find(entryID, LookIn:=xlValues, LookAt:=xlWhole)
    
    If Not f Is Nothing Then
        GetPaymentIDByEntry = ws.Cells(f.row, 1).value
    End If
End Function
Private Sub LogAction(msg As String)
    ' Empty or call PublicModule's LogAction
    On Error Resume Next
    Application.Run "PublicModule.LogAction", msg
    On Error GoTo 0
End Sub

Private Sub LogError(msg As String)
    Debug.Print "ERROR: " & msg
End Sub
