VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDelivery 
   Caption         =   "Customer Delivery"
   ClientHeight    =   13710
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   19755
   OleObjectBlob   =   "frmDelivery.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmDelivery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'========== MODULE VARIABLES ==========
Private mEntryID As String
Private mAssignID As String
Private mCustomerID As String
Private mPhotoPath As String
Private isLoading As Boolean
Private mIsDelivered As Boolean
Private mCurrentOTP As String
Private mOTPTimestamp As Date
Private mIsDeliveryConfirmed As Boolean
Private m_DeliveryEntryID As String
Private mDeliveryEntryID As String

Private Declare PtrSafe Function ShellExecute Lib "shell32.dll" _
    Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, _
    ByVal lpFile As String, ByVal lpParameters As String, _
    ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'=====================================================
' USERFORM INITIALIZE
'=====================================================
Private Sub UserForm_Initialize()
    Me.caption = "PRODUCT DELIVERY - GLOBAL SOFT"
    txtReturnDate.value = Format(Date, "dd-mm-yyyy")
    mPhotoPath = ""
    mIsDelivered = False
    isLoading = False
    
    AddMinMaxButtons Me
    Call SetupPaymentListView
    Call SetBottomButtonsEnabled(False)
    
    With cmbSelectType
        .Clear
        .AddItem "ALL"
        .AddItem "ACTIVE"
        .AddItem "COMPLETED"
        .AddItem "PENDING"
        .AddItem "RETURN"
        .AddItem "REJECTED"
        .AddItem "DELIVERED"
        .ListIndex = 0
    End With

    cmbEntryID.Clear
    LoadReturnModes
    LoadReturnNames
    ClearAllFields
    
    Dim entryID As String
    entryID = Trim(Me.Tag)
    If entryID <> "" Then
        Me.caption = "PRODUCT DELIVERY - EDIT (" & entryID & ")"
        Me.Tag = ""
    End If
    
        ' ===== SAME CUSTOMER CHECKBOX — SAFE INITIALIZE =====
    On Error Resume Next
   With Me.chkSameCustomer
        .value = 0
        .visible = True
        .enabled = True
        .caption = "Delivered By Same Customer"  ' <-- NEW CAPTION
    End With
    
    ' Form khulte hi entries load karo
    LoadEntriesByType "ALL"
End Sub
Private Sub cmbSelectType_Change()
    If Trim(cmbSelectType.value & "") = "" Then Exit Sub
    LoadEntriesByType cmbSelectType.value
End Sub

'========== PARSE ENTRY ID ==========
Private Function GetEntryIDFromDisplay(displayText As String) As String
    If displayText = "" Then
        GetEntryIDFromDisplay = ""
        Exit Function
    End If
    Dim pos As Long
    pos = InStr(displayText, " (")
    If pos > 0 Then
        GetEntryIDFromDisplay = Trim(Left(displayText, pos - 1))
    Else
        GetEntryIDFromDisplay = Trim(displayText)
    End If
End Function

'========== LOAD ENTRIES BY TYPE — FINAL FIXED VERSION ==========
Private Sub LoadEntriesByType(entryType As String)
    Dim wsJob As Worksheet, wsAssign As Worksheet
    Dim lastRow As Long, i As Long
    Dim entryID As String, status As String
    Dim addedCount As Long
    Dim jobRows As Long, assignRows As Long
    Dim bAdd As Boolean
    Dim uniqueList As String
    Dim displayText As String
    
    uniqueList = ","  ' Unique EntryID tracker (comma separated)
    
    On Error Resume Next
    Set wsJob = ThisWorkbook.Sheets("Job_Product")
    Set wsAssign = ThisWorkbook.Sheets("Assign_Master")
    On Error GoTo 0
    
    isLoading = True
    cmbEntryID.Clear
    mIsDelivered = False
    addedCount = 0
    jobRows = 0
    assignRows = 0
    
    ' ===== FIRST: LOAD from Job_Product (Col 1=EntryID, Col 4=Status) =====
    If Not wsJob Is Nothing Then
        lastRow = wsJob.Cells(wsJob.Rows.count, 1).End(xlUp).row
        jobRows = lastRow - 1
        
        For i = 2 To lastRow
            entryID = Trim(wsJob.Cells(i, 1).value & "")
            status = UCase(Trim(wsJob.Cells(i, 4).value & ""))
            
            If entryID = "" Then GoTo NextJobRow
            
            bAdd = False
            Select Case UCase(Trim(entryType))
                Case "ALL"
    ' ????? ?? entries ?? Delivery ?? ??? Ready ???? ???
    If status = "ACTIVE" Or status = "PENDING" Or status = "IN PROGRESS" Or _
       status = "ASSIGNED" Or status = "SENT" Or status = "PENDING PARTS" Or _
       status = "RETURNED" Or status = "REJECTED" Or status = "" Then
        bAdd = True
    End If
                Case Else
                    If status = UCase(Trim(entryType)) Then bAdd = True
            End Select
            
            If bAdd Then
                displayText = entryID & " (" & IIf(status = "", "NO STATUS", status) & ")"
                ' Sirf unique entry add karo
                If InStr(uniqueList, "," & entryID & ",") = 0 Then
                    cmbEntryID.AddItem displayText
                    uniqueList = uniqueList & entryID & ","
                    addedCount = addedCount + 1
                End If
            End If
NextJobRow:
        Next i
    End If
    
    ' ===== SECOND: LOAD from Assign_Master (Col 2=EntryID, Col 32=AssignStatus) =====
    If Not wsAssign Is Nothing Then
        lastRow = wsAssign.Cells(wsAssign.Rows.count, 1).End(xlUp).row
        assignRows = lastRow - 1
        
        For i = 2 To lastRow
            entryID = Trim(wsAssign.Cells(i, 2).value & "")
            status = UCase(Trim(wsAssign.Cells(i, 32).value & ""))
            
            If entryID = "" Then GoTo NextAssignRow
            
            bAdd = False
            Select Case UCase(Trim(entryType))
                Case "ALL"
                    If status = "RETURNED" Or status = "PENDING" Or status = "COMPLETED" Or status = "ACTIVE" Or status = "" Then
                        bAdd = True
                    End If
                Case "DELIVERED"
                    ' Skip — Assign_Master me DELIVERED nahi hota
                Case "ACTIVE", "PENDING"
                    If status = UCase(Trim(entryType)) Or status = "" Then bAdd = True
                Case Else
                    If status = UCase(Trim(entryType)) Then bAdd = True
            End Select
            
            If bAdd Then
                ' Sirf tab add karo jab pehle se Job_Product me nahi mila
                If InStr(uniqueList, "," & entryID & ",") = 0 Then
                    displayText = entryID & " (" & IIf(status = "", "PENDING", status) & ")"
                    cmbEntryID.AddItem displayText
                    uniqueList = uniqueList & entryID & ","
                    addedCount = addedCount + 1
                End If
            End If
NextAssignRow:
        Next i
    End If
    
    isLoading = False
    
    If cmbEntryID.ListCount = 0 Then
        MsgBox "Koi entry nahi mili!" & vbCrLf & _
               "Filter Type: " & entryType & vbCrLf & _
               "Job_Product me rows: " & jobRows & vbCrLf & _
               "Assign_Master me rows: " & assignRows & vbCrLf & _
               "Tip: Job_Product ke Column D (Status) me ACTIVE/PENDING hona chahiye.", vbExclamation
    End If
End Sub
Private Sub cmbEntryID_Change()
    If isLoading Then Exit Sub
    
    Dim rawEntryID As String
    rawEntryID = GetEntryIDFromDisplay(cmbEntryID.value)
    
    If rawEntryID = "" Then Exit Sub
    
    mEntryID = rawEntryID
    m_DeliveryEntryID = rawEntryID
    
    Dim wsDel As Worksheet
    On Error Resume Next
    Set wsDel = ThisWorkbook.Sheets("Delivery_Master")
    On Error GoTo 0
    mIsDelivered = IsDelivered(mEntryID, wsDel)
    
    If mIsDelivered Then
        btnDeliver.enabled = False
        btnDeliver.caption = "ALREADY DELIVERED"
    Else
        btnDeliver.enabled = True
        btnDeliver.caption = "DELIVER"
    End If
    
    LoadEntryDetailsUniversal mEntryID
    LoadWorkDone mEntryID
    LoadAccessories mEntryID
    
    If mIsDelivered Then
        LoadDeliveryDetails mEntryID
    End If
    
        ' PAYMENT DETAILS LOAD
    Call LoadPaymentDetails(mEntryID)
End Sub

'========== LOAD DELIVERY DETAILS (Already Delivered) ==========
Private Sub LoadDeliveryDetails(entryID As String)
    Dim wsDel As Worksheet
    Dim lastRow As Long, i As Long
    
    On Error Resume Next
    Set wsDel = ThisWorkbook.Sheets("Delivery_Master")
    On Error GoTo 0
    If wsDel Is Nothing Then Exit Sub
    
    lastRow = wsDel.Cells(wsDel.Rows.count, 1).End(xlUp).row
    For i = 2 To lastRow
        If Trim(wsDel.Cells(i, 2).value) = entryID Then
            Dim retMode As String, retName As String
            retMode = Trim(wsDel.Cells(i, 9).value & "")
            retName = Trim(wsDel.Cells(i, 10).value & "")
            
            If retMode <> "" Then
                If IsInCombo(cmbReturnMode, retMode) Then
                    cmbReturnMode.value = retMode
                Else
                    cmbReturnMode.AddItem retMode
                    cmbReturnMode.value = retMode
                End If
            End If
            
            If retName <> "" Then
                If IsInCombo(cmbReturnName, retName) Then
                    cmbReturnName.value = retName
                Else
                    cmbReturnName.AddItem retName
                    cmbReturnName.value = retName
                End If
            End If
            
            txtReturnMobile.value = wsDel.Cells(i, 11).value & ""
            txtReturnAddress.value = wsDel.Cells(i, 12).value & ""
            txtReturnDocket.value = wsDel.Cells(i, 13).value & ""
            txtReturnDate.value = wsDel.Cells(i, 14).value & ""
            txtRemarks.value = wsDel.Cells(i, 16).value & ""
            
            Dim photoPath As String
            photoPath = Trim(wsDel.Cells(i, 15).value & "")
            If photoPath <> "" And Dir(photoPath) <> "" Then
                On Error Resume Next
                imgDeliveryPhoto.Picture = LoadPicture(photoPath)
                If Err.Number <> 0 Then Err.Clear
                On Error GoTo 0
                mPhotoPath = photoPath
            End If
            
            Exit For
        End If
    Next i
End Sub

'========== LOAD ENTRY DETAILS + PHOTOS ==========
Private Sub LoadEntryDetailsUniversal(entryID As String)
    Dim wsJob As Worksheet, wsAssign As Worksheet, wsCust As Worksheet
    Dim lastRow As Long, i As Long
    Dim foundInJob As Boolean, foundInAssign As Boolean
    Dim custPhotoName As String
    
    Set wsJob = ThisWorkbook.Sheets("Job_Product")
    Set wsAssign = ThisWorkbook.Sheets("Assign_Master")
    Set wsCust = ThisWorkbook.Sheets("Customer_Master")
    
    foundInJob = False
    foundInAssign = False
    mAssignID = ""
    mCustomerID = ""
    custPhotoName = ""
    
    ' === Search Job_Product ===
    lastRow = wsJob.Cells(wsJob.Rows.count, 1).End(xlUp).row
    For i = 2 To lastRow
        If Trim(wsJob.Cells(i, 1).value) = entryID Then
            foundInJob = True
            mCustomerID = Trim(wsJob.Cells(i, 2).value)
            
            lblEntryIDValue.caption = entryID
            lblProductValue.caption = wsJob.Cells(i, 6).value
            lblCompanyValue.caption = wsJob.Cells(i, 7).value
            lblModelValue.caption = wsJob.Cells(i, 8).value
            lblOriginalSerialValue.caption = wsJob.Cells(i, 9).value
            lblFinalSerialValue.caption = wsJob.Cells(i, 9).value
            lblWarrantyStatusValue.caption = wsJob.Cells(i, 12).value
            Exit For
        End If
    Next i
    
    ' === Search Assign_Master ===
    lastRow = wsAssign.Cells(wsAssign.Rows.count, 1).End(xlUp).row
    For i = 2 To lastRow
        If Trim(wsAssign.Cells(i, 2).value) = entryID Then
            foundInAssign = True
            mAssignID = wsAssign.Cells(i, 1).value
            If mCustomerID = "" Then mCustomerID = Trim(wsAssign.Cells(i, 3).value)
            
            If Not foundInJob Then
                lblEntryIDValue.caption = entryID
                lblProductValue.caption = wsAssign.Cells(i, 6).value
                lblCompanyValue.caption = wsAssign.Cells(i, 7).value
                lblModelValue.caption = wsAssign.Cells(i, 8).value
                lblOriginalSerialValue.caption = wsAssign.Cells(i, 9).value
                lblFinalSerialValue.caption = wsAssign.Cells(i, 9).value
                lblWarrantyStatusValue.caption = wsAssign.Cells(i, 10).value
            End If
            
            'If Not mIsDelivered Then
              '  Dim retMode As String, retName As String
             '   retMode = Trim(wsAssign.Cells(i, 23).value & "")
              '  retName = Trim(wsAssign.Cells(i, 24).value & "")
                
              '  If retMode <> "" Then
              '      If IsInCombo(cmbReturnMode, retMode) Then
              '          cmbReturnMode.value = retMode
              '      Else
               '         cmbReturnMode.AddItem retMode
              '          cmbReturnMode.value = retMode
               '     End If
            '    Else
             '       cmbReturnMode.value = ""
             '   End If
             '
             '   If retName <> "" Then
              '      If IsInCombo(cmbReturnName, retName) Then
               '         cmbReturnName.value = retName
               '     Else
                '        cmbReturnName.AddItem retName
               '         cmbReturnName.value = retName
                '    End If
              ' Else
             '       cmbReturnName.value = ""
              '  End If
                
              '  txtReturnDocket.value = wsAssign.Cells(i, 25).value & ""
             '   txtReturnDate.value = wsAssign.Cells(i, 26).value & ""
             '   txtReturnMobile.value = wsAssign.Cells(i, 15).value & ""
            '    txtReturnAddress.value = wsAssign.Cells(i, 17).value & ""
          '  End If
            
            Exit For
        End If
    Next i
    
    ' === Customer Details + Photo ===
If mCustomerID <> "" Then
    lastRow = wsCust.Cells(wsCust.Rows.count, 1).End(xlUp).row
    For i = 2 To lastRow
        If Trim(wsCust.Cells(i, 1).value) = mCustomerID Then
            
            ' ?????? ?? ?? lines SWAP ???
             lblCustomerValue.caption = wsCust.Cells(i, 2).value   ' Name (Column B)
            lblMobileValue.caption = wsCust.Cells(i, 3).value     ' Mobile (Column
            lblAddressValue.caption = wsCust.Cells(i, 4).value     ' Address (Column D)
            
            custPhotoName = Trim(wsCust.Cells(i, 7).value & "")
            Exit For
        End If
    Next i
End If
    
    ' === LOAD CUSTOMER PHOTO ===
    If custPhotoName <> "" Then
        LoadCustomerPhoto custPhotoName
    Else
        On Error Resume Next
        Set imgCustomer.Picture = Nothing
        On Error GoTo 0
    End If
    
    ' === LOAD PRODUCT PHOTO ===
    LoadProductPhoto entryID
    
    If Not foundInJob And Not foundInAssign Then
        MsgBox "Entry record nahi mila kisi bhi sheet mein!", vbExclamation
    End If
End Sub

'========== LOAD CUSTOMER PHOTO ==========
Private Sub LoadCustomerPhoto(photoName As String)
    Dim folderPath As String
    Dim fullPath As String
    
    On Error Resume Next
    
    folderPath = Trim(ThisWorkbook.Sheets("Settings").Range("B3").value)
    If folderPath = "" Then
        folderPath = ThisWorkbook.path & "\GLOBAL_SOFT_DATA\Images\Customers\"
    End If
    
    If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
    
    fullPath = folderPath & photoName
    
    If Dir(fullPath) = "" Then
        fullPath = ThisWorkbook.path & "\GLOBAL_SOFT_DATA\Images\Customers\" & photoName
    End If
    
    If Dir(fullPath) <> "" Then
        imgCustomer.Picture = LoadPicture(fullPath)
    Else
        Set imgCustomer.Picture = Nothing
    End If
    
    On Error GoTo 0
End Sub

'========== LOAD PRODUCT PHOTO ==========
Private Sub LoadProductPhoto(entryID As String)
    Dim basePath As String
    Dim folderPath As String
    Dim photoPath As String
    
    On Error Resume Next
    
    basePath = Trim(ThisWorkbook.Sheets("Settings").Range("B2").value)
    If basePath = "" Then
        basePath = ThisWorkbook.path & "\GLOBAL_SOFT_DATA\Images\Products"
    End If
    
    If Right(basePath, 1) <> "\" Then basePath = basePath & "\"
    
    folderPath = basePath & entryID & "\"
    photoPath = folderPath & "Photo1.jpg"
    
    If Dir(photoPath) = "" Then
        photoPath = ThisWorkbook.path & "\GLOBAL_SOFT_DATA\Images\Products\" & entryID & "\Photo1.jpg"
    End If
    
    If Dir(photoPath) <> "" Then
        imgProductAfter.Picture = LoadPicture(photoPath)
    Else
        Set imgProductAfter.Picture = Nothing
    End If
    
    On Error GoTo 0
End Sub

'========== LOAD WORK DONE — WARRANTY + SERVICE BOTH ==========
Private Sub LoadWorkDone(entryID As String)
    Dim ws As Worksheet, lastRow As Long, i As Long
    Dim allText As String
    
    allText = ""
    
    ' === WARRANTY ENTRY (WAR*) ===
    If UCase(Left(entryID, 3)) = "WAR" Then
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets("Warranty_Return_Master")
        On Error GoTo 0
        
        If Not ws Is Nothing Then
            lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
            For i = 2 To lastRow
                If Trim(ws.Cells(i, 3).value) = entryID Then
                    Dim retStatus As String, oldSerial As String, newSerial As String
                    Dim rejectReason As String, verifiedBy As String
                    Dim returnDate As String, vendorName As String
                    
                    retStatus = Trim(ws.Cells(i, 19).value & "")
                    oldSerial = Trim(ws.Cells(i, 9).value & "")
                    newSerial = Trim(ws.Cells(i, 21).value & "")
                    rejectReason = Trim(ws.Cells(i, 22).value & "")
                    verifiedBy = Trim(ws.Cells(i, 20).value & "")
                    returnDate = Trim(ws.Cells(i, 17).value & "")
                    vendorName = Trim(ws.Cells(i, 10).value & "")
                    
                    allText = "WARRANTY RETURN STATUS: " & retStatus
                    If returnDate <> "" Then allText = allText & vbCrLf & "Return Date: " & returnDate
                    If vendorName <> "" Then allText = allText & vbCrLf & "Vendor: " & vendorName
                    
                    If UCase(retStatus) = "RECEIVE_REPAIRED" Then
                        If verifiedBy <> "" Then allText = allText & vbCrLf & "Verified By: " & verifiedBy
                        allText = allText & vbCrLf & "Serial: " & oldSerial & " (Same Serial Repaired)"
                    ElseIf UCase(retStatus) = "RECEIVE_REPLACED" Then
                        allText = allText & vbCrLf & "Old Serial: " & oldSerial
                        If newSerial <> "" Then allText = allText & vbCrLf & "New Serial: " & newSerial
                    ElseIf UCase(retStatus) = "RECEIVE_REJECTED" Then
                        If rejectReason <> "" Then allText = allText & vbCrLf & "Reject Reason: " & rejectReason
                    End If
                    
                    Exit For
                End If
            Next i
        End If
        
        If allText = "" Then
            On Error Resume Next
            Set ws = ThisWorkbook.Sheets("Assign_Master")
            On Error GoTo 0
            If Not ws Is Nothing Then
                lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
                For i = 2 To lastRow
                    If Trim(ws.Cells(i, 2).value) = entryID Then
                        Dim assignStatus As String, sendDate As String, expectedDate As String
                        assignStatus = Trim(ws.Cells(i, 32).value & "")
                        sendDate = Trim(ws.Cells(i, 26).value & "")
                        expectedDate = Trim(ws.Cells(i, 27).value & "")
                        
                        allText = "WARRANTY STATUS: " & assignStatus
                        If sendDate <> "" Then allText = allText & vbCrLf & "Sent Date: " & sendDate
                        If expectedDate <> "" Then allText = allText & vbCrLf & "Expected Return: " & expectedDate
                        Exit For
                    End If
                Next i
            End If
        End If
        
    ' === SERVICE ENTRY (SER*) ===
    ElseIf UCase(Left(entryID, 3)) = "SER" Then
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets("Service_WorkLog_Expense")
        On Error GoTo 0
        
        If Not ws Is Nothing Then
            lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
            For i = 2 To lastRow
                If Trim(ws.Cells(i, 3).value) = entryID Then
                    Dim problemFound As String, solutionApplied As String
                    Dim partsUsed As String, laborChg As String
                    Dim transportChg As String, otherChg As String
                    Dim totalExp As String, workStatus As String
                    
                    problemFound = Trim(ws.Cells(i, 14).value & "")
                    solutionApplied = Trim(ws.Cells(i, 15).value & "")
                    partsUsed = Trim(ws.Cells(i, 16).value & "")
                    laborChg = Trim(ws.Cells(i, 17).value & "")
                    transportChg = Trim(ws.Cells(i, 18).value & "")
                    otherChg = Trim(ws.Cells(i, 19).value & "")
                    totalExp = Trim(ws.Cells(i, 20).value & "")
                    workStatus = Trim(ws.Cells(i, 13).value & "")
                    
                    allText = "SERVICE WORK STATUS: " & workStatus
                    If problemFound <> "" Then allText = allText & vbCrLf & "Problem Found: " & problemFound
                    If solutionApplied <> "" Then allText = allText & vbCrLf & "Solution Applied: " & solutionApplied
                    If partsUsed <> "" Then allText = allText & vbCrLf & "Parts Used: " & partsUsed
                    
                    Dim expenseText As String
                    expenseText = ""
                    If laborChg <> "" And laborChg <> "0" Then expenseText = expenseText & "Labor: Rs." & laborChg & " | "
                    If transportChg <> "" And transportChg <> "0" Then expenseText = expenseText & "Transport: Rs." & transportChg & " | "
                    If otherChg <> "" And otherChg <> "0" Then expenseText = expenseText & "Other: Rs." & otherChg & " | "
                    If totalExp <> "" And totalExp <> "0" Then expenseText = expenseText & "TOTAL: Rs." & totalExp
                    
                    If expenseText <> "" Then
                        allText = allText & vbCrLf & "Expenses: " & expenseText
                    End If
                    
                    Exit For
                End If
            Next i
        End If
        
        If allText = "" Then
            On Error Resume Next
            Set ws = ThisWorkbook.Sheets("Assign_Master")
            On Error GoTo 0
            If Not ws Is Nothing Then
                lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
                For i = 2 To lastRow
                    If Trim(ws.Cells(i, 2).value) = entryID Then
                        Dim serStatus As String
                        serStatus = Trim(ws.Cells(i, 32).value & "")
                        allText = "SERVICE STATUS: " & serStatus & " (Work Log Not Created Yet)"
                        Exit For
                    End If
                Next i
            End If
        End If
    End If
    
    txtWorkDone.value = allText
End Sub

Private Sub LoadAccessories(entryID As String)
    Dim ws As Worksheet, lastRow As Long, i As Long
    Dim accName As String, cbIndex As Long
    Dim ctrl As Control, foundCtrl As Control
    
    ' ===== FIX: Sirf accessory checkboxes hide karo =====
    ' chkSameCustomer ko protect karo
    For Each ctrl In Me.Controls
        If TypeName(ctrl) = "CheckBox" Then
            If ctrl.name <> "chkSameCustomer" Then
                ctrl.caption = ""
                ctrl.value = False
                ctrl.visible = False
            End If
        End If
    Next ctrl
    
    ' ? PURANA LOOP DELETE KIYA GAYA — ab sirf ek loop hai
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Job_Accessory")
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    cbIndex = 1
    For i = 2 To lastRow
        If Trim(ws.Cells(i, 1).value) = entryID Then
            accName = Trim(ws.Cells(i, 4).value)
            If accName <> "" And cbIndex <= 11 Then
                If cbIndex = 6 Then cbIndex = 7
                Set foundCtrl = Nothing
                On Error Resume Next
                Set foundCtrl = Me.Controls("CheckBox" & cbIndex)
                On Error GoTo 0
                If Not foundCtrl Is Nothing Then
                    foundCtrl.caption = " " & accName
                    foundCtrl.visible = True
                    foundCtrl.value = False
                    cbIndex = cbIndex + 1
                End If
            End If
        End If
    Next i
End Sub

'========== RETURN MODE/NAME ==========
Private Sub LoadReturnModes()
    Dim ws As Worksheet, lastRow As Long, i As Long
    On Error Resume Next: Set ws = ThisWorkbook.Sheets("Delivery_data"): On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    cmbReturnMode.Clear
    For i = 2 To lastRow
        If Trim(ws.Cells(i, 2).value) <> "" Then cmbReturnMode.AddItem Trim(ws.Cells(i, 2).value)
    Next i
End Sub

Private Sub LoadReturnNames()
    Dim ws As Worksheet, lastRow As Long, i As Long
    On Error Resume Next: Set ws = ThisWorkbook.Sheets("Delivery_data"): On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    cmbReturnName.Clear
    For i = 2 To lastRow
        If Trim(ws.Cells(i, 3).value) <> "" Then cmbReturnName.AddItem Trim(ws.Cells(i, 3).value)
    Next i
End Sub

Private Sub cmbReturnName_Change()
    If cmbReturnName.value = "" Then Exit Sub
    Dim ws As Worksheet, lastRow As Long, i As Long
    Set ws = ThisWorkbook.Sheets("Delivery_data")
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    For i = 2 To lastRow
        If Trim(ws.Cells(i, 3).value) = cmbReturnName.value Then
            txtReturnMobile.value = Trim(ws.Cells(i, 4).value)
            txtReturnAddress.value = Trim(ws.Cells(i, 5).value)
            Exit For
        End If
    Next i
End Sub

'========== ADD RETURN MODE/NAME ==========
Private Sub btnAddReturnMode_Click()
    Dim newMode As String
    newMode = InputBox("Naya Return Mode likho:", "Add Return Mode")
    If newMode <> "" Then
        Dim ws As Worksheet, lr As Long
        Set ws = ThisWorkbook.Sheets("Delivery_data")
        lr = ws.Cells(ws.Rows.count, 1).End(xlUp).row + 1
        ws.Cells(lr, 1).value = "MODE" & Format(lr - 1, "000")
        ws.Cells(lr, 2).value = newMode
        cmbReturnMode.AddItem newMode
        cmbReturnMode.value = newMode
    End If
End Sub

Private Sub btnAddReturnName_Click()
    Dim newName As String
    newName = InputBox("Naya Courier Name likho:", "Add Courier Name")
    If newName = "" Then Exit Sub
    Dim mob As String, addr As String
    mob = InputBox("Mobile Number:", "Courier Details")
    addr = InputBox("Address:", "Courier Details")
    Dim ws As Worksheet, lr As Long
    Set ws = ThisWorkbook.Sheets("Delivery_data")
    lr = ws.Cells(ws.Rows.count, 1).End(xlUp).row + 1
    ws.Cells(lr, 1).value = "COUR" & Format(lr - 1, "000")
    ws.Cells(lr, 3).value = newName
    ws.Cells(lr, 4).value = mob
    ws.Cells(lr, 5).value = addr
    cmbReturnName.AddItem newName
    cmbReturnName.value = newName
    txtReturnMobile.value = mob
    txtReturnAddress.value = addr
End Sub

Private Sub btnDeliveryPhoto_Click()
    Dim picPath As Variant
    picPath = Application.GetOpenFilename( _
        FileFilter:="Image Files (*.jpg;*.jpeg;*.png;*.bmp), *.jpg;*.jpeg;*.png;*.bmp", _
        title:="Delivery Photo Select Karo")
    If picPath <> False Then
        mPhotoPath = CStr(picPath)
        On Error Resume Next
        imgDeliveryPhoto.Picture = LoadPicture(mPhotoPath)
        If Err.Number <> 0 Then Err.Clear
        On Error GoTo 0
    End If
End Sub

Private Sub btnDeliver_Click()
    Dim rawEntryID As String
    rawEntryID = GetEntryIDFromDisplay(cmbEntryID.value)
    
    If rawEntryID = "" Then
        MsgBox "Pehle Entry ID select karo!", vbExclamation, "Required"
        cmbEntryID.SetFocus: Exit Sub
    End If
    
    If mIsDelivered Then
        MsgBox "Yeh product already delivered hai!", vbExclamation, "Already Delivered"
        Exit Sub
    End If
    
    ' ===== GUARANTEED SAME CUSTOMER CHECK =====
    Dim isSameCust As Boolean
    isSameCust = False
    
    On Error Resume Next
    ' Try multiple methods to detect checked state
    If Me.chkSameCustomer.value = 1 Then isSameCust = True
    If Me.chkSameCustomer.value = True Then isSameCust = True  ' Backup check
    On Error GoTo 0
    
    ' ===== VALIDATION =====
    If Not isSameCust Then
        ' NORMAL: All fields mandatory
        If Trim(cmbReturnMode.value) = "" Then
            MsgBox "Return Mode select karo!", vbExclamation
            cmbReturnMode.SetFocus: Exit Sub
        End If
        
        If Trim(cmbReturnName.value) = "" Then
            MsgBox "Courier/Return Name daalo!", vbExclamation
            cmbReturnName.SetFocus: Exit Sub
        End If
        
        If Trim(txtReturnMobile.value) = "" Then
            MsgBox "Return Mobile Number daalo!", vbExclamation
            txtReturnMobile.SetFocus: Exit Sub
        End If
        
        If Trim(txtReturnAddress.value) = "" Then
            MsgBox "Return Address daalo!", vbExclamation
            txtReturnAddress.SetFocus: Exit Sub
        End If
        
        If Trim(txtReturnDate.value) = "" Then
            MsgBox "Return Date daalo!", vbExclamation
            txtReturnDate.SetFocus: Exit Sub
        End If
    Else
        ' SAME CUSTOMER: Auto-fill if empty
        On Error Resume Next
        If Trim(cmbReturnMode.value) = "" Then cmbReturnMode.value = "SELF"
        If Trim(cmbReturnName.value) = "" Then cmbReturnName.value = lblCustomerValue.caption
        If Trim(txtReturnMobile.value) = "" Then txtReturnMobile.value = lblMobileValue.caption
        If Trim(txtReturnAddress.value) = "" Then txtReturnAddress.value = lblAddressValue.caption
        On Error GoTo 0
    End If
    
   
    
    ' === OTP GENERATE & SEND ===
    Randomize
    mCurrentOTP = Format(Int(Rnd() * 900000) + 100000, "000000")
    mOTPTimestamp = Now
    
    Dim customerMobile As String
    customerMobile = lblMobileValue.caption
    
    If Trim(customerMobile) = "" Or Len(Trim(customerMobile)) <> 10 Then
        MsgBox "Customer mobile number valid nahi hai! OTP nahi bhej sakta.", vbExclamation
        Exit Sub
    End If
    
    ' Send WhatsApp OTP
    Call SendDeliveryOTP(customerMobile, mCurrentOTP, rawEntryID)
    
    ' Show InputBox for OTP
    Dim enteredOTP As String
    enteredOTP = InputBox("OTP sent to Customer WhatsApp" & vbCrLf & _
                          "Mobile: " & customerMobile & vbCrLf & vbCrLf & _
                          "Enter 6-digit OTP:", "Verify Delivery OTP")
    
    If Trim(enteredOTP) = "" Then
        MsgBox "OTP Required! Delivery cancelled.", vbExclamation
        mCurrentOTP = ""
        Exit Sub
    End If
    
    ' 10 minute Expiry Check
    Dim timeDiff As Double
    timeDiff = DateDiff("n", mOTPTimestamp, Now)
    If timeDiff > 10 Then
        MsgBox "OTP Expired! 10 minutes ho gaye." & vbCrLf & "Delivery cancelled.", vbCritical
        mCurrentOTP = ""
        Exit Sub
    End If
    
    ' Verify OTP
    If enteredOTP = mCurrentOTP Then
        MsgBox "OTP Verified Successfully!" & vbCrLf & "Delivery Confirm ho raha hai...", vbInformation
        mCurrentOTP = ""
        
        ' ?????? Actual Delivery Save
        Call ConfirmDelivery(rawEntryID)
    Else
        MsgBox "Invalid OTP! Delivery cancelled.", vbCritical
        mCurrentOTP = ""
        Exit Sub
    End If
End Sub
Private Sub btnClear_Click()
    ClearAllFields
    cmbEntryID.value = ""
    mEntryID = "": mAssignID = "": mCustomerID = ""
    mPhotoPath = ""
    mIsDelivered = False
    mIsDeliveryConfirmed = False   ' ?? Reset confirmation
    mCurrentOTP = ""
    
    btnDeliver.enabled = True
    btnDeliver.caption = "DELIVER"
    
    ' ?????? Bottom buttons disable ???
    Call SetBottomButtonsEnabled(False)
End Sub

Private Sub btnCancel_Click()
    If mIsDeliveryConfirmed Then
        If MsgBox("Delivery already confirmed hai. Form band karna hai?", vbQuestion + vbYesNo, "Confirm") = vbNo Then Exit Sub
    Else
        If MsgBox("Form band karna hai?", vbQuestion + vbYesNo, "Confirm") = vbNo Then Exit Sub
    End If
    Unload Me
End Sub

Private Sub ClearAllFields()
    Dim ctrl As Control
    
    lblEntryIDValue.caption = ""
    lblCustomerValue.caption = ""
    lblCompanyValue.caption = ""
    lblProductValue.caption = ""
    lblModelValue.caption = ""
    lblOriginalSerialValue.caption = ""
    lblFinalSerialValue.caption = ""
    lblWarrantyStatusValue.caption = ""
    lblMobileValue.caption = ""
    lblAddressValue.caption = ""
    txtWorkDone.value = ""
    
    ' Clear Payment ListView
    Me.lstPaymentDetails.ListItems.Clear
    
    ' ===== FIX: chkSameCustomer ko exclude karo =====
    ' Sirf ek loop — chkSameCustomer protected
    For Each ctrl In Me.Controls
        If TypeName(ctrl) = "CheckBox" Then
            If ctrl.name <> "chkSameCustomer" Then
                ctrl.caption = ""
                ctrl.value = False
                ctrl.visible = False
            End If
        End If
    Next ctrl
    
    ' ? PURANA LOOP DELETE KIYA GAYA
    
    cmbReturnMode.value = ""
    cmbReturnName.value = ""
    txtReturnMobile.value = ""
    txtReturnAddress.value = ""
    txtReturnDocket.value = ""
    txtReturnDate.value = Format(Date, "dd-mm-yyyy")
    txtRemarks.value = ""
    
    On Error Resume Next
    imgCustomer.Picture = Nothing
    imgProductAfter.Picture = Nothing
    imgDeliveryPhoto.Picture = Nothing
    On Error GoTo 0
    mPhotoPath = ""
    
    ' Clear Payment ListView
    Me.lstPaymentDetails.ListItems.Clear
End Sub
Private Function IsDelivered(entryID As String, wsDel As Worksheet) As Boolean
    Dim lastRow As Long, i As Long
    Dim wsJob As Worksheet
    
    ' === Step 1: Check Delivery_Master ===
    If wsDel Is Nothing Then
        IsDelivered = False
        Exit Function
    End If
    
    lastRow = wsDel.Cells(wsDel.Rows.count, 1).End(xlUp).row
    If lastRow < 2 Then
        IsDelivered = False
        Exit Function
    End If
    
    For i = 2 To lastRow
        If Trim(wsDel.Cells(i, 2).value & "") = Trim(entryID & "") Then
            ' === Step 2: Also check Job_Product Status ===
            On Error Resume Next
            Set wsJob = ThisWorkbook.Sheets("Job_Product")
            On Error GoTo 0
            
            If Not wsJob Is Nothing Then
                Dim jobLastRow As Long, j As Long
                jobLastRow = wsJob.Cells(wsJob.Rows.count, 1).End(xlUp).row
                For j = 2 To jobLastRow
                    If Trim(wsJob.Cells(j, 1).value & "") = Trim(entryID & "") Then
                        If UCase(Trim(wsJob.Cells(j, 4).value & "")) = "DELIVERED" Then
                            IsDelivered = True
                        Else
                            ' Delivery_Master me entry hai, but Job_Product me "DELIVERED" nahi
                            ' This means incomplete delivery - treat as NOT delivered
                            IsDelivered = False
                        End If
                        Exit Function
                    End If
                Next j
            End If
            
            IsDelivered = True
            Exit Function
        End If
    Next i
    
    IsDelivered = False
End Function

Private Function IsInCombo(cmb As ComboBox, val As String) As Boolean
    Dim i As Long
    For i = 0 To cmb.ListCount - 1
        If cmb.List(i) = val Then IsInCombo = True: Exit Function
    Next i
    IsInCombo = False
End Function








'=====================================================
' PAYMENT HELPER FUNCTIONS
'=====================================================
Public Function GetPaymentIDByEntry(entryID As String) As String
    Dim ws As Worksheet
    Dim r As Long
    
    GetPaymentIDByEntry = ""
    If Trim(entryID) = "" Then Exit Function
    
    Set ws = ThisWorkbook.Sheets("Payment_Master")
    
    For r = ws.Cells(ws.Rows.count, 1).End(xlUp).row To 2 Step -1
        If UCase(Trim(ws.Cells(r, 2).value)) = UCase(Trim(entryID)) Then
            GetPaymentIDByEntry = CStr(ws.Cells(r, 1).value)
            Exit Function
        End If
    Next r
End Function

Public Sub DeleteOldPayments(paymentID As String)
    Dim ws As Worksheet
    Dim r As Long
    
    If Trim(paymentID) = "" Then Exit Sub
    Set ws = ThisWorkbook.Sheets("Payment_Master")
    
    For r = ws.Cells(ws.Rows.count, 1).End(xlUp).row To 2 Step -1
        If UCase(Trim(ws.Cells(r, 1).value)) = UCase(Trim(paymentID)) Then
            ws.Rows(r).Delete
        End If
    Next r
End Sub

Public Function GeneratePaymentID() As String
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim lastID As String
    Dim num As Long
    
    Set ws = ThisWorkbook.Sheets("Payment_Master")
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    
    If lastRow < 2 Then
        GeneratePaymentID = "PAY00001"
    Else
        lastID = CStr(ws.Cells(lastRow, 1).value)
        If Left(lastID, 3) = "PAY" Then
            num = val(Mid(lastID, 4)) + 1
            GeneratePaymentID = "PAY" & Format(num, "00000")
        Else
            GeneratePaymentID = "PAY00001"
        End If
    End If
End Function

Public Sub SavePayment_Consolidated(paymentID As String, entryID As String, customerID As String, _
    chargeNames As String, totalReceipt As Double, payModes As String, totalPayment As Double, _
    expTypes As String, totalExpense As Double, remark As String)
    
    Dim ws As Worksheet
    Dim r As Long
    
    Set ws = ThisWorkbook.Sheets("Payment_Master")
    r = ws.Cells(ws.Rows.count, 1).End(xlUp).row + 1
    
    ws.Cells(r, 1).value = paymentID
    ws.Cells(r, 2).value = entryID
    ws.Cells(r, 3).value = customerID
    ws.Cells(r, 4).value = chargeNames
    ws.Cells(r, 5).value = totalPayment
    ws.Cells(r, 6).value = payModes
    ws.Cells(r, 7).value = totalReceipt
    ws.Cells(r, 8).value = expTypes
    ws.Cells(r, 9).value = totalExpense
    ws.Cells(r, 10).value = (totalPayment + totalExpense) - totalReceipt
    ws.Cells(r, 11).value = Now
    ws.Cells(r, 12).value = remark
End Sub




Private Sub SetBottomButtonsEnabled(enabled As Boolean)
    btnWhatsApp.enabled = enabled
    btnPrintDelivery.enabled = enabled
    btnExportDeliveryPDF.enabled = enabled
    btnSendEmail.enabled = enabled
End Sub
'========== SEND DELIVERY OTP — FULL FORMAT WITH ACCESSORIES ==========
Private Sub SendDeliveryOTP(mobile As String, otp As String, entryID As String)
    Dim url As String
    Dim message As String
    Dim wsh As Object
    Dim wsConfig As Worksheet
    
    Dim companyName As String
    Dim address1 As String, address2 As String
    Dim city As String, pin As String, state As String, country As String
    Dim companyMobile As String
    Dim fullAddress As String
    Dim accList As String
    Dim ctrl As Control
    
    ' === Read Software_Config ===
    On Error Resume Next
    Set wsConfig = ThisWorkbook.Sheets("Software_Config")
    If Not wsConfig Is Nothing Then
        companyName = Trim(wsConfig.Range("B2").value & "")
        address1 = Trim(wsConfig.Range("B3").value & "")
        address2 = Trim(wsConfig.Range("B4").value & "")
        city = Trim(wsConfig.Range("B5").value & "")
        pin = Trim(wsConfig.Range("B6").value & "")
        state = Trim(wsConfig.Range("B7").value & "")
        country = Trim(wsConfig.Range("B8").value & "")
        companyMobile = Trim(wsConfig.Range("B9").value & "")
    End If
    On Error GoTo 0
    
    ' === Build Full Address ===
    fullAddress = address1
    If address2 <> "" Then fullAddress = fullAddress & ", " & address2
    If city <> "" Then fullAddress = fullAddress & ", " & city
    If pin <> "" Then fullAddress = fullAddress & " - " & pin
    If state <> "" Then fullAddress = fullAddress & ", " & state
    If country <> "" Then fullAddress = fullAddress & ", " & country
    
    ' === FALLBACKS ===
    If Trim(companyName) = "" Then companyName = "GLOBAL IT SOLUTIONS"
    If Len(Trim(fullAddress)) < 15 Then fullAddress = "MAIN ROAD, IN FRONT OF AAHAR KENDRA, NAYAGARH - 752069, ODISHA, INDIA"
    If Trim(companyMobile) = "" Or Len(Trim(companyMobile)) < 10 Then companyMobile = "9777971045"
    
    ' === COLLECT ACCESSORIES ===
    accList = ""
    For Each ctrl In Me.Controls
        If TypeName(ctrl) = "CheckBox" Then
            If ctrl.visible And ctrl.value = True And Trim(ctrl.caption) <> "" Then
                If accList = "" Then
                    accList = Trim(ctrl.caption)
                Else
                    accList = accList & ", " & Trim(ctrl.caption)
                End If
            End If
        End If
    Next ctrl
    
    ' === BUILD MESSAGE ===
    message = "*GLOBAL IT SOLUTIONS*" & vbCrLf & _
              fullAddress & vbCrLf & _
              "WhatsApp & Call: " & companyMobile & vbCrLf & _
              "----------------------------------------" & vbCrLf & vbCrLf & _
              "*PRODUCT DELIVERY VERIFICATION*" & vbCrLf & vbCrLf & _
              "*Customer:*" & vbCrLf & _
              "ID: " & mCustomerID & " | Name: " & lblCustomerValue.caption & vbCrLf & _
              "Mobile: " & lblMobileValue.caption & vbCrLf
    If lblAddressValue.caption <> "" Then
        message = message & "Address: " & lblAddressValue.caption & vbCrLf
    End If
    message = message & vbCrLf & _
              "*Product:*" & vbCrLf & _
              "Entry: " & entryID & vbCrLf & _
              "Item: " & lblProductValue.caption & " | " & lblCompanyValue.caption & vbCrLf & _
              "Model: " & lblModelValue.caption & " | Serial: " & lblOriginalSerialValue.caption & vbCrLf & _
              "Warranty: " & lblWarrantyStatusValue.caption & vbCrLf & vbCrLf
    
    ' ?????? ACCESSORIES SECTION
    If accList <> "" Then
        message = message & "*Accessories Packed:*" & vbCrLf & accList & vbCrLf & vbCrLf
    End If
    
    ' === DELIVERY DETAILS ===
    If cmbReturnMode.value <> "" Then
        message = message & "*Delivery By:*" & vbCrLf & _
                  "Mode: " & cmbReturnMode.value & " | Name: " & cmbReturnName.value & vbCrLf & _
                  "Mobile: " & txtReturnMobile.value & vbCrLf & _
                  "Address: " & txtReturnAddress.value & vbCrLf
        If txtReturnDocket.value <> "" Then
            message = message & "Docket#: " & txtReturnDocket.value & vbCrLf
        End If
        message = message & "Date: " & txtReturnDate.value & vbCrLf
        If txtRemarks.value <> "" Then
            message = message & "Remark: " & txtRemarks.value & vbCrLf
        End If
        message = message & vbCrLf
    End If
    
    ' === OTP ===
    message = message & "----------------------------------------" & vbCrLf & _
              "*DELIVERY OTP: " & otp & "*" & vbCrLf & _
              "----------------------------------------" & vbCrLf & vbCrLf & _
              "Valid for *10 minutes* only." & vbCrLf & _
              "Do NOT share this OTP with anyone." & vbCrLf & vbCrLf & _
              "Enter this OTP to CONFIRM delivery."
    
    ' === SEND ===
    url = "whatsapp://send?phone=91" & Trim(mobile) & "&text=" & URLEncodeDelivery(message)
    
    Set wsh = CreateObject("WScript.Shell")
    wsh.Run url, 1, False
    Set wsh = Nothing
    
    MsgBox "Delivery OTP sent!" & vbCrLf & "Mobile: " & mobile, vbInformation, "OTP Sent"
End Sub
'========== URL ENCODE FOR DELIVERY OTP ==========
Private Function URLEncodeDelivery(text As String) As String
    Dim i As Integer, char As String, result As String, CharCode As Long
    
    If Len(Trim(text)) = 0 Then
        URLEncodeDelivery = ""
        Exit Function
    End If
    
    For i = 1 To Len(text)
        char = Mid(text, i, 1)
        CharCode = Asc(char)
        
        If (CharCode >= 65 And CharCode <= 90) Or _
           (CharCode >= 97 And CharCode <= 122) Or _
           (CharCode >= 48 And CharCode <= 57) Then
            result = result & char
        ElseIf char = " " Then
            result = result & "%20"
        ElseIf char = vbLf Then
            result = result & "%0A"
        ElseIf char = vbCr Then
            ' Ignore
        ElseIf CharCode < 16 Then
            result = result & "%0" & Hex(CharCode)
        ElseIf CharCode <= 255 Then
            result = result & "%" & Hex(CharCode)
        Else
            result = result & "%3F"
        End If
    Next i
    
    URLEncodeDelivery = result
End Function

'========== CONFIRM DELIVERY — OTP Verify ???? ?? ??? ?? call ???? ==========
Private Sub ConfirmDelivery(rawEntryID As String)
    Dim wsDel As Worksheet, wsJob As Worksheet
    Dim delRow As Long, i As Long, lastRow As Long
    Dim delID As String, accReturned As String
    Dim ctrl As Control
    
    Set wsDel = ThisWorkbook.Sheets("Delivery_Master")
    Set wsJob = ThisWorkbook.Sheets("Job_Product")
    
    delRow = wsDel.Cells(wsDel.Rows.count, 1).End(xlUp).row + 1
    delID = "DEL" & Format(delRow - 1, "00000")
    
    ' Accessories collect karo
    accReturned = ""
    For Each ctrl In Me.Controls
        If TypeName(ctrl) = "CheckBox" Then
            If ctrl.visible And ctrl.value = True Then
                If accReturned = "" Then accReturned = Trim(ctrl.caption) Else accReturned = accReturned & ", " & Trim(ctrl.caption)
            End If
        End If
    Next ctrl
    
    ' Save to Delivery_Master
    wsDel.Cells(delRow, 1).value = delID
    wsDel.Cells(delRow, 2).value = rawEntryID
    wsDel.Cells(delRow, 3).value = mCustomerID
    wsDel.Cells(delRow, 4).value = lblCustomerValue.caption
    wsDel.Cells(delRow, 5).value = lblProductValue.caption
    wsDel.Cells(delRow, 6).value = lblCompanyValue.caption
    wsDel.Cells(delRow, 7).value = lblModelValue.caption
    wsDel.Cells(delRow, 8).value = txtWorkDone.value
    wsDel.Cells(delRow, 9).value = cmbReturnMode.value
    wsDel.Cells(delRow, 10).value = cmbReturnName.value
    wsDel.Cells(delRow, 11).value = txtReturnMobile.value
    wsDel.Cells(delRow, 12).value = txtReturnAddress.value
    wsDel.Cells(delRow, 13).value = txtReturnDocket.value
    wsDel.Cells(delRow, 14).value = txtReturnDate.value
    wsDel.Cells(delRow, 15).value = mPhotoPath
    wsDel.Cells(delRow, 16).value = txtRemarks.value
    wsDel.Cells(delRow, 17).value = accReturned
    wsDel.Cells(delRow, 18).value = Format(Now, "dd-mm-yyyy hh:mm")
    wsDel.Cells(delRow, 19).value = "Delivered"
    
    ' Update Job_Product status
    lastRow = wsJob.Cells(wsJob.Rows.count, 1).End(xlUp).row
    For i = 2 To lastRow
        If Trim(wsJob.Cells(i, 1).value) = rawEntryID Then
            wsJob.Cells(i, 4).value = "DELIVERED"
            wsJob.Cells(i, 5).value = Format(Now, "dd-mm-yyyy hh:mm")
            Exit For
        End If
    Next i
    
    ' ?????? Mark as confirmed and enable buttons
    mIsDeliveryConfirmed = True
    mIsDelivered = True
    
     ' ===== SAME CUSTOMER CHECK =====
    Dim deliveredBy As String
    Dim isSameCustomer As Boolean
    
    On Error Resume Next
    isSameCustomer = (Me.chkSameCustomer.value = 1)
    On Error GoTo 0
    
    If isSameCustomer Then
        deliveredBy = lblCustomerValue.caption & " (SAME CUSTOMER - SELF)"
    Else
        deliveredBy = cmbReturnName.value
    End If
    
    ' Save to Delivery_Master — deliveredBy use karo
    wsDel.Cells(delRow, 10).value = deliveredBy  ' Delivered By Name
    
    ' ?????? Bottom buttons enable ??? — ?? Print/PDF/WhatsApp/Email ??? ?????
    Call SetBottomButtonsEnabled(True)
    
    ' Deliver button disable ??? — Already delivered
    btnDeliver.enabled = False
    btnDeliver.caption = "DELIVERED"
    
    MsgBox "PRODUCT DELIVERED SUCCESSFULLY!" & vbCrLf & _
           "Delivery ID: " & delID & vbCrLf & vbCrLf & _
           "Ab WhatsApp, Print, PDF, Email kar sakte ho.", vbInformation, "GLOBAL SOFT - DELIVERY CONFIRMED"
               ' ?????? AUTO WhatsApp Delivery Confirmation
    Call SendDeliveryConfirmationWhatsApp(rawEntryID, delID)
End Sub
'========== AUTO WHATSAPP — DELIVERY CONFIRMATION TO CUSTOMER ==========
Private Sub SendDeliveryConfirmationWhatsApp(entryID As String, delID As String)
    Dim mobile As String
    Dim message As String
    Dim wsh As Object
    Dim url As String
    Dim wsConfig As Worksheet
    
    Dim companyName As String, address1 As String, address2 As String
    Dim city As String, pin As String, state As String, country As String
    Dim companyMobile As String
    Dim fullAddress As String
    Dim accList As String
    Dim ctrl As Control
    
    ' ?????? ?? line ADD ???
    Dim i As Long
    
    ' === Customer Mobile ===
    mobile = Trim(lblMobileValue.caption)
    If mobile = "" Or Len(mobile) <> 10 Then
        ' Fallback — Customer_Master ??
        Dim wsCust As Worksheet
        Set wsCust = ThisWorkbook.Sheets("Customer_Master")
        For i = 2 To wsCust.Cells(wsCust.Rows.count, 1).End(xlUp).row
            If Trim(wsCust.Cells(i, 1).value) = mCustomerID Then
                mobile = Trim(wsCust.Cells(i, 2).value)
                Exit For
            End If
        Next i
    End If
    If mobile = "" Or Len(mobile) <> 10 Then Exit Sub
    
    ' === Software_Config ===
    On Error Resume Next
    Set wsConfig = ThisWorkbook.Sheets("Software_Config")
    If Not wsConfig Is Nothing Then
        companyName = Trim(wsConfig.Range("B2").value & "")
        address1 = Trim(wsConfig.Range("B3").value & "")
        address2 = Trim(wsConfig.Range("B4").value & "")
        city = Trim(wsConfig.Range("B5").value & "")
        pin = Trim(wsConfig.Range("B6").value & "")
        state = Trim(wsConfig.Range("B7").value & "")
        country = Trim(wsConfig.Range("B8").value & "")
        companyMobile = Trim(wsConfig.Range("B9").value & "")
    End If
    On Error GoTo 0
    
    ' Build Address
    fullAddress = address1
    If address2 <> "" Then fullAddress = fullAddress & ", " & address2
    If city <> "" Then fullAddress = fullAddress & ", " & city
    If pin <> "" Then fullAddress = fullAddress & " - " & pin
    If state <> "" Then fullAddress = fullAddress & ", " & state
    If country <> "" Then fullAddress = fullAddress & ", " & country
    
    If Trim(companyName) = "" Then companyName = "GLOBAL IT SOLUTIONS"
    If Len(Trim(fullAddress)) < 15 Then fullAddress = "MAIN ROAD, IN FRONT OF AAHAR KENDRA, NAYAGARH - 752069, ODISHA, INDIA"
    If Trim(companyMobile) = "" Or Len(Trim(companyMobile)) < 10 Then companyMobile = "9777971045"
    
    ' === Collect Accessories ===
    accList = ""
    For Each ctrl In Me.Controls
        If TypeName(ctrl) = "CheckBox" Then
            If ctrl.visible And ctrl.value = True And Trim(ctrl.caption) <> "" Then
                If accList = "" Then accList = Trim(ctrl.caption) Else accList = accList & ", " & Trim(ctrl.caption)
            End If
        End If
    Next ctrl
    
    ' === BUILD SUCCESS MESSAGE ===
    message = "*GLOBAL IT SOLUTIONS*" & vbCrLf & _
              fullAddress & vbCrLf & _
              "WhatsApp & Call: " & companyMobile & vbCrLf & _
              "----------------------------------------" & vbCrLf & vbCrLf & _
              "*PRODUCT DELIVERED SUCCESSFULLY!*" & vbCrLf & vbCrLf & _
              "Dear " & lblCustomerValue.caption & "," & vbCrLf & vbCrLf & _
              "Your product has been delivered." & vbCrLf & vbCrLf & _
              "*Product:*" & vbCrLf & _
              "Entry ID: " & entryID & vbCrLf & _
              "Item: " & lblProductValue.caption & " | " & lblCompanyValue.caption & vbCrLf & _
              "Model: " & lblModelValue.caption & vbCrLf & _
              "Serial: " & lblOriginalSerialValue.caption & vbCrLf & vbCrLf
    
    If accList <> "" Then
        message = message & "*Accessories:* " & accList & vbCrLf & vbCrLf
    End If
    
    message = message & "*Delivery Details:*" & vbCrLf & _
              "Delivery ID: " & delID & vbCrLf & _
              "Mode: " & cmbReturnMode.value & vbCrLf & _
              "Delivered By: " & cmbReturnName.value & vbCrLf & _
              "Date: " & txtReturnDate.value & vbCrLf & vbCrLf & _
              "----------------------------------------" & vbCrLf & vbCrLf & _
              "Thank you for choosing GLOBAL IT SOLUTIONS!" & vbCrLf & _
              "For any query, call: " & companyMobile
    
    ' === SEND ===
    url = "whatsapp://send?phone=91" & Trim(mobile) & "&text=" & URLEncodeDelivery(message)
    
    Set wsh = CreateObject("WScript.Shell")
    wsh.Run url, 1, False
    Set wsh = Nothing
End Sub

Private Sub btnWhatsApp_Click()
    If Not mIsDeliveryConfirmed Then
        MsgBox "Pehle Delivery Confirm karo! OTP verification required.", vbExclamation
        Exit Sub
    End If
    
    ' Find Delivery ID for this entry
    Dim wsDel As Worksheet
    Dim lastRow As Long, i As Long
    Dim delID As String
    Set wsDel = ThisWorkbook.Sheets("Delivery_Master")
    lastRow = wsDel.Cells(wsDel.Rows.count, 1).End(xlUp).row
    For i = 2 To lastRow
        If Trim(wsDel.Cells(i, 2).value) = mEntryID Then
            delID = wsDel.Cells(i, 1).value
            Exit For
        End If
    Next i
    
    If delID = "" Then delID = "DEL00000"
    
    ' Resend confirmation
    Call SendDeliveryConfirmationWhatsApp(mEntryID, delID)
    
    MsgBox "Delivery confirmation WhatsApp resent!", vbInformation
End Sub

Private Sub btnPrintDelivery_Click()
    On Error GoTo ErrorHandler
    
    ' ===== ENTRY ID CHECK =====
    If Trim(m_DeliveryEntryID) = "" Then
        MsgBox "Entry ID not found! Pehle entry select karo.", vbExclamation
        Exit Sub
    End If
    
    ' ===== OTP CHECK =====
    If Not mIsDeliveryConfirmed Then
        MsgBox "Pehle Delivery Confirm karo! OTP verification required.", vbExclamation
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    Dim wsDelivery As Worksheet
    Set wsDelivery = GenerateDeliverySheet(m_DeliveryEntryID)
    
    If wsDelivery Is Nothing Then
        MsgBox "Delivery receipt banane me error!", vbCritical
        GoTo Cleanup
    End If
    
    ' ===== PRINT =====
    wsDelivery.PrintOut Copies:=1, Preview:=False
    
Cleanup:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Exit Sub
    
ErrorHandler:
    MsgBox "Print Error: " & Err.Description, vbCritical
    Resume Cleanup
End Sub

Private Function CleanFileName(ByVal strName As String) As String
    Dim invalidChars As String, i As Integer
    invalidChars = "\/:*?""<>|"
    For i = 1 To Len(invalidChars)
        strName = Replace(strName, Mid(invalidChars, i, 1), "_")
    Next i
    CleanFileName = Trim(strName)
End Function
Private Function GenerateDeliverySheet(Optional ByVal pEntryID As String = "") As Worksheet
    On Error GoTo ErrorHandler
    
    Dim wsCust As Worksheet, wsProd As Worksheet, wsAcc As Worksheet
    Dim wsPay As Worksheet, wsJob As Worksheet, wsConfig As Worksheet, wsSet As Worksheet
    Dim ws As Worksheet
    Dim r As Long, i As Long, accRow As Long, lastRowAcc As Long
    Dim compName As String, addr1 As String, addr2 As String, compMob As String, compEmail As String
    Dim logoPath As String, finalLogoPath As String, pageSize As String
    Dim eID As String, custID As String, custName As String, custMob As String
    Dim custAddress As String, custEmail As String, custGST As String
    Dim prodName As String, prodComp As String, prodModel As String, prodSerial As String, prodProb As String
    Dim entryDate As String, verifyType As String, verifyName As String, jobStatus As String
    Dim shpLogo As Shape
    Dim termsFound As Boolean, termRow As Long, t As Integer
    Dim terms(1 To 8) As String
    Dim totalCharges As Double, totalExp As Double, totalRec As Double, dueAmt As Double
    Dim payRow As Long, lastRowPay As Long, foundPayment As Boolean, payID As String
    
    ' ===== PHOTO SETTINGS VARIABLES =====
    Dim showCustPic As String, showProdPic As String, showAccPic As String
    Dim custPicPath As String, prodPicPath As String, accPicPath As String
    Dim custPhotoName As String, custPhotoFull As String
    Dim prodPhotoFull As String
    Dim accPhotoName As String, accPhotoFull As String
    Dim photoLeft As Single, photoTop As Single
    Dim picShape As Shape
    
    ' ===== DELIVERY FORM DATA =====
    Dim deliveryDate As String, deliveredBy As String, otpStatus As String
    Dim prodCondition As String, deliveryRemarks As String
    
    On Error Resume Next
    deliveryDate = Trim(Me.txtReturnDate.value)
    deliveredBy = Trim(Me.cmbReturnName.value)
    otpStatus = "VERIFIED"
    prodCondition = GetProductCondition(eID)
    deliveryRemarks = Trim(Me.txtRemarks.value)
    On Error GoTo ErrorHandler
    
    If deliveryDate = "" Then deliveryDate = Format(Date, "dd-MMM-yyyy")
    If deliveredBy = "" Then deliveredBy = "Service Engineer"
    
    ' ===== ENTRY ID SET =====
    If Trim(pEntryID) <> "" Then
        eID = Trim(pEntryID)
    ElseIf Trim(m_DeliveryEntryID) <> "" Then
        eID = Trim(m_DeliveryEntryID)
    Else
        eID = ""
    End If
    
    If eID = "" Then
        MsgBox "Entry ID missing for Delivery Receipt!", vbExclamation
        Set GenerateDeliverySheet = Nothing
        Exit Function
    End If
    
    ' ===== SHEETS =====
    Set wsCust = ThisWorkbook.Sheets("Customer_Master")
    Set wsProd = ThisWorkbook.Sheets("Job_Product")
    Set wsAcc = ThisWorkbook.Sheets("Job_Accessory")
    Set wsPay = ThisWorkbook.Sheets("Payment_Master")
    Set wsJob = ThisWorkbook.Sheets("Job_Master")
    Set wsConfig = ThisWorkbook.Sheets("Software_Config")
    
    ' ===== READ SETTINGS =====
    On Error Resume Next
    Set wsSet = ThisWorkbook.Sheets("Settings")
    If Not wsSet Is Nothing Then
        pageSize = UCase(Trim(wsSet.Range("B20").value & ""))
        showCustPic = UCase(Trim(wsSet.Range("B21").value & ""))
        showProdPic = UCase(Trim(wsSet.Range("B22").value & ""))
        showAccPic = UCase(Trim(wsSet.Range("B23").value & ""))
        custPicPath = Trim(wsSet.Range("B3").value & "")
        prodPicPath = Trim(wsSet.Range("B2").value & "")
        accPicPath = Trim(wsSet.Range("B6").value & "")
        logoPath = wsSet.Range("B11").value & ""
    End If
    On Error GoTo ErrorHandler
    
    If pageSize = "" Then pageSize = "A4"
    If custPicPath = "" Then custPicPath = ThisWorkbook.path & "\GLOBAL_SOFT_DATA\Images\Customers\"
    If prodPicPath = "" Then prodPicPath = ThisWorkbook.path & "\GLOBAL_SOFT_DATA\Images\Products\"
    If accPicPath = "" Then accPicPath = ThisWorkbook.path & "\GLOBAL_SOFT_DATA\Images\Accessories\"
    If logoPath = "" Then logoPath = ThisWorkbook.path & "\GLOBAL_SOFT_DATA\Images\Logo\"
    
    If Right(custPicPath, 1) <> "\" Then custPicPath = custPicPath & "\"
    If Right(prodPicPath, 1) <> "\" Then prodPicPath = prodPicPath & "\"
    If Right(accPicPath, 1) <> "\" Then accPicPath = accPicPath & "\"
    If Right(logoPath, 1) <> "\" Then logoPath = logoPath & "\"
    
    ' ===== COMPANY INFO =====
    compName = Trim(wsConfig.Range("B2").value & "")
    addr1 = Trim(wsConfig.Range("B3").value & "")
    addr2 = Trim(wsConfig.Range("B4").value & "")
    compMob = Trim(wsConfig.Range("B9").value & "")
    compEmail = Trim(wsConfig.Range("B10").value & "")
    If compName = "" Then compName = "GLOBAL IT SOLUTIONS"
    
    ' ===== FIND PRODUCT =====
    Dim prodFound As Boolean
    prodFound = False
    Dim lastRowProd As Long
    lastRowProd = wsProd.Cells(wsProd.Rows.count, 1).End(xlUp).row
    
    For i = 2 To lastRowProd
        If Trim(UCase(wsProd.Cells(i, 1).value & "")) = Trim(UCase(eID)) Then
            prodFound = True
            custID = Trim(wsProd.Cells(i, 2).value & "")
            prodName = wsProd.Cells(i, 6).value & ""
            prodComp = wsProd.Cells(i, 7).value & ""
            prodModel = wsProd.Cells(i, 8).value & ""
            prodSerial = wsProd.Cells(i, 9).value & ""
            prodProb = wsProd.Cells(i, 13).value & ""
            Exit For
        End If
    Next i
    
    If Not prodFound Then
        MsgBox "Product details not found for: " & eID, vbExclamation
        Set GenerateDeliverySheet = Nothing
        Exit Function
    End If
    
    ' ===== FIND CUSTOMER =====
    custName = "": custMob = "": custAddress = "": custEmail = "": custGST = ""
    custPhotoName = ""
    Dim lastRowCust As Long
    lastRowCust = wsCust.Cells(wsCust.Rows.count, 1).End(xlUp).row
        For i = 2 To lastRowCust
        If UCase(Trim(wsCust.Cells(i, 1).value & "")) = UCase(custID) Then
            custName = wsCust.Cells(i, 2).value & ""   ' ? Column B = Name
            custMob = wsCust.Cells(i, 3).value & ""      ' ? Column C = Mobile
            custAddress = wsCust.Cells(i, 4).value & ""
            custEmail = wsCust.Cells(i, 5).value & ""
            custGST = wsCust.Cells(i, 6).value & ""
            custPhotoName = Trim(wsCust.Cells(i, 7).value & "")
            Exit For
        End If
    Next i
    
    ' ===== FIND JOB MASTER =====
    entryDate = "": verifyType = "": verifyName = "": jobStatus = ""
    Dim lastRowJob As Long
    lastRowJob = wsJob.Cells(wsJob.Rows.count, 1).End(xlUp).row
    For i = 2 To lastRowJob
        If Trim(UCase(wsJob.Cells(i, 1).value & "")) = Trim(UCase(eID)) Then
            entryDate = wsJob.Cells(i, 8).value & ""
            verifyType = wsJob.Cells(i, 6).value & ""
            verifyName = wsJob.Cells(i, 7).value & ""
            jobStatus = wsJob.Cells(i, 10).value & ""
            Exit For
        End If
    Next i
    If entryDate = "" Then entryDate = Format(Date, "dd-MMM-yyyy")
    
    ' ===== LOGO PATH =====
    finalLogoPath = logoPath & "LOGO.png"
    If Dir(finalLogoPath) = "" Then finalLogoPath = logoPath & "LOGO.jpg"
    If Dir(finalLogoPath) = "" Then finalLogoPath = logoPath & "logo.png"
    If Dir(finalLogoPath) = "" Then finalLogoPath = logoPath & "logo.jpg"
    
    ' ===== CREATE TEMP SHEET =====
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Sheets("Temp_Delivery").Delete
    On Error GoTo 0
    
    Set ws = ThisWorkbook.Sheets.Add
    ws.name = "Temp_Delivery"
    
    ' ===== PAGE SETUP =====
    With ws.PageSetup
        If pageSize = "A5" Then
            .PaperSize = xlPaperA5
            .LeftMargin = Application.InchesToPoints(0.2)
            .RightMargin = Application.InchesToPoints(0.2)
            .TopMargin = Application.InchesToPoints(0.2)
            .BottomMargin = Application.InchesToPoints(0.2)
        Else
            .PaperSize = xlPaperA4
            .LeftMargin = Application.InchesToPoints(0.35)
            .RightMargin = Application.InchesToPoints(0.35)
            .TopMargin = Application.InchesToPoints(0.35)
            .BottomMargin = Application.InchesToPoints(0.35)
        End If
        .Orientation = xlPortrait
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
    End With
    
    ' ===== COLUMNS =====
    If pageSize = "A5" Then
        ws.Columns("A").ColumnWidth = 1.5
        ws.Columns("B").ColumnWidth = 7
        ws.Columns("C").ColumnWidth = 16
        ws.Columns("D").ColumnWidth = 13
        ws.Columns("E").ColumnWidth = 13
        ws.Columns("F").ColumnWidth = 13
        ws.Columns("G").ColumnWidth = 13
        ws.Columns("H").ColumnWidth = 1.5
    Else
        ws.Columns("A").ColumnWidth = 2
        ws.Columns("B").ColumnWidth = 9
        ws.Columns("C").ColumnWidth = 20
        ws.Columns("D").ColumnWidth = 16
        ws.Columns("E").ColumnWidth = 16
        ws.Columns("F").ColumnWidth = 16
        ws.Columns("G").ColumnWidth = 16
        ws.Columns("H").ColumnWidth = 2
    End If
    
    ' ===== DEFAULT TERMS =====
    terms(1) = "1. Product delivered in tested/working condition."
    terms(2) = "2. Customer has verified & accepted the product physically."
    terms(3) = "3. Warranty void if seal broken by unauthorized person."
    terms(4) = "4. Company not responsible for data loss after delivery."
    terms(5) = "5. All dues must be cleared before delivery."
    terms(6) = "6. Original receipt mandatory for future warranty claims."
    terms(7) = "7. No return accepted once product is delivered & signed."
    terms(8) = "8. For any issue, contact within 7 days of delivery."
    
    ' ===== BUILD RECEIPT =====
    r = 1
    
    ' --- ROW 1: COMPANY NAME ---
    With ws.Range("B" & r & ":G" & r)
        .Merge: .value = compName
        .Font.Size = IIf(pageSize = "A5", 20, 26): .Font.Bold = True: .Font.name = "Arial"
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
    End With
    ws.Rows(r).RowHeight = IIf(pageSize = "A5", 28, 36)
    
    ' LOGO
    If Dir(finalLogoPath) <> "" Then
        On Error Resume Next
        Set shpLogo = ws.Shapes.AddPicture(finalLogoPath, msoFalse, msoTrue, 0, 0, -1, -1)
        If Not shpLogo Is Nothing Then
            With shpLogo
                .LockAspectRatio = msoTrue
                .Height = IIf(pageSize = "A5", 80, 100)
                .Top = ws.Range("B" & r).Top + 2
                .Left = ws.Range("A" & r).Left + 2
                If .Width > IIf(pageSize = "A5", 70, 100) Then .Width = IIf(pageSize = "A5", 70, 100)
            End With
        End If
        On Error GoTo ErrorHandler
    End If
    r = r + 1
    
    ' --- ROW 2: ADDRESS ---
    With ws.Range("B" & r & ":G" & r)
        .Merge: .value = addr1
        .Font.Size = IIf(pageSize = "A5", 10, 12): .Font.name = "Arial"
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
    End With
    ws.Rows(r).RowHeight = IIf(pageSize = "A5", 16, 20)
    r = r + 1
    
    ' --- ROW 3: ADDRESS 2 ---
    If addr2 <> "" Then
        With ws.Range("B" & r & ":G" & r)
            .Merge: .value = addr2
            .Font.Size = IIf(pageSize = "A5", 10, 12)
            .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        End With
        ws.Rows(r).RowHeight = IIf(pageSize = "A5", 16, 20)
        r = r + 1
    End If
    
    ' --- ROW 4: MOBILE & EMAIL ---
    With ws.Range("B" & r & ":G" & r)
        .Merge
        .value = "Mobile: " & compMob & "  |  Email: " & compEmail
        .Font.Size = IIf(pageSize = "A5", 10, 12): .Font.Bold = True
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
    End With
    ws.Rows(r).RowHeight = IIf(pageSize = "A5", 18, 24)
    r = r + 1
    
    ' --- ROW 5: DELIVERY REPORT HEADER (GREEN) ---
    With ws.Range("B" & r & ":G" & r)
        .Merge
        .value = "PRODUCT DELIVERY REPORT"
        .Font.Bold = True: .Font.Size = IIf(pageSize = "A5", 12, 15)
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(0, 128, 0)
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
    End With
    ws.Rows(r).RowHeight = IIf(pageSize = "A5", 24, 30)
    r = r + 1
    
        ' --- ROW 6: CUSTOMER DETAILS HEADER (NAVY BLUE) ---
    ' LEFT SIDE — Customer Details
    With ws.Range("B" & r & ":D" & r)
        .Merge
        .value = "CUSTOMER DETAILS (" & custID & ")"
        .Font.Bold = True: .Font.Size = IIf(pageSize = "A5", 10, 12)
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(0, 0, 128)
        .HorizontalAlignment = xlLeft: .VerticalAlignment = xlCenter
    End With
    
    ' RIGHT SIDE — Delivery Date
    With ws.Range("E" & r & ":G" & r)
        .Merge
        .value = "DELIVERY DATE: " & deliveryDate
        .Font.Bold = True: .Font.Size = IIf(pageSize = "A5", 10, 12)
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(0, 0, 128)
        .HorizontalAlignment = xlRight: .VerticalAlignment = xlCenter
    End With
    
    ws.Rows(r).RowHeight = IIf(pageSize = "A5", 22, 26)
    r = r + 1
    
    ' --- CUSTOMER DETAILS SECTION ---
    Dim custDetailStart As Long
    custDetailStart = r
    
    ' ROW 7: NAME
    ws.Range("B" & r).value = "Name:"
    ws.Range("B" & r).Font.Bold = True: ws.Range("B" & r).Font.Size = IIf(pageSize = "A5", 9, 11)
    ws.Range("C" & r & ":D" & r).Merge
    ws.Range("C" & r).value = custName
    ws.Range("C" & r).Font.Size = IIf(pageSize = "A5", 9, 11)
    r = r + 1
    
    ' ROW 8: MOBILE
    ws.Range("B" & r).value = "Mobile:"
    ws.Range("B" & r).Font.Bold = True: ws.Range("B" & r).Font.Size = IIf(pageSize = "A5", 9, 11)
    ws.Range("C" & r).value = custMob
    ws.Range("C" & r).Font.Size = IIf(pageSize = "A5", 9, 11)
    ws.Range("C" & r).HorizontalAlignment = xlLeft
    r = r + 1
    
    ' ROW 9: EMAIL (compact)
    If custEmail <> "" Then
        ws.Range("B" & r).value = "Email:"
        ws.Range("B" & r).Font.Bold = True: ws.Range("B" & r).Font.Size = IIf(pageSize = "A5", 9, 11)
        ws.Range("C" & r & ":D" & r).Merge
        ws.Range("C" & r).value = custEmail
        ws.Range("C" & r).Font.Size = IIf(pageSize = "A5", 9, 11)
        r = r + 1
    End If
    
    ' ROW 10: ADDRESS
    If custAddress <> "" Then
        ws.Range("B" & r).value = "Address:"
        ws.Range("B" & r).Font.Bold = True: ws.Range("B" & r).Font.Size = IIf(pageSize = "A5", 9, 11)
        ws.Range("C" & r & ":E" & r).Merge
        ws.Range("C" & r).value = Replace(Replace(custAddress, vbCrLf, ", "), vbLf, ", ")
        ws.Range("C" & r).Font.Size = IIf(pageSize = "A5", 9, 11)
        ws.Range("C" & r).HorizontalAlignment = xlLeft
        ws.Rows(r).RowHeight = IIf(pageSize = "A5", 16, 20)
        r = r + 1
    End If
    
    ' ROW 11: GST
    If custGST <> "" Then
        ws.Range("B" & r).value = "GST:"
        ws.Range("B" & r).Font.Bold = True: ws.Range("B" & r).Font.Size = IIf(pageSize = "A5", 9, 11)
        ws.Range("C" & r).value = custGST
        ws.Range("C" & r).Font.Size = IIf(pageSize = "A5", 9, 11)
        r = r + 1
    End If
    
    ' ===== CUSTOMER PHOTO — RIGHT SIDE, BIGGER =====
    If showCustPic = "YES" And custPhotoName <> "" Then
        custPhotoFull = custPicPath & custPhotoName
        If Dir(custPhotoFull) = "" Then
            custPhotoFull = ThisWorkbook.path & "\GLOBAL_SOFT_DATA\Images\Customers\" & custPhotoName
        End If
        If Dir(custPhotoFull) <> "" Then
            On Error Resume Next
            photoTop = ws.Range("B" & custDetailStart).Top + 2
            ' RIGHT SIDE — Column F start
            photoLeft = ws.Range("F" & custDetailStart).Left + 5
            Set picShape = ws.Shapes.AddPicture(custPhotoFull, msoFalse, msoTrue, photoLeft, photoTop, -1, -1)
            If Not picShape Is Nothing Then
                picShape.LockAspectRatio = msoTrue
                If pageSize = "A5" Then
                    If picShape.Width > 100 Then picShape.Width = 100
                    If picShape.Height > 80 Then picShape.Height = 80
                Else
                    If picShape.Width > 140 Then picShape.Width = 140
                    If picShape.Height > 120 Then picShape.Height = 120
                End If
            End If
            On Error GoTo ErrorHandler
        End If
    End If
    
    r = r + 1 ' GAP
    
    ' --- PRODUCT DETAILS HEADER (RED/ORANGE) ---
    Dim prodHeaderRow As Long
    prodHeaderRow = r
    With ws.Range("B" & r & ":G" & r)
        .Merge
        .value = "PRODUCT DETAILS (" & eID & ")"
        .Font.Bold = True: .Font.Size = IIf(pageSize = "A5", 10, 12)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlLeft: .VerticalAlignment = xlCenter
        If UCase(Left(eID, 3)) = "SER" Then
            .Interior.Color = RGB(255, 165, 0)
        Else
            .Interior.Color = RGB(200, 0, 0)
        End If
    End With
    ws.Rows(r).RowHeight = IIf(pageSize = "A5", 22, 26)
    r = r + 1
    
    ' Product Info — compact 2 columns
    ws.Range("B" & r).value = "Product:"
    ws.Range("B" & r).Font.Bold = True: ws.Range("B" & r).Font.Size = IIf(pageSize = "A5", 9, 11)
    ws.Range("C" & r & ":D" & r).Merge
    ws.Range("C" & r).value = prodName: ws.Range("C" & r).Font.Size = IIf(pageSize = "A5", 9, 11)
    ws.Range("E" & r).value = "Company:"
    ws.Range("E" & r).Font.Bold = True: ws.Range("E" & r).Font.Size = IIf(pageSize = "A5", 9, 11)
    ws.Range("F" & r & ":G" & r).Merge
    ws.Range("F" & r).value = prodComp: ws.Range("F" & r).Font.Size = IIf(pageSize = "A5", 9, 11)
    r = r + 1
    
    ws.Range("B" & r).value = "Model:"
    ws.Range("B" & r).Font.Bold = True: ws.Range("B" & r).Font.Size = IIf(pageSize = "A5", 9, 11)
    ws.Range("C" & r).value = prodModel: ws.Range("C" & r).Font.Size = IIf(pageSize = "A5", 9, 11)
    ws.Range("D" & r).value = "Serial:"
    ws.Range("D" & r).Font.Bold = True: ws.Range("D" & r).Font.Size = IIf(pageSize = "A5", 9, 11)
    ws.Range("E" & r & ":F" & r).Merge
    ws.Range("E" & r).value = prodSerial: ws.Range("E" & r).Font.Size = IIf(pageSize = "A5", 9, 11)
    r = r + 1
    
    ' Problem — compact single line
    If prodProb <> "" Then
        With ws.Range("B" & r & ":G" & r)
            .Merge
            .value = "Problem: " & prodProb
            .Font.Color = RGB(200, 0, 0): .Font.Size = IIf(pageSize = "A5", 9, 10)
        End With
        r = r + 1
    End If
    
    ' Entry Date & Status
    ws.Range("B" & r).value = "Entry Date:"
    ws.Range("B" & r).Font.Bold = True: ws.Range("B" & r).Font.Size = IIf(pageSize = "A5", 9, 10)
    ws.Range("C" & r).value = entryDate: ws.Range("C" & r).Font.Size = IIf(pageSize = "A5", 9, 10)
    ws.Range("E" & r).value = "Job Status:"
    ws.Range("E" & r).Font.Bold = True: ws.Range("E" & r).Font.Size = IIf(pageSize = "A5", 9, 10)
    ws.Range("F" & r & ":G" & r).Merge
    ws.Range("F" & r).value = jobStatus: ws.Range("F" & r).Font.Size = IIf(pageSize = "A5", 9, 10)
    r = r + 1
    
    ' ===== PRODUCT PHOTO — RIGHT SIDE, BIGGER =====
    If showProdPic = "YES" Then
        prodPhotoFull = prodPicPath & eID & "\Photo1.jpg"
        If Dir(prodPhotoFull) = "" Then
            prodPhotoFull = ThisWorkbook.path & "\GLOBAL_SOFT_DATA\Images\Products\" & eID & "\Photo1.jpg"
        End If
        If Dir(prodPhotoFull) <> "" Then
            On Error Resume Next
            photoTop = ws.Range("B" & prodHeaderRow).Top + 35
            ' RIGHT SIDE — Column F start
            photoLeft = ws.Range("F" & prodHeaderRow).Left + 35
            Set picShape = ws.Shapes.AddPicture(prodPhotoFull, msoFalse, msoTrue, photoLeft, photoTop, -1, -1)
            If Not picShape Is Nothing Then
                picShape.LockAspectRatio = msoTrue
                If pageSize = "A5" Then
                    If picShape.Width > 80 Then picShape.Width = 80
                    If picShape.Height > 80 Then picShape.Height = 80
                Else
                    If picShape.Width > 140 Then picShape.Width = 140
                    If picShape.Height > 120 Then picShape.Height = 120
                End If
            End If
            On Error GoTo ErrorHandler
        End If
    End If
    
    r = r + 1 ' GAP
    
    ' ===== ACCESSORIES RETURNED =====
    lastRowAcc = wsAcc.Cells(wsAcc.Rows.count, 1).End(xlUp).row
    Dim accCount As Long
    accCount = 0
    For accRow = 2 To lastRowAcc
        If Trim(UCase(wsAcc.Cells(accRow, 1).value & "")) = Trim(UCase(eID)) Then accCount = accCount + 1
    Next accRow
    
    Dim accPhotoList As String
    accPhotoList = ""
    
    If accCount > 0 Then
        With ws.Range("B" & r & ":G" & r)
            .Merge
            .value = "ACCESSORIES RETURNED"
            .Font.Color = RGB(0, 0, 150): .Font.Bold = True: .Font.Size = IIf(pageSize = "A5", 10, 11)
        End With
        r = r + 1
        
        ' Compact Header
        ws.Range("B" & r).value = "#": ws.Range("B" & r).Font.Bold = True: ws.Range("B" & r).Font.Size = IIf(pageSize = "A5", 8, 9)
        ws.Range("B" & r).HorizontalAlignment = xlCenter
        ws.Range("C" & r).value = "Accessory": ws.Range("C" & r).Font.Bold = True: ws.Range("C" & r).Font.Size = IIf(pageSize = "A5", 8, 9)
        ws.Range("D" & r).value = "Brand": ws.Range("D" & r).Font.Bold = True: ws.Range("D" & r).Font.Size = IIf(pageSize = "A5", 8, 9)
        ws.Range("E" & r).value = "Serial": ws.Range("E" & r).Font.Bold = True: ws.Range("E" & r).Font.Size = IIf(pageSize = "A5", 8, 9)
        ws.Range("F" & r & ":G" & r).Merge
        ws.Range("F" & r).value = "Status": ws.Range("F" & r).Font.Bold = True: ws.Range("F" & r).Font.Size = IIf(pageSize = "A5", 8, 9)
        r = r + 1
        
        Dim accIndex As Integer
        accIndex = 1
        For accRow = 2 To lastRowAcc
            If Trim(UCase(wsAcc.Cells(accRow, 1).value & "")) = Trim(UCase(eID)) Then
                ws.Rows(r).RowHeight = IIf(pageSize = "A5", 18, 22)
                ws.Range("B" & r).value = accIndex
                ws.Range("B" & r).Font.Size = IIf(pageSize = "A5", 8, 9): ws.Range("B" & r).HorizontalAlignment = xlCenter
                ws.Range("C" & r).value = wsAcc.Cells(accRow, 4).value: ws.Range("C" & r).Font.Size = IIf(pageSize = "A5", 8, 9)
                ws.Range("D" & r).value = wsAcc.Cells(accRow, 5).value: ws.Range("D" & r).Font.Size = IIf(pageSize = "A5", 8, 9)
                ws.Range("E" & r).value = wsAcc.Cells(accRow, 6).value: ws.Range("E" & r).Font.Size = IIf(pageSize = "A5", 8, 9)
                ws.Range("F" & r & ":G" & r).Merge
                ws.Range("F" & r).value = wsAcc.Cells(accRow, 7).value: ws.Range("F" & r).Font.Size = IIf(pageSize = "A5", 8, 9)
                
                Dim accPhotoNameCell As String
                accPhotoNameCell = Trim(wsAcc.Cells(accRow, 8).value & "")
                If accPhotoNameCell <> "" Then
                    If accPhotoList = "" Then
                        accPhotoList = accPhotoNameCell
                    Else
                        accPhotoList = accPhotoList & "|" & accPhotoNameCell
                    End If
                End If
                
                r = r + 1
                accIndex = accIndex + 1
                If accIndex > 6 Then Exit For
            End If
        Next accRow
        
        ' ===== ACCESSORY PHOTOS INSERT =====
        If showAccPic = "YES" And accPhotoList <> "" Then
            Dim accPhotos() As String
            accPhotos = Split(accPhotoList, "|")
            Dim accPicIndex As Integer
            Dim accLeft As Single, accTop As Single
            accLeft = ws.Range("B" & r).Left + 5
            accTop = ws.Range("B" & r).Top + 3
            
            For accPicIndex = 0 To UBound(accPhotos)
                If accPicIndex >= 4 Then Exit For
                accPhotoFull = accPicPath & accPhotos(accPicIndex)
                If Dir(accPhotoFull) = "" Then
                    accPhotoFull = ThisWorkbook.path & "\GLOBAL_SOFT_DATA\Images\Accessories\" & accPhotos(accPicIndex)
                End If
                If Dir(accPhotoFull) <> "" Then
                    On Error Resume Next
                    Set picShape = ws.Shapes.AddPicture(accPhotoFull, msoFalse, msoTrue, accLeft, accTop, -1, -1)
                    If Not picShape Is Nothing Then
                        picShape.LockAspectRatio = msoTrue
                        If pageSize = "A5" Then
                            If picShape.Width > 40 Then picShape.Width = 40
                            If picShape.Height > 40 Then picShape.Height = 40
                        Else
                            If picShape.Width > 55 Then picShape.Width = 55
                            If picShape.Height > 55 Then picShape.Height = 55
                        End If
                        accLeft = accLeft + picShape.Width + 6
                    End If
                    On Error GoTo ErrorHandler
                End If
            Next accPicIndex
            
            If accPhotoList <> "" Then
                ws.Rows(r).RowHeight = IIf(pageSize = "A5", 50, 65)
                r = r + 1
            End If
        End If
        
        r = r + 1
    End If
    
    ' --- PAYMENT STATUS ---
    lastRowPay = wsPay.Cells(wsPay.Rows.count, 1).End(xlUp).row
    foundPayment = False: totalCharges = 0: totalExp = 0: totalRec = 0: payID = ""
    
    For payRow = 2 To lastRowPay
        If Trim(UCase(wsPay.Cells(payRow, 2).value & "")) = Trim(UCase(eID)) Then
            foundPayment = True
            payID = wsPay.Cells(payRow, 1).value & ""
            If IsNumeric(wsPay.Cells(payRow, 5).value) Then totalCharges = CDbl(wsPay.Cells(payRow, 5).value)
            If IsNumeric(wsPay.Cells(payRow, 7).value) Then totalRec = CDbl(wsPay.Cells(payRow, 7).value)
            If IsNumeric(wsPay.Cells(payRow, 9).value) Then totalExp = CDbl(wsPay.Cells(payRow, 9).value)
            Exit For
        End If
    Next
    
    If foundPayment Then
        dueAmt = (totalCharges + totalExp) - totalRec
        With ws.Range("B" & r & ":G" & r)
            .Merge
            .value = "PAYMENT STATUS  |  Pay ID: " & payID
            .Font.Color = RGB(255, 255, 255): .Font.Bold = True: .Font.Size = IIf(pageSize = "A5", 10, 11)
            .Interior.Color = RGB(0, 100, 0)
            .HorizontalAlignment = xlLeft
        End With
        r = r + 1
        
        With ws.Range("B" & r & ":G" & r)
            .Merge
            If totalExp > 0 Then
                .value = "Charges: Rs." & Format(totalCharges, "0.00") & "  |  Expense: Rs." & Format(totalExp, "0.00") & "  |  Received: Rs." & Format(totalRec, "0.00") & "  |  Due: Rs." & Format(dueAmt, "0.00")
            Else
                .value = "Charges: Rs." & Format(totalCharges, "0.00") & "  |  Received: Rs." & Format(totalRec, "0.00") & "  |  Due: Rs." & Format(dueAmt, "0.00")
            End If
            If dueAmt > 0 Then
                .Font.Color = RGB(200, 0, 0): .Font.Bold = True
            Else
                .Font.Color = RGB(0, 100, 0)
            End If
            .Font.Size = IIf(pageSize = "A5", 9, 10): .HorizontalAlignment = xlLeft
        End With
        r = r + 1
    Else
        With ws.Range("B" & r & ":G" & r)
            .Merge
            .value = "Payment: No transactions found"
            .Font.Color = RGB(128, 128, 128): .Font.Size = IIf(pageSize = "A5", 9, 10)
        End With
        r = r + 1
    End If
    
    r = r + 1 ' GAP
    
    ' --- DELIVERY DETAILS HEADER (DARK GREEN) ---
    With ws.Range("B" & r & ":G" & r)
        .Merge
        .value = "DELIVERY DETAILS"
        .Font.Bold = True: .Font.Size = IIf(pageSize = "A5", 10, 12)
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(0, 100, 0)
        .HorizontalAlignment = xlLeft: .VerticalAlignment = xlCenter
    End With
    ws.Rows(r).RowHeight = IIf(pageSize = "A5", 22, 28)
    r = r + 1
    
    ' Delivery Date & Delivered By
    ws.Range("B" & r).value = "Delivery Date:"
    ws.Range("B" & r).Font.Bold = True: ws.Range("B" & r).Font.Size = IIf(pageSize = "A5", 9, 10)
    ws.Range("C" & r & ":D" & r).Merge
    ws.Range("C" & r).value = deliveryDate: ws.Range("C" & r).Font.Size = IIf(pageSize = "A5", 9, 10)
    ws.Range("E" & r).value = "Delivered By:"
    ws.Range("E" & r).Font.Bold = True: ws.Range("E" & r).Font.Size = IIf(pageSize = "A5", 9, 10)
    ws.Range("F" & r & ":G" & r).Merge
    ws.Range("F" & r).value = deliveredBy: ws.Range("F" & r).Font.Size = IIf(pageSize = "A5", 9, 10)
    r = r + 1
    
    ' OTP & Condition
    ws.Range("B" & r).value = "OTP Status:"
    ws.Range("B" & r).Font.Bold = True: ws.Range("B" & r).Font.Size = IIf(pageSize = "A5", 9, 10)
    ws.Range("C" & r).value = otpStatus: ws.Range("C" & r).Font.Size = IIf(pageSize = "A5", 9, 10)
    If UCase(otpStatus) = "VERIFIED" Then
        ws.Range("C" & r).Font.Color = RGB(0, 128, 0)
        ws.Range("C" & r).Font.Bold = True
    End If
    ws.Range("E" & r).value = "Condition:"
    ws.Range("E" & r).Font.Bold = True: ws.Range("E" & r).Font.Size = IIf(pageSize = "A5", 9, 10)
    ws.Range("F" & r & ":G" & r).Merge
    ws.Range("F" & r).value = prodCondition: ws.Range("F" & r).Font.Size = IIf(pageSize = "A5", 9, 10)
    r = r + 1
    
    ' Remarks
    If deliveryRemarks <> "" Then
        With ws.Range("B" & r & ":G" & r)
            .Merge
            .value = "Remarks: " & deliveryRemarks
            .Font.Size = IIf(pageSize = "A5", 9, 10): .Font.Italic = True
            .Font.Color = RGB(128, 0, 0)
        End With
        r = r + 1
    End If
    
    r = r + 1 ' GAP
    
    
    
    ' ===== FOOTER: TERMS + SIGNATURES =====
    Call AddDeliveryFooter(ws, r, termsFound, wsSet, terms, compName, pageSize)
    
    ws.Activate
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Set GenerateDeliverySheet = ws
    Exit Function
    
ErrorHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "Delivery Sheet Error: " & Err.Description, vbCritical
    Set GenerateDeliverySheet = Nothing
End Function
Private Sub AddDeliveryFooter(ws As Worksheet, ByRef r As Long, termsFound As Boolean, wsSet As Worksheet, terms() As String, ByVal compName As String, ByVal pageSize As String)
    Dim termRow As Long, t As Integer, foundAny As Boolean
    Dim maxTerms As Integer
    
    ' A4 me 6 terms max, A5 me 4 terms max
    maxTerms = IIf(pageSize = "A5", 4, 6)
    
    r = r + 1
    
    ' TERMS HEADER
    With ws.Range("B" & r & ":G" & r)
        .Merge
        .value = "TERMS & CONDITIONS"
        .Font.Bold = True: .Font.Size = IIf(pageSize = "A5", 10, 11)
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(128, 0, 0)
        .HorizontalAlignment = xlCenter
    End With
    ws.Rows(r).RowHeight = IIf(pageSize = "A5", 18, 22)
    r = r + 1
    
    ' Terms Content — compact
    foundAny = False
    If termsFound And Not wsSet Is Nothing Then
        For termRow = 33 To 40
            If Trim(wsSet.Cells(termRow, 2).value & "") <> "" Then
                foundAny = True
                With ws.Range("B" & r & ":G" & r)
                    .Merge
                    .value = wsSet.Cells(termRow, 2).value
                    .Font.Size = IIf(pageSize = "A5", 7, 8): .WrapText = True
                End With
                ws.Rows(r).RowHeight = IIf(pageSize = "A5", 12, 14)
                r = r + 1
                If (r - 1) >= maxTerms Then Exit For
            End If
        Next termRow
    End If
    
    If Not foundAny Then
        For t = 1 To maxTerms
            With ws.Range("B" & r & ":G" & r)
                .Merge
                .value = terms(t)
                .Font.Size = IIf(pageSize = "A5", 7, 8): .WrapText = True
            End With
            ws.Rows(r).RowHeight = IIf(pageSize = "A5", 12, 14)
            r = r + 1
        Next t
    End If
    
    ' SIGNATURES — compact
    r = r + 1
    
    ' Signature Lines
    With ws.Range("B" & r & ":C" & r)
        .Merge: .value = ""
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlBottom
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeTop).Weight = xlThin
    End With
    With ws.Range("F" & r & ":G" & r)
        .Merge: .value = ""
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlBottom
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeTop).Weight = xlThin
    End With
    r = r + 1
    
    ' Labels
    With ws.Range("B" & r & ":C" & r)
        .Merge
        .value = "Customer Signature"
        .Font.Bold = True: .Font.Size = IIf(pageSize = "A5", 9, 10)
        .HorizontalAlignment = xlCenter
    End With
    With ws.Range("F" & r & ":G" & r)
        .Merge
        .value = "Authorized Signature"
        .Font.Bold = True: .Font.Size = IIf(pageSize = "A5", 9, 10)
        .HorizontalAlignment = xlCenter
    End With
    r = r + 1
    
    ' For Company
    With ws.Range("F" & r & ":G" & r)
        .Merge
        .value = "For " & compName
        .Font.Italic = True: .Font.Size = IIf(pageSize = "A5", 9, 10): .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
    r = r + 1
    
    ' THANK YOU
    r = r + 1
    With ws.Range("B" & r & ":G" & r)
        .Merge
        .value = "*** Thank You for choosing " & compName & " ***"
        .Font.Bold = True: .Font.Size = IIf(pageSize = "A5", 10, 11)
        .Font.name = "Arial": .Font.Color = RGB(0, 100, 0)
        .HorizontalAlignment = xlCenter
        .Interior.Color = RGB(255, 250, 205)
    End With
    ws.Rows(r).RowHeight = IIf(pageSize = "A5", 20, 24)
    r = r + 1
End Sub
'========================================
' FRMDELIVERY — BTNSENDEMAIL_CLICK
' CDO Direct Background Send (Like Assignment)
'========================================
Private Sub btnSendEmail_Click()
    On Error GoTo ErrorHandler
    
    ' ===== DELIVERY CONFIRMED CHECK =====
    If Not mIsDeliveryConfirmed Then
        MsgBox "Pehle Delivery Confirm karo! OTP verification required.", vbExclamation
        Exit Sub
    End If
    
    If Trim(mEntryID) = "" Then
        MsgBox "Entry ID not found!", vbExclamation
        Exit Sub
    End If
    
    ' ===== GET CUSTOMER EMAIL =====
    Dim wsCust As Worksheet, custEmail As String
    Dim lastRow As Long, i As Long
    Set wsCust = ThisWorkbook.Sheets("Customer_Master")
    lastRow = wsCust.Cells(wsCust.Rows.count, 1).End(xlUp).row
    For i = 2 To lastRow
        If Trim(wsCust.Cells(i, 1).value) = mCustomerID Then
            custEmail = Trim(wsCust.Cells(i, 5).value & "")
            Exit For
        End If
    Next i
    
    If custEmail = "" Then
        MsgBox "Customer email not found in Customer_Master (Column E)!", vbExclamation
        Exit Sub
    End If
    
    ' ===== READ SOFTWARE_CONFIG =====
    Dim wsConfig As Worksheet
    Dim senderEmail As String, usePassword As String
    Dim smtpServer As String, smtpPort As String, sslEnable As String
    Set wsConfig = ThisWorkbook.Sheets("Software_Config")
    
    senderEmail = Trim(wsConfig.Range("B17").value & "")
    smtpServer = Trim(wsConfig.Range("B19").value & "")
    smtpPort = Trim(wsConfig.Range("B21").value & "")
    sslEnable = UCase(Trim(wsConfig.Range("B22").value & ""))
    
    If Trim(wsConfig.Range("B20").value & "") <> "" Then
        usePassword = Trim(wsConfig.Range("B20").value & "")
    Else
        usePassword = Trim(wsConfig.Range("B18").value & "")
    End If
    
    If senderEmail = "" Or smtpServer = "" Or usePassword = "" Then
        MsgBox "Email settings missing! Check Software_Config B17-B22", vbExclamation
        Exit Sub
    End If
    If smtpPort = "" Then smtpPort = "465"
    
    Dim compName As String, compMob As String
    compName = Trim(wsConfig.Range("B2").value & "")
    If compName = "" Then compName = "GLOBAL IT SOLUTIONS"
    compMob = Trim(wsConfig.Range("B9").value & "")
    
    ' ===== GET DELIVERY ID =====
    Dim wsDel As Worksheet, delID As String
    Set wsDel = ThisWorkbook.Sheets("Delivery_Master")
    lastRow = wsDel.Cells(wsDel.Rows.count, 1).End(xlUp).row
    For i = 2 To lastRow
        If Trim(wsDel.Cells(i, 2).value) = mEntryID Then
            delID = wsDel.Cells(i, 1).value
            Exit For
        End If
    Next i
    If delID = "" Then delID = "DEL00000"
    
    ' ===== BUILD EMAIL BODY =====
    Dim sBody As String
    sBody = "==============================================" & vbCrLf
    sBody = sBody & "  " & UCase(compName) & vbCrLf
    sBody = sBody & "  Mobile: " & compMob & vbCrLf
    sBody = sBody & "==============================================" & vbCrLf & vbCrLf
    sBody = sBody & "PRODUCT DELIVERY REPORT" & vbCrLf
    sBody = sBody & "----------------------------------------------" & vbCrLf
    sBody = sBody & "Delivery ID   : " & delID & vbCrLf
    sBody = sBody & "Entry ID      : " & mEntryID & vbCrLf
    sBody = sBody & "Date          : " & Format(Date, "dd-mm-yyyy") & vbCrLf & vbCrLf
    sBody = sBody & "CUSTOMER DETAILS" & vbCrLf
    sBody = sBody & "----------------------------------------------" & vbCrLf
    sBody = sBody & "Customer ID   : " & mCustomerID & vbCrLf
    sBody = sBody & "Name          : " & lblCustomerValue.caption & vbCrLf
    sBody = sBody & "Mobile        : " & lblMobileValue.caption & vbCrLf & vbCrLf
    sBody = sBody & "PRODUCT DETAILS" & vbCrLf
    sBody = sBody & "----------------------------------------------" & vbCrLf
    sBody = sBody & "Product       : " & lblProductValue.caption & vbCrLf
    sBody = sBody & "Company       : " & lblCompanyValue.caption & vbCrLf
    sBody = sBody & "Model         : " & lblModelValue.caption & vbCrLf
    sBody = sBody & "Serial        : " & lblOriginalSerialValue.caption & vbCrLf & vbCrLf
    sBody = sBody & "DELIVERY DETAILS" & vbCrLf
    sBody = sBody & "----------------------------------------------" & vbCrLf
    sBody = sBody & "Mode          : " & cmbReturnMode.value & vbCrLf
    sBody = sBody & "Delivered By  : " & cmbReturnName.value & vbCrLf
    sBody = sBody & "Date          : " & txtReturnDate.value & vbCrLf
    sBody = sBody & "OTP Status    : VERIFIED" & vbCrLf
    
    ' Payment Summary (if any)
    Dim paySummary As String
    paySummary = GetDeliveryPaySummary(mEntryID)
    If paySummary <> "" Then
        sBody = sBody & vbCrLf & "PAYMENT STATUS" & vbCrLf
        sBody = sBody & "----------------------------------------------" & vbCrLf
        sBody = sBody & paySummary & vbCrLf
    End If
    
    sBody = sBody & vbCrLf & "----------------------------------------------" & vbCrLf
    sBody = sBody & "This is a computer generated delivery report." & vbCrLf
    sBody = sBody & "For any query, contact: " & compMob & vbCrLf & vbCrLf
    sBody = sBody & "Thanks & Regards," & vbCrLf
    sBody = sBody & "Team " & compName & vbCrLf
    sBody = sBody & "=============================================="
    
    ' ===== SEND VIA CDO =====
    Dim cdoMsg As Object, cdoConf As Object
    Set cdoMsg = CreateObject("CDO.Message")
    Set cdoConf = CreateObject("CDO.Configuration")
    
    With cdoConf.Fields
        .item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate ") = 1
        .item("http://schemas.microsoft.com/cdo/configuration/smtpserver ") = GetSMTPServerAddress(smtpServer)
        .item("http://schemas.microsoft.com/cdo/configuration/smtpserverport ") = CInt(smtpPort)
        .item("http://schemas.microsoft.com/cdo/configuration/sendusername ") = senderEmail
        .item("http://schemas.microsoft.com/cdo/configuration/sendpassword ") = usePassword
        .item("http://schemas.microsoft.com/cdo/configuration/sendusing ") = 2
        If sslEnable = "YES" Then
            .item("http://schemas.microsoft.com/cdo/configuration/smtpusessl ") = True
        Else
            .item("http://schemas.microsoft.com/cdo/configuration/smtpusessl ") = False
        End If
        .item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout ") = 60
        .Update
    End With
    
    With cdoMsg
        Set .Configuration = cdoConf
        .From = senderEmail
        .To = custEmail
        .subject = "Product Delivered — " & mEntryID & " — " & lblProductValue.caption
        .TextBody = sBody
        .Send
    End With
    
    Set cdoMsg = Nothing
    Set cdoConf = Nothing
    
    MsgBox "Email sent successfully!" & vbCrLf & "To: " & custEmail, vbInformation, "Mail Sent"
    Exit Sub
    
ErrorHandler:
    MsgBox "Email Error: " & Err.Description & vbCrLf & vbCrLf & "Check:" & vbCrLf & "1. Internet connected?" & vbCrLf & "2. App Password correct in B20?" & vbCrLf & "3. Gmail: Enable 2-Step Verification & use App Password", vbCritical
End Sub

'========================================
' HELPER: Get SMTP Server Address
'========================================
Private Function GetSMTPServerAddress(smtpServer As String) As String
    Dim s As String
    s = Trim(LCase(smtpServer))
    If s = "" Then
        GetSMTPServerAddress = "smtp.gmail.com"
    ElseIf InStr(s, "gmail") > 0 Then
        GetSMTPServerAddress = "smtp.gmail.com"
    ElseIf InStr(s, "outlook") > 0 Or InStr(s, "hotmail") > 0 Then
        GetSMTPServerAddress = "smtp-mail.outlook.com"
    ElseIf InStr(s, "yahoo") > 0 Then
        GetSMTPServerAddress = "smtp.mail.yahoo.com"
    Else
        GetSMTPServerAddress = smtpServer
    End If
End Function


'========================================
' HELPER: Payment Summary for Delivery Email
'========================================
Private Function GetDeliveryPaySummary(entryID As String) As String
    On Error Resume Next
    Dim ws As Worksheet, lastRow As Long, r As Long
    Dim tc As Double, tr As Double, te As Double, td As Double, hasData As Boolean
    Set ws = ThisWorkbook.Sheets("Payment_Master")
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    For r = 2 To lastRow
        If UCase(Trim(ws.Cells(r, 2).value)) = UCase(Trim(entryID)) Then
            hasData = True: tc = tc + val(ws.Cells(r, 5).value): tr = tr + val(ws.Cells(r, 7).value)
            te = te + val(ws.Cells(r, 9).value): td = td + val(ws.Cells(r, 10).value)
        End If
    Next r
    If hasData Then GetDeliveryPaySummary = "Charges: Rs." & Format(tc, "0.00") & " | Expense: Rs." & Format(te, "0.00") & " | Received: Rs." & Format(tr, "0.00") & " | Due: Rs." & Format(td, "0.00") Else GetDeliveryPaySummary = ""
    On Error GoTo 0
End Function


'========== GET PRODUCT CONDITION BASED ON ACTUAL STATUS ==========
Private Function GetProductCondition(entryID As String) As String
    Dim wsAssign As Worksheet, wsWar As Worksheet, wsSer As Worksheet
    Dim lastRow As Long, i As Long, j As Long
    Dim assignStatus As String, retStatus As String, workStatus As String
    Dim solutionApplied As String, newSerial As String, rejectReason As String
    
    ' DEFAULT
    GetProductCondition = "Pending - Under Process"
    
    ' ===== STEP 1: Check Assign_Master Status (Column 32) =====
    On Error Resume Next
    Set wsAssign = ThisWorkbook.Sheets("Assign_Master")
    On Error GoTo 0
    
    If wsAssign Is Nothing Then Exit Function
    
    lastRow = wsAssign.Cells(wsAssign.Rows.count, 1).End(xlUp).row
    For i = 2 To lastRow
        If Trim(wsAssign.Cells(i, 2).value & "") = entryID Then
            assignStatus = UCase(Trim(wsAssign.Cells(i, 32).value & ""))
            
            Select Case assignStatus
                Case "PENDING"
                    ' Default pending
                    GetProductCondition = "Pending - Under Process"
                    
                    ' Check Warranty_Return_Master for actual outcome
                    On Error Resume Next
                    Set wsWar = ThisWorkbook.Sheets("Warranty_Return_Master")
                    On Error GoTo 0
                    If Not wsWar Is Nothing Then
                        Dim warLastRow As Long
                        warLastRow = wsWar.Cells(wsWar.Rows.count, 1).End(xlUp).row
                        For j = 2 To warLastRow
                            If Trim(wsWar.Cells(j, 3).value & "") = entryID Then
                                retStatus = UCase(Trim(wsWar.Cells(j, 19).value & ""))
                                Select Case retStatus
                                    Case "RECEIVE_REPAIRED"
                                        GetProductCondition = "Working / Tested OK (Repaired)"
                                    Case "RECEIVE_REPLACED"
                                        newSerial = Trim(wsWar.Cells(j, 21).value & "")
                                        If newSerial <> "" Then
                                            GetProductCondition = "Replaced with New Unit | New Serial: " & newSerial
                                        Else
                                            GetProductCondition = "Replaced with New Unit"
                                        End If
                                    Case "RECEIVE_REJECTED"
                                        rejectReason = Trim(wsWar.Cells(j, 22).value & "")
                                        If rejectReason <> "" Then
                                            GetProductCondition = "Rejected / Not Repairable | Reason: " & rejectReason
                                        Else
                                            GetProductCondition = "Rejected / Not Repairable"
                                        End If
                                    Case "RETURNED"
                                        GetProductCondition = "Returned from Vendor"
                                End Select
                                Exit For
                            End If
                        Next j
                    End If
                    Exit Function
                    
                Case "ACTIVE"
                    GetProductCondition = "Service In Progress - Not Completed"
                    On Error Resume Next
                    Set wsSer = ThisWorkbook.Sheets("Service_WorkLog_Expense")
                    On Error GoTo 0
                    If Not wsSer Is Nothing Then
                        Dim serLastRow As Long
                        serLastRow = wsSer.Cells(wsSer.Rows.count, 1).End(xlUp).row
                        For j = 2 To serLastRow
                            If Trim(wsSer.Cells(j, 3).value & "") = entryID Then
                                workStatus = UCase(Trim(wsSer.Cells(j, 13).value & ""))
                                solutionApplied = UCase(Trim(wsSer.Cells(j, 15).value & ""))
                                Select Case workStatus
                                    Case "COMPLETED"
                                        If InStr(solutionApplied, "REPLACE") > 0 Then
                                            GetProductCondition = "Part(s) Replaced - Working / Tested OK"
                                        Else
                                            GetProductCondition = "Working / Tested OK (Service Completed)"
                                        End If
                                    Case "PENDING", "IN_PROGRESS"
                                        GetProductCondition = "Service In Progress - Not Completed"
                                    Case "RETURN"
                                        GetProductCondition = "Returned to Customer"
                                    Case "REJECTED"
                                        GetProductCondition = "Rejected / Not Repairable"
                                End Select
                                Exit For
                            End If
                        Next j
                    End If
                    Exit Function
                    
                Case "COMPLETED"
                    GetProductCondition = "Work Completed - Tested OK"
                    Exit Function
                Case "RETURNED"
                    GetProductCondition = "Returned to Customer"
                    Exit Function
                Case "REJECTED"
                    GetProductCondition = "Rejected / Not Repairable"
                    Exit Function
                Case "DELIVERED"
                    GetProductCondition = "Delivered to Customer"
                    Exit Function
            End Select
            Exit For
        End If
    Next i
End Function

Private Sub btnExportDeliveryPDF_Click()
    On Error GoTo ErrorHandler
    
    ' ===== ENTRY ID CHECK =====
    If Trim(m_DeliveryEntryID) = "" Then
        MsgBox "Entry ID not found! Pehle entry select karo.", vbExclamation
        Exit Sub
    End If
    
    ' ===== OTP CHECK =====
    If Not mIsDeliveryConfirmed Then
        MsgBox "Pehle Delivery Confirm karo! OTP verification required.", vbExclamation
        Exit Sub
    End If
    
    ' ===== FOLDER =====
    Dim configFolder As String
    configFolder = ThisWorkbook.path & "\Configuration\"
    If Dir(configFolder, vbDirectory) = "" Then MkDir configFolder
    
    ' ===== FILE NAME =====
    Dim safeName As String, fileName As String, pdfPath As String
    safeName = CleanFileName(mCustomerID & "_" & m_DeliveryEntryID)
    fileName = "DELIVERY_" & safeName & "_" & Format(Now, "ddmmyyyy_hhmmss") & ".pdf"
    pdfPath = configFolder & fileName
    
    ' ===== GENERATE & EXPORT =====
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    Dim wsDelivery As Worksheet
    Set wsDelivery = GenerateDeliverySheet(m_DeliveryEntryID)
    
    If wsDelivery Is Nothing Then
        MsgBox "PDF generation failed!", vbCritical
        GoTo Cleanup
    End If
    
    ' Export to PDF — Same format as printout
    wsDelivery.ExportAsFixedFormat Type:=xlTypePDF, fileName:=pdfPath, Quality:=xlQualityStandard
    
    ' Temp sheet delete
    Application.DisplayAlerts = False
    wsDelivery.Delete
    Application.DisplayAlerts = True
    
    ' Open PDF
    ShellExecute 0, "open", pdfPath, vbNullString, vbNullString, 1
    
    MsgBox "PDF Saved & Opened!" & vbCrLf & pdfPath, vbInformation
    
Cleanup:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Exit Sub
    
ErrorHandler:
    MsgBox "PDF Error: " & Err.Description, vbCritical
    Resume Cleanup
End Sub
Private Sub btnSearch_Click()
    Dim searchText As String
    searchText = Trim(txtSearch.value & "")
    If searchText = "" Then
        MsgBox "Search text daalo! (Name, Mobile, Model, Serial, Product)", vbExclamation
        txtSearch.SetFocus
        Exit Sub
    End If
    
    Dim wsJob As Worksheet, wsCust As Worksheet
    Dim lastRow As Long, i As Long, j As Long
    Dim foundEntries As String
    Dim matchCount As Long
    
    Set wsJob = ThisWorkbook.Sheets("Job_Product")
    Set wsCust = ThisWorkbook.Sheets("Customer_Master")
    
    foundEntries = ""
    matchCount = 0
    
    ' === SEARCH 1: Job_Product (Product, Company, Model, Serial, Problem) ===
    lastRow = wsJob.Cells(wsJob.Rows.count, 1).End(xlUp).row
    For i = 2 To lastRow
        Dim eID As String, cID As String, prod As String, comp As String
        Dim modl As String, ser As String, prob As String
        eID = Trim(wsJob.Cells(i, 1).value & "")
        cID = Trim(wsJob.Cells(i, 2).value & "")
        prod = Trim(wsJob.Cells(i, 6).value & "")
        comp = Trim(wsJob.Cells(i, 7).value & "")
        modl = Trim(wsJob.Cells(i, 8).value & "")
        ser = Trim(wsJob.Cells(i, 9).value & "")
        prob = Trim(wsJob.Cells(i, 13).value & "")
        
        If InStr(1, eID, searchText, vbTextCompare) > 0 Or _
           InStr(1, prod, searchText, vbTextCompare) > 0 Or _
           InStr(1, comp, searchText, vbTextCompare) > 0 Or _
           InStr(1, modl, searchText, vbTextCompare) > 0 Or _
           InStr(1, ser, searchText, vbTextCompare) > 0 Or _
           InStr(1, prob, searchText, vbTextCompare) > 0 Then
            foundEntries = foundEntries & eID & " | " & GetCustomerName(cID) & " | " & prod & " (" & comp & ")" & vbCrLf
            matchCount = matchCount + 1
        End If
    Next i
    
    ' === SEARCH 2: Customer_Master (Name, Mobile, Address) ===
    lastRow = wsCust.Cells(wsCust.Rows.count, 1).End(xlUp).row
    For i = 2 To lastRow
        Dim cName As String, cMob As String, cAddr As String
        cID = Trim(wsCust.Cells(i, 1).value & "")
        cName = Trim(wsCust.Cells(i, 3).value & "")
        cMob = Trim(wsCust.Cells(i, 2).value & "")
        cAddr = Trim(wsCust.Cells(i, 4).value & "")
        
        If InStr(1, cName, searchText, vbTextCompare) > 0 Or _
           InStr(1, cMob, searchText, vbTextCompare) > 0 Or _
           InStr(1, cAddr, searchText, vbTextCompare) > 0 Then
            ' Find all entries for this customer
            For j = 2 To wsJob.Cells(wsJob.Rows.count, 1).End(xlUp).row
                If Trim(wsJob.Cells(j, 2).value & "") = cID Then
                    eID = Trim(wsJob.Cells(j, 1).value & "")
                    prod = Trim(wsJob.Cells(j, 6).value & "")
                    comp = Trim(wsJob.Cells(j, 7).value & "")
                    ' Avoid duplicates
                    If InStr(foundEntries, eID & " |") = 0 Then
                        foundEntries = foundEntries & eID & " | " & cName & " | " & prod & " (" & comp & ")" & vbCrLf
                        matchCount = matchCount + 1
                    End If
                End If
            Next j
        End If
    Next i
    
    If matchCount = 0 Then
        MsgBox "Koi entry nahi mili!" & vbCrLf & "Search: " & searchText, vbExclamation
        Exit Sub
    End If
    
    ' === SHOW RESULTS & SELECT ===
    Dim selectedEntry As String
    selectedEntry = InputBox("Select Entry ID from list:" & vbCrLf & vbCrLf & foundEntries & vbCrLf & "Type exact Entry ID:", "Search Results (" & matchCount & " found)")
    
    If Trim(selectedEntry) = "" Then Exit Sub
    
    ' === DIRECT LOAD (Bypass ComboBox) ===
    mEntryID = Trim(selectedEntry)
    m_DeliveryEntryID = mEntryID
    
    ' Load all details
    Call LoadEntryDetailsUniversal(mEntryID)
    Call LoadWorkDone(mEntryID)
    Call LoadAccessories(mEntryID)
    
    ' Check delivery status
    Dim wsDel As Worksheet
    On Error Resume Next
    Set wsDel = ThisWorkbook.Sheets("Delivery_Master")
    On Error GoTo 0
    mIsDelivered = IsDelivered(mEntryID, wsDel)
    
    If mIsDelivered Then
        btnDeliver.enabled = False
        btnDeliver.caption = "ALREADY DELIVERED"
        Call LoadDeliveryDetails(mEntryID)
        Call SetBottomButtonsEnabled(True)
    Else
        btnDeliver.enabled = True
        btnDeliver.caption = "DELIVER"
        Call SetBottomButtonsEnabled(False)
    End If
    
    ' Load payment
    Call LoadPaymentDetails(mEntryID)
    
    ' Show in combo for reference
    cmbEntryID.value = mEntryID
    
    MsgBox "Entry loaded: " & mEntryID & vbCrLf & "Customer: " & lblCustomerValue.caption, vbInformation
End Sub

' ===== HELPER: Get Customer Name =====
Private Function GetCustomerName(custID As String) As String
    Dim ws As Worksheet, lr As Long, i As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Customer_Master")
    On Error GoTo 0
    If ws Is Nothing Then
        GetCustomerName = ""
        Exit Function
    End If
    lr = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    For i = 2 To lr
        If Trim(ws.Cells(i, 1).value & "") = custID Then
            GetCustomerName = Trim(ws.Cells(i, 3).value & "")
            Exit Function
        End If
    Next i
    GetCustomerName = ""
End Function
'===========================================
' LOAD PAYMENT LIST — ENTRY FORM STYLE
'===========================================
Public Sub LoadPaymentDetails(entryID As String)
    Dim ws As Worksheet
    Dim lastRow As Long, r As Long
    Dim itm As listItem

    If Trim(entryID) = "" Then
        Me.lstPaymentDetails.ListItems.Clear
        Call UpdatePaymentButtons
        Exit Sub
    End If

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Payment_Master")
    On Error GoTo 0

    If ws Is Nothing Then
        Me.lstPaymentDetails.ListItems.Clear
        Call UpdatePaymentButtons
        Exit Sub
    End If

    Me.lstPaymentDetails.ListItems.Clear
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row

    For r = 2 To lastRow
        If Trim(ws.Cells(r, 2).value & "") = Trim(entryID & "") Then
            If Trim(ws.Cells(r, 6).value & "") = "" Then   ' Active (DeletedBy blank)
                Set itm = Me.lstPaymentDetails.ListItems.Add
                itm.text = ws.Cells(r, 1).value              ' A: PaymentID
                itm.SubItems(1) = ws.Cells(r, 2).value       ' B: EntryID
                itm.SubItems(2) = ws.Cells(r, 8).value       ' H: Type
                itm.SubItems(3) = ws.Cells(r, 9).value       ' I: Category
                itm.SubItems(4) = ws.Cells(r, 15).value      ' O: PaymentMode
                itm.SubItems(5) = ws.Cells(r, 10).value      ' J: Name
                itm.SubItems(6) = ws.Cells(r, 11).value      ' K: Company
                itm.SubItems(7) = ws.Cells(r, 12).value      ' L: Model
                itm.SubItems(8) = ws.Cells(r, 13).value      ' M: Serial
                itm.SubItems(9) = ws.Cells(r, 14).value      ' N: Qty
                itm.SubItems(10) = Format(ws.Cells(r, 16).value, "0.00") ' P: Amount
                itm.SubItems(11) = ws.Cells(r, 17).value     ' Q: Warranty
            End If
        End If
    Next r

    Call UpdatePaymentButtons
End Sub

'===========================================
' CHECK IF PAYMENT EXISTS
'===========================================
Private Function GetExistingPaymentID(entryID As String) As String
    Dim ws As Worksheet
    Dim r As Long, lastRow As Long

    If Trim(entryID) = "" Then Exit Function

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Payment_Master")
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row

    For r = 2 To lastRow
        If Trim(ws.Cells(r, 2).value & "") = Trim(entryID & "") Then
            If Trim(ws.Cells(r, 6).value & "") = "" Then
                GetExistingPaymentID = ws.Cells(r, 1).value
                Exit Function
            End If
        End If
    Next r

    GetExistingPaymentID = ""
End Function

'===========================================
' UPDATE BUTTON STATES
' Payment exists = Edit Enable, Add Disable
' No payment = Add Enable, Edit Disable
'===========================================
Private Sub UpdatePaymentButtons()
    If Me.lstPaymentDetails.ListItems.count > 0 Then
        btnAddPayment.enabled = False
        btnEditPayment.enabled = True
    Else
        btnAddPayment.enabled = True
        btnEditPayment.enabled = False
    End If
    Me.Repaint
End Sub

'===========================================
' ADD CHARGES BUTTON
'===========================================
Private Sub btnAddPayment_Click()
    Dim payForm As frmPaymentSection
    Dim entryID As String
    Dim existingPayID As String

    entryID = Trim(mEntryID)
    If entryID = "" Then
        MsgBox "Please select Entry ID first!", vbExclamation
        Exit Sub
    End If

    existingPayID = GetExistingPaymentID(entryID)

    Set payForm = New frmPaymentSection
    Set payForm.parentForm = Me
    
    ' ===== YE LINE ADD KARO =====
    payForm.customerID = mCustomerID
    ' ============================

    payForm.cmbEntryID.value = entryID
    payForm.cmbEntryID.enabled = False

    If existingPayID <> "" Then
        payForm.LoadPaymentForEdit existingPayID
    End If

    payForm.Show vbModal

    Call LoadPaymentDetails(entryID)
    Set payForm = Nothing
End Sub
'===========================================
' EDIT CHARGES BUTTON
'===========================================
Private Sub btnEditPayment_Click()
    Dim payID As String

    On Error Resume Next
    payID = Me.lstPaymentDetails.selectedItem.text
    On Error GoTo 0

    If payID = "" Then
        MsgBox "Please select a payment first!", vbExclamation
        Exit Sub
    End If

    Dim payForm As New frmPaymentSection
    Set payForm.parentForm = Me

    payForm.LoadPaymentForEdit payID
    payForm.cmbEntryID.enabled = False

    payForm.Show vbModal

    Call LoadPaymentDetails(mEntryID)
End Sub

'===========================================
' LISTVIEW DOUBLE CLICK
'===========================================
Private Sub lstPaymentDetails_DblClick()
    Call btnEditPayment_Click
End Sub

'===========================================
' REFRESH WRAPPER (frmPaymentSection callback)
'===========================================
Public Sub LoadPaymentList(customerID As String)
    Call LoadPaymentDetails(mEntryID)
End Sub

    '===========================================
' SETUP PAYMENT LISTVIEW — ENTRY FORM STYLE
'===========================================
Sub SetupPaymentListView()
    With Me.lstPaymentDetails
        .View = lvwReport
        .FullRowSelect = True
        .Gridlines = True
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "Payment ID", 80
        .ColumnHeaders.Add , , "Entry ID", 80
        .ColumnHeaders.Add , , "Type", 80
        .ColumnHeaders.Add , , "Category", 100
        .ColumnHeaders.Add , , "PayMode", 80
        .ColumnHeaders.Add , , "Name", 120
        .ColumnHeaders.Add , , "Company", 100
        .ColumnHeaders.Add , , "Model", 100
        .ColumnHeaders.Add , , "Serial", 100
        .ColumnHeaders.Add , , "Qty", 50
        .ColumnHeaders.Add , , "Amount", 80
        .ColumnHeaders.Add , , "Warranty", 80
    End With
End Sub
Private Sub UserForm_Activate()
    Dim entryID As String
    Dim i As Long
    Dim found As Boolean
    
    entryID = Trim(Me.Tag)
    
    If entryID <> "" Then
        ' Check if entryID exists in combo list
        found = False
        For i = 0 To cmbEntryID.ListCount - 1
            If UCase(Trim(cmbEntryID.List(i))) = UCase(entryID) Then
                found = True
                Exit For
            End If
        Next i
        
        ' If not found, add it temporarily
        If Not found Then
            cmbEntryID.AddItem entryID
        End If
        
        ' Now set value safely
        On Error Resume Next
        cmbEntryID.value = entryID
        On Error GoTo 0
        
        Me.Tag = ""
    End If
End Sub

'=====================================================
' SAME CUSTOMER CHECKBOX — HIDE/SHOW RETURN DETAILS
'=====================================================
Private Sub chkSameCustomer_Click()
    ' ===== SAFETY CHECK =====
    On Error Resume Next
    If Me.chkSameCustomer Is Nothing Then
        MsgBox "Checkbox control not found!", vbExclamation
        Exit Sub
    End If
    On Error GoTo 0
    
    Dim chkValue As Integer
    chkValue = 0
    
    ' Safe value read
    On Error Resume Next
    chkValue = Me.chkSameCustomer.value
    If Err.Number <> 0 Then
        Err.Clear
        chkValue = 0
    End If
    On Error GoTo 0
    
    ' ===== CHECKED (1) = HIDE RETURN DETAILS =====
    If chkValue = 1 Then
        ' --- Labels hide ---
        SetControlVisible "lblReturnMode", False
        SetControlVisible "lblReturnName", False
        SetControlVisible "lblReturnMobile", False
        SetControlVisible "lblReturnAddress", False
        SetControlVisible "lblReturnDocket", False
        SetControlVisible "lblReturnDate", False
        
        ' --- Inputs hide ---
        SetControlVisible "cmbReturnMode", False
        SetControlVisible "cmbReturnName", False
        SetControlVisible "txtReturnMobile", False
        SetControlVisible "txtReturnAddress", False
        SetControlVisible "txtReturnDocket", False
        
        ' --- Buttons hide ---
        SetControlVisible "btnAddReturnMode", False
        SetControlVisible "btnAddReturnName", False
        
        ' --- Auto Fill Customer Details ---
        On Error Resume Next
        cmbReturnMode.value = "SELF"
        cmbReturnName.value = lblCustomerValue.caption
        txtReturnMobile.value = lblMobileValue.caption
        txtReturnAddress.value = lblAddressValue.caption
        On Error GoTo 0
        
    ' ===== UNCHECKED (0) = SHOW RETURN DETAILS =====
    Else
        ' --- Labels show ---
        SetControlVisible "lblReturnMode", True
        SetControlVisible "lblReturnName", True
        SetControlVisible "lblReturnMobile", True
        SetControlVisible "lblReturnAddress", True
        SetControlVisible "lblReturnDocket", True
        SetControlVisible "lblReturnDate", True
        
        ' --- Inputs show ---
        SetControlVisible "cmbReturnMode", True
        SetControlVisible "cmbReturnName", True
        SetControlVisible "txtReturnMobile", True
        SetControlVisible "txtReturnAddress", True
        SetControlVisible "txtReturnDocket", True
        
        ' --- Buttons show ---
        SetControlVisible "btnAddReturnMode", True
        SetControlVisible "btnAddReturnName", True
        
        ' --- Clear values ---
        On Error Resume Next
        cmbReturnMode.value = ""
        cmbReturnName.value = ""
        txtReturnMobile.value = ""
        txtReturnAddress.value = ""
        txtReturnDocket.value = ""
        On Error GoTo 0
    End If
    
    Me.Repaint
End Sub

'===========================================
' HELPER: Hide Control by Name
'===========================================
Private Sub HideControl(ctrlName As String)
    On Error Resume Next
    Me.Controls(ctrlName).visible = False
    On Error GoTo 0
End Sub

'===========================================
' HELPER: Show Control by Name
'===========================================
Private Sub ShowControl(ctrlName As String)
    On Error Resume Next
    Me.Controls(ctrlName).visible = True
    On Error GoTo 0
End Sub
'=====================================================
' SAFE CONTROL VISIBILITY SETTER
'=====================================================
Private Sub SetControlVisible(ctrlName As String, visible As Boolean)
    On Error Resume Next
    Me.Controls(ctrlName).visible = visible
    On Error GoTo 0
End Sub
