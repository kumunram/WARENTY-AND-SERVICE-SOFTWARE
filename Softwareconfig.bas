Attribute VB_Name = "Softwareconfig"
Sub CreateConfigSheet()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Software_Config")
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.name = "Software_Config"
    End If
    
    ' Headers
    ws.Range("A1").value = "Setting_Name"
    ws.Range("B1").value = "Value"
    
    ' Company Info (Tab 1)
    ws.Range("A2").value = "Company_Name"
    ws.Range("A3").value = "Address_Line1"
    ws.Range("A4").value = "Address_Line2"
    ws.Range("A5").value = "City"
    ws.Range("A6").value = "PIN_Code"
    ws.Range("A7").value = "State"
    ws.Range("A8").value = "Country"
    ws.Range("A9").value = "Mobile_Number"
    ws.Range("A10").value = "Email_ID"
    ws.Range("A11").value = "Website"
    ws.Range("A12").value = "Logo_Path"
    ws.Range("A13").value = "Receipt_Header"
    
    ' Communication Settings (Tab 2)
    ws.Range("A16").value = "WhatsApp_Number"
    ws.Range("A17").value = "Email_ID"
    ws.Range("A18").value = "Email_Password"
    ws.Range("A19").value = "SMTP_Server"
    ws.Range("A20").value = "SMTP_Port"
    ws.Range("A21").value = "SSL_Enable"
    
    ws.visible = xlSheetVeryHidden
    
    MsgBox "Config sheet created!", vbInformation
End Sub

Sub AddFinancialSettings()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Software_Config")
    
    ' Financial & Legal Settings
    ws.Range("A23").value = "Country"
    ws.Range("A24").value = "State"
    ws.Range("A25").value = "GST_Number"
    ws.Range("A26").value = "Service_Tax_Number"
    ws.Range("A27").value = "Judicial_Area"
    ws.Range("A28").value = "Financial_Start"
    ws.Range("A29").value = "Financial_End"
    ws.Range("A30").value = "Current_Year"
    
    MsgBox "Financial settings added!", vbInformation
End Sub
Sub AddLicenseSettings()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Software_Config")
    
    ' License & Security Settings
    ws.Range("A31").value = "License_Type"
    ws.Range("A32").value = "Activation_Date"
    ws.Range("A33").value = "Expiry_Date"
    ws.Range("A34").value = "Renew_Date"
    ws.Range("A35").value = "Days_Remaining"
    ws.Range("A36").value = "Status"
    ws.Range("A37").value = "Auto_Archive"        ' Yes/No (Checkbox)
    ws.Range("A38").value = "Archive_Location"
    ws.Range("A44").value = "Archiving_Mode"      ' Daily/Weekly/Monthly/Yearly (ComboBox)
    
    MsgBox "License settings structure added!", vbInformation
End Sub
Sub AddUpdateSettings()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Software_Config")
    
    ' Software Update Settings
    ws.Range("A56").value = "Installed_Version"
    ws.Range("A51").value = "Release_Date"
    ws.Range("A52").value = "License_Status"
    ws.Range("A53").value = "New_File_Path"
    ws.Range("A54").value = "Backup_Location"
    ws.Range("A55").value = "Restore_File_Path"
    
    MsgBox "Update settings added!", vbInformation
End Sub

Sub Create_ServiceStation_Register_Sheet()
    Dim ws As Worksheet
    Dim sheetExists As Boolean
    
    ' Check if sheet already exists
    sheetExists = False
    For Each ws In ThisWorkbook.Worksheets
        If ws.name = "ServiceStation_Register" Then
            sheetExists = True
            Exit For
        End If
    Next ws
    
    ' If sheet exists, delete it
    If sheetExists Then
        Application.DisplayAlerts = False
        ThisWorkbook.Worksheets("ServiceStation_Register").Delete
        Application.DisplayAlerts = True
    End If
    
    ' Create new sheet
    Set ws = ThisWorkbook.Worksheets.Add
    ws.name = "ServiceStation_Register"
    
    ' Add Headers (Row 1)
    With ws
        .Range("A1").value = "StationID"
        .Range("B1").value = "StationName"
        .Range("C1").value = "ContactPerson"
        .Range("D1").value = "Mobile"
        .Range("E1").value = "Email"
        .Range("F1").value = "Address"
        .Range("G1").value = "City"
        .Range("H1").value = "State"
        .Range("I1").value = "PIN"
        .Range("J1").value = "GSTNumber"
        .Range("K1").value = "ServicesProvided"
        .Range("L1").value = "Status"
        .Range("M1").value = "CreatedDate"
        
        ' Format Headers
        .Range("A1:M1").Font.Bold = True
        .Range("A1:M1").Interior.Color = RGB(200, 220, 255)
        .Range("A1:M1").HorizontalAlignment = xlCenter
        
        ' Set Column Widths
        .Columns("A").ColumnWidth = 12
        .Columns("B").ColumnWidth = 25
        .Columns("C").ColumnWidth = 20
        .Columns("D").ColumnWidth = 15
        .Columns("E").ColumnWidth = 25
        .Columns("F").ColumnWidth = 30
        .Columns("G").ColumnWidth = 15
        .Columns("H").ColumnWidth = 15
        .Columns("I").ColumnWidth = 10
        .Columns("J").ColumnWidth = 15
        .Columns("K").ColumnWidth = 25
        .Columns("L").ColumnWidth = 12
        .Columns("M").ColumnWidth = 18
        
        ' Add Sample Data (Row 2)
        .Range("A2").value = "STA00001"
        .Range("B2").value = "Dell Authorized Service"
        .Range("C2").value = "Mr. Sharma"
        .Range("D2").value = "9876543210"
        .Range("E2").value = "dell.service@example.com"
        .Range("F2").value = "123, MG Road"
        .Range("G2").value = "Mumbai"
        .Range("H2").value = "Maharashtra"
        .Range("I2").value = "400001"
        .Range("J2").value = "GST123456789"
        .Range("K2").value = "Laptop, Desktop"
        .Range("L2").value = "ACTIVE"
        .Range("M2").value = Format(Now, "dd-mm-yyyy")
        
        ' Add Sample Data (Row 3)
        .Range("A3").value = "STA00002"
        .Range("B3").value = "HP Service Center"
        .Range("C3").value = "Mr. Kumar"
        .Range("D3").value = "9876543211"
        .Range("E3").value = "hp.service@example.com"
        .Range("F3").value = "456, Park Street"
        .Range("G3").value = "Delhi"
        .Range("H3").value = "Delhi"
        .Range("I3").value = "110001"
        .Range("J3").value = "GST987654321"
        .Range("K3").value = "Laptop, Printer"
        .Range("L3").value = "ACTIVE"
        .Range("M3").value = Format(Now, "dd-mm-yyyy")
    End With
    
    MsgBox "ServiceStation_Register Sheet Created Successfully!", vbInformation
End Sub

Sub Create_Assign_Master_Sheet()
    Dim ws As Worksheet
    Dim sheetExists As Boolean
    
    ' Check if sheet already exists
    sheetExists = False
    For Each ws In ThisWorkbook.Worksheets
        If ws.name = "Assign_Master" Then
            sheetExists = True
            Exit For
        End If
    Next ws
    
    ' If sheet exists, delete it
    If sheetExists Then
        Application.DisplayAlerts = False
        ThisWorkbook.Worksheets("Assign_Master").Delete
        Application.DisplayAlerts = True
    End If
    
    ' Create new sheet
    Set ws = ThisWorkbook.Worksheets.Add
    ws.name = "Assign_Master"
    
    ' Add Headers (Row 1)
    With ws
        .Range("A1").value = "AssignID"
        .Range("B1").value = "EntryID"
        .Range("C1").value = "CustomerID"
        .Range("D1").value = "CustomerMobile"
        .Range("E1").value = "CustomerName"
        .Range("F1").value = "ProductType"
        .Range("G1").value = "Company"
        .Range("H1").value = "Model"
        .Range("I1").value = "SerialNumber"
        .Range("J1").value = "WarrantyStatus"
        .Range("K1").value = "AssignType"
        .Range("L1").value = "AssignMode"
        .Range("M1").value = "AssignToID"
        .Range("N1").value = "AssignToName"
        .Range("O1").value = "AssignToMobile"
        .Range("P1").value = "AssignToEmail"
        .Range("Q1").value = "AssignToAddress"
        .Range("R1").value = "OurCompanyName"
        .Range("S1").value = "OurCompanyAddress"
        .Range("T1").value = "OurCompanyMobile"
        .Range("U1").value = "OurCompanyEmail"
        .Range("V1").value = "AssignDate"
        .Range("W1").value = "CourierMode"
        .Range("X1").value = "CourierName"
        .Range("Y1").value = "DocketNumber"
        .Range("Z1").value = "SendDate"
        .Range("AA1").value = "ExpectedReturnDate"
        .Range("AB1").value = "CourierCharges"
        .Range("AC1").value = "ProblemDescription"
        .Range("AD1").value = "AccessoriesSent"
        .Range("AE1").value = "Remarks"
        .Range("AF1").value = "AssignStatus"
        .Range("AG1").value = "CreatedBy"
        .Range("AH1").value = "CreatedDate"
        
        ' Format Headers
        .Range("A1:AH1").Font.Bold = True
        .Range("A1:AH1").Interior.Color = RGB(200, 255, 200)
        .Range("A1:AH1").HorizontalAlignment = xlCenter
        
        ' Set Column Widths
        .Columns("A").ColumnWidth = 12
        .Columns("B").ColumnWidth = 12
        .Columns("C").ColumnWidth = 12
        .Columns("D").ColumnWidth = 15
        .Columns("E").ColumnWidth = 20
        .Columns("F").ColumnWidth = 15
        .Columns("G").ColumnWidth = 15
        .Columns("H").ColumnWidth = 15
        .Columns("I").ColumnWidth = 15
        .Columns("J").ColumnWidth = 12
        .Columns("K").ColumnWidth = 12
        .Columns("L").ColumnWidth = 15
        .Columns("M").ColumnWidth = 12
        .Columns("N").ColumnWidth = 25
        .Columns("O").ColumnWidth = 15
        .Columns("P").ColumnWidth = 25
        .Columns("Q").ColumnWidth = 30
        .Columns("R").ColumnWidth = 20
        .Columns("S").ColumnWidth = 30
        .Columns("T").ColumnWidth = 15
        .Columns("U").ColumnWidth = 25
        .Columns("V").ColumnWidth = 12
        .Columns("W").ColumnWidth = 12
        .Columns("X").ColumnWidth = 15
        .Columns("Y").ColumnWidth = 12
        .Columns("Z").ColumnWidth = 12
        .Columns("AA").ColumnWidth = 15
        .Columns("AB").ColumnWidth = 12
        .Columns("AC").ColumnWidth = 30
        .Columns("AD").ColumnWidth = 25
        .Columns("AE").ColumnWidth = 25
        .Columns("AF").ColumnWidth = 15
        .Columns("AG").ColumnWidth = 12
        .Columns("AH").ColumnWidth = 18
    End With
    
    MsgBox "Assign_Master Sheet Created Successfully!", vbInformation
End Sub

Sub Create_Assign_Accessories_Sheet()
    Dim ws As Worksheet
    Dim sheetExists As Boolean
    
    ' Check if sheet already exists
    sheetExists = False
    For Each ws In ThisWorkbook.Worksheets
        If ws.name = "Assign_Accessories" Then
            sheetExists = True
            Exit For
        End If
    Next ws
    
    ' If sheet exists, delete it
    If sheetExists Then
        Application.DisplayAlerts = False
        ThisWorkbook.Worksheets("Assign_Accessories").Delete
        Application.DisplayAlerts = True
    End If
    
    ' Create new sheet
    Set ws = ThisWorkbook.Worksheets.Add
    ws.name = "Assign_Accessories"
    
    ' Add Headers (Row 1)
    With ws
        .Range("A1").value = "AssignID"
        .Range("B1").value = "EntryID"
        .Range("C1").value = "AccessoryName"
        .Range("D1").value = "SerialNumber"
        .Range("E1").value = "PhotoPath"
        .Range("F1").value = "SentDate"
        .Range("G1").value = "Returned"
        .Range("H1").value = "ReturnDate"
        .Range("I1").value = "Condition"
        
        ' Format Headers
        .Range("A1:I1").Font.Bold = True
        .Range("A1:I1").Interior.Color = RGB(255, 255, 200)
        .Range("A1:I1").HorizontalAlignment = xlCenter
        
        ' Set Column Widths
        .Columns("A").ColumnWidth = 12
        .Columns("B").ColumnWidth = 12
        .Columns("C").ColumnWidth = 20
        .Columns("D").ColumnWidth = 15
        .Columns("E").ColumnWidth = 40
        .Columns("F").ColumnWidth = 12
        .Columns("G").ColumnWidth = 10
        .Columns("H").ColumnWidth = 12
        .Columns("I").ColumnWidth = 15
    End With
    
    MsgBox "Assign_Accessories Sheet Created Successfully!", vbInformation
End Sub

Sub Create_Warranty_Return_Master_Sheet()
    Dim ws As Worksheet
    Dim sheetExists As Boolean
    
    ' Check if sheet already exists
    sheetExists = False
    For Each ws In ThisWorkbook.Worksheets
        If ws.name = "Warranty_Return_Master" Then
            sheetExists = True
            Exit For
        End If
    Next ws
    
    ' If sheet exists, delete it
    If sheetExists Then
        Application.DisplayAlerts = False
        ThisWorkbook.Worksheets("Warranty_Return_Master").Delete
        Application.DisplayAlerts = True
    End If
    
    ' Create new sheet
    Set ws = ThisWorkbook.Worksheets.Add
    ws.name = "Warranty_Return_Master"
    
    ' Add Headers (Row 1)
    With ws
        .Range("A1").value = "ReturnID"
        .Range("B1").value = "AssignID"
        .Range("C1").value = "EntryID"
        .Range("D1").value = "CustomerID"
        .Range("E1").value = "CustomerName"
        .Range("F1").value = "ProductType"
        .Range("G1").value = "Company"
        .Range("H1").value = "Model"
        .Range("I1").value = "OriginalSerial"
        .Range("J1").value = "VendorID"
        .Range("K1").value = "VendorName"
        .Range("L1").value = "VendorMobile"
        .Range("M1").value = "ReturnMode"
        .Range("N1").value = "CourierName"
        .Range("O1").value = "DocketNumber"
        .Range("P1").value = "ReturnDate"
        .Range("Q1").value = "ReturnCharges"
        .Range("R1").value = "ReturnStatus"
        .Range("S1").value = "NewSerial"
        .Range("T1").value = "RejectReason"
        .Range("U1").value = "ReturnPhotoPath"
        .Range("V1").value = "VendorInvoiceNumber"
        .Range("W1").value = "VendorInvoiceAmount"
        .Range("X1").value = "Remarks"
        .Range("Y1").value = "CreatedDate"
        
        ' Format Headers
        .Range("A1:Y1").Font.Bold = True
        .Range("A1:Y1").Interior.Color = RGB(255, 200, 200)
        .Range("A1:Y1").HorizontalAlignment = xlCenter
        
        ' Set Column Widths
        .Columns("A").ColumnWidth = 12
        .Columns("B").ColumnWidth = 12
        .Columns("C").ColumnWidth = 12
        .Columns("D").ColumnWidth = 12
        .Columns("E").ColumnWidth = 20
        .Columns("F").ColumnWidth = 15
        .Columns("G").ColumnWidth = 15
        .Columns("H").ColumnWidth = 15
        .Columns("I").ColumnWidth = 15
        .Columns("J").ColumnWidth = 12
        .Columns("K").ColumnWidth = 25
        .Columns("L").ColumnWidth = 15
        .Columns("M").ColumnWidth = 12
        .Columns("N").ColumnWidth = 15
        .Columns("O").ColumnWidth = 12
        .Columns("P").ColumnWidth = 12
        .Columns("Q").ColumnWidth = 12
        .Columns("R").ColumnWidth = 15
        .Columns("S").ColumnWidth = 15
        .Columns("T").ColumnWidth = 30
        .Columns("U").ColumnWidth = 40
        .Columns("V").ColumnWidth = 15
        .Columns("W").ColumnWidth = 15
        .Columns("X").ColumnWidth = 25
        .Columns("Y").ColumnWidth = 18
    End With
    
    MsgBox "Warranty_Return_Master Sheet Created Successfully!", vbInformation
End Sub

Sub Create_Service_WorkLog_Expense_Sheet()
    Dim ws As Worksheet
    Dim sheetExists As Boolean
    
    ' Check if sheet already exists
    sheetExists = False
    For Each ws In ThisWorkbook.Worksheets
        If ws.name = "Service_WorkLog_Expense" Then
            sheetExists = True
            Exit For
        End If
    Next ws
    
    ' If sheet exists, delete it
    If sheetExists Then
        Application.DisplayAlerts = False
        ThisWorkbook.Worksheets("Service_WorkLog_Expense").Delete
        Application.DisplayAlerts = True
    End If
    
    ' Create new sheet
    Set ws = ThisWorkbook.Worksheets.Add
    ws.name = "Service_WorkLog_Expense"
    
    ' Add Headers (Row 1)
    With ws
        .Range("A1").value = "LogID"
        .Range("B1").value = "AssignID"
        .Range("C1").value = "EntryID"
        .Range("D1").value = "CustomerID"
        .Range("E1").value = "ProductType"
        .Range("F1").value = "Company"
        .Range("G1").value = "Model"
        .Range("H1").value = "SerialNumber"
        .Range("I1").value = "EngineerID"
        .Range("J1").value = "EngineerName"
        .Range("K1").value = "WorkStartDate"
        .Range("L1").value = "WorkEndDate"
        .Range("M1").value = "WorkStatus"
        .Range("N1").value = "ProblemFound"
        .Range("O1").value = "ProblemSolution"
        .Range("P1").value = "PartsUsed"
        .Range("Q1").value = "LaborCharges"
        .Range("R1").value = "TransportCharges"
        .Range("S1").value = "OtherCharges"
        .Range("T1").value = "TotalExpense"
        .Range("U1").value = "CustomerApproval"
        .Range("V1").value = "ApprovalDate"
        .Range("W1").value = "ExpensePhotoPath"
        .Range("X1").value = "Remarks"
        .Range("Y1").value = "CreatedDate"
        
        ' Format Headers
        .Range("A1:Y1").Font.Bold = True
        .Range("A1:Y1").Interior.Color = RGB(255, 220, 180)
        .Range("A1:Y1").HorizontalAlignment = xlCenter
        
        ' Set Column Widths
        .Columns("A").ColumnWidth = 12
        .Columns("B").ColumnWidth = 12
        .Columns("C").ColumnWidth = 12
        .Columns("D").ColumnWidth = 12
        .Columns("E").ColumnWidth = 15
        .Columns("F").ColumnWidth = 15
        .Columns("G").ColumnWidth = 15
        .Columns("H").ColumnWidth = 15
        .Columns("I").ColumnWidth = 12
        .Columns("J").ColumnWidth = 20
        .Columns("K").ColumnWidth = 14
        .Columns("L").ColumnWidth = 14
        .Columns("M").ColumnWidth = 15
        .Columns("N").ColumnWidth = 30
        .Columns("O").ColumnWidth = 30
        .Columns("P").ColumnWidth = 40
        .Columns("Q").ColumnWidth = 12
        .Columns("R").ColumnWidth = 15
        .Columns("S").ColumnWidth = 12
        .Columns("T").ColumnWidth = 12
        .Columns("U").ColumnWidth = 15
        .Columns("V").ColumnWidth = 14
        .Columns("W").ColumnWidth = 40
        .Columns("X").ColumnWidth = 25
        .Columns("Y").ColumnWidth = 18
    End With
    
    MsgBox "Service_WorkLog_Expense Sheet Created Successfully!", vbInformation
End Sub

Sub Create_Delivery_Master_Sheet()
    Dim ws As Worksheet
    Dim sheetExists As Boolean
    
    ' Check if sheet already exists
    sheetExists = False
    For Each ws In ThisWorkbook.Worksheets
        If ws.name = "Delivery_Master" Then
            sheetExists = True
            Exit For
        End If
    Next ws
    
    ' If sheet exists, delete it
    If sheetExists Then
        Application.DisplayAlerts = False
        ThisWorkbook.Worksheets("Delivery_Master").Delete
        Application.DisplayAlerts = True
    End If
    
    ' Create new sheet
    Set ws = ThisWorkbook.Worksheets.Add
    ws.name = "Delivery_Master"
    
    ' Add Headers (Row 1)
    With ws
        .Range("A1").value = "DeliveryID"
        .Range("B1").value = "EntryID"
        .Range("C1").value = "CustomerID"
        .Range("D1").value = "CustomerName"
        .Range("E1").value = "CustomerMobile"
        .Range("F1").value = "ProductType"
        .Range("G1").value = "Company"
        .Range("H1").value = "Model"
        .Range("I1").value = "OriginalSerial"
        .Range("J1").value = "FinalSerial"
        .Range("K1").value = "DeliveryType"
        .Range("L1").value = "DeliveryDate"
        .Range("M1").value = "WorkDoneDescription"
        .Range("N1").value = "AccessoriesReturned"
        .Range("O1").value = "CourierChargesPaid"
        .Range("P1").value = "CourierAmount"
        .Range("Q1").value = "ServiceCharges"
        .Range("R1").value = "AmountPaid"
        .Range("S1").value = "PaymentMode"
        .Range("T1").value = "BalanceAmount"
        .Range("U1").value = "DeliveredBy"
        .Range("V1").value = "CustomerSignaturePath"
        .Range("W1").value = "DeliveryPhotoPath"
        .Range("X1").value = "Remarks"
        
        ' Format Headers
        .Range("A1:X1").Font.Bold = True
        .Range("A1:X1").Interior.Color = RGB(180, 255, 180)
        .Range("A1:X1").HorizontalAlignment = xlCenter
        
        ' Set Column Widths
        .Columns("A").ColumnWidth = 12
        .Columns("B").ColumnWidth = 12
        .Columns("C").ColumnWidth = 12
        .Columns("D").ColumnWidth = 20
        .Columns("E").ColumnWidth = 15
        .Columns("F").ColumnWidth = 15
        .Columns("G").ColumnWidth = 15
        .Columns("H").ColumnWidth = 15
        .Columns("I").ColumnWidth = 15
        .Columns("J").ColumnWidth = 15
        .Columns("K").ColumnWidth = 12
        .Columns("L").ColumnWidth = 12
        .Columns("M").ColumnWidth = 30
        .Columns("N").ColumnWidth = 25
        .Columns("O").ColumnWidth = 15
        .Columns("P").ColumnWidth = 12
        .Columns("Q").ColumnWidth = 12
        .Columns("R").ColumnWidth = 12
        .Columns("S").ColumnWidth = 12
        .Columns("T").ColumnWidth = 12
        .Columns("U").ColumnWidth = 15
        .Columns("V").ColumnWidth = 40
        .Columns("W").ColumnWidth = 40
        .Columns("X").ColumnWidth = 25
    End With
    
    MsgBox "Delivery_Master Sheet Created Successfully!", vbInformation
End Sub

Sub Update_Software_Config_CompanyDetails()
    Dim ws As Worksheet
    Dim sheetExists As Boolean
    
    ' Check if sheet exists
    sheetExists = False
    For Each ws In ThisWorkbook.Worksheets
        If ws.name = "Software_Config" Then
            sheetExists = True
            Exit For
        End If
    Next ws
    
    ' If sheet doesn't exist, create it
    If Not sheetExists Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.name = "Software_Config"
    Else
        Set ws = ThisWorkbook.Worksheets("Software_Config")
    End If
    
    ' Add/Update Company Details
    With ws
        ' Row 2-10: Company Information
        .Range("A2").value = "Company_Name"
        If .Range("B2").value = "" Then .Range("B2").value = "GLOBAL IT SOLUTIONS"
        
        .Range("A3").value = "Address_Line1"
        If .Range("B3").value = "" Then .Range("B3").value = "123, Main Road"
        
        .Range("A4").value = "Address_Line2"
        If .Range("B4").value = "" Then .Range("B4").value = ""
        
        .Range("A5").value = "City"
        If .Range("B5").value = "" Then .Range("B5").value = "Mumbai"
        
        .Range("A6").value = "PIN_Code"
        If .Range("B6").value = "" Then .Range("B6").value = "400001"
        
        .Range("A7").value = "State"
        If .Range("B7").value = "" Then .Range("B7").value = "Maharashtra"
        
        .Range("A8").value = "Country"
        If .Range("B8").value = "" Then .Range("B8").value = "India"
        
        .Range("A9").value = "Mobile_Number"
        If .Range("B9").value = "" Then .Range("B9").value = "9876543210"
        
        .Range("A10").value = "Email_ID"
        If .Range("B10").value = "" Then .Range("B10").value = "globalitsolutions@gmail.com"
        
        ' Format
        .Range("A2:A10").Font.Bold = True
        .Columns("A").ColumnWidth = 20
        .Columns("B").ColumnWidth = 35
    End With
    
    MsgBox "Software_Config Updated with Company Details!" & vbCrLf & _
           "Please verify/update the details in Column B", vbInformation
End Sub

Sub Create_ALL_Assignment_Sheets()
    ' Create all 6 sheets in one go
    Call Create_ServiceStation_Register_Sheet
    Call Create_Assign_Master_Sheet
    Call Create_Assign_Accessories_Sheet
    Call Create_Warranty_Return_Master_Sheet
    Call Create_Service_WorkLog_Expense_Sheet
    Call Create_Delivery_Master_Sheet
    Call Update_Software_Config_CompanyDetails
    
    MsgBox "All 6 Sheets Created Successfully!" & vbCrLf & vbCrLf & _
           "1. ServiceStation_Register" & vbCrLf & _
           "2. Assign_Master" & vbCrLf & _
           "3. Assign_Accessories" & vbCrLf & _
           "4. Warranty_Return_Master" & vbCrLf & _
           "5. Service_WorkLog_Expense" & vbCrLf & _
           "6. Delivery_Master" & vbCrLf & vbCrLf & _
           "Software_Config also updated!", vbInformation
End Sub

