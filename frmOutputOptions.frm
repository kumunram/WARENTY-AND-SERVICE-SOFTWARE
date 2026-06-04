VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmOutputOptions 
   Caption         =   "frmOutputOptions"
   ClientHeight    =   2805
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7560
   OleObjectBlob   =   "frmOutputOptions.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmOutputOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim eID As String
Dim prodRow As Long

#If VBA7 Then
    Private Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
        ByVal hwnd As LongPtr, ByVal lpOperation As String, ByVal lpFile As String, _
        ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As LongPtr
#Else
    Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
        ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
        ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
#End If

' PUBLIC VARIABLES
Public parentForm As Object
Private m_CustomerID As String
Private m_EntryID As String
Private m_CustomerName As String
Private m_Mobile As String
Private m_DataLoaded As Boolean

Private Sub btnCancelOutput_Click()
Unload Me
End Sub

'========================================
' EMAIL BUTTON - USE COMMUNICATION SETTINGS
'========================================
Private Sub btnSendMail_Click()
    On Error GoTo ErrorHandler
    
    ' ===== CHECK =====
    If Trim(m_CustomerID) = "" Then
        MsgBox "No Customer Selected!", vbExclamation
        Exit Sub
    End If
    
    ' ===== EMAIL SETTINGS FROM Communication Settings (B17-B22) =====
    Dim wsConfig As Worksheet
    Dim senderEmail As String, emailPassword As String, appPassword As String
    Dim smtpServer As String, smtpPort As String, sslEnable As String
    
    Set wsConfig = ThisWorkbook.Sheets("Software_Config")
    
    ' Communication Settings Tab ?? ?? (B17-B22) - ??? cells!
    senderEmail = Trim(wsConfig.Range("B17").value & "")    ' Email ID
    emailPassword = Trim(wsConfig.Range("B18").value & "")  ' Email Password (regular)
    smtpServer = Trim(wsConfig.Range("B19").value & "")     ' SMTP Server (Gmail/Outlook/etc)
    appPassword = Trim(wsConfig.Range("B20").value & "")    ' APP PASSWORD (important!)
    smtpPort = Trim(wsConfig.Range("B21").value & "")       ' Port (587/465)
    sslEnable = UCase(Trim(wsConfig.Range("B22").value & "")) ' SSL (Yes/No)
    
    ' Check if configured
    If senderEmail = "" Or smtpServer = "" Then
        MsgBox "Email settings not configured!" & vbCrLf & _
               "Please set Email in Communication Settings first.", vbExclamation
        Exit Sub
    End If
    
    ' Use App Password if available, else use regular password
    Dim usePassword As String
    If appPassword <> "" Then
        usePassword = appPassword    ' App Password??
    Else
        usePassword = emailPassword  ' Regular password
    End If
    
    If usePassword = "" Then
        MsgBox "Please enter Email Password or App Password in Communication Settings!", vbExclamation
        Exit Sub
    End If
    
    ' Default port if empty
    If smtpPort = "" Then smtpPort = "587"
    
    ' ===== CUSTOMER EMAIL =====
    Dim wsCust As Worksheet
    Dim custEmail As String
    Dim lastRow As Long, i As Long
    
    Set wsCust = ThisWorkbook.Sheets("Customer_Master")
    lastRow = wsCust.Cells(wsCust.Rows.count, 1).End(xlUp).row
    
    For i = 2 To lastRow
        If UCase(Trim(wsCust.Cells(i, 1).value & "")) = UCase(m_CustomerID) Then
            custEmail = Trim(wsCust.Cells(i, 5).value & "")
            Exit For
        End If
    Next i
    
    If custEmail = "" Then
        MsgBox "Customer email not found!", vbExclamation
        Exit Sub
    End If
    
    ' ===== GENERATE FULL RECEIPT HTML =====
    Dim htmlReceipt As String
    htmlReceipt = GenerateReceiptHTML()
    
    If htmlReceipt = "" Then
        MsgBox "Failed to generate receipt!", vbCritical
        Exit Sub
    End If
    
    ' ===== SEND EMAIL VIA CDO =====
    ' ===== SEND EMAIL VIA CDO =====
Dim cdoMsg As Object
Dim cdoConf As Object
    
' SMTP Configuration
Set cdoMsg = CreateObject("CDO.Message")
Set cdoConf = CreateObject("CDO.Configuration")

With cdoConf.Fields
    .item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
    .item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = GetSMTPServerAddress(smtpServer)
    .item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = CInt(smtpPort)
    .item("http://schemas.microsoft.com/cdo/configuration/sendusername") = senderEmail
    .item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = usePassword
    .item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
    
    If sslEnable = "YES" Then
    .item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
Else
    .item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = False
End If

.item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
    .Update
End With
    
    ' Email Setup
    Dim compName As String
    compName = Trim(wsConfig.Range("B2").value & "")
    If compName = "" Then compName = "GLOBAL IT SOLUTIONS"
    
    With cdoMsg
        Set .Configuration = cdoConf
        .From = senderEmail
        .To = custEmail
        .subject = "Service Receipt - " & m_CustomerName & " (" & m_CustomerID & ")"
        .htmlBody = htmlReceipt
        .Send
    End With
    
    ' Cleanup
    Set cdoMsg = Nothing
    Set cdoConf = Nothing
    
    MsgBox "Email sent successfully!" & vbCrLf & _
           "To: " & custEmail, vbInformation
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Email Error: " & Err.Description & vbCrLf & vbCrLf & _
           "Check:" & vbCrLf & _
           "1. App Password ??? ???" & vbCrLf & _
           "2. Communication Settings ??? Gmail + App Password set ???" & vbCrLf & _
           "3. Internet connected ???", vbCritical
End Sub

'========================================
' GET SMTP SERVER ADDRESS FROM NAME
'========================================
Private Function GetSMTPServerAddress(ByVal serverName As String) As String
    Select Case Trim(LCase(serverName))
        Case "gmail": GetSMTPServerAddress = "smtp.gmail.com"
        Case "outlook": GetSMTPServerAddress = "smtp.office365.com"
        Case "yahoo": GetSMTPServerAddress = "smtp.mail.yahoo.com"
        Case "hotmail": GetSMTPServerAddress = "smtp.live.com"
        Case Else: GetSMTPServerAddress = serverName  ' Custom
    End Select
End Function
'========================================
' GENERATE RECEIPT AS HTML FOR EMAIL
'========================================
Private Function GenerateReceiptHTML() As String
    On Error GoTo ErrorHandler
    
    Dim wsCust As Worksheet, wsProd As Worksheet, wsAcc As Worksheet
    Dim wsPay As Worksheet, wsJob As Worksheet, wsConfig As Worksheet
    Dim html As String, i As Long, entryNum As Long
    Dim compName As String, addr1 As String, addr2 As String
    Dim compMob As String, compEmail As String
    Dim entryRows As Collection, prodRow As Long, eID As String
    Dim custAddress As String, custGST As String, custEmail As String
    Dim totalEntries As Long
    
    ' Set worksheets
    Set wsCust = ThisWorkbook.Sheets("Customer_Master")
    Set wsProd = ThisWorkbook.Sheets("Job_Product")
    Set wsAcc = ThisWorkbook.Sheets("Job_Accessory")
    Set wsPay = ThisWorkbook.Sheets("Payment_Master")
    Set wsJob = ThisWorkbook.Sheets("Job_Master")
    Set wsConfig = ThisWorkbook.Sheets("Software_Config")
    
    ' Company Info from Software_Config
    compName = Trim(wsConfig.Range("B2").value & "")
    addr1 = Trim(wsConfig.Range("B3").value & "")
    addr2 = Trim(wsConfig.Range("B4").value & "")
    compMob = Trim(wsConfig.Range("B9").value & "")
    compEmail = Trim(wsConfig.Range("B10").value & "")
    
    If compName = "" Then compName = "GLOBAL IT SOLUTIONS"
    
    ' Customer Details
    custAddress = ""
    custGST = ""
    custEmail = ""
    For i = 2 To wsCust.Cells(wsCust.Rows.count, 1).End(xlUp).row
        If UCase(Trim(wsCust.Cells(i, 1).value & "")) = UCase(m_CustomerID) Then
            custAddress = wsCust.Cells(i, 4).value & ""
            custEmail = wsCust.Cells(i, 5).value & ""
            custGST = wsCust.Cells(i, 6).value & ""
            Exit For
        End If
    Next i
    
    ' Get all entries for this customer
    Set entryRows = New Collection
    For i = 2 To wsProd.Cells(wsProd.Rows.count, 1).End(xlUp).row
        If Trim(UCase(wsProd.Cells(i, 2).value & "")) = Trim(UCase(m_CustomerID)) Then
            entryRows.Add i
        End If
    Next i
    totalEntries = entryRows.count
    
    ' ===== BUILD HTML =====
    html = "<!DOCTYPE html>"
    html = html & "<html><head>"
    html = html & "<style>"
    html = html & "body { font-family: Arial, sans-serif; margin: 20px; background: #f5f5f5; }"
    html = html & ".receipt { max-width: 800px; margin: 0 auto; background: white; border: 2px solid #333; }"
    html = html & ".header { background: #006400; color: white; text-align: center; padding: 15px; }"
    html = html & ".header h1 { margin: 0; font-size: 24px; }"
    html = html & ".company-info { text-align: center; padding: 10px; border-bottom: 1px solid #ccc; }"
    html = html & ".customer-header { background: #000080; color: white; padding: 10px; font-weight: bold; }"
    html = html & ".section { padding: 10px; border-bottom: 1px solid #ddd; }"
    html = html & ".entry-header { background: #c80000; color: white; padding: 8px; font-weight: bold; margin-top: 10px; }"
    html = html & ".entry-header.service { background: #ffa500; }"
    html = html & "table { width: 100%; border-collapse: collapse; margin: 10px 0; }"
    html = html & "th, td { padding: 8px; text-align: left; border: 1px solid #ddd; }"
    html = html & "th { background: #f0f0f0; font-weight: bold; }"
    html = html & ".payment { background: #006400; color: white; padding: 8px; }"
    html = html & ".terms { background: #800000; color: white; padding: 10px; font-weight: bold; }"
    html = html & ".footer { text-align: center; padding: 15px; background: #ffffcd; color: #006400; font-weight: bold; }"
    html = html & ".label { font-weight: bold; color: #333; }"
    html = html & "</style></head><body>"
    
    ' Receipt Container
    html = html & "<div class='receipt'>"
    
    ' Company Header
    html = html & "<div class='header'>"
    html = html & "<h1>" & compName & "</h1>"
    html = html & "</div>"
    
    ' Company Info
    html = html & "<div class='company-info'>"
    If addr1 <> "" Then html = html & "<p>" & addr1 & "</p>"
    If addr2 <> "" Then html = html & "<p>" & addr2 & "</p>"
    html = html & "<p><strong>Mobile: " & compMob & " | Email: " & compEmail & "</strong></p>"
    html = html & "</div>"
    
    ' Title
    html = html & "<div class='header'>"
    html = html & "<h2>WARRANTY AND SERVICE REPORT</h2>"
    html = html & "</div>"
    
    ' Customer Header
    html = html & "<div class='customer-header'>"
    html = html & "CUSTOMER ID: " & m_CustomerID
    html = html & "</div>"
    
    ' Customer Details
    html = html & "<div class='section'>"
    html = html & "<table>"
    html = html & "<tr><td class='label'>Name:</td><td>" & m_CustomerName & "</td></tr>"
    html = html & "<tr><td class='label'>Mobile:</td><td>" & m_Mobile & "</td></tr>"
    If custEmail <> "" Then html = html & "<tr><td class='label'>Email:</td><td>" & custEmail & "</td></tr>"
    If custAddress <> "" Then html = html & "<tr><td class='label'>Address:</td><td>" & Replace(custAddress, vbCrLf, ", ") & "</td></tr>"
    If custGST <> "" Then html = html & "<tr><td class='label'>GST:</td><td>" & custGST & "</td></tr>"
    html = html & "</table>"
    html = html & "</div>"
    
    ' Entries Loop
    For entryNum = 1 To totalEntries
        prodRow = entryRows(entryNum)
        eID = wsProd.Cells(prodRow, 1).value & ""
        
        ' Entry Type & Color
        Dim entryType As String, entryClass As String
        If UCase(Left(eID, 3)) = "WAR" Then
            entryType = "WARRANTY"
            entryClass = "entry-header"
        ElseIf UCase(Left(eID, 3)) = "SER" Then
            entryType = "SERVICE"
            entryClass = "entry-header service"
        Else
            entryType = "ENTRY"
            entryClass = "entry-header"
        End If
        
        ' Entry Date from Job_Master
        Dim entryDate As String, verifyType As String, verifyName As String
        entryDate = ""
        verifyType = ""
        verifyName = ""
        For i = 2 To wsJob.Cells(wsJob.Rows.count, 1).End(xlUp).row
            If Trim(UCase(wsJob.Cells(i, 1).value & "")) = Trim(UCase(eID)) Then
                entryDate = wsJob.Cells(i, 8).value & ""
                verifyType = wsJob.Cells(i, 6).value & ""
                verifyName = wsJob.Cells(i, 7).value & ""
                Exit For
            End If
        Next i
        If entryDate = "" Then entryDate = Format(Date, "dd-MMM-yyyy")
        
        ' Entry Header
        html = html & "<div class='" & entryClass & "'>"
        html = html & entryType & " ID: " & eID & " | Date: " & entryDate
        html = html & "</div>"
        
        ' Product Details
        html = html & "<div class='section'>"
        html = html & "<table>"
        html = html & "<tr><td class='label'>Product:</td><td>" & wsProd.Cells(prodRow, 6).value & "</td>"
        html = html & "<td class='label'>Company:</td><td>" & wsProd.Cells(prodRow, 7).value & "</td></tr>"
        html = html & "<tr><td class='label'>Model:</td><td>" & wsProd.Cells(prodRow, 8).value & "</td>"
        html = html & "<td class='label'>Serial:</td><td>" & wsProd.Cells(prodRow, 9).value & "</td></tr>"
        
        ' Problem
        Dim prob As String
        prob = wsProd.Cells(prodRow, 13).value & ""
        If prob <> "" Then
            html = html & "<tr><td class='label'>Problem:</td><td colspan='3' style='color: #c80000;'>" & prob & "</td></tr>"
        End If
        html = html & "</table>"
        
        ' Accessories
        Dim accRow As Long, lastRowAcc As Long, accCount As Integer
        lastRowAcc = wsAcc.Cells(wsAcc.Rows.count, 1).End(xlUp).row
        accCount = 0
        
        For accRow = 2 To lastRowAcc
            If Trim(UCase(wsAcc.Cells(accRow, 1).value & "")) = Trim(UCase(eID)) Then
                accCount = accCount + 1
            End If
        Next accRow
        
        If accCount > 0 Then
            html = html & "<p class='label'>ACCESSORIES:</p>"
            html = html & "<table>"
            html = html & "<tr><th>S.No</th><th>Accessory Name</th><th>Brand</th><th>Serial</th><th>Received</th></tr>"
            
            Dim accIdx As Integer
accIdx = 1
For accRow = 2 To lastRowAcc
    If Trim(UCase(wsAcc.Cells(accRow, 1).value & "")) = Trim(UCase(eID)) Then
        html = html & "<tr>"
        html = html & "<td>" & accIdx & "</td>"
                    html = html & "<td>" & wsAcc.Cells(accRow, 4).value & "</td>"
                    html = html & "<td>" & wsAcc.Cells(accRow, 5).value & "</td>"
                    html = html & "<td>" & wsAcc.Cells(accRow, 6).value & "</td>"
                    html = html & "<td>" & wsAcc.Cells(accRow, 7).value & "</td>"
                    html = html & "</tr>"
                    accIdx = accIdx + 1
        If accIdx > 4 Then Exit For  ' idx ?? ??? accIdx
    End If
Next accRow
            html = html & "</table>"
        End If
        
        ' Payment
        Dim lastRowPay As Long, payRow As Long
        Dim totalCharges As Double, totalExp As Double, totalRec As Double
        Dim dueAmt As Double, payID As String
        Dim foundPayment As Boolean
        
        lastRowPay = wsPay.Cells(wsPay.Rows.count, 1).End(xlUp).row
        foundPayment = False
        totalCharges = 0: totalExp = 0: totalRec = 0
        
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
            html = html & "<div class='payment'>"
            html = html & "Payment Details | Entry ID: " & eID & " | Pay ID: " & payID
            html = html & "</div>"
            html = html & "<table>"
            html = html & "<tr>"
            html = html & "<td>Service Charge: <strong>Rs." & Format(totalCharges, "0.00") & "</strong></td>"
            If totalExp > 0 Then html = html & "<td>Expense: <strong>Rs." & Format(totalExp, "0.00") & "</strong></td>"
            html = html & "<td>Received: <strong>Rs." & Format(totalRec, "0.00") & "</strong></td>"
            html = html & "<td>Due: <strong>Rs." & Format(dueAmt, "0.00") & "</strong></td>"
            html = html & "</tr>"
            html = html & "</table>"
        End If
        
        ' Verification
        If verifyType <> "" Or verifyName <> "" Then
            html = html & "<p style='background: #f0f8ff; padding: 8px; border: 2px solid #000080; text-align: center;'>"
            html = html & "<strong>Verified Type:</strong> " & verifyType & " | <strong>Verified By:</strong> " & verifyName
            html = html & "</p>"
        End If
        
        html = html & "</div>"
    Next entryNum
    
    ' Terms & Conditions
    html = html & "<div class='terms'>TERMS & CONDITIONS</div>"
    html = html & "<div class='section'>"
    html = html & "<ol>"
    html = html & "<li>WARRANTY: Coverage as per manufacturer's policy only.</li>"
    html = html & "<li>TIMELINE: Repair confirmation within 30 days.</li>"
    html = html & "<li>RECEIPT: Original receipt mandatory for collection.</li>"
    html = html & "<li>DATA: Not responsible for data loss during repair.</li>"
    html = html & "<li>DAMAGE: Not liable for pre-existing damage.</li>"
    html = html & "<li>REPLACEMENT: No refund if condition remains same.</li>"
    html = html & "<li>VOID: Warranty void if seal removed/tampered.</li>"
    html = html & "<li>LEGAL: All disputes subject to local jurisdiction.</li>"
    html = html & "</ol>"
    html = html & "</div>"
    
    ' Footer
    html = html & "<div class='footer'>"
    html = html & "*** Thank You for choosing " & compName & "! ***"
    html = html & "</div>"
    
    ' Close containers
    html = html & "</div>"  ' receipt
    html = html & "</body></html>"
    
    GenerateReceiptHTML = html
    Exit Function
    
ErrorHandler:
    GenerateReceiptHTML = ""
End Function

'========================================
' USERFORM INITIALIZE
'========================================
Private Sub UserForm_Initialize()
    On Error Resume Next
    Dim wsConfig As Worksheet
    Set wsConfig = ThisWorkbook.Sheets("Software_Config")
    If Not wsConfig Is Nothing Then
        Me.caption = "Output Options - " & wsConfig.Range("B2").value
    End If
    SetupButtonStyles
    
    m_CustomerID = ""
    m_EntryID = ""
    m_CustomerName = ""
    m_Mobile = ""
    m_DataLoaded = False
End Sub

'========================================
' USERFORM ACTIVATE
'========================================
Private Sub UserForm_Activate()
    On Error Resume Next
    
    If m_DataLoaded Then
        Debug.Print "Data already loaded, skipping Activate load"
        Exit Sub
    End If
    
    If m_CustomerID = "" And Not parentForm Is Nothing Then
        Dim custID As String
        
        On Error Resume Next
        custID = parentForm.txtCustomerID.caption
        If custID = "" Then custID = parentForm.txtCustomerID.text
        If custID = "" Then custID = parentForm.lblCustomerID.caption
        On Error GoTo 0
        
        If custID <> "" Then
            Debug.Print "Loading from parentForm: " & custID
            LoadData custID
        Else
            MsgBox "Could not get Customer ID from parent form!", vbExclamation
        End If
    End If
End Sub

'========================================
' LOAD DATA
'========================================
Public Sub LoadData(ByVal paramID As String, Optional ByVal custID As String = "")
    On Error Resume Next
    
    Debug.Print "LoadData called with paramID: '" & paramID & "', custID: '" & custID & "'"
    
    Dim cleanID As String
    cleanID = Trim(paramID)
    
    If cleanID = "" Then
        Debug.Print "Empty paramID received!"
        Exit Sub
    End If
    
    m_CustomerID = ""
    m_EntryID = ""
    m_CustomerName = ""
    m_Mobile = ""
    
    Dim isEntryID As Boolean
    isEntryID = (Left(UCase(cleanID), 3) = "WAR" Or Left(UCase(cleanID), 3) = "SER")
    
    If isEntryID Then
        m_EntryID = cleanID
        If Trim(custID) <> "" Then
            m_CustomerID = Trim(custID)
            Debug.Print "EntryID with direct CustomerID: " & m_CustomerID
        Else
            m_CustomerID = GetCustomerIDFromEntryID(m_EntryID)
            Debug.Print "EntryID detected. Converted to CustomerID: '" & m_CustomerID & "'"
        End If
    Else
        m_CustomerID = cleanID
        m_EntryID = ""
        Debug.Print "CustomerID detected: '" & m_CustomerID & "'"
    End If
    
    If m_CustomerID <> "" Then
        LoadCustomerDetailsFromMaster
        m_DataLoaded = True
        Debug.Print "Loaded Customer: " & m_CustomerName & ", Mobile: " & m_Mobile
    Else
        Debug.Print "ERROR: CustomerID is empty after conversion!"
    End If
End Sub

'========================================
' GET CUSTOMERID FROM ENTRYID
'========================================
Private Function GetCustomerIDFromEntryID(eID As String) As String
    Dim wsProd As Worksheet
    Dim lastRow As Long, i As Long

    On Error Resume Next
    Set wsProd = ThisWorkbook.Sheets("Job_Product")
    On Error GoTo 0

    If wsProd Is Nothing Then
        GetCustomerIDFromEntryID = ""
        Exit Function
    End If

    lastRow = wsProd.Cells(wsProd.Rows.count, 1).End(xlUp).row

    For i = 2 To lastRow
        If Trim(UCase(wsProd.Cells(i, 1).value & "")) = Trim(UCase(eID)) Then
            GetCustomerIDFromEntryID = Trim(wsProd.Cells(i, 2).value & "")
            Exit Function
        End If
    Next i

    GetCustomerIDFromEntryID = ""
End Function

'========================================
' LOAD CUSTOMER DETAILS FROM MASTER
'========================================
Private Sub LoadCustomerDetailsFromMaster()
    Dim wsCust As Worksheet
    Dim lastRow As Long, i As Long
    
    On Error Resume Next
    Set wsCust = ThisWorkbook.Sheets("Customer_Master")
    If wsCust Is Nothing Then
        Debug.Print "Customer_Master sheet not found!"
        Exit Sub
    End If
    
    lastRow = wsCust.Cells(wsCust.Rows.count, 1).End(xlUp).row
    
    For i = 2 To lastRow
        If UCase(Trim(wsCust.Cells(i, 1).value & "")) = UCase(m_CustomerID) Then
            m_Mobile = wsCust.Cells(i, 2).value & ""
            m_CustomerName = wsCust.Cells(i, 3).value & ""
            Debug.Print "Found in Customer_Master: " & m_CustomerName
            Exit For
        End If
    Next i
End Sub

'========================================
' PRINT BUTTON
'========================================
Private Sub btnPrintReceipt_Click()
    On Error GoTo ErrorHandler
    
    If Trim(m_CustomerID) = "" Then
        MsgBox "Customer ID not found! Please check data loading.", vbExclamation
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    Dim wsReceipt As Worksheet
    Set wsReceipt = GenerateReceiptSheet
    
    If wsReceipt Is Nothing Then
        MsgBox "Receipt generation failed!", vbCritical
        GoTo Cleanup
    End If
    
    wsReceipt.PrintOut Copies:=1, Preview:=False
    
Cleanup:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Exit Sub
    
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
    Resume Cleanup
End Sub
'========================================
' PDF BUTTON - ??????? ?????
'========================================
Private Sub btnExportPDF_Click()
    On Error GoTo ErrorHandler
    
    If Trim(m_CustomerID) = "" Then
        MsgBox "No Customer Selected!", vbExclamation
        Exit Sub
    End If
    
    ' ===== Configuration Folder ????? (Images ?? ????) =====
    Dim configFolder As String
    configFolder = ThisWorkbook.path & "\Configuration\"
    
    ' ??? GLOBAL_SOFT_DATA ?? ???? ????? ?? ?? ???? ??? ????:
    ' configFolder = ThisWorkbook.path & "\GLOBAL_SOFT_DATA\Configuration\"
    
    ' Folder ???? ?? ?? ??? ??
    If Dir(configFolder, vbDirectory) = "" Then
        MkDir configFolder
    End If
    
    ' ===== File Name ?????: EntryID_Username_DateTime.pdf =====
    Dim entryIDForFile As String
    Dim safeCustomerName As String
    Dim fileName As String
    Dim pdfPath As String
    
    ' Entry ID ??? (??? Entry ID ?? ???? ?? ?? ??, ???? Customer ID)
    If Trim(m_EntryID) <> "" Then
        entryIDForFile = m_EntryID
    Else
        entryIDForFile = m_CustomerID
    End If
    
    ' Customer Name ??? ?? Invalid Characters ????
    ' Special Characters ????
safeCustomerName = m_CustomerName
safeCustomerName = Replace(safeCustomerName, "\", "_")
safeCustomerName = Replace(safeCustomerName, "/", "_")
safeCustomerName = Replace(safeCustomerName, ":", "_")
safeCustomerName = Replace(safeCustomerName, "*", "_")
safeCustomerName = Replace(safeCustomerName, "?", "_")
safeCustomerName = Replace(safeCustomerName, """", "_")
safeCustomerName = Replace(safeCustomerName, "<", "_")
safeCustomerName = Replace(safeCustomerName, ">", "_")
safeCustomerName = Replace(safeCustomerName, "|", "_")
    If safeCustomerName = "" Then safeCustomerName = "Customer"
    
    ' Format: WAR001_Rajesh_08042026_143022.pdf
    fileName = entryIDForFile & "_" & safeCustomerName & "_" & Format(Now, "ddmmyyyy_hhmmss") & ".pdf"
    pdfPath = configFolder & fileName
    
    ' ===== Receipt Generate ???? =====
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    Dim wsReceipt As Worksheet
    Set wsReceipt = GenerateReceiptSheet
    
    If wsReceipt Is Nothing Then
        MsgBox "Failed to create receipt!", vbCritical
        GoTo Cleanup
    End If
    
    ' ===== PDF Export ???? =====
    wsReceipt.ExportAsFixedFormat Type:=xlTypePDF, fileName:=pdfPath, Quality:=xlQualityStandard
    
    ' ===== Temp Sheet Delete ???? =====
    Application.DisplayAlerts = False
    wsReceipt.Delete
    Application.DisplayAlerts = True
    
    ' ===== PDF Open ???? (Adobe Reader/Browser ???) =====
    ShellExecute 0, "open", pdfPath, vbNullString, vbNullString, 1
    
    MsgBox "PDF Opened!" & vbCrLf & "Auto-Saved at: " & pdfPath, vbInformation
    
Cleanup:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Exit Sub
    
ErrorHandler:
    MsgBox "PDF Error: " & Err.Description, vbCritical
    Resume Cleanup
End Sub
' ?? Function File Name ??? ???? ?? ??? ??
Private Function CleanFileName(ByVal strName As String) As String
    Dim invalidChars As String
    Dim i As Integer
    invalidChars = "\/:*?""<>|"
    For i = 1 To Len(invalidChars)
        strName = Replace(strName, Mid(invalidChars, i, 1), "_")
    Next i
    CleanFileName = Trim(strName)
End Function
Private Function GenerateReceiptSheet() As Worksheet
    On Error GoTo ErrorHandler
    
    Dim wsCust As Worksheet, wsProd As Worksheet, wsAcc As Worksheet, wsPay As Worksheet, wsJob As Worksheet
    Dim wsSettings As Worksheet, wsSet As Worksheet, ws As Worksheet, wsConfig As Worksheet
    Dim r As Long, i As Long, termRow As Long, t As Integer
    Dim compAddr As String, compMob As String, compName As String, logoPath As String, compEmail As String
    Dim custEmail As String, custGST As String, pProblem As String
    Dim eID As String, payID As String, entryDate As String
    Dim entryRows As Collection
    Dim lastRowProd As Long, prodRow As Long, accRow As Long, lastRowAcc As Long, accCount As Long
    Dim lastRowPay As Long, payRow As Long, jobRow As Long
    Dim foundPayment As Boolean, termsFound As Boolean
    Dim totalCharges As Double, totalExp As Double, totalRec As Double, dueAmt As Double
    Dim verifyType As String, verifyName As String
    Dim pageSize As String, terms(1 To 8) As String
    Dim shpLogo As Shape, shpPhoto As Shape
    Dim showCustPhoto As String, showProdPhoto As String, showAccPhoto As String
    Dim custPhotoPath As String, prodPhotoPath As String, accPhotoPath As String
    
    ' PAGINATION VARIABLES
    Dim totalEntries As Long, entryNum As Long
    Dim entriesOnPage As Long
    Dim totalPages As Long, currentPage As Long
    Const MAX_ENTRIES = 2
    
    ' Delete old sheet
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    On Error Resume Next
    ThisWorkbook.Sheets("Temp_Receipt").Delete
    On Error GoTo 0
    
    ' Create new sheet
    Set ws = ThisWorkbook.Sheets.Add
    ws.name = "Temp_Receipt"
    
    ' Page Setup
    On Error Resume Next
    Set wsSettings = ThisWorkbook.Sheets("Settings")
    pageSize = UCase(wsSettings.Range("B20").value & "")
    showCustPhoto = UCase(wsSettings.Range("B21").value & "")
    showProdPhoto = UCase(wsSettings.Range("B22").value & "")
    showAccPhoto = UCase(wsSettings.Range("B23").value & "")
    custPhotoPath = wsSettings.Range("B3").value & ""
    prodPhotoPath = wsSettings.Range("B2").value & ""
    accPhotoPath = wsSettings.Range("B6").value & ""
    logoPath = wsSettings.Range("B11").value & ""
    On Error GoTo 0
    
    ' Setup paths
    If custPhotoPath <> "" And Right(custPhotoPath, 1) <> "\" Then custPhotoPath = custPhotoPath & "\"
    If prodPhotoPath <> "" And Right(prodPhotoPath, 1) <> "\" Then prodPhotoPath = prodPhotoPath & "\"
    If accPhotoPath <> "" And Right(accPhotoPath, 1) <> "\" Then accPhotoPath = accPhotoPath & "\"
    If prodPhotoPath = "" Then prodPhotoPath = ThisWorkbook.path & "\GLOBAL_SOFT_DATA\Images\Products\"
    If accPhotoPath = "" Then accPhotoPath = ThisWorkbook.path & "\GLOBAL_SOFT_DATA\Images\Accessories\"
    If logoPath = "" Then logoPath = ThisWorkbook.path & "\GLOBAL_SOFT_DATA\Images\Logo\"
    If Right(logoPath, 1) <> "\" Then logoPath = logoPath & "\"
    
    ' Check logo file
    Dim finalLogoPath As String
    finalLogoPath = logoPath & "LOGO.png"
    If Dir(finalLogoPath) = "" Then finalLogoPath = logoPath & "LOGO.jpg"
    If Dir(finalLogoPath) = "" Then finalLogoPath = logoPath & "logo.png"
    If Dir(finalLogoPath) = "" Then finalLogoPath = logoPath & "logo.jpg"
    
    If pageSize = "" Then pageSize = "A4"
    
    ' Calculate total pages
    Set entryRows = New Collection
    Set wsProd = ThisWorkbook.Sheets("Job_Product")
    lastRowProd = wsProd.Cells(wsProd.Rows.count, 1).End(xlUp).row
    For i = 2 To lastRowProd
        If Trim(UCase(wsProd.Cells(i, 2).value & "")) = Trim(UCase(m_CustomerID)) Then
            entryRows.Add i
        End If
    Next i
    
    totalEntries = entryRows.count
    totalPages = ((totalEntries + MAX_ENTRIES - 1) \ MAX_ENTRIES)
    If totalPages < 1 Then totalPages = 1
    
    ' Page Setup
    With ws.PageSetup
        If pageSize = "A5" Then
            .PaperSize = xlPaperA5
            .LeftMargin = Application.InchesToPoints(0.25)
            .RightMargin = Application.InchesToPoints(0.25)
            .TopMargin = Application.InchesToPoints(0.25)
            .BottomMargin = Application.InchesToPoints(0.25)
        Else
            .PaperSize = xlPaperA4
            .LeftMargin = Application.InchesToPoints(0.5)
            .RightMargin = Application.InchesToPoints(0.5)
            .TopMargin = Application.InchesToPoints(0.5)
            .BottomMargin = Application.InchesToPoints(0.5)
        End If
        .Orientation = xlPortrait
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
    End With
    
    ' Column widths
    If pageSize = "A5" Then
        ws.Columns("A").ColumnWidth = 2
        ws.Columns("B").ColumnWidth = 8
        ws.Columns("C").ColumnWidth = 20
        ws.Columns("D").ColumnWidth = 15
        ws.Columns("E").ColumnWidth = 15
        ws.Columns("F").ColumnWidth = 15
        ws.Columns("G").ColumnWidth = 15
        ws.Columns("H").ColumnWidth = 2
    Else
        ws.Columns("A").ColumnWidth = 2
        ws.Columns("B").ColumnWidth = 10
        ws.Columns("C").ColumnWidth = 22
        ws.Columns("D").ColumnWidth = 18
        ws.Columns("E").ColumnWidth = 18
        ws.Columns("F").ColumnWidth = 18
        ws.Columns("G").ColumnWidth = 18
        ws.Columns("H").ColumnWidth = 2
    End If
    
    ' Set worksheets
    Set wsCust = ThisWorkbook.Sheets("Customer_Master")
    Set wsAcc = ThisWorkbook.Sheets("Job_Accessory")
    Set wsPay = ThisWorkbook.Sheets("Payment_Master")
    Set wsJob = ThisWorkbook.Sheets("Job_Master")
    Set wsConfig = ThisWorkbook.Sheets("Software_Config")
    
    ' Load Terms & Conditions
    On Error Resume Next
    Set wsSet = ThisWorkbook.Sheets("Settings")
    termsFound = False
    If Not wsSet Is Nothing Then
        For termRow = 25 To 32
            If Trim(wsSet.Cells(termRow, 2).value & "") <> "" Then
                termsFound = True
                Exit For
            End If
        Next termRow
    End If
    On Error GoTo 0
    
    If Not termsFound Then
        terms(1) = "1. WARRANTY: Coverage as per manufacturer's policy only."
        terms(2) = "2. TIMELINE: Repair confirmation within 30 days."
        terms(3) = "3. RECEIPT: Original receipt mandatory for collection."
        terms(4) = "4. DATA: Not responsible for data loss during repair."
        terms(5) = "5. DAMAGE: Not liable for pre-existing damage."
        terms(6) = "6. REPLACEMENT: No refund if condition remains same."
        terms(7) = "7. VOID: Warranty void if seal removed/tampered."
        terms(8) = "8. LEGAL: All disputes subject to Nayagarh jurisdiction."
    End If
    
    ' ========== READ COMPANY DATA FROM Software_Config ==========
    compName = Trim(wsConfig.Range("B2").value & "")
    Dim addr1 As String, addr2 As String
    addr1 = Trim(wsConfig.Range("B3").value & "")
    addr2 = Trim(wsConfig.Range("B4").value & "")
    compMob = Trim(wsConfig.Range("B9").value & "")
    compEmail = Trim(wsConfig.Range("B10").value & "")
    
    ' Main loop
    r = 1
    currentPage = 0
    entriesOnPage = 0
    
    For entryNum = 1 To totalEntries
        
        ' NEW PAGE CHECK
                If entryNum > 1 And entriesOnPage >= MAX_ENTRIES Then
            Call AddPageFooter(ws, r, termsFound, wsSet, terms, verifyType, verifyName)
            r = r + 1
            ws.Rows(r).PageBreak = xlPageBreakManual
            entriesOnPage = 0
            r = r + 1
        End If
        
        currentPage = currentPage + 1
        
        ' PAGE HEADER
        If entriesOnPage = 0 Then
            
            ' ========== ROW 1: COMPANY NAME ==========
            With ws.Range("B" & r & ":G" & r)
                .Merge
                .value = compName
                .Font.Size = 26
                .Font.Bold = True
                .Font.name = "Arial"
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
            End With
            ws.Rows(r).RowHeight = 40
            
            ' LOGO - LEFT SIDE
            If Dir(finalLogoPath) <> "" Then
                On Error Resume Next
                Set shpLogo = ws.Shapes.AddPicture(finalLogoPath, msoFalse, msoTrue, 0, 0, -1, -1)
                If Not shpLogo Is Nothing Then
                    With shpLogo
                        .LockAspectRatio = msoTrue
                        .Height = 100
                        .Top = ws.Range("B" & r).Top + 5
                        .Left = ws.Range("A" & r).Left + 5
                        If .Width > 130 Then .Width = 130
                    End With
                End If
                On Error GoTo ErrorHandler
            End If
            r = r + 1

            ' ========== ROW 2: ADDRESS LINE 1 ==========
            With ws.Range("B" & r & ":G" & r)
                .Merge
                .value = addr1
                .Font.Size = 13
                .Font.name = "Arial"
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
            End With
            ws.Rows(r).RowHeight = 22
            r = r + 1
            
            ' ========== ROW 3: ADDRESS LINE 2 ==========
            If addr2 <> "" Then
                With ws.Range("B" & r & ":G" & r)
                    .Merge
                    .value = addr2
                    .Font.Size = 13
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlCenter
                End With
                ws.Rows(r).RowHeight = 22
                r = r + 1
            End If

            ' ========== ROW 4: MOBILE & EMAIL ==========
            With ws.Range("B" & r & ":G" & r)
                .Merge
                .value = "Mobile: " & compMob & "  |  Email: " & compEmail
                .Font.Size = 13
                .Font.Bold = True
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
            End With
            ws.Rows(r).RowHeight = 26
            r = r + 1
            
            ' ========== ROW 5: WARRANTY HEADER - NO GAP ==========
            With ws.Range("B" & r & ":G" & r)
                .Merge
                .value = "WARRANTY AND SERVICE REPORT"
                .Font.Bold = True
                .Font.Size = 16
                .Font.Color = RGB(255, 255, 255)
                .Interior.Color = RGB(0, 128, 0)
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
            End With
            ws.Rows(r).RowHeight = 32
            r = r + 1
            
            ' ========== ROW 6: CUSTOMER ID HEADER - NO GAP ==========
            With ws.Range("B" & r & ":G" & r)
                .Merge
                .value = "CUSTOMER ID: " & m_CustomerID
                .Font.Bold = True
                .Font.Size = 13
                .Font.Color = RGB(255, 255, 255)
                .Interior.Color = RGB(0, 0, 128)
                .HorizontalAlignment = xlLeft
                .VerticalAlignment = xlCenter
            End With
            ws.Rows(r).RowHeight = 28
            
            ' CUSTOMER PHOTO - Positioned properly (moved down)
            If showCustPhoto = "YES" And custPhotoPath <> "" Then
                If Dir(custPhotoPath & m_Mobile & ".jpg") <> "" Then
                    On Error Resume Next
                    Set shpPhoto = ws.Shapes.AddPicture(custPhotoPath & m_Mobile & ".jpg", msoFalse, msoTrue, 0, 0, 100, 80)
                    If Not shpPhoto Is Nothing Then
                        With shpPhoto
                            .LockAspectRatio = msoTrue
                            .Width = 100
                            .Height = 80
                            .Top = ws.Range("F" & r).Top + 30
                            .Left = ws.Range("F" & r).Left + 40
                        End With
                    End If
                    On Error GoTo ErrorHandler
                End If
            End If
            r = r + 1

            ' ========== ROW 7: CUSTOMER NAME ==========
            ws.Range("B" & r).value = "Name:"
            ws.Range("B" & r).Font.Bold = True
            ws.Range("B" & r).Font.Size = 12
            ws.Range("C" & r & ":E" & r).Merge
            ws.Range("C" & r).value = m_CustomerName
            ws.Range("C" & r).Font.Size = 12
            r = r + 1

            ' ========== ROW 8: MOBILE (LEFT ALIGNED) ==========
            ws.Range("B" & r).value = "Mobile:"
            ws.Range("B" & r).Font.Bold = True
            ws.Range("B" & r).Font.Size = 12
            ws.Range("C" & r).value = m_Mobile
            ws.Range("C" & r).Font.Size = 12
            ws.Range("C" & r).HorizontalAlignment = xlLeft
            r = r + 1
            
            ' Get customer details
            custEmail = ""
            custGST = ""
            Dim custAddress As String
            custAddress = ""
            For i = 2 To wsCust.Cells(wsCust.Rows.count, 1).End(xlUp).row
                If UCase(Trim(wsCust.Cells(i, 1).value & "")) = UCase(m_CustomerID) Then
                    custEmail = wsCust.Cells(i, 5).value & ""
                    custGST = wsCust.Cells(i, 6).value & ""
                    custAddress = wsCust.Cells(i, 4).value & ""
                    Exit For
                End If
            Next i
            
            ' ========== ROW 9: EMAIL ==========
            If custEmail <> "" Then
                ws.Range("B" & r).value = "Email:"
                ws.Range("B" & r).Font.Bold = True
                ws.Range("B" & r).Font.Size = 12
                ws.Range("C" & r & ":E" & r).Merge
                ws.Range("C" & r).value = custEmail
                ws.Range("C" & r).Font.Size = 12
                r = r + 1
            End If

            ' ========== ROW 10: ADDRESS (CENTERED) ==========
            If custAddress <> "" Then
    
    ws.Range("B" & r).value = "Address:"
    ws.Range("B" & r).Font.Bold = True
    ws.Range("B" & r).Font.Size = 12
    
    ws.Range("C" & r & ":G" & r).Merge
    ws.Range("C" & r).value = Replace(Replace(custAddress, vbCrLf, ", "), vbLf, ", ")
    ws.Range("C" & r).Font.Size = 12
    ws.Range("C" & r).HorizontalAlignment = xlLeft
    
    ws.Rows(r).RowHeight = 22
    
    r = r + 1
End If
            
            ' ========== ROW 11: GST ==========
            If custGST <> "" Then
                ws.Range("B" & r).value = "GST:"
                ws.Range("B" & r).Font.Bold = True
                ws.Range("B" & r).Font.Size = 12
                ws.Range("C" & r).value = custGST
                ws.Range("C" & r).Font.Size = 12
                r = r + 1
            End If
            
            r = r + 1
        End If
        
        ' ========== ENTRY DATA ==========
        entriesOnPage = entriesOnPage + 1
        prodRow = entryRows(entryNum)
        eID = wsProd.Cells(prodRow, 1).value & ""
        
        ' Get Entry Date and Verify Info from Job_Master
        entryDate = ""
        verifyType = ""
        verifyName = ""
        For i = 2 To wsJob.Cells(wsJob.Rows.count, 1).End(xlUp).row
            If Trim(UCase(wsJob.Cells(i, 1).value & "")) = Trim(UCase(eID)) Then
                entryDate = wsJob.Cells(i, 8).value & ""
                verifyType = wsJob.Cells(i, 6).value & ""
                verifyName = wsJob.Cells(i, 7).value & ""
                If entryDate = "" Then entryDate = Format(Date, "dd-MMM-yyyy")
                Exit For
            End If
        Next i
        If entryDate = "" Then entryDate = Format(Date, "dd-MMM-yyyy")

        ' Entry Header - LEFT ALIGNED
        With ws.Range("B" & r & ":G" & r)
            .Merge
            .value = "ENTRY ID: " & eID & "  |  Date: " & entryDate
            .Font.Bold = True
            .Font.Size = 13
            .Font.Color = RGB(255, 255, 255)
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter
            If UCase(Left(eID, 3)) = "WAR" Then
                .Interior.Color = RGB(200, 0, 0)
            ElseIf UCase(Left(eID, 3)) = "SER" Then
                .Interior.Color = RGB(255, 165, 0)
            Else
                .Interior.Color = RGB(0, 100, 150)
            End If
        End With
        ws.Rows(r).RowHeight = 28
        r = r + 1

        ' Product Details
        ws.Range("B" & r).value = "Product:"
        ws.Range("B" & r).Font.Bold = True
        ws.Range("B" & r).Font.Size = 11
        ws.Range("C" & r & ":D" & r).Merge
        ws.Range("C" & r).value = wsProd.Cells(prodRow, 6).value
        ws.Range("C" & r).Font.Size = 11
        ws.Range("E" & r).value = "Company:"
        ws.Range("E" & r).Font.Bold = True
        ws.Range("E" & r).Font.Size = 11
        ws.Range("F" & r & ":G" & r).Merge
        ws.Range("F" & r).value = wsProd.Cells(prodRow, 7).value
        ws.Range("F" & r).Font.Size = 11
        r = r + 1
        
        ws.Range("B" & r).value = "Model:"
        ws.Range("B" & r).Font.Bold = True
        ws.Range("B" & r).Font.Size = 11
        ws.Range("C" & r).value = wsProd.Cells(prodRow, 8).value
        ws.Range("C" & r).Font.Size = 11
        ws.Range("D" & r).value = "Serial:"
        ws.Range("D" & r).Font.Bold = True
        ws.Range("D" & r).Font.Size = 11
        ws.Range("E" & r & ":F" & r).Merge
        ws.Range("E" & r).value = wsProd.Cells(prodRow, 9).value
        ws.Range("E" & r).Font.Size = 11
        r = r + 1

        ' Problem
        pProblem = wsProd.Cells(prodRow, 13).value & ""
        If pProblem <> "" Then
            With ws.Range("B" & r & ":G" & r)
                .Merge
                .value = "Problem: " & pProblem
                .Font.Color = RGB(200, 0, 0)
                .Font.Size = 11
            End With
            r = r + 1
        End If

        ' PRODUCT PHOTO
        If showProdPhoto = "YES" And prodPhotoPath <> "" Then
            For i = 1 To 4
                If Dir(prodPhotoPath & eID & "\Photo" & i & ".jpg") <> "" Then
                    On Error Resume Next
                    Set shpPhoto = ws.Shapes.AddPicture(prodPhotoPath & eID & "\Photo" & i & ".jpg", msoFalse, msoTrue, 0, 0, 90, 90)
                    If Not shpPhoto Is Nothing Then
                        With shpPhoto
                            .LockAspectRatio = msoTrue
                            .Width = 90
                            .Height = 90
                            .Top = ws.Range("G" & (r - 3)).Top + 6
                            .Left = ws.Range("G" & (r - 3)).Left - 5
                        End With
                    End If
                    On Error GoTo ErrorHandler
                    Exit For
                End If
            Next i
        End If
        
        ' Accessories
        lastRowAcc = wsAcc.Cells(wsAcc.Rows.count, 1).End(xlUp).row
        accCount = 0
        For accRow = 2 To lastRowAcc
            If Trim(UCase(wsAcc.Cells(accRow, 1).value & "")) = Trim(UCase(eID)) Then accCount = accCount + 1
        Next accRow
        
        If accCount > 0 Then
            With ws.Range("B" & r & ":G" & r)
                .Merge
                .value = "ACCESSORIES:"
                .Font.Color = RGB(0, 0, 150)
                .Font.Bold = True
                .Font.Size = 11
            End With
            r = r + 1
            
            ' Header row
            ws.Range("B" & r).value = "S.No"
            ws.Range("B" & r).Font.Bold = True
            ws.Range("B" & r).Font.Size = 10
            ws.Range("B" & r).HorizontalAlignment = xlCenter
            
            ws.Range("C" & r).value = "Accessory Name"
            ws.Range("C" & r).Font.Bold = True
            ws.Range("C" & r).Font.Size = 10
            
            ws.Range("D" & r).value = "Brand"
            ws.Range("D" & r).Font.Bold = True
            ws.Range("D" & r).Font.Size = 10
            
            ws.Range("E" & r).value = "Serial"
            ws.Range("E" & r).Font.Bold = True
            ws.Range("E" & r).Font.Size = 10
            
            ws.Range("F" & r & ":G" & r).Merge
            ws.Range("F" & r).value = "Received"
            ws.Range("F" & r).Font.Bold = True
            ws.Range("F" & r).Font.Size = 10
            ws.Range("F" & r).HorizontalAlignment = xlLeft
            r = r + 1
            
            Dim accIndex As Integer
            accIndex = 1
            For accRow = 2 To lastRowAcc
                If Trim(UCase(wsAcc.Cells(accRow, 1).value & "")) = Trim(UCase(eID)) Then
                    ws.Rows(r).RowHeight = 26
                    
                    ' Accessory Photo - BIGGER
                    If showAccPhoto = "YES" And accPhotoPath <> "" Then
                        If Dir(accPhotoPath & eID & "\Acc" & accIndex & ".jpg") <> "" Then
                            On Error Resume Next
                            Set shpPhoto = ws.Shapes.AddPicture(accPhotoPath & eID & "\Acc" & accIndex & ".jpg", msoFalse, msoTrue, 0, 0, 40, 30)
                            If Not shpPhoto Is Nothing Then
                                With shpPhoto
                                    .LockAspectRatio = msoTrue
                                    .Width = 40
                                    .Height = 30
                                    .Top = ws.Range("B" & r).Top - 15
                                    .Left = ws.Range("B" & r).Left + 5
                                End With
                            End If
                            On Error GoTo ErrorHandler
                        End If
                    End If
                    
                    ' Accessory Details
                    ws.Range("B" & r).value = accIndex
                    ws.Range("B" & r).Font.Size = 10
                    ws.Range("B" & r).HorizontalAlignment = xlCenter
                    
                    ws.Range("C" & r).value = wsAcc.Cells(accRow, 4).value
                    ws.Range("C" & r).Font.Size = 10
                    
                    ws.Range("D" & r).value = wsAcc.Cells(accRow, 5).value
                    ws.Range("D" & r).Font.Size = 10
                    
                    ws.Range("E" & r).value = wsAcc.Cells(accRow, 6).value
                    ws.Range("E" & r).Font.Size = 10
                    
                    ws.Range("F" & r & ":G" & r).Merge
                    ws.Range("F" & r).value = wsAcc.Cells(accRow, 7).value
                    ws.Range("F" & r).Font.Size = 10
                    ws.Range("F" & r).HorizontalAlignment = xlLeft
                    
                    r = r + 1
                    accIndex = accIndex + 1
                    If accIndex > 4 Then Exit For
                End If
            Next accRow
        End If

        ' Payment - ALL IN ONE LINE
        lastRowPay = wsPay.Cells(wsPay.Rows.count, 1).End(xlUp).row
        foundPayment = False
        totalCharges = 0
        totalExp = 0
        totalRec = 0
        payID = ""

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
            
            ' Payment Header - ONE LINE
            With ws.Range("B" & r & ":G" & r)
                .Merge
                .value = "Payment Details  |  Entry ID: " & eID & "  |  Pay ID: " & payID
                .Font.Color = RGB(255, 255, 255)
                .Font.Bold = True
                .Font.Size = 11
                .Interior.Color = RGB(0, 100, 0)
                .HorizontalAlignment = xlLeft
            End With
            r = r + 1
            
            ' All amounts in ONE LINE
            With ws.Range("B" & r & ":G" & r)
                .Merge
                If totalExp > 0 Then
                    .value = "Service Charge: Rs." & Format(totalCharges, "0.00") & "  |  Expense: Rs." & Format(totalExp, "0.00") & "  |  Received: Rs." & Format(totalRec, "0.00") & "  |  Due: Rs." & Format(dueAmt, "0.00")
                Else
                    .value = "Service Charge: Rs." & Format(totalCharges, "0.00") & "  |  Received: Rs." & Format(totalRec, "0.00") & "  |  Due: Rs." & Format(dueAmt, "0.00")
                End If
                .Font.Color = RGB(0, 100, 0)
                .Font.Size = 11
                .HorizontalAlignment = xlLeft
            End With
            r = r + 1
        Else
            With ws.Range("B" & r & ":G" & r)
                .Merge
                .value = "Payment: No transactions found"
                .Font.Color = RGB(128, 128, 128)
                .Font.Size = 11
            End With
            r = r + 1
        End If

        ' Verification - CENTERED
        If verifyType <> "" Or verifyName <> "" Then
            With ws.Range("B" & r & ":G" & r)
                .Merge
                .value = "Verified Type: " & verifyType & "  |  Verified By: " & verifyName
                .Font.Bold = True
                .Font.Size = 11
                .Font.Color = RGB(0, 0, 128)
                .HorizontalAlignment = xlCenter
                .Interior.Color = RGB(240, 248, 255)
            End With
            
            With ws.Range("B" & r & ":G" & r).Borders
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .Color = RGB(0, 0, 128)
            End With
            
            ws.Rows(r).RowHeight = 26
            r = r + 1
        End If
        
        r = r + 1
        
    Next entryNum
    
    ' Add FOOTER to last page
    Call AddPageFooter(ws, r, termsFound, wsSet, terms)
   
    ws.Activate
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Set GenerateReceiptSheet = ws
    Exit Function
    
ErrorHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "Error: " & Err.Description, vbCritical
    Set GenerateReceiptSheet = Nothing

End Function

' Footer subroutine
Private Sub AddPageFooter(ws As Worksheet, ByRef r As Long, termsFound As Boolean, wsSet As Worksheet, terms() As String, Optional vType As String = "", Optional vName As String = "")
    Dim termRow As Long, t As Integer
    
    r = r + 1
    
    
    
    ' TERMS & CONDITIONS
    With ws.Range("B" & r & ":G" & r)
        .Merge
        .value = "TERMS & CONDITIONS"
        .Font.Bold = True
        .Font.Size = 12
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(128, 0, 0)
        .HorizontalAlignment = xlCenter
    End With
    ws.Rows(r).RowHeight = 24
    r = r + 1
    
    ' Terms Content
    If termsFound Then
        For termRow = 25 To 32
            If Trim(wsSet.Cells(termRow, 2).value & "") <> "" Then
                With ws.Range("B" & r & ":G" & r)
                    .Merge
                    .value = wsSet.Cells(termRow, 2).value
                    .Font.Size = 9
                    .WrapText = True
                End With
                ws.Rows(r).RowHeight = 16
                r = r + 1
            End If
        Next termRow
    Else
        For t = 1 To 8
            With ws.Range("B" & r & ":G" & r)
                .Merge
                .value = terms(t)
                .Font.Size = 9
                .WrapText = True
            End With
            ws.Rows(r).RowHeight = 16
            r = r + 1
        Next t
    End If
    
    
    
       ' ===== SIGNATURES SECTION =====
    r = r + 2  ' ????? Gap ??????
    
    ' --- ???? ???? ????? (????? ???? ???) ---
    ' Customer Signature Line (Left Side)
    With ws.Range("B" & r & ":C" & r)
        .Merge
        .value = ""
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeTop).Weight = xlThin
    End With
    
    ' Authorized Signature Line (Right Side)
    With ws.Range("F" & r & ":G" & r)
        .Merge
        .value = ""
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeTop).Weight = xlThin
    End With
    
    ' ??? ???? ???? ??????? ?????
    r = r + 1
    
    ' Customer Signature Text
    With ws.Range("B" & r & ":C" & r)
        .Merge
        .value = "Customer Signature"
        .Font.Bold = True
        .Font.Size = 11
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlTop
    End With
    
    ' Authorized Signature Text
    With ws.Range("F" & r & ":G" & r)
        .Merge
        .value = "Authorized Signature"
        .Font.Bold = True
        .Font.Size = 11
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlTop
    End With
   
        ' For GLOBAL IT SOLUTIONS (???? ?? ???? ??)
    r = r + 1
    
    With ws.Range("F" & r & ":G" & r)
        .Merge
        .value = "For GLOBAL IT SOLUTIONS"
        .Font.Italic = True
        .Font.Size = 11
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
   
    ' THANK YOU
    r = r + 2
    With ws.Range("B" & r & ":G" & r)
        .Merge
        .value = "*** Thank You for choosing US ***"
        .Font.Bold = True
        .Font.Size = 12
        .Font.name = "Arial"
        .Font.Color = RGB(0, 100, 0)
        .HorizontalAlignment = xlCenter
        .Interior.Color = RGB(255, 250, 205)
    End With
    ws.Rows(r).RowHeight = 26
    
    r = r + 2
End Sub
'========================================
' WHATSAPP BUTTON - ??????? (Direct App)
'========================================
Private Sub btnWhatsApp_Click()
    On Error GoTo ErrorHandler
    
    If Trim(m_Mobile) = "" Then
        MsgBox "Mobile number not found!", vbExclamation
        Exit Sub
    End If
    
    ' ?????? ???? ??? ???? (10 ?????)
    Dim cleanMobile As String
    cleanMobile = FormatMobileNumber(m_Mobile)
    
    ' 91 ????? ??? ???? ?? ?? ?? (wa.me ??? ???????)
    If Left(cleanMobile, 2) = "91" And Len(cleanMobile) > 10 Then
        cleanMobile = Right(cleanMobile, 10)
    End If
    
    ' ???? ????? ?????
    Dim msg As String
    msg = BuildWhatsAppMessage()
    
    ' URL Encode ???? (spaces ?? special chars ?? ???)
    msg = URLEncode(msg)
    
    ' ***** Direct WhatsApp Desktop App ????? *****
    ' ??? ??? ????, ???? ?? ??????
    Dim url As String
    url = "whatsapp://send?phone=91" & cleanMobile & "&text=" & msg
    
    ' Direct App ????? ?? ??? ShellExecute
    ShellExecute 0, "open", url, vbNullString, vbNullString, 1
    
    Exit Sub
    
ErrorHandler:
    MsgBox "WhatsApp open karne me error: " & Err.Description, vbExclamation
End Sub
'========================================
' URL Encode Function (WhatsApp ?? ??? ?????)
'========================================
Private Function URLEncode(ByVal StringVal As String) As String
    Dim i As Integer
    Dim CharCode As Integer
    Dim char As String
    Dim OutString As String
    
    For i = 1 To Len(StringVal)
        char = Mid(StringVal, i, 1)
        CharCode = Asc(char)
        
        ' Safe characters ?? ???? ?? ??????
        If (CharCode >= 48 And CharCode <= 57) Or _
           (CharCode >= 65 And CharCode <= 90) Or _
           (CharCode >= 97 And CharCode <= 122) Or _
           char = "-" Or char = "_" Or char = "." Or char = "~" Then
            OutString = OutString & char
        ElseIf char = " " Then
            ' Space ?? + ?? %20 ?? ??????? ???? (WhatsApp %0A ????? ?? ???????? ?? ???)
            OutString = OutString & "%20"
        ElseIf char = vbLf Or char = vbCr Then
            ' New Line ?? ??? %0A (Line Feed)
            OutString = OutString & "%0A"
        Else
            ' ???? ?? ?? Hex ??? Convert ????
            OutString = OutString & "%" & Right("0" & Hex(CharCode), 2)
        End If
    Next i
    
    URLEncode = OutString
End Function
Private Function BuildWhatsAppMessage() As String
    On Error Resume Next
    
    Dim msg As String
    Dim wsCust As Worksheet, wsProd As Worksheet, wsAcc As Worksheet, wsJob As Worksheet
    Dim wsConfig As Worksheet
    Dim i As Long, entryNum As Long
    Dim compName As String, compAddr1 As String, compAddr2 As String, compMob As String
    Dim entryRows As Collection
    Dim prodRow As Long, eID As String
    Dim custAddress As String, custEmail As String
    
    ' === SHEETS SET ===
    Set wsCust = ThisWorkbook.Sheets("Customer_Master")
    Set wsProd = ThisWorkbook.Sheets("Job_Product")
    Set wsAcc = ThisWorkbook.Sheets("Job_Accessory")
    Set wsJob = ThisWorkbook.Sheets("Job_Master")
    Set wsConfig = ThisWorkbook.Sheets("Software_Config")
    
    ' === COMPANY INFO ===
    compName = Trim(wsConfig.Range("B2").value & "")
    If compName = "" Then compName = "GLOBAL IT SOLUTIONS"
    
    compAddr1 = Trim(wsConfig.Range("B3").value & "")
    compAddr2 = Trim(wsConfig.Range("B4").value & "")
    If compAddr2 = "" Then compAddr2 = "Nayagarh-752069, Odisha"
    
    compMob = Trim(wsConfig.Range("B9").value & "")
    If compMob = "" Then compMob = "9777971045"
    
    ' === CUSTOMER DETAILS ===
    custAddress = ""
    custEmail = ""
    For i = 2 To wsCust.Cells(wsCust.Rows.count, 1).End(xlUp).row
        If UCase(Trim(wsCust.Cells(i, 1).value & "")) = UCase(m_CustomerID) Then
            custAddress = wsCust.Cells(i, 4).value & ""
            custEmail = wsCust.Cells(i, 5).value & ""
            Exit For
        End If
    Next i
    
    ' === BUILD MESSAGE (????????? ???????) ===
    ' NOTE: vbLf ???? ??? (????? ????), vbCrLf ?? ???? ??? (??? ??? ???? ??)
    
    msg = "*" & compName & "*" & vbLf & _
          compAddr1 & vbLf & _
          compAddr2 & vbLf & _
          "WhatsApp & Call: " & compMob & vbLf & _
          "------------------------------" & vbLf & _
          "*YOUR PRODUCT RECEIVED*" & vbLf & _
          "------------------------------" & vbLf & _
          "*ID:* " & m_CustomerID & vbLf & _
          "*Name:* " & m_CustomerName & vbLf & _
          "*Mobile:* " & m_Mobile
    
    ' Address (??? ?? ??)
    If custAddress <> "" Then
        msg = msg & vbLf & "*Address:* " & Replace(custAddress, vbCrLf, ", ")
    End If
    
    ' Email (??? ?? ??)
    If custEmail <> "" Then
        msg = msg & vbLf & "*Email:* " & custEmail
    End If
    
    ' === ENTRIES ===
    Set entryRows = New Collection
    For i = 2 To wsProd.Cells(wsProd.Rows.count, 1).End(xlUp).row
        If Trim(UCase(wsProd.Cells(i, 2).value & "")) = Trim(UCase(m_CustomerID)) Then
            entryRows.Add i
        End If
    Next i
    
    For entryNum = 1 To entryRows.count
        prodRow = entryRows(entryNum)
        eID = wsProd.Cells(prodRow, 1).value & ""
        
        ' Entry Type
        Dim entryType As String
        If UCase(Left(eID, 3)) = "WAR" Then
            entryType = "WARRANTY"
        ElseIf UCase(Left(eID, 3)) = "SER" Then
            entryType = "SERVICE"
        Else
            entryType = "ENTRY"
        End If
        
        ' Date from Job_Master
        Dim entryDate As String
        entryDate = ""
        For i = 2 To wsJob.Cells(wsJob.Rows.count, 1).End(xlUp).row
            If Trim(UCase(wsJob.Cells(i, 1).value & "")) = Trim(UCase(eID)) Then
                entryDate = wsJob.Cells(i, 8).value & ""
                Exit For
            End If
        Next i
        If entryDate = "" Then entryDate = Format(Date, "dd-MMM-yyyy")
        
        ' Entry Details (?????????)
        msg = msg & vbLf & "------------------------------" & vbLf & _
              "*" & entryType & " ID: " & eID & "*" & vbLf & _
              "Date: " & entryDate & vbLf & _
              "Product: " & wsProd.Cells(prodRow, 6).value & " | " & wsProd.Cells(prodRow, 7).value & vbLf & _
              "Model: " & wsProd.Cells(prodRow, 8).value & " | S/N: " & wsProd.Cells(prodRow, 9).value
        
        ' Problem (?? ?? ???? ???)
        Dim prob As String
        prob = wsProd.Cells(prodRow, 13).value & ""
        If prob <> "" Then
            msg = msg & vbLf & "Issue: " & prob
        End If
        
        ' Accessories
        Dim accRow As Long, lastRowAcc As Long, accCount As Integer
        lastRowAcc = wsAcc.Cells(wsAcc.Rows.count, 1).End(xlUp).row
        accCount = 0
        
        For accRow = 2 To lastRowAcc
            If Trim(UCase(wsAcc.Cells(accRow, 1).value & "")) = Trim(UCase(eID)) Then
                accCount = accCount + 1
                If accCount = 1 Then msg = msg & vbLf & "Accessories:"
                
                Dim accName As String, accReceived As String
                accName = wsAcc.Cells(accRow, 4).value & ""
                accReceived = wsAcc.Cells(accRow, 7).value & ""
                
                msg = msg & vbLf & accCount & ". " & accName
                If accReceived <> "" Then msg = msg & " (" & accReceived & ")"
                
                If accCount >= 4 Then Exit For
            End If
        Next accRow
        
    Next entryNum
    
    ' Footer
    msg = msg & vbLf & "------------------------------" & vbLf & _
          "*Thank You for choosing " & compName & "!*"
    
    BuildWhatsAppMessage = msg
End Function
Function CleanText(txt As String) As String
    txt = Replace(txt, vbCrLf, " ")
    txt = Replace(txt, vbLf, " ")
    txt = Replace(txt, vbCr, " ")
    CleanText = Trim(txt)
End Function



Private Sub SetupButtonStyles()
    On Error Resume Next
    
    ' Print - ???? ???? (?????????)
    Me.btnPrintReceipt.BackColor = RGB(41, 128, 185)
    Me.btnPrintReceipt.ForeColor = RGB(255, 255, 255)  ' ???? ???????
    
    ' PDF - ????? ???
    Me.btnExportPDF.BackColor = RGB(192, 57, 43)
    Me.btnExportPDF.ForeColor = RGB(255, 255, 255)
    
    ' WhatsApp - ?????? ????? (????? ?????)
    Me.btnWhatsApp.BackColor = RGB(7, 94, 84)
    Me.btnWhatsApp.ForeColor = RGB(255, 255, 255)
    
    ' Email - ????? ????
    Me.btnSendMail.BackColor = RGB(52, 152, 219)
    Me.btnSendMail.ForeColor = RGB(255, 255, 255)
    
    
    
    ' Cancel - ????
    Me.btnCancelOutput.BackColor = RGB(127, 140, 141)
    Me.btnCancelOutput.ForeColor = RGB(255, 255, 255)
End Sub



