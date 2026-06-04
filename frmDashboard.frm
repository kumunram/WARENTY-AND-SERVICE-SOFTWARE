VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDashboard 
   Caption         =   "GLOBAL SOFT Dashboard"
   ClientHeight    =   13710
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   22320
   OleObjectBlob   =   "frmDashboard.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmDashboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit




'========================================
' USERFORM INITIALIZE
'========================================
Private Sub UserForm_Initialize()
    On Error GoTo InitError
    
    ' === ???????? ???? Init ??? ???? ===
    
    ' Min/Max buttons — NORMAL FLOW ??? ??? (Error handler ??? ????)
    AddMinMaxButtons Me
    
    Exit Sub

InitError:
    MsgBox "Error in frmDashboard Initialize" & vbCrLf & _
           "Error: " & Err.Number & " - " & Err.Description, vbCritical
    On Error Resume Next
    If Not IsLoggedIn Then
        Unload Me
        Exit Sub
    End If
    On Error GoTo 0
End Sub

Private Sub UserForm_Activate()
    On Error Resume Next
     
    ' === SHOW CURRENT USER ===
    If Not lblCurrentUser Is Nothing Then
        If CurrentUserName <> "" Then
            lblCurrentUser.caption = "User: " & UCase(CurrentUserName) & " (" & UCase(CurrentUserRole) & ")"
        Else
            lblCurrentUser.caption = "User: GUEST"
        End If
    End If

    ' === HEADER ===
    Me.caption = "Service Management System - Dashboard"
    lblDate.caption = Format(Date, "dd-mmm-yyyy")
    lblTime.caption = Format(Time, "hh:mm:ss AM/PM")

    ' === LOAD COMPANY INFO ===
    LoadCompanyInfo

    ' === SETUP & REFRESH LISTVIEWS ===
    SetupListViews
    RefreshDashboard

    ' === APPLY PERMISSIONS ===
    ApplyDashboardPermissions

    ' ============================================
    ' AUTO BACKUP ON OPEN (SIRF EK BAAR)
    ' ============================================
    If Not modAutoBackup.g_OpenBackupDone Then
        If Now - modAutoBackup.GetLastBackupTime >= TimeValue("06:00:00") Then
            modAutoBackup.DoAutoBackup True
        End If
        modAutoBackup.g_OpenBackupDone = True
    End If
    
    ' Timer active rakho
    modAutoBackup.ScheduleNextBackup
    
    ' Min/Max — Activate ??? ?? (100% guarantee)
    AddMinMaxButtons Me

    On Error GoTo 0
End Sub

'========================================
' APPLY PERMISSIONS
'========================================
Private Sub ApplyDashboardPermissions()
    On Error Resume Next

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim userFound As Boolean

    Set ws = ThisWorkbook.Sheets("User_Permission")
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    userFound = False

    For i = 2 To lastRow
        If UCase(Trim(ws.Cells(i, 1).value & "")) = UCase(CurrentUserName) Then
            userFound = True
            If Not btnDashboard Is Nothing Then btnDashboard.enabled = (UCase(Trim(ws.Cells(i, "D").value & "")) = "YES")
            If Not btnNewEntry Is Nothing Then btnNewEntry.enabled = (UCase(Trim(ws.Cells(i, "E").value & "")) = "YES")
            If Not btnAssign Is Nothing Then btnAssign.enabled = (UCase(Trim(ws.Cells(i, "F").value & "")) = "YES")
            If Not btnService Is Nothing Then btnService.enabled = (UCase(Trim(ws.Cells(i, "G").value & "")) = "YES")
            If Not btnCustomer Is Nothing Then btnCustomer.enabled = (UCase(Trim(ws.Cells(i, "H").value & "")) = "YES")
            If Not btnEngineer Is Nothing Then btnEngineer.enabled = (UCase(Trim(ws.Cells(i, "I").value & "")) = "YES")
            If Not btnReports Is Nothing Then btnReports.enabled = (UCase(Trim(ws.Cells(i, "J").value & "")) = "YES")
            If Not btnSettings Is Nothing Then btnSettings.enabled = (UCase(Trim(ws.Cells(i, "K").value & "")) = "YES")
            Exit For
        End If
    Next i

    ' Hardcoded Admin/Developer = everything enabled
    If Not userFound Then
        If Not btnDashboard Is Nothing Then btnDashboard.enabled = True
        If Not btnNewEntry Is Nothing Then btnNewEntry.enabled = True
        If Not btnAssign Is Nothing Then btnAssign.enabled = True
        If Not btnService Is Nothing Then btnService.enabled = True
        If Not btnCustomer Is Nothing Then btnCustomer.enabled = True
        If Not btnEngineer Is Nothing Then btnEngineer.enabled = True
        If Not btnReports Is Nothing Then btnReports.enabled = True
        If Not btnSettings Is Nothing Then btnSettings.enabled = True
    End If

    ' Backup always enabled
    If Not btnBackup Is Nothing Then
        btnBackup.enabled = True
        btnBackup.visible = True
    End If

    ' User Permission button
    If Not btnUserPermission Is Nothing Then
        btnUserPermission.visible = True
        If UCase(CurrentUserRole) = "ADMIN" Or UCase(CurrentUserRole) = "DEVELOPER" Then
            btnUserPermission.enabled = True
        Else
            btnUserPermission.enabled = False
        End If
    End If

    ' Exit & Logout always enabled
    If Not btnExit Is Nothing Then btnExit.enabled = True
    If Not btnLogout Is Nothing Then btnLogout.enabled = True

    On Error GoTo 0
End Sub

'========================================
' NAVIGATION BUTTONS
'========================================
Private Sub btnNewEntry_Click()
    On Error GoTo ErrorHandler
    frmEntryWizard.Show vbModal
    RefreshDashboard
    Exit Sub
ErrorHandler:
    MsgBox "New Entry button error: " & Err.Description, vbCritical, "Error in btnNewEntry"
End Sub

Private Sub btnAssign_Click()
    On Error GoTo handler
    frmAssignmentDashboard.Show vbModal
    RefreshDashboard
    Exit Sub
handler:
    MsgBox "Assign Work button error: " & Err.Description, vbCritical, "Error in btnAssign"
End Sub

Private Sub btnService_Click()
    On Error GoTo handler
    frmStatusManager.Show vbModal
    RefreshDashboard
    Exit Sub
handler:
    MsgBox "Service Update button error: " & Err.Description, vbCritical, "Error in btnService"
End Sub

Private Sub btnCustomer_Click()
    On Error GoTo handler
    frmCustomerList.Show vbModal
    Exit Sub
handler:
    MsgBox "Customer DB button error: " & Err.Description, vbCritical, "Error in btnCustomer"
End Sub

Private Sub btnEngineer_Click()
    On Error GoTo handler
    frmEngineerRegister.Show vbModal
    RefreshDashboard
    Exit Sub
handler:
    MsgBox "Engineer Register button error: " & Err.Description, vbCritical, "Error in btnEngineer"
End Sub

Private Sub btnReports_Click()
    On Error GoTo handler
    frmReports.Show vbModal
    Exit Sub
handler:
    MsgBox "Reports button error: " & Err.Description, vbCritical, "Error in btnReports"
End Sub

Private Sub btnSettings_Click()
    On Error GoTo handler
    frmSetting1.Show vbModal
    RefreshDashboard
    Exit Sub
handler:
    MsgBox "Settings button error: " & Err.Description, vbCritical, "Error in btnSettings"
End Sub

Private Sub btnBackup_Click()
    On Error GoTo handler
    frmBackup.Show vbModal
    RefreshDashboard
    Exit Sub
handler:
    MsgBox "Backup button error: " & Err.Description, vbCritical, "Error in btnBackup"
End Sub

Private Sub btnUserPermission_Click()
    frmUserPermission.Show vbModal
End Sub

Private Sub btnExit_Click()
    On Error Resume Next
    If MsgBox("Exit Application?", vbQuestion + vbYesNo, "Exit") = vbYes Then
        ThisWorkbook.Save
        Unload Me
    End If
    On Error GoTo 0
End Sub

Private Sub btnLogout_Click()
    On Error Resume Next
    IsLoggedIn = False
    CurrentUserName = ""
    CurrentUserRole = ""
    
    Me.Hide                    ' ???? ???????? ?????
    frmLogin.Show vbModal      ' ????? ??????? ?????
    Unload Me                  ' ??? ???????? ??? ???
    
    On Error GoTo 0
End Sub

'========================================
' LABEL CLICKS
'========================================
Private Sub lblTitle_Click()
    MsgBox lblTitle.caption & " - Warranty & Service Management", vbInformation, "About"
End Sub

Private Sub lblCompany_Click()
    frmCompanySettings.Show vbModal
    LoadCompanyInfo
End Sub

Private Sub lblAddress_Click()
    frmCompanySettings.Show vbModal
    LoadCompanyInfo
End Sub

Private Sub lblMobile_Click()
    MsgBox "Contact: " & lblMobile.caption, vbInformation, "Contact"
End Sub

Private Sub lblDate_Click()
    lblDate.caption = Format(Date, "dd-mmm-yyyy")
End Sub

Private Sub lblTime_Click()
    lblTime.caption = Format(Time, "hh:mm:ss AM/PM")
End Sub

Private Sub lblVersion_Click()
    MsgBox "Software Version: " & lblVersion.caption, vbInformation, "Version"
End Sub

Private Sub lblDeveloper_Click()
    MsgBox "Developer: " & lblDeveloper.caption & vbCrLf & "Helpline: " & lblHelpLine.caption, vbInformation, "Developer"
End Sub

Private Sub lblHelpLine_Click()
    MsgBox "HelpLine: " & lblHelpLine.caption, vbInformation, "HelpLine"
End Sub

'========================================
' LISTVIEW EVENTS
'========================================
Private Sub lstAssignPending1_BeforeLabelEdit(Cancel As Integer)
    Cancel = True
End Sub

Private Sub lstVendorPending1_BeforeLabelEdit(Cancel As Integer)
    Cancel = True
End Sub

Private Sub lstEngineerWork1_BeforeLabelEdit(Cancel As Integer)
    Cancel = True
End Sub

Private Sub lstReadyDelivery1_BeforeLabelEdit(Cancel As Integer)
    Cancel = True
End Sub




'========================================
' LISTVIEW DOUBLE CLICKS
'========================================
Private Sub lstAssignPending1_DblClick()
    If lstAssignPending1.selectedItem Is Nothing Then Exit Sub
    Dim entryID As String
    entryID = lstAssignPending1.selectedItem.text
    Me.Hide
    frmEntryWizard.Tag = entryID
    frmEntryWizard.Show vbModal
    Me.Show
    RefreshDashboard
End Sub

Private Sub lstVendorPending1_DblClick()
    If lstVendorPending1.selectedItem Is Nothing Then Exit Sub
    Dim entryID As String
    entryID = lstVendorPending1.selectedItem.text
    Me.Hide
    frmWarrantyAssignment.Tag = entryID
    frmWarrantyAssignment.Show vbModal
    Me.Show
    RefreshDashboard
End Sub

Private Sub lstEngineerWork1_DblClick()
    If lstEngineerWork1.selectedItem Is Nothing Then Exit Sub
    Dim entryID As String
    entryID = lstEngineerWork1.selectedItem.text
    Me.Hide
    frmServiceAssignment.Tag = entryID
    frmServiceAssignment.Show vbModal
    Me.Show
    RefreshDashboard
End Sub

Private Sub lstReadyDelivery1_DblClick()
    If lstReadyDelivery1.selectedItem Is Nothing Then Exit Sub
    Dim entryID As String
    entryID = lstReadyDelivery1.selectedItem.text
    Me.Hide
    frmDelivery.Tag = entryID
    frmDelivery.Show vbModal
    Me.Show
    RefreshDashboard
End Sub

'========================================
' LOAD COMPANY INFO
'========================================
Private Sub LoadCompanyInfo()
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Software_Config")
    If ws Is Nothing Then GoTo SetDefaults

    Dim companyName As String, addr1 As String, addr2 As String
    Dim city As String, pin As String, state As String, country As String
    Dim mobile As String

    companyName = Trim(CStr(ws.Range("B2").value))
    addr1 = Trim(CStr(ws.Range("B3").value))
    addr2 = Trim(CStr(ws.Range("B4").value))
    city = Trim(CStr(ws.Range("B5").value))
    pin = Trim(CStr(ws.Range("B6").value))
    state = Trim(CStr(ws.Range("B7").value))
    country = Trim(CStr(ws.Range("B8").value))
    mobile = Trim(CStr(ws.Range("B9").value))

    Dim swName As String, swVer As String, devName As String, devMobile As String
    Dim lastCol As Long, c As Long
    swName = "": swVer = "": devName = "": devMobile = ""

    lastCol = ws.Cells(16, ws.Columns.count).End(xlToLeft).Column
    For c = 1 To lastCol
        Select Case LCase(Replace(Trim(CStr(ws.Cells(16, c).value)), " ", ""))
            Case "software_name": swName = Trim(CStr(ws.Cells(17, c).value))
            Case "software_version": swVer = Trim(CStr(ws.Cells(17, c).value))
            Case "developer": devName = Trim(CStr(ws.Cells(17, c).value))
            Case "developermobile": devMobile = Trim(CStr(ws.Cells(17, c).value))
        End Select
    Next c

    lblCompany.caption = companyName
    lblTitle.caption = "Warranty & Service Management"
    lblSoftware.caption = swName

    Dim fullAddress As String
    fullAddress = ""
    If addr1 <> "" Then fullAddress = addr1
    If addr2 <> "" Then fullAddress = fullAddress & ", " & addr2
    If city <> "" Then fullAddress = fullAddress & ", " & city
    If pin <> "" Then fullAddress = fullAddress & " - " & pin
    If state <> "" Then fullAddress = fullAddress & ", " & state
    If country <> "" Then fullAddress = fullAddress & ", " & country

    lblAddress.caption = fullAddress
    lblMobile.caption = mobile

    If swVer <> "" Then lblVersion.caption = swVer Else lblVersion.caption = "v1.0"
    If devName <> "" Then lblDeveloper.caption = devName Else lblDeveloper.caption = "Developer"
    If devMobile <> "" Then lblHelpLine.caption = devMobile Else lblHelpLine.caption = "Helpline"

    On Error GoTo 0
    Exit Sub

SetDefaults:
    lblTitle.caption = "GLOBAL SOFT"
    lblSoftware.caption = "Warranty & Service Management"
    lblAddress.caption = "Company Address"
    lblMobile.caption = "Mobile"
    lblCompany.caption = "Company Name"
    lblVersion.caption = "v1.0"
    lblDeveloper.caption = "Developer"
    lblHelpLine.caption = "Helpline"
End Sub

'========================================
' LOGO DRAG
'========================================
Private Sub imgLogo_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)
    Cancel = True
    Effect = fmDropEffectNone
End Sub
Private Sub btnDiagnostic_Click()
    frmDiagnostic.Show vbModal
End Sub
'========================================
' SETUP LISTVIEWS - 5 SECTIONS
'========================================
Private Sub SetupListViews()
    Dim i As Integer
    Dim listViews(1 To 5) As Object  ' <--- Changed from 4 to 5

    Set listViews(1) = lstActiveEntries      ' <--- NEW: Active Entries
    Set listViews(2) = lstAssignPending1
    Set listViews(3) = lstVendorPending1
    Set listViews(4) = lstEngineerWork1
    Set listViews(5) = lstReadyDelivery1

    For i = 1 To 5  ' <--- Changed from 4 to 5
        With listViews(i)
            .ListItems.Clear
            .ColumnHeaders.Clear
            .View = lvwReport
            .FullRowSelect = True
            .Gridlines = True

            .ColumnHeaders.Add , , "EntryID", 70
            .ColumnHeaders.Add , , "Type", 80
            .ColumnHeaders.Add , , "Status", 100
            .ColumnHeaders.Add , , "Customer Name", 140
            .ColumnHeaders.Add , , "Mobile", 90
            .ColumnHeaders.Add , , "Product", 120
            .ColumnHeaders.Add , , "Company", 100
            .ColumnHeaders.Add , , "Model", 100
            .ColumnHeaders.Add , , "Serial", 120
            .ColumnHeaders.Add , , "Entry Date", 90   ' <--- Assign Date ki jagah Entry Date
            .ColumnHeaders.Add , , "Days", 60          ' <--- Days since entry created
        End With
    Next i
End Sub
'========================================
' DASHBOARD REFRESH - 5 SECTIONS
'========================================
Public Sub RefreshDashboard()
    On Error GoTo ErrorHandler
    
    lstActiveEntries.ListItems.Clear
    lstAssignPending1.ListItems.Clear
    lstVendorPending1.ListItems.Clear
    lstEngineerWork1.ListItems.Clear
    lstReadyDelivery1.ListItems.Clear

    ' "Active Entries" ? "Ready to Assign"
    LoadListViewByStatus lstActiveEntries, "ACTIVE", "Ready to Assign"
    LoadListViewByStatus lstAssignPending1, "ASSIGNED", "Assign Pending"
    LoadListViewByStatus lstVendorPending1, "ASSIGNED", "Vendor Pending"
    LoadListViewByStatus lstEngineerWork1, "ASSIGNED", "Engineer Work"
    LoadListViewByStatus lstReadyDelivery1, "READY", "Ready Delivery"
    
    Exit Sub
ErrorHandler:
    MsgBox "Dashboard Error: " & Err.Description, vbCritical, "Error"
End Sub
'========================================
' LOAD LISTVIEW - UPDATED FOR ACTIVE ENTRIES
'========================================
Private Sub LoadListViewByStatus(lv As Object, statusFilter As String, sectionName As String)
    On Error Resume Next
    Dim wsEntry As Worksheet, wsCust As Worksheet, wsAssign As Worksheet
    Dim lastRowEntry As Long, lastRowCust As Long, lastRowAssign As Long
    Dim i As Long, j As Long, k As Long   ' <--- j ADDED HERE!
    Dim li As Object
    Dim entryID As String, entryType As String, status As String
    Dim custID As String, custName As String, mobile As String
    Dim product As String, company As String, model As String, serial As String
    Dim entryDate As String, daysPending As Long
    Dim count As Long

    Set wsEntry = ThisWorkbook.Sheets("Job_Product")
    Set wsCust = ThisWorkbook.Sheets("Customer_Master")
    Set wsAssign = ThisWorkbook.Sheets("Assign_Master")

    lastRowEntry = wsEntry.Cells(wsEntry.Rows.count, 1).End(xlUp).row
    lastRowCust = wsCust.Cells(wsCust.Rows.count, 1).End(xlUp).row
    lastRowAssign = wsAssign.Cells(wsAssign.Rows.count, 1).End(xlUp).row

    count = 0

    For i = 2 To lastRowEntry
        entryID = Trim(wsEntry.Cells(i, 1).value)
        entryType = UCase(Trim(wsEntry.Cells(i, 3).value))
        status = UCase(Trim(wsEntry.Cells(i, 4).value))
        
        ' ==========================================
' SECTION FILTER LOGIC (SIRF FILTER - NO COLOR)
' ==========================================
Select Case sectionName
    
    Case "Ready to Assign"
        ' === Only ACTIVE status (not assigned yet) ===
        If status <> "ACTIVE" Then GoTo NextRow
        
    Case "Assign Pending"
        ' Any type that is ASSIGNED
        If status <> "ASSIGNED" Then GoTo NextRow
        
    Case "Vendor Pending"
        ' Only WARRANTY + ASSIGNED (need to send to vendor)
        If status <> "ASSIGNED" Then GoTo NextRow
        If entryType <> "WARRANTY" Then GoTo NextRow
        
    Case "Engineer Work"
        ' Only SERVICE + ASSIGNED (engineer working)
        If status <> "ASSIGNED" Then GoTo NextRow
        If entryType <> "SERVICE" Then GoTo NextRow
        
    Case "Ready Delivery"
        ' Ready to deliver
        If status <> "COMPLETED" And status <> "READY" Then GoTo NextRow
        
End Select

        If entryID = "" Then GoTo NextRow

        ' ==========================================
        ' CUSTOMER LOOKUP (Fixed for swapped columns)
        ' ==========================================
        custID = Trim(wsEntry.Cells(i, 2).value)
        custName = "": mobile = ""
        
        For k = 2 To lastRowCust
            If UCase(Trim(wsCust.Cells(k, 1).value)) = UCase(custID) Then
                ' Auto-detect: Col 2 = Name, Col 3 = Mobile (current swapped)
                custName = Trim(wsCust.Cells(k, 2).value)   ' B: Name
                mobile = Trim(wsCust.Cells(k, 3).value)      ' C: Mobile
                Exit For
            End If
        Next k

        product = Trim(wsEntry.Cells(i, 6).value)
        company = Trim(wsEntry.Cells(i, 7).value)
        model = Trim(wsEntry.Cells(i, 8).value)
        serial = Trim(wsEntry.Cells(i, 9).value)

        ' ==========================================
        ' DATE & DAYS CALCULATION
        ' ==========================================
        Select Case sectionName
            
            Case "Active Entries"
                ' Entry Date from Job_Product Col 5 (Entry Date/Time)
                entryDate = wsEntry.Cells(i, 5).value
                If IsDate(entryDate) Then
                    daysPending = DateDiff("d", CDate(entryDate), Date)
                Else
                    daysPending = 0
                    entryDate = ""
                End If
                
            Case Else
                ' Assign Date from Assign_Master
                entryDate = ""
                For j = 2 To lastRowAssign
                    If UCase(Trim(wsAssign.Cells(j, 2).value)) = UCase(entryID) Then
                        entryDate = wsAssign.Cells(j, 22).value
                        Exit For
                    End If
                Next j
                
                If IsDate(entryDate) Then
                    daysPending = Date - CDate(entryDate)
                Else
                    daysPending = 0
                End If
                
        End Select

        ' ==========================================
' ADD TO LISTVIEW
' ==========================================
Set li = lv.ListItems.Add(text:=entryID)
li.SubItems(1) = entryType
li.SubItems(2) = status
li.SubItems(3) = custName
li.SubItems(4) = mobile
li.SubItems(5) = product
li.SubItems(6) = company
li.SubItems(7) = model
li.SubItems(8) = serial
li.SubItems(9) = entryDate
li.SubItems(10) = daysPending

' ==========================================
' COLOR CODING (YAHAN AANA CHAHIYE - li create hone ke BAAD)
' ==========================================
Select Case sectionName
    Case "Ready to Assign"
        ' Red if pending too long (>7 days), Orange if >3 days, Green if fresh
        If daysPending > 7 Then
            li.ForeColor = RGB(220, 0, 0)      ' Dark Red - Urgent
            li.Bold = True
        ElseIf daysPending > 3 Then
            li.ForeColor = RGB(255, 140, 0)    ' Orange - Warning
            li.Bold = True
        Else
            li.ForeColor = RGB(0, 150, 0)      ' Green - Fresh
            li.Bold = False
        End If
        
    Case "Assign Pending"
        li.ForeColor = RGB(0, 0, 200): li.Bold = True
        
    Case "Vendor Pending"
        li.ForeColor = RGB(0, 100, 200): li.Bold = True
        
    Case "Engineer Work"
        li.ForeColor = RGB(200, 100, 0): li.Bold = True
        
    Case "Ready Delivery"
        li.ForeColor = RGB(150, 0, 150): li.Bold = True
        
End Select

        count = count + 1
NextRow:
    Next i

    UpdateCountLabel sectionName, count
End Sub
'========================================
' UPDATE COUNT LABEL - 5 SECTIONS
'========================================
Private Sub UpdateCountLabel(sectionName As String, count As Long)
    On Error Resume Next
    Select Case sectionName
        Case "Ready to Assign": lblActiveEntries.caption = "Ready to Assign (" & count & ")"  ' <--- CHANGED
        Case "Assign Pending": lblAssignPending.caption = "Assign Pending (" & count & ")"
        Case "Vendor Pending": lblVendorPending.caption = "Vendor Pending (" & count & ")"
        Case "Engineer Work": lblEngineerWork.caption = "Engineer Work (" & count & ")"
        Case "Ready Delivery": lblReadyDelivery.caption = "Ready Delivery (" & count & ")"
    End Select
End Sub
'========================================
' ACTIVE ENTRIES DOUBLE CLICK
'========================================
Private Sub lstActiveEntries_DblClick()
    If lstActiveEntries.selectedItem Is Nothing Then Exit Sub
    
    Dim entryID As String
    entryID = lstActiveEntries.selectedItem.text
    
    ' Open Entry Wizard to edit/assign this active entry
    Me.Hide
    frmEntryWizard.Tag = entryID
    frmEntryWizard.Show vbModal
    Me.Show
    RefreshDashboard
End Sub
