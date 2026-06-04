Attribute VB_Name = "modSystem"
Option Explicit


Public CurrentEntryID As String

'=================== ID GENERATION ===================

Function GenerateEntryID(entryType As String) As String
    Dim ws As Worksheet, lastRow As Long, nextNum As Long
    Set ws = Sheets("Job_Master")
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    
    If lastRow < 2 Then nextNum = 1 Else nextNum = lastRow
    
    If entryType = "WARRANTY" Then
        GenerateEntryID = "WAR" & Format(nextNum, "00000")
    Else
        GenerateEntryID = "SER" & Format(nextNum, "00000")
    End If
End Function

Function GetNextCustomerID() As String
    Dim ws As Worksheet, lastRow As Long, num As Long
    Set ws = Sheets("Customer_Master")
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    
    If lastRow < 2 Then
        GetNextCustomerID = "CUST0001"
    Else
        num = val(Mid(ws.Cells(lastRow, 1).value, 5)) + 1
        GetNextCustomerID = "CUST" & Format(num, "00000")
    End If
End Function

'=================== OTP SYSTEM ===================

Function MobileExists(mobile As String) As Boolean
    Dim ws As Worksheet, f As Range
    Set ws = Sheets("Customer_Master")
    Set f = ws.Columns(2).Find(What:=mobile, LookIn:=xlValues, LookAt:=xlWhole)
    MobileExists = Not f Is Nothing
End Function





'=================== PHOTO HANDLING ===================

Public Sub LoadCustomerPhoto(mobile As String)
    Dim photoPath As String
    On Error Resume Next
    photoPath = ThisWorkbook.path & "\GLOBAL_SOFT_DATA\Images\Customers\" & mobile & ".jpg"
    If Dir(photoPath) <> "" Then
        'frmEntryWizard.imgCustomerPhoto.Picture = LoadPicture(photoPath)
    Else
        'frmEntryWizard.imgCustomerPhoto.Picture = LoadPicture()
    End If
    On Error GoTo 0
End Sub

Function GetLatestCameraPhoto() As String
    Dim fso As Object, folder As Object, file As Object
    Dim LatestFile As Object, cameraPath As String
    
    cameraPath = Environ("USERPROFILE") & "\Pictures\Camera Roll"
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(cameraPath) Then Exit Function
    
    Set folder = fso.GetFolder(cameraPath)
    For Each file In folder.files
        If LatestFile Is Nothing Then
            Set LatestFile = file
        ElseIf file.DateLastModified > LatestFile.DateLastModified Then
            Set LatestFile = file
        End If
    Next
    
    If Not LatestFile Is Nothing Then GetLatestCameraPhoto = LatestFile.path
End Function

'=================== ACCESSORY MANAGEMENT ===================

Public Sub LoadAccessoriesData(entryID As String)
    Dim ws As Worksheet, lastRow As Long, i As Long
    Set ws = ThisWorkbook.Sheets("Job_Accessory")
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    
    For i = 2 To lastRow
        If ws.Cells(i, 1).value = entryID Then
            'Data load ??? (Forms ?? ??? Connect ?????)
            Debug.Print "Accessory Found: " & ws.Cells(i, 2).value
        End If
    Next i
End Sub

Public Sub SaveAccessoryPhotos(ByVal entryID As String)
    Dim basePath As String, folderPath As String
    Dim i As Integer
    
    basePath = ThisWorkbook.path & "\GLOBAL_SOFT_DATA\Images\Accessories\"
    folderPath = basePath & entryID & "\"
    
    If Dir(basePath, vbDirectory) = "" Then MkDir basePath
    If Dir(folderPath, vbDirectory) = "" Then MkDir folderPath
    
    'Photos Save ??? (???? Form Controls ?? Logic Add ?????)
    LogAction "Accessory Photos Saved: " & entryID
End Sub

'=================== DATA CHECK FUNCTIONS ===================

Function EntryExists(entryID As String) As Boolean
    Dim ws As Worksheet, f As Range
    Set ws = Sheets("Job_Master")
    Set f = ws.Columns(1).Find(entryID, LookIn:=xlValues, LookAt:=xlWhole)
    EntryExists = Not f Is Nothing
End Function

Function CustomerExists(customerID As String) As Boolean
    Dim ws As Worksheet, f As Range
    Set ws = Sheets("Customer_Master")
    Set f = ws.Columns(1).Find(customerID, LookIn:=xlValues, LookAt:=xlWhole)
    CustomerExists = Not f Is Nothing
End Function

Function CustomerHasJobs(customerID As String) As Boolean
    Dim ws As Worksheet, f As Range
    Set ws = Sheets("Job_Master")
    Set f = ws.Columns(2).Find(customerID, LookIn:=xlValues, LookAt:=xlWhole)
    CustomerHasJobs = Not f Is Nothing
End Function

'=================== UTILITY FUNCTIONS ===================

Sub MarkDeleted(entryID As String)
    Dim ws As Worksheet, f As Range
    Set ws = Sheets("Job_Master")
    Set f = ws.Columns(1).Find(entryID, LookIn:=xlValues, LookAt:=xlWhole)
    If Not f Is Nothing Then
        ws.Cells(f.row, 4).value = "DELETED"
        ws.Cells(f.row, 5).value = Now
        LogAction "Marked Deleted: " & entryID
    End If
End Sub

Sub LogAction(msg As String)
    '?? Perfect_Safety_Module ??? ?? ??, ???? Reference ?? ???
    Application.Run "Perfect_Safety_Module.LogAction", msg
End Sub



Public Function FormatMobileNumber(ByVal mob As String) As String
    Dim i As Integer
    Dim ch As String
    Dim result As String
    
    ' Only digits ??????
    For i = 1 To Len(mob)
        ch = Mid(mob, i, 1)
        If ch Like "[0-9]" Then
            result = result & ch
        End If
    Next i
    
    ' ??? 10 digit ?? ?????? ?? ?? last 10 ??
    If Len(result) > 10 Then
        result = Right(result, 10)
    End If
    
    FormatMobileNumber = result
End Function


Public Function GetSMTPServerAddress(ByVal serverName As String) As String
    Select Case Trim(LCase(serverName))
        Case "gmail": GetSMTPServerAddress = "smtp.gmail.com"
        Case "outlook": GetSMTPServerAddress = "smtp.office365.com"
        Case "yahoo": GetSMTPServerAddress = "smtp.mail.yahoo.com"
        Case "hotmail": GetSMTPServerAddress = "smtp.live.com"
        Case Else: GetSMTPServerAddress = serverName
    End Select
End Function
