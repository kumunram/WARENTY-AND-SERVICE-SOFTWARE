Attribute VB_Name = "modMinMax"
Option Explicit
 
'========================================
' API DECLARATIONS FOR MIN/MAX BUTTONS
'========================================
#If VBA7 Then
    Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As Long) As LongPtr
    Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
    Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
    Private Declare PtrSafe Function GetSystemMenu Lib "user32" (ByVal hwnd As LongPtr, ByVal bRevert As Long) As LongPtr
    Private Declare PtrSafe Function DeleteMenu Lib "user32" (ByVal hMenu As LongPtr, ByVal nPosition As Long, ByVal wFlags As Long) As Long
    Private Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal hwnd As LongPtr) As Long
    Private Declare PtrSafe Function SetWindowPos Lib "user32" (ByVal hwnd As LongPtr, ByVal hWndInsertAfter As LongPtr, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
#Else
    Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
    Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
    Private Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
    Private Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
#End If
 
Private Const GWL_STYLE = (-16)
Private Const WS_MINIMIZEBOX = &H20000
Private Const WS_MAXIMIZEBOX = &H10000
Private Const WS_SYSMENU = &H80000
Private Const WS_THICKFRAME = &H40000
Private Const MF_BYCOMMAND = &H0&
Private Const SC_CLOSE = &HF060
Private Const HWND_TOP = 0
Private Const SWP_FRAMECHANGED = &H20
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOZORDER = &H4
Private Const SWP_NOOWNERZORDER = &H200
Private Const SWP_NOACTIVATE = &H10
 
'========================================
' ADD MIN/MAX BUTTONS TO ANY USERFORM
' FIXED: Removed SWP_SHOWWINDOW (was causing premature show)
'========================================
Public Sub AddMinMaxButtons(frm As Object)
    On Error Resume Next
 
    #If VBA7 Then
        Dim hwnd As LongPtr
        Dim lStyle As LongPtr
    #Else
        Dim hwnd As Long
        Dim lStyle As Long
    #End If
 
    ' Find the UserForm window handle
    hwnd = FindWindow("ThunderDFrame", frm.caption)
    If hwnd = 0 Then hwnd = FindWindow("ThunderDFrame", vbNullString)
    If hwnd = 0 Then hwnd = FindWindow("ThunderXFrame", frm.caption)
    If hwnd = 0 Then hwnd = FindWindow("ThunderXFrame", vbNullString)
 
    ' If form window not yet created (called too early), just exit cleanly
    If hwnd = 0 Then Exit Sub
 
    ' Set min/max style flags
    lStyle = GetWindowLong(hwnd, GWL_STYLE)
    lStyle = lStyle Or WS_MINIMIZEBOX
    lStyle = lStyle Or WS_MAXIMIZEBOX
    lStyle = lStyle Or WS_SYSMENU
    lStyle = lStyle Or WS_THICKFRAME
    SetWindowLong hwnd, GWL_STYLE, lStyle
 
    ' Redraw frame WITHOUT forcing show (SWP_NOACTIVATE instead of SWP_SHOWWINDOW)
    SetWindowPos hwnd, HWND_TOP, 0, 0, 0, 0, _
        SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOZORDER Or _
        SWP_NOOWNERZORDER Or SWP_FRAMECHANGED Or SWP_NOACTIVATE
 
    On Error GoTo 0
End Sub
 
'========================================
' REMOVE CLOSE BUTTON (optional helper)
'========================================
Public Sub RemoveCloseButton(frm As Object)
    On Error Resume Next
 
    #If VBA7 Then
        Dim hwnd As LongPtr
        Dim hMenu As LongPtr
    #Else
        Dim hwnd As Long
        Dim hMenu As Long
    #End If
 
    hwnd = FindWindow("ThunderDFrame", frm.caption)
    If hwnd = 0 Then hwnd = FindWindow("ThunderXFrame", frm.caption)
    If hwnd = 0 Then Exit Sub
 
    hMenu = GetSystemMenu(hwnd, 0)
    If hMenu = 0 Then Exit Sub
 
    DeleteMenu hMenu, SC_CLOSE, MF_BYCOMMAND
    DrawMenuBar hwnd
 
    On Error GoTo 0
End Sub




