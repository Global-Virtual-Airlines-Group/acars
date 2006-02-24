Attribute VB_Name = "SystemTray"
Option Explicit

'User-defined variable to pass to the Shell_NotiyIcon function
Private Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uId As Long
    uFlags As Long
    uCallBackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

'Constants for the Shell_NotifyIcon function
Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2

Private Const WM_MOUSEMOVE = &H200

Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4

Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_RBUTTONDBLCLK = &H206
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205

Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_APPWINDOW = &H40000

Private Const SW_HIDE = 0
Private Const SW_SHOW = 5

'Declare the API function calls
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" _
    (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
    
Dim nid As NOTIFYICONDATA

'Taskbar manipulation constants
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
    (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
    (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Public Sub AddIcon(ByVal ToolTip As String, IconData As Long)
    On Error GoTo ErrorHandler
    
    'Add icon to system tray
    With nid
        .cbSize = Len(nid)
        .hWnd = frmMain.hWnd
        .uId = vbNull
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uCallBackMessage = WM_MOUSEMOVE
        .hIcon = IconData
        .szTip = ToolTip & vbNullChar
    End With
    
    Call Shell_NotifyIcon(NIM_ADD, nid)
    
ExitSub:
    Exit Sub
ErrorHandler:
    Screen.MousePointer = vbDefault
    Resume ExitSub

End Sub

Public Sub RemoveIcon()
    Call Shell_NotifyIcon(NIM_DELETE, nid)
End Sub

Public Sub ModifyIcon(ByVal ToolTip As String, IconData As Long)
    On Error GoTo ErrorHandler
    
    'Update icon in system tray
    With nid
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .hIcon = IconData
        .szTip = ToolTip & vbNullChar
    End With
    
    Call Shell_NotifyIcon(NIM_MODIFY, nid)
    
ExitSub:
    Exit Sub
ErrorHandler:
    Screen.MousePointer = vbDefault
    Resume ExitSub

End Sub

Public Sub TaskBarHide(ByVal id As Long)
    ShowWindow id, SW_HIDE
    SetWindowLong id, GWL_EXSTYLE, (GetWindowLong(id, GWL_EXSTYLE) And (Not WS_EX_APPWINDOW))
    ShowWindow id, SW_SHOW
End Sub

Public Sub TaskBarShow(ByVal id As Long)
    ShowWindow id, SW_HIDE
    SetWindowLong id, GWL_EXSTYLE, (GetWindowLong(id, GWL_EXSTYLE) Or WS_EX_APPWINDOW)
    ShowWindow id, SW_SHOW
End Sub
