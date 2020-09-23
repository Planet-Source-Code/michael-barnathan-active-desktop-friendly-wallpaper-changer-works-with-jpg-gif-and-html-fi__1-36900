Attribute VB_Name = "TrayModule"
 Public Type NOTIFYICONDATA
   cbSize As Long
   hWnd As Long
   uId As Long
   uFlags As Long
   uCallBackMessage As Long
   hIcon As Long
   szTip As String * 128
   dwState As Long
   dwStateMask As Long
   szInfo As String * 256
   uTimeout As Long
   szInfoTitle As String * 64
   dwInfoFlags As Long
End Type

Public Const NOTIFYICON_VERSION = 3
Public Const NOTIFYICON_OLDVERSION = 0
Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const NIM_SETFOCUS = &H3
Public Const NIM_SETVERSION = &H4
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4
Public Const NIF_STATE = &H8
Public Const NIF_INFO = &H10
Public Const NIS_HIDDEN = &H1
Public Const NIS_SHAREDICON = &H2
Public Const NIIF_NONE = &H0
Public Const NIIF_WARNING = &H2
Public Const NIIF_ERROR = &H3
Public Const NIIF_INFO = &H1
Public Const NIIF_GUID = &H4

Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_RBUTTONDBLCLK = &H206
 
Public Declare Function SetForegroundWindow Lib "user32" _
(ByVal hWnd As Long) As Long
Public Declare Function Shell_NotifyIcon Lib "shell32" _
Alias "Shell_NotifyIconA" _
(ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

Public TrayIcon As NOTIFYICONDATA
Function RespondToTray(X As Single, Control As Object)
          'Call this sub from the mousemove event on a form
          'Event occurs when the mouse pointer is within the rectangular
          'boundaries of the icon in the taskbar status area.
          RespondToTray = 0
          Dim msg As Long
          Dim sFilter As String
          If Control.ScaleMode <> 3 Then msg = X / Screen.TwipsPerPixelX Else: msg = X
          msg = X
          Select Case msg
             Case WM_LBUTTONDOWN
             Case WM_LBUTTONUP
             Case WM_LBUTTONDBLCLK 'Left button double-clicked
             RespondToTray = 1
             Case WM_RBUTTONDOWN 'Right button pressed
             RespondToTray = 2
             Case WM_RBUTTONUP
             Case WM_RBUTTONDBLCLK
          End Select
End Function

Sub AddToTray(TrayHandle As Long, TrayIconImage As Long, TrayText As String, Optional BalloonTitle As String, Optional BalloonText As String, Optional BalloonIcon As Long, Optional TrayTimeOut As Long = 3000)
        'Set the individual values of the NOTIFYICONDATA data type.
        With TrayIcon
        .cbSize = Len(TrayIcon)
        If TrayHandle <> 0 Then
        .hWnd = TrayHandle
        End If
        .uId = vbNull
        If BalloonTitle <> "" Then
        .uFlags = NIF_ICON Or NIF_INFO Or NIF_MESSAGE
        Else
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        End If
        .uCallBackMessage = WM_MOUSEMOVE
        If TrayIconImage <> 0 Then
        .hIcon = TrayIconImage
        End If
        If TrayText <> "" Then
        .szTip = TrayText & vbNullChar
        End If
        .dwState = 0
        .dwStateMask = 0
        If BalloonTitle <> "" Or BalloonText <> "" Or BalloonIcon <> 0 Then
        If BalloonText <> "" Then
        .szInfo = BalloonText & Chr(0)
        End If
        If BalloonTitle <> "" Then
        .szInfoTitle = BalloonTitle & Chr(0)
        End If
        If BalloonIcon <> 0 Then
        .dwInfoFlags = BalloonIcon
        End If
        If TrayTimeOut > 0 Then
        .uTimeout = TrayTimeOut
        End If
        End If
        End With
  'Call the Shell_NotifyIcon function to add the icon to the taskbar
  'status area.
   Shell_NotifyIcon NIM_ADD, TrayIcon
End Sub

Sub ModifyTray(Optional TrayHandle As Long, Optional TrayIconImage As Long, Optional TrayText As String, Optional BalloonTitle As String, Optional BalloonText As String, Optional BalloonIcon As Long, Optional TrayTimeOut As Long)
        With TrayIcon
        .cbSize = Len(TrayIcon)
        If TrayHandle <> 0 Then
        .hWnd = TrayHandle
        End If
        .uId = vbNull
        If BalloonTitle <> "" Then
        .uFlags = NIF_ICON Or NIF_INFO Or NIF_MESSAGE
        End If
        .uCallBackMessage = WM_MOUSEMOVE
        .hIcon = TrayIconImage
        If TrayText <> "" Then
        .szTip = TrayText & vbNullChar
        End If
        .dwState = 0
        .dwStateMask = 0
        If BalloonTitle <> "" Or BalloonText <> "" Or BalloonIcon <> 0 Then
        If BalloonText <> "" Then
        .szInfo = BalloonText & Chr(0)
        End If
        If BalloonTitle <> "" Then
        .szInfoTitle = BalloonTitle & Chr(0)
        End If
        If BalloonIcon <> 0 Then
        .dwInfoFlags = BalloonIcon
        End If
        If TrayTimeOut > 0 Then
        .uTimeout = TrayTimeOut
        End If
        End If
         End With
   Shell_NotifyIcon NIM_MODIFY, TrayIcon
End Sub

Sub RemoveFromTray()
'If not called when the program execution ends, the tray icon will "ghost"
Shell_NotifyIcon NIM_DELETE, TrayIcon
End Sub
Sub ShowFormAgain(TrayForm As Form)
RemoveFromTray
TrayForm.Show
End Sub
