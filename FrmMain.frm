VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmMain 
   Caption         =   "Wallpaper Rotator"
   ClientHeight    =   1545
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5130
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   103
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   342
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin MSComDlg.CommonDialog CMDlg1 
      Left            =   0
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      Height          =   540
      Left            =   -5000
      Picture         =   "FrmMain.frx":0442
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   2
      Top             =   960
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Copyright 2002 by Michael Barnathan"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   360
      TabIndex        =   1
      Top             =   720
      Width           =   4230
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Wallpaper rotation utility"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   600
      TabIndex        =   0
      Top             =   120
      Width           =   3750
   End
   Begin VB.Menu mnuoptions 
      Caption         =   "Options"
      Visible         =   0   'False
      Begin VB.Menu mnurefresh 
         Caption         =   "&Refresh Wallpaper"
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuday 
         Caption         =   "&Sunday..."
         Index           =   1
      End
      Begin VB.Menu mnuday 
         Caption         =   "&Monday..."
         Index           =   2
      End
      Begin VB.Menu mnuday 
         Caption         =   "&Tuesday..."
         Index           =   3
      End
      Begin VB.Menu mnuday 
         Caption         =   "&Wednesday"
         Index           =   4
      End
      Begin VB.Menu mnuday 
         Caption         =   "T&hursday..."
         Index           =   5
      End
      Begin VB.Menu mnuday 
         Caption         =   "&Friday..."
         Index           =   6
      End
      Begin VB.Menu mnuday 
         Caption         =   "S&aturday..."
         Index           =   7
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuabout 
         Caption         =   "Abou&t"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As String, ByVal fuWinIni As Long) As Long
Const SPIF_UPDATEINIFILE = &H1
Const SPI_SETDESKWALLPAPER = 20
Const SPIF_SENDWININICHANGE = &H2

Dim DayDescription(1 To 7) As String


Private Sub RefreshWallpaper(DayOfWeek)
On Error Resume Next
Dim OldSetting$

OldSetting$ = ""
OldSetting$ = QueryValue(HKEY_CURRENT_USER, "Software\Wallpaper Rotator\Settings", DayDescription(DayOfWeek))
OldSetting$ = Trim$(OldSetting$)

If OldSetting$ = "" Then OldSetting$ = "Error: No data entered for " & DayDescription(DayOfWeek)
If Left$(OldSetting$, 5) = "Error" Then ModifyTray Me.hWnd, Me.Icon, Me.Caption, "Error reading from registry", OldSetting$, 3, 5000: Exit Sub
TestForFileExist = FileLen(OldSetting$)

If TestForFileExist = 0 Or Err > 0 Then ModifyTray Me.hWnd, Me.Icon, Me.Caption, Error$(Err) & " while reading " & OldSetting$, OldSetting$, 3, 5000: Exit Sub

'SystemParametersInfo SPI_SETDESKWALLPAPER, 0&, OldSetting$, SPIF_UPDATEINIFILE Or SPIF_SENDWININICHANGE
SetWallpaper OldSetting$, WPSTYLE_STRETCH, False

ModifyTray Me.hWnd, Me.Icon, Me.Caption, "Wallpaper set", "Your wallpaper has been changed to your " & DayDescription(DayOfWeek) & " preference", 1, 2000
End Sub
Private Sub SetWallpaper(ByVal Filename As String, ByVal Wallpaperstyle As Long, ByVal ForceEnable As Boolean)

  Dim ActiveDesktop1 As ActiveDesktop
  Dim Component1 As COMPONENTSOPT
  Dim Wallpaper1 As WALLPAPEROPT
  
  Set ActiveDesktop1 = New ActiveDesktop
  
  Component1.dwSize = Len(Component1)
  
  ActiveDesktop1.GetDesktopItemOptions Component1, 0&
  If Component1.fActiveDesktop = 0 And fForce Then
    Component1.fActiveDesktop = 1 'Enable Active Desktop if it must be enabled
    ActiveDesktop1.SetDesktopItemOptions Component1, 0&
  End If
  
Wallpaper1.dwSize = Len(Wallpaper1)
Wallpaper1.dwStyle = Wallpaperstyle
  
  ActiveDesktop1.SetWallpaperOptions Wallpaper1, 0&
  ActiveDesktop1.SetWallpaper Filename, 0&
  ActiveDesktop1.ApplyChanges AD_APPLY_ALL

  Set ActiveDesktop1 = Nothing 'Clean up after we're done
  
End Sub

Private Sub Form_Load()
DayDescription(1) = "Sunday"
DayDescription(2) = "Monday"
DayDescription(3) = "Tuesday"
DayDescription(4) = "Wednesday"
DayDescription(5) = "Thursday"
DayDescription(6) = "Friday"
DayDescription(7) = "Saturday"

AddToTray Me.hWnd, Me.Icon, Me.Caption
RefreshWallpaper Weekday(Now)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Response = RespondToTray(X, FrmMain)
If Response = 1 Then Call mnuabout_Click
If Response = 2 Then PopupMenu mnuoptions, , , , mnuabout
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = 0 Then Cancel = True: Me.Hide: Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
RemoveFromTray
End Sub

Private Sub mnuabout_Click()
Me.Show
End Sub

Private Sub mnuday_Click(Index As Integer)
Dim OldSetting$

OldSetting$ = ""
OldSetting$ = QueryValue(HKEY_CURRENT_USER, "Software\Wallpaper Rotator\Settings", DayDescription(Index))
If Left$(OldSetting$, 5) = "Error" Then MsgBox OldSetting$, 48, App.ProductName

CMDlg1.DialogTitle = "Choose wallpaper to use for " & DayDescription(Index)
CMDlg1.Filter = "Image files (*.jpg; *.gif; *.bmp; *.tga; *.*htm*; *.png)|*.jpg; *.gif; *.bmp; *.tga; *.*htm*; *.png|All Files (*.*)|*.*"
CMDlg1.FilterIndex = 1
CMDlg1.Flags = &H80000 Or &H1000 Or &H4 Or &H200000 Or &H800
CMDlg1.CancelError = True
On Error Resume Next
If OldSetting$ <> "" Then CMDlg1.Filename = OldSetting$
CMDlg1.ShowOpen
If Err.Number = cdlCancel Then Exit Sub

On Error Resume Next
CreateNewKey "Software\Wallpaper Rotator\Settings", HKEY_CURRENT_USER
SetKeyValue HKEY_CURRENT_USER, "Software\Wallpaper Rotator\Settings", DayDescription(Index), CMDlg1.Filename, REG_SZ

ModifyTray Me.hWnd, Me.Icon, Me.Caption, DayDescription(Index) & "'s wallpaper changed", "Using " & CMDlg1.Filename, 1
End Sub

Private Sub mnuexit_Click()
Unload Me
End Sub

Private Sub mnurefresh_Click()
RefreshWallpaper Weekday(Now)
End Sub
