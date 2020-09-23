VERSION 5.00
Begin VB.Form FMain 
   BackColor       =   &H80000004&
   Caption         =   "Invisible form"
   ClientHeight    =   3255
   ClientLeft      =   3675
   ClientTop       =   5040
   ClientWidth     =   6135
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3255
   ScaleWidth      =   6135
   Begin VB.PictureBox picClock 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   2
      Left            =   1200
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   2
      Top             =   600
      Width           =   375
   End
   Begin VB.PictureBox picClock 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   1
      Left            =   600
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   1
      Top             =   600
      Width           =   375
   End
   Begin VB.PictureBox picClock 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   0
      Left            =   0
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   0
      Top             =   600
      Width           =   375
   End
   Begin BinaryClock.cSysTray cSysTray 
      Index           =   2
      Left            =   1200
      Top             =   0
      _ExtentX        =   900
      _ExtentY        =   900
      InTray          =   0   'False
      TrayIcon        =   "FMain.frx":0000
      TrayTip         =   "VB 5 - SysTray Control."
   End
   Begin BinaryClock.cSysTray cSysTray 
      Index           =   1
      Left            =   600
      Top             =   0
      _ExtentX        =   900
      _ExtentY        =   900
      InTray          =   0   'False
      TrayIcon        =   "FMain.frx":1D0A
      TrayTip         =   "VB 5 - SysTray Control."
   End
   Begin VB.Timer Timer 
      Interval        =   1000
      Left            =   0
      Top             =   1200
   End
   Begin BinaryClock.cSysTray cSysTray 
      Index           =   0
      Left            =   0
      Top             =   0
      _ExtentX        =   900
      _ExtentY        =   900
      InTray          =   0   'False
      TrayIcon        =   "FMain.frx":3A14
      TrayTip         =   "Nerdklok"
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "Menu"
      Begin VB.Menu mnuCancel 
         Caption         =   "&Cancel"
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHours 
         Caption         =   "&Hours"
      End
      Begin VB.Menu mnuMinutes 
         Caption         =   "&Minutes"
      End
      Begin VB.Menu mnuSeconds 
         Caption         =   "&Seconds"
      End
      Begin VB.Menu mnuSep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "FMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function CreateIconIndirect Lib "user32" (ByRef pIconInfo As ICONINFO) As Long
Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
Private Declare Function CreateBitmapIndirect Lib "gdi32" (lpbm As BITMAP) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (ByRef lpOSVERSIONINFO As OSVERSIONINFO) As Long

Private Type ICONINFO
   fIcon As Long
   xHotspot As Long
   yHotspot As Long
   hbmMask As Long
   hbmColor As Long
End Type
Private Type BITMAP
   bmType As Long
   bmWidth As Long
   bmHeight As Long
   bmWidthBytes As Long
   bmPlanes As Integer
   bmBitsPixel As Integer
   lpbmBits As Long
End Type
Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

Private Const VER_PLATFORM_WIN32_NT = 2

Private hIcon0 As Long
Private hIcon1 As Long
Private hIcon2 As Long

Private hMask As Long
Private laMask(0 To 15) As Integer

Private II As ICONINFO
Private bWindowsXP As Boolean
Private bFirstTime As Boolean

Private Sub cSysTray_MouseUp(Index As Integer, Button As Integer, Id As Long)
    If Button = vbRightButton Then _
        PopupMenu mnuPopup, , , , mnuExit
End Sub

Private Sub Form_Load()
    ' Windows XP?
    Dim osv As OSVERSIONINFO
    osv.dwOSVersionInfoSize = Len(osv)
    Call GetVersionEx(osv)
    If osv.dwPlatformId = VER_PLATFORM_WIN32_NT And _
       osv.dwMajorVersion = 5 And _
       osv.dwMinorVersion >= 1 Then ' Windows XP
       
       bWindowsXP = True
       
    End If
    
    cSysTray(0).gTrayId = 0
    cSysTray(1).gTrayId = 1
    cSysTray(2).gTrayId = 2
    
    mnuHours.Checked = CBool(GetSetting("DigitalClock", "Config", "Hours", "True"))
    mnuMinutes.Checked = CBool(GetSetting("DigitalClock", "Config", "Minutes", "True"))
    mnuSeconds.Checked = CBool(GetSetting("DigitalClock", "Config", "Seconds", "True"))

    With picClock(0)
        .Width = Screen.TwipsPerPixelX * 18
        .Height = Screen.TwipsPerPixelY * 18
        .ScaleMode = 3
        .ScaleWidth = 16
        .ScaleHeight = 16
        .AutoRedraw = True
    End With
    With picClock(1)
        .Width = Screen.TwipsPerPixelX * 18
        .Height = Screen.TwipsPerPixelY * 18
        .ScaleMode = 3
        .ScaleWidth = 16
        .ScaleHeight = 16
        .AutoRedraw = True
    End With
    With picClock(2)
        .Width = Screen.TwipsPerPixelX * 18
        .Height = Screen.TwipsPerPixelY * 18
        .ScaleMode = 3
        .ScaleWidth = 16
        .ScaleHeight = 16
        .AutoRedraw = True
    End With
    
    Hide
    
    MaakANDMask
    
    bFirstTime = True
    Timer_Timer
End Sub

Private Sub Form_Unload(Cancel As Integer)
    DestroyIcon hIcon0
    DestroyIcon hIcon1
    DestroyIcon hIcon2
    DeleteObject hMask
    
    SaveSetting "DigitalClock", "Config", "Hours", CStr(mnuHours.Checked)
    SaveSetting "DigitalClock", "Config", "Minutes", CStr(mnuMinutes.Checked)
    SaveSetting "DigitalClock", "Config", "Seconds", CStr(mnuSeconds.Checked)
    
End Sub

Private Sub mnuExit_Click()
    Timer.Enabled = False
    cSysTray(0).InTray = False
    cSysTray(1).InTray = False
    cSysTray(2).InTray = False
    Unload Me
End Sub

Private Sub mnuHours_Click()
    mnuHours.Checked = Not mnuHours.Checked
    If mnuHours.Checked = False Then cSysTray(0).InTray = False
End Sub

Private Sub mnuMinutes_Click()
    mnuMinutes.Checked = Not mnuMinutes.Checked
    If mnuMinutes.Checked = False Then cSysTray(1).InTray = False
End Sub

Private Sub mnuSeconds_Click()
    mnuSeconds.Checked = Not mnuSeconds.Checked
    If mnuSeconds.Checked = False Then cSysTray(2).InTray = False
End Sub

Private Sub Timer_Timer()
    ' Check icons
    If (cSysTray(0).InTray = False And mnuHours.Checked) _
    Or (cSysTray(1).InTray = False And mnuMinutes.Checked) _
    Or (cSysTray(2).InTray = False And mnuSeconds.Checked) Then
        cSysTray(0).InTray = False
        cSysTray(1).InTray = False
        cSysTray(2).InTray = False

        If bWindowsXP Then
            cSysTray(2).InTray = mnuSeconds.Checked
            cSysTray(1).InTray = mnuMinutes.Checked
            cSysTray(0).InTray = mnuHours.Checked
        Else
            cSysTray(0).InTray = mnuHours.Checked
            cSysTray(1).InTray = mnuMinutes.Checked
            cSysTray(2).InTray = mnuSeconds.Checked
        End If
        
        bFirstTime = True
    End If

    Dim hrs As Long
    Dim min As Long
    Dim sec As Long
    
    hrs = Hour(Now)
    min = Minute(Now)
    sec = Second(Now)
    
    ' Draw clock in pictureboxes
    picClock(0).Cls
    picClock(1).Cls
    picClock(2).Cls
    Call DrawDots(0, hrs)
    Call DrawDots(1, min)
    Call DrawDots(2, sec)
    
    ' Change Icons
    If min = 0 Or bFirstTime Then ' Needs only to be repainted once every hour
        With II
            .fIcon = 1
            .hbmColor = picClock(0).Image.Handle
            .hbmMask = hMask
            .xHotspot = 0
            .yHotspot = 0
        End With
        DestroyIcon hIcon0
        hIcon0 = CreateIconIndirect(II)
        cSysTray(0).SetTrayIcon hIcon0
    End If
    
    If sec < 3 Or bFirstTime Then ' Repaint once every minute, the 3 is to
                                  ' make sure a second skip won't prevent a redraw
        With II
            .fIcon = 1
            .hbmColor = picClock(1).Image.Handle
            .hbmMask = hMask
            .xHotspot = 0
            .yHotspot = 0
        End With
        DestroyIcon hIcon1
        hIcon1 = CreateIconIndirect(II)
        cSysTray(1).SetTrayIcon hIcon1
    End If
    
    With II
        .fIcon = 1
        .hbmColor = picClock(2).Image.Handle
        .hbmMask = hMask
        .xHotspot = 0
        .yHotspot = 0
    End With
    DestroyIcon hIcon2
    hIcon2 = CreateIconIndirect(II)
    cSysTray(2).SetTrayIcon hIcon2
    
    ' Set traytip
    cSysTray(0).TrayTip = CStr(hrs) & ":" & Format(min, "00") & ":" & Format(sec, "00")
    cSysTray(1).TrayTip = CStr(hrs) & ":" & Format(min, "00") & ":" & Format(sec, "00")
    cSysTray(2).TrayTip = CStr(hrs) & ":" & Format(min, "00") & ":" & Format(sec, "00")
    
    bFirstTime = False
End Sub

Private Sub DrawDots(Index As Long, Number As Long)
    Dim eerste As Long
    Dim tweede As Long
    Dim offset As Long

    Dim cVol As Long
    Dim cLeeg As Long
    
    cVol = RGB(0, 0, 255)
    cLeeg = RGB(255, 255, 255)

    eerste = Number \ 10
    tweede = Number Mod 10

    offset = 2

    blokje Index, offset, 12, IIf(eerste And 1, cVol, cLeeg)
    blokje Index, offset, 9, IIf(eerste And 2, cVol, cLeeg)
    blokje Index, offset, 6, IIf(eerste And 4, cVol, cLeeg)
    blokje Index, offset, 3, IIf(eerste And 8, cVol, cLeeg)
    
    blokje Index, offset + 8, 12, IIf(tweede And 1, cVol, cLeeg)
    blokje Index, offset + 8, 9, IIf(tweede And 2, cVol, cLeeg)
    blokje Index, offset + 8, 6, IIf(tweede And 4, cVol, cLeeg)
    blokje Index, offset + 8, 3, IIf(tweede And 8, cVol, cLeeg)
    
End Sub

Private Sub blokje(ByVal Index As Long, ByVal x As Long, ByVal y As Long, ByVal c As OLE_COLOR)
    picClock(Index).Line (x, y)-(x + 4, y + 1), c, BF
End Sub

Private Sub MaakANDMask()
    laMask(0) = &HFFFF
    laMask(1) = &HFFFF
    laMask(2) = &HFFFF
    laMask(3) = &HC1C1
    laMask(4) = &HC1C1
    laMask(5) = &HFFFF
    laMask(6) = &HC1C1
    laMask(7) = &HC1C1
    laMask(8) = &HFFFF
    laMask(9) = &HC1C1
    laMask(10) = &HC1C1
    laMask(11) = &HFFFF
    laMask(12) = &HC1C1
    laMask(13) = &HC1C1
    laMask(14) = &HFFFF
    laMask(15) = &HFFFF
        
    Dim bm As BITMAP
    With bm
        .bmBitsPixel = 1
        .bmHeight = 16
        .bmPlanes = 1
        .bmType = 0
        .bmWidth = 16
        .bmWidthBytes = 2
        .lpbmBits = VarPtr(laMask(0))
    End With
    
    hMask = CreateBitmapIndirect(bm)
End Sub
