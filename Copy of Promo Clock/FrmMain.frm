VERSION 5.00
Object = "{27395F88-0C0C-101B-A3C9-08002B2F49FB}#1.1#0"; "PICCLP32.OCX"
Begin VB.Form FrmMain 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2760
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7680
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2760
   ScaleWidth      =   7680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox PicHook 
      Height          =   255
      Left            =   5580
      MousePointer    =   1  'Arrow
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   4
      Top             =   2100
      Visible         =   0   'False
      Width           =   255
   End
   Begin PicClip.PictureClip StdClip 
      Left            =   2220
      Top             =   1800
      _ExtentX        =   5133
      _ExtentY        =   3175
      _Version        =   393216
      Rows            =   4
      Cols            =   2
      Picture         =   "FrmMain.frx":23D2
   End
   Begin VB.Timer Timer 
      Interval        =   1
      Left            =   6960
      Top             =   2340
   End
   Begin VB.PictureBox PicDisplay 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   675
      Left            =   0
      ScaleHeight     =   45
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   345
      TabIndex        =   3
      Top             =   0
      Width           =   5175
   End
   Begin VB.PictureBox PicCanvas 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   1815
      TabIndex        =   2
      Top             =   2040
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.PictureBox PicAlarm 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   0
      Picture         =   "FrmMain.frx":135E4
      ScaleHeight     =   27
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   1
      Top             =   1560
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.PictureBox PicSkin 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   720
      Left            =   0
      Picture         =   "FrmMain.frx":14556
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   462
      TabIndex        =   0
      Top             =   780
      Visible         =   0   'False
      Width           =   6930
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    

Private TmrStep As String, MStep As String

Private Sub Form_Load()
    
    SnoozeTime = 10
    
    If App.PrevInstance Then
        End
    End If

    TrayI.cbSize = Len(TrayI)
    'Set the window's handle
    TrayI.hWnd = PicHook.hWnd
    'Application-defined identifier of the taskbar icon
    TrayI.uId = 1&
    'Set the flags
    TrayI.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    'Set the callback message
    TrayI.ucallbackMessage = WM_LBUTTONDOWN
    TrayI.hIcon = FrmSettings.Icon
    TrayI.szTip = "Desktop Clock" & Chr$(0)
    'Create the icon
    Shell_NotifyIcon NIM_ADD, TrayI

    MStep = Time$
    Call sOperatingSystemString
    
    If OSBase <> "NT" Then 'Obase = NT Or 9x\Me
        Dim Msg As String

        TrayI.cbSize = Len(TrayI)
        TrayI.hWnd = FrmMain.PicHook.hWnd
        TrayI.uId = 1&
        Shell_NotifyIcon NIM_DELETE, TrayI
        
        Msg = "Desktop Clock Utilizes Functions Available Only On NT Operationg Systems" & vbCrLf
        Msg = Msg & "Such as Windows NT, 2000 And XP. Desktop Clock Will Now Terminate." & vbCrLf & vbCrLf
        Msg = Msg & "PCS Apologizes For Any Inconvenience Caused."
        MsgBox Msg, vbOKOnly + vbExclamation
        
        End
    Else
        REGKEY = "HKEY_LOCAL_MACHINE\Software\PCS\Promotional Software\Desktop Clock\Settings\"
        GetRegSettings
        ApplySettings
        NewSkin
    End If
    
    AlarmSetup
    ReDrawClock
    
End Sub

Private Sub PicDisplay_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Select Case Button
        Case 1
            ReleaseCapture
            SendMessage hWnd, WM_NCLBUTTONDOWN, HTCAPTION, ByVal 0&
        Case 2
            SetWindowPos FrmMain.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
            FrmSettings.MnuHide.Visible = False
            PopupMenu FrmSettings.MnuMenu
    End Select
End Sub

Private Sub PicHook_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Dim Msg As Long

    Msg = X / Screen.TwipsPerPixelX
    If Msg = WM_LBUTTONDBLCLK Then
        'Left button double click
        FrmSettings.Show
        FrmSettings.SetFocus
    ElseIf Msg = WM_RBUTTONUP Then
        'Right button click
        FrmSettings.MnuHide.Visible = True
        PopupMenu FrmSettings.MnuMenu
    End If
    
End Sub

Private Sub Timer_Timer()
    
    Dim AlarmSet As String
    
    'If Set Make Sure The Clock Is Always TopMost
    If Attribs.Clock.AlwaysOntop = True And FrmSettings.MnuMenu.Visible = False And MaskPrepared = True Then
        SetWindowPos FrmMain.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
    Else
        SetWindowPos FrmMain.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
    End If
    
    'The Most Important Requirement Is To Have The Clocks Mask Drawn,
    'Only Continue With The Rest Of The Routine If The Mask Is Prepared
    If MaskPrepared = True And TmrStep <> Time$ Then
        TmrStep = Time$
        
            TrayI.szTip = PicDisplay.ToolTipText & Chr$(0)
            'Update SysTray icon
            Shell_NotifyIcon NIM_MODIFY, TrayI
        
        'If The Alarm Time Has Been Reached And The Alarm Is On
        If Time$ = Attribs.Alarm.AlarmTime And Attribs.Alarm.AlarmActivated = False And Attribs.Alarm.AlarmSet = True Then
            Attribs.Alarm.AlarmActivated = True
            Attribs.Alarm.AlarmSet = False
            
            'The AlarmSet Is Set To False, Save The Settings
            SetRegSettings
            
            If Attribs.Alarm.AlarmType = 0 Or Attribs.Alarm.AlarmType = 2 Then
                'The Alarm Is Visual Or Both (AV)
                FrmAlarm.Show
            Else
                'The Alarm Is Only Audio Which Can Be Stopped With The Now Visible
                'Alarm Menu On The Clock
                MCISend "Close All", Me
                MCISend "setaudio CH1 volume to " & Attribs.Alarm.AlarmVolume, Me
                MCISend "Open """ & Attribs.Alarm.AlarmSound & """ Alias CH1", Me
                MCISend "Play CH1 From 0 Repeat", Me
            End If
            'Show The Alarm Controls In The Menu
            FrmSettings.MnuAlarm.Visible = True
            FrmSettings.spc2.Visible = True
        End If
        
        'Update The Remain Display Every Minute
        If Minute(MStep) <> Minute(Time$) Then
            MStep = Time$
            AlarmSetup
        End If
        
        'Information Regarding The ToolTipText
        If Attribs.Alarm.AlarmSet = True Then
            AlarmSet = " : Alarm On. Remain(" & Remain & ")"
        Else
            AlarmSet = " : Alarm Off"
        End If
        PicDisplay.ToolTipText = WeekdayName(Weekday(Date$, vbMonday), False, vbMonday) & " " & Day(Now) & " " & MonthName(Month(Now), False) & " " & Year(Now) & AlarmSet
        
        ReDrawClock
        
        If Me.Visible = False Then Me.Visible = True
        
    ElseIf MaskPrepared = False Then
        DrawTime
    End If

End Sub
