VERSION 5.00
Object = "{27395F88-0C0C-101B-A3C9-08002B2F49FB}#1.1#0"; "PICCLP32.OCX"
Begin VB.Form FrmAlarm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   2115
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6150
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2115
   ScaleWidth      =   6150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin PicClip.PictureClip AlarmClip 
      Left            =   60
      Top             =   2160
      _ExtentX        =   5133
      _ExtentY        =   1588
      _Version        =   393216
      Rows            =   2
      Cols            =   2
      Picture         =   "FrmAlarm.frx":0000
   End
   Begin VB.PictureBox StdBtnSnooze 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Left            =   3060
      Picture         =   "FrmAlarm.frx":8932
      ScaleHeight     =   450
      ScaleWidth      =   1455
      TabIndex        =   2
      Tag             =   "0\1"
      ToolTipText     =   "Snooze"
      Top             =   1620
      Width           =   1455
   End
   Begin VB.PictureBox StdBtnStop 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Left            =   4620
      Picture         =   "FrmAlarm.frx":ABAC
      ScaleHeight     =   450
      ScaleWidth      =   1455
      TabIndex        =   1
      Tag             =   "2\3"
      ToolTipText     =   "Stop Alarm"
      Top             =   1620
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Interval        =   75
      Left            =   4140
      Top             =   0
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2010
      Left            =   60
      Picture         =   "FrmAlarm.frx":CE26
      ScaleHeight     =   2010
      ScaleWidth      =   1905
      TabIndex        =   0
      Top             =   60
      Width           =   1905
   End
   Begin VB.Label LblMessage 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "What Time Is This Time? The Chosen Time To Adhere To."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   1155
      Left            =   2100
      TabIndex        =   3
      Top             =   360
      Width           =   3915
   End
End
Attribute VB_Name = "FrmAlarm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private IX As Single
Private IY As Single

Private Sub Form_Activate()
    
    Me.Caption = "Desktop Alarm: " & Attribs.Alarm.AlarmTime
    
    'If The Alarm Type Is AV(2) Then Play The Alarm Sound
    If Attribs.Alarm.AlarmType = 2 Then
        MCISend "Close All", Me
        MCISend "Open """ & Attribs.Alarm.AlarmSound & """ Alias CH1", Me
        MCISend "setaudio CH1 volume to " & Attribs.Alarm.AlarmVolume, Me
        MCISend "Play CH1 From 0 Repeat", Me
    End If
    
    StdBtnSnooze.ToolTipText = "Snooze Alarm For " & SnoozeTime & " Minutes."
    
    LblMessage.Caption = Attribs.Alarm.AlarmNote
    
End Sub

Private Sub Form_Load()
    
    Randomize Timer
    
    IX = Picture1.Left
    IY = Picture1.Left
    
    'Set Button Images
    DoButton StdBtnSnooze, AlarmClip, False
    DoButton StdBtnStop, AlarmClip, False
    
    'Skin The Buttons
    SkinPicture StdBtnSnooze, StdBtnSnooze
    SkinPicture StdBtnStop, StdBtnStop
    
End Sub

Private Sub StdBtnSnooze_Click()
    
    MCISend "Close All", Me
    FrmSettings.MnuAlarm.Visible = False
    FrmSettings.spc2.Visible = False
    
    Attribs.Alarm.AlarmActivated = False
    
    Attribs.Alarm.AlarmTime = Snooze(Attribs.Alarm.AlarmTime, SnoozeTime)
    Attribs.Alarm.AlarmSet = True
    SetRegSettings
    
    AlarmSetup
    If FrmAlarm.Visible = True Then FrmAlarm.Visible = False
    
End Sub

Private Sub StdBtnSnooze_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    DoButton StdBtnSnooze, AlarmClip, True
    
End Sub

Private Sub StdBtnSnooze_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    DoButton StdBtnSnooze, AlarmClip, False
    
End Sub

Private Sub StdBtnStop_Click()
    
    MCISend "Close All", Me
    FrmSettings.MnuAlarm.Visible = False
    FrmSettings.spc2.Visible = False
    
    If FrmAlarm.Visible = True Then FrmAlarm.Visible = False
    Attribs.Alarm.AlarmActivated = False
    
End Sub

Private Sub StdBtnStop_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    DoButton StdBtnStop, AlarmClip, True
    
End Sub

Private Sub StdBtnStop_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    DoButton StdBtnStop, AlarmClip, False
    
End Sub

Private Sub Timer1_Timer()
    
    'Jiggle The Picture And Make Sure The Alarm Window Is Topmost
    If Me.Visible = True Then
        SetWindowPos FrmMain.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
        Picture1.Top = IY + Int(Rnd * 90) - 45
        Picture1.Left = IX + Int(Rnd * 90) - 45
    End If
    
End Sub
