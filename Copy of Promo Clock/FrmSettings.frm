VERSION 5.00
Object = "{27395F88-0C0C-101B-A3C9-08002B2F49FB}#1.1#0"; "PICCLP32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmSettings 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "DC - Settings"
   ClientHeight    =   5280
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5955
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   5955
   StartUpPosition =   2  'CenterScreen
   Begin PicClip.PictureClip MscClip 
      Left            =   960
      Top             =   4800
      _ExtentX        =   1349
      _ExtentY        =   450
      _Version        =   393216
      Cols            =   3
      Picture         =   "FrmSettings.frx":058A
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   120
      Top             =   4800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame FrameAlarm 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Alarm Settings"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   60
      TabIndex        =   11
      Top             =   3480
      Width           =   5835
      Begin VB.HScrollBar Volume 
         Height          =   195
         LargeChange     =   10
         Left            =   4500
         Max             =   100
         TabIndex        =   27
         Top             =   420
         Width           =   1215
      End
      Begin VB.CommandButton CmdText 
         Height          =   315
         Left            =   3420
         Style           =   1  'Graphical
         TabIndex        =   26
         Tag             =   "P"
         ToolTipText     =   "Enter A Message To Be Displayed With Your Alarm"
         Top             =   360
         Width           =   315
      End
      Begin VB.CommandButton CmdPlay 
         Height          =   315
         Left            =   3780
         Style           =   1  'Graphical
         TabIndex        =   25
         Tag             =   "P"
         Top             =   360
         Width           =   315
      End
      Begin VB.CommandButton CmdSound 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   4140
         Picture         =   "FrmSettings.frx":1038
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Change Alarm Sound"
         Top             =   360
         Width           =   315
      End
      Begin VB.OptionButton OptAlarm 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Both"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   1860
         TabIndex        =   23
         Top             =   900
         Width           =   1035
      End
      Begin VB.OptionButton OptAlarm 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Audio Alarm"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   1860
         TabIndex        =   22
         Top             =   600
         Width           =   1215
      End
      Begin VB.OptionButton OptAlarm 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Visual Alarm"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   1860
         TabIndex        =   21
         Top             =   300
         Width           =   1275
      End
      Begin VB.ComboBox CmbMinute 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "FrmSettings.frx":137A
         Left            =   1020
         List            =   "FrmSettings.frx":137C
         TabIndex        =   18
         Text            =   "00"
         Top             =   660
         Width           =   615
      End
      Begin VB.ComboBox CmbHour 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   300
         TabIndex        =   17
         Text            =   "00"
         Top             =   660
         Width           =   675
      End
      Begin VB.CheckBox ChkAlarm 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "Alarm On"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4320
         TabIndex        =   16
         Top             =   780
         Width           =   1395
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Minute:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1020
         TabIndex        =   20
         Top             =   420
         Width           =   555
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Hour:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   300
         TabIndex        =   19
         Top             =   420
         Width           =   555
      End
   End
   Begin VB.PictureBox StdBtnApply 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Left            =   2940
      Picture         =   "FrmSettings.frx":137E
      ScaleHeight     =   450
      ScaleWidth      =   1455
      TabIndex        =   8
      Tag             =   "2\3"
      ToolTipText     =   "Save All Setting Changes"
      Top             =   4740
      Width           =   1455
   End
   Begin VB.PictureBox StdBtnClose 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Left            =   4440
      Picture         =   "FrmSettings.frx":35F8
      ScaleHeight     =   450
      ScaleWidth      =   1455
      TabIndex        =   7
      Tag             =   "4\5"
      ToolTipText     =   "Close Settings Window"
      Top             =   4740
      Width           =   1455
   End
   Begin VB.Frame FrmSkin 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Clock Skin Settings"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   60
      TabIndex        =   5
      Top             =   2280
      Width           =   5835
      Begin VB.PictureBox StdBtnNewSkin 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   450
         Left            =   4200
         Picture         =   "FrmSettings.frx":5872
         ScaleHeight     =   450
         ScaleWidth      =   1455
         TabIndex        =   6
         Tag             =   "0\1"
         ToolTipText     =   "Select A New Skin For Your Clock"
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label LblNewSkin 
         BackStyle       =   0  'Transparent
         Caption         =   "Default Skin"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   195
         Left            =   1260
         TabIndex        =   14
         ToolTipText     =   "Click Apply To Switch To This Skin"
         Top             =   720
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.Label LblNSDesc 
         BackStyle       =   0  'Transparent
         Caption         =   "Switch To:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   195
         Left            =   300
         TabIndex        =   13
         Top             =   720
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label LblSkin 
         BackStyle       =   0  'Transparent
         Caption         =   "Default Skin"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1260
         TabIndex        =   10
         ToolTipText     =   "The Name Of Your Current Clock Skin"
         Top             =   480
         Width           =   2655
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Current Skin:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   300
         TabIndex        =   9
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.Frame FrmGeneral 
      BackColor       =   &H00E0E0E0&
      Caption         =   "General Clock Settings "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1095
      Left            =   60
      TabIndex        =   2
      Top             =   1020
      Width           =   5835
      Begin VB.CheckBox ChkOntop 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Clock Always Ontop"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3180
         TabIndex        =   15
         ToolTipText     =   "Your Clock Will Always Be Visible"
         Top             =   660
         Width           =   1935
      End
      Begin VB.CheckBox ChkWindows 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Start  With Windows"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3180
         TabIndex        =   12
         Top             =   360
         Width           =   2235
      End
      Begin VB.CheckBox ChkDispAlarm 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Display Alarm Status"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   300
         TabIndex        =   4
         Top             =   660
         Width           =   2895
      End
      Begin VB.CheckBox ChkRemPos 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Remember Clocks Screen Position"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   300
         TabIndex        =   3
         Top             =   360
         Width           =   2895
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8115
      Begin VB.Line Line1 
         BorderColor     =   &H00FFC0C0&
         X1              =   0
         X2              =   8100
         Y1              =   900
         Y2              =   900
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Desktop Clock Settings"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFC0C0&
         Height          =   555
         Left            =   240
         TabIndex        =   1
         Top             =   180
         Width           =   6735
      End
   End
   Begin VB.Menu MnuMenu 
      Caption         =   "MnuMenu"
      Visible         =   0   'False
      Begin VB.Menu MnuAlarm 
         Caption         =   "Alarm"
         Visible         =   0   'False
         Begin VB.Menu MnuAlarmStop 
            Caption         =   "Stop"
         End
         Begin VB.Menu MnuAlarmSnooze 
            Caption         =   "Snooze"
         End
      End
      Begin VB.Menu spc2 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu MnuSettings 
         Caption         =   "Settings"
      End
      Begin VB.Menu spc 
         Caption         =   "-"
      End
      Begin VB.Menu MnuAbout 
         Caption         =   "About Desktop Clock"
      End
      Begin VB.Menu MnuExit 
         Caption         =   "Exit Desktop Clock"
      End
      Begin VB.Menu MnuHide 
         Caption         =   "HideMenu"
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "FrmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ChkAlarm_Click()
    
    ApplyEnabled StdBtnApply, True
    
End Sub

Private Sub ChkDispAlarm_Click()
    
    ApplyEnabled StdBtnApply, True
    
End Sub

Private Sub ChkOntop_Click()
    
    ApplyEnabled StdBtnApply, True
    
End Sub

Private Sub ChkRemPos_Click()
    
    ApplyEnabled StdBtnApply, True
    
End Sub

Private Sub ChkWindows_Click()
    
    ApplyEnabled StdBtnApply, True
    
End Sub

Private Sub CmbHour_Change()
    
    ApplyEnabled StdBtnApply, True
    
End Sub

Private Sub CmbHour_Click()
    
    ApplyEnabled StdBtnApply, True
    
End Sub

Private Sub CmbHour_GotFocus()
    
    If Not IsNumeric(CmbHour.Text) Then
        CmbHour.Text = Hour(Attribs.Alarm.AlarmTime)
    End If
    
    If Val(CmbHour.Text) > 23 Or Val(CmbHour.Text) < 0 Then
        CmbHour.Text = Hour(Attribs.Alarm.AlarmTime)
    End If
    
End Sub

Private Sub CmbHour_KeyPress(KeyAscii As Integer)
    
    If Not IsNumeric(Chr(KeyAscii)) Then
        If Not KeyAscii = 8 Then
            KeyAscii = 0
        End If
    End If
    
End Sub

Private Sub CmbHour_LostFocus()
    
    CmbHour_GotFocus
    
End Sub

Private Sub CmbMinute_Change()
    
    ApplyEnabled StdBtnApply, True
    
End Sub

Private Sub CmbMinute_Click()
    
    ApplyEnabled StdBtnApply, True
    
End Sub

Private Sub CmbMinute_GotFocus()

    If Not IsNumeric(CmbMinute.Text) Then
        CmbMinute.Text = Minute(Attribs.Alarm.AlarmTime)
    End If
    
    If Val(CmbMinute.Text) > 59 Or Val(CmbMinute.Text) < 0 Then
        CmbMinute.Text = Minute(Attribs.Alarm.AlarmTime)
    End If
    
End Sub

Private Sub CmbMinute_KeyPress(KeyAscii As Integer)
    
    If Not IsNumeric(Chr(KeyAscii)) Then
        If Not KeyAscii = 8 Then
            KeyAscii = 0
        End If
    End If
    
End Sub

Private Sub CmbMinute_LostFocus()
    
    CmbMinute_GotFocus
    
End Sub

Private Sub CmdPlay_Click()
    
    MCISend "setaudio CH1 volume to " & Attribs.Alarm.AlarmVolume, Me
    
    If CmdPlay.Tag = "P" Then
        CmdPlay.Tag = "S"
        MCISend "Play CH1 From 0 Repeat", Me
        CmdPlay.ToolTipText = "Stop (" & GetFileName(TempSound) & ")"
        CmdPlay.Picture = MscClip.GraphicCell(0)
    Else
        CmdPlay.Tag = "P"
        MCISend "Stop CH1", Me
        CmdPlay.Picture = MscClip.GraphicCell(1)
        CmdPlay.ToolTipText = "Play (" & GetFileName(TempSound) & ")"
    End If
    
    
            
            
    
End Sub

Private Sub CmdSound_Click()
On Error GoTo CDError
    
    MCISend "Close All", Me
    CmdPlay.Tag = "P"
    CmdPlay.Picture = MscClip.GraphicCell(1)
    
    
    CD.CancelError = True
    CD.InitDir = App.path
    CD.Filter = "Mp3 Sounds (*.Mp3)|*.Mp3|All Files (*.*)|*.*"
    CD.Flags = &H81804
    CD.DialogTitle = "Select A New Alarm Sound."
    
    CD.ShowOpen
    
    TempSound = CD.FileName
    ApplyEnabled StdBtnApply, True
    
    
    MCISend "Open """ & TempSound & """ Alias CH1", Me
    CmdPlay.ToolTipText = "Play (" & GetFileName(TempSound) & ")"
    
CDError:
MCISend "Open """ & TempSound & """ Alias CH1", Me
End Sub

Private Sub CmdText_Click()
        
    TempString = InputBox("Enter Your Alarm Message.", , Attribs.Alarm.AlarmNote)
    If StripNonChar(TempString) = "" Then
        TempString = Attribs.Alarm.AlarmNote
    Else
        ApplyEnabled StdBtnApply, True
    End If
    
End Sub

Private Sub Form_Activate()
    
    FrmMain.Enabled = False
    
    ApplyEnabled StdBtnApply, False
    
    If Attribs.Skin.Skin = "Default" Or Attribs.Skin.Skin = "" Then
        Attribs.Skin.Skin = "Default"
        LblSkin.Caption = "Default"
    Else
        LblSkin.Caption = GetFileName(Attribs.Skin.Skin)
    End If
    
    ChkRemPos.Value = Cbl(Attribs.Clock.RemLocation)
    ChkDispAlarm.Value = Cbl(Attribs.Clock.DispAlarm)
    ChkWindows.Value = Cbl(Attribs.Clock.StartWithWindows)
    ChkOntop.Value = Cbl(Attribs.Clock.AlwaysOntop)
    ChkAlarm.Value = Cbl(Attribs.Alarm.AlarmSet)
    
    LblNSDesc.Visible = False
    LblNewSkin.Visible = False
    
    TempSkin = Attribs.Skin.Skin
    TempString = Attribs.Alarm.AlarmNote
    
    If Attribs.Alarm.AlarmSet = False Then
        CmbHour.Text = Format(Hour(Now), "00")
        CmbMinute = Format(Minute(Now), "00")
    Else
        CmbHour.Text = Format(Hour(Attribs.Alarm.AlarmTime), "00")
        CmbMinute.Text = Format(Minute(Attribs.Alarm.AlarmTime), "00")
    End If
    
    AlarmSetup
    
    CmdPlay.Picture = MscClip.GraphicCell(1)
    CmdPlay.Tag = "P"
    
    
    
    CmdText.Picture = MscClip.GraphicCell(2)
    
    'Override Alarm Type If No Sound File Is Available
    If Attribs.Alarm.AlarmSound <> "None" Then
        OptAlarm(1).Enabled = True
        OptAlarm(2).Enabled = True
        OptAlarm(Attribs.Alarm.AlarmType).Value = True
    Else
        OptAlarm(0).Value = True
        OptAlarm(1).Enabled = False
        OptAlarm(2).Enabled = False
    End If
    
    If FrmAlarm.Visible = True Then FrmAlarm.Visible = False
    
    MCISend "Close All", Me
    MCISend "Open """ & Attribs.Alarm.AlarmSound & """ Alias CH1", Me
    
    TempSound = Attribs.Alarm.AlarmSound
    CmdSound.ToolTipText = "Change Alarm Sound (" & GetFileName(TempSound) & ")"
    CmdPlay.ToolTipText = "Play (" & GetFileName(TempSound) & ")"
    
    ApplyEnabled StdBtnApply, False
    
    Volume.Value = Int(Attribs.Alarm.AlarmVolume / 10)
    
End Sub

Private Sub Form_Load()
    
    Dim LP As Single
    
    DoButton StdBtnApply, FrmMain.StdClip, False
    DoButton StdBtnClose, FrmMain.StdClip, False
    DoButton StdBtnNewSkin, FrmMain.StdClip, False
    
    SkinPicture StdBtnApply, StdBtnApply
    SkinPicture StdBtnClose, StdBtnClose
    SkinPicture StdBtnNewSkin, StdBtnNewSkin
           
    For LP = 0 To 23
        CmbHour.AddItem Format(LP, "00")
    Next LP
               
    For LP = 0 To 55 Step 5
        CmbMinute.AddItem Format(LP, "00")
    Next LP
               
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Cancel = -1
    
End Sub

Private Sub MnuAbout_Click()
    
    FrmAbout.Show
    
End Sub

Private Sub MnuAlarmSnooze_Click()
    
    MCISend "Close All", Me
    
    FrmSettings.MnuAlarm.Visible = False
    FrmSettings.spc2.Visible = False
    
    Attribs.Alarm.AlarmActivated = False
    Attribs.Alarm.AlarmTime = Snooze(Attribs.Alarm.AlarmTime, SnoozeTime)
    Attribs.Alarm.AlarmSet = True
    SetRegSettings
    
    AlarmSetup
    
End Sub

Private Sub MnuAlarmStop_Click()
    
    MCISend "Close All", Me
    FrmSettings.MnuAlarm.Visible = False
    FrmSettings.spc2.Visible = False
    If FrmAlarm.Visible = True Then FrmAlarm.Visible = False
    Attribs.Alarm.AlarmActivated = False
    
End Sub

Private Sub MnuExit_Click()
    Dim Resp As Long
    
    MCISend "Close All", Me
    
    Resp = MsgBox("Do You Really Want To Exit Desktop Clock?", vbYesNo + vbExclamation)
    
    If Resp = vbYes Then
         'remove the icon
        TrayI.cbSize = Len(TrayI)
        TrayI.hWnd = FrmMain.PicHook.hWnd
        TrayI.uId = 1&
        Shell_NotifyIcon NIM_DELETE, TrayI
        SetRegSettings
        End
    End If
    
End Sub

Private Sub MnuHide_Click()
    
    MnuHide.Visible = False
    FrmMain.Enabled = True
    
End Sub

Private Sub MnuSettings_Click()
    
    FrmSettings.Show
    
End Sub

Private Sub OptAlarm_Click(Index As Integer)
    
    ApplyEnabled StdBtnApply, True
    
End Sub




Private Sub StdBtnApply_Click()
       
    If Not IsNumeric(CmbHour.Text) Then
        CmbHour.Text = Hour(Attribs.Alarm.AlarmTime)
    End If
    
    If Not IsNumeric(CmbMinute.Text) Then
        CmbMinute.Text = Minute(Attribs.Alarm.AlarmTime)
    End If
    
    If Val(CmbHour.Text) > 23 Or Val(CmbHour.Text) < 0 Then
        CmbHour.Text = Hour(Attribs.Alarm.AlarmTime)
    End If
    
    If Val(CmbMinute.Text) > 59 Or Val(CmbMinute.Text) < 0 Then
        CmbMinute.Text = Minute(Attribs.Alarm.AlarmTime)
    End If
    
    ApplyEnabled StdBtnApply, False
    
    'If The Skin Has Changed Then Update The Clock
    If TempSkin <> Attribs.Skin.Skin Then
        LblNewSkin.Caption = "Updating Skin."
        Attribs.Skin.Skin = TempSkin
        NewSkin
        ReDrawClock
        Do
            DoEvents
        Loop Until FrmMain.Visible = True
        LblSkin.Caption = GetFileName(Attribs.Skin.Skin)
    End If
    
    LblNSDesc.Visible = False
    LblNewSkin.Visible = False
    
    If Trim(TempString) = "" Then TempString = "Your Alarm Time Has Been Reached"
    
    Attribs.Alarm.AlarmNote = TempString
    Attribs.Alarm.AlarmSound = TempSound
    Attribs.Clock.DispAlarm = CBool(ChkDispAlarm.Value)
    Attribs.Clock.StartWithWindows = CBool(ChkWindows.Value)
    Attribs.Clock.RemLocation = CBool(ChkRemPos.Value)
    Attribs.Clock.AlwaysOntop = CBool(ChkOntop.Value)
    Attribs.Alarm.AlarmSet = CBool(ChkAlarm.Value)
    
    If OptAlarm(0).Value = True Then Attribs.Alarm.AlarmType = 0
    If OptAlarm(1).Value = True Then Attribs.Alarm.AlarmType = 1
    If OptAlarm(2).Value = True Then Attribs.Alarm.AlarmType = 2
    
    'Store The Alarm Time In The Settings
    Attribs.Alarm.AlarmTime = Format(Left$(CmbHour.Text, 2), "00") & ":" & Format(Left$(CmbMinute.Text, 2), "00") & ":00"
    AlarmSetup
    
    'Remove Or Apply The Always Ontop API Setting
    If Attribs.Clock.AlwaysOntop = True Then
        SetWindowPos FrmMain.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
    Else
        SetWindowPos FrmMain.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
    End If
    
    MaskPrepared = False
    DrawTime
    ReDrawClock
    
    'Prevent The Clock From Being Completely Hidden By Mistake
    If FrmMain.Left < 0 - FrmMain.Width / 2 Then FrmMain.Left = 0
    If FrmMain.Top < 0 - FrmMain.Height / 2 Then FrmMain.Top = 0
    If FrmMain.Left > Screen.Width Then FrmMain.Left = Screen.Width - FrmMain.Width
    If FrmMain.Top > Screen.Height Then FrmMain.Top = Screen.Height - FrmMain.Height
    
    'Override Alarm Type If No Sound File Is Available
    If Attribs.Alarm.AlarmSound <> "None" Then
        OptAlarm(1).Enabled = True
        OptAlarm(2).Enabled = True
        OptAlarm(Attribs.Alarm.AlarmType).Value = True
    Else
        OptAlarm(0).Value = True
        OptAlarm(1).Enabled = False
        OptAlarm(2).Enabled = False
    End If
    
    'Reset Tooltips, Buttons And Channel 1 Audio File To The New Sound File
    MCISend "Close All", Me
    CmdPlay.Tag = "P"
    CmdPlay.Picture = MscClip.GraphicCell(1)
    CmdPlay.ToolTipText = "Play (" & GetFileName(TempSound) & ")"
    CmdSound.ToolTipText = "Change Alarm Sound (" & GetFileName(Attribs.Alarm.AlarmSound) & ")"
    MCISend "Open """ & Attribs.Alarm.AlarmSound & """ Alias CH1", Me
    
    Dim path As Long
    
    'Start With Windows Registry Entry
    If Attribs.Clock.StartWithWindows = True Then
        If RegOpenKeyEx(HKEY_CURRENT_USER, SUREGKEY, 0, KEY_WRITE, path) Then Exit Sub
        RegSetValueEx path, App.Title & ".Exe", 0, REG_SZ, ByVal App.path & "\" & App.Title & ".Exe", Len(App.path & "\" & App.Title & ".Exe")
    Else
        If RegOpenKeyEx(HKEY_CURRENT_USER, SUREGKEY, 0, KEY_WRITE, path) Then Exit Sub
        RegDeleteValue path, App.Title & ".Exe"
    End If
    
    Attribs.Alarm.AlarmVolume = Volume.Value * 10
    
    SetRegSettings
    Me.SetFocus
End Sub

Private Sub StdBtnApply_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    DoButton StdBtnApply, FrmMain.StdClip, True
    
End Sub

Private Sub StdBtnApply_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    DoButton StdBtnApply, FrmMain.StdClip, False
    
End Sub

Private Sub StdBtnClose_Click()
    
    Me.Hide
    MCISend "Close All", Me
    FrmMain.Enabled = True
    FrmMain.ZOrder (0)
        
End Sub

Private Sub StdBtnClose_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    DoButton StdBtnClose, FrmMain.StdClip, True
    
End Sub

Private Sub StdBtnClose_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    DoButton StdBtnClose, FrmMain.StdClip, False
    
End Sub

Private Sub StdBtnNewSkin_Click()
On Error GoTo CDError
    
    CD.CancelError = True
    If FSys.FolderExists(App.path & "\Clock Skins") = True Then
        CD.InitDir = App.path & "\Clock Skins"
    Else
        CD.InitDir = App.path
    End If
    
    CD.Filter = "Bitmap Skins (*.Bmp)|*.Bmp|JPeg Skins (*.Jpg)|*.Jpg)|All Files (*.*)|*.*"
    CD.Flags = &H81804
    CD.DialogTitle = "Select A New Skin."
    
    CD.ShowOpen
    
    TempSkin = CD.FileName
    LblNewSkin.Caption = GetFileName(TempSkin)
    LblNSDesc.Visible = True
    LblNewSkin.Visible = True
    ApplyEnabled StdBtnApply, True
    
    
CDError:
End Sub

Private Sub StdBtnNewSkin_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    DoButton StdBtnNewSkin, FrmMain.StdClip, True
    
End Sub

Private Sub StdBtnNewSkin_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    DoButton StdBtnNewSkin, FrmMain.StdClip, False
    
End Sub

Private Sub Volume_Change()
    
    ApplyEnabled StdBtnApply, True
    MCISend "setaudio CH1 volume to " & (Volume.Value * 10), Me
    
End Sub

Private Sub Volume_Scroll()
    
    ApplyEnabled StdBtnApply, True
    MCISend "setaudio CH1 volume to " & (Volume.Value * 10), Me
    
End Sub
