VERSION 5.00
Begin VB.Form FrmAbout 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About Desktop Clock"
   ClientHeight    =   3270
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4230
   Icon            =   "FrmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   4230
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame 
      BackColor       =   &H00FFFFFF&
      Height          =   3195
      Left            =   188
      TabIndex        =   0
      Top             =   0
      Width           =   3855
      Begin VB.PictureBox StdBtnClose 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   450
         Left            =   2340
         Picture         =   "FrmAbout.frx":058A
         ScaleHeight     =   450
         ScaleWidth      =   1455
         TabIndex        =   3
         Tag             =   "4\5"
         ToolTipText     =   "Close About Window"
         Top             =   2700
         Width           =   1455
      End
      Begin VB.PictureBox PicLogo 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2400
         Left            =   120
         Picture         =   "FrmAbout.frx":2804
         ScaleHeight     =   2400
         ScaleWidth      =   3660
         TabIndex        =   1
         Top             =   180
         Width           =   3660
      End
      Begin VB.Label LblVersion 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Version:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFC0C0&
         Height          =   255
         Left            =   60
         TabIndex        =   2
         Top             =   2880
         Width           =   2295
      End
   End
End
Attribute VB_Name = "FrmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
    
    LblVersion.Caption = "Version: ProM " & App.Major & "." & App.Minor
    DoButton StdBtnClose, FrmMain.StdClip, False
    
    SkinPicture StdBtnClose, StdBtnClose
    
End Sub

Private Sub Picture1_Click()

End Sub

Private Sub StdBtnClose_Click()
    
    Me.Hide
    FrmMain.Show
    
End Sub

Private Sub StdBtnClose_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    DoButton StdBtnClose, FrmMain.StdClip, True
    
End Sub

Private Sub StdBtnClose_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
        DoButton StdBtnClose, FrmMain.StdClip, False
    
End Sub
