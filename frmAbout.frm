VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00400000&
   Caption         =   "About Rcon Unlimited"
   ClientHeight    =   4560
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6045
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4560
   ScaleWidth      =   6045
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   25
      Left            =   5400
      Top             =   3120
   End
   Begin VB.PictureBox PicMain 
      Appearance      =   0  'Flat
      BackColor       =   &H00400000&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      ScaleHeight     =   465
      ScaleWidth      =   4305
      TabIndex        =   4
      Top             =   4080
      Width           =   4335
      Begin VB.PictureBox Pic2 
         BackColor       =   &H00400000&
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   4320
         ScaleHeight     =   375
         ScaleWidth      =   15735
         TabIndex        =   5
         Top             =   0
         Width           =   15735
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   $"frmAbout.frx":210A
            ForeColor       =   &H00FF8080&
            Height          =   255
            Left            =   120
            TabIndex        =   6
            Top             =   120
            Width           =   13815
         End
      End
   End
   Begin VB.CommandButton CmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   4440
      TabIndex        =   3
      Top             =   3960
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   1500
      Left            =   0
      Picture         =   "frmAbout.frx":21D3
      Top             =   0
      Width           =   6000
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Thanks goes to the guys at Fragomatic for making Rcon Commander open source, so I learnt how to make this program a reality!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   3360
      Width           =   5295
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAbout.frx":6807
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   2520
      Width           =   5295
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAbout.frx":689F
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   5775
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdClose_Click()
    Unload Me
End Sub

Private Sub Timer1_Timer()
    Pic2.Left = Pic2.Left - 10
    If Pic2.Left <= -Pic2.Width Then Pic2.Left = PicMain.Width
End Sub
