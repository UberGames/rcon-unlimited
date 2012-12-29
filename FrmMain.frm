VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rcon Unlimited v1.0"
   ClientHeight    =   8115
   ClientLeft      =   3315
   ClientTop       =   3060
   ClientWidth     =   11970
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   541
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   798
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   120
      TabIndex        =   9
      Top             =   0
      Width           =   8535
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   2880
         Top             =   120
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   3360
         Top             =   120
      End
      Begin VB.Timer Timer1a 
         Interval        =   2000
         Left            =   5640
         Top             =   120
      End
      Begin MSWinsockLib.Winsock Winsock1 
         Left            =   2400
         Top             =   120
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         Protocol        =   1
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   7800
         Top             =   240
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label lblPing 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   5760
         TabIndex        =   22
         Top             =   600
         Width           =   2535
      End
      Begin VB.Label Label1b 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Server Ping:"
         Height          =   255
         Left            =   4440
         TabIndex        =   21
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label1a 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Version:"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label2a 
         Alignment       =   1  'Right Justify
         Caption         =   "Host Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label Label3a 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Game Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label4a 
         Alignment       =   1  'Right Justify
         Caption         =   "Map:"
         Height          =   255
         Left            =   4440
         TabIndex        =   16
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label Label5a 
         Alignment       =   1  'Right Justify
         Caption         =   "Players:"
         Height          =   255
         Left            =   4440
         TabIndex        =   15
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label lblHostName 
         Height          =   255
         Left            =   1440
         TabIndex        =   14
         Top             =   120
         Width           =   3255
      End
      Begin VB.Label lblGameName 
         Height          =   255
         Left            =   1440
         TabIndex        =   13
         Top             =   360
         Width           =   3255
      End
      Begin VB.Label lblVersion 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   1440
         TabIndex        =   12
         Top             =   600
         Width           =   3255
      End
      Begin VB.Label lblMap 
         Height          =   255
         Left            =   5760
         TabIndex        =   11
         Top             =   120
         Width           =   2535
      End
      Begin VB.Label lblMaxClients 
         Height          =   255
         Left            =   5760
         TabIndex        =   10
         Top             =   360
         Width           =   2535
      End
   End
   Begin VB.Frame Frame3 
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Top             =   7320
      Width           =   11775
      Begin VB.CommandButton CmdStatus 
         Caption         =   "Status"
         Height          =   375
         Left            =   120
         TabIndex        =   26
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton CmdInfo 
         Caption         =   "Server Info"
         Height          =   375
         Left            =   1440
         TabIndex        =   25
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Change Speed"
         Height          =   375
         Left            =   8640
         TabIndex        =   24
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton CmdGrav 
         Caption         =   "Change Gravity"
         Height          =   375
         Left            =   7080
         TabIndex        =   23
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton CmdRestart 
         Caption         =   "Shutdown Server"
         Height          =   375
         Left            =   10200
         TabIndex        =   8
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton CmdKick 
         Caption         =   "Kick"
         Height          =   375
         Left            =   4200
         TabIndex        =   7
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton CmdBan 
         Caption         =   "Ban"
         Height          =   375
         Left            =   5640
         TabIndex        =   6
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton CmdSay 
         Caption         =   "Say"
         Height          =   375
         Left            =   2760
         TabIndex        =   5
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.TextBox txtOutput 
      BackColor       =   &H00400000&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   5655
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   960
      Width           =   8655
   End
   Begin VB.Frame Frame1 
      Caption         =   "Command"
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   6600
      Width           =   8415
      Begin MSWinsockLib.Winsock UDPClient 
         Left            =   1440
         Top             =   120
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         Protocol        =   1
      End
      Begin VB.ComboBox cboCommand 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "FrmMain.frx":210A
         Left            =   120
         List            =   "FrmMain.frx":210C
         TabIndex        =   3
         Top             =   240
         Width           =   6720
      End
      Begin VB.CommandButton CmdSend 
         Caption         =   "Send"
         Default         =   -1  'True
         Height          =   375
         Left            =   7080
         TabIndex        =   2
         Top             =   240
         Width           =   1215
      End
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   7215
      Left            =   8760
      TabIndex        =   20
      Top             =   0
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   12726
      View            =   3
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   16777215
      BackColor       =   4194304
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Player"
         Object.Width           =   2999
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Frags"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Ping"
         Object.Width           =   1235
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnudisconnect 
         Caption         =   "Disconnect"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnusep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuquit 
         Caption         =   "Quit"
      End
   End
   Begin VB.Menu mnuedit 
      Caption         =   "Edit"
      Begin VB.Menu mnucopy 
         Caption         =   "Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuselall 
         Caption         =   "Select All"
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu mnulog 
      Caption         =   "Log"
      Begin VB.Menu mnuclearlog 
         Caption         =   "Clear Log Window"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnusaveas 
         Caption         =   "Save Log Window as"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnusep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuautosave 
         Caption         =   "Auto-Save Log"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuviewlog 
         Caption         =   "View Auto-Save Log"
      End
      Begin VB.Menu mnuclear 
         Caption         =   "Clear Auto-Save Log"
      End
   End
   Begin VB.Menu mnuopt 
      Caption         =   "Commands"
      Begin VB.Menu mnustatus 
         Caption         =   "Get Server Status"
         Shortcut        =   {F8}
      End
      Begin VB.Menu mnuinfo 
         Caption         =   "Get Server Information"
         Shortcut        =   {F9}
      End
      Begin VB.Menu mnusep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuac 
         Caption         =   "Auto Command Complete"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnusta 
         Caption         =   "Auto-Query Status"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnusepe 
         Caption         =   "-"
      End
      Begin VB.Menu nuautocomplete 
         Caption         =   "Edit AutoComplete Command List..."
      End
   End
   Begin VB.Menu mnuview 
      Caption         =   "View"
      Begin VB.Menu mnucommands 
         Caption         =   "Commands List"
      End
      Begin VB.Menu mnuvaraible 
         Caption         =   "Variables List"
      End
      Begin VB.Menu mnusep7 
         Caption         =   "-"
      End
      Begin VB.Menu mnureadme 
         Caption         =   "Readme"
      End
   End
   Begin VB.Menu mnuhelp 
      Caption         =   "Help"
      Begin VB.Menu mnufeedback 
         Caption         =   "Submit Feedback"
      End
      Begin VB.Menu mnubug 
         Caption         =   "Submit Bug Report"
      End
      Begin VB.Menu mnusep6 
         Caption         =   "-"
      End
      Begin VB.Menu mnucom 
         Caption         =   "Community Links"
         Begin VB.Menu mnugame1 
            Caption         =   "Return To Castle Wolfenstein"
            Begin VB.Menu mnulink4 
               Caption         =   "Offical Site"
            End
            Begin VB.Menu mnulink1 
               Caption         =   "RTCWFiles.com"
            End
         End
         Begin VB.Menu mnugame2 
            Caption         =   "Star Trek Voyager: Elite Force"
            Begin VB.Menu mnulink5 
               Caption         =   "Offical Site"
            End
            Begin VB.Menu mnulink2 
               Caption         =   "EFFiles.com"
            End
         End
         Begin VB.Menu mnugame3 
            Caption         =   "Star Trek: Elite Force 2"
            Begin VB.Menu mnulink8 
               Caption         =   "Offical Site"
            End
            Begin VB.Menu mnulink3 
               Caption         =   "EF2Files.com"
            End
         End
         Begin VB.Menu mnumisclinks 
            Caption         =   "Misc Links"
            Begin VB.Menu mnulink9 
               Caption         =   "SFEF Clan"
            End
            Begin VB.Menu mnulink11 
               Caption         =   "AQRP"
            End
         End
         Begin VB.Menu mnulink10 
            Caption         =   "GamingForums.com"
         End
      End
      Begin VB.Menu mnusep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuemail 
         Caption         =   "Email Me"
      End
      Begin VB.Menu mnusep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuabout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strData As String
Dim lastCommand As String
Dim ctr As Integer
Dim ping1 As Currency
Dim hptimer As Boolean
Dim autorefreshon As Boolean
Dim strBaseGame As String
Dim strMod As String
Dim hnparam As String
Dim players As Integer
Dim maxplayers As Integer
Dim l1so As Boolean
Dim l2so As Boolean
Dim l2sortcolumn As Integer
Dim l1sortcolumn As Integer
Dim strGame As String
Dim strCmdLine As String
Dim strExePath As String
Dim autorefreshtime As Integer
Dim closebrowser As Boolean
Dim CanComplete As Boolean



Private Sub cboCommand_Change()
    Dim start As Integer
    Dim NewCmd As String
    
    On Error Resume Next

    If mnuac.Checked = True And CanComplete = True Then
        NewCmd = LookupCommand(cboCommand.Text)
        start = Len(cboCommand)

        cboCommand.Text = cboCommand.Text & NewCmd
        cboCommand.SelStart = start
        cboCommand.SelLength = Len(NewCmd)
    End If
End Sub


Private Sub cboCommand_KeyDown(KeyCode As Integer, Shift As Integer)
   ' If KeyCode = vbKeyUp Then
   '     cboCommand.Text = lastCommand
   ' End If
    
    If mnuac.Checked Then
        If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then
            CanComplete = False
        ElseIf KeyCode = vbKeySpace Or KeyCode = vbKeyTab Then
            cboCommand.SelStart = Len(cboCommand)
            CanComplete = True
        Else
            CanComplete = True
        End If
    End If
End Sub

Private Sub CmdBan_Click()
    Dim playerIp As String
    
    
    playerIp = InputBox("What is the IP of the player you wish to ban?", "Rcon Unlimited")
    
    If playerIp <> "" Then SendCommand "addip " + playerIp
End Sub

Private Sub CmdGrav_Click()
    Dim grav As String
    
    
    grav = InputBox("What is the gravity you wish to set the server at?", "Rcon Unlimited", "800")
    
    If grav <> "" Then SendCommand "g_gravity " + grav
End Sub

Private Sub CmdInfo_Click()
    SendCommand "serverinfo"
End Sub

Private Sub CmdKick_Click()
    Dim playerID As String
    
    
    playerID = InputBox("What is the num of the player you wish to kick?", "Rcon Unlimited")
    
    If playerID <> "" Then SendCommand "kick " + playerID
End Sub

Private Sub CmdRestart_Click()
    If MsgBox("Are you sure you wish to shutdown the server?", vbYesNo, "Rcon Unlimited") = vbYes Then SendCommand "quit"
End Sub

Private Sub CmdSay_Click()
    Dim tosay As String
    
    tosay = InputBox("What do you want to say to the players." + vbNewLine + "Note you will not be able to see what they say back", "Rcon Unlimited")
    If tosay <> "" Then SendCommand "svsay " + tosay
End Sub

Private Sub CmdSend_Click()
    SendCommand Trim(cboCommand)
    
    If Not InHistory(cboCommand) Then
            cboCommand.AddItem (cboCommand.Text)
    End If
    
    lastCommand = cboCommand.Text
    cboCommand.Text = ""
    cboCommand.SetFocus
End Sub

Private Sub CmdStatus_Click()
    SendCommand
End Sub

Private Sub Command1_Click()
    Dim speed As String
    
    
    speed = InputBox("What is the speed you wish to set the server at?", "Rcon Unlimited", "250")
    
    If speed <> "" Then SendCommand "g_speed " + speed
End Sub

Private Sub Form_Load()
    On Error Resume Next
    
    GetCommandList
    
    If GetSetting("Rcon Unlimited", "settings", "acc") = "0" Then
        mnuac.Checked = False
    Else
        mnuac.Checked = True
    End If
    
    If GetSetting("Rcon Unlimited", "settings", "autosave") = "0" Then
        mnuautosave.Checked = False
    Else
        mnuautosave.Checked = True
    End If
    
    If GetSetting("Rcon Unlimited", "settings", "sta") = "0" Then
        mnusta.Checked = False
    Else
        mnusta.Checked = True
        SendCommand
    End If
                
    Q3_sendData
End Sub

Private Sub mnuabout_Click()
    frmAbout.Show
End Sub

Private Sub mnuac_Click()
    If mnuac.Checked = True Then
        mnuac.Checked = False
        SaveSetting "Rcon Unlimited", "settings", "acc", "0"
    Else
        mnuac.Checked = True
        SaveSetting "Rcon Unlimited", "settings", "acc", "1"
    End If
End Sub

Private Sub mnuautosave_Click()
    If mnuautosave.Checked = True Then
        mnuautosave.Checked = False
        SaveSetting "Rcon Unlimited", "settings", "autosave", "0"
    Else
        mnuautosave.Checked = True
        SaveSetting "Rcon Unlimited", "settings", "autosave", "1"
    End If
End Sub

Private Sub mnubug_Click()
    OpenURL "mailto:phenix@sg15.com?subject=Bug Report: Rcon Unlimited v1.0"
End Sub

Private Sub mnuclear_Click()
    If FileExists(CheckPath(App.Path) & "autosave.log") Then
        Kill (App.Path & "\autosave.log")
    End If
End Sub

Private Sub mnuclearlog_Click()
    txtOutput = ""
    cboCommand.SetFocus
End Sub

Private Sub mnucommands_Click()
    If FileExists(CheckPath(App.Path) & "commands.txt") Then
        Shell "notepad.exe " & CheckPath(App.Path) & "commands.txt", vbNormalFocus
    Else
        MsgBox "The commands file could not be found.", vbInformation, "Rcon Unlimited"
    End If
End Sub

Private Sub mnucopy_Click()
    On Error Resume Next
    Clipboard.Clear
    Clipboard.SetText txtOutput.SelText
End Sub

Private Sub mnudisconnect_Click()
    If MsgBox("Are you sure you wish to disconnect from this server?", vbYesNo, "Rcon Unlimited") = vbYes Then
        FrmStartup.Show
        Unload Me
    End If
End Sub

Private Sub mnuemail_Click()
    OpenURL "mailto:phenix@sg15.com"
End Sub

Private Sub mnufeedback_Click()
    OpenURL "mailto:phenix@sg15.com?subject=Feedback: Rcon Unlimited v1.0"
End Sub

Private Sub mnuinfo_Click()
    SendCommand ("serverinfo")
End Sub

Private Sub mnulink1_Click()
    OpenURL "http://www.rtcwfiles.com"
End Sub

Private Sub mnulink10_Click()
    OpenURL "http://www.gamingforums.com"
End Sub

Private Sub mnulink11_Click()
    OpenURL "http://www.alpha-quadrant.net"
End Sub

Private Sub mnulink2_Click()
    OpenURL "http://www.effiles.com"
End Sub

Private Sub mnulink3_Click()
    OpenURL "http://www.ef2files.com"
End Sub

Private Sub mnulink4_Click()
    OpenURL "http://www.castlewolfenstein.com"
End Sub

Private Sub mnulink5_Click()
    OpenURL "http://www.ravensoft.com/eliteforce/"
End Sub

Private Sub mnulink8_Click()
    OpenURL "http://gaming.startrek.com/games/eliteforce2"
End Sub

Private Sub mnulink9_Click()
    OpenURL "http://www.sfefclan.com"
End Sub

Private Sub mnuquit_Click()
    If MsgBox("Are you sure you wish to close Rcon Unlimited?", vbYesNo, "Rcon Unlimited") = vbYes Then
        End
    End If
End Sub

Private Sub mnureadme_Click()
    If FileExists(CheckPath(App.Path) & "readme.txt") Then
        Shell "notepad.exe " & CheckPath(App.Path) & "readme.txt", vbNormalFocus
    Else
        MsgBox "The Read-Me file could not be found.", vbInformation, "Rcon Unlimited"
    End If
End Sub

Private Sub mnusaveas_Click()
    CommonDialog1.Filter = "Log File (*.log)|*.log|Text File (*.txt)|*.txt|All Files (*.*)|*.*|"
    CommonDialog1.Action = 2
    On Error Resume Next
    Open CommonDialog1.FileName For Output As #1
        Print #1, txtOutput
    Close #1
End Sub

Private Sub mnuselall_Click()
    txtOutput.SetFocus
    txtOutput.SelStart = 0
    txtOutput.SelLength = Len(txtOutput)
End Sub

Private Sub mnusta_Click()
    If mnusta.Checked = True Then
        mnusta.Checked = False
        SaveSetting "Rcon Unlimited", "settings", "sta", "0"
    Else
        mnusta.Checked = True
        SaveSetting "Rcon Unlimited", "settings", "sta", "1"
    End If
End Sub

Private Sub mnustatus_Click()
    SendCommand
End Sub

Private Sub mnuvaraible_Click()
    If FileExists(CheckPath(App.Path) & "variables.txt") Then
        Shell "notepad.exe " & CheckPath(App.Path) & "variables.txt", vbNormalFocus
    Else
        MsgBox "The variables file could not be found.", vbInformation, "Rcon Unlimited"
    End If
End Sub

Private Sub mnuviewlog_Click()
    If FileExists(CheckPath(App.Path) & "autosave.log") Then
        Shell "notepad.exe " & CheckPath(App.Path) & "autosave.log", vbNormalFocus
    Else
        MsgBox "Auto log is empty.", vbInformation, "Rcon Unlimited"
    End If
End Sub

Private Sub nuautocomplete_Click()
    If FileExists(CheckPath(App.Path) & "commands.dat") Then
        Shell "notepad.exe " & CheckPath(App.Path) & "commands.dat", vbNormalFocus
    End If
End Sub

Private Sub SendCommand(Optional Cmd As String = "status")
    On Error Resume Next
    Dim sStamp As String
    
    sStamp = vbCrLf & vbCrLf & Time & ":> " & Cmd & vbCrLf & String(Len(Time & ":> " & Cmd), "-")
    
    UDPClient.RemoteHost = Ip
    UDPClient.RemotePort = CLng(Port)
    UDPClient.Connect
    UDPClient.SendData Chr(255) & Chr(255) & Chr(255) & Chr(255) & "rcon " & Pass & " " & Cmd
    
    txtOutput = txtOutput & sStamp & vbCrLf
    txtOutput.SelStart = Len(txtOutput)
    If Len(txtOutput) >= 24000 Then txtOutput = ""
    
    Log (sStamp + vbNewLine)
End Sub

Private Sub Timer1a_Timer()
    On Error Resume Next
    Q3_sendData
End Sub


Private Sub UDPClient_DataArrival(ByVal bytesTotal As Long)
    On Error Resume Next
    
    UDPClient.GetData strData, vbString
    
    strData = Replace(strData, Chr(255) & Chr(255) & Chr(255) & Chr(255) & "print" & vbLf, "")
    strData = Replace(strData, vbLf, vbCrLf)
    strData = StripColours(strData)
    
    If Len(txtOutput) >= 24000 Then txtOutput = ""
    txtOutput = txtOutput & strData
    txtOutput.SelStart = Len(txtOutput)
    
    Log (strData)
    
End Sub

Private Function StripColours(name As String) As String
    Dim i As Integer
    i = 1
    Dim toBeReturned As String
    Dim temp As String
    Dim temp2 As String
    While i < Len(name) + 1
        temp = Mid$(name, i, 1)
        If i = Len(name) Then
            temp2 = i
        Else
            temp2 = Mid$(name, i + 1, 1)
        End If
        
        If temp = "^" And (temp2 >= "0" And temp2 <= "9") Then
            i = i + 1
        Else
            toBeReturned = toBeReturned & temp
        End If
        i = i + 1
    Wend
    StripColours = toBeReturned
    
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    ColumnHeaderClick Me.ListView1, ColumnHeader
End Sub

Public Sub ColumnHeaderClick(LV As ListView, ByVal ColumnHeader As MSComctlLib.ColumnHeader)
  Dim LvIndex As Long
  On Error Resume Next
    LvIndex = LV.Index
    If Err Then
      LvIndex = 0
    End If
  On Error GoTo 0
  LV.ColumnHeaders.Item(LV.SortKey + 1).Icon = 0
  If LV.SortKey = ColumnHeader.Index - 1 Then
    If LV.SortOrder = lvwAscending Then
      LV.SortOrder = lvwDescending
    Else
      LV.SortOrder = lvwAscending
    End If
    SaveSetting "Rcon Unlimited", "settings", "SortOrder" & LV.name & LvIndex, LV.SortOrder
  Else
    LV.SortKey = ColumnHeader.Index - 1
    SaveSetting "Rcon Unlimited", "settings", "SortKey" & LV.name & LvIndex, LV.SortKey
  End If

  LV.SetFocus
  DoEvents
  If LV.ListItems.count > 0 Then
    LV.ListItems(LV.SelectedItem.Index).EnsureVisible
  End If
End Sub

Private Sub winsock1_DataArrival(ByVal bytesTotal As Long)
    On Error GoTo error_handler
    Dim y As String
    Winsock1.GetData y
    Dim temp As String
    Dim ping2, freq As Currency
    Dim itmx As ListItem
      
    If hptimer Then
        QueryPerformanceCounter ping2
        QueryPerformanceFrequency freq
        ping2 = Int(((ping2 - ping1) / freq * 1000))
    Else
        ping2 = ctr
    End If
    
    Timer1.Enabled = False
    Dim rules As String
    Dim players As String
    Dim rule As String
    Dim rvalue As String
    Dim pname As String
    Dim pfrags As String
    Dim pping As String
    ListView1.ListItems.Clear
    
    If (InStr(1, y, "statusResponse", vbBinaryCompare)) > 0 Then
        rules = Mid$(y, InStr(1, y, vbLf, vbBinaryCompare) + 1)
        players = Mid$(rules, InStr(1, rules, vbLf, vbBinaryCompare) + 1)
        rules = Left$(rules, InStr(1, rules, vbLf, vbBinaryCompare))
       
        Dim maxClients As String
        
        While rules <> vbLf
            rules = Mid$(rules, 2)
            rule = Mid$(rules, 1, InStr(1, rules, "\", vbBinaryCompare) - 1)
            rules = Mid$(rules, Len(rule) + 2)
            If InStr(1, rules, "\", vbBinaryCompare) <> 0 Then
                rvalue = Mid$(rules, 1, InStr(1, rules, "\", vbBinaryCompare) - 1)
            Else
                rvalue = Mid$(rules, 1, Len(rules) - 1)
            End If
            rules = Mid$(rules, Len(rvalue) + 1)
            If rule = "sv_hostname" Then lblHostName = rvalue
            If rule = "gamename" Then lblGameName = rvalue
            If rule = "version" Then lblVersion = rvalue
            If rule = "mapname" Then lblMap = rvalue
            If rule = "sv_maxclients" Then maxClients = rvalue
        Wend
        
        While players <> ""
            pfrags = Left$(players, InStr(1, players, " ", vbBinaryCompare) - 1)
            players = Mid$(players, Len(pfrags) + 2)
            pping = Left$(players, InStr(1, players, " ", vbBinaryCompare) - 1)
            players = Mid$(players, Len(pping) + 3)
            pname = Left$(players, InStr(1, players, Chr$(34), vbBinaryCompare) - 1)
            players = Mid$(players, Len(pname) + 3)
            pname = StripColours(pname)
            lv1add pname, pfrags, pping
        Wend
        If ping2 > 200 Then
            lblPing.ForeColor = vbRed
        Else
            lblPing.ForeColor = vbBlack
        End If
        lblPing.Caption = ping2 & " ms"

      
        players = ListView1.ListItems.count
        lblMaxClients.Caption = players & "/" & maxClients
        
        ListView1.HideSelection = True
    End If

Exit Sub

error_handler:
    Timer1.Enabled = False
End Sub

Private Sub Timer1_Timer()
    ctr = ctr + 10
    If ctr > 1200 Then
        Timer1.Enabled = False
        Winsock1.Close
        lblPing.Caption = "Timed out"
    End If
End Sub

Private Sub Timer2_Timer()
    Timer2.Enabled = False
    
    Timer2.Interval = autorefreshtime * 1000
    Timer2.Enabled = True
End Sub

Private Sub lv1add(i1 As Variant, i2 As Variant, i3 As Variant)
        Dim itmx As ListItem
        Set itmx = ListView1.ListItems.Add(, , i1)
        itmx.SubItems(1) = i2
        itmx.SubItems(2) = i3
        Set itmx = Nothing
End Sub

Private Sub Q3_sendData()
    On Error GoTo error_handler
    Winsock1.RemoteHost = Ip
    Winsock1.RemotePort = Port
    If Winsock1.State = 1 Then Winsock1.Close
    ctr = 0
    Winsock1.Bind
    If QueryPerformanceCounter(ping1) Then
        hptimer = True
    Else
        hptimer = False
    End If
    Timer1.Enabled = True
    Winsock1.SendData ("ÿÿÿÿgetstatus")
    
    Exit Sub
error_handler:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error"
End Sub
