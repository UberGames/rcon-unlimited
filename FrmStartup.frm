VERSION 5.00
Begin VB.Form FrmStartup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rcon Unlimited"
   ClientHeight    =   4155
   ClientLeft      =   720
   ClientTop       =   1350
   ClientWidth     =   6000
   Icon            =   "FrmStartup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   277
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   400
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdConnect 
      Caption         =   "Connect"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   13
      ToolTipText     =   "Click Here to loadup the Main Rcon Unlimited Window"
      Top             =   3720
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "Server Detials"
      Height          =   2055
      Left            =   0
      TabIndex        =   0
      Top             =   1560
      Width           =   5895
      Begin VB.CommandButton CmdRemove 
         Caption         =   "Delete Server"
         Height          =   255
         Left            =   4560
         TabIndex        =   12
         ToolTipText     =   "Delete this server from Rcon Unlimited's Memory"
         Top             =   1680
         Width           =   1215
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "Save Settings"
         Height          =   255
         Left            =   3240
         TabIndex        =   11
         ToolTipText     =   "Save this server to Rcon Unlimited's memory"
         Top             =   1680
         Width           =   1215
      End
      Begin VB.CheckBox chkDefault 
         Caption         =   "Set as Default Server"
         Height          =   255
         Left            =   3360
         TabIndex        =   10
         ToolTipText     =   "If this is checked, each time Rcon Unlimited loads this server will be selected automaticly."
         Top             =   960
         Width           =   2295
      End
      Begin VB.CheckBox ChkPass 
         Caption         =   "Remember Rcon Password"
         Height          =   255
         Left            =   3360
         TabIndex        =   9
         ToolTipText     =   "Do you want Rcon Unlimited to remember the password, or leave it?"
         Top             =   1320
         Width           =   2295
      End
      Begin VB.TextBox txtPass 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1440
         PasswordChar    =   "*"
         TabIndex        =   8
         ToolTipText     =   "What is the servers Rcon Password?"
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox txtPort 
         Height          =   285
         Left            =   1440
         TabIndex        =   6
         ToolTipText     =   "Put the port of the server in here."
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox txtIP 
         Height          =   285
         Left            =   1440
         TabIndex        =   5
         ToolTipText     =   "Put the IP address of the server in here."
         Top             =   600
         Width           =   3375
      End
      Begin VB.ComboBox CboNick 
         Height          =   315
         Left            =   1440
         TabIndex        =   4
         ToolTipText     =   "Enter a name here. This helps you identify which server is which."
         Top             =   240
         Width           =   3375
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Rcon Password:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   " Port:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "IP / Address:"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   1215
      End
   End
   Begin VB.Image Image1 
      Height          =   1500
      Left            =   0
      Picture         =   "FrmStartup.frx":210A
      Top             =   0
      Width           =   6000
   End
   Begin VB.Label lblCopyright 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "© 2003 UberGames. All Rights Reserved."
      ForeColor       =   &H80000010&
      Height          =   255
      Left            =   1800
      TabIndex        =   14
      Top             =   3720
      Width           =   4095
   End
End
Attribute VB_Name = "FrmStartup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub LoadCombo()
    On Error GoTo error:
    
    Dim endoflist As Boolean
    Dim ServerNo As Integer
    
    CboNick.Clear
    servers = Split(GetSetting("Rcon Unlimited", "settings", "servers"), "¬¬")
    
    endoflist = False
    ServerNo = 0
    
    While endoflist = False
        If servers(ServerNo) = "" Then
            endoflist = True
        Else
            CboNick.AddItem (servers(ServerNo))
        End If
        ServerNo = ServerNo + 1
    Wend
    
    Exit Sub
    
error:
    CboNick.AddItem ("New Server")
    CboNick.Text = "New Server"
    chkDefault.Value = 1
End Sub

Public Sub LoadSettings()
    If GetSetting("Rcon Unlimited", "IP", CboNick.Text) <> "" Then
        address = Split(GetSetting("Rcon Unlimited", "IP", CboNick.Text), ":")
        txtIP = address(0)
        txtPort = address(1)
        
        txtPass = GetSetting("Rcon Unlimited", "Rcon", CboNick.Text)
        ChkPass.Value = GetSetting("Rcon Unlimited", "Remember", CboNick.Text)
        
        If CboNick.Text = GetSetting("Rcon Unlimited", "settings", "default") Then
            chkDefault.Value = 1
        Else
            chkDefault.Value = 0
        End If
    End If
End Sub
Private Sub CboNick_Click()
    LoadSettings
End Sub

Private Sub CmdConnect_Click()
    If txtIP = "" Then
        MsgBox "Please enter the IP Address of the server.", vbCritical, "Rcon Unlimited"
        txtIP.SetFocus
        Exit Sub
    End If
    
    If txtPort = "" Then
        MsgBox "Please enter the port that the server runs on.", vbCritical, "Rcon Unlimited"
        txtPort.SetFocus
        Exit Sub
    End If
    
    If txtPass = "" Then
        MsgBox "Please enter the rcon password.", vbCritical, "Rcon Unlimited"
        txtPass.SetFocus
        Exit Sub
    End If
    
    Ip = txtIP
    Port = txtPort
    Pass = txtPass
    
    FrmMain.Show
    FrmMain.Caption = "Rcon Unlimited v1.0 - " + CboNick.Text
    Unload Me
End Sub

Private Sub CmdRemove_Click()
    Dim endoflist As Boolean
    Dim ServerNo As Integer
    Dim newServers As String
    
    On Error Resume Next
    
    servers = Split(GetSetting("Rcon Unlimited", "settings", "servers"), "¬¬")
    endoflist = False
    ServerNo = 0
    
    While endoflist = False
        If servers(ServerNo) = "" Then
            endoflist = True
        End If
        ServerNo = ServerNo + 1
    Wend
    
    If ServerNo < 3 Then
        MsgBox "You cannot delete all servers.", vbCritical, "Rcon Unlimited"
        Exit Sub
    End If
    
    If GetSetting("Rcon Unlimited", "settings", "default") = CboNick.Text Then
        MsgBox "You cannot delete the default server. Choose another server as default then you can delete this one."
        Exit Sub
    End If
    
    If MsgBox("Are you sure you wish to delete this server?", vbYesNo, "Rcon Unlimited") = vbNo Then
        Exit Sub
    End If
    
    endoflist = False
    ServerNo = 0
    newServers = ""
    
    While endoflist = False
        If servers(ServerNo) = "" Then
            endoflist = True
        Else
            If servers(ServerNo) <> CboNick.Text Then
                newServers = newServers + servers(ServerNo) + "¬¬"
            End If
        End If
        ServerNo = ServerNo + 1
    Wend
    
    SaveSetting "Rcon Unlimited", "settings", "servers", newServers
    
    DeleteSetting "Rcon Unlimited", "IP", CboNick.Text
    DeleteSetting "Rcon Unlimited", "Remember", CboNick.Text
    DeleteSetting "Rcon Unlimited", "Rcon", CboNick.Text
    
    LoadCombo
    CboNick.Text = GetSetting("Rcon Unlimited", "settings", "default")
    LoadSettings
End Sub

Private Sub CmdSave_Click()
    Dim endoflist As Boolean
    Dim InList As Boolean
    Dim ServerNo As Integer
    
    If CboNick.Text = "" Then
        MsgBox "Please enter a name for this server, so you can idenify it in the list.", vbCritical, "Rcon Unlimited"
        Exit Sub
    End If
    
    If txtIP = "" Then
        MsgBox "Please enter the IP Address of the server.", vbCritical, "Rcon Unlimited"
        Exit Sub
    End If
    
    If txtPort = "" Then
        MsgBox "Please enter the port that the server runs on.", vbCritical, "Rcon Unlimited"
        Exit Sub
    End If
    
    If GetSetting("Rcon Unlimited", "IP", CboNick.Text) <> "" Then
        If MsgBox("A server with this name allready exists, do you want to continue saving. This would overwrite the other server.", vbYesNo, "Rcon Unlimited") = vbNo Then
            Exit Sub
        End If
    End If
    
    SaveSetting "Rcon Unlimited", "IP", CboNick.Text, txtIP + ":" + txtPort
    
    If ChkPass.Value = 1 Then
        SaveSetting "Rcon Unlimited", "Rcon", CboNick.Text, txtPass
        SaveSetting "Rcon Unlimited", "Remember", CboNick.Text, "1"
    Else
        SaveSetting "Rcon Unlimited", "Rcon", CboNick.Text, ""
        SaveSetting "Rcon Unlimited", "Remember", CboNick.Text, "0"
    End If
    
    servers = Split(GetSetting("Rcon Unlimited", "settings", "servers"), "¬¬")
    endoflist = False
    InList = False
    ServerNo = 0
    
    While endoflist = False
        If servers(ServerNo) = "" Then
            endoflist = True
        Else
            If servers(ServerNo) = CboNick.Text Then
                InList = True
            End If
        End If
        ServerNo = ServerNo + 1
    Wend
   
    If InList = False Then
        SaveSetting "Rcon Unlimited", "settings", "servers", GetSetting("Rcon Unlimited", "settings", "servers") + CboNick.Text + "¬¬"
    End If
    
    
    If chkDefault.Value = 1 Then
        SaveSetting "Rcon Unlimited", "settings", "default", CboNick.Text
    End If
    
    one$ = CboNick.Text
    LoadCombo
    CboNick.Text = one$
End Sub

Private Sub Form_Load()
    If GetSetting("Rcon Unlimited", "settings", "servers") = "" Then
        SaveSetting "Rcon Unlimited", "settings", "default", "New Server"
        SaveSetting "Rcon Unlimited", "IP", "New Server", "127.0.0.1:27960"
        SaveSetting "Rcon Unlimited", "Rcon", "New Server", ""
        SaveSetting "Rcon Unlimited", "Remember", "New Server", "0"
        SaveSetting "Rcon Unlimited", "settings", "servers", "New Server¬¬"
    End If
    
    
    LoadCombo
    CboNick.Text = GetSetting("Rcon Unlimited", "settings", "default")
    LoadSettings
End Sub
