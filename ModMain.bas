Attribute VB_Name = "ModMain"
Option Explicit

Public Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Public Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As Long)
Public Declare Function GetCharWidth Lib "gdi32" Alias "GetCharWidthA" (ByVal hdc As Long, ByVal wFirstChar As Long, ByVal wLastChar As Long, lpBuffer As Long) As Long
Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetDesktopWindow Lib "user32" () As Long
Public Const WM_CLOSE = &H10
Private Declare Function SendMessageTimeout Lib "user32" Alias "SendMessageTimeoutA" (ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal fuFlags As Long, ByVal uTimeout As Long, lpdwResult As Long) As Long
Private Const SC_CLOSE = &HF060&
Private Const WM_SYSCOMMAND = &H112
Public Const CLR_INVALID = &HFFFF
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Const PROCESS_ALL_ACCESS = &H1F0FFF

Private Declare Function TerminateProcess& Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long)
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

' API CALLS
Public Declare Function ShellExecute _
    Lib "shell32.dll" _
    Alias "ShellExecuteA" ( _
        ByVal hwnd As Long, _
        ByVal lpOperation As String, _
        ByVal lpFile As String, _
        ByVal lpParameters As String, _
        ByVal lpDirectory As String, _
        ByVal nShowCmd As Long) As Long

Global Ip As String
Global Port As String
Global Pass As String
Public strGame As String

Public Const CMDLIST As String = "cmdlist,serverinfo,status,set,kick,map,sets,say,svsay,vstr,exec,seta,killserver,spdevmap,spmap,devmap,sectorlist,map_restart,dumpuser,systeminfo,heartbeat,vminfo,midiinfo,net_restart,in_restart,writeconfig,changeVectors,quit,meminfo,bind,unbindall,touchFile,dir,path,bindlist,unbind,cvar_restart,cvarlist,reset,setu,toggle,wait,echo"

Public aCmdList() As String
Public hwnd As Long

Public Sub OpenURL(sURL As String)
    On Error Resume Next
    ShellExecute hwnd, "open", sURL, vbNullString, vbNullString, 1
End Sub

Public Function LookupCommand(Cmd As String) As String
    Dim i As Integer
    
    On Error Resume Next
    
    If Cmd <> "" Then
        For i = 0 To UBound(aCmdList)
            If Mid(aCmdList(i), 1, Len(Cmd)) = Cmd Then
                LookupCommand = Mid(aCmdList(i), Len(Cmd) + 1)
                Exit Function
            End If
        Next
    End If
End Function

Public Sub GetCommandList()
    Dim strTemp As String
    Dim sCmdLst As String
    
    ' get the string or the file
    If FileExists(CheckPath(App.Path) & "\commands.dat") Then
        Open CheckPath(App.Path) & "\commands.dat" For Input As #1
        While Not EOF(1)
            Line Input #1, strTemp
            sCmdLst = sCmdLst & Trim(strTemp) & ","
        Wend
        Close #1
        
        aCmdList = Split(sCmdLst, ",")
    Else
        aCmdList = Split(CMDLIST, ",")  ' for emergencies
    End If
End Sub

Public Function FileExists(strPath As String) As Integer
    FileExists = Not (Dir(strPath) = "")
End Function

Public Function CheckPath(Path As String) As String
    If Right(Path, 1) <> "\" And Path <> "" Then Path = Path & "\"
    CheckPath = Path
End Function

Public Sub Log(str As String)
    If FrmMain.mnuautosave.Checked = True Then
        Open CheckPath(App.Path) & "\autosave.log" For Append As #1
        Print #1, str;
        Close #1
    End If
End Sub

Public Function InHistory(cbo As ComboBox) As Boolean
    Dim i As Integer
    
    For i = 0 To cbo.ListCount
        If cbo.List(i) = cbo.Text Then
            InHistory = True
            Exit Function
        End If
    Next
    
    InHistory = False
End Function
