Attribute VB_Name = "MainMod"
Option Explicit

Private Type tagInitCommonControlsEx
   lngSize As Long
   lngICC As Long
End Type
Private Declare Function InitCommonControlsEx Lib "comctl32.dll" (iccex As tagInitCommonControlsEx) As Boolean
Private Const ICC_USEREX_CLASSES = &H200

Public m_bInIDE As Boolean
Public Const SEM_FAILCRITICALERRORS = &H1
Public Const SEM_NOGPFAULTERRORBOX = &H2
Public Const SEM_NOOPENFILEERRORBOX = &H8000
Public Declare Function SetErrorMode Lib "kernel32" (ByVal wMode As Long) As Long

Private Declare Function SetCurrentDirectory Lib "kernel32" Alias "SetCurrentDirectoryA" (ByVal lpPathName As String) As Long

Private Declare Function PathFileExists Lib "shlwapi.dll" Alias "PathFileExistsA" (ByVal pszPath As String) As Long

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function IsIconic Lib "user32" (ByVal hwnd As Long) As Long

Private Const WM_COPYDATA = &H4A
Private Const SMTO_NORMAL = &H0
Private Const SUCCESS As Long = 0
Private Type COPYDATASTRUCT
   dwData As Long
   cbData As Long
   lpData As Long
End Type
Private Declare Function SendMessageTimeout Lib "user32" Alias "SendMessageTimeoutA" (ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, lParam As Any, ByVal fuFlags As Long, ByVal uTimeout As Long, lpdwResult As Long) As Long

Private Const WM_SYSCOMMAND = &H112
Private Const SC_RESTORE = &HF120&
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long

Public AFP As String 'Application full path
Public Mut As New clsMutex

Public Function FileExists(FilePath As String) As Boolean
FileExists = CBool(PathFileExists(FilePath))
End Function

Private Function InitCommonControlsVB() As Boolean
   On Error Resume Next
   Dim iccex As tagInitCommonControlsEx
   
   With iccex
       .lngSize = LenB(iccex)
       .lngICC = ICC_USEREX_CLASSES
   End With
   
   InitCommonControlsEx iccex
   InitCommonControlsVB = (Err.Number = 0)
   On Error GoTo 0
End Function

Public Property Get InIDE() As Boolean
   Debug.Assert (IsInIDE())
   InIDE = m_bInIDE
End Property

Private Function IsInIDE() As Boolean
   m_bInIDE = True
   IsInIDE = m_bInIDE
End Function

Public Sub Main()
On Error Resume Next
Dim sCmd As String

sCmd = Command$

AFP = App.Path
If Right$(AFP, 1) <> "\" Then AFP = AFP & "\"

SetCurrentDirectory AFP

If Not InIDE() Then
    If FileExists(AFP & App.EXEName & ".exe.manifest") = False Then
        Dim FF As Integer
        FF = FreeFile()
        
        Open AFP & App.EXEName & ".exe.manifest" For Output As #FF
            Print #FF, "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "UTF-8" & Chr(34) & " standalone=" & Chr(34) & "yes" & Chr(34) & "?><assembly xmlns=" & Chr(34) & "urn:schemas-microsoft-com:asm.v1" & Chr(34) & " manifestVersion=" & Chr(34) & "1.0" & Chr(34) & "><assemblyIdentity version=" & Chr(34) & "1.0.0.0" & Chr(34) & " processorArchitecture=" & Chr(34) & "X86" & Chr(34) & " name=" & Chr(34) & "XP Styles Manifest" & Chr(34) & " type=" & Chr(34) & "win32" & Chr(34) & " /><description>M.S.co.</description><dependency><dependentAssembly><assemblyIdentity type=" & Chr(34) & "win32" & Chr(34) & " name=" & Chr(34) & "Microsoft.Windows.Common-Controls" & Chr(34) & " version=" & Chr(34) & "6.0.0.0" & Chr(34) & " processorArchitecture=" & Chr(34) & "X86" & Chr(34) & " publicKeyToken=" & Chr(34) & "6595b64144ccf1df" & Chr(34) & " language=" & Chr(34) & "*" & Chr(34) & " /></dependentAssembly></dependency></assembly>"
        Close #FF
        
        'Restart to enable the XP Styles.
        If sCmd <> "" Then
                Shell AFP & App.EXEName & ".exe " & sCmd, vbNormalFocus: Exit Sub
            Else
                Shell AFP & App.EXEName & ".exe", vbNormalFocus: Exit Sub
        End If
    End If
    
    Set Mut = New clsMutex
    If Mut.CheckMutex("YomnaFAE") = False Then
        Dim ohWnd As Long
        ohWnd = FindWindow(vbNullString, "Yomna - File Association Example")
        If ohWnd <> 0 Then
                ActivateWindow ohWnd
                If Command$ <> "" Then
                    Dim aB() As Byte
                    Dim sCst  As COPYDATASTRUCT
                    
                    aB = StrConv(Command$, vbFromUnicode)
                    sCst.dwData = 0
                    sCst.cbData = UBound(aB) + 1
                    sCst.lpData = VarPtr(aB(0))
                    
                    SendMessageTimeout ohWnd, WM_COPYDATA, 0, sCst, SMTO_NORMAL, 10000, SUCCESS
                    
                End If
            Else
                InitCommonControlsVB
                mdiMain.Show
                Exit Sub
        End If
        End
    End If
    
    InitCommonControlsVB
    
End If

mdiMain.Show

End Sub

Private Sub ActivateWindow(Handle As Long)
If CBool(IsIconic(Handle)) = True Then
    PostMessage Handle, WM_SYSCOMMAND, SC_RESTORE, 0&
End If

SetForegroundWindow Handle
BringWindowToTop Handle
End Sub

Public Sub SaveTerminate()
On Error Resume Next

If Not InIDE() Then
   SetErrorMode SEM_NOGPFAULTERRORBOX
End If
End Sub

Public Sub CloseAllForms()
On Error Resume Next

Dim Frm As Form

For Each Frm In Forms
    Unload Frm
    Set Frm = Nothing
Next
End Sub
