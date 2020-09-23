Attribute VB_Name = "SubClasser"
Option Explicit

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Const GWL_WNDPROC = (-4)
Dim PrevProc As Long, Handle As Long

Private Const WM_COPYDATA = &H4A
Private Type COPYDATASTRUCT
   dwData As Long
   cbData As Long
   lpData As Long
End Type
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

Public Sub HookForm(FormName As Form)
Handle = FormName.hwnd
PrevProc = SetWindowLong(Handle, GWL_WNDPROC, AddressOf WindowProc)
End Sub

Public Sub UnHookForm()
SetWindowLong Handle, GWL_WNDPROC, PrevProc
End Sub

Public Function WindowProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Select Case uMsg
    Case WM_COPYDATA
        Dim aB() As Byte
        Dim sCst  As COPYDATASTRUCT
        Dim sNewCmd As String
        
        CopyMemory sCst, ByVal lParam, Len(sCst)
        If (sCst.cbData > 0) Then
            ReDim aB(0 To sCst.cbData - 1) As Byte
            CopyMemory aB(0), ByVal sCst.lpData, sCst.cbData
            sNewCmd = StrConv(aB, vbUnicode)
            
            mdiMain.NewWindow sNewCmd
        End If
End Select

WindowProc = CallWindowProc(PrevProc, hwnd, uMsg, wParam, lParam)
End Function
