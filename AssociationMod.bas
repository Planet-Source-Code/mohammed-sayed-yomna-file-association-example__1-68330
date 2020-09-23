Attribute VB_Name = "AssociationMod"
Option Explicit
Private Const ShellTXT = "Open with Yomna Association Example"

Private Const HWND_BROADCAST As Long = &HFFFF&
Private Const WM_WININICHANGE As Long = &H1A
Private Const WM_SETTINGCHANGE As Long = WM_WININICHANGE
Private Const SPI_SETNONCLIENTMETRICS As Long = 42
Private Const SMTO_ABORTIFHUNG As Long = &H2
Private Const SUCCESS As Long = 0
Private Declare Function SendMessageTimeout Lib "user32" Alias "SendMessageTimeoutA" (ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As String, ByVal fuFlags As Long, ByVal uTimeout As Long, lpdwResult As Long) As Long

Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Public Sub Associate()
SaveString HKEY_CLASSES_ROOT, "." & "YTF", "", "YTF"
SaveString HKEY_CLASSES_ROOT, "YTF", "", "Yomna Text File"
SaveString HKEY_CLASSES_ROOT, "YTF" & "\shell\" & ShellTXT & "\command", "", AFP & App.EXEName & ".exe %1"
SaveString HKEY_CLASSES_ROOT, "YTF" & "\DefaultIcon\", "", AFP & App.EXEName & ".exe,1"

RefreshIcons
End Sub

Public Sub DeAssociate()
DelString HKEY_CLASSES_ROOT, "YTF", ""
DelString HKEY_CLASSES_ROOT, "YTF" & "\shell\" & ShellTXT & "\command", ""
DelString HKEY_CLASSES_ROOT, "YTF" & "\DefaultIcon\", ""

RegDeleteKey HKEY_CLASSES_ROOT, ".YTF"
RegDeleteKey HKEY_CLASSES_ROOT, "YTF" & "\shell\" & ShellTXT & "\command"
RegDeleteKey HKEY_CLASSES_ROOT, "YTF" & "\DefaultIcon\"
RegDeleteKey HKEY_CLASSES_ROOT, "YTF" & "\shell\" & ShellTXT
RegDeleteKey HKEY_CLASSES_ROOT, "YTF" & "\shell\"
RegDeleteKey HKEY_CLASSES_ROOT, "YTF"

RefreshIcons
End Sub

Private Sub RefreshIcons()
'This is very dangerous , I am not responsible about
'any damage occur to your system.

'Please Back up your registry first.

'Some machines does not respond to this routine immediately , why ???!!! ask Microsoft ;)
Dim OrgSize As Long, NewSize As Long

OrgSize = GetString(HKEY_CURRENT_USER, "Control Panel\Desktop\WindowMetrics", "Shell Icon Size")
NewSize = OrgSize - 1

If NewSize > 0 Then
    SaveString HKEY_CURRENT_USER, "Control Panel\Desktop\WindowMetrics", "Shell Icon Size", CStr(NewSize)
    
    LockWindowUpdate GetDesktopWindow()
    SendMessageTimeout HWND_BROADCAST, WM_SETTINGCHANGE, SPI_SETNONCLIENTMETRICS, 0&, SMTO_ABORTIFHUNG, 10000&, SUCCESS
    LockWindowUpdate 0
    
    SaveString HKEY_CURRENT_USER, "Control Panel\Desktop\WindowMetrics", "Shell Icon Size", CStr(OrgSize)
    
    LockWindowUpdate GetDesktopWindow()
    SendMessageTimeout HWND_BROADCAST, WM_SETTINGCHANGE, SPI_SETNONCLIENTMETRICS, 0&, SMTO_ABORTIFHUNG, 10000&, SUCCESS
    LockWindowUpdate 0
End If
End Sub

Public Function Check4Association() As Boolean
Dim sStr1 As String, sStr2 As String
sStr1 = GetString(HKEY_CLASSES_ROOT, ".YTF", "")
sStr2 = GetString(HKEY_CLASSES_ROOT, "YTF" & "\shell\" & ShellTXT & "\command", "")

If sStr1 = "YTF" And sStr2 = AFP & App.EXEName & ".exe %1" Then
        Check4Association = True
    Else
        Check4Association = False
End If

End Function
