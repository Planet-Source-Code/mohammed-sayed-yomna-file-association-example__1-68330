Attribute VB_Name = "Registry"
Option Explicit

Private Const REG_SZ = 1
Private Const REG_BINARY = 3

Public Enum RegClasses
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
    HKEY_PERFORMANCE_DATA = &H80000004
End Enum

Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long

Public Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As RegClasses, ByVal lpSubKey As String) As Long

Private Function RegQueryStringValue(ByVal hKey As Long, ByVal strValueName As String) As String
On Error Resume Next
Dim lResult As Long, lValueType As Long, strBuf As String, lDataBufSize As Long
lResult = RegQueryValueEx(hKey, strValueName, 0, lValueType, ByVal 0, lDataBufSize)
If lResult = 0 Then
    If lValueType = REG_SZ Then
        strBuf = String(lDataBufSize, Chr$(0))
        lResult = RegQueryValueEx(hKey, strValueName, 0, 0, ByVal strBuf, lDataBufSize)
        If lResult = 0 Then
            RegQueryStringValue = Left$(strBuf, InStr(1, strBuf, Chr$(0)) - 1)
        End If
    ElseIf lValueType = REG_BINARY Then
        Dim strData As Integer
        lResult = RegQueryValueEx(hKey, strValueName, 0, 0, strData, lDataBufSize)
        If lResult = 0 Then
            RegQueryStringValue = strData
        End If
    End If
End If
End Function
Public Function GetString(hKey As RegClasses, strPath As String, strValue As String)
Dim Ret As Long
RegOpenKey hKey, strPath, Ret
GetString = RegQueryStringValue(Ret, strValue)
RegCloseKey Ret
End Function
Public Sub SaveString(hKey As RegClasses, strPath As String, strValue As String, strData As String)
Dim Ret As Long
RegCreateKey hKey, strPath, Ret
RegSetValueEx Ret, strValue, 0, REG_SZ, ByVal strData, Len(strData)
RegCloseKey Ret
End Sub
'Public Sub SaveStringLong(hKey As RegClasses, strPath As String, strValue As String, strData As String)
'Dim Ret As Long
'RegCreateKey hKey, strPath, Ret
'RegSetValueEx Ret, strValue, 0, REG_BINARY, strData, 4
'RegCloseKey Ret
'End Sub
Public Sub DelString(hKey As RegClasses, strPath As String, strValue As String)
Dim Ret As Long
RegCreateKey hKey, strPath, Ret
RegDeleteValue Ret, strValue
RegCloseKey Ret
End Sub
