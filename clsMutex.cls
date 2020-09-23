VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMutex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'App.PrevInstance usually fails , I made a web search & I found this
'amazing solution By Hesham A. Amin (hspc)
'http://www.codeguru.com/forum/showthread.php?s=&threadid=293730
'
'Example :
'***********************************************
'Set mut = New clsMutex                         |
'If Not mut.CheckMutex("CodeGuru") Then         |
'    MsgBox "Application is running"            |
'    Unload Me                                  |
'End If                                         |
'***********************************************

Private Declare Function CreateMutex Lib "kernel32" Alias "CreateMutexA" (lpMutexAttributes As SECURITY_ATTRIBUTES, ByVal bInitialOwner As Long, ByVal lpName As String) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function ReleaseMutex Lib "kernel32" (ByVal hMutex As Long) As Long

Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

Private Const ERROR_ALREADY_EXISTS = 183&

Private m_hMutex As Long

Public Function CheckMutex(MutexName As String) As Boolean

Dim S As SECURITY_ATTRIBUTES

m_hMutex = CreateMutex(S, 0, MutexName)

If Err.LastDllError = ERROR_ALREADY_EXISTS Then
    CheckMutex = False
Else
    CheckMutex = True
End If

End Function

Public Sub CloseMutex()
ReleaseMutex m_hMutex
CloseHandle m_hMutex
End Sub
