VERSION 5.00
Begin VB.MDIForm mdiMain 
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   Caption         =   "Yomna - File Association Example"
   ClientHeight    =   6135
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   9300
   Icon            =   "mdiMain.frx":0000
   StartUpPosition =   2  'CenterScreen
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Sav&e As"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "&Close"
         Shortcut        =   ^W
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuAss 
         Caption         =   "Associate me with *.YTF"
      End
   End
   Begin VB.Menu mnuhelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About..."
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "mdiMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub MDIForm_Load()
HookForm Me

If Check4Association = True Then mnuAss.Checked = True

If Command$ <> "" Then
    Dim sCmd As String, sExt As String
    sCmd = Trim$(Command$)
    
    If Right$(sCmd, 1) = Chr$(34) Then sCmd = Left$(sCmd, Len(sCmd) - 1)
    If Left$(sCmd, 1) = Chr$(34) Then sCmd = Right$(sCmd, Len(sCmd) - 1)
    
    sExt = Right$(LCase$(sCmd), 4)
    If sExt = ".ytf" Or sExt = ".txt" Or sExt = ".text" Or sExt = ".log" Or sExt = ".ini" Or sExt = ".inf" Then
                Dim frmNewDoc As New frmDoc
                frmNewDoc.OpenFile sCmd
                frmNewDoc.Show
            Else
                MsgBox "Unsupported format , Can not load this file.", vbCritical + vbApplicationModal, "Error...!"
    End If
End If
End Sub

Private Sub MDIForm_Terminate()
SaveTerminate
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
On Error Resume Next
UnHookForm
Mut.CloseMutex
Kill AFP & App.EXEName & ".exe.manifest"
CloseAllForms
End Sub

Private Sub mnuAbout_Click()
frmAbout.Show vbModal, Me
End Sub

Private Sub mnuAss_Click()
If mnuAss.Checked = True Then
        mnuAss.Checked = False
        DeAssociate
    Else
        mnuAss.Checked = True
        Associate
End If
End Sub

Private Sub mnuClose_Click()
'
End Sub

Private Sub mnuExit_Click()
Unload Me
End Sub

Private Sub mnuNew_Click()
Dim frmNewDoc As New frmDoc
frmNewDoc.Show
End Sub

Public Sub NewWindow(FilePath As String)
If FilePath <> "" Then
    Dim sCmd As String, sExt As String
    sCmd = Trim$(FilePath)
    
    If Right$(sCmd, 1) = Chr$(34) Then sCmd = Left$(sCmd, Len(sCmd) - 1)
    If Left$(sCmd, 1) = Chr$(34) Then sCmd = Right$(sCmd, Len(sCmd) - 1)
    
    sExt = Right$(LCase$(sCmd), 4)
    If sExt = ".ytf" Or sExt = ".txt" Or sExt = ".text" Or sExt = ".log" Or sExt = ".ini" Or sExt = ".inf" Then
                Dim frmNewDoc As New frmDoc
                frmNewDoc.OpenFile sCmd
                frmNewDoc.Show
            Else
                MsgBox "Unsupported format , Can not load this file.", vbCritical + vbApplicationModal, "Error...!"
    End If
End If
End Sub

Private Sub mnuOpen_Click()
'
End Sub

Private Sub mnuSave_Click()
'
End Sub

Private Sub mnuSaveAs_Click()
'
End Sub
