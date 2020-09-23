VERSION 5.00
Begin VB.Form frmDoc 
   Caption         =   "Untitled"
   ClientHeight    =   4590
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7440
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4590
   ScaleWidth      =   7440
   Begin VB.TextBox txtEdit 
      Height          =   2565
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   120
      Width           =   4095
   End
End
Attribute VB_Name = "frmDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Resize()
On Error Resume Next
txtEdit.Height = Me.ScaleHeight - 240
txtEdit.Width = Me.ScaleWidth - 240
End Sub

Public Sub OpenFile(FilePath As String)
Dim FF As Integer
FF = FreeFile()

Open FilePath For Input As #FF
    txtEdit = Input$(LOF(FF), #FF)
Close #FF

End Sub
