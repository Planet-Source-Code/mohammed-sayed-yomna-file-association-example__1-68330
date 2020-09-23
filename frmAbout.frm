VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About...!"
   ClientHeight    =   720
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3225
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   720
   ScaleWidth      =   3225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label lblSite 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "http://www.freewebs.com/msayed/"
      Height          =   195
      Left            =   337
      TabIndex        =   1
      Top             =   360
      Width           =   2550
   End
   Begin VB.Label lblBy 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "By Mohammed Sayed Mohammed"
      Height          =   195
      Left            =   405
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Click()
Unload Me
End Sub

Private Sub lblBy_Click()
Unload Me
End Sub

Private Sub lblSite_Click()
Unload Me
End Sub
