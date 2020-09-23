VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  '¨t²Î¹w³]­È
   Begin VB.CommandButton Command2 
      Caption         =   "Show Modeless Form3"
      Height          =   420
      Left            =   1350
      TabIndex        =   1
      Top             =   1575
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Show Modeless Form2"
      Height          =   420
      Left            =   1350
      TabIndex        =   0
      Top             =   720
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command1_Click()
    If IsFormLoaded("Form2") Then
        RestoreWindow Form2.hwnd
    Else
        Form2.Show
    End If
End Sub

Private Sub Command2_Click()
    If IsFormLoaded("Form3") Then
        RestoreWindow Form3.hwnd
    Else
        Form3.Show
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim frm As Form
    For Each frm In Forms
        Unload frm
    Next
End Sub
