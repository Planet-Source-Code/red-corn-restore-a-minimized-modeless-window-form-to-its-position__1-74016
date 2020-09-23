VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form3"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  '¨t²Î¹w³]­È
   Begin VB.CommandButton Command2 
      Caption         =   "Show Modeless Form2"
      Height          =   420
      Left            =   1170
      TabIndex        =   0
      Top             =   1035
      Width           =   2175
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command2_Click()
    If IsFormLoaded("Form2") Then
        RestoreWindow Form2.hwnd
    Else
        Form2.Show
    End If
End Sub

