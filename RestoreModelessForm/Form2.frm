VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  '¨t²Î¹w³]­È
   Begin VB.CommandButton Command2 
      Caption         =   "Show Modeless Form3"
      Height          =   420
      Left            =   1260
      TabIndex        =   0
      Top             =   1080
      Width           =   2175
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command2_Click()
    If IsFormLoaded("Form3") Then
        RestoreWindow Form3.hwnd
    Else
        Form3.Show
    End If
End Sub

