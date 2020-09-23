Attribute VB_Name = "Module1"
Option Explicit

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type RECT
    Left   As Long
    Top    As Long
    Right  As Long
    Bottom As Long
End Type
Private Type WINDOWPLACEMENT
    Length           As Long
    flags            As Long
    showCmd          As Long
    ptMinPosition    As POINTAPI
    ptMaxPosition    As POINTAPI
    rcNormalPosition As RECT
End Type
Private Declare Function GetWindowPlacement Lib "user32" _
    (ByVal hwnd As Long, lpwndpl As WINDOWPLACEMENT) As Long
Private Declare Function SetForegroundWindow Lib "user32" _
    (ByVal hwnd As Long) As Long
Private Declare Function ShowWindow Lib "user32" _
    (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
    
Private Const SW_SHOW = 5
Private Const SW_RESTORE = 9
Private Const SW_SHOWMINIMIZED = 2
Private Const SW_SHOWMAXIMIZED = 3
Private Const SW_SHOWNORMAL = 1
Public Function IsFormLoaded(strFormName As String) As Boolean

    Dim frm As Form
    For Each frm In Forms
        If frm.Name = strFormName Then
            IsFormLoaded = True
            Exit Function
        End If
    Next
    IsFormLoaded = False
End Function
Public Sub RestoreWindow(ByVal hwnd As Long)
' Get the window's state and activate it.
    Dim lpwndpl   As WINDOWPLACEMENT
    Dim lState As Long
    On Error GoTo ErrHandler
    lState = GetWindowPlacement(hwnd, lpwndpl)
    Select Case lpwndpl.showCmd
    Case SW_SHOWMINIMIZED
        Call ShowWindow(hwnd, SW_RESTORE)
    Case SW_SHOWNORMAL, SW_SHOWMAXIMIZED
        Call ShowWindow(hwnd, SW_SHOW)
    End Select
    Call SetForegroundWindow(hwnd)
    Exit Sub
ErrHandler:
    'do noting
End Sub
