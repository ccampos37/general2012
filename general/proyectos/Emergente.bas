Attribute VB_Name = "Emergente"
Option Explicit
Private Declare Function SetWindowPos Lib "User32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function FindWindow Lib "User32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetWindowRect Lib "User32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetWindowLong Lib "User32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "User32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "User32" (ByVal hwnd As Long, ByVal crey As Byte, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_LAYERED = &H80000
Private Const LWA_ALPHA = &H2&
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_SHOWWINDOW = &H40
Private Const HWND_TOPMOST = -1
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Sub SlideForm(FRM As Form, Direccion As Long, LEVEL As Byte)
Dim Posicion As Integer
Dim Tamaño As Integer
Dim hwnd As Long
Dim res As Long
Dim buffRECT As RECT

hwnd& = FindWindow("Shell_TrayWnd", "")
If hwnd > 0 Then
res = GetWindowRect(hwnd, buffRECT)
If res > 0 Then
Tamaño = CStr(buffRECT.Bottom - buffRECT.Top) * 15
If buffRECT.Left <= 0 And buffRECT.Top > 0 Then Posicion = 1
If buffRECT.Left > 0 And buffRECT.Top <= 0 Then Posicion = 2: Tamaño = (buffRECT.Right - buffRECT.Left) * 15
If buffRECT.Left <= 0 And buffRECT.Top <= 0 And buffRECT.Right < 600 Then Posicion = 3: Tamaño = buffRECT.Right * 15
If buffRECT.Left <= 0 And buffRECT.Top <= 0 And buffRECT.Right > 600 Then Posicion = 4
End If
Else
Posicion = 1
End If

res = SetWindowPos(FRM.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOACTIVATE Or SWP_SHOWWINDOW)
Call SetWindowLong(FRM.hwnd, GWL_EXSTYLE, GetWindowLong(FRM.hwnd, GWL_EXSTYLE) Or WS_EX_LAYERED)
Call SetLayeredWindowAttributes(FRM.hwnd, 0, LEVEL, LWA_ALPHA)

If Direccion = 0 Then
FRM.Height = 0
Select Case Posicion
Case 1
FRM.Move Screen.Width - FRM.Width, Screen.Height - FRM.Height - Tamaño
Case 2
FRM.Move Screen.Width - FRM.Width - Tamaño, Screen.Height - FRM.Height
Case 3
FRM.Move Tamaño, Screen.Height - FRM.Height
Case 4
FRM.Move Screen.Width - FRM.Width, Tamaño
End Select
Do Until FRM.Height = 2000 ' la altura que se quiera
DoEvents
FRM.Height = FRM.Height + 1
If Not Posicion = 4 Then FRM.Top = FRM.Top - 1
Loop
Else
Do Until FRM.Height = 520
DoEvents
FRM.Height = FRM.Height - 1
If Not Posicion = 4 Then FRM.Top = FRM.Top + 1
Loop
Unload FRM
End If
End Sub
