Attribute VB_Name = "modGeneral"
Public Const C_gnAnchoPwd As Byte = 20
Public Const C_gnAnchoSrv As Byte = 20
Public Const C_gnAnchoUsuario As Byte = 20
'Private Declare Function GetComputerName Lib "Kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Public Function BuscaBase(Base As String, Conex As ADODB.Connection) As Boolean
Dim RsAdo As New ADODB.Recordset
Dim c As Integer
On Error GoTo Herr
Set RsAdo = New ADODB.Recordset
BuscaBase = False
RsAdo.Open "EXECUTE SP_DATABASES", Conex, adOpenStatic
If RsAdo.EOF = False Then
   RsAdo.Find "DATABASE_NAME='" & Base & "'"
   c = RsAdo.Bookmark
   BuscaBase = True
End If
Exit Function
Herr:
   BuscaBase = False
End Function

