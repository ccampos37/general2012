Attribute VB_Name = "ModPrincipal"
Option Explicit


Public Sub Main()
 On Error GoTo x
   ColorDesHabilitado = &H80000004
   ColorHabilitado = &H80000005
   vgUSUARIO = "ics"
   
   vgCADENAREPORT = "DSN=DESARROLLO3;DSQ=CONTAPRUEBA;UID=SA"
   'vgCADENAREPORT = "DSN=IVAN;DSQ=CONTAPRUEBA;UID=SA"
   
   vgCN = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=contaprueba;Data Source=DESARROLLO3"
   'vgCN = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=contaprueba;Data Source=IVAN"
   
   strConexion = vgCN
   cn.ConnectionString = strConexion
   cn.CursorLocation = adUseClient
   cn.ConnectionTimeout = 30
   cn.Open
   
   vgCG = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=MARFICE;Data Source=DESARROLLO3"
   'vgCG = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=MARFICE;Data Source=IVAN"
   
   cg.ConnectionString = vgCG
   cg.CursorLocation = adUseClient
   cg.ConnectionTimeout = 30
   cg.Open
   
   MesProceso = 8
   AnnoProceso = 2002
   
   Call ParametroCuenta
   MDIPrincipal.Show
   
x:
   If Err Then
       MsgBox cn.Errors(0).NativeError & "-" & cn.Errors(0).Description
       Err = 0
       Resume Next
   End If
End Sub

Public Sub ParametroCuenta()
 Dim rs As ADODB.Recordset
 Dim cuenta As String
 Dim I As Integer
 Dim j As Integer
 Dim num As Integer
 Set rs = New ADODB.Recordset
 
  Set rs = cn.Execute("SELECT sistemaconfiguracuenta FROM ct_sistema")
  If Not (rs.BOF Or rs.EOF) Then
    cuenta = Trim(rs(0))
    For I = 1 To Len(cuenta)
      If Mid(cuenta, I, 1) = "*" Then num = num + 1
    Next
    ReDim vg_aNIVELES(Len(cuenta) - num)
    j = 0
    For I = 1 To Len(cuenta) Step 2
      vg_aNIVELES(j) = Mid(cuenta, I, 1)
      j = j + 1
    Next
    vgNUMNIVELES = Len(cuenta) - num
  End If
  Set rs = Nothing
End Sub
