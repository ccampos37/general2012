Attribute VB_Name = "Funciones"
Option Explicit
Public Cadenabusca As String
Public cn As New ADODB.Connection
Public cg As New ADODB.Connection

'Variables parametros de componente usercontrol1
'Public a_Array(0 To 12, 0 To 12)

'Variables de acceso de usuario
Public g_usuario As String
Public g_ptoventa As String
Public conexion As String

'Constantes de mensajes para visualizar
Public Const MsgEdit = "No Existen Datos para Editar.. "
Public Const MsgGraba = "Datos Grabados satisfactoriamente...."
Public Const MsgElim = "No Existen Datos a Eliminar.."
Public Const MsgAdd = "Los datos ya existen...Verifique!!!"
Public Const MsgTitle = "AVISO"
'REPORTES
Public Const RutaRep = "\\desarrollo\librerias_controles\Reportes\"
Public Const RutaRepProc = "\\desarrollo\librerias_controles\Reportes\Procesos\"
Public Const CadenaRep = "DSN=DESARROLLO;DSQ=Ventas_Prueba;UID=pirata"


Public Sub Main()
   On Error GoTo nerror
   
   g_usuario = "elozano"
   
   conexion = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=pirata;Initial Catalog=Ventas_Prueba;Data Source=DESARROLLO"
   cn.ConnectionString = conexion
   cn.CursorLocation = adUseClient
   cn.Open
   
   cg.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=pirata;Initial Catalog=MARFICE;Data Source=DESARROLLO"
   cg.CursorLocation = adUseClient
   cg.Open

   MDIMain.Show
   
nerror:
   If Err Then
       MsgBox cn.Errors(0).NativeError & "-" & cn.Errors(0).Description
       Err = 0
       Resume Next
   End If
End Sub


Public Function ValidaDato(ByRef aDato) As String
   If IsNull(aDato) Then
      ValidaDato = "" & aDato
   Else
      ValidaDato = Trim(aDato)
   End If
   
End Function

Public Function MostrarForm(pVentana As Form, pPos As String)
   pVentana.Icon = LoadPicture(App.Path & "\factu.ico")
   If pPos = "C" Then
     pVentana.Left = (Screen.Width - pVentana.Width) / 2
     pVentana.Top = (Screen.Height - pVentana.Height) / 2
   ElseIf pPos = "I" Then
      pVentana.Left = 300
      pVentana.Top = 300
   End If

End Function
Public Function EliminaReg(vcon As ADODB.Connection, xtabla As String, xCondi As String) As Integer
    On Error GoTo nerror
    If MsgBox("Desea Eliminar el Registro?", vbYesNo) = vbYes Then
       vcon.Execute "Delete From " & Trim(xtabla) & " Where " & xCondi
       EliminaReg = 1
    Else
       EliminaReg = 0
    End If
    
nerror:
   If Err Then
      MsgBox vcon.Errors(0).Number & "-" & vcon.Errors(0).Description, vbInformation, MsgTitle
      Err = 0
      Resume Next
   End If
    
End Function

Public Function Limpiartexto(MBox As Object, ninicio As Integer, nfin As Integer, Optional Noincluir1, Optional Noincluir2 As Integer)
 Dim J As Integer
   
   For J = ninicio To nfin
      Select Case J
         Case J <> Noincluir1
            MBox(J) = ""
         Case J <> Noincluir2
            MBox(J) = ""
      End Select
   Next J
End Function
Public Function Listar(vcon As ADODB.Connection, vTabla, xDbGrid As TDBGrid, xCampos, xOrden, ByRef xlongicampo, Optional xCondi As String)
  Dim lista As Integer
  
  On Error GoTo Elista
  If IsNull(xCondi) Or Len(Trim(xCondi)) = 0 Then
    Set xDbGrid.DataSource = vcon.Execute("Select " & xCampos & " From " & Trim(vTabla) & " Order By " & xOrden)
  Else
    If Len(Trim(xOrden)) = 0 Then
      Set xDbGrid.DataSource = vcon.Execute("Select " & xCampos & " From " & Trim(vTabla) & " Where " & xCondi)
    Else
      Set xDbGrid.DataSource = vcon.Execute("Select " & xCampos & " From " & Trim(vTabla) & " Where " & xCondi & " Order By " & xOrden)
   End If
  End If
  If xlongicampo(1) > 0 Then
    For lista = 1 To UBound(xlongicampo)
        xDbGrid.Columns(lista - 1).Width = xlongicampo(lista)
    Next lista
  End If
  xDbGrid.Refresh
  
  
Elista:
   If Err Then
        MsgBox vcon.Errors(0).NativeError & "-" & vcon.Errors(0).Description, vbInformation
        Err = 0
        Resume Next
    End If
End Function



Public Function ActivaTab(pos, nro, xcontrol As SSTab)
   Dim J As Integer
   
   For J = 0 To nro
      xcontrol.TabEnabled(J) = False
   Next J
   xcontrol.TabEnabled(pos) = True
   xcontrol.Tab = pos
End Function



Public Function Chequeo(vcon As ADODB.Connection, vsql As String) As Integer
   Dim rs As New ADODB.Recordset
   On Error GoTo nerror

   Set rs = vcon.Execute(vsql)
   If rs.RecordCount > 0 Then
      Chequeo = 1
   Else
      Chequeo = 0
   End If
   Set rs = Nothing
   
nerror:
  If Err Then
       MsgBox vcon.Errors(0).Number & "-" & vcon.Errors(0).Description, vbInformation, MsgTitle
       Err = 0
       Resume Next
  End If

End Function

Public Function Seguir(MBox As Object, ntecla As Integer)
    If ntecla = 13 Then
        SendKeys "{tab}"
    End If
End Function


Public Sub Imprimir(cNombreReporte As String)
Dim busca As New dll_apisgen.dll_apis
On Error GoTo Errores
   
   ''''''''''''''''''''''''''''''''''''''''''''''''''''
        'Agregar:
   MDIPrincipal.OCrystalReport.Connect = "DSN=" & busca.LeerIni(App.Path & "\Camtex.ini", "Bventas", "dserver", "") & ";" & _
                   "DSQ=" & CStr(cn.DefaultDatabase) & ";" & _
                   "UID=" & busca.LeerIni(App.Path & "\Camtex.ini", "Bventas", "duser", "")
   ''''''''''''''''''''''''''''''''''''''''''''''''''''
   
   MDIPrincipal.OCrystalReport.Destination = crptToWindow
   MDIPrincipal.OCrystalReport.WindowState = crptMaximized
   MDIPrincipal.OCrystalReport.ReportFileName = RutaRep & cNombreReporte
   MDIPrincipal.OCrystalReport.Formulas(0) = "Empresa='" & g_DetalleEmpresa & "'"
   MDIPrincipal.OCrystalReport.DiscardSavedData = True
   MDIPrincipal.OCrystalReport.Action = 1
   
   Exit Sub
   
Errores:
  MsgBox "Error inesperado: " & Err.Number & "  " & Err.Description, vbExclamation
  Err = 0
  Exit Sub
End Sub


