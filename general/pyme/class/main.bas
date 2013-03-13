Attribute VB_Name = "inicio"
'Sistema de Contabilidad Version 1.0
'Visual Basic 6.0 y SQL - 2000
Option Explicit

'Declaraciones de Variables
'Variables Globales

Public Const ColorHabilitado = &H80000004
Public Const ColorDesHabilitado = &H80000005

Public SQL As String

'Constantes de mensajes para visualizar

Public Const MsgEdit = "No Existen Datos para Editar.. "
Public Const MsgGraba = "Datos Grabados satisfactoriamente...."
Public Const MsgElim = "No Existen Datos a Eliminar.."
Public Const MsgAdd = "Los datos ya existen...Verifique!!!"
Public Const MsgTitle = "AVISO"
Public Const Msg29 = "Debe Ingresar Numeros"
Public mensaje1 As String

Public VGCNx As ADODB.Connection             'Conexion de la BD empresa
Public VGcnxCT As ADODB.Connection        'Conexion de Contabilidad
Public VGgeneral As ADODB.Connection      'Conexion de la BD Generales
Public VGconfig As ADODB.Connection      'Conexion de la BD de configuracion

Public VgActivalogin As Long                 'Activa login una sola vez
Public Cuentacodigo As String
Public VGactulizodoc As Boolean
Public VGformatofecha As String

Public VGflaglimpia As Boolean               'Flag que Limpia
Public VGMoverRegistro As Boolean            'Flag al mover el registro
Public VGCommandoSP As ADODB.Command         'De Comando
Public VGDllgeneral As dllgeneral.dll_general 'Dll de Algunas funciones
Public VGdllApi As dll_apisgen.dll_apis
Public VGvarVerifica As Boolean              'Flag Verifica que transaccion es OK (Grabar ,Etc)
Public VGErrorString As String               'Almacena el Error el que hubo en alguna transaccion
Public VGValorCambio As Double               'Almacena el valor del Cambio
Public VGstrConexion As String               'Cadena de Conexion
Public VgMostrar As Boolean
Public vgcont As Integer


Public VGUsuario As String
Public VGComputer As String                  'Nombre de la computadora
Public VGOrden As String
Public VGCODEMPRESA As String * 3          'Codigo de la compannia
Public VGfecha As Date
Public vgsalir As Boolean

Public vgCADENAREPORT2 As String              'Cadena de Reportes Base Marfice
Public Const NUMMAGICO As Integer = 5

Private Type Parametrosdecostos
    
    BaseOrigen As String
    mesesreferencia As Integer
    codigopersonalplantilla As String * 2
    mesdecierre As String * 6
    
    minimoretencion As Double
    NomEmpresa As String
    RucEmpresa As String
    empresacodigo As String * 2
    
    direccionempresa As String
    impresionalta As Boolean
    ctascompra As String
    Auxaut As Boolean
    

    
    CpTiplan As String
    CpOficina As String
    permite_tc As Boolean
    sistemaactivaccostos As Boolean
    sistemaasientoenlinea As Boolean
    sistemactrlgastos As Boolean
    sistemamultiempresas As Boolean
    sistemabancarizacion As Boolean
    sistemabancarizacion01 As Double
    sistemabancarizacion02 As Double
    
    controlaestadosrendicion  As Boolean
    diasatrazorendicion As Integer
    diacierrerendicion As Integer

        
End Type

Private Type parametrosdesistema
    TablaCabcomprob As String
    TablaDetcomprob As String
    Mesproceso As String * 2
    Anoproceso As String * 4
    FechaTrabajo As Date
    RutaReport As String
    
    Usuario As String
    Servidor As String
    BDEmpresa As String
    Pwd      As String
    
    ServidorCT As String
    BDEmpresaCT As String
    UsuarioCT As String
    PwdCT     As String
    
    ServidorGEN As String
    BDEmpresaGEN As String
    UsuarioGEN As String
    PwdGEN As String
    
    carpetareportes As String
End Type
Public Declare Function GetComputerName Lib "Kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public VGParametros As Parametrosdecostos
Public VGParamSistem As parametrosdesistema
Public VGtipo As TIPOSISTEMA
'VGtipo = 2
Public Sub main()
 Dim sFileName As String
 Dim sBD As String
 Dim sBDt As String
 Dim n As String
 Dim RSQL As String
 Dim IASA As String
 On Error GoTo Err
 Set VGdllApi = New dll_apisgen.dll_apis
   
 'Verifica si es Copia Ilegal
 '  Verificar_Sistema
    VGComputer = UCase(ComputerName)
    
     VGsql = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "CONEXION", "SQL", "?")
    VGsql = IIf(VGsql = "?", 0, VGsql)
    
    VGformatofecha = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "CONEXION", "FECHASQL", "?")
    VGformatofecha = IIf(VGformatofecha = "?", "DMY", VGformatofecha)

       'Conexion de General
    
    VGParamSistem.BDEmpresaGEN = Trim(VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "BDGENERAL", "BDDATOS", "?"))
    VGParamSistem.ServidorGEN = Trim(VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "BDGENERAL", "SERVIDOR", "?"))
    VGParamSistem.UsuarioGEN = Trim(VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "BDGENERAL", "USUARIO", "?"))
    VGParamSistem.PwdGEN = Trim(VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "BDGENERAL", "PASSW", "?"))
        
    
    'Conexion de costos
    
    VGParamSistem.BDEmpresa = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "conexion", "BDDATOS", "?")
    VGParamSistem.Servidor = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "conexion", "SERVIDOR", "?")
    VGParamSistem.Usuario = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "conexion", "USUARIO", "?")
    VGParamSistem.Pwd = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "conexion", "PASSW", "?")
    VGOrden = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "conexion", "ORDEN", "?")
   
' reportes
   
   VGParamSistem.RutaReport = Trim(VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "REPORTES", "COSTOS", "?"))
   VGParamSistem.carpetareportes = Trim(VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "conexion", "CARPETAREPORTES", "?"))
   
    'Conexion de Contabilidad
    
    VGParamSistem.BDEmpresaCT = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "CONTABILIDAD", "BDDATOS", "?")
    If VGParamSistem.BDEmpresaCT = "" Then
       VGParamSistem.BDEmpresaCT = VGParamSistem.BDEmpresa
       VGParamSistem.ServidorCT = VGParamSistem.Servidor
       VGParamSistem.UsuarioCT = VGParamSistem.Usuario
       VGParamSistem.PwdCT = VGParamSistem.Pwd
     Else
       VGParamSistem.BDEmpresaCT = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "CONTABILIDAD", "BDDATOS", "?")
       VGParamSistem.ServidorCT = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "CONTABILIDAD", "SERVIDOR", "?")
       VGParamSistem.UsuarioCT = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "CONTABILIDAD", "USUARIO", "?")
       VGParamSistem.PwdCT = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "CONTABILIDAD", "PASSW", "?")
   End If
    If VGParamSistem.RutaReport = "" Or VGParamSistem.RutaReport = "?" Then
       VGParamSistem.RutaReport = App.Path
       VGParamSistem.carpetareportes = "Reportes"
    End If
       
    'Establecer Cadena de Conexión de Reportes
    
    vgCADENAREPORT2 = "DSN=jckconsultores;DSQ=" & VGParamSistem.BDEmpresaGEN & ";UID=" & VGParamSistem.UsuarioGEN & ";PWD=" & VGParamSistem.PwdGEN & ""
          

  FrmLogin.Show
  MDIPrincipal.Caption = "Sistema de Costos" & "    Empresa : " & VGParametros.NomEmpresa & "   Base de datos --> " & VGParamSistem.BDEmpresa
  If vgsalir Then
     If VGCNx.State = 1 Then VGCNx.Close
     If VGcnxCT.State = 1 Then VGcnxCT.Close
        MDIPrincipal.Visible = False
        Exit Sub
   Else
        Call CargarParametrosCostos
 End If

 Exit Sub
Err:
    MsgBox Err.Description, vbExclamation, "Aviso"
    Exit Sub
    Resume

End Sub


Public Sub CargarParametrosCostos()
Dim rssql As New ADODB.Recordset
Set rssql = VGCNx.Execute(" Select * from cs_sistema")
If rssql.RecordCount = 0 Then Exit Sub
VGParametros.BaseOrigen = rssql!BaseOrigen
VGParametros.mesesreferencia = rssql!mesesreferencia
VGParametros.codigopersonalplantilla = rssql!codigopersonalplantilla
VGParametros.mesdecierre = ESNULO(rssql!mesdecierre, "000000")
End Sub


