Attribute VB_Name = "MainCompras"
'Sistema de Contabilidad Version 1.0
'Visual Basic 6.0 y SQL - 2000
Option Explicit

'Declaraciones de Variables
'Variables Globales

Public Const ColorHabilitado = &H80000004
Public Const ColorDesHabilitado = &H80000005


'Constantes de mensajes para visualizar



Public VgActivalogin As Long                 'Activa login una sola vez
Public Cuentacodigo As String
Public VGactulizodoc As Boolean
Public VGformatofecha As String

Public VGflaglimpia As Boolean               'Flag que Limpia
Public VGMoverRegistro As Boolean            'Flag al mover el registro
Public VGCommandoSP As ADODB.Command         'De Comando
Public VGvardllgen As dllgeneral.dll_general 'Dll de Algunas funciones
Public VGdllApi As dll_apisgen.dll_apis
Public VGvarVerifica As Boolean              'Flag Verifica que transaccion es OK (Grabar ,Etc)
Public VGErrorString As String               'Almacena el Error el que hubo en alguna transaccion
Public VGValorCambio As Double               'Almacena el valor del Cambio
Public VGstrConexion As String               'Cadena de Conexion
Public VgMostrar As Boolean
Public vgcont As Integer

Public VG_gNIVELES() As Integer              'Dígitos por Nivel
Public vGUtil(4) As String                  ' Pase de parametros de ayuda
Public VGOrden As String
Public VGfecha As Date

Private Type ParametrosdeCompras
    monedabase As String * 2
    Igv As Double
    minimoretencion As Double
    NomEmpresa As String
    RucEmpresa As String
    empresacodigo As String * 2
    puntovta As String * 2
    cierremes As Boolean
    
    direccionempresa As String
    impresionalta As Boolean
    ctascompra As String
    Auxaut As Boolean
    
    
    xsubasiento As String
    xLibro As String
    xTipAnal As String
    xCtaIGV As String
    xCtaIES As String
    xCtaRTA As String
    xCtaTotal As String
    xCodPercepcion As String
    xcodretencion As String

    
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
    sistemactaajustedeb As String
    sistemactaajustehab As String
    cuentadeCostos As String
    AsientoAutoxCCostos As String
    numeracionautomaticalibro As Boolean
    
    controlaestadosrendicion  As Boolean
    diasatrazorendicion As Integer
    diacierrerendicion As Integer

        
End Type

Private Type ParametrosdeSistema
    TablaCabcomprob As String
    TablaDetcomprob As String
    BDEmpresa As String
    Mesproceso As String * 2
    Anoproceso As String * 4
    FechaTrabajo As Date
    tipoanaliticocodigo As String
    familiaproyectos As String
    
    Usuario As String
    Servidor As String
    RutaReport As String
    PWD      As String
    
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
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public VGParametros As ParametrosdeCompras
Public VGParamSistem As ParametrosdeSistema
Public VGtipo As TIPOSISTEMA
'VGtipo = 2

Public Sub Main()
    'Base de Datos General
    
On Error GoTo Xmain
 '   VGusuario = "03"
    'Leer Ini
    Set VGdllApi = New dll_apisgen.dll_apis
    VGcomputer = UCase(ComputerName)
     VGsql = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "CONEXION", "SQL", "?")
    VGsql = IIf(VGsql = "?", 0, VGsql)
    
    VGformatofecha = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "CONEXION", "FECHASQL", "?")
    VGformatofecha = IIf(VGformatofecha = "?", "DMY", VGformatofecha)
    
    VGParametros.empresacodigo = "01"
    
    'Conexion de General
    VGParamSistem.BDEmpresaGEN = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "BDGENERAL", "BDDATOS", "?")
    VGParamSistem.ServidorGEN = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "BDGENERAL", "SERVIDOR", "?")
    VGParamSistem.UsuarioGEN = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "BDGENERAL", "USUARIO", "?")
    VGParamSistem.PwdGEN = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "BDGENERAL", "PASSW", "?")
        
    
    'Conexion de Compras
    VGParamSistem.BDEmpresa = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "conexion", "BDDATOS", "?")
    VGParamSistem.Servidor = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "conexion", "SERVIDOR", "?")
    VGParamSistem.Usuario = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "conexion", "USUARIO", "?")
    VGParamSistem.PWD = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "conexion", "PASSW", "?")
    VGOrden = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "conexion", "ORDEN", "?")
   
' reportes
   
   VGParamSistem.RutaReport = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "REPORTES", "COMPRAS", "?")
    VGParamSistem.carpetareportes = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "conexion", "CARPETAREPORTES", "?")
        
    'Conexion de Contabilidad
    VGParamSistem.BDEmpresaCT = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "CONTABILIDAD", "BDDATOS", "?")
    If VGParamSistem.BDEmpresaCT = "" Then
       VGParamSistem.BDEmpresaCT = VGParamSistem.BDEmpresa
       VGParamSistem.ServidorCT = VGParamSistem.Servidor
       VGParamSistem.UsuarioCT = VGParamSistem.Usuario
       VGParamSistem.PwdCT = VGParamSistem.PWD
     Else
       VGParamSistem.BDEmpresaCT = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "CONTABILIDAD", "BDDATOS", "?")
       VGParamSistem.ServidorCT = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "CONTABILIDAD", "SERVIDOR", "?")
       VGParamSistem.UsuarioCT = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "CONTABILIDAD", "USUARIO", "?")
       VGParamSistem.PwdCT = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "CONTABILIDAD", "PASSW", "?")
   End If
    If VGParamSistem.BDEmpresa = "?" Or VGParamSistem.Servidor = "?" Then
        MsgBox "No se ha Configurado bien los parametros BDDATOS y SERVIDOR en el archivo " & Chr(13) & _
               App.Path & "\MARFICE.INI"
    End If
    If VGParamSistem.RutaReport = "" Or VGParamSistem.RutaReport = "?" Then
       VGParamSistem.RutaReport = App.Path
       VGParamSistem.carpetareportes = "Reportes"
    End If
       
    'Establecer Cadena de Conexión de Reportes
    
    VGCadenaReport2 = "DSN=jckconsultores;DSQ=" & VGParamSistem.BDEmpresaGEN & ";UID=" & VGParamSistem.UsuarioGEN & ";PWD=" & VGParamSistem.PwdGEN & ""
        
    'Establecer Conexiones
    Set VGGeneral = New ADODB.Connection
    VGGeneral.CursorLocation = adUseClient
    VGGeneral.CommandTimeout = 0
    VGGeneral.ConnectionTimeout = 0
    VGGeneral.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & VGParamSistem.UsuarioGEN & ";Password=" & VGParamSistem.PwdGEN & ";Initial Catalog=" & VGParamSistem.BDEmpresaGEN & ";Data Source=" & VGParamSistem.Servidor
    VGGeneral.Open
   
   'Call AdjuntarServ(VGGeneral, VGParamSistem.Servidor)
   
   'Conexion de Compras el Principal
    Set VGCNx = New ADODB.Connection
    VGCNx.CursorLocation = adUseClient
    VGCNx.CommandTimeout = 0
    VGCNx.ConnectionTimeout = 0
    VGCNx.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & VGParamSistem.Usuario & ";Password=" & VGParamSistem.PWD & ";Initial Catalog=" & VGParamSistem.BDEmpresa & ";Data Source=" & VGParamSistem.Servidor
    VGCNx.Open
    
   
   'Conexion de Cofiguracion
    Set VGConfig = New ADODB.Connection
    VGConfig.CursorLocation = adUseClient
    VGConfig.CommandTimeout = 0
    VGConfig.ConnectionTimeout = 0
    VGConfig.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & VGParamSistem.Usuario & ";Password=" & VGParamSistem.PWD & ";Initial Catalog=bdwenco;Data Source=" & VGParamSistem.Servidor
    VGConfig.Open
    
    
    MDIPrincipal.menu02_03.Visible = True
          
    'Conexion de Contabilidad
    Set VGcnxCT = New ADODB.Connection
    VGcnxCT.CursorLocation = adUseClient
    VGcnxCT.CommandTimeout = 0
    VGcnxCT.ConnectionTimeout = 0
    VGcnxCT.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & VGParamSistem.UsuarioCT & ";Password=" & VGParamSistem.PwdCT & ";Initial Catalog=" & VGParamSistem.BDEmpresaCT & ";Data Source=" & VGParamSistem.ServidorCT
    VGcnxCT.Open
    Call adicionarcampos
  '   Call CargarParametrosCompras
    Call Parametrogastos
    VgActivalogin = 1
    MDIPrincipal.Show
    Exit Sub
Xmain:
    MsgBox err.Description, vbExclamation
    Exit Sub
    Resume
End Sub

Public Sub AdjuntarServ(cnx As ADODB.Connection, Servidor As String)
    On Error GoTo ErrAdjunt
        cnx.Execute "Exec sp_addlinkedserver '" & Servidor & "'"
    Exit Sub
ErrAdjunt:
    Exit Sub
End Sub

