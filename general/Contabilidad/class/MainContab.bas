Attribute VB_Name = "MainContab"
'Sistema de Contabilidad Version 1.0
'Visual Basic 6.0 y SQL - 2000
'Desarrolladores :
'Fernando Cossio Peralta
'Ivan Crispin Sanchez

Option Explicit
'Declaraciones de Variables
'Variables Globales
Public Const ColorHabilitado = &H80000004
Public Const ColorDesHabilitado = &H80000005

Public VgActivalogin As Long                 'Activa login una sola vez


Public VGflaglimpia As Boolean               'Flag que Limpia
Public VGMoverRegistro As Boolean            'Flag al mover el registro
Public VGCommandoSP As ADODB.Command         'De Comando
Public VGvardllgen As dllgeneral.dll_general 'Dll de Algunas funciones
Public VGdllApi As dll_apisgen.dll_apis
Public VGvarVerifica As Boolean              'Flag Verifica que transaccion es OK (Grabar ,Etc)
Public VGErrorString As String               'Almacena el Error el que hubo en alguna transaccion
Public VGValorCambio As Double               'Almacena el valor del Cambio
Public VGRepiteDoc As Boolean                'Flag de Repite de documento en subasiento
Public VGactulizodoc As Boolean              'Flag de Actualizacion del detalle de comprobante
Public VGstrConexion As String               'Cadena de Conexion
Public VgMostrar As Boolean
Public vgcont As Integer
Public Vgdocumentoanulado As String
Public VGformatofecha As String



Public VGnumnivelescuenta As Integer               'Número de Niveles del Plan de Cuentas
Public VG_aNIVELES() As Integer              'Dígitos por Nivel de cuenta
Public VGnumnivelescentrocosto As Integer               'Número de Niveles de los centros de costos
Public VG_cNIVELES() As Integer              'Dígitos por Nivel de costos
Public VGGlosa As String                     'Glosa del Sub Asiento
Public VGMonSubAsiento As String             'Moneda por defecto del Sub Asiento


Public vgCADENAREPORT As String              'Cadena de Reportes Base Empresa

Public strvalor As String
Public strvalor1 As String

Private Type ParametrosdeContabilidad
    monedabase As String * 2
    IGV As Double
    empresacodigo As String * 2
    NomEmpresa As String
    RucEmpresa As String
    CuadreAsiento As Boolean
    ImpresionAsiento As Boolean
    impresionalta As Boolean
    asientocodigo As String * 3
    subasientocodigo As String * 4
    documentoanulado As String * 1
    sistemamonista As Boolean
    sistemactaajustedeb As String
    sistemactaajustehab As String
    
    AsientoAutoxCCostos As String
    cuentadeCostos As String
    puntovta As String * 2
    cierremes As Boolean
    multiempresas As Boolean
    
    
End Type
Private Type ParametrosdeSistema
    TablaCabcomprob As String
    TablaDetcomprob As String
    BDEmpresa As String
    Mesproceso As String * 2
    Anoproceso As String * 4
    FechaTrabajo As Date
    Usuario As String
    Servidor As String
    
    ServidorGEN As String
    UsuarioGEN As String
    PwdGEN As String
    BDEmpresaGEN As String
    
    ServidorCT As String
    UsuarioCT As String
    PwdCT As String
    BDEmpresaCT As String
   
   RutaReport As String
    Pwd      As String
    UsuarioReporte As String
    carpetareportes As String
End Type
Public VGParametros As ParametrosdeContabilidad
Public VGParamSistem As ParametrosdeSistema
Public VGtipo As TIPOSISTEMA

Public Sub Main()
    'Base de Datos General
      
On Error GoTo Xmain
    'VGusuario = "03"
    'Leer Ini
    Set VGdllApi = New dll_apisgen.dll_apis
    VGcomputer = UCase$(ComputerName)
    VGsql = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "CONEXION", "SQL", "?")
    VGsql = IIf(VGsql = "?", 0, VGsql)
   VGParametros.empresacodigo = "01"
   
    VGformatofecha = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "CONEXION", "FECHASQL", "?")
    VGformatofecha = IIf(VGformatofecha = "?", "DMY", VGformatofecha)
    
   
    VGParamSistem.BDEmpresaGEN = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "BDGENERAL", "BDDATOS", "?")
    VGParamSistem.ServidorGEN = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "BDGENERAL", "SERVIDOR", "?")
    VGParamSistem.UsuarioGEN = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "BDGENERAL", "USUARIO", "?")
    VGParamSistem.PwdGEN = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "BDGENERAL", "PASSW", "?")
    
    VGCadenaReport2 = "DSN=jckconsultores;DSQ=" & VGParamSistem.BDEmpresaGEN & ";UID=" & VGParamSistem.UsuarioGEN & ";PWD=" & VGParamSistem.PwdGEN & ""
    
    'Establecer Conexiones
    Set VGGeneral = New ADODB.Connection
    VGGeneral.CursorLocation = adUseClient
    VGGeneral.CommandTimeout = 0
    VGGeneral.ConnectionTimeout = 0
    VGGeneral.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & VGParamSistem.UsuarioGEN & ";Password=" & VGParamSistem.PwdGEN & ";Initial Catalog=" & VGParamSistem.BDEmpresaGEN & ";Data Source=" & VGParamSistem.ServidorGEN
    VGGeneral.Open
    
    If Trim$(VGdllApi.LeerIni(App.Path & "\MARFICE.INI", VGcomputer, "conexion", "?")) <> "?" Then
        VGParamSistem.BDEmpresa = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", VGcomputer, "BDDATOS", "?")
        VGParamSistem.Servidor = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", VGcomputer, "SERVIDOR", "?")
        VGParamSistem.Usuario = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", VGcomputer, "USUARIO", "?")
        VGParamSistem.UsuarioReporte = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", VGcomputer, "USUARIO", "?")
        VGParamSistem.Pwd = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", VGcomputer, "PASSW", "?")
      Else
        VGParamSistem.BDEmpresa = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "CONEXION", "BDDATOS", "?")
        VGParamSistem.Servidor = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "CONEXION", "SERVIDOR", "?")
        VGParamSistem.Usuario = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "CONEXION", "USUARIO", "?")
        VGParamSistem.UsuarioReporte = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "CONEXION", "USUARIO", "?")
        VGParamSistem.Pwd = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "CONEXION", "PASSW", "?")
    End If
    If VGParamSistem.BDEmpresa = "?" Or VGParamSistem.Servidor = "?" Then
        MsgBox "No se ha Configurado bien los parametros BDDATOS y SERVIDOR en el archivo " & Chr(13) & _
               App.Path & "\MARFICE.INI"
    End If
    
    VGParamSistem.RutaReport = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "REPORTES", "CONTABILIDAD", "?")
    VGParamSistem.carpetareportes = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "conexion", "CARPETAREPORTES", "?")
  
    
    'Establecer Cadena de Conexión de Reportes
    vgCADENAREPORT = "DSN=" & VGParamSistem.Servidor & ";DSQ=" & VGParamSistem.BDEmpresa & ";UID=" & VGParamSistem.UsuarioReporte & ";PWD=" & VGParamSistem.Pwd
    Set VGCNx = New ADODB.Connection
    VGCNx.CursorLocation = adUseClient
    VGCNx.CommandTimeout = 0
    VGCNx.ConnectionTimeout = 0
    VGCNx.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & VGParamSistem.Usuario & ";Password=" & VGParamSistem.Pwd & ";Initial Catalog=" & VGParamSistem.BDEmpresa & ";Data Source=" & VGParamSistem.Servidor
    VGCNx.Open
    
    Set VGcnxCT = New ADODB.Connection
    VGcnxCT.CursorLocation = adUseClient
    VGcnxCT.CommandTimeout = 0
    VGcnxCT.ConnectionTimeout = 0
    VGcnxCT.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & VGParamSistem.Usuario & ";Password=" & VGParamSistem.Pwd & ";Initial Catalog=" & VGParamSistem.BDEmpresa & ";Data Source=" & VGParamSistem.Servidor
'    VGCnxCT.ConnectionString = VGcnx.ConnectionString
    VGcnxCT.Open
    
 
   'Conexion de Cofiguracion

    Set VGConfig = New ADODB.Connection
    VGConfig.CursorLocation = adUseClient
    VGConfig.CommandTimeout = 0
    VGConfig.ConnectionTimeout = 0
    VGConfig.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & VGParamSistem.Usuario & ";Password=" & VGParamSistem.Pwd & ";Initial Catalog=bdwenco;Data Source=" & VGParamSistem.Servidor
    VGConfig.Open

    VGtipo = contab
   'Call AdjuntarServ(VGGeneral, VGParamSistem.Servidor)
    'Base de datos de la Empresa
    Call adicionarcampos
 

    VgActivalogin = 1
    MDIPrincipal.Show
    Exit Sub
Xmain:
    MsgBox err.Description, vbExclamation
End Sub

Public Sub AdjuntarServ(cnx As ADODB.Connection, Servidor As String)
    On Error GoTo ErrAdjunt
        cnx.Execute "Exec sp_addlinkedserver '" & Servidor & "'"
    Exit Sub
ErrAdjunt:
    Exit Sub
End Sub
