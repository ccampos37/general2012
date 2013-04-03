Attribute VB_Name = "MainCobrar"
Public Cadenabusca As String

Public cn As New ADODB.Connection

Public VGformatofecha As String
Public VGPlanillaAjuste As Integer

Public VGtipo As TIPOSISTEMA
Public VgActivalogin As Integer

'Variables de Ventana de Ayuda
Public nAyuda As String
Public nDetalle As String
Public nfecha As String

Public g_BaseContab As String
Public g_Empresa As String
Public g_DetalleEmpresa As String
Public g_PedidoPuntoVta As String
Public g_DetallePuntoVta As String

Public VGvardllgen As dllgeneral.dll_general 'Dll de Algunas funciones

Public g_TipoMovi As String
Public g_usuario As String
Public g_ptoventa As String
Public g_pedserie As String
Public g_facserie As String
Public g_bolserie As String
Public g_guiaserie As String

Public punto As PuntoVenta
Public Tipodocu As Tipodocu

'Variables de acceso de usuario
Public Const g_tipoped = "PE"    'PEDIDO
Public Const g_tipofac = "01"    'factura
Public Const g_tipobol = "03"    'boletas
Public Const g_tipoguia = "80"    '9"   'guias B.O
Public Const g_Todos = "99"      'Todos los documentos para Reporte
Public Const g_Eventual = "88888888888" 'cliente eventual

'REPORTES

Public RutaRep  As String
Public RutaRepProc As String

Public VGparametros As Parametrocobrar
Public VGparamsistem As PuntoVenta
Public modoventa As modoventa

Private Type Parametrocobrar         ' Crea Tipo de Empresa
   nombre As String * 2
   moneda As String * 2
   tieneigv As String * 1
   igv As Double
   tienedscto As String * 1
   descuento As Double
   empresacodigo As String * 2
   mensaje As String * 70
   almacen As String * 2
   listapre As String * 1
   tipocambio As Double
   comivende As String * 1
   formaemi As String * 1
   paraboleta As String * 1
   NomEmpresa As String
   RucEmpresa As String
   sistemamultiempresas As Boolean
   contabilizaenlinea As Boolean
   imprimevoucher As Integer
   puntovta As String
   
   sistemactaajustehab As String
   sistemactaajustedeb As String
   
   
End Type


Private Type PuntoVenta   'Crea Punto de Venta

    nropedido As String
    nroguia As String
    nrofactura As String
    nroboleta As String
    nroabono As String
    nrocargo As String
    ventaauto As String
    seriefac As String
    seriebol As String
    serieguia As String
    serieped As String
    serie99 As String
    
    Anoproceso As String
    mesproceso As String
    puntovta As String
    
    bdempresa As String
    Usuario As String
    pwd As String
    servidor As String
    
    bdempresaCONF As String
    
    bdempresaCT As String
    usuarioCT As String
    pwdCT As String
    servidorCT As String
    
    bdempresaGEN As String
    usuarioGEN As String
    pwdGEN As String
    servidorGEN As String
    
    RutaReport As String
    carpetareportes As String
    UsuarioReporte As String
End Type

Private Type modoventa   'Crea modoventa
    descuento As String
    impuestos As String
    nroitem  As Double
    numeraauto As String
    ctrlinventario As String
    unidadmedida As String
    copiasfac As Double
    copiasbol As Double
    copiashoja As Double
    ctacte As String * 1
    ingcliente As String
    ingforma As String
    emitehoja As String
    emiteguia As String
    emitefact As String
    modificaguia As String
    ingpedido As String
    inghoja As String
    ingguia As String
    usafactor As String
    documento As String

End Type


Private Type Tipodocu   'Tipo de documento
    numeauto As String
    numerador As String
End Type


Public Sub Main()
    Set VGdllApi = New dll_apisgen.dll_apis
    VGcomputer = UCase$(ComputerName)
    VGsql = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "CONEXION", "SQL", "?")
    VGsql = IIf(VGsql = "?", 0, VGsql)
    VGparametros.empresacodigo = "01"
   
    VGformatofecha = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "CONEXION", "FECHASQL", "?")
    VGformatofecha = IIf(VGformatofecha = "?", "DMY", VGformatofecha)
    
   
    VGparamsistem.bdempresaGEN = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "BDGENERAL", "BDDATOS", "?")
    VGparamsistem.servidorGEN = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "BDGENERAL", "SERVIDOR", "?")
    VGparamsistem.usuarioGEN = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "BDGENERAL", "USUARIO", "?")
    VGparamsistem.pwdGEN = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "BDGENERAL", "PASSW", "?")
    
    VGparamsistem.bdempresa = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "CONEXION", "BDDATOS", "?")
    VGparamsistem.servidor = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "CONEXION", "SERVIDOR", "?")
    VGparamsistem.Usuario = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "CONEXION", "USUARIO", "?")
    VGparamsistem.pwd = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "CONEXION", "PASSW", "?")
    
    VGparamsistem.bdempresaCONF = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "CONEXION", "BDDATOSCONF", "?")
    If VGparamsistem.bdempresaCONF = "?" Then VGparamsistem.bdempresaCONF = "Bdwenco"
    
    
    VGparamsistem.bdempresaCT = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "CONTABILIDAD", "BDDATOS", "?")
    VGparamsistem.servidorCT = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "CONTABILIDAD", "SERVIDOR", "?")
    VGparamsistem.usuarioCT = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "CONTABILIDAD", "USUARIO", "?")
    VGparamsistem.pwdCT = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "CONTABILIDAD", "PASSW", "?")
    
    VGCadenaReport2 = "DSN=jckconsultores;DSQ=" & VGparamsistem.bdempresaGEN & ";UID=" & VGparamsistem.usuarioGEN & ";PWD=" & VGparamsistem.pwdGEN & ""
    
    'Establecer Conexiones
    Set VGGeneral = New ADODB.Connection
    VGGeneral.CursorLocation = adUseClient
    VGGeneral.CommandTimeout = 0
    VGGeneral.ConnectionTimeout = 0
    VGGeneral.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & VGparamsistem.usuarioGEN & ";Password=" & VGparamsistem.pwdGEN & ";Initial Catalog=" & VGparamsistem.bdempresaGEN & ";Data Source=" & VGparamsistem.servidorGEN
    VGGeneral.Open
    
    VGparamsistem.UsuarioReporte = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", VGcomputer, "USUARIO", "?")
    
    If VGparamsistem.bdempresa = "?" Or VGparamsistem.servidor = "?" Then
        MsgBox "No se ha Configurado bien los parametros BDDATOS y SERVIDOR en el archivo " & Chr(13) & _
               App.Path & "\MARFICE.INI"
    End If
    
    VGparamsistem.RutaReport = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "REPORTES", "CONTABILIDAD", "?")
    VGparamsistem.carpetareportes = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "conexion", "CARPETAREPORTES", "?")
  
    
    'Establecer Cadena de Conexión de Reportes
    vgCADENAREPORT = "DSN=" & VGparamsistem.servidor & ";DSQ=" & VGparamsistem.bdempresa & ";UID=" & VGparamsistem.UsuarioReporte & ";PWD=" & VGparamsistem.pwd
    
    Set VGCNx = New ADODB.Connection
    VGCNx.CursorLocation = adUseClient
    VGCNx.CommandTimeout = 0
    VGCNx.ConnectionTimeout = 0
    VGCNx.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & VGparamsistem.Usuario & ";Password=" & VGparamsistem.pwd & ";Initial Catalog=" & VGparamsistem.bdempresa & ";Data Source=" & VGparamsistem.servidor
    VGCNx.Open
    
    Set VGcnxCT = New ADODB.Connection
    VGcnxCT.CursorLocation = adUseClient
    VGcnxCT.CommandTimeout = 0
    VGcnxCT.ConnectionTimeout = 0
    VGcnxCT.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & VGparamsistem.usuarioCT & ";Password=" & VGparamsistem.pwdCT & ";Initial Catalog=" & VGparamsistem.bdempresaCT & ";Data Source=" & VGparamsistem.servidorCT
'    VGCnxCT.ConnectionString = VGcnx.ConnectionString
    VGcnxCT.Open
    
 
   'Conexion de Cofiguracion

    Set VGConfig = New ADODB.Connection
    VGConfig.CursorLocation = adUseClient
    VGConfig.CommandTimeout = 0
    VGConfig.ConnectionTimeout = 0
    VGConfig.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & VGparamsistem.Usuario & ";Password=" & VGparamsistem.pwd & ";Initial Catalog=" & VGparamsistem.bdempresaCONF & ";Data Source=" & VGparamsistem.servidor
    VGConfig.Open

    VGtipo = cobrar
   'Call AdjuntarServ(VGGeneral, VGParamSistem.Servidor)
    'Base de datos de la Empresa
    Call adicionarcampos
 

    VgActivalogin = 1
    MDIPrincipal.Show
    Exit Sub
Xmain:
    MsgBox Err.Description, vbExclamation

   Call ParametrosFuncionalesCobrar
   Call Cargar_Parametros_Funcionales
   FrmIngreso.Show
   'MDIPrincipal.Show
End Sub

Public Sub ADOConectar()
On Error GoTo error

VGGeneral.CursorLocation = adUseClient
VGGeneral.CommandTimeout = 0
VGGeneral.ConnectionTimeout = 200
VGGeneral.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & VGparamsistem.usuarioGEN & ";Password=" & VGparamsistem.pwdGEN & ";Initial Catalog=" & VGparamsistem.bdempresaGEN & ";Data Source=" & VGparamsistem.servidorGEN
VGGeneral.Open

   
'Conexion de Cofiguracion

Set VGConfig = New ADODB.Connection
VGConfig.CursorLocation = adUseClient
VGConfig.CommandTimeout = 0
VGConfig.ConnectionTimeout = 0
VGConfig.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & VGparamsistem.Usuario & ";Password=" & VGparamsistem.pwd & ";Initial Catalog=bdwenco;Data Source=" & VGparamsistem.servidor
VGConfig.Open
    
'Conexion de inventarios

If VGparamsistem.bdempresa = "" Or VGparamsistem.bdempresa = "?" Then
   Set RSQL = VGConfig.Execute("select empresabaseinventarios from empresa where empresaflaginventarios=1")
   VGparamsistem.bdempresa = RSQL!empresabaseinventarios
End If
Set VGCNx = New ADODB.Connection
VGCNx.CursorLocation = adUseClient
VGCNx.CommandTimeout = 0
VGCNx.ConnectionTimeout = 0
VGCNx.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & VGparamsistem.Usuario & ";Password=" & VGparamsistem.pwd & ";Initial Catalog=" & VGparamsistem.bdempresa & ";Data Source=" & VGparamsistem.servidor
VGCNx.Open
    
'Conexion de Contabilidad

VGcnxCT.CursorLocation = adUseClient
VGcnxCT.CommandTimeout = 0
VGcnxCT.ConnectionTimeout = 0
VGcnxCT.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & VGparamsistem.usuarioCT & ";Password=" & VGparamsistem.pwdCT & ";Initial Catalog=" & VGparamsistem.bdempresaCT & ";Data Source=" & VGparamsistem.servidorCT
VGcnxCT.Open
    
'Call adicionacamposct
Exit Sub

error:
    
MsgBox Err.Description, vbExclamation
Exit Sub
Resume
End Sub

