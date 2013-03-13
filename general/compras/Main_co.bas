Attribute VB_Name = "Modulo"
'Sistema de Contabilidad Version 1.0
'Visual Basic 6.0 y SQL - 2000


Option Explicit

'Declaraciones de Variables
'Variables Globales

Public Const ColorHabilitado = &H80000004
Public Const ColorDesHabilitado = &H80000005

'Constantes de mensajes para visualizar

Public Const MsgEdit = "No Existen Datos para Editar.. "
Public Const MsgGraba = "Datos Grabados satisfactoriamente...."
Public Const MsgElim = "No Existen Datos a Eliminar.."
Public Const MsgAdd = "Los datos ya existen...Verifique!!!"
Public Const MsgTitle = "AVISO"
Public Const Msg29 = "Debe Ingresar Numeros"

Public VGcnx As ADODB.Connection             'Conexion de la BD empresa
Public VGcnxCT As ADODB.Connection        'Conexion de Contabilidad
Public VGcnxCP As ADODB.Connection          'Conexion de Cuentas x Pagar

Public VgActivalogin As Long                 'Activa login una sola vez
Public VGcnxMarfice As ADODB.Connection      'Conexion de la BD Generales
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

Public VG_gNIVELES() As Integer              'Dígitos por Nivel
Public VGnumniveles As Integer               'Número de Niveles del Plan de Cuentas
Public VGnumnivgas As Integer               'Número de Niveles del Plan de gastos
Public vGUtil(4) As String                  ' Pase de parametros de ayuda

Public VGusuario As String
Public VGComputer As String                  'Nombre de la computadora
Public VGOrden As String


Public vgCADENAREPORT2 As String              'Cadena de Reportes Base Marfice


Private Type ParametrosdeCompras
    monedabase As String * 2
    Igv As Double
    NomEmpresa As String
    RucEmpresa As String
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
    
    CpTiplan As String
    CpOficina As String
    permite_tc As Boolean
    sistemaactivaccostos As Boolean
    sistemaasientoenlinea As Boolean
    sistemactrlgastos As Boolean
        
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
    RutaReport As String
    PWD      As String
    
    ServidorCT As String
    BDEmpresaCT As String
    UsuarioCT As String
    PwdCT     As String
    
    ServidorCP As String
    BDEmpresaCP As String
    UsuarioCP As String
    PwdCP     As String
    
    ServidorGEN As String
    BDEmpresaGEN As String
    UsuarioGEN As String
    PwdGEN As String
    
    carpetareportes As String
End Type
Public Cuentacodigo As String
Public VGParamCompra As ParametrosdeCompras
Public VGParamSistem As ParametrosdeSistema
Public Sub Main()
    'Base de Datos General
    
On Error GoTo Xmain
    VGusuario = "03"
    'Leer Ini
    Set VGdllApi = New dll_apisgen.dll_apis
    VGComputer = UCase(ComputerName)
    'Conexion de General
    VGParamSistem.BDEmpresaGEN = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "BDGENERAL", "BDDATOS", "?")
    VGParamSistem.ServidorGEN = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "BDGENERAL", "SERVIDOR", "?")
    VGParamSistem.UsuarioGEN = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "BDGENERAL", "USUARIO", "?")
    VGParamSistem.PwdGEN = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "BDGENERAL", "PASSW", "?")
        
    
    'Conexion de Compras
    VGParamSistem.BDEmpresa = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "COMPRAS", "BDDATOS", "?")
    VGParamSistem.Servidor = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "COMPRAS", "SERVIDOR", "?")
    VGParamSistem.Usuario = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "COMPRAS", "USUARIO", "?")
    VGParamSistem.PWD = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "COMPRAS", "PASSW", "?")
    VGOrden = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "COMPRAS", "ORDEN", "?")
    VGParamSistem.RutaReport = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "COMPRAS", "RUTAREPORTES", "?")
            
        
    'Conexion de Contabilidad
    VGParamSistem.BDEmpresaCT = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "CONTABILIDAD", "BDDATOS", "?")
    VGParamSistem.ServidorCT = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "CONTABILIDAD", "SERVIDOR", "?")
    VGParamSistem.UsuarioCT = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "CONTABILIDAD", "USUARIO", "?")
    VGParamSistem.PwdCT = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "CONTABILIDAD", "PASSW", "?")
    
    'Conexion de Cuentas x Pagar
    VGParamSistem.BDEmpresaCP = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "CXP", "BDDATOS", "?")
    VGParamSistem.ServidorCP = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "CXP", "SERVIDOR", "?")
    VGParamSistem.UsuarioCP = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "CXP", "USUARIO", "?")
    VGParamSistem.PwdCP = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "CXP", "PASSW", "?")
    
    
    
    If VGParamSistem.BDEmpresa = "?" Or VGParamSistem.Servidor = "?" Then
        MsgBox "No se ha Configurado bien los parametros BDDATOS y SERVIDOR en el archivo " & Chr(13) & _
               App.Path & "\MARFICE.INI"
    End If
    If VGParamSistem.RutaReport = "" Or VGParamSistem.RutaReport = "?" Then
       VGParamSistem.RutaReport = App.Path
    End If
       
    'Establecer Cadena de Conexión de Reportes
    
    vgCADENAREPORT2 = "DSN=" & VGParamSistem.Servidor & ";DSQ=" & VGParamSistem.BDEmpresaGEN & ";UID=" & VGParamSistem.Usuario & ";PWD=" & ""
        
    'Establecer Conexiones
    Set VGcnxMarfice = New ADODB.Connection
    VGcnxMarfice.CursorLocation = adUseClient
    VGcnxMarfice.CommandTimeout = 0
    VGcnxMarfice.ConnectionTimeout = 0
    VGcnxMarfice.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & VGParamSistem.UsuarioGEN & ";Password=" & VGParamSistem.PwdGEN & ";Initial Catalog=" & VGParamSistem.BDEmpresaGEN & ";Data Source=" & VGParamSistem.Servidor
    VGcnxMarfice.Open
   'Call AdjuntarServ(VGcnxMarfice, VGParamSistem.Servidor)
   
   'Conexion de Compras el Principal
    Set VGcnx = New ADODB.Connection
    VGcnx.CursorLocation = adUseClient
    VGcnx.CommandTimeout = 0
    VGcnx.ConnectionTimeout = 0
    VGcnx.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & VGParamSistem.Usuario & ";Password=" & VGParamSistem.PWD & ";Initial Catalog=" & VGParamSistem.BDEmpresa & ";Data Source=" & VGParamSistem.Servidor
    VGcnx.Open
    MDIPrincipal.menu02_03.Visible = True
          
    'Conexion de Contabilidad
    Set VGcnxCT = New ADODB.Connection
    VGcnxCT.CursorLocation = adUseClient
    VGcnxCT.CommandTimeout = 0
    VGcnxCT.ConnectionTimeout = 0
    VGcnxCT.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & VGParamSistem.UsuarioCT & ";Password=" & VGParamSistem.PwdCT & ";Initial Catalog=" & VGParamSistem.BDEmpresaCT & ";Data Source=" & VGParamSistem.ServidorCT
    VGcnxCT.Open
    
    'Conexion de Cuentas por Pagar
    Set VGcnxCP = New ADODB.Connection
    VGcnxCP.CursorLocation = adUseClient
    VGcnxCP.CommandTimeout = 0
    VGcnxCP.ConnectionTimeout = 0
    VGcnxCP.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & VGParamSistem.UsuarioCP & ";Password=" & VGParamSistem.PwdCP & ";Initial Catalog=" & VGParamSistem.BDEmpresaCP & ";Data Source=" & VGParamSistem.ServidorCP
    VGcnxCP.Open
    
    Call CargarParametrosCompras
    Call Parametrogastos
    VgActivalogin = 1
    MDIPrincipal.Show
    Exit Sub
Xmain:
    MsgBox Err.Description, vbExclamation
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
