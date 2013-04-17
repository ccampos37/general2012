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
Public VGSalir As Boolean
Public vGCadenaReport As String    'Cadena de Reportes Base Empresa

Public strvalor As String
Public strvalor1 As String

Public VGParametros As ParametrosdeContabilidad
Public VGParamSistem As ParametrosdeSistema
Public VGtipo As TIPOSISTEMA

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
    MultiEmpresas As Boolean
    
    
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
    
    BDempresaCONF As String
    
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
    CarpetaReportes As String
End Type



Public Sub AdjuntarServ(cnx As ADODB.Connection, Servidor As String)
    On Error GoTo ErrAdjunt
        cnx.Execute "Exec sp_addlinkedserver '" & Servidor & "'"
    Exit Sub
ErrAdjunt:
    Exit Sub
End Sub
