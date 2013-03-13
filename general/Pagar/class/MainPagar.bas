Attribute VB_Name = "MainPagar"
Option Explicit
Public Cadenabusca As String


Public VGvardllgen As dllgeneral.dll_general 'Dll de Algunas funciones
Public VGdllApi As dll_apisgen.dll_apis

'Variables de Ventana de Ayuda
Public nAyuda As String
Public nDetalle As String
Public nSaldo As Double   'Adicional para el Saldo Pendiente
Public VG_String As String


Public VGaplicaciones As Double


Public g_Empresa As String
Public g_DetalleEmpresa As String
Public g_PedidoPuntoVta As String
Public g_DetallePuntoVta As String

Public g_TipoMovi As String
Public g_ptoventa As String
Public g_pedserie As String
Public g_facserie As String
Public g_bolserie As String
Public g_guiaserie As String
Public VGformatofecha As String

Public punto As PuntoVenta

Public modoventa As modoventa
Public Tipodocu As Tipodocu

'Variables de acceso de usuario
Public Const g_tipoped = "PE"    'PEDIDO
Public Const g_tipofac = "01"    'factura
Public Const g_tipobol = "03"    'boletas
Public Const g_tipoguia = "80"    '9"   'guias B.O
Public Const g_Todos = "99"      'Todos los documentos para Reporte
Public Const g_Eventual = "88888888888" 'Proveedor eventual

'Constantes de mensajes para visualizar

Public Const MsgElim = "No Existen Datos a Eliminar.."
Public Const MsgAdd = "Los datos ya existen...Verifique!!!"
Public Const Msg29 = "Debe Ingresar Numeros"

Public struser As String
Public strserver As String


Public VGparametros As ParametroPagar

Private Type ParametroPagar         ' Parametros de Empresa
   NomEmpresa As String
   RucEmpresa As String
   empresacodigo As String
   puntovta As String
   nombre As String
   moneda As String
   tieneigv As String
   tienedscto As String
   descuento As Double
   igv As Double
   mensaje As String
   almacen As String
   listapre As String
   tipocambio As Double
   comivende As String
   formaemi As String
   paraboleta As String
   sistemamultiempresas As Boolean
   contabilizaenlinea As Boolean
   sistemactaajustehab As String
   sistemactaajustedeb As String
End Type

Public VGparamsistem As PuntoVenta

Private Type PuntoVenta   'Crea Punto de Venta
    
    FechaTrabajo As Date
    Mesproceso As String
    Anoproceso As String
    
    RutaReport As String
    RutaCarpeta As String
    carpetareportes As String
    
    BDEmpresa As String
    Usuario As String
    Servidor As String
    PWD      As String
    
    ServidorGEN As String
    BDEmpresaGEN As String
    UsuarioGEN As String
    PwdGEN As String
    
    ServidorCT As String
    BDEmpresaCT As String
    UsuarioCT As String
    PwdCT As String
    
    puntovta As String
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
    ctacte As String
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
    aplicacion As String
End Type

Public VGtipo As TIPOSISTEMA

Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

