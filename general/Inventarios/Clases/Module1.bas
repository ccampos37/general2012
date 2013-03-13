Attribute VB_Name = "inicio"
Option Explicit

'*************************
Public g_DetalleEmpresa As String
Public g_usuario As String
Public g_PedidoPuntoVta As String
Public g_DetallePuntoVta As String
Public g_tipofac As String
Public g_tipoped As String
'*************************

Public nAyuda1   As String
Public nMoneda As String

Public VGSeleccion As Integer    ' modificar o adicionar. descripcion en el formulario registro

Public VGmodovta As String
Public VGCodigo As String
Public VGAlma As String
         'Codigo de la compannia
Public vgtipo As TIPOSISTEMA
Public VGOrden As String

Public VGformatofecha As String
Public VGflagconversioncodigo As Boolean
Public VGtransf As Integer

'***
Public VGServer As String
Public VGBUsuario As String
Public VGPassw As String

Public VGServer2 As String
Public VGBase2 As String
Public VGBUsuario2 As String
Public VGPassw2 As String

Public VGDIRE As String
Public VGBase3 As String

Public VGbase4 As String
Public VGxx As ADODB.Connection
Public VGCommandoSP As ADODB.Command         'De Comando


Public VGdllApi As dll_apisgen.dll_apis
Public VGvardllgen As dllgeneral.dll_general 'Dll de Algunas funciones

Public GPunto As String   'punto de venta

'*********

Public VGRegEnt  As Integer         ' registro de entrada o salida
Public VGSoles As Boolean           'indica si se trabaja con S/. o $
Public VGtransp As Boolean          'indica el trasnportista ,true para llamar del manteniminento,false  de g. remision
Public VGForm   As Integer            'indica el formulario en uso
Public VGForm1   As Integer          ' indica el formulario de procedencia para la ayuda
Public VGtipocreacion As Integer   'Para el modificar
Public VGabrev As String              ' codigo de unidad
Public VGcrea As Boolean
           'Codigo del almacen
Public VGval As Boolean               'Indicca si es valorizado
Public VGtipoAprobacion As Integer

Public VGActualizar As Boolean    'Para el caso de modificar  y restaurar informacion  en caso de no modificar
Public VGElimina As Boolean          'Para el caso de utilizar el formulario de eliminar y anular
Public VGAyuClie As Boolean
Public VGGuiaSal As Boolean        'Para el caso del form de Guia de salida en que puede crear o  modificar parcial
Public VGRuta As String                 'ruta de la base de datos
Public VGTipCamb As Double        'Tipo de Cambios
Public VGCodMon As String       'Tipo de Moneda
'Public VGWrk As Workspace
'Public VGBaseDatos As Database
Public VGSALIR As Boolean
Public VGEstadomodi As Boolean    'Estado de modificacion

Public VGValnuevo  As Boolean           'Para doc valorizados
Public VGUsua  As String

Public VGRclie As Boolean
Public cAnexo As String
Public VGIASA As String                     'Codigo de la empresa IASA, aplicacion personalizada
Public vGAdmLog As Boolean             'Login del Administrador
Public VGNameCont  As String             'Nombre de contabilidad
Public VGContTra As String                 'NombredBD de trasacciones de Contabilidad
Public VGAutomatico As Boolean         'Indica si la numeracion no es editable
Public VGcc As Integer     'Indica si el reporte espor centro de costo o autorizado
Public stockcomp As String     ' controla stock comprometido
Public VGOrdenes As Integer    ' controla requerimientos de ordenes
  


 Public cRuta6 As String
 Public cNomBd  As String
 Public cNomBd5 As String
 Public Const cNomBd6 As String = "BdWenco.Mdb"
 Public cNomBd4 As String
 Public cNomBd2 As String
 Public cRuta5 As String
 Public cRuta2 As String                    'Contiene la ruta de Bd, incluyend el nombre  ***********
 Public cRuta3 As String
 Public cRuta4 As String
 Public sName As String

Public VGNomAlm As String              ' Nombre del almacen
Public RUTA As String                        'Indica solo la carpeta donde se ha instalado


'Variables globales para Administradores
Public VGTEMP As String
Public VGADM_CODIDO As String
Public VGADM_NOMBRE As String
Public VGADM_PASSWORD As String
Public VGAdmLogin As Boolean
'Variables globales para Empresas

Public VGusuariocodigo As String
Public VGUSU_PASSWORD As String
Public VG_FecTrab As Date
Public VGcod As String                          'Se utiliza para las consultas
Public vGUtil(4) As String                        'Se para los pases de ayuda
Public arrayserie()   As String                'Ingreso masivo de serie

'RMM**************************************
'Public ClsTock As New ClsTock
Public ClsTDoc As New ClasDocumento
Public ClsTock As New ClasMovimientos
Public VGLadrillera As Boolean
Dim cConexAux As New ADODB.Connection

'Variables de Ventana de Ayuda
Public nAyuda As String
Public nDetalle As String
Public ndensidad As Long
Public modoventa As modoventa
Public VGParametros As Parametrosdeempresa

Public Type Parametrosdeempresa
 
    monedabase As String
    Igv As Double
    direccionempresa As String
    impresionalta As Boolean
    ctascompra As String
    tipodevalorizacion As String
    SaldosvalorxAlmacen As Integer
    Valorestadooccodigo As Integer   '' valor de estado de requerimiento
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
    TipoValorizacion As Integer  ' 1: empresa 2' x establecimiento
    SaldoConsolidadoxPedidos As Integer
     
 
 VGLongCodigo As Integer    'Inddica la long de codigo de un articulo
 NomEmpresa As String
 RucEmpresa As String
 puntovta As String
 empresacodigo As String
 mesproceso As String
 VGporcentajeimpto As Double
 MontoexeneradoLiqCompra As Double
 nombreguia As String
 nombrefactura As String
 minimoretencion As Double
 auxaut As Boolean
 sistemaactivaccostos As Boolean
 sistemaasientoenlinea As Boolean
 sistemactrlgastos As Boolean
 sistemamultiempresas As Boolean
 sistemabancarizacion As Boolean
 sistemabancarizacion01 As Double
 sistemabancarizacion02 As Double

 PermiteRequerimientos As Boolean
 PermiteIngresosconRequerimientos As Boolean
 
 tipocreacioncodigo As String
 tipogeneracioncodigo As Double
 
End Type




Public Type modoventa   'Crea modoventa
    descuento As String
    impuestos As String
    nroitem  As Double
    numeraauto As String
    ctrlinventario As String
    unidadmedida As String
    copiasfac As Double
    copiasbol As Double
    copiashoja As Double
    copiasguiarem As Double
    ctacte As String
    ingcliente As String
    ingforma As String
    emiteguia As String
    emitefact As String
    modificaguia As String
    ingpedido As String
    inghoja As String
    ingguia As String
    usafactor As String
    documento As String
    almacenes As String


    valorizaliqcompra As String

End Type

Public VGParamSistem As ParametrosdeSistema

Public Type ParametrosdeSistema
    TablaCabcomprob As String
    tabladetcomprob As String
    mesproceso As String
    AnoProceso As String
    fechatrabajo As Date
    RutaReport As String
    Servidor As String
    BDEmpresa As String
    Usuario As String
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
    
    tipoanaliticocodigo As String
End Type




'/-------------------------------------
'/ Modificado por Carlos Enrique
'/                2005/09/24
'/ Decia: Public nsaldo As Integer
'/ Debe decir: Public nsaldo As long
'/-------------------------------------
Public nSaldo As Long

Public g_ptoventa As String

