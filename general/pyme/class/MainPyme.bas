Attribute VB_Name = "MainPyme"
Option Explicit

'*************************
Public g_DetalleEmpresa As String
Public g_usuario As String
Public g_PedidoPuntoVta As String
Public g_DetallePuntoVta As String
Public g_TipoMovi As Integer

 Public g_pedserie As String
 Public g_facserie As String
 Public g_bolserie As String
 Public g_guiaserie As String
Public g_ticserieb As String
Public g_ticserief As String
Public g_GuiaRemSerie As String



Public Const g_tipoped = "PE"    'PEDIDO
Public Const g_tipocot = "CO"    'PEDIDO
Public Const g_tipoticketB = "14"    'Tickets boletas
Public Const g_tipoticketf = "15"    'facturas
Public Const g_tipofac = "01"    'factura
Public Const g_tipobol = "03"    'boletas
Public Const g_tipoguia = "80"    '9"   'guias B.O
Public Const g_Todos = "99"      'Todos los documentos para Reporte
Public Const g_Eventual = "88888888888" 'cliente eventual
Public Const g_tipoguiaRem = "GR" ' guias de remision
'*************************

Public nAyuda1   As String
Public nMoneda As String
Public VgModificar As Integer
Public VGSeleccion As Integer    ' modificar o adicionar. descripcion en el formulario registro

Public VGmodovta As String
Public VGCodigo As String
Public VGAlma As String
Public VGoficina As String
         'Codigo de la compannia

Public VGOrden As String

Public VGformatofecha As String
Public VGflagconversioncodigo As Boolean
Public VGtransf As Integer

'***
Public VGServer As String

Public VGPassw As String

Public VGServer2 As String
Public VGBase2 As String
Public VGBUsuario2 As String
Public VGPassw2 As String

Public VGDIRE As String
Public VGBase3 As String

Public VGbase4 As String
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
            'el usuario de la aplicacion
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

Public NombreImpresora As String

'Variables globales para Administradores
Public VGTEMP As String
Public VGADM_CODIDO As String
Public VGADM_NOMBRE As String
Public VGADM_PASSWORD As String
Public VGAdmLogin As Boolean
'Variables globales para Empresas

Public VGusuariocodigo As String
Public VGUsuarioPassword As String
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

Public vgtipo As TIPOSISTEMA
Public VGParametros As Parametrosdeempresa
Public Type Parametrosdeempresa
    
    empresacodigoretencion As String
    porcentajeretencion As Double
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
    cierremes As Boolean
    nropedido As String
    
    multifacturas As Boolean
    empresaasientosautomaticos As String
    CpTiplan As String
    CpOficina As String
    permite_tc As Boolean
    TipoValorizacion As Integer  ' 1: empresa 2' x establecimiento
    
     controlaestadosrendicion As Boolean
    diasatrazorendicion As Integer
 
     VGLongCodigo As Integer    'Inddica la long de codigo de un articulo
    NomEmpresa As String
     RucEmpresa As String
     puntovta As String
     empresacodigo As String
     MesProceso As String
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
 
 sistemactaajustehab As String
 sistemactaajustedeb As String

 nrofactura As String
 ventaauto As String
 nroboleta As String
 nroguia As String
 administraproyectos As Integer
 SaldoConsolidadoxPedidos As Integer
 listaPuntoVtas As String
 listacajas As String
 
 multiguias As Boolean
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
    copiastic As Integer
    copiasGr As Integer
    emitehoja As String
    valorizaliqcompra As String
    canje As Boolean
    
End Type

Public VGParamSistem As ParametrosdeSistema
Public Type ParametrosdeSistema
    TablaCabcomprob As String
    tabladetcomprob As String
    MesProceso As String
    AnoProceso As String
    fechatrabajo As Date
    RutaReport As String
    
    Servidor As String
    BDEmpresa As String
    Usuario As String
    Pwd      As String
    
    BDEmpresaCONF As String
    
    ServidorCT As String
    BDEmpresaCT As String
    UsuarioCT As String
    PwdCT     As String
    
    
    ServidorGEN As String
    BDEmpresaGEN As String
    UsuarioGEN As String
    PwdGEN As String
    
    Mensaje As String

    listapre As String
    carpetareportes As String
    stockcomp As Integer
    tipocambio As Double
    almacen As String
    comivende As String
    tipoanaliticocodigo As String
    
    Nombre As String
    tienedscto As String
    descuentoas As Double
    moneda As String
    tieneigv As String
    Igv As Double
    formaemi As String
    paraboleta As String
    paramvtamasivo As String
    kitvirtual As String
    tesoreriaenlinea As String
   descuento As String
End Type




'/-------------------------------------
'/ Modificado por Carlos Enrique
'/                2005/09/24
'/ Decia: Public nsaldo As Integer
'/ Debe decir: Public nsaldo As long
'/-------------------------------------
Public nsaldo As Long
Public g_ptoventa As String

Public Sub CargarParametrosVentas()
Dim rs As New ADODB.Recordset
Dim rb As New ADODB.Recordset
Dim averi As New dllgeneral.dll_general
   
     
   g_PedidoPuntoVta = "vt_Tempopedido" & Trim(g_ptoventa)
   g_DetallePuntoVta = "vt_Tempodetallepedido" & Trim(g_ptoventa)
   
   Set rs = VGCNx.Execute("select TOP 1 * from vt_parametroventa ")
   If rs.RecordCount > 0 Then
       VGParamSistem.Nombre = Escadena(rs!empresacodigo)
      VGParamSistem.tienedscto = Escadena(rs!paramvtaestdesc)
      VGParamSistem.descuento = IIf(IsNull(rs!paramvtadescto), 0, CDbl(rs!paramvtadescto))
      VGParamSistem.moneda = Escadena(rs!monedacodigo)
      VGParamSistem.tieneigv = IIf(IsNull(rs!paramvtaestigv) Or rs!paramvtaestigv = 0, "0", "1")
      VGParamSistem.Igv = IIf(IsNull(rs!paramvtaporcigv), 0, CDbl(rs!paramvtaporcigv))
      VGParamSistem.almacen = Escadena(rs!almacencodigo)
      VGParamSistem.Mensaje = Escadena(rs!paramvtamensaje)
      VGParamSistem.listapre = IIf(IsNull(rs!paramvtalistaprec), "", Escadena(rs!paramvtalistaprec))
      VGParamSistem.tipocambio = IIf(IsNull(rs!paramvtatipcambref), CDbl(0), CDbl(rs!paramvtatipcambref))
      VGParamSistem.comivende = IIf(IsNull(rs!paramvtacomisionvendedor), "", Escadena(rs!paramvtacomisionvendedor))
      VGParamSistem.formaemi = IIf(IsNull(rs!paramvtaformaemision), "", Escadena(rs!paramvtaformaemision))
      VGParamSistem.paraboleta = IIf(IsNull(rs!paramvtaboleta) Or rs!paramvtaboleta = 0, "0", "1")
      VGParamSistem.paramvtamasivo = IIf(IsNull(rs!paramvtamasivo) Or rs!paramvtamasivo = 0, "0", "1")
      VGParamSistem.stockcomp = IIf(IsNull(rs!stockcomp) Or rs!stockcomp = 0, "0", "1")
      VGParamSistem.kitvirtual = IIf(IsNull(rs!kitvirtual) Or rs!kitvirtual = 0, "0", "1")
      VGParamSistem.tesoreriaenlinea = IIf(IsNull(rs!tesoreriaenlinea) Or rs!tesoreriaenlinea = 0, "0", "1")
      
   
   End If
   rs.Close
   
   SQL = "select * from vt_puntovtadocumento where puntovtacodigo='" & g_ptoventa & "' AND empresacodigo='" & VGParametros.empresacodigo & "'"
   Set rb = VGCNx.Execute(SQL)
   If rb.RecordCount > 0 Then
          rb.MoveFirst
          Do Until rb.EOF
             If g_tipobol = Escadena(rb!documentocodigo) Then
                g_bolserie = Escadena(rb!puntovtadocserie)
             ElseIf g_tipofac = Escadena(rb!documentocodigo) Then
                g_facserie = Escadena(rb!puntovtadocserie)
             ElseIf g_tipoped = Escadena(rb!documentocodigo) Then
                g_pedserie = Escadena(rb!puntovtadocserie)
             ElseIf g_tipoguia = Escadena(rb!documentocodigo) Then
                g_guiaserie = Escadena(rb!puntovtadocserie)
             ElseIf g_tipoticketB = Escadena(rb!documentocodigo) Then
                g_ticserieb = Escadena(rb!puntovtadocserie)
             ElseIf g_tipoticketf = Escadena(rb!documentocodigo) Then
                g_ticserief = Escadena(rb!puntovtadocserie)
             End If
             rb.MoveNext
          Loop
   End If
   rb.Close
   Set rb = Nothing
   
   Set rs = VGCNx.Execute("select * from vt_puntoventa where puntovtacodigo='" & g_ptoventa & "'")
   If rs.RecordCount > 0 Then
        VGParametros.puntovta = Escadena(rs!puntovtacodigo)
        VGParametros.nropedido = Escadena(IIf(IsNull(rs!puntovtanropedido) Or rs!puntovtanropedido = 0, "0", "1"))
        VGParametros.nrofactura = Escadena(IIf(IsNull(rs!puntovtanrofact) Or rs!puntovtanropedido = 0, "0", "1"))
        VGParametros.nroguia = Escadena(IIf(IsNull(rs!puntovtanroguiarem) Or rs!puntovtanropedido = 0, "0", "1"))
   '     VGParametros.nroabono = Escadena(IIf(IsNull(rs!puntovtanotaabono) Or rs!puntovtanropedido = 0, "0", "1"))
   '    VGParametros.nrocargo = Escadena(IIf(IsNull(rs!puntovtanotacargo) Or rs!puntovtanropedido = 0, "0", "1"))
        VGParametros.ventaauto = Escadena(IIf(IsNull(rs!puntovtaautomat) Or rs!puntovtanropedido = 0, "0", "1"))
    '    VGParametros.nroticket = Escadena(IIf(IsNull(rs!puntovtaticket) Or rs!puntovtanropedido = 0, "0", "1"))
   '     VGParametros.codigocajaVtas = IIf(Escadena(rs!codigocajaVtas) = "", "01", Escadena(rs!codigocajaVtas))
        VGParametros.administraproyectos = ESNULO(rs!administraproyectos, 0)
   
   End If
   rs.Close
  VGCNx.Execute "set dateformat dmy"    '--seteo de formato de fecha
   If averi.VerificaDatoExistente(VGCNx, "select * from sysobjects where name LIKE '" & g_PedidoPuntoVta & "%'") = 0 Then
        VGCNx.Execute "select * into " & g_PedidoPuntoVta & " from vt_pedido"
        VGCNx.Execute "delete from " & g_PedidoPuntoVta
   End If
   If averi.VerificaDatoExistente(VGCNx, "select * from sysobjects where name LIKE '" & g_DetallePuntoVta & "%'") = 0 Then
        VGCNx.Execute "select * into " & g_DetallePuntoVta & " from vt_detallepedido"
        VGCNx.Execute "delete from " & g_DetallePuntoVta
   End If
      
      
   If averi.VerificaDatoExistente(VGCNx, "select * from sysobjects where name Like 'cotizalibre%'") = 0 Then
        VGCNx.Execute "select * into cotizalibre from vt_pedido"
        VGCNx.Execute "delete from cotizalibre"
   End If
   If averi.VerificaDatoExistente(VGCNx, "select * from sysobjects where name Like 'detallecotizalibre%'") = 0 Then
        VGCNx.Execute "select * into detallecotizalibre from vt_detallepedido"
        VGCNx.Execute "delete from detallecotizalibre"
   End If
   
      
   If averi.VerificaDatoExistente(VGCNx, "select * from sysobjects where name Like 'jtempo%'") = 0 Then
        VGCNx.Execute "select * into jtempo from vt_pedido"
        VGCNx.Execute "delete from jtempo"
   End If
   If averi.VerificaDatoExistente(VGCNx, "select * from sysobjects where name Like 'jdetatempo%'") = 0 Then
        VGCNx.Execute "select * into jdetatempo from vt_detallepedido"
        VGCNx.Execute "delete from jdetatempo"
   End If
   
   If averi.VerificaDatoExistente(VGCNx, "select * from sysobjects where name Like 'gtempfile%'") = 0 Then
        VGCNx.Execute "Create table gtempfile " & _
                    "( detpedcantpedida char(8)," & _
                    "productocodigo char(20)," & _
                    "productodescripcion char(100)," & _
                    "detpedmontoprecvta float," & _
                    "detpedimpbruto float," & _
                    "detpeddsctoxitem float," & _
                    "detpedfactorconv float," & _
                    "unidadcodigo char(3))"
   End If
   
   If averi.VerificaDatoExistente(VGCNx, "select * from sysobjects where name Like 'tempfile%'") = 0 Then
        VGCNx.Execute "Create table tempfile " & _
                    "( detpedcantpedida char(8)," & _
                    "productocodigo char(20)," & _
                    "productodescripcion char(100)," & _
                    "detpedmontoprecvta float," & _
                    "detpedimpbruto float," & _
                    "detpeddsctoxitem float," & _
                    "detpedfactorconv float," & _
                    "unidadcodigo char(3))"
   End If
   Set rs = VGCNx.Execute("select * from ct_sistema")
   If rs.RecordCount > 0 Then
     VGParametros.sistemactaajustehab = RTrim(rs!sistemactaajustehab)
     VGParametros.sistemactaajustedeb = RTrim(rs!sistemactaajustedeb)
   End If
   Set rs = VGCNx.Execute("select * from vt_sistema")
   If rs.RecordCount > 0 Then
     VGParamSistem.tipoanaliticocodigo = rs!tipoanaliticocodigo
   End If
   Set rs = VGCNx.Execute("select * from dbo.te_parametroempresa")
   If rs.RecordCount > 0 Then
     VGParametros.empresaasientosautomaticos = ESNULO(rs!empresaasientosautomaticos, "0")
   End If
If IsNumeric(VGParamSistem.AnoProceso) And IsNumeric(VGParametros.MesProceso) Then
        SQL = "select * from ct_cierremensual where empresacodigo='" & VGParametros.empresacodigo & "' and " _
        & " anio='" & VGParamSistem.AnoProceso & "' and mes=" & Trim(VGParametros.MesProceso) & " "
        Set rs = VGCNx.Execute(SQL)
        If rs.RecordCount > 0 Then VGParametros.cierremes = IIf(rs!Ventas = True, True, False)
End If


End Sub
