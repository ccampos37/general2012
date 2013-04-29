Attribute VB_Name = "MainVentas"
Option Explicit
Public Cadenabusca As String

Public VgModificar As Byte
'Variables de Ventana de Ayuda
Public nAyuda As String
Public nDetalle As String

Public vgorden As String

Public VGdllApi As dll_apisgen.dll_apis

'Variables de acceso de usuario

Public g_Ip As String
Public g_PedidoPuntoVta As String
Public g_DetallePuntoVta As String

Public g_TipoMovi As String
Public g_usuario As String
Public g_ptoventa As String
Public g_pedserie As String
Public g_facserie As String
Public g_ticserieb As String
Public g_ticserief As String
Public g_bolserie As String
Public g_guiaserie As String
Public g_GuiaRemSerie As String


Public vgtipo As TIPOSISTEMA
Public VGformatofecha As String




'REPORTES

Public Const g_tipoped = "PE"    'PEDIDO
Public Const g_tipocot = "CO"    'PEDIDO
Public Const g_tipoticketB = "14"    'Tickets boletas
Public Const g_tipoticketf = "15"    'facturas
Public Const g_tipofac = "01"    'factura
Public Const g_tipobol = "03"    'boletas
Public Const g_tipoguia = "80"    '9"   'guias B.O
Public Const g_tipoguiaRem = "GR"    'GR"   'guias REM
Public Const g_Todos = "99"      'Todos los documentos para Reporte
Public Const g_Eventual = "88888888888" 'cliente eventual

Public VGParametros  As PuntoVenta
Public modoventa As modoventa
Public VGParamSistem As ParametrosdeSistema

Private Type ParametrosdeSistema         ' Crea Tipo de Empresa
   nombre As String * 2
   moneda As String * 2
   tieneigv As String * 1
   tienedscto As String * 1
   descuento As Double
   Igv As Double
   mensaje As String * 70
   almacen As String * 2
   listapre As String * 1
   tipocambio As Double
   comivende As String * 1
   formaemi As String * 1
   paraboleta As String * 1
   paramvtamasivo As String * 1
   stockcomp As String * 1
   AnoProceso As String * 4
   kitvirtual As String * 1
   tesoreriaenlinea As Integer
   tipoanaliticocodigo As String
   familiaproyectos As String
  
   BDEmpresa As String
   Usuario As String
   Servidor As String
   PWD      As String
   UsuarioReporte As String
   BDempresaCONF As String
   
   ServidorGEN As String
   BDEmpresaGEN As String
   UsuarioGEN As String
   PwdGEN As String
    
    ServidorCT As String
    BDEmpresaCT As String
    UsuarioCT As String
    PwdCT As String
    
    
    carpetareportes As String
    Rutareport As String
    FechaTrabajo As Date
End Type


Private Type PuntoVenta   'Crea Punto de Venta
    puntovta As String * 2
    nropedido As String * 1
    nroguia As String * 1
    nroticket As String * 1     'agregado para ticket
    nrofactura As String * 1
    nroboleta As String * 1
    nroabono As String * 1
    nrocargo As String * 1
    ventaauto As String * 1
    seriefac As String * 3
    seriebol As String * 3
    serieguia As String * 3
    serieped As String * 3
    serie99 As String * 3
    nomempresa As String * 30
    RucEmpresa As String * 11
    empresacodigo As String * 2
    sistemamultiempresas As Boolean
    multiguias As Boolean
    multifacturas As Boolean
    multiboletas As Boolean
    cajerocodigo As String * 2
    sistemactaajustehab As String
    sistemactaajustedeb As String
    codigocajaVtas As String
    cierremes As Boolean
    mesproceso As String
    administraproyectos As Integer
    empresaasientosautomaticos As String
    
    listaPuntoVtas As String
End Type


Private Type modoventa   'Crea modoventa
    descuento As String * 1
    impuestos As String * 1
    nroitem  As Double
    numeraauto As String * 1
    ctrlinventario As String * 1
    unidadmedida As String * 1
    copiasfac As Double
    copiastic As Double
    copiasbol As Double
    copiashoja As Double
    copiasGr As Double
    ctacte As String * 1
    ingcliente As String * 1
    ingforma As String * 1
    emitehoja As String * 1
    emiteguia As String * 1
    emitefact As String * 1
    modificaguia As String * 1
    ingpedido As String * 1
    inghoja As String * 1
    ingguia As String * 1
    usafactor As String * 1
    documento As String * 2
    almacenes As String * 100
    canje As Boolean
End Type


Public Sub Cargar_Parametros_Funcionales()
   Dim rs As New ADODB.Recordset
   Dim rb As New ADODB.Recordset
   Dim averi As New dllgeneral.dll_general
       
   
   g_PedidoPuntoVta = "vt_Tempopedido" & Trim(g_ptoventa)
   g_DetallePuntoVta = "vt_Tempodetallepedido" & Trim(g_ptoventa)
   
   Set rs = VGCNx.Execute("select TOP 1 * from vt_parametroventa ")
   If rs.RecordCount > 0 Then
      VGParamSistem.nombre = Escadena(rs!empresacodigo)
      VGParamSistem.tienedscto = Escadena(rs!paramvtaestdesc)
      VGParamSistem.descuento = IIf(IsNull(rs!paramvtadescto), 0, CDbl(rs!paramvtadescto))
      VGParamSistem.moneda = Escadena(rs!monedacodigo)
      VGParamSistem.tieneigv = IIf(IsNull(rs!paramvtaestigv) Or rs!paramvtaestigv = 0, "0", "1")
      VGParamSistem.Igv = IIf(IsNull(rs!paramvtaporcigv), 0, CDbl(rs!paramvtaporcigv))
      VGParamSistem.almacen = Escadena(rs!almacencodigo)
      VGParamSistem.mensaje = Escadena(rs!paramvtamensaje)
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
              ElseIf g_tipoguiaRem = Escadena(rb!documentocodigo) Then
                g_GuiaRemSerie = Escadena(rb!puntovtadocserie)
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
        VGParametros.nroabono = Escadena(IIf(IsNull(rs!puntovtanotaabono) Or rs!puntovtanropedido = 0, "0", "1"))
        VGParametros.nrocargo = Escadena(IIf(IsNull(rs!puntovtanotacargo) Or rs!puntovtanropedido = 0, "0", "1"))
        VGParametros.ventaauto = Escadena(IIf(IsNull(rs!puntovtaautomat) Or rs!puntovtanropedido = 0, "0", "1"))
        VGParametros.nroticket = Escadena(IIf(IsNull(rs!puntovtaticket) Or rs!puntovtanropedido = 0, "0", "1"))
        VGParametros.codigocajaVtas = IIf(Escadena(rs!codigocajaVtas) = "", "01", Escadena(rs!codigocajaVtas))
        VGParametros.administraproyectos = ESNULO(rs!administraproyectos, 0)
   
   End If
   rs.Close
   Set rs = Nothing
   
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
     VGParamSistem.familiaproyectos = rs!familiaproyectos
   End If
   Set rs = VGCNx.Execute("select * from dbo.te_parametroempresa")
   If rs.RecordCount > 0 Then
     VGParametros.empresaasientosautomaticos = ESNULO(rs!empresaasientosautomaticos, "0")
   End If
If IsNumeric(VGParamSistem.AnoProceso) And IsNumeric(VGParametros.mesproceso) Then
        SQL = "select * from ct_cierremensual where empresacodigo='" & VGParametros.empresacodigo & "' and " _
        & " anio='" & VGParamSistem.AnoProceso & "' and mes=" & Trim(VGParametros.mesproceso) & " "
        Set rs = VGCNx.Execute(SQL)
        If rs.RecordCount > 0 Then VGParametros.cierremes = IIf(rs!Ventas = True, True, False)
End If


End Sub

