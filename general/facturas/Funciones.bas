Attribute VB_Name = "Funciones"
Option Explicit
Public Cadenabusca As String

Public VGconfig As New ADODB.Connection
Public VGgeneral As New ADODB.Connection
Public VGcnxCT As New ADODB.Connection
Public VGCNx As New ADODB.Connection
Public VgModificar As Byte
'Variables de Ventana de Ayuda
Public nAyuda As String
Public nDetalle As String
Public VGcomputer As String
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



Public vgtipo As TIPOSISTEMA
Public VGCODEMPRESA As String
Public VGformatofecha As String




'REPORTES

'Public Const VGParamSistem.Rutareport = "\\desarrollo\librerias_controles\Reportes\"
Public VGcadenareport2 As String

Public Const g_TipoSol = "01"
Public Const g_TipoDolar = "02"

Public Const g_tipoped = "PE"    'PEDIDO
Public Const g_tipocot = "CO"    'PEDIDO
Public Const g_tipoticketB = "14"    'Tickets boletas
Public Const g_tipoticketf = "15"    'facturas
Public Const g_tipofac = "01"    'factura
Public Const g_tipobol = "03"    'boletas
Public Const g_tipoguia = "BO"    '9"   'guias B.O
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

End Type


Public Function FechS(Fecha As Variant, tipo As TIPFECHA) As Variant
Dim h As Date
Dim fechaAux As Double
On Error GoTo ErrorFecha
   h = CDate(Fecha)
   Select Case tipo
      Case Sqlf: 'Para transformar al sql
         fechaAux = DateSerial(Year(Fecha), Month(Fecha), Day(Fecha)) - 2
      Case Adof: 'Para transformar al ado Y AL ACCESS
         fechaAux = DateSerial(Year(Fecha), Month(Fecha), Day(Fecha))
   End Select
   FechS = fechaAux
   Exit Function
ErrorFecha:
   Select Case tipo
      Case Sqlf: FechS = "Null"
      Case Adof: FechS = Null
   End Select
End Function


Public Function TraeTipoCambio(vfecha As Date, vcon As ADODB.Connection) As String
    Dim rsbuscn As New ADODB.Recordset
    
    Set rsbuscn = vcon.Execute("select * from ct_tipocambio where tipocambiofecha='" & vfecha & "'")
    If rsbuscn.RecordCount > 0 Then
        TraeTipoCambio = rsbuscn!tipocambioventa
    Else
        TraeTipoCambio = "1.00"
    End If
    Set rsbuscn = Nothing

End Function


Public Function TraePrecio(vlista As String, vArti As String, vcon As ADODB.Connection, vAlma As String) As Double
    Dim rsbuscn As New ADODB.Recordset
    
    Set rsbuscn = vcon.Execute("select * from listapre" & Trim(vlista) & " where productocodigo='" & vArti & "'")
    If rsbuscn.RecordCount > 0 Then
        TraePrecio = IIf(IsNull(rsbuscn!productoprecvta), 1, rsbuscn!productoprecvta)
    Else
        TraePrecio = 1
    End If
    Set rsbuscn = Nothing

End Function



Public Function TraeDataSerie(nsql As String, vcon As ADODB.Connection) As String
    Dim rsbuscn As New ADODB.Recordset
    Set rsbuscn = vcon.Execute(nsql)
    If rsbuscn.RecordCount > 0 Then
        TraeDataSerie = rsbuscn.Fields(0)    '  !puntovtadoccorr
    Else
        TraeDataSerie = "1"
    End If
    Set rsbuscn = Nothing

End Function

Public Sub CargarTipo(xcombo As ComboBox, xtipo)
  
  Select Case xtipo
    Case 1     '--condicion documento--
     xcombo.Clear
     xcombo.AddItem "0-Activo"
     xcombo.AddItem "1-Anulado"
     xcombo.ListIndex = 0
   Case 2   '--tipodocumento --
     xcombo.Clear
     xcombo.AddItem g_tipobol & "-Boleta"
     xcombo.AddItem g_tipofac & "-Factura"
     xcombo.AddItem g_tipoguia & "-B.O."
     xcombo.AddItem g_tipoticketB & "-Ticket Boleta"
     xcombo.AddItem g_tipoticketf & "-Ticket factura"
     xcombo.ListIndex = 0
   Case 3   '---estado
     xcombo.Clear
     xcombo.AddItem "S-SI"
     xcombo.AddItem "N-NO"
     xcombo.ListIndex = 0
   Case 4  '-- Tipo persona
     xcombo.Clear
     xcombo.AddItem "1-NATURAL"
     xcombo.AddItem "2-JURIDICA"
     xcombo.ListIndex = 0
   Case 5  '-tipo pais
     xcombo.Clear
     xcombo.AddItem "1-PERUANA"
     xcombo.AddItem "2-EXTRANJERA"
     xcombo.ListIndex = 0
   Case 6   '--todos los tipodocumentos --
     xcombo.Clear
     xcombo.AddItem g_tipobol & "-Boleta"
     xcombo.AddItem g_tipofac & "-Factura"
     xcombo.AddItem g_tipoguia & "-B.O."
     xcombo.AddItem g_tipoped & "-Pedido"
     xcombo.AddItem g_tipoticketB & "-Ticket Boleta"
     xcombo.AddItem g_tipoticketf & "-Ticket factura"
     xcombo.AddItem "07-Ticket factura"
     xcombo.ListIndex = 0
     
  End Select
End Sub


Public Sub Main()
   

   Call Configurar_Conexiones
   Call adicionarcampos
   'Call Cargar_Parametros_Funcionales
   FrmIngreso.Show
  
   
   
End Sub

Public Sub Configurar_Conexiones()
  
  On Error GoTo nerror
   ' *****Archivo de Punto de Venta*******
   Dim VGdllApi As New dll_apisgen.dll_apis
   Dim strbase As String
   Dim struser As String
   Dim strpass As String
   Dim strserver As String
   
   Dim strconecta As String
 
    VGcomputer = UCase(ComputerName)
    VGsql = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "CONEXION", "SQL", "?")
    VGsql = IIf(VGsql = "?", 0, VGsql)
    
    VGformatofecha = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "CONEXION", "FECHASQL", "?")
    VGformatofecha = IIf(VGformatofecha = "?", "DMY", VGformatofecha)
    
    VGParametros.empresacodigo = "01"
    
    
      VGParamSistem.BDEmpresaGEN = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "BDGENERAL", "BDDATOS", "?")
    VGParamSistem.ServidorGEN = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "BDGENERAL", "SERVIDOR", "?")
    VGParamSistem.UsuarioGEN = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "BDGENERAL", "USUARIO", "?")
    VGParamSistem.PwdGEN = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "BDGENERAL", "PASSW", "?")
        
    
    'Conexion de Compras
    VGParamSistem.BDEmpresa = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "conexion", "BDDATOS", "?")
    VGParamSistem.Servidor = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "conexion", "SERVIDOR", "?")
    VGParamSistem.Usuario = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "conexion", "USUARIO", "?")
    VGParamSistem.PWD = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "conexion", "PASSW", "?")
    vgorden = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "conexion", "ORDEN", "?")
   
' reportes
   
   VGParamSistem.Rutareport = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "REPORTES", "VENTAS", "?")
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
    If VGParamSistem.Rutareport = "" Or VGParamSistem.Rutareport = "?" Then
       VGParamSistem.Rutareport = App.Path
       VGParamSistem.carpetareportes = "Reportes"
    End If
       
    'Establecer Cadena de Conexión de Reportes
    
    VGcadenareport2 = "DSN=jckconsultores;DSQ=" & VGParamSistem.BDEmpresaGEN & ";UID=" & VGParamSistem.UsuarioGEN & ";PWD=" & VGParamSistem.PwdGEN & ""
        
    'Establecer Conexiones
    Set VGgeneral = New ADODB.Connection
    VGgeneral.CursorLocation = adUseClient
    VGgeneral.CommandTimeout = 0
    VGgeneral.ConnectionTimeout = 0
    VGgeneral.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & VGParamSistem.UsuarioGEN & ";Password=" & VGParamSistem.PwdGEN & ";Initial Catalog=" & VGParamSistem.BDEmpresaGEN & ";Data Source=" & VGParamSistem.Servidor
    VGgeneral.Open
   
   'Call AdjuntarServ(VGGeneral, VGParamSistem.Servidor)
   
   'Conexion de Compras el Principal
    Set VGCNx = New ADODB.Connection
    VGCNx.CursorLocation = adUseClient
    VGCNx.CommandTimeout = 0
    VGCNx.ConnectionTimeout = 0
    VGCNx.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & VGParamSistem.Usuario & ";Password=" & VGParamSistem.PWD & ";Initial Catalog=" & VGParamSistem.BDEmpresa & ";Data Source=" & VGParamSistem.Servidor
    VGCNx.Open
    
   
   'Conexion de Cofiguracion
    Set VGconfig = New ADODB.Connection
    VGconfig.CursorLocation = adUseClient
    VGconfig.CommandTimeout = 0
    VGconfig.ConnectionTimeout = 0
    VGconfig.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & VGParamSistem.Usuario & ";Password=" & VGParamSistem.PWD & ";Initial Catalog=bdwenco;Data Source=" & VGParamSistem.Servidor
    VGconfig.Open
  

      
   Exit Sub
nerror:
   If Err Then
       MsgBox "Comunicarse con Sistemas " & Chr(13) & Chr(10) & Err.Number & "-" & Err.Description
       Err = 0
       End
   End If

End Sub

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


Public Function MostrarForm(pVentana As Form, pPos As String)
   pVentana.Icon = LoadPicture(App.Path & "\factu.ico")
   
   If pPos = "C" Then
     pVentana.Left = ((Screen.Width - pVentana.Width) / 2) - 350
     pVentana.Top = ((Screen.Height - pVentana.Height) / 2) - 350
   ElseIf pPos = "I" Then
      pVentana.Left = 300
      pVentana.Top = 300
   ElseIf pPos = "M" And pVentana.Visible = False Then
      pVentana.Caption = pVentana.Caption & "  " & VGParamSistem.BDEmpresa
      pVentana.Width = Screen.Width
   ElseIf pPos = "C1" Then
     pVentana.Left = ((Screen.Width - pVentana.Width) / 2) - 350
     pVentana.Top = ((Screen.Height - pVentana.Height) / 2) - 350
     Exit Function
   ElseIf pPos = "C2" Then
     pVentana.Left = ((Screen.Width - pVentana.Width) / 2) - 350
     pVentana.Top = ((Screen.Height - pVentana.Height) / 2) - 350
     Exit Function
   End If
   If pPos = "M" Then
      pVentana.StatusBar1.Panels(1).Text = "EMPRESA: " & VGParametros.nomempresa
      pVentana.StatusBar1.Panels(2).Text = "PTO. VENTA: " & g_ptoventa
      pVentana.StatusBar1.Panels(2).Alignment = sbrLeft
   Else
 '     pVentana.StatusBar1.Panels(1).Text = "FORMATO : " & Escadena(pVentana.Caption)
 '     pVentana.StatusBar1.Panels(2).Text = "USUARIO: " & g_usuario
 '     pVentana.StatusBar1.Panels(2).Alignment = sbrLeft
   End If
'   pVentana.StatusBar1.Panels(1).Alignment = sbrLeft
'   pVentana.StatusBar1.Panels(3).Text = "FECHA :" & Format(Date, "dd/mm/yyyy")
'   pVentana.StatusBar1.Panels(3).Alignment = sbrRight
'   pVentana.StatusBar1.Panels(4).Text = "HORA :" & Format(Time, "hh:mm:ss")
'   pVentana.StatusBar1.Panels(4).Alignment = sbrRight
 
End Function

Public Function Limpiartexto(MBox As Object, ninicio As Integer, nfin As Integer, Optional Noincluir1, Optional Noincluir2 As Integer)
 Dim J As Integer
 If IsMissing(Noincluir1) Then
    Noincluir1 = -1
 End If
 If IsMissing(Noincluir2) Then
    Noincluir2 = -1
 End If
   
   For J = ninicio To nfin

         If J = Noincluir1 Or J = Noincluir2 Then
         Else
            MBox(J) = ""
         End If
 
   Next J
End Function

Public Function numero(Number) As String
   Dim aValor As Double
   If IsNull(Number) Or Len(Trim(Number)) = 0 Then
     aValor = 0
   Else
     aValor = Number
   End If
   numero = Right(Space(12) & Trim(Format(aValor, "####,###0.0000")), 14)
End Function

Public Function Escadena(pDato) As String
   If IsNull(pDato) Or Len(Trim(pDato)) = 0 Then
     Escadena = ""
   Else
     Escadena = Trim(pDato)
   End If
End Function

Public Function Seguir(MBox As Object, ntecla As Integer)
If ntecla = 13 Then SendKeys "{tab}"
End Function


Public Function DatoMoneda(xValor As String) As String
   Dim rmone As New ADODB.Recordset
   
   Set rmone = VGCNx.Execute("select * from gr_moneda where monedacodigo='" & xValor & "'")
   If rmone.RecordCount > 0 Then
       DatoMoneda = Escadena(rmone!monedasimbolo) & " ."
   Else
       DatoMoneda = " "
   End If
   rmone.Close
   Set rmone = Nothing

End Function

Public Function VerificaCombo(xcombo As ComboBox, ncadena As String) As Long
    Dim J, k As Long
    On Error GoTo nerror
    VerificaCombo = -1
    If xcombo.ListCount > 0 Then
      VerificaCombo = 0
      For J = 0 To xcombo.ListCount - 1
         xcombo.ListIndex = J
         k = InStr(xcombo, "-")
         If k > 1 Then
           If Left(xcombo.Text, k - 1) = ncadena Then
             VerificaCombo = J
             Exit For
           End If
         Else
           If xcombo.Text = ncadena Then
             VerificaCombo = J
             Exit For
           End If
         End If
      Next J

    End If
    
nerror:
  If Err Then
    MsgBox Err.Number & "-" & Err.Description
    Err = 0
    Resume Next
  End If
End Function

Public Sub Imprimir(cNombreReporte As String)
Dim busca As New dll_apisgen.dll_apis
On Error GoTo Errores

   MDIPrincipal.oCrystalReport.Reset
   MDIPrincipal.oCrystalReport.Destination = crptToWindow
   MDIPrincipal.oCrystalReport.WindowState = crptMaximized
   'MDIPrincipal.oCrystalReport.ReportFileName = VGParamSistem.Rutareport & cNombreReporte
   
   MDIPrincipal.oCrystalReport.LogOnServer "pdssql.dll", _
         busca.LeerIni(App.Path & "\Marfice.ini", "CONEXION", "SERVIDOR", ""), _
         busca.LeerIni(App.Path & "\Marfice.ini", "CONEXION", "BDDATOS", ""), _
         busca.LeerIni(App.Path & "\Marfice.ini", "CONEXION", "USUARIO", ""), _
         busca.LeerIni(App.Path & "\Marfice.ini", "CONEXION", "PASSW", "")
   MDIPrincipal.oCrystalReport.Connect = _
        "DSN=" & busca.LeerIni(App.Path & "\Marfice.ini", "CONEXION", "SERVIDOR", "") & ";" & _
        "DSQ=" & busca.LeerIni(App.Path & "\Marfice.ini", "CONEXION", "BDDATOS", "") & ";" & _
        "UID=" & busca.LeerIni(App.Path & "\Marfice.ini", "CONEXION", "USUARIO", "") & ";" & _
        "PWD=" & busca.LeerIni(App.Path & "\Marfice.ini", "CONEXION", "PASSW", "")
   
   MDIPrincipal.oCrystalReport.DiscardSavedData = True
   MDIPrincipal.oCrystalReport.formulas(0) = "Empresa='" & VGParametros.nomempresa & "'"
   MDIPrincipal.oCrystalReport.Action = 1
  
   Exit Sub
   
Errores:
  MsgBox "Error inesperado: " & Err.Number & "  " & Err.Description, vbExclamation
  Err = 0
  Exit Sub
  
End Sub



Public Function aImpresora(wFile)
  Dim wbat, wcade As String
  Dim X As Long
  Dim rnrofile As Double
  
  On Error GoTo nerror
  rnrofile = CInt((90 * (Rnd(10) + 1)))
  wbat = "c:\printer" & CStr(Val(Right(wFile, 5))) & ".bat"
  Open wbat For Output As #rnrofile
  Print #rnrofile, "@echo off"
  Print #rnrofile, "Type " & wFile & " >" & Left(Printer.Port, Len(Printer.Port) - 1)
  Print #rnrofile, "cls"
  Print #rnrofile, "exit"
  Close #rnrofile
  wcade = "start /m " & Trim(wbat)
  X = Shell(wcade, vbHide)
  DoEvents
  
nerror:
   If Err Then
      MsgBox "Error: " & Err.Number & "-" & Err.Description, vbCritical, "AVISO"
      Err = 0
   End If

End Function
Public Sub GeneraAsientoEnlineaTesor(Fecha As Date, Empresa As String, m_opcion As String, Nrecibo As String, op As Integer, comprobconta As String, monedacodigo As String, cajabanco As String, m_tipovoucher As String)
Dim rsparimpo As ADODB.Recordset
Dim numerror As Integer
Dim Comando As ADODB.Command
numerror = 0
On Error GoTo Proceso

   VGCNx.BeginTrans

Set rsparimpo = New ADODB.Recordset
SQL = " Select * From  ct_importartesoreria Where tipooperacion <>'T' and Left(Upper(tipocajabanco),1) ='" & m_opcion & "' And monedacodigo='" & monedacodigo & "' "
Set rsparimpo = VGcnxCT.Execute(SQL)
If rsparimpo.RecordCount() > 0 Then

   Set Comando = New ADODB.Command
   With Comando
        .CommandType = adCmdStoredProc
        .CommandText = "te_GeneraAsientosTesoreriaLinea_pro"
        .CommandTimeout = 0
        .ActiveConnection = VGgeneral
        .Parameters.Refresh
        .Parameters("@BaseConta") = VGcnxCT.DefaultDatabase
        .Parameters("@BaseVenta") = VGCNx.DefaultDatabase
        .Parameters("@empresa") = Empresa
        .Parameters("@Asiento") = rsparimpo!asiento
        .Parameters("@SubAsiento") = rsparimpo!SubAsiento
        .Parameters("@Libro") = rsparimpo!Libro
         
        .Parameters("@Mes") = Format(Month(Fecha), "00")
        .Parameters("@Ano") = Year(Fecha)
            
        .Parameters("@tipanal") = "002"
        .Parameters("@Compu") = VGcomputer
        .Parameters("@Usuario") = VGParamSistem.Usuario
        .Parameters("@TipoMov") = Trim(UCase(m_opcion))
        .Parameters("@Nrecibo") = Nrecibo
        .Parameters("@op") = op
        .Parameters("@comprobconta") = comprobconta
        .Parameters("@ajustehaber") = VGParametros.sistemactaajustehab
        .Parameters("@ajustedebe") = VGParametros.sistemactaajustedeb
        .Execute
   End With
   If numerror = 0 Then
        VGCNx.CommitTrans
        Screen.MousePointer = 1
        MsgBox "La Contabilizacion se Realizo Satisfactoriamente", vbInformation, "Sistema de Tesoreria"
   End If
End If
Exit Sub
Proceso:
   numerror = 1
   Screen.MousePointer = 1
    MsgBox Err.Description
    VGCNx.RollbackTrans
   Exit Sub
   Resume
End Sub

Public Sub exportarExcel(RSSQL1 As ADODB.Recordset, Titulo As String)
On Error GoTo ErrorExcel
Dim objExcel As Excel.Application
Dim HNom As Integer 'Horizontal
Dim VNom As Integer 'Vertical
Dim Hdatos As Integer 'Horizontal
Dim Vdatos As Integer 'Vertical
Dim cuentaNombres As Integer
Dim cuentadatos As Integer
Dim i As Integer
Dim n As Integer
Dim J As Integer

If RSSQL1.RecordCount <> 0 Then
   'Crear un objeto del tipo excel.application

   cuentaNombres = RSSQL1.Fields.Count
   cuentadatos = RSSQL1.RecordCount

   Set objExcel = New Excel.Application
   objExcel.Visible = True
   objExcel.SheetsInNewWorkbook = 1
   objExcel.Workbooks.Add

    'PONER UN TITULO
    objExcel.ActiveSheet.Cells(1, 1) = "EXPORTAR A EXCEL - " + Titulo
    objExcel.ActiveSheet.Cells(2, 1) = cuentadatos
    objExcel.ActiveSheet.Cells(2, 2) = cuentaNombres
    With objExcel.ActiveSheet.Cells(1, 1).Font
      .Color = vbBlack
      .Size = 12
      .Bold = True
   End With

   'UTILIZAMOS LAS VARIABLES PARA LA UBICACION DE NUESTROS TEXTOS
   HNom = 1
   VNom = 4
   Vdatos = 5
   Hdatos = 1


   'AGREGAMOS LOS REGISTROS (RECUERDEN QUE NO IMPORTA CUANTAS COLUMNAS O REGISTROS TENGAMOS EL BUCLE_
   'FUNCIONA SEGUN EL NUMERO DE CABECERAS Y REGISTROS
  
    For i = 0 To (cuentaNombres - 1)
       objExcel.ActiveSheet.Cells(VNom, HNom) = RSSQL1.Fields(i).Name
       objExcel.ActiveSheet.Range(objExcel.ActiveSheet.Cells(VNom, HNom), objExcel.ActiveSheet.Cells(VNom, HNom)).HorizontalAlignment = xlHAlignCenterAcrossSelection
       With objExcel.ActiveSheet.Cells(VNom, HNom).Font
          .Size = 12
          .Color = vbRed
          .Bold = True
       End With
       RSSQL1.MoveFirst
       For n = 1 To RSSQL1.RecordCount
         objExcel.ActiveSheet.Cells(Vdatos, Hdatos) = RSSQL1.Fields(i).Value
         objExcel.ActiveSheet.Cells(Vdatos, Hdatos).Font.Size = 10
         Vdatos = Vdatos + 1
         RSSQL1.MoveNext
       Next
       HNom = HNom + 1
       Hdatos = Hdatos + 1
       Vdatos = 5
       RSSQL1.MoveFirst
   Next i
   'AHORA LE ASIGNAMOS UN TAMAÑO A CADA COLUMNA SEGUN NESECITEMOS
    objExcel.Columns("B").ColumnWidth = 15.43
    objExcel.Columns("C").ColumnWidth = 15.43
    objExcel.Columns("D").ColumnWidth = 25.86
    objExcel.Columns("E").ColumnWidth = 15.83
End If
Exit Sub
ErrorExcel:
MsgBox Err.Description
End Sub

Public Sub importarExcel(Tabla As Recordset, hojadecalculo As String, DataGrid1 As DataGrid)
    'Variable de tipo Aplicación de Excel
    Dim objExcel As Excel.Application
 
    'Una variable de tipo Libro de Excel
    Dim xLibro As Excel.Workbook
    Dim fila As Integer
    Dim Columna As Integer
        'creamos un nuevo objeto excel
    Set objExcel = New Excel.Application
    objExcel.SheetsInNewWorkbook = 1
  
    'Usamos el método open para abrir el archivo que está _
     en el directorio del programa llamado archivo.xls
    Set xLibro = objExcel.Workbooks.Open("" & hojadecalculo & "")
  
    'Hacemos el Excel Visible
    objExcel.Visible = True
    Set DataGrid1.DataSource = Tabla
    With xLibro
  
        ' Hacemos referencia a la Hoja
        With .Sheets(1)

            'Recorremos la fila desde la 1 hasta la 7
            For fila = 5 To 30000
                If .Cells(fila, 1) = "" Then Exit For
                Tabla.AddNew
                For Columna = 1 To 12
                'Agregamos el valor de la fila que _
                 corresponde a la columna 2
                Tabla.Fields(Columna - 1) = .Cells(fila, Columna)
                Next
                Tabla.Update
                DataGrid1.Refresh
            Next
         
        End With
    End With
    Tabla.Update
    Tabla.UpdateBatch adAffectAllChapters
 
  
    'Eliminamos los objetos si ya no los usamos
    Set objExcel = Nothing
    Set xLibro = Nothing
  
End Sub



