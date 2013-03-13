Attribute VB_Name = "Funciones"
Option Explicit
Public Cadenabusca As String

Public cn As New ADODB.Connection
Public VGGeneral As New ADODB.Connection
Public VGcnxCT As New ADODB.Connection
Public VGCNx As New ADODB.Connection
Public vgconfig As New ADODB.Connection
Public VGCODEMPRESA As String
Public VGCadenaReport2 As String
Public VGComputer As String
Public VGformatofecha As String
Public VGPlanillaAjuste As Integer

'Variables de Ventana de Ayuda
Public nAyuda As String
Public nDetalle As String
Public nfecha As String
Public nsaldo As Double
Public nestado As Integer

Public g_BaseContab As String
Public g_Empresa As String
Public g_DetalleEmpresa As String
Public g_PedidoPuntoVta As String
Public g_DetallePuntoVta As String
Public SQL As String

Public VGvardllgen As dllgeneral.dll_general 'Dll de Algunas funciones
Public VGdllApi As dll_apisgen.dll_apis

Public g_TipoMovi As String
Public g_usuario As String
Public g_ptoventa As String
Public g_pedserie As String
Public g_facserie As String
Public g_bolserie As String
Public g_guiaserie As String

Public Const g_TipoSol = "01"
Public Const g_TipoDolar = "02"

Public punto As PuntoVenta
Public Tipodocu As Tipodocu

'Variables de acceso de usuario
Public Const g_tipoped = "PE"    'PEDIDO
Public Const g_tipofac = "01"    'factura
Public Const g_tipobol = "03"    'boletas
Public Const g_tipoguia = "80"    '9"   'guias B.O
Public Const g_Todos = "99"      'Todos los documentos para Reporte
Public Const g_Eventual = "88888888888" 'cliente eventual

'Constantes de mensajes para visualizar
Public Const MsgEdit = "No Existen Datos para Editar.. "
Public Const MsgGraba = "Datos Grabados satisfactoriamente...."
Public Const MsgElim = "No Existen Datos a Eliminar.."
Public Const MsgAdd = "Los datos ya existen...Verifique!!!"
Public Const MsgTitle = "AVISO"
Public Const Msg29 = "Debe Ingresar Numeros"

'REPORTES

Public RutaRep  As String
Public RutaRepProc As String


Private Type Parametrocobrar         ' Crea Tipo de Empresa
   puntovta As String
   nombre As String
   moneda As String
   tieneigv As String
   igv As Double
   tienedscto As String
   descuento As Double
   empresacodigo As String
   mensaje As String
   almacen As String
   listapre As String
   tipocambio As Double
   comivende As String
   formaemi As String
   paraboleta As String
   NomEmpresa As String
   RucEmpresa As String
   sistemamultiempresas As Boolean
   contabilizaenlinea As Boolean
   cierremes As Boolean

   
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
    
    bdempresa As String
    Usuario As String
    pwd As String
    servidor As String
    
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
End Type
Public VGparametros As Parametrocobrar
Public VGparamsistem As PuntoVenta
Public modoventa As modoventa
Public VGtipo As TIPOSISTEMA

Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long


'FIXIT: Declare 'FechS' and 'fecha' con un tipo de datos de enlace en tiempo de compilación     FixIT90210ae-R1672-R1B8ZE
Public Function FechS(Fecha As Variant, Tipo As TIPFECHA) As Variant
Dim H As Date
Dim fechaAux As Double
fechaAux = 0
On Error GoTo ErrorFecha
   H = CDate(Fecha)
   Select Case Tipo
      Case Sqlf: 'Para transformar al sql
         fechaAux = DateSerial(Year(Fecha), Month(Fecha), Day(Fecha)) - 2
      Case Adof: 'Para transformar al ado Y AL ACCESS
         fechaAux = DateSerial(Year(Fecha), Month(Fecha), Day(Fecha))
   End Select
   FechS = fechaAux
   Exit Function
ErrorFecha:
   Select Case Tipo
      Case Sqlf: FechS = Null
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

Public Function TraeDataSerie(nsql As String, vcon As ADODB.Connection) As String
    Dim rsbuscn As New ADODB.Recordset
    
    Set rsbuscn = vcon.Execute(nsql)
    If rsbuscn.RecordCount > 0 Then
        TraeDataSerie = rsbuscn!puntovtadoccorr
    Else
        TraeDataSerie = "1"
    End If
    Set rsbuscn = Nothing

End Function

'FIXIT: Declare 'xtipo' con un tipo de datos de enlace en tiempo de compilación            FixIT90210ae-R1672-R1B8ZE
Public Sub CargarTipo(xcombo As ComboBox, xtipo As Integer)
  Dim adll2 As New dllgeneral.dll_general
  
  Select Case xtipo
    Case 1     '--condicion documento--
     xcombo.Clear
     xcombo.AddItem "0-Activo"
     xcombo.AddItem "1-Anulado"
     xcombo.ListIndex = 0
   Case 2   '--tipodocumento --
     xcombo.Clear
     Call adll2.llenacombo(xcombo, "select documentocodigo,documentodescripcion from vt_documento", VGCNx)
'     xcombo.AddItem g_tipobol & "-Boleta"
'     xcombo.AddItem g_tipofac & "-Factura"
'     xcombo.AddItem g_tipoguia & "-B.O."
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
   Case 6   '--todos los tipos documentos --
     xcombo.Clear
     Call adll2.llenacombo(xcombo, "select documentocodigo,documentodescripcion from vt_documento ", VGCNx)
     'xcombo.AddItem g_tipobol & "-Boleta"
     'xcombo.AddItem g_tipofac & "-Factura"
     'xcombo.AddItem g_tipoguia & "-B.O."
     'xcombo.AddItem g_tipoped & "-Pedido"
     xcombo.ListIndex = 0
     
  End Select
End Sub

Public Sub Main()
   Call Configurar_Conexiones
   Call Cargar_Parametros_Funcionales
   FrmIngreso.Show
   'MDIPrincipal.Show
End Sub

Public Sub Configurar_Conexiones()
  
  On Error GoTo nerror
   ' *****Archivo de Punto de Venta*******
   Dim busca As New dll_apisgen.dll_apis
   Dim strbase As String
   Dim struser As String
   Dim strpass As String
   Dim strserver As String
   
   Dim strconecta As String
   
    VGComputer = UCase$(ComputerName)
    VGsql = busca.LeerIni(App.Path & "\MARFICE.INI", "CONEXION", "SQL", "?")
    VGsql = IIf(VGsql = "?", 0, VGsql)
    
    VGformatofecha = busca.LeerIni(App.Path & "\MARFICE.INI", "CONEXION", "FECHASQL", "?")
    VGformatofecha = IIf(VGformatofecha = "?", "DMY", VGformatofecha)
    
    VGparametros.empresacodigo = "01"

   
   VGparamsistem.bdempresa = busca.LeerIni(App.Path & "\Marfice.ini", "CONEXION", "BDDATOS", "")
   VGparamsistem.Usuario = busca.LeerIni(App.Path & "\Marfice.ini", "CONEXION", "USUARIO", "")
   VGparamsistem.pwd = busca.LeerIni(App.Path & "\Marfice.ini", "CONEXION", "PASSW", "")
   VGparamsistem.servidor = busca.LeerIni(App.Path & "\Marfice.ini", "CONEXION", "SERVIDOR", "")
   strconecta = "Provider=SQLOLEDB.1;Persist Security Info=False;Password=" & VGparamsistem.pwd & ";User ID=" & VGparamsistem.Usuario & ";Initial Catalog=" & VGparamsistem.bdempresa & ";Data Source=" & VGparamsistem.servidor
   
   VGCNx.ConnectionTimeout = 0
   VGCNx.CursorLocation = adUseClient
   VGCNx.ConnectionString = strconecta
   VGCNx.Open
    
    'Conexion de configuracion
   strconecta = "Provider=SQLOLEDB.1;Persist Security Info=False;Password=" & VGparamsistem.pwd & ";User ID=" & VGparamsistem.Usuario & ";Initial Catalog=bdwenco;Data Source=" & VGparamsistem.servidor
   
   vgconfig.ConnectionTimeout = 0
   vgconfig.CursorLocation = adUseClient
   vgconfig.ConnectionString = strconecta
   vgconfig.Open
   
   
   'Base de Datos de Contabilidad
   
   VGparamsistem.bdempresaCT = busca.LeerIni(App.Path & "\Marfice.ini", "CONTABILIDAD", "BDDATOS", "")
   VGparamsistem.usuarioCT = busca.LeerIni(App.Path & "\Marfice.ini", "CONTABILIDAD", "USUARIO", "")
   VGparamsistem.pwdCT = busca.LeerIni(App.Path & "\Marfice.ini", "CONTABILIDAD", "PASSW", "")
   VGparamsistem.servidorCT = busca.LeerIni(App.Path & "\Marfice.ini", "CONTABILIDAD", "SERVIDOR", "")
   strconecta = "Provider=SQLOLEDB.1;Persist Security Info=False;Password=" & VGparamsistem.pwdCT & ";User ID=" & VGparamsistem.usuarioCT & ";Initial Catalog=" & VGparamsistem.bdempresaCT & ";Data Source=" & VGparamsistem.servidorCT
   
   g_BaseContab = strbase
   VGcnxCT.ConnectionTimeout = 0
   VGcnxCT.CursorLocation = adUseClient
   VGcnxCT.ConnectionString = strconecta
   VGcnxCT.Open
    
   'Base de Datos General
   
    VGparamsistem.bdempresaGEN = busca.LeerIni(App.Path & "\MARFICE.INI", "BDGENERAL", "BDDATOS", "?")
    VGparamsistem.servidorGEN = busca.LeerIni(App.Path & "\MARFICE.INI", "BDGENERAL", "SERVIDOR", "?")
    VGparamsistem.usuarioGEN = busca.LeerIni(App.Path & "\MARFICE.INI", "BDGENERAL", "USUARIO", "?")
    VGparamsistem.pwdGEN = busca.LeerIni(App.Path & "\MARFICE.INI", "BDGENERAL", "PASSW", "?")
    strconecta = "Provider=SQLOLEDB.1;Persist Security Info=False;Password=" & VGparamsistem.pwdGEN & ";User ID=" & VGparamsistem.usuarioGEN & ";Initial Catalog=" & VGparamsistem.bdempresaGEN & ";Data Source=" & VGparamsistem.servidorGEN
 
   VGGeneral.CursorLocation = adUseClient
   VGGeneral.ConnectionString = strconecta
   VGGeneral.Open
   
   
' reportes
   
   VGparamsistem.RutaReport = busca.LeerIni(App.Path & "\MARFICE.INI", "REPORTES", "COBRAR", "?")
     VGparamsistem.carpetareportes = busca.LeerIni(App.Path & "\MARFICE.INI", "conexion", "CARPETAREPORTES", "?")
   
   'Establecer Cadena de Conexión de Reportes
    
    VGCadenaReport2 = "DSN=jckconsultores;DSQ=" & VGparamsistem.bdempresaGEN & ";UID=" & VGparamsistem.usuarioGEN & ";PWD=" & VGparamsistem.pwdGEN & ""
        
   
   Exit Sub
  
nerror:
   If Err Then
       MsgBox "Comunicarse con Sistemas " & Chr(13) & Chr(10) & Err.Number & "-" & Err.Description
       Err = 0
       Exit Sub
       Resume
   End If

End Sub

Public Sub Cargar_Parametros_Funcionales()
   Dim rs As New ADODB.Recordset
   Dim rb As New ADODB.Recordset
   Dim averi As New dllgeneral.dll_general
       
    
   g_PedidoPuntoVta = "Tempopedido" & RTrim$(g_ptoventa)
   g_DetallePuntoVta = "Tempodetallepedido" & RTrim$(g_ptoventa)
   
   Set rs = VGCNx.Execute("select top 1 * from vt_parametroventa ")
   If rs.RecordCount > 0 Then
      VGparametros.nombre = Escadena(rs!empresacodigo)
      VGparametros.tienedscto = Escadena(rs!paramvtaestdesc)
      VGparametros.descuento = IIf(IsNull(rs!paramvtadescto), 0, CDbl(rs!paramvtadescto))
      VGparametros.moneda = Escadena(rs!monedacodigo)
      VGparametros.tieneigv = IIf(IsNull(rs!paramvtaestigv) Or rs!paramvtaestigv = 0, "0", "1")
      VGparametros.igv = IIf(IsNull(rs!paramvtaporcigv), 0, CDbl(rs!paramvtaporcigv))
      VGparametros.almacen = Escadena(rs!almacencodigo)
      VGparametros.mensaje = Escadena(rs!paramvtamensaje)
      VGparametros.listapre = IIf(IsNull(rs!paramvtalistaprec), "", Escadena(rs!paramvtalistaprec))
      VGparametros.tipocambio = IIf(IsNull(rs!paramvtatipcambref), CDbl(0), CDbl(rs!paramvtatipcambref))
      VGparametros.comivende = IIf(IsNull(rs!paramvtacomisionvendedor), "", Escadena(rs!paramvtacomisionvendedor))
      VGparametros.formaemi = IIf(IsNull(rs!paramvtaformaemision), "", Escadena(rs!paramvtaformaemision))
      VGparametros.paraboleta = IIf(IsNull(rs!paramvtaboleta) Or rs!paramvtaboleta = 0, "0", "1")
       
   End If
   rs.Close
   Set rb = VGCNx.Execute("select * from vt_puntovtadocumento where puntovtacodigo='" & g_ptoventa & "'")
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
             End If
             rb.MoveNext
          Loop
   End If
   rb.Close
   Set rb = Nothing
   
   Set rs = VGCNx.Execute("select * from vt_puntoventa where puntovtacodigo='" & g_ptoventa & "'")
   If rs.RecordCount > 0 Then
        VGparametros.puntovta = Escadena(rs!puntovtacodigo)
        punto.nropedido = Escadena(IIf(IsNull(rs!puntovtanropedido) Or rs!puntovtanropedido = 0, "0", "1"))
        punto.nrofactura = Escadena(IIf(IsNull(rs!puntovtanrofact) Or rs!puntovtanropedido = 0, "0", "1"))
        punto.nroguia = Escadena(IIf(IsNull(rs!puntovtanroguiarem) Or rs!puntovtanropedido = 0, "0", "1"))
        punto.nroabono = Escadena(IIf(IsNull(rs!puntovtanotaabono) Or rs!puntovtanropedido = 0, "0", "1"))
        punto.nrocargo = Escadena(IIf(IsNull(rs!puntovtanotacargo) Or rs!puntovtanropedido = 0, "0", "1"))
        punto.ventaauto = Escadena(IIf(IsNull(rs!puntovtaautomat) Or rs!puntovtanropedido = 0, "0", "1"))
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
   
   Set rs = VGCNx.Execute("select top 1 * from co_sistema ")
   If ExisteElem(1, VGCNx, "co_sistema", "sistemamultiempresas") Then
       VGparametros.sistemamultiempresas = IIf(averi.ESNULO(rs!sistemamultiempresas, 0) = 0, False, True)
    End If
End Sub

'FIXIT: Declare 'MostrarForm' con un tipo de datos de enlace en tiempo de compilación      FixIT90210ae-R1672-R1B8ZE
Public Function MostrarForm(pVentana As Form, pPos As String)
   pVentana.Icon = LoadPicture(App.Path & "\Cuenta.ico")
   
   If pPos = "C" Then
     pVentana.Left = ((Screen.Width - pVentana.Width) / 2) - 350
     pVentana.Top = ((Screen.Height - pVentana.Height) / 2) - 350
   ElseIf pPos = "I" Then
      pVentana.Left = 300
      pVentana.Top = 300
   ElseIf pPos = "M" And pVentana.Visible = False Then
      pVentana.Caption = pVentana.Caption & "  " & g_DetalleEmpresa
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
   pVentana.Panel.Panels(1).Width = (pVentana.Width / 4)
   If pPos = "M" Then
      pVentana.Panel.Panels(1).Width = ((pVentana.Width - 2600) / 4)
      pVentana.Panel.Panels(1).Text = "EMPRESA: " & g_DetalleEmpresa
      pVentana.Panel.Panels(2).Text = "PTO. VENTA: " & g_ptoventa
      pVentana.Panel.Panels(2).Alignment = sbrLeft
      pVentana.Panel.Panels(2).Width = (pVentana.Width / 4)
   Else
      pVentana.Panel.Panels(1).Text = "FORMATO : " & Escadena(pVentana.Caption)
      pVentana.Panel.Panels(2).Text = "USUARIO: " & g_usuario
      pVentana.Panel.Panels(2).Alignment = sbrLeft
      pVentana.Panel.Panels(2).Width = (pVentana.Width / 4)
   End If
   pVentana.Panel.Panels(1).Alignment = sbrLeft
   pVentana.Panel.Panels(3).Text = "FECHA :" & Format(Date, "dd/mm/yyyy")
   pVentana.Panel.Panels(3).Alignment = sbrRight
   pVentana.Panel.Panels(3).Width = (pVentana.Width / 4)
   pVentana.Panel.Panels(4).Text = "HORA :" & Format(Time, "hh:mm:ss")
   pVentana.Panel.Panels(4).Alignment = sbrRight
   pVentana.Panel.Panels(4).Width = (pVentana.Width / 4)

End Function

'FIXIT: Declare 'Limpiartexto' and 'MBox' and 'Noincluir1' con un tipo de datos de enlace en tiempo de compilación     FixIT90210ae-R1672-R1B8ZE
Public Function Limpiartexto(MBox As Object, ninicio As Integer, nfin As Integer, Optional Noincluir1, Optional Noincluir2 As Integer)
 Dim J As Integer
 If IsMissing(Noincluir1) Then
    Noincluir1 = -1
 End If
 If IsMissing(Noincluir2) Then
    Noincluir2 = -1
 End If
   
   For J = ninicio To nfin
      'Select Case J
         If J = Noincluir1 Or J = Noincluir2 Then
         Else
            MBox(J) = ""
         End If
         'If  Then
         '   MBox(J) = ""
          'End If
   Next J
End Function

'FIXIT: Declare 'Number' con un tipo de datos de enlace en tiempo de compilación           FixIT90210ae-R1672-R1B8ZE
Public Function Numero(Number) As String
   Dim aValor As Double
   If IsNull(Number) Or Len(RTrim$(Number)) = 0 Then
     aValor = 0
   Else
     aValor = Number
   End If
   Numero = RTrim$(Right$(Space(10) & RTrim$(Format(Number, "#######0.00")), 12))
End Function

'FIXIT: Declare 'pDato' con un tipo de datos de enlace en tiempo de compilación            FixIT90210ae-R1672-R1B8ZE
Public Function Escadena(pDato) As String
   If IsNull(pDato) Then
     Escadena = ""
     Exit Function
    ElseIf Len(RTrim$(pDato)) = 0 Then
           Escadena = ""
        Else
         Escadena = RTrim$(pDato)
   End If
End Function

'FIXIT: Declare 'xcampo' con un tipo de datos de enlace en tiempo de compilación           FixIT90210ae-R1672-R1B8ZE
Public Sub PlanillaTotales(xrb As ADODB.Recordset, xcampo, xdepo As Label)
    Dim asumar As Double
    Dim rs As ADODB.Recordset
    asumar = 0
    Set rs = New ADODB.Recordset
    If xrb.RecordCount > 0 Then
        xrb.MoveFirst
        Do Until xrb.EOF
            Set rs = VGCNx.Execute("select tdocumentotipo from cc_tipodocumento where tdocumentocodigo='" & xrb.Fields(2) & "'")
            If rs.BOF Or rs.EOF Then
                MsgBox "Falta completar el tipo (C)argo/(A)bono en el Maestro de Documentos", vbInformation, "Sistema Cuentas x Cobrar"
                Exit Do
            End If
            If rs(0) = "A" Then
               asumar = asumar - CDbl(xrb.Fields(xcampo))
            Else
               asumar = asumar + CDbl(xrb.Fields(xcampo))
            End If
            xrb.MoveNext
        Loop
    End If
    xdepo = Numero(asumar)
    Set rs = Nothing
End Sub

'FIXIT: Declare 'xcampo' con un tipo de datos de enlace en tiempo de compilación           FixIT90210ae-R1672-R1B8ZE
Public Sub PlanillaTotalesCanje(xrb As ADODB.Recordset, xcampo, xdepo As Label)
    Dim asumar As Double
    Dim rs As ADODB.Recordset
    asumar = 0
    Set rs = New ADODB.Recordset
    If xrb.RecordCount > 0 Then
        xrb.MoveFirst
        Do Until xrb.EOF
            Set rs = VGCNx.Execute("select tdocumentotipo from cc_tipodocumento where tdocumentocodigo='" & xrb.Fields(0) & "'")
            If rs.BOF Or rs.EOF Then
                MsgBox "Falta completar el tipo (C)argo/(A)bono en el Maestro de Documentos", vbInformation, "Sistema Cuentas x Cobrar"
                Exit Do
            End If
            If rs(0) = "A" Then
               asumar = asumar - CDbl(xrb.Fields(xcampo))
            Else
               asumar = asumar + CDbl(xrb.Fields(xcampo))
            End If
            xrb.MoveNext
        Loop
    End If
    xdepo = Numero(asumar)
    Set rs = Nothing
End Sub

'FIXIT: Declare 'Seguir' and 'MBox' con un tipo de datos de enlace en tiempo de compilación     FixIT90210ae-R1672-R1B8ZE
Public Function Seguir(MBox As Object, ntecla As Integer)
    If ntecla = 13 Then
        SendKeys "{tab}"
    End If
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
'FIXIT: Declare 'J' con un tipo de datos de enlace en tiempo de compilación                FixIT90210ae-R1672-R1B8ZE
    Dim J, k As Long
    On Error GoTo nerror
    VerificaCombo = -1
    If xcombo.ListCount > 0 Then
      VerificaCombo = 0
      For J = 0 To xcombo.ListCount - 1
         xcombo.ListIndex = J
         k = InStr(xcombo, "-")
         If k > 1 Then
           If Left$(xcombo.Text, k - 1) = ncadena Then
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
   MDIPrincipal.oCrystalReport.ReportFileName = RutaRep & cNombreReporte
   
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
   MDIPrincipal.oCrystalReport.Formulas(0) = "Empresa='" & g_DetalleEmpresa & "'"
   MDIPrincipal.oCrystalReport.Action = 1
  
   Exit Sub
   
Errores:
  MsgBox "Error inesperado: " & Err.Number & "  " & Err.Description, vbExclamation
  Err = 0
  Exit Sub
End Sub

'FIXIT: Declare 'aImpresora' and 'wFile' con un tipo de datos de enlace en tiempo de compilación     FixIT90210ae-R1672-R1B8ZE
Public Function aImpresora(wFile)
'FIXIT: Declare 'wbat' con un tipo de datos de enlace en tiempo de compilación             FixIT90210ae-R1672-R1B8ZE
  Dim wbat, wcade As String
  Dim X As Long
  Dim rnrofile As Double
  
  On Error GoTo nerror
  rnrofile = CInt((90 * (Rnd(10) + 1)))
  wbat = "c:\printer" & CStr(Val(Right$(wFile, 5))) & ".bat"
  Open wbat For Output As #rnrofile
'FIXIT: Print method no tiene equivalente en Visual Basic .NET y no se actualizará.        FixIT90210ae-R7594-R67265
  Print #rnrofile, "@echo off"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
'FIXIT: Print method no tiene equivalente en Visual Basic .NET y no se actualizará.        FixIT90210ae-R7594-R67265
  Print #rnrofile, "Type " & wFile & " >" & Left$(Printer.Port, Len(Printer.Port) - 1)
'FIXIT: Print method no tiene equivalente en Visual Basic .NET y no se actualizará.        FixIT90210ae-R7594-R67265
  Print #rnrofile, "cls"
'FIXIT: Print method no tiene equivalente en Visual Basic .NET y no se actualizará.        FixIT90210ae-R7594-R67265
  Print #rnrofile, "exit"
  Close #rnrofile
  wcade = "start /m " & RTrim$(wbat)
  X = Shell(wcade, vbHide)
  DoEvents
  
nerror:
   If Err Then
      MsgBox "Error: " & Err.Number & "-" & Err.Description, vbCritical, "AVISO"
      Err = 0
   End If

End Function


Public Function DatoTipoCambio(xCn As ADODB.Connection, xfecha As String) As Double
  Dim rs As New ADODB.Recordset
  Dim SQL As String
  SQL = "select tipocambiocompra,tipocambioventa from ct_tipocambio "
  SQL = SQL & "Where tipocambiofecha='" & Format(xfecha, "dd/mm/yyyy") & "'"
  Set rs = xCn.Execute(SQL)
  If Not (rs.EOF Or rs.BOF) Then
     DatoTipoCambio = Format(rs(1), "#####0.###0")
  Else
     DatoTipoCambio = Format(1, "#####0.###0")
  End If
  Set rs = Nothing
End Function

Private Sub CrystOrden(ByRef cry As CrystalReport, cad As String)
Dim pos As Integer, cadaux As String, i As Integer
Dim Valor As String
    Do While True
        pos = InStr(1, cad, ",", vbTextCompare)
        i = 0
        If pos = 0 Then Exit Do
        Valor = Left$(cad, pos - 1)
        cry.SortFields(i) = Valor
        i = i + 1
        cad = Right$(cad, (Len(cad) - pos))
    Loop
End Sub

Public Sub ModoEditable(flagModo As Boolean, Formu As Form, cnameCtrX As String)
 Dim i As Integer
    Dim Control As Control
    For Each Control In Formu.Controls
'FIXIT: Reemplazar la función 'UCase' con la función 'UCase$'.                             FixIT90210ae-R9757-R1B8ZE
'FIXIT: Reemplazar la función 'UCase' con la función 'UCase$'.                             FixIT90210ae-R9757-R1B8ZE
       If UCase$(Control.Name) <> UCase$(cnameCtrX) Then
           If TypeOf Control Is TextBox Then Control.Enabled = flagModo
           If TypeOf Control Is TextFer.TxFer Then Control.Enabled = flagModo
           If TypeOf Control Is CheckBox Then Control.Enabled = flagModo
           If TypeOf Control Is ctrlayuda_f.Ctr_Ayuda Then Control.Enabled = flagModo
       End If
    Next
End Sub

Public Sub LimpiarForm(Formu As Form, cnameCtrX As String)
 Dim i As Integer
    Dim Control As Control
    For Each Control In Formu.Controls
'FIXIT: Reemplazar la función 'UCase' con la función 'UCase$'.                             FixIT90210ae-R9757-R1B8ZE
'FIXIT: Reemplazar la función 'UCase' con la función 'UCase$'.                             FixIT90210ae-R9757-R1B8ZE
       If UCase$(Control.Name) <> UCase$(cnameCtrX) Then
           If TypeOf Control Is TextBox Then Control.Text = Empty
           If TypeOf Control Is TextFer.TxFer Then Control.Text = Empty
'FIXIT: 'Value' no es una propiedad del objeto genérico 'Control' en Visual Basic .NET. Para obtener acceso a 'Value', declare 'Control' utilizando su tipo real en lugar de 'Control'     FixIT90210ae-R1460-RCFE85
           If TypeOf Control Is CheckBox Then Control.Value = 0
           If TypeOf Control Is ctrlayuda_f.Ctr_Ayuda Then
'FIXIT: 'xclave' no es una propiedad del objeto genérico 'Control' en Visual Basic .NET. Para obtener acceso a 'xclave', declare 'Control' utilizando su tipo real en lugar de 'Control'     FixIT90210ae-R1460-RCFE85
              Control.xclave = Empty
'FIXIT: 'xnombre' no es una propiedad del objeto genérico 'Control' en Visual Basic .NET. Para obtener acceso a 'xnombre', declare 'Control' utilizando su tipo real en lugar de 'Control'     FixIT90210ae-R1460-RCFE85
              Control.xnombre = Empty
           End If
       End If
    Next
End Sub

'FIXIT: Declare 'Espunto' and 'texto' con un tipo de datos de enlace en tiempo de compilación     FixIT90210ae-R1672-R1B8ZE
Public Function Espunto(ByRef texto As Variant) As Variant
    If RTrim$(texto) = "." Then
        Espunto = "0"
      Else
        Espunto = texto
    End If
End Function

Public Function DevSaldoVerLimiteCred(ByVal Cliente As String, ByVal Doc As String, ByVal SumaMont _
                As Double, ByVal mon As String, Optional ByRef saldo As Double) As Boolean
Dim cad As String
Dim rs As ADODB.Recordset
Dim monto As Double
  DevSaldoVerLimiteCred = True
  Set rs = New ADODB.Recordset
  monto = 1E+17
  cad = "SELECT A.Clientecodigo,A.codgrup,B.coddoc,SaldoSoles=isnull(A.SaldoSoles,0),SaldoDolares=Isnull(A.SaldoDolares,0) " & _
        "From cc_ClientexGrupoCred A " & _
        "Inner join c_docxlimicred B " & _
        "on A.codgrup=B.codgrup " & _
        "where " & _
        "A.Clientecodigo='" & RTrim$(Cliente) & "' and " & _
        "B.coddoc='" & RTrim$(Doc) & "'"
  rs.Open cad, VGCNx, adOpenKeyset, adLockReadOnly
  If rs.RecordCount > 0 Then monto = IIf(mon = "01", rs!SaldoSoles, rs!SaldoDolares)
  If SumaMont > monto Then
    DevSaldoVerLimiteCred = False
    saldo = monto
  End If
End Function


