Attribute VB_Name = "Funciones"
Option Explicit
Public Cadenabusca As String

Public cConexCom As ADODB.Connection   'BdComun

Public cn As New ADODB.Connection
Public cg As New ADODB.Connection
Public cnconta As New ADODB.Connection
Public cbdatos As New ADODB.Connection

'Variables de Ventana de Ayuda
Public nAyuda As String
Public nDetalle As String


'Variables de acceso de usuario
Public Const g_tipoped = "PE"    'PEDIDO
Public Const g_tipofac = "01"    'factura
Public Const g_tipobol = "03"    'boletas
Public Const g_tipoguia = "80"    '9"   'guias B.O
Public Const g_Todos = "99"      'Todos los documentos para Reporte
Public Const g_Eventual = "88888888888" 'cliente eventual

Public g_Empresa As String
Public g_DetalleEmpresa As String
Public g_PedidoPuntoVta As String
Public g_DetallePuntoVta As String

Public g_TipoMovi As String
Public g_usuario As String
Public g_ptoventa As String
Public g_pedserie As String
Public g_facserie As String
Public g_bolserie As String
Public g_guiaserie As String

   Public strbase As String
   Public struser As String
   Public strpass As String
   Public strserver As String


Public Const g_TipoSol = "01"
Public Const g_TipoDolar = "02"

Public punto As PuntoVenta
Public parametro As ParametroEmpresa
Public modoventa As modoventa


'Constantes de mensajes para visualizar
Public Const MsgEdit = "No Existen Datos para Editar.. "
Public Const MsgGraba = "Datos Grabados satisfactoriamente...."
Public Const MsgElim = "No Existen Datos a Eliminar.."
Public Const MsgAdd = "Los datos ya existen...Verifique!!!"
Public Const MsgTitle = "AVISO"
Public Const Msg29 = "Debe Ingresar Numeros"

'REPORTES

'Public Const RutaRep = "\\desarrollo\librerias_controles\Reportes\"
Public RutaRep  As String      '= "\\desarrollo\librerias_controles\Reportes\"
'Public Const RutaRepProc = "\\desarrollo\librerias_controles\Reportes\Procesos\"
Public RutaRepProc As String    '= "\\desarrollo\librerias_controles\Reportes\Procesos\"

'Public Const CadenaRep = "DSN=DESARROLLO;DSQ=Ventas_Prueba;UID=pirata"


Private Type ParametroEmpresa         ' Crea Tipo de Empresa
   nombre As String * 2
   moneda As String * 2
   tieneigv As String * 1
   tienedscto As String * 1
   descuento As Double
   igv As Double
   mensaje As String * 70
   almacen As String * 2
   listapre As String * 1
   tipocambio As Double
   comivende As String * 1
   formaemi As String * 1
   paraboleta As String * 1
End Type


Private Type PuntoVenta   'Crea Punto de Venta
    puntovta As String * 2
    nropedido As String * 1
    nroguia As String * 1
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
End Type


Private Type modoventa   'Crea modoventa
    descuento As String * 1
    impuestos As String * 1
    nroitem  As Double
    numeraauto As String * 1
    ctrlinventario As String * 1
    unidadmedida As String * 1
    copiasfac As Double
    copiasbol As Double
    copiashoja As Double
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
End Type

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


Public Sub CargarTipo(xcombo As ComboBox, xtipo)
  Select Case xtipo
    Case 1     '--condicion documento--
     xcombo.Clear
     xcombo.AddItem "0-Activo"
     xcombo.AddItem "1-Anulado"
     xcombo.ListIndex = 0
   Case 2   '--tipodocumento --
     xcombo.Clear
     xcombo.AddItem "B-Boleta"
     xcombo.AddItem "F-Factura"
     xcombo.AddItem "G-B.O."
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
     xcombo.ListIndex = 0
     
  End Select
End Sub


Public Sub Main()
   

   Call Configurar_Conexiones
 
   'F.Show
   
End Sub

Public Sub Configurar_Conexiones()
  
  On Error GoTo nerror
   ' *****Archivo de Punto de Venta*******
   Dim busca As New dll_apisgen.dll_apis
   'Dim strbase As String
   'Dim struser As String
   'Dim strpass As String
   'Dim strserver As String
   
   Dim strconecta As String
   
   


   
      
   'RutaRep = busca.LeerIni(App.Path & "\Camtex.ini", "Reporte", "repo", "")
   'RutaRepProc = busca.LeerIni(App.Path & "\Camtex.ini", "Reporte", "opera", "")
   
   
   strbase = busca.LeerIni(App.Path & "\Camtex.ini", "Bdatos", "dbase", "")
   struser = busca.LeerIni(App.Path & "\Camtex.ini", "Bdatos", "duser", "")
   strpass = busca.LeerIni(App.Path & "\Camtex.ini", "Bdatos", "dpass", "")
   strserver = busca.LeerIni(App.Path & "\Camtex.ini", "Bdatos", "dserver", "")
    
  strconecta = "Provider=SQLOLEDB.1;Persist Security Info=False;Password=" & strpass & ";User ID=" & struser & ";Initial Catalog=" & strbase & ";Data Source=" & strserver
  ' cbdatos.Close
   cn.ConnectionTimeout = 0
   cn.CursorLocation = adUseClient
   cn.ConnectionString = strconecta
   cn.Open
  ' cbdatos.Open "DSN=dsn_almacen;UID=sistemas;PWD=sistemas"        '"DSN=PC01;DSQ=dsn_datos;UID=marlene;PWD=marlene"
   
    
   'strbase = busca.LeerIni(App.Path & "\Camtex.ini", "Bconta", "dbase", "")
   'struser = busca.LeerIni(App.Path & "\Camtex.ini", "Bconta", "duser", "")
   'strpass = busca.LeerIni(App.Path & "\Camtex.ini", "Bconta", "dpass", "")
   'strserver = busca.LeerIni(App.Path & "\Camtex.ini", "Bconta", "dserver", "")
    
   'strconecta = "Provider=SQLOLEDB.1;Persist Security Info=False;Password=" & strpass & ";User ID=" & struser & ";password=" & strpass & ";Initial Catalog=" & strbase & ";Data Source=" & strserver
    
   'Base de Datos a Conectar
 '  cnconta.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=CONTAPRUEBA;Data Source=DESARROLLO3"
 '  cnconta.ConnectionTimeout = 0
 '  cnconta.CursorLocation = adUseClient
 '  cnconta.Open

   'strbase = busca.LeerIni(App.Path & "\Camtex.ini", "Bgeneral", "dbase", "")
   'struser = busca.LeerIni(App.Path & "\Camtex.ini", "Bgeneral", "duser", "")
   'strpass = busca.LeerIni(App.Path & "\Camtex.ini", "Bgeneral", "dpass", "")
   'strserver = busca.LeerIni(App.Path & "\Camtex.ini", "Bgeneral", "dserver", "")
   
   'strconecta = "Provider=SQLOLEDB.1;Persist Security Info=False;Password=" & strpass & ";User ID=" & struser & ";Initial Catalog=" & strbase & ";Data Source=" & strserver


'   cg.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=pirata;Initial Catalog=MARFICE;Data Source=DESARROLLO"
'   cg.ConnectionTimeout = 0
   'cg.CursorLocation = adUseClient
   'cg.ConnectionString = strconecta
   'cg.Open
   
   'cg.Open "DSN=dsn_general;UID=sistemas;PWD=sistemas"        '"DSN=PC01;DSQ=dsn_datos;UID=marlene;PWD=marlene"

'
'   strbase = busca.LeerIni(App.Path & "\Camtex.ini", "Bventas", "dbase", "")
'   struser = busca.LeerIni(App.Path & "\Camtex.ini", "Bventas", "duser", "")
'   strpass = busca.LeerIni(App.Path & "\Camtex.ini", "Bventas", "dpass", "")
'   strserver = busca.LeerIni(App.Path & "\Camtex.ini", "Bventas", "dserver", "")
'
'   strconecta = "Provider=SQLOLEDB.1;Persist Security Info=False;Password=" & strpass & ";User ID=" & struser & ";Initial Catalog=" & strbase & ";Data Source=" & strserver
'
'
'   cn.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=pirata;Initial Catalog=VENTAS_PRUEBA;Data Source=DESARROLLO"
'   cn.ConnectionTimeout = 0
  ' cn.CursorLocation = adUseClient
  ' cn.ConnectionString = strconecta
  ' cn.Open
    'cn.Open "DSN=dsn_ventas;UID=sistemas;PWD=sistemas"        '"DSN=PC01;DSQ=dsn_datos;UID=marlene;PWD=marlene"


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
       
 '  g_ptoventa = "01"
 '  g_Empresa = "01"
 '  g_usuario = "elozano"
   
   g_PedidoPuntoVta = "Tempopedido" & Trim(g_ptoventa)
   g_DetallePuntoVta = "Tempodetallepedido" & Trim(g_ptoventa)
   
   Set rs = cn.Execute("select * from vt_parametroventa where empresacodigo='" & g_Empresa & "'")
   If rs.RecordCount > 0 Then
      parametro.nombre = Escadena(rs!empresacodigo)
      parametro.tienedscto = Escadena(rs!paramvtaestdesc)
      parametro.descuento = IIf(IsNull(rs!paramvtadescto), 0, CDbl(rs!paramvtadescto))
      parametro.moneda = Escadena(rs!monedacodigo)
      parametro.tieneigv = IIf(IsNull(rs!paramvtaestigv) Or rs!paramvtaestigv = 0, "0", "1")
      parametro.igv = IIf(IsNull(rs!paramvtaporcigv), 0, CDbl(rs!paramvtaporcigv))
      parametro.almacen = Escadena(rs!almacencodigo)
      parametro.mensaje = Escadena(rs!paramvtamensaje)
      parametro.listapre = IIf(IsNull(rs!paramvtalistaprec), "", Escadena(rs!paramvtalistaprec))
      parametro.tipocambio = IIf(IsNull(rs!paramvtatipcambref), CDbl(0), CDbl(rs!paramvtatipcambref))
      parametro.comivende = IIf(IsNull(rs!paramvtacomisionvendedor), "", Escadena(rs!paramvtacomisionvendedor))
      parametro.formaemi = IIf(IsNull(rs!paramvtaformaemision), "", Escadena(rs!paramvtaformaemision))
      parametro.paraboleta = IIf(IsNull(rs!paramvtaboleta) Or rs!paramvtaboleta = 0, "0", "1")
       
   End If
   rs.Close
   Set rb = cn.Execute("select * from vt_puntovtadocumento where puntovtacodigo='" & g_ptoventa & "'")
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
   
   Set rs = cn.Execute("select * from vt_puntoventa where puntovtacodigo='" & g_ptoventa & "'")
   If rs.RecordCount > 0 Then
        punto.puntovta = Escadena(rs!puntovtacodigo)
        punto.nropedido = Escadena(IIf(IsNull(rs!puntovtanropedido) Or rs!puntovtanropedido = 0, "0", "1"))
        punto.nrofactura = Escadena(IIf(IsNull(rs!puntovtanrofact) Or rs!puntovtanropedido = 0, "0", "1"))
        punto.nroguia = Escadena(IIf(IsNull(rs!puntovtanroguiarem) Or rs!puntovtanropedido = 0, "0", "1"))
        punto.nroabono = Escadena(IIf(IsNull(rs!puntovtanotaabono) Or rs!puntovtanropedido = 0, "0", "1"))
        punto.nrocargo = Escadena(IIf(IsNull(rs!puntovtanotacargo) Or rs!puntovtanropedido = 0, "0", "1"))
        punto.ventaauto = Escadena(IIf(IsNull(rs!puntovtaautomat) Or rs!puntovtanropedido = 0, "0", "1"))
   End If
   rs.Close
   Set rs = Nothing
   
   cn.Execute "set dateformat dmy"    '--seteo de formato de fecha
   If averi.VerificaDatoExistente(cn, "select * from sysobjects where name LIKE '" & g_PedidoPuntoVta & "%'") = 0 Then
        cn.Execute "select * into " & g_PedidoPuntoVta & " from vt_pedido"
        cn.Execute "delete from " & g_PedidoPuntoVta
   End If
   If averi.VerificaDatoExistente(cn, "select * from sysobjects where name LIKE '" & g_DetallePuntoVta & "%'") = 0 Then
        cn.Execute "select * into " & g_DetallePuntoVta & " from vt_detallepedido"
        cn.Execute "delete from " & g_DetallePuntoVta
   End If
      
      
   If averi.VerificaDatoExistente(cn, "select * from sysobjects where name Like 'cotizalibre%'") = 0 Then
        cn.Execute "select * into cotizalibre from vt_pedido"
        cn.Execute "delete from cotizalibre"
   End If
   If averi.VerificaDatoExistente(cn, "select * from sysobjects where name Like 'detallecotizalibre%'") = 0 Then
        cn.Execute "select * into detallecotizalibre from vt_detallepedido"
        cn.Execute "delete from detallecotizalibre"
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
'   If pPos = "M" Then
'      pVentana.Panel.Panels(1).Width = ((pVentana.Width - 2600) / 4)
'      pVentana.Panel.Panels(1).Text = "EMPRESA: " & g_DetalleEmpresa
'      pVentana.Panel.Panels(2).Text = "PTO. VENTA: " & g_ptoventa
'      pVentana.Panel.Panels(2).Alignment = sbrLeft
'      pVentana.Panel.Panels(2).Width = (pVentana.Width / 4)
'   Else
'      pVentana.Panel.Panels(1).Text = "FORMATO : " & Escadena(pVentana.Caption)
'      pVentana.Panel.Panels(2).Text = "USUARIO: " & g_usuario
'      pVentana.Panel.Panels(2).Alignment = sbrLeft
'      pVentana.Panel.Panels(2).Width = (pVentana.Width / 4)
'   End If
'   pVentana.Panel.Panels(1).Alignment = sbrLeft
'   pVentana.Panel.Panels(3).Text = "FECHA :" & Format(Date, "dd/mm/yyyy")
'   pVentana.Panel.Panels(3).Alignment = sbrRight
'   pVentana.Panel.Panels(3).Width = (pVentana.Width / 4)
'   pVentana.Panel.Panels(4).Text = "HORA :" & Format(Time, "hh:mm:ss")
'   pVentana.Panel.Panels(4).Alignment = sbrRight
'   pVentana.Panel.Panels(4).Width = (pVentana.Width / 4)

End Function

'Public Function Limpiartexto(MBox As Object, ninicio As Integer, nfin As Integer)
'   Dim J As Integer
'   For J = ninicio To nfin
'       MBox(J) = ""
'   Next J
'End Function


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

Public Function Numero(Number) As String
   Dim aValor As Double
   If IsNull(Number) Or Len(Trim(Number)) = 0 Then
     aValor = 0
   Else
     aValor = Number
   End If
   Numero = Right(Space(10) & Trim(Format(Number, "##,###0.00")), 10)
End Function

Public Function Escadena(pDato) As String
   If IsNull(pDato) Or Len(Trim(pDato)) = 0 Then
     Escadena = ""
   Else
     Escadena = Trim(pDato)
   End If
End Function

Public Function Seguir(MBox As Object, ntecla As Integer)
    If ntecla = 13 Then
        SendKeys "{tab}"
    End If
End Function


Public Function DatoMoneda(xValor As String) As String
   Dim rmone As New ADODB.Recordset
   
   Set rmone = cn.Execute("select * from gr_moneda where monedacodigo='" & xValor & "'")
   If rmone.RecordCount > 0 Then
       DatoMoneda = Escadena(rmone!monedasimbolo) & "/."
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
On Error GoTo Errores
   'MDIPrincipal.oCrystalReport.Destination = crptToWindow
   'MDIPrincipal.oCrystalReport.WindowState = crptMaximized
   'MDIPrincipal.oCrystalReport.ReportFileName = RutaRep & cNombreReporte
   ''MDIPrincipal.oCrystalReport.Connect = CadenaRep
   'MDIPrincipal.oCrystalReport.DiscardSavedData = True
   'MDIPrincipal.oCrystalReport.Action = 1
'   MDIPrincipal.oCrystalReport.Destination = crptToWindow
'   MDIPrincipal.oCrystalReport.WindowState = crptMaximized
'   MDIPrincipal.oCrystalReport.ReportFileName = RutaRep & cNombreReporte
'   MDIPrincipal.oCrystalReport.Formulas(0) = "Empresa='" & g_DetalleEmpresa & "'"
'   MDIPrincipal.oCrystalReport.DiscardSavedData = True
'   MDIPrincipal.oCrystalReport.Action = 1
   
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

