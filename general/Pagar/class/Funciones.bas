Attribute VB_Name = "Funciones"
Option Explicit
Public Cadenabusca As String


Public VGvardllgen As dllgeneral.dll_general 'Dll de Algunas funciones
Public VGdllApi As dll_apisgen.dll_apis

'Variables de Ventana de Ayuda
Public nAyuda As String
Public nDetalle As String
Public nSaldo As Double   'Adicional para el Saldo Pendiente
Public VG_String As String

Public VGcomputer As String
Public VGaplicaciones As Double


Public g_Empresa As String
Public g_DetalleEmpresa As String
Public g_PedidoPuntoVta As String
Public g_DetallePuntoVta As String

Public g_TipoMovi As String
Public VGUsuario As String
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

Public Sub Configurar_Conexiones()
  
  On Error GoTo nerror
   ' *****Archivo de Punto de Venta*******
   Dim VGdllApi As New dll_apisgen.dll_apis
   Dim strbase As String
   Dim strpass As String
   Dim strconecta As String
   
   ' reportes
   
   VGparamsistem.RutaReport = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "REPORTES", "PAGAR", "?")
   VGparamsistem.carpetareportes = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "conexion", "CARPETAREPORTES", "?")
   VGcomputer = ComputerName()
   
   VGsql = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "conexion", "SQL", "?")
   VGsql = IIf(VGsql = "?", 0, VGsql)
   
    VGformatofecha = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "CONEXION", "FECHASQL", "?")
    VGformatofecha = IIf(VGformatofecha = "?", "MDY", VGformatofecha)

    'Conexion de pagar
    
    VGparamsistem.BDEmpresa = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "conexion", "BDDATOS", "?")
    VGparamsistem.Servidor = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "conexion", "SERVIDOR", "?")
    VGparamsistem.Usuario = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "conexion", "USUARIO", "?")
    VGparamsistem.PWD = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "conexion", "PASSW", "?")

    Set VGCNx = New ADODB.Connection
    VGCNx.CursorLocation = adUseClient
    VGCNx.CommandTimeout = 0
    VGCNx.ConnectionTimeout = 0
    VGCNx.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & VGparamsistem.Usuario & ";Password=" & VGparamsistem.PWD & ";Initial Catalog=" & VGparamsistem.BDEmpresa & ";Data Source=" & VGparamsistem.Servidor
    VGCNx.Open

   'Conexion de configuracion
   strconecta = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & VGparamsistem.Usuario & ";Password=" & VGparamsistem.PWD & ";Initial Catalog=bdwenco;Data Source=" & VGparamsistem.Servidor
   
   VGconfig.ConnectionTimeout = 0
   VGconfig.CursorLocation = adUseClient
   VGconfig.ConnectionString = strconecta
   VGconfig.Open
       
   'Conexion de Contabilidad
    
    VGparamsistem.BDEmpresaCT = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "CONTABILIDAD", "BDDATOS", "?")
    If VGparamsistem.BDEmpresaCT = "" Then
       VGparamsistem.BDEmpresaCT = VGparamsistem.BDEmpresa
       VGparamsistem.ServidorCT = VGparamsistem.Servidor
       VGparamsistem.UsuarioCT = VGparamsistem.Usuario
       VGparamsistem.PwdCT = VGparamsistem.PWD
     Else
       VGparamsistem.BDEmpresaCT = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "CONTABILIDAD", "BDDATOS", "?")
       VGparamsistem.ServidorCT = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "CONTABILIDAD", "SERVIDOR", "?")
       VGparamsistem.UsuarioCT = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "CONTABILIDAD", "USUARIO", "?")
       VGparamsistem.PwdCT = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "CONTABILIDAD", "PASSW", "?")
   End If
    
   strconecta = "Provider=SQLOLEDB.1;Persist Security Info=False;Password=" & VGparamsistem.PwdCT & ";User ID=" & VGparamsistem.UsuarioCT & ";password=" & VGparamsistem.PwdCT & ";Initial Catalog=" & VGparamsistem.BDEmpresaCT & ";Data Source=" & VGparamsistem.ServidorCT
 
 
   VGcnxCT.ConnectionTimeout = 0
   VGcnxCT.CursorLocation = adUseClient
   VGcnxCT.ConnectionString = strconecta
   VGcnxCT.Open

    'Conexion de General
    VGparamsistem.BDEmpresaGEN = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "BDGENERAL", "BDDATOS", "?")
    VGparamsistem.ServidorGEN = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "BDGENERAL", "SERVIDOR", "?")
    VGparamsistem.UsuarioGEN = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "BDGENERAL", "USUARIO", "?")
    VGparamsistem.PwdGEN = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "BDGENERAL", "PASSW", "?")
    
   strconecta = "Provider=SQLOLEDB.1;Persist Security Info=False;Password=" & VGparamsistem.PwdGEN & ";User ID=" & VGparamsistem.UsuarioGEN & ";Initial Catalog=" & VGparamsistem.BDEmpresaGEN & ";Data Source=" & VGparamsistem.ServidorGEN

'   VGgeneral.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=pirata;Initial Catalog=MARFICE;Data Source=DESARROLLO"
   VGgeneral.ConnectionTimeout = 0
   VGgeneral.CursorLocation = adUseClient
   VGgeneral.ConnectionString = strconecta
   VGgeneral.Open
   
    VGCadenaReport2 = "DSN=jckconsultores;DSQ=" & VGparamsistem.BDEmpresaGEN & ";UID=" & VGparamsistem.UsuarioGEN & ";PWD=" & VGparamsistem.PwdGEN & ""

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
   Set VGvardllgen = New dllgeneral.dll_general
   g_PedidoPuntoVta = "Tempopedido" & Trim$(g_ptoventa)
   g_DetallePuntoVta = "Tempodetallepedido" & Trim$(g_ptoventa)
   
   Set rs = VGCNx.Execute("select top 1 * from vt_parametroventa ")
   If rs.RecordCount > 0 Then
      g_Empresa = Escadena(rs!empresacodigo)
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
    Set rs = VGCNx.Execute("select top 1 * from cp_parametros ")
   If ExisteElem(1, VGCNx, "cp_parametros", "contabilizaenlinea") Then
       VGparametros.contabilizaenlinea = IIf(VGvardllgen.ESNULO(rs!contabilizaenlinea, 0) = 0, False, True)
    End If
   
   Set rs = VGCNx.Execute("select top 1 * from co_sistema ")
   If ExisteElem(1, VGCNx, "co_sistema", "sistemamultiempresas") Then
       VGparametros.sistemamultiempresas = IIf(VGvardllgen.ESNULO(rs!sistemamultiempresas, 0) = 0, False, True)
    End If
   Set rs = VGCNx.Execute("select * from vt_puntoventa where puntovtacodigo='" & g_ptoventa & "'")
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
   If VGparametros.sistemamultiempresas = False Then VGparametros.empresacodigo = "01"
   VGCNx.Execute "set dateformat dmy"    '--seteo de formato de fecha

End Sub

Public Function MostrarForm(pVentana As Form, pPos As String)
   pVentana.Icon = LoadPicture(App.Path & "\Cuenta.ico")
   
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
'   pVentana.Panel.Panels(1).Width = (pVentana.Width / 4)
   If pPos = "M" Then
      pVentana.Panel.Panels(1).Width = ((pVentana.Width - 2600) / 4)
      pVentana.Panel.Panels(1).Text = "EMPRESA: " & g_DetalleEmpresa
      pVentana.Panel.Panels(1).Alignment = sbrLeft
      pVentana.Panel.Panels(2).Text = "Usuario: " & UCase$(VGUsuario)
      pVentana.Panel.Panels(2).Alignment = sbrLeft
      pVentana.Panel.Panels(2).Width = (pVentana.Width / 5)
      pVentana.Panel.Panels(3).Text = "Base: " & UCase$(VGCNx.DefaultDatabase)
      pVentana.Panel.Panels(3).Alignment = sbrLeft
      pVentana.Panel.Panels(3).Width = (pVentana.Width / 5)
      pVentana.Panel.Panels(4).Text = "Servidor: " & UCase$(strserver)
      pVentana.Panel.Panels(4).Alignment = sbrLeft
      pVentana.Panel.Panels(4).Width = (pVentana.Width / 5)
      pVentana.Panel.Panels(5).Text = "Fecha :" & Format(Date, "dd/mm/yyyy")
      pVentana.Panel.Panels(5).Alignment = sbrRight
   Else
      pVentana.Panel.Panels(1).Text = "FORMATO : " & Escadena(pVentana.Caption)
      pVentana.Panel.Panels(1).Width = 3800
      pVentana.Panel.Panels(2).Text = "USUARIO: " & VGUsuario
      pVentana.Panel.Panels(2).Alignment = sbrLeft
      pVentana.Panel.Panels(2).Width = (pVentana.Width / 4)
      pVentana.Panel.Panels(3).Text = "Fecha: " & Format(Date, "dd/mm/yyyy")
      pVentana.Panel.Panels(3).Width = 2200
      pVentana.Panel.Panels(4).Text = "Hora: " & Format(Time, "hh:mm:ss")
      pVentana.Panel.Panels(4).Width = 2200
   End If

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

Public Function numero(Number) As String
   Dim aValor As Double
   If IsNull(Number) Or Len(Trim$(Number)) = 0 Then
     aValor = 0
   Else
     aValor = Number
   End If
   numero = Format(Number, "##,###0.00")
End Function

Public Function Numero_Formato(Number) As String
   Dim aValor As Double
   If IsNull(Number) Or Len(Trim$(Number)) = 0 Then
     aValor = 0
   Else
     aValor = Number
   End If
   Numero_Formato = Right$(Space(10) & Trim$(Format(Number, "##,###0.0000")), 10)
End Function

Public Function Escadena(pdato) As String
   If IsNull(pdato) Or Len(Trim$(pdato)) = 0 Then
     Escadena = ""
   Else
     Escadena = Trim$(pdato)
   End If
End Function


Public Sub PlanillaTotales(xrb As ADODB.Recordset, xcampo, xdepo As Label)
    Dim asumar As Double
    asumar = 0
    If xrb.RecordCount > 0 Then
        xrb.MoveFirst
        Do Until xrb.EOF
            asumar = asumar + CDbl(xrb.Fields(xcampo))
            xrb.MoveNext
        Loop
    End If
    xdepo = numero(asumar)
End Sub

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
With MDIPrincipal.CryRptProc
   .Reset
   .Destination = crptToWindow
   .WindowState = crptMaximized
   .ReportFileName = VGparamsistem.RutaReport & cNombreReporte
   .LogOnServer "pdssql.dll", VGparamsistem.ServidorGEN, VGparamsistem.BDEmpresaGEN, VGparamsistem.UsuarioGEN, ""
   .Connect = VGCadenaReport2
   .DiscardSavedData = True
   .Formulas(0) = "Empresa='" & g_DetalleEmpresa & "'"
   .Action = 1
     Exit Sub
End With
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
  wbat = "c:\printer" & CStr(Val(Right$(wFile, 5))) & ".bat"
  Open wbat For Output As #rnrofile
  Print #rnrofile, "@echo off"
  Print #rnrofile, "Type " & wFile & " >" & Left$(Printer.Port, Len(Printer.Port) - 1)
  Print #rnrofile, "cls"
  Print #rnrofile, "exit"
  Close #rnrofile
  wcade = "start /m " & Trim$(wbat)
  X = Shell(wcade, vbHide)
  DoEvents
  
nerror:
   If Err Then
      MsgBox "Error: " & Err.Number & "-" & Err.Description, vbCritical, "AVISO"
      Err = 0
   End If
End Function

Public Function DatoTipoCambio(xCn As ADODB.Connection, xFec As String) As Double
  Dim rs As New ADODB.Recordset
  Dim SQL As String
  SQL = "select tipocambiocompra,tipocambioventa from ct_tipocambio "
  SQL = SQL & "Where tipocambiofecha='" & Format(xFec, "dd/mm/yyyy") & "'"
  Set rs = xCn.Execute(SQL)
  If Not (rs.EOF Or rs.BOF) Then
     DatoTipoCambio = Format(rs(1), "#####0.###0")
  Else
     DatoTipoCambio = Format(1, "#####0.###0")
  End If
  Set rs = Nothing
End Function
Public Function FechS(Fecha As Variant, tipo As TIPFECHA) As Variant
Dim H As Date
Dim fechaAux As Double
On Error GoTo ErrorFecha
   H = CDate(Fecha)
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
Public Sub contabilizaenlinea(empresa As String, nn As Integer, comprobconta As String, tipo As String, numero As String)
Dim VGCommandoSP As ADODB.Command
Dim rsparimpo As New ADODB.Recordset
On Error GoTo genasiento
Screen.MousePointer = 11
Set rsparimpo = New ADODB.Recordset
Set rsparimpo = VGCNx.Execute("Select * From  ct_importarPagar ")
If rsparimpo.RecordCount() = 0 Then
   Screen.MousePointer = 1
   MsgBox "Verifique el parametro del asiento de Pagar en ct_importarPagar ", vbInformation, "Sistema de Tesoreria"
   Exit Sub
End If

Set VGCommandoSP = New ADODB.Command
Set VGCommandoSP.ActiveConnection = VGgeneral
    VGCommandoSP.CommandText = "cp_GeneraAsientoPagarenLinea_pro"
    VGCommandoSP.CommandType = adCmdStoredProc
    VGCommandoSP.Prepared = True
    VGCommandoSP.CommandTimeout = 0
    VGCommandoSP.Parameters.Refresh
 
    With VGCommandoSP
         .Parameters("@BaseConta") = VGcnxCT.DefaultDatabase
         .Parameters("@BaseVenta") = VGCNx.DefaultDatabase
         .Parameters("@empresa") = empresa
         .Parameters("@Asiento") = rsparimpo!asientocodigo
         .Parameters("@SubAsiento") = rsparimpo!subasientocodigo
         .Parameters("@Libro") = rsparimpo!librocodigo
         .Parameters("@Mes") = Format(VGparamsistem.Mesproceso, "00")
         .Parameters("@Ano") = VGparamsistem.Anoproceso
         .Parameters("@tipanal") = "001"
         .Parameters("@compu") = VGcomputer
         .Parameters("@usuario") = VGUsuario
         .Parameters("@ajustedebe") = "771100"
         .Parameters("@ajustehaber") = "671100"
         .Parameters("@numero") = tipo + numero
         .Parameters("@op") = 1
         .Parameters("@comprobconta") = comprobconta
         .Execute
     End With
Set rsparimpo = Nothing
Exit Sub
genasiento:
    Screen.MousePointer = 1
    VGCNx.RollbackTrans
    MsgBox "Hubo Errores al momento que se genero la Planilla  " & numero & Chr(13) & Err.Description
    Exit Sub
    Resume Next
End Sub
Public Sub Main()
   Call Configurar_Conexiones
   Call adicionarcampos
   FrmIngreso.Show
   'MDIPrincipal.Show
End Sub

