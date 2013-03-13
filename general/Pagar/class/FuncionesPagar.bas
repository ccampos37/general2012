Attribute VB_Name = "FuncionesPagar"


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

    VGCNx.CursorLocation = adUseClient
    VGCNx.CommandTimeout = 0
    VGCNx.ConnectionTimeout = 0
    VGCNx.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & VGparamsistem.Usuario & ";Password=" & VGparamsistem.PWD & ";Initial Catalog=" & VGparamsistem.BDEmpresa & ";Data Source=" & VGparamsistem.Servidor
    VGCNx.Open

   'Conexion de configuracion
   strconecta = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & VGparamsistem.Usuario & ";Password=" & VGparamsistem.PWD & ";Initial Catalog=bdwenco;Data Source=" & VGparamsistem.Servidor
   
   VGConfig.ConnectionTimeout = 0
   VGConfig.CursorLocation = adUseClient
   VGConfig.ConnectionString = strconecta
   VGConfig.Open
       
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
   
   VGGeneral.ConnectionTimeout = 0
   VGGeneral.CursorLocation = adUseClient
   VGGeneral.ConnectionString = strconecta
   VGGeneral.Open
   
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


Public Function Numero_Formato(Number) As String
   Dim aValor As Double
   If IsNull(Number) Or Len(Trim$(Number)) = 0 Then
     aValor = 0
   Else
     aValor = Number
   End If
   Numero_Formato = Right$(Space(10) & Trim$(Format(Number, "##,###0.0000")), 10)
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
Set VGCommandoSP.ActiveConnection = VGGeneral
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

