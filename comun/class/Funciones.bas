Attribute VB_Name = "FuncionesCobrar"
Option Explicit

Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

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
   
    VGComputer = UCase(ComputerName)
    VGsql = busca.LeerIni(App.Path & "\MARFICE.INI", "CONEXION", "SQL", "?")
    VGsql = IIf(VGsql = "?", 0, VGsql)
    
    VGformatofecha = busca.LeerIni(App.Path & "\MARFICE.INI", "CONEXION", "FECHASQL", "?")
    VGformatofecha = IIf(VGformatofecha = "?", "DMY", VGformatofecha)
    
   
   VGParamSistem.BDEmpresa = busca.LeerIni(App.Path & "\Marfice.ini", "CONEXION", "BDDATOS", "")
   VGParamSistem.Usuario = busca.LeerIni(App.Path & "\Marfice.ini", "CONEXION", "USUARIO", "")
   VGParamSistem.Pwd = busca.LeerIni(App.Path & "\Marfice.ini", "CONEXION", "PASSW", "")
   VGParamSistem.Servidor = busca.LeerIni(App.Path & "\Marfice.ini", "CONEXION", "SERVIDOR", "")
   strconecta = "Provider=SQLOLEDB.1;Persist Security Info=False;Password=" & VGParamSistem.Pwd & ";User ID=" & VGParamSistem.Usuario & ";Initial Catalog=" & VGParamSistem.BDEmpresa & ";Data Source=" & VGParamSistem.Servidor
   
   VGCNx.ConnectionTimeout = 0
   VGCNx.CursorLocation = adUseClient
   VGCNx.ConnectionString = strconecta
   VGCNx.Open
    
    'Conexion de configuracion
   strconecta = "Provider=SQLOLEDB.1;Persist Security Info=False;Password=" & VGParamSistem.Pwd & ";User ID=" & VGParamSistem.Usuario & ";Initial Catalog=bdwenco;Data Source=" & VGParamSistem.Servidor
   
   VGconfig.ConnectionTimeout = 0
   VGconfig.CursorLocation = adUseClient
   VGconfig.ConnectionString = strconecta
   VGconfig.Open
   
   
   'Base de Datos de Contabilidad
   
   VGParamSistem.BDEmpresaCT = busca.LeerIni(App.Path & "\Marfice.ini", "CONTABILIDAD", "BDDATOS", "")
   VGParamSistem.UsuarioCT = busca.LeerIni(App.Path & "\Marfice.ini", "CONTABILIDAD", "USUARIO", "")
   VGParamSistem.PwdCT = busca.LeerIni(App.Path & "\Marfice.ini", "CONTABILIDAD", "PASSW", "")
   VGParamSistem.ServidorCT = busca.LeerIni(App.Path & "\Marfice.ini", "CONTABILIDAD", "SERVIDOR", "")
   strconecta = "Provider=SQLOLEDB.1;Persist Security Info=False;Password=" & VGParamSistem.PwdCT & ";User ID=" & VGParamSistem.UsuarioCT & ";Initial Catalog=" & VGParamSistem.BDEmpresaCT & ";Data Source=" & VGParamSistem.ServidorCT
   
   g_BaseContab = strbase
   VGCnxCT.ConnectionTimeout = 0
   VGCnxCT.CursorLocation = adUseClient
   VGCnxCT.ConnectionString = strconecta
   VGCnxCT.Open
    
   'Base de Datos General
   
    VGParamSistem.BDEmpresaGEN = busca.LeerIni(App.Path & "\MARFICE.INI", "BDGENERAL", "BDDATOS", "?")
    VGParamSistem.ServidorGEN = busca.LeerIni(App.Path & "\MARFICE.INI", "BDGENERAL", "SERVIDOR", "?")
    VGParamSistem.UsuarioGEN = busca.LeerIni(App.Path & "\MARFICE.INI", "BDGENERAL", "USUARIO", "?")
    VGParamSistem.PwdGEN = busca.LeerIni(App.Path & "\MARFICE.INI", "BDGENERAL", "PASSW", "?")
    strconecta = "Provider=SQLOLEDB.1;Persist Security Info=False;Password=" & VGParamSistem.PwdGEN & ";User ID=" & VGParamSistem.UsuarioGEN & ";Initial Catalog=" & VGParamSistem.BDEmpresaGEN & ";Data Source=" & VGParamSistem.ServidorGEN
 
   VGgeneral.CursorLocation = adUseClient
   VGgeneral.ConnectionString = strconecta
   VGgeneral.Open
   
   
' reportes
   
   VGParamSistem.RutaReport = busca.LeerIni(App.Path & "\MARFICE.INI", "REPORTES", "COBRAR", "?")
    VGParamSistem.carpetareportes = busca.LeerIni(App.Path & "\MARFICE.INI", "conexion", "CARPETAREPORTES", "?")
   
   'Establecer Cadena de Conexión de Reportes
    
    VGCadenaReport2 = "DSN=jckconsultores;DSQ=" & VGParamSistem.BDEmpresaGEN & ";UID=" & VGParamSistem.UsuarioGEN & ";PWD=" & VGParamSistem.PwdGEN & ""
        
   
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
       
    
   g_PedidoPuntoVta = "Tempopedido" & Trim(g_ptoventa)
   g_DetallePuntoVta = "Tempodetallepedido" & Trim(g_ptoventa)
   
   Set rs = VGCNx.Execute("select * from vt_parametroventa ")
   If rs.RecordCount > 0 Then
      VGParametros.Nombre = Escadena(rs!empresacodigo)
      VGParametros.tienedscto = Escadena(rs!paramvtaestdesc)
      VGParametros.descuento = IIf(IsNull(rs!paramvtadescto), 0, CDbl(rs!paramvtadescto))
      VGParametros.moneda = Escadena(rs!monedacodigo)
      VGParametros.tieneigv = IIf(IsNull(rs!paramvtaestigv) Or rs!paramvtaestigv = 0, "0", "1")
       VGParametros.Igv = IIf(IsNull(rs!paramvtaporcigv), 0, CDbl(rs!paramvtaporcigv))
      VGParametros.almacen = Escadena(rs!almacencodigo)
      VGParametros.Mensaje = Escadena(rs!paramvtamensaje)
      VGParametros.listapre = IIf(IsNull(rs!paramvtalistaprec), "", Escadena(rs!paramvtalistaprec))
      VGParametros.tipocambio = IIf(IsNull(rs!paramvtatipcambref), CDbl(0), CDbl(rs!paramvtatipcambref))
      VGParametros.comivende = IIf(IsNull(rs!paramvtacomisionvendedor), "", Escadena(rs!paramvtacomisionvendedor))
      VGParametros.formaemi = IIf(IsNull(rs!paramvtaformaemision), "", Escadena(rs!paramvtaformaemision))
      VGParametros.paraboleta = IIf(IsNull(rs!paramvtaboleta) Or rs!paramvtaboleta = 0, "0", "1")
       
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
        VGParamSistem.puntovta = Escadena(rs!puntovtacodigo)
        VGParamSistem.nropedido = Escadena(IIf(IsNull(rs!puntovtanropedido) Or rs!puntovtanropedido = 0, "0", "1"))
        VGParamSistem.nrofactura = Escadena(IIf(IsNull(rs!puntovtanrofact) Or rs!puntovtanropedido = 0, "0", "1"))
        VGParamSistem.nroguia = Escadena(IIf(IsNull(rs!puntovtanroguiarem) Or rs!puntovtanropedido = 0, "0", "1"))
        VGParamSistem.nroabono = Escadena(IIf(IsNull(rs!puntovtanotaabono) Or rs!puntovtanropedido = 0, "0", "1"))
        VGParamSistem.nrocargo = Escadena(IIf(IsNull(rs!puntovtanotacargo) Or rs!puntovtanropedido = 0, "0", "1"))
        VGParamSistem.ventaauto = Escadena(IIf(IsNull(rs!puntovtaautomat) Or rs!puntovtanropedido = 0, "0", "1"))
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
       VGParametros.sistemamultiempresas = IIf(averi.ESNULO(rs!sistemamultiempresas, 0) = 0, False, True)
    End If
    If VGParametros.sistemamultiempresas = False Then
       VGParametros.empresacodigo = "01"
    End If
   Set rs = VGCNx.Execute("select top 1 * from cc_sistema ")
   If ExisteElem(1, VGCNx, "cc_sistema", "imprimevoucher") Then
       VGParametros.imprimevoucher = ESNULO(rs!imprimevoucher, 0)
    End If
    
End Sub

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
    xdepo = numero(asumar)
    Set rs = Nothing
End Sub

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
    xdepo = numero(asumar)
    Set rs = Nothing
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


Public Sub imprimir(cNombreReporte As String)
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
   MDIPrincipal.oCrystalReport.formulas(0) = "Empresa='" & g_DetalleEmpresa & "'"
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
Dim pos As Integer, cadaux As String, I As Integer
Dim valor As String
    Do While True
        pos = InStr(1, cad, ",", vbTextCompare)
        I = 0
        If pos = 0 Then Exit Do
        valor = Left(cad, pos - 1)
        cry.SortFields(I) = valor
        I = I + 1
        cad = Right(cad, (Len(cad) - pos))
    Loop
End Sub

Public Sub ModoEditable(flagModo As Boolean, Formu As Form, cnameCtrX As String)
 Dim I As Integer
    Dim Control As Control
    For Each Control In Formu.Controls
       If UCase(Control.Name) <> UCase(cnameCtrX) Then
           If TypeOf Control Is TextBox Then Control.Enabled = flagModo
           If TypeOf Control Is TextFer.TxFer Then Control.Enabled = flagModo
           If TypeOf Control Is CheckBox Then Control.Enabled = flagModo
           If TypeOf Control Is ctrlayuda_f.Ctr_Ayuda Then Control.Enabled = flagModo
       End If
    Next
End Sub

Public Sub LimpiarForm(Formu As Form, cnameCtrX As String)
 Dim I As Integer
    Dim Control As Control
    For Each Control In Formu.Controls
       If UCase(Control.Name) <> UCase(cnameCtrX) Then
           If TypeOf Control Is TextBox Then Control.text = Empty
           If TypeOf Control Is TextFer.TxFer Then Control.text = Empty
           If TypeOf Control Is CheckBox Then Control.Value = 0
           If TypeOf Control Is ctrlayuda_f.Ctr_Ayuda Then
              Control.xclave = Empty
              Control.xnombre = Empty
           End If
       End If
    Next
End Sub

Public Function Espunto(ByRef texto As Variant) As Variant
    If Trim(texto) = "." Then
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
        "A.Clientecodigo='" & Trim(Cliente) & "' and " & _
        "B.coddoc='" & Trim(Doc) & "'"
  rs.Open cad, VGCNx, adOpenKeyset, adLockReadOnly
  If rs.RecordCount > 0 Then monto = IIf(mon = "01", rs!SaldoSoles, rs!SaldoDolares)
  If SumaMont > monto Then
    DevSaldoVerLimiteCred = False
    saldo = monto
  End If
End Function


