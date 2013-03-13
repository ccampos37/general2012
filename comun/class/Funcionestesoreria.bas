Attribute VB_Name = "FuncionesTesoreria"
  Option Explicit

Public Sub Main()
  VGComputer = UCase(ComputerName)
   Call Configurar_Conexiones
   Call adicionarcampos
'   Call Cargar_Parametros_Funcionales
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
   
' reportes
   
   VGParamSistem.RutaReport = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "REPORTES", "TESORERIA", "?")
   VGParamSistem.carpetareportes = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "conexion", "CARPETAREPORTES", "?")
   
   VGsql = VGdllApi.LeerIni(App.Path & "\Marfice.ini", "conexion", "SQL", "")
   VGsql = IIf(VGsql = "", 1, VGsql)
   
    VGformatofecha = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "CONEXION", "FECHASQL", "?")
    VGformatofecha = IIf(VGformatofecha = "?", "MDY", VGformatofecha)

   VGParametros.empresacodigo = "01"
   
   RutaRep = VGdllApi.LeerIni(App.Path & "\Marfice.ini", "reportes", "TESORERIA", "")
   RutaRepProc = RutaRep
   strbase = VGdllApi.LeerIni(App.Path & "\Marfice.ini", "conexion", "RUTAREPORTES", "")
   If strbase <> "" Then
      RutaRep = RutaRep + Trim(strbase)
      RutaRepProc = RutaRepProc + Trim(strbase)
   End If
    
' conexion de tesoreria

    VGParamSistem.BDEmpresa = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "conexion", "BDDATOS", "?")
    VGParamSistem.Servidor = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "conexion", "SERVIDOR", "?")
    VGParamSistem.Usuario = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "conexion", "USUARIO", "?")
    VGParamSistem.Pwd = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "conexion", "PASSW", "?")

    Set VGCNx = New ADODB.Connection
    VGCNx.CursorLocation = adUseClient
    VGCNx.CommandTimeout = 0
    VGCNx.ConnectionTimeout = 0
    VGCNx.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & VGParamSistem.Usuario & ";Password=" & VGParamSistem.Pwd & ";Initial Catalog=" & VGParamSistem.BDEmpresa & ";Data Source=" & VGParamSistem.Servidor
    VGCNx.Open
    
    
    ' Bdwenco
    
    Set VGConfig = New ADODB.Connection
    VGConfig.CursorLocation = adUseClient
    VGConfig.CommandTimeout = 0
    VGConfig.ConnectionTimeout = 0
    VGConfig.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & VGParamSistem.Usuario & ";Password=" & VGParamSistem.Pwd & ";Initial Catalog=Bdwenco;Data Source=" & VGParamSistem.Servidor
    VGConfig.Open
    
    
    'Conexion de Contabilidad
    
    VGParamSistem.BDEmpresaCT = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "CONTABILIDAD", "BDDATOS", "?")
    If VGParamSistem.BDEmpresaCT = "" Then
       VGParamSistem.BDEmpresaCT = VGParamSistem.BDEmpresa
       VGParamSistem.ServidorCT = VGParamSistem.Servidor
       VGParamSistem.UsuarioCT = VGParamSistem.Usuario
       VGParamSistem.PwdCT = VGParamSistem.Pwd
     Else
       VGParamSistem.BDEmpresaCT = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "CONTABILIDAD", "BDDATOS", "?")
       VGParamSistem.ServidorCT = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "CONTABILIDAD", "SERVIDOR", "?")
       VGParamSistem.UsuarioCT = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "CONTABILIDAD", "USUARIO", "?")
       VGParamSistem.PwdCT = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "CONTABILIDAD", "PASSW", "?")
   End If
   
    'Conexion de Contabilidad
    Set VGCnxCT = New ADODB.Connection
    VGCnxCT.CursorLocation = adUseClient
    VGCnxCT.CommandTimeout = 0
    VGCnxCT.ConnectionTimeout = 0
    VGCnxCT.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & VGParamSistem.UsuarioCT & ";Password=" & VGParamSistem.PwdCT & ";Initial Catalog=" & VGParamSistem.BDEmpresaCT & ";Data Source=" & VGParamSistem.ServidorCT
    VGCnxCT.Open
    
    VGParamSistem.BDEmpresaGEN = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "BDGENERAL", "BDDATOS", "?")
    VGParamSistem.ServidorGEN = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "BDGENERAL", "SERVIDOR", "?")
    VGParamSistem.UsuarioGEN = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "BDGENERAL", "USUARIO", "?")
    VGParamSistem.PwdGEN = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "BDGENERAL", "PASSW", "?")
        
    'Establecer Conexiones del Marfice
    
    
    Set VGGeneral = New ADODB.Connection
    VGGeneral.CursorLocation = adUseClient
    VGGeneral.CommandTimeout = 0
    VGGeneral.ConnectionTimeout = 0
    VGGeneral.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & VGParamSistem.UsuarioGEN & ";Password=" & VGParamSistem.PwdGEN & ";Initial Catalog=" & VGParamSistem.BDEmpresaGEN & ";Data Source=" & VGParamSistem.Servidor
    VGGeneral.Open
  
      VGCadenaReport2 = "DSN=jckconsultores;DSQ=" & VGParamSistem.BDEmpresaGEN & ";UID=" & VGParamSistem.UsuarioGEN & ";PWD=" & VGParamSistem.PwdGEN & ""

  
   Exit Sub

nerror:
   If Err Then
       MsgBox "Comunicarse con Sistemas " & Chr(13) & Chr(10) & Err.Number & "-" & Err.Description
       Err = 0
       End
   End If
End Sub

Public Sub Cargar_Parametros_Funcionales()
   Dim rsaux As New ADODB.Recordset
   Dim rb As New ADODB.Recordset
   Dim VGvardllgen As New dllgeneral.dll_general
   
   Set rsaux = Nothing
   
   Set rsaux = VGCNx.Execute("select top 1 * from te_parametroempresa ")
   If rsaux.RecordCount > 0 Then
        VGCodEmpresa = Escadena(rsaux!empresacodigo)
        VGParametros.descripcion = Escadena(rsaux!empresarazonsocial)
        VGParametros.tipocambio = numero(rsaux!empresatipocambio)
        VGParametros.controlarefe = Escadena(rsaux!empresacontrolarefe)
        VGParametros.numeauto = Escadena(rsaux!empresanumeauto)
        VGParametros.controlacodigocaja = Escadena(rsaux!empresacontrolacodcaja)
        VGParametros.saldocontadispo = Escadena(rsaux!empresacontrolasaldocontabledispo)
        VGParametros.controlacobranzacheq = Escadena(rsaux!empresanocontrolcobranzacheque)
        VGParametros.impresioncheq = Escadena(rsaux!empresaimpresioncheque)
        VGParametros.controlacuenta = Escadena(rsaux!empresacontrolactacontable)
        VGParametros.transferencia = Escadena(rsaux!empresanumtransferencia)
        VGParametros.transferenciaegreso = Escadena(rsaux!empresatransaccionegreso)
        VGParametros.transferenciaingreso = Escadena(rsaux!empresatransaccioningreso)
        VGParametros.empresacodigoretencion = Escadena(rsaux!empresacodigoretencion)
        VGParametros.porcentajeretencion = numero(rsaux!porcentajeretencion)
        VGParametros.empresaretencion = numero(rsaux!empresaretencion)
        VGParametros.codigooperaciontransferencia = Escadena(rsaux!codigooperaciontransferencia)
   End If
   
   Set rsaux = VGCNx.Execute("select * from co_sistema")
   If rsaux.RecordCount > 0 Then
    VGParametros.monedabase = Trim(rsaux!monedacodigo)
    VGParametros.ctascompra = ArmaCriterioComodin(rsaux!sistemactacomp, "cuentacodigo")
    VGParametros.igv = rsaux!sistemaigv / 100
    
    'Parametros Exclusivos para la generacion de asientos a contabilidad
    
    VGParametros.xLibro = VGvardllgen.ESNULO(rsaux!sistemalibro, "")
    VGParametros.xTipAnal = VGvardllgen.ESNULO(rsaux!sistematipanal, "00")
    VGParametros.xsubasiento = VGvardllgen.ESNULO(rsaux!sistemasubasiento, "00")
    VGParametros.xCtaIGV = VGvardllgen.ESNULO(rsaux!sistemactaIGV, "00")
    VGParametros.xCtaIES = VGvardllgen.ESNULO(rsaux!sistemactaIES, "00")
    VGParametros.xCtaRTA = VGvardllgen.ESNULO(rsaux!sistemactaRTA, "00")
    VGParametros.auxaut = True ' Se tiene que crear el campo para controlar auxiliar automatico
    
    'Cargar parametros para pasar a cuentas por cobrar
    
    VGParametros.CpTiplan = VGvardllgen.ESNULO(rsaux!sistematipoplan, "00")
    VGParametros.CpOficina = VGvardllgen.ESNULO(rsaux!sistemaoficina, "00")
    
    VGParametros.xCtaTotal = rsaux!sistemactatotal
    VGParametros.permite_tc = IIf(VGvardllgen.ESNULO(rsaux!permite_tc, 0) = 0, False, True)
    VGParametros.sistemaactivaccostos = IIf(VGvardllgen.ESNULO(rsaux!sistemaactivaccostos, 0) = 0, False, True)
    VGParametros.sistemaasientoenlinea = IIf(VGvardllgen.ESNULO(rsaux!sistemaasientoenlinea, 0) = 0, False, True)
    VGParametros.sistemactrlgastos = IIf(VGvardllgen.ESNULO(rsaux!sistemactrlgastos, 0) = 0, False, True)
    
    If ExisteElem(1, VGCNx, "co_sistema", "sistemamultiempresas") Then
       VGParametros.sistemamultiempresas = IIf(VGvardllgen.ESNULO(rsaux!sistemamultiempresas, 0) = 0, False, True)
    End If
    VGParametros.minimoretencion = IIf(VGvardllgen.ESNULO(rsaux!sistemaminimoretencion, 0) = 0, 99999, rsaux!sistemaminimoretencion)
    VGParametros.sistemabancarizacion = IIf(VGvardllgen.ESNULO(rsaux!bancarizacion, 0) = 0, 0, rsaux!bancarizacion)
    VGParametros.sistemabancarizacion01 = IIf(VGvardllgen.ESNULO(rsaux!minimobancarizacion01, 0) = 0, 9999999, rsaux!minimobancarizacion01)
    VGParametros.sistemabancarizacion02 = IIf(VGvardllgen.ESNULO(rsaux!minimobancarizacion02, 0) = 0, 9999999, rsaux!minimobancarizacion02)
    
    VGParametros.controlaestadosrendicion = IIf(VGvardllgen.ESNULO(rsaux!controlaestadosrendicion, 0) = 0, 0, rsaux!controlaestadosrendicion)
    VGParametros.diasatrazorendicion = IIf(VGvardllgen.ESNULO(rsaux!diasatrazorendicion, 0) = 0, 0, rsaux!diasatrazorendicion)
    VGParametros.diacierrerendicion = IIf(VGvardllgen.ESNULO(rsaux!diacierrerendicion, 0) = 0, 1, rsaux!diacierrerendicion)

   End If
   Set rsaux = New ADODB.Recordset
   rsaux.Open "select sistemaultimonivel,sistemaultimonivelcostos from  ct_sistema", VGCnxCT, adOpenKeyset, adLockReadOnly
   If rsaux.RecordCount = 0 Then Exit Sub
   VGnumniveles = rsaux!sistemaultimonivel
   VGnumnivcos = ESNULO(rsaux!sistemaultimonivelcostos, 1)
    
    Set rsaux = New ADODB.Recordset
    rsaux.Open "select sistemaultimonivel from  co_sistema", VGCNx, adOpenKeyset, adLockReadOnly
    If rsaux.RecordCount = 0 Then Exit Sub
    VGnumnivgas = rsaux!sistemaultimonivel
    
    Set rsaux = New ADODB.Recordset
    Set rsaux = VGCNx.Execute("select * from  vt_sistema")
    If rsaux.RecordCount = 0 Then Exit Sub
    VGParamSistem.tipoanaliticocodigo = rsaux!tipoanaliticocodigo
    
    Set rsaux = New ADODB.Recordset
    rsaux.Open "select top 1 sistemactaajustedeb,sistemactaajustehab from  ct_sistema", VGCNx, adOpenKeyset, adLockReadOnly
    If rsaux.RecordCount = 0 Then Exit Sub
    
    VGParametros.sistemactaajustedeb = RTrim(rsaux!sistemactaajustedeb)
    VGParametros.sistemactaajustehab = RTrim(rsaux!sistemactaajustehab)
    
    
    VGCNx.Execute "set dateformat dmy"    '--seteo de formato de fecha
    
End Sub


Public Function MontoCero(Number) As String
   Dim aValor As Double
   If IsNull(Number) Or Len(Trim(Number)) = 0 Then
     aValor = 0
   Else
     aValor = Number
   End If
   MontoCero = Right(Space(10) & Trim(Format(aValor, "##,###0.0000")), 10)
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


Public Sub Imprimir(cNombreReporte As String)
Dim VGdllApi As New dll_apisgen.dll_apis
On Error GoTo Errores

   MDIPrincipal.CryRptProc.Reset
   MDIPrincipal.CryRptProc.Destination = crptToWindow
   MDIPrincipal.CryRptProc.WindowState = crptMaximized
   MDIPrincipal.CryRptProc.ReportFileName = RutaRep & cNombreReporte
   
   MDIPrincipal.CryRptProc.LogOnServer "pdssql.dll", _
         VGdllApi.LeerIni(App.Path & "\Marfice.ini", "BDGENERAL", "bddatos", ""), _
         VGdllApi.LeerIni(App.Path & "\Marfice.ini", "BDGENERAL", "servidor", ""), _
         VGdllApi.LeerIni(App.Path & "\Marfice.ini", "BDGENERAL", "usuario", ""), _
         VGdllApi.LeerIni(App.Path & "\Marfice.ini", "BDGENERAL", "passw", "")
   MDIPrincipal.CryRptProc.Connect = _
        "DSN=" & VGdllApi.LeerIni(App.Path & "\Marfice.ini", "BDGENERAL", "servidor", "") & ";" & _
        "DSQ=" & VGdllApi.LeerIni(App.Path & "\Marfice.ini", "BDGENERAL", "bddatos", "") & ";" & _
        "UID=" & VGdllApi.LeerIni(App.Path & "\Marfice.ini", "BDGENERAL", "usuario", "") & ";" & _
        "PWD=" & VGdllApi.LeerIni(App.Path & "\Marfice.ini", "BDGENERAL", "passw", "")
   
   MDIPrincipal.CryRptProc.DiscardSavedData = True
   MDIPrincipal.CryRptProc.formulas(0) = "Empresa='" & g_DetalleEmpresa & "'"
   MDIPrincipal.CryRptProc.Action = 1
  
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
Public Sub PropCrystal(ByRef CrystalRpt As CrystalReport)
    CrystalRpt.WindowShowCancelBtn = True
    CrystalRpt.WindowShowCloseBtn = True
    CrystalRpt.WindowShowExportBtn = True
    CrystalRpt.WindowShowGroupTree = True
    CrystalRpt.WindowShowNavigationCtls = True
    CrystalRpt.WindowShowPrintBtn = True
    CrystalRpt.WindowShowPrintSetupBtn = True
    CrystalRpt.WindowShowProgressCtls = True
    CrystalRpt.WindowShowSearchBtn = True
    CrystalRpt.WindowShowZoomCtl = True
End Sub

Public Sub ImprimirRecibo(Nrecibo As String)
Dim arrform() As Variant, arrparm() As Variant
Dim rs As ADODB.Recordset
Dim xmonto As Double
Dim monto As String
Dim SQL As String
    ReDim arrparm(2)
    ReDim arrform(2)
    '@Base  ,@Nrecibo
    
    Set rs = New ADODB.Recordset
    SQL = "select a.monedacodigo,b.detrec_monedacancela,sum(b.detrec_importesoles) as detrec_importesoles, sum(b.detrec_importedolares) as detrec_importedolares, a.cabrec_tipocambio, a.cabrec_numreciboegreso "
    SQL = SQL & "FROM te_cabecerarecibos a, te_detallerecibos b "
    SQL = SQL & "WHERE a.cabrec_numrecibo=b.cabrec_numrecibo AND "
    SQL = SQL & " b.detalle_no_saldos<>'1' and  "
    SQL = SQL & "a.cabrec_numrecibo='" & Nrecibo & "' "
    SQL = SQL & "Group by a.monedacodigo,b.detrec_monedacancela,a.cabrec_tipocambio,a.cabrec_numreciboegreso"
    Set rs = VGCNx.Execute(SQL)
    If rs.BOF Or rs.EOF Then Exit Sub
    If rs.Fields("monedacodigo") = "01" Then
       If rs.Fields("detrec_monedacancela") = "01" Then
          xmonto = rs.Fields("detrec_importesoles")
       Else
         xmonto = rs.Fields("detrec_importedolares") * rs.Fields("cabrec_tipocambio")
       End If
    Else
       If rs.Fields("detrec_monedacancela") = "01" Then
         xmonto = rs.Fields("detrec_importesoles") / rs.Fields("cabrec_tipocambio")
       Else
         xmonto = rs.Fields("detrec_importedolares")
       End If
    End If
    
    If rs.RecordCount > 0 Then
       monto = Format(xmonto, "#########.00")
       monto = monto + 0.001
       arrparm(0) = VGParamSistem.BDEmpresa
       arrparm(1) = Nrecibo
       arrform(0) = "@NumeroLetras='" & NUMLET(monto) & "'"
       If rs.Fields("cabrec_numreciboegreso") <> Empty Then
          arrform(1) = "@NroTransferencia='" & "Nro Transferencia: " & rs.Fields("cabrec_numreciboegreso") & "'"
       Else
          arrform(1) = "@NroTransferencia='" & rs.Fields("cabrec_numreciboegreso") & "'"
       End If
'       Call ImpresionRptProc("Te_Voucher.rpt", arrform, arrparm, , "Impresion de recibos")
       Call ImpresionRpt_SubRpt_Proc("Te_Voucher.rpt", arrform, arrparm, "Te_Voucher_sub.rpt", , "Impresion de recibos")
    Else
       MsgBox "No existen datos del Nº de Recibo " & Str(Nrecibo)
    End If
    
    rs.Close
    Set rs = Nothing
    
End Sub

Public Sub ImprimirComprobanteretencion(Nrecibo As String)
Dim arrform(1) As Variant, arrparm(4) As Variant
Dim rs As ADODB.Recordset
Dim RSQL As New ADODB.Recordset
Dim xmonto As Double
Dim monto As String
Dim SQL As String
   '@Base  ,@Nrecibo
    
    Set rs = New ADODB.Recordset
    SQL = "select b.detrec_monedacancela,sum(b.detrec_importesoles) as detrec_importesoles, sum(b.detrec_importedolares) as detrec_importedolares, a.cabrec_tipocambio, a.cabrec_numreciboegreso "
    SQL = SQL & "FROM te_cabecerarecibos a, te_detallerecibos b "
    SQL = SQL & "WHERE a.cabrec_numrecibo=b.cabrec_numrecibo AND "
    SQL = SQL & " B.detalle_no_saldos ='1' and "
    SQL = SQL & "a.cabrec_numrecibo='" & Nrecibo & "' "
    SQL = SQL & "Group by b.detrec_monedacancela,a.cabrec_tipocambio,a.cabrec_numreciboegreso"
    Set rs = VGCNx.Execute(SQL)
    If rs.BOF Or rs.EOF Then Exit Sub
 '   Do Until rs.EOF()
    If rs.Fields("detrec_monedacancela") = "01" Then
       xmonto = rs.Fields("detrec_importesoles")
     Else
       xmonto = rs.Fields("detrec_importedolares") * rs.Fields("cabrec_tipocambio")
    End If
    If rs.RecordCount > 0 Then
       monto = Format(xmonto, "#########.00")
       monto = monto + 0.001
       arrparm(0) = VGParamSistem.BDEmpresa
       arrparm(1) = Nrecibo
       SQL = "select top 1 detrec_ndqc from te_detallerecibos where detrec_ndqc>'0' "
       SQL = SQL & " and detrec_tdqc='" & VGParametros.empresacodigoretencion & "'"
       SQL = SQL & " and cabrec_numrecibo='" & Nrecibo & "'"
       Set RSQL = VGCNx.Execute(SQL)
       arrparm(2) = "0000000"
       If RSQL.RecordCount() > 0 Then
          arrparm(2) = Trim(RSQL!detrec_ndqc)
       End If
       arrparm(3) = VGParametros.porcentajeretencion
       arrform(0) = "@NumeroLetras='" & NUMLET(monto) & " Nuevos soles '"
       Call ImpresionRptProc("Te_ComprobanteRetencion.rpt", arrform, arrparm, , "Impresion de Comprobantes de Retencion")
    Else
       MsgBox "No existen datos del Nº de Recibo " & Str(Nrecibo)
    End If
    
    rs.Close
    Set rs = Nothing
    
End Sub
Public Function Espunto(ByRef texto As Variant) As Variant
    If Trim(texto) = "." Then
        Espunto = "0"
      Else
        Espunto = texto
    End If
End Function
Public Sub GeneraAsientoEnlineaTesor(Fecha As Date, empresa As String, m_opcion As String, Nrecibo As String, op As Integer, comprobconta As String, monedacodigo As String, cajabanco As String, m_tipovoucher As String)
Dim rsparimpo As ADODB.Recordset
Dim numerror As Integer
Dim Comando As ADODB.Command
numerror = 0
On Error GoTo Proceso

   VGCNx.BeginTrans

Set rsparimpo = New ADODB.Recordset

rsparimpo.Open "Select * From  ct_importartesoreria Where tipooperacion ='" & UCase(m_opcion) & "' ", VGCnxCT, adOpenKeyset, adLockReadOnly
If rsparimpo.RecordCount() > 0 Then

   Set Comando = New ADODB.Command
   With Comando
        .CommandType = adCmdStoredProc
        .CommandText = "te_GeneraAsientosTesoreriaLinea_pro"
        .CommandTimeout = 0
        .ActiveConnection = VGGeneral
        .Parameters.Refresh
        .Parameters("@BaseConta") = VGCnxCT.DefaultDatabase
        .Parameters("@BaseVenta") = VGCNx.DefaultDatabase
        .Parameters("@empresa") = empresa
        .Parameters("@Asiento") = rsparimpo!asiento
        .Parameters("@SubAsiento") = rsparimpo!SubAsiento
        .Parameters("@Libro") = rsparimpo!libro
         
        .Parameters("@Mes") = Format(Month(Fecha), "00")
        .Parameters("@Ano") = Year(Fecha)
            
        .Parameters("@tipanal") = "002"
        .Parameters("@Compu") = VGComputer
        .Parameters("@Usuario") = VGParamSistem.Usuario
        .Parameters("@TipoMov") = Trim(UCase(m_tipovoucher))
        .Parameters("@Nrecibo") = Nrecibo
        .Parameters("@op") = op
        .Parameters("@comprobconta") = comprobconta
  '      .Parameters("@ajustehaber") = VGParametros.sistemactaajustehab
  '      .Parameters("@ajustedebe") = VGParametros.sistemactaajustedeb
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
Public Sub GeneraAsientoEnlineaTesorTransfer(empresa As String, Fecha As Date, Nrecibo As String)
Dim rsparimpo As ADODB.Recordset
Dim Comando As ADODB.Command
On Error GoTo Procesotransf
    Set rsparimpo = New ADODB.Recordset
    rsparimpo.Open "Select * From  ct_importartesoreria Where Left(Upper(tipooperacion),1) ='T'", VGCnxCT, adOpenKeyset, adLockReadOnly
    Set Comando = New ADODB.Command
        With Comando
            .CommandType = adCmdStoredProc
            .CommandText = "te_GeneraAsientosTesoreriaTransflinea_pro"
            .ActiveConnection = VGGeneral
            .Parameters.Refresh
            .Parameters("@BaseConta") = VGCnxCT.DefaultDatabase
            .Parameters("@BaseVenta") = VGCNx.DefaultDatabase
            .Parameters("@empresa") = empresa
            .Parameters("@Asiento") = rsparimpo!asiento
            .Parameters("@SubAsiento") = rsparimpo!SubAsiento
            .Parameters("@Libro") = rsparimpo!libro
            
            .Parameters("@Mes") = Format(Month(Fecha), "00")
            .Parameters("@Ano") = Year(Fecha)
            
            .Parameters("@Compu") = VGComputer
            .Parameters("@Usuario") = VGParamSistem.Usuario
            .Parameters("@Ntransfer") = Nrecibo
            .Parameters("@ajustehaber") = VGParametros.sistemactaajustehab
            .Parameters("@ajustedebe") = VGParametros.sistemactaajustedeb
            .Execute
        End With
        Screen.MousePointer = 1
        MsgBox "La Contabilizacion se Realizo Satisfactoriamente", vbInformation, "Sistema de Tesoreria"
        Exit Sub
Procesotransf:
        Screen.MousePointer = 1
        MsgBox Err.Description
        Exit Sub
        Resume
End Sub

Public Function ArmaCriterioComodin(cad As String, Campo As String) As String
Dim pos As Integer, cadaux As String, criterio As String
Dim Valor As String
    criterio = ""
    Do While True
        pos = InStr(1, cad, "%", vbTextCompare)
        If pos = 0 Then Exit Do
        Valor = "'" & Left(cad, pos) & "'"
        cad = Right(cad, (Len(cad) - pos))
        criterio = criterio & Campo & " like " & Valor & " or "
    Loop
    ArmaCriterioComodin = Left(criterio, Len(criterio) - 3)
End Function
Public Function FiltroCcosto(codigo As String, ByRef flag As Boolean) As String
Dim rsaux As ADODB.Recordset
Dim AuxCad As String, filtro As String
Dim tipocontrol As String
    Set rsaux = New ADODB.Recordset
    flag = False
    tipocontrol = 1
    If tipocontrol = 0 Then
       filtro = "centrocostocodigo<>'00' and centrocostotipo='6'"
       rsaux.Open "Select criterio=isnull(conceptotextccosto,''),flagx=isnull(conceptosiccosto,0)  From te_conceptocaja Where conceptocodigo='" & Trim(codigo) & "'", VGCNx, adOpenKeyset, adLockReadOnly
       If rsaux.RecordCount > 0 Then
          flag = rsaux!flagx
          If flag Then
             If rsaux!criterio <> "" Then filtro = filtro & " and (" & ArmaCriterioComodin(rsaux!criterio, "centrocostocodigo") & ")"
          End If
       End If
     Else
       filtro = "gastoscodigo<>'00'"
       rsaux.Open "Select criterio=isnull(conceptotextccosto,''),flagx=isnull(conceptosiccosto,0)  From te_conceptocaja Where conceptocodigo='" & Trim(codigo) & "'", VGCNx, adOpenKeyset, adLockReadOnly
       If rsaux.RecordCount > 0 Then
          flag = rsaux!flagx
'          If flag Then
'             If rsaux!criterio <> "" Then Filtro = Filtro & " and (" & ArmaCriterioComodin(rsaux!criterio, "gastoscodigo") & ")"
'          End If
      End If
    End If
    FiltroCcosto = filtro
End Function

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
Public Sub HabilitarDetalle(flag As Boolean, framex As Frame, formx As Form)
'FCP
On Error Resume Next
framex.Enabled = flag
Dim Control As Control
    For Each Control In formx.Controls
        If UCase(Control.Container.Name) = UCase(framex.Name) Then
            Control.Enabled = flag
        End If
    Next
End Sub

Public Sub ClearControlsInframe(framex As Frame, formx As Form)
'FCP
On Error Resume Next
    Dim Control As Control
    For Each Control In formx.Controls
        If UCase(Control.Container.Name) = UCase(framex.Name) Then
            If UCase(Left(Control.Name, 2)) <> "LE" Then
                If TypeOf Control Is TextBox Then Control.Text = ""
                If TypeOf Control Is TextFer.TxFer Then Control.Text = ""
                If TypeOf Control Is Label Then Control.Caption = ""
                'If TypeOf Control Is DTPicker Then Control.Value = Date
            End If
        End If
    Next
End Sub
Public Function ActualizaNumeroAuto(Tabla As String, op As String, cnx As ADODB.Connection, Optional tipo As Integer) As Long
Dim rsaux As ADODB.Recordset
On Error GoTo errornum
    Set rsaux = New ADODB.Recordset
    Select Case op
        Case 1
            If (frmMantRecibos.TxingEgr.Text) = "I" Then
               rsaux.Open "SELECT top 1 Numx=isnull(empresanumeingreso,1) from te_parametroempresa ", cnx, adOpenKeyset, adLockReadOnly
             Else
               rsaux.Open "SELECT top 1 Numx=isnull(empresanumegreso,1) from te_parametroempresa ", cnx, adOpenKeyset, adLockReadOnly
            End If
    End Select
    
    If rsaux.EOF Or rsaux.BOF Then
      ActualizaNumeroAuto = 1
      Exit Function
    Else
      ActualizaNumeroAuto = rsaux!Numx
    End If
    Set rsaux = New ADODB.Recordset
    If op = 1 And tipo = 1 Then
       If tipo = 1 Then
          Set rsaux = cnx.Execute("update te_parametroempresa  set empresanumeingreso = empresanumeingreso + 1")
        Else
          Set rsaux = cnx.Execute("update te_parametroempresa  set empresanumegreso = empresanumegreso + 1")
       End If
    End If
    Exit Function
errornum:
    ActualizaNumeroAuto = -1
End Function
Public Function recibosrendicion(tipo As Integer, rs As Recordset)
Dim SQL As String
Dim xx As String
rs.MoveFirst
If Not ExisteElem(0, VGCNx, VGComputer + "_cajaconcil") Then
   SQL = " Create Table " & VGComputer & "_cajaconcil (recibo VarChar(6))"
 Else
    SQL = " delete " & VGComputer & "_cajaconcil"
End If
VGCNx.Execute (SQL)
If tipo = "1" Then
   Do Until rs.EOF()
     If IsNull(rs!chkconcil) = False And rs!chkconcil Then
           SQL = "insert into " & VGComputer & "_cajaconcil (recibo) values ('" & rs!cabrec_numrecibo & "')"
           VGCNx.Execute (SQL)
      Else
        xx = 12
     End If
     rs.MoveNext
   Loop
End If
If tipo = "2" Then
 
   Do Until rs.EOF()
 '  If rs!cabrec_numrecibo = "213069" Then
 '     xx = 11
 '  End If
    If Not (IsNull(rs!chkconcil) = False And rs!chkconcil) Then
        SQL = "insert into " & VGComputer & "_cajaconcil (recibo) values ('" & rs!cabrec_numrecibo & "')"
        VGCNx.Execute (SQL)
     End If
     rs.MoveNext
   Loop
End If
recibosrendicion = VGComputer + "_cajaconcil"
rs.MoveFirst
End Function
Public Function UltNumeroAuto(Tabla As String, op As String, cnx As ADODB.Connection) As Long
Dim rsaux As ADODB.Recordset
On Error GoTo errornum
    Set rsaux = New ADODB.Recordset
    Select Case op
        Case 1
'            rsaux.Open "SELECT Numx=isnull(IDENT_CURRENT('" & TABLA & "'),0)", cnx, adOpenKeyset, adLockReadOnly
            rsaux.Open "SELECT top 1 Numx=isnull(cabprovinumero,1) from co_sistema ", cnx, adOpenKeyset, adLockReadOnly
    End Select
    If rsaux.EOF Or rsaux.BOF Then
      UltNumeroAuto = 1
      Exit Function
    Else
      UltNumeroAuto = rsaux!Numx
      Set rsaux = New ADODB.Recordset
    End If
    Exit Function
errornum:
    UltNumeroAuto = -1
End Function

Public Function Valida(Modulo As String) As Boolean
Dim Rshabilita As ADODB.Recordset
Valida = True
Set Rshabilita = VGCNx.Execute("select * from ct_cierremensual where empresacodigo='" & FrmContabilizaTesoreria.Ctr_Ayuempresa.xclave & "' " _
& " and mes='" & VGParamSistem.MesProceso & "' and anio='" & VGParamSistem.AnoProceso & "'")
If Rshabilita.RecordCount > 0 Then
    Select Case Modulo
    Case "Facturacion":
        Valida = IIf(Rshabilita!Ventas = False, True, False)
    Case "Cobranza":
        Valida = IIf(Rshabilita!cobrar = False, True, False)
    Case "Pagar":
        Valida = IIf(Rshabilita!pagar = False, True, False)
    Case "Contabilidad":
        Valida = IIf(Rshabilita!contabilidad = False, True, False)
    Case "Provisiones":
        Valida = IIf(Rshabilita!compras = False, True, False)
    End Select
End If

End Function
