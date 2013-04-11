Attribute VB_Name = "MainTesoreria"
Public Cadenabusca As String

'Variables de Ventana de Ayuda
Public nAyuda As String
Public nDetalle As String
Public nAyuda1 As String
Public nMoneda As String

Public VGdllApi As New dll_apisgen.dll_apis
    
Public g_DetalleEmpresa As String
Public g_PedidoPuntoVta As String
Public g_DetallePuntoVta As String

Public Cuentacodigo As String
Public VGactulizodoc As Boolean
Public VGformatofecha As String
Public VGmodifica As Integer

Public VGflaglimpia As Boolean               'Flag que Limpia
Public VGMoverRegistro As Boolean            'Flag al mover el registro
Public VGvarVerifica As Boolean              'Flag Verifica que transaccion es OK (Grabar ,Etc)
Public VGErrorString As String               'Almacena el Error el que hubo en alguna transaccion
Public VGoficina As String * 3
Public VGCommandoSP As ADODB.Command         'De Comando
Public VGvardllgen As dllgeneral.dll_general 'Dll de Algunas funciones


'REPORTES
Public RutaRep  As String      '= "\\desarrollo\librerias_controles\Reportes\"
Public RutaRepProc As String    '= "\\desarrollo\librerias_controles\Reportes\Procesos\"
             'Cadena de Reportes Base Marfice


Public VGParametros As ParametrosdeTesoreria
Public VGParamSistem As ParametrosdeSistema
Public VGtipo As TIPOSISTEMA

Private Type ParametrosdeTesoreria        ' Crea Datos de Empresa
   NomEmpresa As String
   RucEmpresa As String
   auxaut As Boolean
   monedabase As String * 2
   ctascompra As String
   igv As Double
   minimoretencion As Double
   empresacodigo As String * 2
   puntovta As String * 2
   numeracionautomaticalibro As Boolean
   
   sistemaasientoenlinea As Boolean
   descripcion As String * 30
   tipocambio As Double
   controlarefe As String * 1
   numeauto As String * 1
   controlacodigocaja As String * 1
   saldocontadispo As String * 1
   controlacobranzacheq As String * 1
   impresioncheq As String * 1
   listaclientes As String * 1
   listaproveedor As String * 1
   controlacuenta As String * 1
   transferencia As String * 1
   transferenciaegreso As String * 2
   transferenciaingreso As String * 2
   codigooperaciontransferencia As String * 2
   
   sistemaactivaccostos As Boolean
   sistemactrlgastos As Boolean
   sistemamultiempresas As Boolean
   sistemaultimonivel As String * 1
   sistemactaajustedeb As String
   sistemactaajustehab As String
       
    xsubasiento As String
    xLibro As String
    xTipAnal As String
    xCtaIGV As String
    xCtaIES As String
    xCtaRTA As String
    xCtaTotal As String
    xcodretencion As String
    
    CpTiplan As String
    CpOficina As String
    permite_tc As Boolean
    cierremes As Boolean
    AsientoAutoxCCostos As Integer

   porcentajeretencion As Double
   empresacodigoretencion As String * 2
   empresaretencion As String * 1
   sistemaminimoretencion As Double
   sistemanumnivelcosto As String * 1
   sistemabancarizacion As Boolean
   sistemabancarizacion01 As Double
   sistemabancarizacion02 As Double
   controlaestadosrendicion  As Boolean
   diasatrazorendicion As Integer
   diacierrerendicion As Integer
   listacajas As String
End Type

Private Type ParametrosdeSistema
    RutaReport As String
    carpetareportes As String
    AnoProceso As String
    MesProceso As String
    fechatrabajo As Date
    TablaCabcomprob As String
    tabladetcomprob As String
    tipoanaliticocodigo As String
    
    BDEmpresa As String
    Servidor As String
    Usuario As String
    Pwd      As String
    
    BDempresaCONF As String
    UsuarioReporte As String
    
    ServidorCT As String
    BDEmpresaCT As String
    UsuarioCT As String
    PwdCT As String

    ServidorGEN As String
    BDEmpresaGEN As String
    UsuarioGEN As String
    PwdGEN As String

End Type


Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Public Sub Main()
  VGcomputer = UCase(ComputerName)
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
        VGParametros.empresaretencion = numeroEntero(rsaux!empresaretencion)
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

