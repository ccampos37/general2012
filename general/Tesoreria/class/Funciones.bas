Attribute VB_Name = "Funciones"
Option Explicit
Public Cadenabusca As String

Public VGcnx As New ADODB.Connection
Public cg As New ADODB.Connection
Public cnconta As New ADODB.Connection

'Variables de Ventana de Ayuda
Public nAyuda As String
Public nDetalle As String
Public nAyuda1 As String
Public nMoneda As String

Public g_Empresa As String
Public g_DetalleEmpresa As String
Public g_PedidoPuntoVta As String
Public g_DetallePuntoVta As String

Public VGfecha As Date
Public g_TipoMovi As String
Public g_usuario As String
Public g_ptoventa As String
Public g_pedserie As String
Public g_facserie As String
Public g_bolserie As String
Public g_guiaserie As String
Public VGComputer As String   'Nombre de la computadora


Public Const g_TipoSol = "01"
Public Const g_TipoDolar = "02"

'Public parametro As ParametroEmpresa
Public Empresa As DatoEmpresa
Public VGParamSistem As ParametrosdeSistema

'Variables de acceso de usuario
Public Const g_Eventual = "88888888888" 'cliente eventual

'Constantes de mensajes para visualizar
Public Const MsgEdit = "No Existen Datos para Editar.. "
Public Const MsgGraba = "Datos Grabados satisfactoriamente...."
Public Const MsgElim = "No Existen Datos a Eliminar.."
Public Const MsgAdd = "Los datos ya existen...Verifique!!!"
Public Const MsgTitle = "AVISO"
Public Const Msg29 = "Debe Ingresar Numeros"

'REPORTES
Public RutaRep  As String      '= "\\desarrollo\librerias_controles\Reportes\"
Public RutaRepProc As String    '= "\\desarrollo\librerias_controles\Reportes\Procesos\"

'Private Type ParametroEmpresa         ' Crea Tipo de Empresa
'   nombre As String * 2
'   descripcion As String * 50
'   moneda As String * 2
'   tieneigv As String * 1
'   tienedscto As String * 1
'   descuento As Double
'   igv As Double
'   mensaje As String * 70
'   almacen As String * 2
'   listapre As String * 1
'   tipocambio As Double
'   comivende As String * 1
'   formaemi As String * 1
'   paraboleta As String * 1
'End Type

Private Type DatoEmpresa         ' Crea Datos de Empresa
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
   sistemaactivaccostos As Boolean
   sistemactrlgastos As Boolean
   sistemamultiempresas As Boolean
   sistemaultimonivel As String * 1
   
End Type

Private Type ParametrosdeSistema
    BDEmpresa As String
    BDConfigura As String
    Servidor As String
    RutaReport As String
    Usuario As String
    Pwd      As String
    carpetareportes As String
    Año As String
    Mes As String
End Type
Public Enum TIPFECHA
   Sqlf = 1
   Adof = 2
End Enum

Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Property Get ComputerName() As Variant
    Dim sName As String
    Dim iRetVal As Long
    Dim ipos As Integer
    
    sName = Space$(255)
    iRetVal = GetComputerName(sName, 255&)
    If iRetVal = 0 Then
      ComputerName = ""
      Exit Property
    End If
    ipos = InStr(sName, Chr$(0))
    ComputerName = Left$(sName, ipos - 1)
End Property

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
  Dim adll2 As New dllgeneral.dll_general
  
  Select Case xtipo
    Case 1     '--condicion documento--
     xcombo.Clear
     xcombo.AddItem "0-Activo"
     xcombo.AddItem "1-Anulado"
     xcombo.ListIndex = 0
   Case 2   '--tipodocumento --
     xcombo.Clear
     Call adll2.llenacombo(xcombo, "select documentocodigo,documentodescripcion from vt_documento", VGcnx)
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
     Call adll2.llenacombo(xcombo, "select documentocodigo,documentodescripcion from vt_documento ", VGcnx)
     'xcombo.AddItem g_tipobol & "-Boleta"
     'xcombo.AddItem g_tipofac & "-Factura"
     'xcombo.AddItem g_tipoguia & "-B.O."
     'xcombo.AddItem g_tipoped & "-Pedido"
     xcombo.ListIndex = 0
     
  End Select
End Sub

Public Sub Main()
  VGComputer = "##" + UCase(ComputerName)
   Call Configurar_Conexiones
   Call Cargar_Parametros_Funcionales
   FrmIngreso.Show
  
End Sub

Public Sub Configurar_Conexiones()
  
  On Error GoTo nerror
   ' *****Archivo de Punto de Venta*******
   Dim Busca As New dll_apisgen.dll_apis
   Dim strbase As String
   Dim struser As String
   Dim strpass As String
   Dim strserver As String
   
   Dim strconecta As String
   
   RutaRep = Busca.LeerIni(App.Path & "\Marfice.ini", "Reporte", "repo", "")
   RutaRepProc = Busca.LeerIni(App.Path & "\Marfice.ini", "Reporte", "opera", "")
    
   strbase = Busca.LeerIni(App.Path & "\Marfice.ini", "CONTABILIDAD", "bddatos", "")
   struser = Busca.LeerIni(App.Path & "\Marfice.ini", "CONTABILIDAD", "usuario", "")
   strpass = Busca.LeerIni(App.Path & "\Marfice.ini", "CONTABILIDAD", "passw", "")
   strserver = Busca.LeerIni(App.Path & "\Marfice.ini", "CONTABILIDAD", "servidor", "")
   strconecta = "Provider=SQLOLEDB.1;Persist Security Info=False;Password=" & strpass & ";User ID=" & struser & ";password=" & strpass & ";Initial Catalog=" & strbase & ";Data Source=" & strserver
   
   cnconta.CursorLocation = adUseClient
   cnconta.ConnectionString = strconecta
   cnconta.CommandTimeout = 0
   cnconta.ConnectionTimeout = 0
   cnconta.Open

   strbase = Busca.LeerIni(App.Path & "\Marfice.ini", "BDGENERAL", "bddatos", "")
   struser = Busca.LeerIni(App.Path & "\Marfice.ini", "BDGENERAL", "usuario", "")
   strpass = Busca.LeerIni(App.Path & "\Marfice.ini", "BDGENERAL", "passw", "")
   strserver = Busca.LeerIni(App.Path & "\Marfice.ini", "BDGENERAL", "servidor", "")
   strconecta = "Provider=SQLOLEDB.1;Persist Security Info=False;Password=" & strpass & ";User ID=" & struser & ";Initial Catalog=" & strbase & ";Data Source=" & strserver
   
   cg.CursorLocation = adUseClient
   cg.ConnectionString = strconecta
   cg.CommandTimeout = 0
   cg.ConnectionTimeout = 0
   cg.Open
   VGParamSistem.BDConfigura = strbase
   
   strbase = Busca.LeerIni(App.Path & "\Marfice.ini", "CONEXION", "bddatos", "")
   struser = Busca.LeerIni(App.Path & "\Marfice.ini", "CONEXION", "usuario", "")
   strpass = Busca.LeerIni(App.Path & "\Marfice.ini", "CONEXION", "passw", "")
   strserver = Busca.LeerIni(App.Path & "\Marfice.ini", "CONEXION", "servidor", "")
   strconecta = "Provider=SQLOLEDB.1;Persist Security Info=False;Password=" & strpass & ";User ID=" & struser & ";Initial Catalog=" & strbase & ";Data Source=" & strserver

   VGcnx.CursorLocation = adUseClient
   VGcnx.ConnectionString = strconecta
   VGcnx.Open
   
   VGParamSistem.BDEmpresa = strbase
   VGParamSistem.Servidor = strserver
   VGParamSistem.Usuario = struser
   VGParamSistem.Pwd = strpass
   
   VGParamSistem.Año = "2002"
   VGParamSistem.Mes = "12"
   
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
   Dim VGvardllgen As New dllgeneral.dll_general
   
   Set rs = VGcnx.Execute("select * from vt_parametroventa where empresacodigo='" & g_Empresa & "'")
   If rs.RecordCount > 0 Then
 '     parametro.nombre = Escadena(rs!empresacodigo)
 '     parametro.tienedscto = Escadena(rs!paramvtaestdesc)
 '     parametro.descuento = IIf(IsNull(rs!paramvtadescto), 0, CDbl(rs!paramvtadescto))
 '     parametro.moneda = Escadena(rs!monedacodigo)
 '     parametro.tieneigv = IIf(IsNull(rs!paramvtaestigv) Or rs!paramvtaestigv = 0, "0", "1")
 '     parametro.igv = IIf(IsNull(rs!paramvtaporcigv), 0, CDbl(rs!paramvtaporcigv))
 '     parametro.almacen = Escadena(rs!almacencodigo)
 '     parametro.mensaje = Escadena(rs!paramvtamensaje)
 '     parametro.listapre = IIf(IsNull(rs!paramvtalistaprec), "", Escadena(rs!paramvtalistaprec))
 '     parametro.tipocambio = IIf(IsNull(rs!paramvtatipcambref), CDbl(0), CDbl(rs!paramvtatipcambref))
 '     parametro.comivende = IIf(IsNull(rs!paramvtacomisionvendedor), "", Escadena(rs!paramvtacomisionvendedor))
 '     parametro.formaemi = IIf(IsNull(rs!paramvtaformaemision), "", Escadena(rs!paramvtaformaemision))
 '     parametro.paraboleta = IIf(IsNull(rs!paramvtaboleta) Or rs!paramvtaboleta = 0, "0", "1")
   End If
   rs.Close
   Set rs = Nothing
   
   Set rs = VGcnx.Execute("select * from te_parametroempresa where empresacodigo='" & g_Empresa & "'")
   If rs.RecordCount > 0 Then
        Empresa.descripcion = Escadena(rs!empresarazonsocial)
        Empresa.tipocambio = Numero(rs!empresatipocambio)
        Empresa.controlarefe = Escadena(rs!empresacontrolarefe)
        Empresa.numeauto = Escadena(rs!empresanumeauto)
        Empresa.controlacodigocaja = Escadena(rs!empresacontrolacodcaja)
        Empresa.saldocontadispo = Escadena(rs!empresacontrolasaldocontabledispo)
        Empresa.controlacobranzacheq = Escadena(rs!empresanocontrolcobranzacheque)
        Empresa.impresioncheq = Escadena(rs!empresaimpresioncheque)
       ' Empresa.listaclientes = Escadena(rs!empresalistaestadoclientes)
       ' Empresa.listaproveedor = Escadena(rs!empresalistaestadoproveedor)
        Empresa.controlacuenta = Escadena(rs!empresacontrolactacontable)
        Empresa.transferencia = Escadena(rs!empresanumtransferencia)
        Empresa.transferenciaegreso = Escadena(rs!empresatransaccionegreso)
        Empresa.transferenciaingreso = Escadena(rs!empresatransaccioningreso)
   End If
   rs.Close
   Set rs = Nothing
   
   Set rs = VGcnx.Execute("select * from co_sistema")
   If rs.RecordCount > 0 Then
    Empresa.sistemaactivaccostos = IIf(VGvardllgen.ESNULO(rs!sistemaactivaccostos, 0) = 0, False, True)
    Empresa.sistemactrlgastos = IIf(VGvardllgen.ESNULO(rs!sistemactrlgastos, 0) = 0, False, True)
    If ExisteElem(1, VGcnx, "co_sistema", "sistemamultiempresas") Then
       Empresa.sistemamultiempresas = IIf(VGvardllgen.ESNULO(rs!sistemamultiempresas, 0) = 0, False, True)
    End If
    Empresa.sistemaultimonivel = Escadena(rs!sistemaultimonivel)
   End If
   rs.Close
   Set rs = Nothing
   
   
   
   VGcnx.Execute "set dateformat dmy"    '--seteo de formato de fecha

End Sub

Public Function MostrarForm(pVentana As Form, pPos As String)
   pVentana.Icon = LoadPicture(App.Path & "\tesoro.ico")
   
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
      pVentana.Panel.Panels(2).Text = "BASE DATOS: " & UCase(VGcnx.DefaultDatabase)
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
   pVentana.Panel.Panels(3).Alignment = sbrCenter
   pVentana.Panel.Panels(3).Width = (pVentana.Width / 4)
   pVentana.Panel.Panels(4).Text = "HORA :" & Format(Time, "hh:mm:ss")
   pVentana.Panel.Panels(4).Alignment = sbrRight
   pVentana.Panel.Panels(4).Width = (pVentana.Width / 4)

End Function

Public Function Limpiartexto(MBox As Object, ninicio As Integer, nfin As Integer, Optional Noincluir1, Optional Noincluir2 As Integer)
 Dim j As Integer
 If IsMissing(Noincluir1) Then
    Noincluir1 = -1
 End If
 If IsMissing(Noincluir2) Then
    Noincluir2 = -1
 End If
   
 For j = ninicio To nfin
   If j = Noincluir1 Or j = Noincluir2 Then
   Else
      MBox(j) = Empty
   End If
 Next j
End Function

Public Function Numero(Number) As String
   Dim aValor As Double
   If IsNull(Number) Or Len(Trim(Number)) = 0 Then
     aValor = 0
   Else
     aValor = Number
   End If
   Numero = Right(Space(10) & Trim(Format(aValor, "##,###0.00")), 10)
End Function
Public Function MontoCero(Number) As String
   Dim aValor As Double
   If IsNull(Number) Or Len(Trim(Number)) = 0 Then
     aValor = 0
   Else
     aValor = Number
   End If
   MontoCero = Right(Space(10) & Trim(Format(aValor, "##,###0.0000")), 10)
End Function

Public Function Escadena(pDato) As String
   If IsNull(pDato) Or Len(Trim(pDato)) = 0 Then
     Escadena = ""
   Else
     Escadena = Trim(pDato)
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
    xdepo = Numero(asumar)
End Sub

Public Function Seguir(MBox As Object, ntecla As Integer)
    If ntecla = 13 Then
        SendKeys "{tab}"
    End If
End Function

Public Function DatoMoneda(xValor As String) As String
   Dim rmone As New ADODB.Recordset
   
   Set rmone = VGcnx.Execute("select * from gr_moneda where monedacodigo='" & xValor & "'")
   If rmone.RecordCount > 0 Then
       DatoMoneda = Escadena(rmone!monedasimbolo) & " ."
   Else
       DatoMoneda = " "
   End If
   rmone.Close
   Set rmone = Nothing

End Function

Public Function VerificaCombo(xcombo As ComboBox, ncadena As String) As Long
    Dim j, k As Long
    On Error GoTo nerror
    VerificaCombo = -1
    If xcombo.ListCount > 0 Then
      VerificaCombo = 0
      For j = 0 To xcombo.ListCount - 1
         xcombo.ListIndex = j
         k = InStr(xcombo, "-")
         If k > 1 Then
           If Left(xcombo.Text, k - 1) = ncadena Then
             VerificaCombo = j
             Exit For
           End If
         Else
           If xcombo.Text = ncadena Then
             VerificaCombo = j
             Exit For
           End If
         End If
      Next j

    End If
    
nerror:
  If Err Then
    MsgBox Err.Number & "-" & Err.Description
    Err = 0
    Resume Next
  End If
End Function

Public Sub Imprimir(cNombreReporte As String)
Dim Busca As New dll_apisgen.dll_apis
On Error GoTo Errores

   MDIPrincipal.oCrystalReport.Reset
   MDIPrincipal.oCrystalReport.Destination = crptToWindow
   MDIPrincipal.oCrystalReport.WindowState = crptMaximized
   MDIPrincipal.oCrystalReport.ReportFileName = RutaRep & cNombreReporte
   
   MDIPrincipal.oCrystalReport.LogOnServer "pdssql.dll", _
         Busca.LeerIni(App.Path & "\Marfice.ini", "BDGENERAL", "bddatos", ""), _
         Busca.LeerIni(App.Path & "\Marfice.ini", "BDGENERAL", "servidor", ""), _
         Busca.LeerIni(App.Path & "\Marfice.ini", "BDGENERAL", "usuario", ""), _
         Busca.LeerIni(App.Path & "\Marfice.ini", "BDGENERAL", "passw", "")
   MDIPrincipal.oCrystalReport.Connect = _
        "DSN=" & Busca.LeerIni(App.Path & "\Marfice.ini", "BDGENERAL", "servidor", "") & ";" & _
        "DSQ=" & Busca.LeerIni(App.Path & "\Marfice.ini", "BDGENERAL", "bddatos", "") & ";" & _
        "UID=" & Busca.LeerIni(App.Path & "\Marfice.ini", "BDGENERAL", "usuario", "") & ";" & _
        "PWD=" & Busca.LeerIni(App.Path & "\Marfice.ini", "BDGENERAL", "passw", "")
   
   MDIPrincipal.oCrystalReport.DiscardSavedData = True
   MDIPrincipal.oCrystalReport.Formulas(0) = "Empresa='" & g_DetalleEmpresa & "'"
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

Sub ImpresionRptProc(cNombreReporte As String, PFormulas(), Param(), Optional orden As String, Optional Titulo As String)
Dim I As Integer
Dim vgCADENAREPORT2 As String
On Error GoTo X
    vgCADENAREPORT2 = "DSN=" & VGParamSistem.Servidor & ";DSQ=" & VGParamSistem.BDConfigura & ";UID=" & VGParamSistem.Usuario & ";PWD=" & VGParamSistem.Pwd
    Screen.MousePointer = 11
    With MDIPrincipal.CryRptProc
        .Reset
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .WindowTitle = Titulo
        Call PropCrystal(MDIPrincipal.CryRptProc)
        .ReportFileName = RutaRepProc & cNombreReporte
        .LogOnServer "pdssql.dll", VGParamSistem.Servidor, VGParamSistem.BDConfigura, VGParamSistem.Usuario, VGParamSistem.Pwd
        .Connect = vgCADENAREPORT2
        If UBound(PFormulas) > 0 Then
            For I = 0 To UBound(PFormulas) - 1
                .Formulas(I) = PFormulas(I)
            Next
        End If
        .DiscardSavedData = True
        If UBound(Param) > 0 Then
            For I = 0 To UBound(Param) - 1
                .StoredProcParam(I) = Param(I)
            Next
        End If
        If orden <> "" Then Call CrystOrden(MDIPrincipal.CryRptProc, orden)
        If .Status <> 2 Then .Action = 1
    End With
    Screen.MousePointer = 1
    Exit Sub
X:
  If Err.Number = 9 Then Resume Next
  Screen.MousePointer = 1
  MsgBox "Error inesperado: " & Err.Number & "  " & Err.Description, vbExclamation
End Sub

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
Dim xMonto As Double
Dim monto As String
Dim SQL As String
    ReDim arrparm(2)
    ReDim arrform(3)
    '@Base  ,@Nrecibo
    
    Set rs = New ADODB.Recordset
    SQL = "select a.monedacodigo,b.detrec_monedacancela,sum(b.detrec_importesoles) as detrec_importesoles, sum(b.detrec_importedolares) as detrec_importedolares, a.cabrec_tipocambio, a.cabrec_numreciboegreso "
    SQL = SQL & "FROM te_cabecerarecibos a, te_detallerecibos b "
    SQL = SQL & "WHERE a.cabrec_numrecibo=b.cabrec_numrecibo AND "
    SQL = SQL & "a.cabrec_numrecibo='" & Trim(Str(Nrecibo)) & "' "
    SQL = SQL & "Group by a.monedacodigo,b.detrec_monedacancela,a.cabrec_tipocambio,a.cabrec_numreciboegreso"
    Set rs = VGcnx.Execute(SQL)
    If rs.BOF Or rs.EOF Then Exit Sub
    If rs.Fields("monedacodigo") = "01" Then
       If rs.Fields("detrec_monedacancela") = "01" Then
          xMonto = rs.Fields("detrec_importesoles")
       Else
         xMonto = rs.Fields("detrec_importedolares") * rs.Fields("cabrec_tipocambio")
       End If
    Else
       If rs.Fields("detrec_monedacancela") = "01" Then
         xMonto = rs.Fields("detrec_importesoles") / rs.Fields("cabrec_tipocambio")
       Else
         xMonto = rs.Fields("detrec_importedolares")
       End If
    End If
    
    If rs.RecordCount > 0 Then
       monto = Format(xMonto, "#########.00")
       monto = monto + 0.001
       arrparm(0) = VGParamSistem.BDEmpresa
       arrparm(1) = Nrecibo
       arrform(0) = "@Emp='" & Empresa.descripcion & "'"
       arrform(1) = "@NumeroLetras='" & NUMLET(monto) & "'"
       If rs.Fields("cabrec_numreciboegreso") <> Empty Then
          arrform(2) = "@NroTransferencia='" & "Nro Transferencia: " & rs.Fields("cabrec_numreciboegreso") & "'"
       Else
          arrform(2) = "@NroTransferencia='" & rs.Fields("cabrec_numreciboegreso") & "'"
       End If
       Call ImpresionRptProc("TeVoucher.rpt", arrform, arrparm)
    Else
       MsgBox "No existen datos del Nº de Recibo " & Str(Nrecibo)
    End If
    
    rs.Close
    Set rs = Nothing
    
   'Carlos dice que borres de CONTABILIDAD DOS el registro=30 del me de agosto de la tabla CT030108. dice que coordines con melissa
    
End Sub

Public Function NUMLET(num As String) As String
Dim cLET As String
Dim cWork As String
Dim cUNIDAD As String
Dim cDECENA As String
Dim cCENTENA As String
Dim nMODULUS As Integer
Dim NI As Integer
Dim nK As Integer
Dim Lit1 As String
Dim Lit2 As String
Dim Lit3 As String
Dim Lit4 As String
Dim Lit5 As String
Lit1 = "Uno    Doc    Trec   Cuatroc  Quin   Seisc  Setec  Ochoc  Novec  "
Lit2 = "Diez     Veinte   Treinta  Cuarenta CincuentaSesenta  Setenta  Ochenta  Noventa  "
Lit3 = "Once      Doce      Trece     Catorce   Quince    Dieciseis DiecisieteDieciocho Diecinueve"
Lit4 = "Uno   Dos   Tres  CuatroCinco Seis  Siete Ocho  Nueve "
Lit5 = "Millon    Billon    Trillon   CuatrillonQuintillon"
'Proceso Input = Num , Output = Let

cLET = ""

'Dim NUM As Double
'NUM = Val(NUMx)

If num > 0.99 Then
    'Separa los Enteros en una Cadena de Caracteres
     If InStr(1, Trim(Str(num)), ".", 0) > 0 Then
        cWork = Mid(Trim(Str(num)), 1, InStr(1, Trim(Str(num)), ".", 0) - 1)
     Else
        cWork = Str(num)
     End If
     nMODULUS = Int(Len(Trim(cWork)) / 3)
     nMODULUS = Len(Trim(cWork)) - (nMODULUS * 3)
     
     If nMODULUS > 0 Then
        cWork = String(3 - nMODULUS, "0") & Trim(cWork)
     End If
     
     nK = (Len(Trim(cWork)) / 3) - 1
    'Procesa de Mil en Mil
     NI = 1
     Do While NI < Len(Trim(cWork)) - 1
        cCENTENA = Mid(Trim(cWork), NI, 1)
        cDECENA = Mid(Trim(cWork), NI + 1, 1)
        cUNIDAD = Mid(Trim(cWork), NI + 2, 1)
        'Centenas
        If cCENTENA <> "0" Then
            If cCENTENA = "1" Then
                cLET = cLET & "Cien "
                If cDECENA <> "0" Or cUNIDAD <> "0" Then
                    cLET = Mid(cLET, 1, (Len(cLET) - 1)) & "to "
                End If
            Else
                cLET = cLET & Trim(Mid(Lit1, ((Val(cCENTENA) - 1) * 7) + 1, 7)) & "ientos "
            End If
        End If
        'Decenas
        If cDECENA <> "0" Then
            If cDECENA = "1" And cUNIDAD <> "0" Then
                If ((Val(cUNIDAD) - 1) * 10) + 1 > 0 Then cLET = cLET + Trim(Mid(Lit3, ((Val(cUNIDAD) - 1) * 10) + 1, 10))
            Else
                If ((Val(cDECENA) - 1) * 9) + 1 > 0 Then cLET = cLET + Trim(Mid(Lit2, ((Val(cDECENA) - 1) * 9) + 1, 9))
            End If
        End If
        'Unidades
        If cUNIDAD <> "0" Then
            If cDECENA > "1" Then
                cLET = Mid(cLET, 1, (Len(cLET) - 1)) & "i"
                If ((Val(cUNIDAD) - 1) * 6) + 1 > 0 Then cLET = cLET + LCase(Trim(Mid(Lit4, ((Val(cUNIDAD) - 1) * 6) + 1, 6)))
            Else
                If cDECENA < "1" Then
                    If ((Val(cUNIDAD) - 1) * 6) + 1 > 0 Then cLET = cLET + Trim(Mid(Lit4, ((Val(cUNIDAD) - 1) * 6) + 1, 6))
                End If
            End If
        End If
        cLET = cLET & " "
        'Pone Miles o Millones
        If nK > 0 Then
            If cCENTENA & cDECENA & cUNIDAD = "001" Then
                cLET = Mid(cLET, 1, Len(cLET) - 2) & " "
            End If
            nMODULUS = Int(nK / 2)
            nMODULUS = nK - (nMODULUS * 2)
            If nMODULUS = 0 Then
                cLET = cLET + Trim(Mid(Lit5, (((nK / 2) - 1) * 10) + 1, 10))
                If cCENTENA & cDECENA & cUNIDAD = "001" Or num > 1999999 Then
                    cLET = cLET & "es "
                Else
                    cLET = cLET & " "
                End If
            Else
                If cCENTENA & cDECENA & cUNIDAD > "000" Then
                    cLET = cLET & "Mil "
                End If
            End If
            nK = nK - 1
        End If
        NI = NI + 3
    Loop
    cLET = cLET & "con "
End If
If InStr(1, Trim(Str(num)), ".", 0) > 0 Then
    cLET = cLET + Mid(Trim(Str(num)), InStr(1, Trim(Str(num)), ".", 0) + 1, 2) & "/100" & " "
Else
    cLET = cLET + "00/100" & " "
End If
NUMLET = cLET
End Function
Public Function Espunto(ByRef texto As Variant) As Variant
    If Trim(texto) = "." Then
        Espunto = "0"
      Else
        Espunto = texto
    End If
End Function
Public Sub GeneraAsientoEnlineaTesor(fecha As Date, m_opcion As String, Nrecibo As String, OP As Integer, comprobconta As String)
Dim rsparimpo As ADODB.Recordset
Dim Comando As ADODB.Command
    Set rsparimpo = New ADODB.Recordset
    rsparimpo.Open "Select * From  ct_importartesoreria Where Left(Upper(tipooperacion),1) ='" & UCase(m_opcion) & "'", cnconta, adOpenKeyset, adLockReadOnly
    
    Set Comando = New ADODB.Command
        'cg.BeginTrans
        With Comando
            .CommandType = adCmdStoredProc
            .CommandText = "te_GeneraAsientosTesoreriaLinea_pro"
            .CommandTimeout = 0
            .ActiveConnection = cg
            .Parameters.Refresh
            .Parameters("@BaseConta") = cnconta.DefaultDatabase
            .Parameters("@BaseVenta") = VGcnx.DefaultDatabase
            .Parameters("@Asiento") = rsparimpo!Asiento
            .Parameters("@SubAsiento") = rsparimpo!SubAsiento
            .Parameters("@Libro") = rsparimpo!Libro
            
            .Parameters("@Mes") = Format(Month(fecha), "00")
            .Parameters("@Ano") = Year(fecha)
            
            .Parameters("@tipanal") = "002"
            .Parameters("@Compu") = VGComputer
            .Parameters("@Usuario") = VGParamSistem.Usuario
            .Parameters("@TipoMov") = Trim(UCase(m_opcion))
            .Parameters("@Nrecibo") = Nrecibo
            .Parameters("@op") = OP
            .Parameters("@comprobconta") = comprobconta
            .Execute
        End With
        'cg.CommitTrans
        Screen.MousePointer = 1
        MsgBox "La Operacion se Realizo Satisfactoriamente", vbInformation, "Sistema de Tesoreria"
        Exit Sub
Proceso:
        Screen.MousePointer = 1
        'cg.RollbackTrans
        MsgBox Err.Description
End Sub
Public Sub GeneraAsientoEnlineaTesorTransfer(fecha As Date, Nrecibo As String)
Dim rsparimpo As ADODB.Recordset
Dim Comando As ADODB.Command
    Set rsparimpo = New ADODB.Recordset
    rsparimpo.Open "Select * From  ct_importartesoreria Where Left(Upper(tipooperacion),1) ='T'", cnconta, adOpenKeyset, adLockReadOnly
    Set Comando = New ADODB.Command
        With Comando
            .CommandType = adCmdStoredProc
            .CommandText = "te_GeneraAsientosTesoreriaTransflinea_pro"
            .ActiveConnection = cg
            .Parameters.Refresh
            .Parameters("@BaseConta") = cnconta.DefaultDatabase
            .Parameters("@BaseVenta") = VGcnx.DefaultDatabase
            .Parameters("@Asiento") = rsparimpo!Asiento
            .Parameters("@SubAsiento") = rsparimpo!SubAsiento
            .Parameters("@Libro") = rsparimpo!Libro
            
            .Parameters("@Mes") = Format(Month(fecha), "00")
            .Parameters("@Ano") = Year(fecha)
            
            .Parameters("@Compu") = VGComputer
            .Parameters("@Usuario") = VGParamSistem.Usuario
            .Parameters("@Ntransfer") = Nrecibo
            .Execute
        End With
        Screen.MousePointer = 1
        MsgBox "La Operacion se Realizo Satisfactoriamente", vbInformation, "Sistema de Tesoreria"
        Exit Sub
Proceso:
        Screen.MousePointer = 1
        MsgBox Err.Description
End Sub

Public Function ArmaCriterioComodin(cad As String, Campo As String) As String
Dim pos As Integer, cadaux As String, criterio As String
Dim valor As String
    criterio = ""
    Do While True
        pos = InStr(1, cad, "%", vbTextCompare)
        If pos = 0 Then Exit Do
        valor = "'" & Left(cad, pos) & "'"
        cad = Right(cad, (Len(cad) - pos))
        criterio = criterio & Campo & " like " & valor & " or "
    Loop
    ArmaCriterioComodin = Left(criterio, Len(criterio) - 3)
End Function
Public Function FiltroCcosto(Codigo As String, ByRef flag As Boolean) As String
Dim rsaux As ADODB.Recordset
Dim AuxCad As String, Filtro As String
Dim tipocontrol As String
    Set rsaux = New ADODB.Recordset
    flag = False
    tipocontrol = 1
    If tipocontrol = 0 Then
       Filtro = "centrocostocodigo<>'00' and centrocostotipo='6'"
       rsaux.Open "Select criterio=isnull(conceptotextccosto,''),flagx=isnull(conceptosiccosto,0)  From te_conceptocaja Where conceptocodigo='" & Trim(Codigo) & "'", VGcnx, adOpenKeyset, adLockReadOnly
       If rsaux.RecordCount > 0 Then
          flag = rsaux!flagx
          If flag Then
             If rsaux!criterio <> "" Then Filtro = Filtro & " and (" & ArmaCriterioComodin(rsaux!criterio, "centrocostocodigo") & ")"
          End If
       End If
     Else
       Filtro = "gastoscodigo<>'00'"
       rsaux.Open "Select criterio=isnull(conceptotextccosto,''),flagx=isnull(conceptosiccosto,0)  From te_conceptocaja Where conceptocodigo='" & Trim(Codigo) & "'", VGcnx, adOpenKeyset, adLockReadOnly
       If rsaux.RecordCount > 0 Then
          flag = rsaux!flagx
'          If flag Then
'             If rsaux!criterio <> "" Then Filtro = Filtro & " and (" & ArmaCriterioComodin(rsaux!criterio, "gastoscodigo") & ")"
'          End If
      End If
    End If
    FiltroCcosto = Filtro
End Function
Public Function ExisteElem(Tip As Integer, VGcnx As ADODB.Connection, TABLA As String, _
        Optional Campo As String) As Boolean
'Funcion que devuelve un valor verdadero si es que encuentra el elemento
'Creado por Fernando Cossio
    Dim SQL As String
    Dim rsaux As New ADODB.Recordset
   '*------------------------------*
   '0 Si Existe la tabla
   '1 Si Existe el Campo
   ExisteElem = False
   TABLA = UCase(TABLA): Campo = UCase(Campo)
On Error GoTo ErrExiste
   SQL = ""
    Select Case Tip
        Case 0:
            SQL = "Select Top 1 * From " & TABLA
        Case 1:
            SQL = "Select Top 1 " & Campo & " From " & TABLA
    End Select
    rsaux.Open SQL, VGcnx
    ExisteElem = True
    Exit Function
ErrExiste:
    ExisteElem = False
End Function

Public Function FechS(fecha As Variant, Tipo As TIPFECHA) As Variant
Dim H As Date
Dim fechaAux As Double
On Error GoTo ErrorFecha
   H = CDate(fecha)
   Select Case Tipo
      Case Sqlf: 'Para transformar al sql
        fechaAux = DateSerial(Year(fecha), Month(fecha), Day(fecha)) - 2
      Case Adof: 'Para transformar al ado Y AL ACCESS
         fechaAux = DateSerial(Year(fecha), Month(fecha), Day(fecha))
   End Select
   FechS = fechaAux
   Exit Function
ErrorFecha:
   Select Case Tipo
      Case Sqlf: FechS = "Null"
      Case Adof: FechS = Null
   End Select
End Function
Sub ImpresionRpt_SubRpt_Proc(cNombreReporte As String, PFormulas(), Param(), cNombreSubRpt As String, Optional orden As String, Optional Titulo As String)
Dim Busca As New dll_apis
Dim I As Integer
On Error GoTo X
    Screen.MousePointer = 11
    With MDIPrincipal.oCrystalReport
        .Reset
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .WindowTitle = Titulo
        .ReportFileName = RutaRepProc & cNombreReporte
   '     .LogOnServer "pdssql.dll", _
   '      Busca.LeerIni(App.Path & "\Marfice.ini", "BDGENERAL", "bddatos", ""), _
   '      Busca.LeerIni(App.Path & "\Marfice.ini", "BDGENERAL", "servidor", ""), _
   '      Busca.LeerIni(App.Path & "\Marfice.ini", "BDGENERAL", "usuario", ""), _
   '      Busca.LeerIni(App.Path & "\Marfice.ini", "BDGENERAL", "passw", "")
        .Connect = _
        "DSN=" & Busca.LeerIni(App.Path & "\Marfice.ini", "BDGENERAL", "servidor", "") & ";" & _
        "DSQ=" & Busca.LeerIni(App.Path & "\Marfice.ini", "BDGENERAL", "bddatos", "") & ";" & _
        "UID=" & Busca.LeerIni(App.Path & "\Marfice.ini", "BDGENERAL", "usuario", "") & ";" & _
        "PWD=" & Busca.LeerIni(App.Path & "\Marfice.ini", "BDGENERAL", "passw", "")
        Call PropCrystal(MDIPrincipal.oCrystalReport)
        .Formulas(0) = "@Empresa='" & g_DetalleEmpresa & "'"
        .Formulas(1) = "@Ruc='" & "20293847038" & "'"
        If UBound(PFormulas) > 0 Then
            For I = 0 To UBound(PFormulas) - 1
                .Formulas(2 + I) = PFormulas(I)
            Next
        End If
        .DiscardSavedData = True
        If UBound(Param) > 0 Then
            For I = 0 To UBound(Param) - 1
                .StoredProcParam(I) = Param(I)
            Next
        End If
        '***Para el SubReporte
        .SubreportToChange = cNombreSubRpt
   '     .LogOnServer "pdssql.dll",
   '      Busca.LeerIni(App.Path & "\Marfice.ini", "BDGENERAL", "bddatos", ""), _
   '      Busca.LeerIni(App.Path & "\Marfice.ini", "BDGENERAL", "servidor", ""), _
   '      Busca.LeerIni(App.Path & "\Marfice.ini", "BDGENERAL", "usuario", ""), _
   '      Busca.LeerIni(App.Path & "\Marfice.ini", "BDGENERAL", "passw", "")
        .Connect = _
        "DSN=" & Busca.LeerIni(App.Path & "\Marfice.ini", "BDGENERAL", "servidor", "") & ";" & _
        "DSQ=" & Busca.LeerIni(App.Path & "\Marfice.ini", "BDGENERAL", "bddatos", "") & ";" & _
        "UID=" & Busca.LeerIni(App.Path & "\Marfice.ini", "BDGENERAL", "usuario", "") & ";" & _
        "PWD=" & Busca.LeerIni(App.Path & "\Marfice.ini", "BDGENERAL", "passw", "")
        If UBound(Param) > 0 Then
            For I = 0 To UBound(Param) - 1
                .StoredProcParam(I) = Param(I)
            Next
        End If
        .StoredProcParam(3) = "1"
        If .Status <> 2 Then .Action = 1
    End With
    Screen.MousePointer = 1
    Exit Sub
X:
  If Err.Number = 9 Then Resume Next
  Screen.MousePointer = 1
  MsgBox "Error inesperado: " & Err.Number & "  " & Err.Description, vbExclamation
End Sub
