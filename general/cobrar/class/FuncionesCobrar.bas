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
Public Sub ParametrosFuncionalesCobrar()
   Dim rs2 As New ADODB.Recordset
   Dim rb As New ADODB.Recordset
   Set VGvardllgen = New dllgeneral.dll_general
    Set rs2 = VGCNx.Execute("select top 1 * from cc_parametros ")
   If ExisteElem(1, VGCNx, "cc_parametros", "contabilizaenlinea") Then
       VGparametros.contabilizaenlinea = IIf(VGvardllgen.ESNULO(rs2!contabilizaenlinea, 0) = 0, False, True)
    End If
   
   Set rs2 = VGCNx.Execute("select top 1 * from co_sistema ")
   If ExisteElem(1, VGCNx, "co_sistema", "sistemamultiempresas") Then
       VGparametros.sistemamultiempresas = IIf(VGvardllgen.ESNULO(rs2!sistemamultiempresas, 0) = 0, False, True)
    End If
    
    If VGparametros.sistemamultiempresas = False Then VGparametros.empresacodigo = "01"
   VGCNx.Execute "set dateformat dmy"    '--seteo de formato de fecha

End Sub


Public Sub Cargar_Parametros_Funcionales()
   Dim rs1 As New ADODB.Recordset
   Dim rs2 As New ADODB.Recordset
   
   Dim rb As New ADODB.Recordset
   Dim averi As New dllgeneral.dll_general
       
    Set rs1 = Nothing
   g_PedidoPuntoVta = "Tempopedido" & Trim(g_ptoventa)
   g_DetallePuntoVta = "Tempodetallepedido" & Trim(g_ptoventa)
   Set rs2 = Nothing
   Set rs2 = VGCNx.Execute("select * from vt_parametroventa ")
   If rs2.RecordCount > 0 Then
      VGparametros.nombre = Escadena(rs2!empresacodigo)
      VGparametros.tienedscto = Escadena(rs2!paramvtaestdesc)
      VGparametros.descuento = IIf(IsNull(rs2!paramvtadescto), 0, CDbl(rs2!paramvtadescto))
      VGparametros.moneda = Escadena(rs2!monedacodigo)
      VGparametros.tieneigv = IIf(IsNull(rs2!paramvtaestigv) Or rs2!paramvtaestigv = 0, "0", "1")
       VGparametros.igv = IIf(IsNull(rs2!paramvtaporcigv), 0, CDbl(rs2!paramvtaporcigv))
      VGparametros.almacen = Escadena(rs2!almacencodigo)
      VGparametros.mensaje = Escadena(rs2!paramvtamensaje)
      VGparametros.listapre = IIf(IsNull(rs2!paramvtalistaprec), "", Escadena(rs2!paramvtalistaprec))
      VGparametros.tipocambio = IIf(IsNull(rs2!paramvtatipcambref), CDbl(0), CDbl(rs2!paramvtatipcambref))
      VGparametros.comivende = IIf(IsNull(rs2!paramvtacomisionvendedor), "", Escadena(rs2!paramvtacomisionvendedor))
      VGparametros.formaemi = IIf(IsNull(rs2!paramvtaformaemision), "", Escadena(rs2!paramvtaformaemision))
      VGparametros.paraboleta = IIf(IsNull(rs2!paramvtaboleta) Or rs2!paramvtaboleta = 0, "0", "1")
       
   End If
   rs2.Close
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
   
   Set rs2 = VGCNx.Execute("select * from vt_puntoventa where puntovtacodigo='" & g_ptoventa & "'")
   If rs2.RecordCount > 0 Then
        VGparamsistem.puntovta = Escadena(rs2!puntovtacodigo)
        VGparamsistem.nropedido = Escadena(IIf(IsNull(rs2!puntovtanropedido) Or rs2!puntovtanropedido = 0, "0", "1"))
        VGparamsistem.nrofactura = Escadena(IIf(IsNull(rs2!puntovtanrofact) Or rs2!puntovtanropedido = 0, "0", "1"))
        VGparamsistem.nroguia = Escadena(IIf(IsNull(rs2!puntovtanroguiarem) Or rs2!puntovtanropedido = 0, "0", "1"))
        VGparamsistem.nroabono = Escadena(IIf(IsNull(rs2!puntovtanotaabono) Or rs2!puntovtanropedido = 0, "0", "1"))
        VGparamsistem.nrocargo = Escadena(IIf(IsNull(rs2!puntovtanotacargo) Or rs2!puntovtanropedido = 0, "0", "1"))
        VGparamsistem.ventaauto = Escadena(IIf(IsNull(rs2!puntovtaautomat) Or rs2!puntovtanropedido = 0, "0", "1"))
   End If
   rs2.Close
   Set rs2 = Nothing
   
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
   
   Set rs2 = VGCNx.Execute("select top 1 * from co_sistema ")
   If ExisteElem(1, VGCNx, "co_sistema", "sistemamultiempresas") Then
       VGparametros.sistemamultiempresas = IIf(averi.ESNULO(rs2!sistemamultiempresas, 0) = 0, False, True)
    End If
    If VGparametros.sistemamultiempresas = False Then
       VGparametros.empresacodigo = "01"
    End If
   Set rs2 = VGCNx.Execute("select top 1 * from cc_sistema ")
   If ExisteElem(1, VGCNx, "cc_sistema", "imprimevoucher") Then
       VGparametros.imprimevoucher = ESNULO(rs2!imprimevoucher, 0)
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

Private Sub CrystOrden(ByRef cry As CrystalReport, cad As String)
Dim pos As Integer, cadaux As String, i As Integer
Dim Valor As String
    Do While True
        pos = InStr(1, cad, ",", vbTextCompare)
        i = 0
        If pos = 0 Then Exit Do
        Valor = Left(cad, pos - 1)
        cry.SortFields(i) = Valor
        i = i + 1
        cad = Right(cad, (Len(cad) - pos))
    Loop
End Sub

Public Sub ModoEditable(flagModo As Boolean, Formu As Form, cnameCtrX As String)
 Dim i As Integer
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
 Dim i As Integer
    Dim Control As Control
    For Each Control In Formu.Controls
       If UCase(Control.Name) <> UCase(cnameCtrX) Then
           If TypeOf Control Is TextBox Then Control.Text = Empty
           If TypeOf Control Is TextFer.TxFer Then Control.Text = Empty
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
