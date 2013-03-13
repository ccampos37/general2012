Attribute VB_Name = "FuncionesTesoreria"
  Option Explicit
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
Dim rsql As New ADODB.Recordset
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
       Set rsql = VGCNx.Execute(SQL)
       arrparm(2) = "0000000"
       If rsql.RecordCount() > 0 Then
          arrparm(2) = Trim(rsql!detrec_ndqc)
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
            If tipo = 1 Then
               Set rsaux = cnx.Execute("SELECT top 1 Numx=isnull(empresanumeingreso,1) from te_parametroempresa ")
             Else
               Set rsaux = VGCNx.Execute("SELECT top 1 Numx=isnull(empresanumegreso,1) from te_parametroempresa ")
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
If Not ExisteElem(0, VGCNx, VGcomputer + "_cajaconcil") Then
   SQL = " Create Table " & VGcomputer & "_cajaconcil (recibo VarChar(6))"
 Else
    SQL = " delete " & VGcomputer & "_cajaconcil"
End If
VGCNx.Execute (SQL)
If tipo = "1" Then
   Do Until rs.EOF()
     If IsNull(rs!chkconcil) = False And rs!chkconcil Then
           SQL = "insert into " & VGcomputer & "_cajaconcil (recibo) values ('" & rs!cabrec_numrecibo & "')"
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
        SQL = "insert into " & VGcomputer & "_cajaconcil (recibo) values ('" & rs!cabrec_numrecibo & "')"
        VGCNx.Execute (SQL)
     End If
     rs.MoveNext
   Loop
End If
recibosrendicion = VGcomputer + "_cajaconcil"
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


