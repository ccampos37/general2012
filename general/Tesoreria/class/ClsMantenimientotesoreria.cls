VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsMantMovimientos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Enum CamposTesoreria
    formapago = 1
    gastostes = 2
    glosates = 3
    analiticotes = 4
    ccostotes = 5
    importesoles = 6
    importedolares = 7
    banco = 8
    fechacancelacion = 9
    concepto = 10
    cheque = 11
    ctacte = 12
    docpago = 13
    rendicion = 14
    costo = 15

End Enum
Public Sub CargarAyudas(Optional optipo As Integer = 0)
If optipo = 0 Then
   With frmMantRecibos
    Call .CtrAyu_Moneda.conexion(VGCNx): .CtrAyu_Moneda.Filtro = "monedacodigo<>'00'"
    Call .Ctr_AyuMonedacancela.conexion(VGCNx): .Ctr_AyuMonedacancela.Filtro = "monedacodigo<>'00'"
    Call .CtrAyu_ModoOperacion.conexion(VGCNx): .CtrAyu_ModoOperacion.Filtro = "operacioncodigo <>'00'"
    Call .Ctr_AyuAnalitico.conexion(VGCNx)
    Call .Ctr_Ayuctacte.conexion(VGCNx)
    Call .Ctr_Ayugastos.conexion(VGCNx)
    Call .CtrAyu_Proveedor.conexion(VGCNx): .CtrAyu_Proveedor.Filtro = "clientecodigo <>'00'"
    Call .Ctr_AyudaCaja.conexion(VGCNx): .Ctr_AyudaCaja.Filtro = " isnull(CajaCuentaxRendir,0)=" & .m_cajachica
    Call .Ctr_Ayuconcepto.conexion(VGCNx)
    Call .Ctr_AyuBanco.conexion(VGCNx): .Ctr_AyuBanco.Filtro = "bancocodigo<>'00' "
    Call .Ctr_AyuFormaPago.conexion(VGCNx): .Ctr_AyuFormaPago.Filtro = "tdocumentotipo = 'A'"
    Call .Ctr_AyudaOficina.conexion(VGCNx)
    Call .Ctr_AyuCcosto.conexion(VGCNx): .Ctr_AyuCcosto.Filtro = "centrocostocodigo <> '00'"
    Call .Ctr_Ayuempresa.conexion(VGCNx): .Ctr_Ayuempresa.Filtro = "empresacodigo<>'00'"
End With
End If
If optipo = 1 Then
   With frmModrecibos
    Call .CtrAyu_Moneda.conexion(VGCNx): .CtrAyu_Moneda.Filtro = "monedacodigo<>'00'"
    Call .Ctr_AyuMonedacancela.conexion(VGCNx): .Ctr_AyuMonedacancela.Filtro = "monedacodigo<>'00'"
    Call .CtrAyu_ModoOperacion.conexion(VGCNx): .CtrAyu_ModoOperacion.Filtro = "operacioncodigo <>'00'"
    Call .Ctr_AyuAnalitico.conexion(VGCNx)
    Call .Ctr_Ayuctacte.conexion(VGCNx)
    Call .Ctr_Ayugastos.conexion(VGCNx)
    Call .CtrAyu_Proveedor.conexion(VGCNx): .CtrAyu_Proveedor.Filtro = "clientecodigo <>'00'"
    Call .Ctr_AyudaCaja.conexion(VGCNx): .Ctr_AyudaCaja.Filtro = " isnull(CajaCuentaxRendir,0)=" & .m_cajachica
    Call .Ctr_Ayuconcepto.conexion(VGCNx)
    Call .Ctr_AyuBanco.conexion(VGCNx): .Ctr_AyuBanco.Filtro = "bancocodigo<>'00' "
    Call .Ctr_AyuFormaPago.conexion(VGCNx): .Ctr_AyuFormaPago.Filtro = "tdocumentotipo = 'A'"
    Call .Ctr_AyudaOficina.conexion(VGCNx)
    Call .Ctr_AyuCcosto.conexion(VGCNx): .Ctr_AyuCcosto.Filtro = "centrocostocodigo <> '00'"
    Call .Ctr_Ayuempresa.conexion(VGCNx): .Ctr_Ayuempresa.Filtro = "empresacodigo<>'00'"
End With
End If
End Sub
Public Sub CreaRsTempDetalle(ByRef rs As ADODB.Recordset)
    With rs.Fields
        Call .Append("item", adVarChar, 3)
        Call .Append("tipodoc_concepto", adVarChar, 2, adFldIsNullable)
        Call .Append("numdocumento", adVarChar, 14, adFldIsNullable)
        Call .Append("tdqc", adVarChar, 2, adFldIsNullable)
        Call .Append("ndqc", adVarChar, 11, adFldIsNullable)
        Call .Append("cajabanco1", adVarChar, 2, adFldIsNullable)
        Call .Append("numctacte", adVarChar, 30, adFldIsNullable)
        Call .Append("monedadocumento", adVarChar, 2, adFldIsNullable)
        Call .Append("monedacancela", adVarChar, 2, adFldIsNullable)
        Call .Append("importesoles", adDouble)
        Call .Append("importedolares", adDouble)
        Call .Append("fechacancela", adDate)
        Call .Append("entidad", adVarChar, 11, adFldIsNullable)
        Call .Append("costos", adVarChar, 10, adFldIsNullable)
        Call .Append("gastos", adVarChar, 10, adFldIsNullable)
        Call .Append("observacion", adVarChar, 50, adFldIsNullable)
        Call .Append("rendicionnumero", adVarChar, 10)
        Call .Append("clientecodigo", adVarChar, 11)
   End With

End Sub
Public Sub AñadiralDetalle(ByRef rs As ADODB.Recordset)
Dim vardllgen As New dllgeneral.dll_general
    Call Limpia
    With rs
        .AddNew
        !Item = Format(Val(vardllgen.ESNULO(MaxItem(rs, "Item"), 0) + 1), "000")
        .Update
    End With
End Sub
Public Sub Limpia(Optional optipo As Integer = 0)
    VGflaglimpia = False
If optipo = 0 Then
With frmMantRecibos
        .Ctr_Ayugastos.xclave = "": .Ctr_Ayugastos.xnombre = ""
        .Txsoles.Text = ""
        .Txdolares.Text = ""
        .Txtglosa.Text = ""
        .Ctr_AyuCcosto.xclave = "": .Ctr_AyuCcosto.xnombre = ""
        .Ctr_AyuBanco.xclave = "": .Ctr_AyuFormaPago.xclave = ""
    End With
End If
If optipo = 1 Then
With frmModrecibos
        .Ctr_Ayugastos.xclave = "": .Ctr_Ayugastos.xnombre = ""
        .Txsoles.Text = ""
        .Txdolares.Text = ""
        .Txtglosa.Text = ""
        .Ctr_AyuCcosto.xclave = "": .Ctr_AyuCcosto.xnombre = ""
        .Ctr_AyuBanco.xclave = "": .Ctr_AyuFormaPago.xclave = ""
    End With
End If
VGflaglimpia = True
End Sub
Public Sub ActualizarDetalle(ByRef rs As ADODB.Recordset, op As Campos)
    Dim vardllgen As New dllgeneral.dll_general
    If rs Is Nothing Then Exit Sub
    If rs.RecordCount = 0 Then Exit Sub
  '  VGactulizodoc = False
    If VGMoverRegistro Then Exit Sub
    Screen.MousePointer = 11
    With rs
        '!detcomprobitem = Format(Val(vardllgen.ESNULO(MaxItem(rs, "detcomprobitem"), 0) + 1), "00000")
        Select Case op
            Case formapago
                !tdqc = frmModrecibos.Ctr_AyuFormaPago.xclave
                !costos = "00"
                !gastos = "00"
           Case gastos
                !gastos = frmModrecibos.Ctr_Ayugastos.xclave
            Case glosa: !observacion = Trim(frmModrecibos.Txtglosa.Text)
            Case analitico: !analitico = frmModrecibos.Ctr_AyuAnalitico.xclave
            Case costo: !costos = frmModrecibos.Ctr_AyuCcosto.xclave
            Case importesoles: !importesoles = vardllgen.ESNULO(Espunto(frmModrecibos.Txsoles.valor), 0)
            Case importedolares: !importedolares = vardllgen.ESNULO(Espunto(frmModrecibos.Txdolares.valor), 0)
            Case banco: !cajabanco1 = vardllgen.ESNULO(Espunto(frmModrecibos.Ctr_AyuBanco.xclave), 0)
            Case fechacancelacion: !fechacancelacion = frmModrecibos.DTPFechacancelacion
            Case concepto: !tipodoc_concepto = frmModrecibos.Ctr_AyuCcosto.xclave
            Case cheque: !ndqc = frmModrecibos.TxNroPago.Text
            Case ctacte: !numctacte = frmModrecibos.Ctr_Ayuctacte.xclave
            Case docpago: !numdocumento = frmModrecibos.TxNroPagado
            Case rendicion: !rendicionnumero = frmModrecibos.TxtRendicion
        
        End Select
    End With
    
    Screen.MousePointer = 0
End Sub
Private Function MaxItem(ByVal rs As ADODB.Recordset, ByVal Campo As String) As Variant
Dim rsauxiliar As New ADODB.Recordset
Set rsauxiliar = rs.Clone(adLockReadOnly)
On Error GoTo errormaxitem
    rsauxiliar.Sort = Campo & " desc"
    rsauxiliar.MoveFirst
    MaxItem = rsauxiliar.Fields(Campo).Value
    Exit Function
errormaxitem:
    MaxItem = Null
End Function
Public Sub GeneraAsientoenLine(ByVal op As Integer, ByVal Nprovi As String, ByVal Comprob_Contable As String)
On Error GoTo genasiento
    Screen.MousePointer = 11
    'Generando los Analticos que no Esten en contabilidad
    Set VGCommandoSP = New ADODB.Command
    Set VGvardllgen = New dllgeneral.dll_general
    VGCommandoSP.ActiveConnection = VGgeneral
    VGCommandoSP.CommandType = adCmdStoredProc
    VGCommandoSP.CommandText = "co_VerificaAnaliticoenLinea"
    VGCommandoSP.Parameters.Refresh
    With VGCommandoSP
        .Parameters("@BaseConta") = VGParamSistem.BDEmpresa
        .Parameters("@BaseCompra") = VGParamSistem.BDEmpresa
        .Parameters("@Ano") = VGParamSistem.AnoProceso
        .Parameters("@Mes") = VGParamSistem.MesProceso
        .Parameters("@tipanal") = "03"
        .Parameters("@User") = VGusuario
        .Parameters("@Nprovi") = Nprovi
        .Execute
    End With
    'Generando el Asiento en contabilidad
'    Set VGCommandoSP = New ADODB.Command
'    Set VGvardllgen = New dllgeneral.dll_general
'
'    VGCommandoSP.ActiveConnection = VGGeneral
'    VGCommandoSP.CommandType = adCmdStoredProc
'    VGCommandoSP.CommandText = "co_generaasientolinea_pro"
'    VGCommandoSP.Parameters.Refresh
'    With VGCommandoSP
'        .Parameters("@BaseConta") = VGParamSistem.BDEmpresa
'        .Parameters("@BaseCompra") = VGParamSistem.BDEmpresa
'        .Parameters("@SubAsiento") = "03"
'        .Parameters("@Libro") = "99"
'        .Parameters("@mes") = VGParamSistem.mesproceso
'        .Parameters("@Ano") = VGParamSistem.AnoProceso
'        .Parameters("@ctatotal") = VGParametros.xCtaTotal
'        .Parameters("@ctaIGV") = VGParametros.xCtaIGV
'        .Parameters("@ctaIES") = VGParametros.xCtaIES
'        .Parameters("@ctaRTA") = VGParametros.xCtaRTA
'        .Parameters("@tipanal") = VGParametros.xTipAnal
'        .Parameters("@Compu") = VGComputer
 '       .Parameters("@Usuario") = VGParamSistem.Usuario
 '       .Parameters("@Oficina") = Format(VGParametros.CpOficina, "00")
 '       .Parameters("@Nprovi") = Nprovi
 '       .Parameters("@op") = OP
 '       .Parameters("@comprobconta") = Comprob_Contable
 '       .Execute
 '   End With
    
 '   'Actualizando las Glosas de Cabecera y Detalle
  ''  Set VGCommandoSP = New ADODB.Command
 '   Set VGvardllgen = New dllgeneral.dll_general
 '   VGCommandoSP.ActiveConnection = VGGeneral
  '  VGCommandoSP.CommandType = adCmdStoredProc
 '   VGCommandoSP.CommandText = "co_GrabaGlosasProvisionLinea_pro"
 '   VGCommandoSP.Parameters.Refresh
 '   With VGCommandoSP
 '       .Parameters("@BaseConta") = VGParamSistem.BDEmpresa
 '       .Parameters("@BaseCompra") = VGParamSistem.BDEmpresa
 '       .Parameters("@mes") = VGParamSistem.mesproceso
 '       .Parameters("@Ano") = VGParamSistem.Anoproceso
 '       .Parameters("@Nprovi") = Nprovi
 '       .Execute
 '   End With
 '
 '   'Actualizando los Registros que no se incluyen en el Reg. Compras
 '   Set VGCommandoSP = New ADODB.Command
 '   Set VGvardllgen = New dllgeneral.dll_general
 '   VGCommandoSP.ActiveConnection = VGGeneral
 '   VGCommandoSP.CommandType = adCmdStoredProc
 ''   VGCommandoSP.CommandText = "co_RegComprasNoIncluyenenLinea_pro"
 '   VGCommandoSP.Parameters.Refresh
 '   With VGCommandoSP
 '       .Parameters("@BaseCompra") = VGParamSistem.BDEmpresa
 '       .Parameters("@BaseConta") = VGParamSistem.BDEmpresa
 '       .Parameters("@Asiento") = "081"
 '       .Parameters("@mes") = VGParamSistem.mesproceso
 '       .Parameters("@Ano") = VGParamSistem.AnoProceso
 '       .Parameters("@Nprovi") = Nprovi
 '       .Execute
 '   End With
 '
 '   MsgBox "Se Realizo la Operacion Satisfactoriamente"
 '   Screen.MousePointer = 1
 '   Exit Sub
genasiento:
'    Screen.MousePointer = 1
'    VGvarVerifica = False
'    VGErrorString = "Error en Grabar Cabecera " & Chr(13) & Err.Description
End Sub

Public Sub ActualizaCorrelComprob(ByVal numero As Double)
On Error GoTo Actualizacorre
    Set VGCommandoSP = New ADODB.Command
    VGCommandoSP.ActiveConnection = VGgeneral
    VGCommandoSP.CommandType = adCmdStoredProc
    VGCommandoSP.CommandText = "co_actcorraux_pro"
    VGCommandoSP.Parameters.Refresh
    With VGCommandoSP
        .Parameters("@Base") = VGParamSistem.BDEmpresa
        .Parameters("@Ano") = VGParamSistem.AnoProceso
        .Parameters("@Mes") = VGParamSistem.MesProceso
        .Parameters("@numero") = numero
        .Execute
    End With
    
    Exit Sub
Actualizacorre:
    VGvarVerifica = False
    VGErrorString = "Error en Actualizar el Numero de Comprobante Auxiliar " & Chr(13) & Err.Description
End Sub

Public Sub GrabarCP_abono(op As Integer, Optional numero As Long = 0, Optional ByVal rs As Recordset)
Set VGCommandoSP = New ADODB.Command
VGCommandoSP.ActiveConnection = VGgeneral
VGCommandoSP.CommandType = adCmdStoredProc
VGCommandoSP.CommandText = "cp_abonadocumento_pro"
VGCommandoSP.Parameters.Refresh
rs.MoveFirst
Do While Not rs.EOF()
   With VGCommandoSP
        .Parameters("@Base") = VGParamSistem.BDEmpresa
        .Parameters("@tipo") = "2"
        .Parameters("@tipoplanilla") = "TE"
        .Parameters("@numplanilla") = frmMantRecibos.lbNumComprobCab
        .Parameters("@documentoabono") = rs!tipodoc_concepto
        .Parameters("@abononumdoc") = rs!numdocumento
        .Parameters("@abonocantd") = rs!tdqc
        .Parameters("@abonocannro") = rs!ndqc
        .Parameters("@banco") = rs!cajabanco1
        .Parameters("@ctabanco") = rs!numctacte
        .Execute
    End With
    rs.MoveNext
Loop
End Sub
Public Sub Grabaren_Tesoreria(ByVal op As Integer, Optional ByVal Numeroprovi As Long = 0, Optional ByVal rs As Recordset, Optional ByRef Xnum As String)
'On Error GoTo ErrorGrabaTesore
Dim numero As String
Set VGvardllgen = New dllgeneral.dll_general
Dim rb As ADODB.Recordset
Dim Item As Integer
   'Obtener el Ultimo Numero Correlativo de las cajas
    Dim opaux As Integer
    opaux = op
    Set VGCommandoSP = New ADODB.Command
    Set VGvardllgen = New dllgeneral.dll_general
    VGCommandoSP.ActiveConnection = VGgeneral
    VGCommandoSP.CommandType = adCmdStoredProc
    VGCommandoSP.CommandText = "te_abonadocumento_pro"
    VGCommandoSP.Parameters.Refresh
    With VGCommandoSP
        .Parameters("@base") = VGParamSistem.BDEmpresa
        .Parameters("@tipo") = IIf(op = 3, 2, op)
        If op = 2 Or op = 3 Then
            'Set rb = VGcnx.Execute("Select cabprovinumtesor  From " & VGParamSistem.TablaCabcomprob & " Where cabprovinumero=" & frmMantRecibos.lbNumComprobCab.Caption)
            numero = frmMantRecibos.lbNumComprobCab
            
          ElseIf op = 1 Then
            Set rb = VGCNx.Execute("select * from te_parametroempresa where empresacodigo='01'")
            If rb.RecordCount > 0 Then
                numero = Format(CDbl(VGvardllgen.ESNULO(rb!empresanumegreso, "0")) + 1, "000000")
                VGCNx.Execute "Update te_parametroempresa Set empresanumegreso='" & numero & "' where empresacodigo='01'"
                'VGcnx.Execute "Update " & VGParamSistem.TablaCabcomprob & " Set cabprovinumtesor='" & Numero & _
                '              "' Where cabprovinumero=" & Numeroprovi
            End If
          Else
          numero = frmModrecibos.lbNumComprobCab
        End If
        Xnum = numero
        .Parameters("@estadoreg") = ""
        .Parameters("@numrecibo") = numero
        If op = 3 Then
            .Execute
        End If
        If op = 2 Then
            .Execute
            .Parameters("@tipo") = 1
           op = 1
        End If
        
        'Este para que al eliminar no utilizar estos parametros
        If op = 1 Then
            .Parameters("@controlctacte") = "N"
            .Parameters("@vendedorcodigo") = frmMantRecibos.Ctr_AyudaOficina.xclave
            .Parameters("@cajacodigo") = frmMantRecibos.Ctr_AyudaCaja.xclave
            .Parameters("@clientecodigo") = frmMantRecibos.CtrAyu_Proveedor.xclave
            .Parameters("@descripcion") = ""
            .Parameters("@operacion") = frmMantRecibos.CtrAyu_ModoOperacion.xclave
            .Parameters("@monedacodigo") = frmMantRecibos.CtrAyu_Moneda.xclave
            .Parameters("@ingsal") = frmMantRecibos.TxingEgr.Text
            .Parameters("@tipocambio") = CDbl(frmMantRecibos.lb_vcambio.Caption)
            .Parameters("@totsoles") = IIf(frmMantRecibos.CtrAyu_Moneda.xclave = "01", CDbl(frmMantRecibos.Txsoles.valor), Round(CDbl(frmMantRecibos.Txdolares.valor) * CDbl(frmMantRecibos.lb_vcambio.Caption), 2))
            .Parameters("@totdolares") = IIf(frmMantRecibos.CtrAyu_Moneda.xclave <> "01", CDbl(frmMantRecibos.Txdolares.valor), Round(CDbl(frmMantRecibos.Txsoles.valor) / CDbl(frmMantRecibos.lb_vcambio.Caption), 2))
            .Parameters("@fechadocumento") = frmMantRecibos.Dtp_FechaDoc.Value
            .Parameters("@observa") = ""
            .Parameters("@transferauto") = ""
            .Parameters("@numreciboegreso") = ""
            .Parameters("@usuario") = VGusuario
            .Parameters("@fechaact") = Now
            .Parameters("@empresa") = frmMantRecibos.Ctr_Ayuempresa.xclave
            .Execute
        End If
        If op = 4 Then
            .Parameters("@fechadocumento") = frmModrecibos.Dtp_FechaDoc.Value
            .Parameters("@usuario") = VGusuario
            .Parameters("@empresa") = frmModrecibos.Ctr_Ayuempresa.xclave
            .Execute
        End If
    End With
op = opaux
Set VGCommandoSP = New ADODB.Command
VGCommandoSP.ActiveConnection = VGgeneral
VGCommandoSP.CommandType = adCmdStoredProc
VGCommandoSP.CommandText = "te_abonadetalledocumento_pro"
VGCommandoSP.Parameters.Refresh
If op = 3 Then
   With VGCommandoSP
        .Parameters("@base") = VGParamSistem.BDEmpresa
        .Parameters("@tipo") = IIf(op = 3, 2, op)
        .Parameters("@numrecibo") = numero
        If op = 3 Then
            .Execute
        End If
   End With
Else
  rs.MoveFirst
  Item = 1
  rs.MoveFirst
  While Not rs.EOF
      With VGCommandoSP
        .Parameters("@base") = VGParamSistem.BDEmpresa
        .Parameters("@tipo") = IIf(op = 3, 2, op)
        .Parameters("@numrecibo") = numero
        If op = 3 Then
            .Execute
        End If
        If op = 2 Then
            If Item = 1 Then .Execute
            .Parameters("@tipo") = 1
            op = 1
        End If
        If op = 1 Then
            .Parameters("@estadoreg") = ""
            .Parameters("@item") = Item
            .Parameters("@emisioncheque") = "C"
            .Parameters("@tipodocconcepto") = rs.Fields("tipodoc_concepto")
            .Parameters("@numdocumento") = rs.Fields("numdocumento")
            .Parameters("@carabo") = "A"
            
            .Parameters("@formacan") = ""
            
            .Parameters("@tdqc") = rs.Fields("tdqc")
            .Parameters("@ndqc") = rs.Fields("ndqc")
            
            .Parameters("@tipocajabanco") = "C"
            
            .Parameters("@cajabanco") = rs.Fields("cajabanco1")
            .Parameters("@numctacte") = rs.Fields("numctacte")
            
            .Parameters("@adicionactacte") = "P"
            
            .Parameters("@monedadocumento") = frmMantRecibos.CtrAyu_Moneda.xclave
            .Parameters("@monedacancela") = rs.Fields("monedacancela")
            
            
            .Parameters("@importesoles") = rs.Fields("importesoles")
            .Parameters("@importedolares") = rs.Fields("importedolares")
            .Parameters("@contabledisponi") = "S"
            .Parameters("@fechacancela") = rs.Fields("fechacancela")
            .Parameters("@observacion") = rs.Fields("observacion")
            .Parameters("@gastos") = rs.Fields("gastos")
            .Parameters("@usuario") = VGusuario
            .Parameters("@fechaact") = rs.Fields("fechacancela")
            .Parameters("@entidad") = rs.Fields("entidad")
            .Parameters("@centrocosto") = rs.Fields("costos")
            .Execute
         End If
         If op = 4 Then
            .Parameters("@estadoreg") = ""
            .Parameters("@item") = Trim(Item)
            .Parameters("@tdqc") = rs.Fields("tdqc")
            .Parameters("@ndqc") = rs.Fields("ndqc")
            .Parameters("@fechacancela") = rs.Fields("fechacancela")
            .Parameters("@observacion") = rs.Fields("observacion")
            .Parameters("@gastos") = rs.Fields("gastos")
            .Parameters("@usuario") = VGusuario
            .Parameters("@fechaact") = Now
            .Parameters("@entidad") = rs.Fields("entidad")
            .Parameters("@centrocosto") = rs.Fields("costos")
            .Execute
         End If
    End With
    Item = Item + 1
    rs.MoveNext
  Wend
End If
Exit Sub

ErrorGrabaTesore:
    VGvarVerifica = False
    VGErrorString = "Error en Grabar en Cuentas por Pagar " & Chr(13) & Err.Description
End Sub
Public Sub MostrarCabecera(ByVal rs As ADODB.Fields, Optional optipo As Integer = 0)
Dim rsaux As ADODB.Fields
Dim pos As Integer
Set rsaux = rs
Set VGvardllgen = New dll_general
If optipo = 0 Then
With frmMantRecibos
     .lbNumComprobCab.Caption = rsaux!cabrec_numrecibo
     .CtrAyu_Proveedor.xclave = rsaux!clientecodigo: .CtrAyu_Proveedor.Ejecutar
     .CtrAyu_Moneda.xclave = rsaux!monedacodigo: .CtrAyu_Moneda.Ejecutar
     .CtrAyu_ModoOperacion.xclave = rsaux!operacioncodigo: .CtrAyu_ModoOperacion.Ejecutar
     .ChkCtaCte.Value = IIf(rsaux!controlctacte = 0 Or Trim(rsaux!controlctacte) = "", 0, 1)
     .ChkActCaja.Value = IIf(rsaux!controlctacte = 0 Or Trim(rsaux!controlctacte) = "", 0, 1)
   
   'La Serie y El Numero de Documento
     
     .TxingEgr.Text = rsaux!cabrec_ingsal
     .CtrAyu_Proveedor.Tag = Trim(rsaux!clientecodigo)
     .Dtp_FechaDoc.Value = rsaux!cabrec_fechadocumento
     .TxNref.Text = VGvardllgen.ESNULO(rsaux!cabrec_numreciboegreso, "")
     .TxTotBruto.Text = Format(rsaux!cabrec_totsoles, "###,###,###.00"): .TxTotBruto.valor = Format(rsaux!cabrec_totsoles, "#0.00")
     .TxTotIGV.Text = Format(rsaux!cabrec_totdolares, "###,###,###.00"): .TxTotIGV.valor = Format(rsaux!cabrec_totdolares, "#0.00")
     .lb_vcambio.Caption = Format(rsaux!cabrec_tipocambio, "#0.000")
        
     .Ctr_AyudaOficina.xclave = VGvardllgen.ESNULO(rsaux!vendedorcodigo, "00"): .Ctr_AyudaOficina.Ejecutar
     .Ctr_AyudaCaja.xclave = VGvardllgen.ESNULO(rsaux!cajacodigo, "00"): .Ctr_AyudaCaja.Ejecutar
     .Ctr_Ayuempresa.xclave = VGvardllgen.ESNULO(rsaux!empresacodigo, "00")
     If .Ctr_Ayuempresa.xclave <> "00" Then .Ctr_AyuCcosto.Filtro = "empresacodigo='" & .Ctr_Ayuempresa.xclave & "'"
  End With
End If
If optipo = 1 Then
With frmModrecibos
     .lbNumComprobCab.Caption = rsaux!cabrec_numrecibo
     .CtrAyu_Proveedor.xclave = rsaux!clientecodigo: .CtrAyu_Proveedor.Ejecutar
     .CtrAyu_Moneda.xclave = rsaux!monedacodigo: .CtrAyu_Moneda.Ejecutar
     .CtrAyu_ModoOperacion.xclave = rsaux!operacioncodigo: .CtrAyu_ModoOperacion.Ejecutar
     .ChkCtaCte.Value = IIf(rsaux!controlctacte = 0 Or Trim(rsaux!controlctacte) = "", 0, 1)
     .ChkActCaja.Value = IIf(rsaux!controlctacte = 0 Or Trim(rsaux!controlctacte) = "", 0, 1)
   
   'La Serie y El Numero de Documento
     
     .TxingEgr.Text = rsaux!cabrec_ingsal
     .CtrAyu_Proveedor.Tag = Trim(rsaux!clientecodigo)
     .Dtp_FechaDoc.Value = rsaux!cabrec_fechadocumento
     .TxNref.Text = VGvardllgen.ESNULO(rsaux!cabrec_numreciboegreso, "")
     .TxTotBruto.Text = Format(rsaux!cabrec_totsoles, "###,###,###.00"): .TxTotBruto.valor = Format(rsaux!cabrec_totsoles, "#0.00")
     .TxTotIGV.Text = Format(rsaux!cabrec_totdolares, "###,###,###.00"): .TxTotIGV.valor = Format(rsaux!cabrec_totdolares, "#0.00")
     .lb_vcambio.Caption = Format(rsaux!cabrec_tipocambio, "#0.000")
        
     .Ctr_AyudaOficina.xclave = VGvardllgen.ESNULO(rsaux!vendedorcodigo, "00"): .Ctr_AyudaOficina.Ejecutar
     .Ctr_AyudaCaja.xclave = VGvardllgen.ESNULO(rsaux!cajacodigo, "00"): .Ctr_AyudaCaja.Ejecutar
     .Ctr_Ayuempresa.xclave = VGvardllgen.ESNULO(rsaux!empresacodigo, "00")
     If .Ctr_Ayuempresa.xclave <> "00" Then .Ctr_AyuCcosto.Filtro = "empresacodigo='" & .Ctr_Ayuempresa.xclave & "'"
  End With
End If
End Sub
Public Sub MostrarDetalle(ByRef rs As ADODB.Recordset, Optional optipo As Integer = 0)
Dim rsaux As ADODB.Recordset, i As Integer
On Error GoTo err1
Set rsaux = New ADODB.Recordset
    Set VGCommandoSP = New ADODB.Command
    Set VGvardllgen = New dllgeneral.dll_general
    'Elimar los Detalle antes de grabar
    VGCommandoSP.ActiveConnection = VGgeneral
    VGCommandoSP.CommandType = adCmdStoredProc
    VGCommandoSP.CommandText = "te_abonadetalledocumento_pro"
    VGCommandoSP.Parameters.Refresh
    With VGCommandoSP
        .Parameters("@base") = VGParamSistem.BDEmpresa
        If optipo = 0 Then .Parameters("@numrecibo") = frmMantRecibos.lbNumComprobCab
        If optipo = 1 Then .Parameters("@numrecibo") = frmModrecibos.lbNumComprobCab
        .Parameters("@tipo") = 3
        Set rsaux = .Execute
    End With
    If rsaux.RecordCount() > 0 Then
       rsaux.MoveFirst
    End If
    While Not rsaux.EOF
        rs.AddNew
        For i = 0 To rsaux.Fields.Count - 1
            rs.Fields(rsaux.Fields(i).Name) = IIf(IsNull(rsaux.Fields(i).Value), "0", Trim(rsaux.Fields(i).Value))
        Next
        rs.Update
        rsaux.MoveNext
    Wend
    If rs.RecordCount > 0 Then
        rs.AbsolutePosition = 1
        If optipo = 0 Then
           frmMantRecibos.TDBG_Det.SetFocus
           frmMantRecibos.lbnregdetalle.Caption = Format(rs.RecordCount, "0 ")
         Else
           frmModrecibos.TDBG_Det.SetFocus
           frmModrecibos.lbnregdetalle.Caption = Format(rs.RecordCount, "0 ")
         End If
      Else
      If optipo = 0 Then frmMantRecibos.lbnregdetalle.Caption = Format(0, "0 ")
      If optipo = 1 Then frmModrecibos.lbnregdetalle.Caption = Format(0, "0 ")
    End If
    Exit Sub
err1:
Resume Next
End Sub
Public Function ValidarGrabarCabecera(NR As Long) As Boolean
ValidarGrabarCabecera = False
With frmModrecibos
    'Validando Que exista por lo menos un registro en el detalle
    Set VGvardllgen = New dllgeneral.dll_general
    If NR = 0 Then
        MsgBox "Por lo menos debe haber ingresado un registro de detalle", vbInformation
        .CtrAyu_ModoOperacion.SetFocus
        Exit Function
    End If
    If Trim(.CtrAyu_ModoOperacion.xclave) = "" Or .CtrAyu_ModoOperacion.xclave = "00" Then
        MsgBox "Tiene que ingresar un modo de compra", vbInformation
        .CtrAyu_ModoOperacion.SetFocus
        Exit Function
    End If
    If .CtrAyu_Proveedor.Visible = True Then
    If Trim(.CtrAyu_Proveedor.xclave) = "" Or .CtrAyu_Proveedor.xclave = "00" Then
        MsgBox "Tiene que ingresar un Proveedor", vbInformation
        .CtrAyu_Proveedor.SetFocus
        Exit Function
    End If
    End If
    If Trim(.Ctr_AyudaOficina.xclave) = "" Or .Ctr_AyudaOficina.xclave = "00" Then
        MsgBox "Tiene que ingresar un Codigo de Oficina", vbInformation
        .Ctr_AyudaOficina.SetFocus
        Exit Function
    End If
    If Trim(.TxingEgr.Text) = "" Then
        MsgBox "Tiene que ingresar Tipo de recibo ) I = Ingresos , E = Egresos ", vbInformation
        .TxingEgr.SetFocus
        Exit Function
    End If
    If Trim(.CtrAyu_Moneda.xclave) = "" Or .CtrAyu_Moneda.xclave = "00" Then
        MsgBox "Tiene que seleccionar una moneda ", vbInformation
        .CtrAyu_Moneda.SetFocus
        Exit Function
    End If
    If VGvardllgen.ESNULO(.lb_vcambio.Caption, 0) = 0 Then
        MsgBox "Debe de escoger una fecha que exista tipo de cambio", vbInformation
        Exit Function
    End If
    
    If .ChkActCaja = 1 And Trim(.Ctr_AyudaCaja.xclave) = "" Then
        MsgBox "No se ha ingresado codigo de caja , ingrese por favor ", vbExclamation
          .Ctr_AyudaCaja.SetFocus
          Exit Function
    
    End If
    If VGParametros.sistemamultiempresas = True And Trim(.Ctr_Ayuempresa.xclave) = "" Then
       MsgBox "Tiene que Seleccionar un codigo de empresa", vbInformation
       .Ctr_Ayuempresa.SetFocus
            Exit Function
        End If

    ValidarGrabarCabecera = True
End With
End Function
Public Function ValidarGrabarDetalle() As Boolean
    ValidarGrabarDetalle = False
    Set VGvardllgen = New dllgeneral.dll_general
With frmModrecibos
    If VGParametros.sistemactrlgastos Then
       If (.Ctr_Ayugastos.Enabled And .Ctr_Ayugastos.Visible) And (Trim(.Ctr_Ayugastos.xclave) = "" Or Trim(.Ctr_Ayugastos.xclave) = "00") Then
           MsgBox "Debe ingresar la Cuenta de gastos ", vbInformation
          .Ctr_Ayugastos.SetFocus
          Exit Function
       End If
    End If
    If CDbl(VGvardllgen.ESNULO(Espunto(.Txdolares.valor), 0)) = 0 And CDbl(VGvardllgen.ESNULO(Espunto(.Txsoles.valor), 0)) = 0 Then
        MsgBox "Por lo menos el Importe en Soles o el Importe En Dolares tiene que ser mayor a cero"
        .Txsoles.SetFocus
        Exit Function
    End If
    If (.Ctr_AyuCcosto.Enabled And .Ctr_AyuCcosto.Visible) And (Trim(.Ctr_AyuCcosto.xclave) = "") Then
        MsgBox "Debe ingresar El cento de costo", vbInformation
        .Ctr_AyuCcosto.SetFocus
        Exit Function
    End If
    ValidarGrabarDetalle = True
End With
End Function
Public Function ValidarRsDetalle(ByRef rs As Recordset) As Boolean
Dim doc As String, docaux As String
    Set VGvardllgen = New dllgeneral.dll_general
    ValidarRsDetalle = False
    rs.AbsolutePosition = 1
    While Not rs.EOF
        If Not ValidarGrabarDetalle Then Exit Function
        rs.MoveNext
    Wend
    ValidarRsDetalle = True
End Function
Public Function LimpiarCab()
    With frmMantRecibos
        .lbNumComprobCab.Caption = ""
        .CtrAyu_Proveedor.xclave = "": .CtrAyu_Proveedor.xnombre = ""
        .TxNref.Text = "": .Dtp_FechaDoc.Value = VGParamSistem.fechatrabajo
        .CtrAyu_Moneda.xclave = "": .CtrAyu_Moneda.xnombre = ""
        .CtrAyu_ModoOperacion.xclave = "": .CtrAyu_ModoOperacion.xnombre = ""
        .ChkCtaCte.Value = 0: .ChkRegComp.Value = 0: .ChkActCaja.Value = 0
        .Ctr_AyudaCaja.xclave = "": .Ctr_AyudaCaja.xnombre = ""
        .Ctr_AyudaOficina.xclave = "": .Ctr_AyudaOficina.xnombre = ""
        'Limpiar Totales
        .TxTotBruto.Text = "": .TxTotBruto.valor = ""
        .TxTotIGV.Text = "": .TxTotIGV.valor = ""
        .Ctr_Ayuempresa.xclave = "": .Ctr_Ayuempresa.xnombre = ""
    End With
End Function

