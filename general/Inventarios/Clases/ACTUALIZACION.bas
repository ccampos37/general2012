Attribute VB_Name = "ACTUALIZACION"

Sub ACTUALIZACION2001()
  
  If Not ExisteElem(0, VGCNx, "AL_CIERRESMENSUALES") Then
        SQL = " Create Table AL_CIERRESMENSUALES (CIERRMES Text(6),CIERRFECH DATETIME, CIERROPER TEXT(15) , " & _
        " CONSTRAINT Clave PRIMARY KEY (CIERRMES))"
        VGCNx.Execute SQL
  End If


End Sub
Public Function NumeroAuxiliar(mes As Integer, Optional ByRef numero As Long, Optional ByRef AnoProceso As String) As String
On Error GoTo Errnum
Dim rsaux As ADODB.Recordset
Dim cad As String
    Set rsaux = New ADODB.Recordset
    cad = "Select isnull(mes" & Trim(Format(mes, "00")) & ",0)+1 as numcorrelativo   From co_correlames " & _
          "Where Ano='" & AnoProceso & "'"
          
    rsaux.Open cad, VGCNx, adOpenKeyset, adLockReadOnly
               
    If rsaux.RecordCount > 0 Then
       NumeroAuxiliar = Trim(Format(rsaux!numcorrelativo, "00000"))
       numero = rsaux!numcorrelativo
       Else
        NumeroAuxiliar = "00"
        numero = 0
    End If
    Exit Function
Errnum:
    VGvarVerifica = False
    VGErrorString = "Error en Numero de Comprobante " & Chr(13) & Err.Description
End Function
Public Sub ActualizaCorrelComprob(ByVal numero As Double, ByVal Fecha As Date)
On Error GoTo Actualizacorre
    Set VGCommandoSP = New ADODB.Command
    VGCommandoSP.ActiveConnection = VGgeneral
    VGCommandoSP.CommandType = adCmdStoredProc
    VGCommandoSP.CommandText = "co_actcorraux_pro"
    VGCommandoSP.Parameters.Refresh
    With VGCommandoSP
        .Parameters("@Base") = VGCNx.DefaultDatabase
        .Parameters("@Ano") = LTrim(Year(Fecha))
        .Parameters("@Mes") = Format(Month(Fecha), "00")
        .Parameters("@numero") = numero
        .Execute
    End With
    Exit Sub
Actualizacorre:
    VGvarVerifica = False
    VGErrorString = "Error en Actualizar el Numero de Comprobante Auxiliar " & Chr(13) & Err.Description
End Sub

Public Sub ActualizaCorrelAuxiliar(ByVal numero As Double, ByVal Fecha As Date)
On Error GoTo Actualizacorre
    Set VGCommandoSP = New ADODB.Command
    VGCommandoSP.ActiveConnection = VGgeneral
    VGCommandoSP.CommandType = adCmdStoredProc
    VGCommandoSP.CommandText = "co_actcorraux_pro"
    VGCommandoSP.Parameters.Refresh
    With VGCommandoSP
        .Parameters("@Base") = VGCNx.DefaultDatabase
        .Parameters("@Ano") = LTrim(Year(Fecha))
        .Parameters("@Mes") = Format(Month(Fecha), "00")
        .Parameters("@numero") = numero
        .Execute
    End With
    Exit Sub
Actualizacorre:
    VGvarVerifica = False
    VGErrorString = "Error en Actualizar el Numero de Comprobante Auxiliar " & Chr(13) & Err.Description
End Sub
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
Public Function ValidarGrabarCabecera(NR As Long) As Boolean
ValidarGrabarCabecera = False
Dim montosoles As Double
With FrmValorizacionArticulos
    'Validando Que exista por lo menos un registro en el detalle
    Set VGvardllgen = New dllgeneral.dll_general
    If NR = 0 Then
        MsgBox "Por lo menos debe haber ingresado un registro de detalle", vbInformation
        FrmValorizacionArticulos.CtrAyu_Modoprovi.SetFocus
        Exit Function
    End If
    If Trim(.CtrAyu_Modoprovi.xclave) = "" Or .CtrAyu_Modoprovi.xclave = "00" Then
        MsgBox "Tiene que ingresar un modo de compra", vbInformation
        .CtrAyu_Modoprovi.SetFocus
        Exit Function
    End If
    If Trim(.CtrAyu_Proveedor.xclave) = "" Or .CtrAyu_Proveedor.xclave = "00" Then
        MsgBox "Tiene que ingresar un Proveedor", vbInformation
        .CtrAyu_Proveedor.SetFocus
        Exit Function
    End If
    If Trim(.CtrAyu_TipDoc.xclave) = "" Or .CtrAyu_TipDoc.xclave = "00" Then
        MsgBox "Tiene que ingresar un Tipo de Documento", vbInformation
        .CtrAyu_TipDoc.SetFocus
        Exit Function
    End If
    If Trim(.TxSerie.text) = "" Then
        MsgBox "Tiene que ingresar la Serie del Documento", vbInformation
        .TxSerie.SetFocus
        Exit Function
    End If
    If Trim(.TxNdoc.text) = "" Then
        MsgBox "Tiene que ingresar el Numero de Documento", vbInformation
        .TxNdoc.SetFocus
        Exit Function
    End If
    If Trim(.CtrAyu_moneda.xclave) = "" Or .CtrAyu_moneda.xclave = "00" Then
        MsgBox "Tiene que seleccionar una moneda ", vbInformation
        .CtrAyu_moneda.SetFocus
        Exit Function
    End If
    If Trim(.CtrAyu_TipCompra.xclave) = "" Or .CtrAyu_TipCompra.xclave = "00" Then
        MsgBox "Tiene que seleccionar una moneda ", vbInformation
        .CtrAyu_TipCompra.SetFocus
        Exit Function
    End If
    montosoles = VGvardllgen.ESNULO(.lb_vcambio.Caption, 0)
    If VGvardllgen.ESNULO(.lb_vcambio.Caption, 0) = 0 Then
        MsgBox "Debe de escoger una fecha que exista tipo de cambio", vbInformation
        Exit Function
    End If
    
'    If .totalcomprobante <> .TxTotImpCompra.text And .estadorendicion = 1 Then
'            MsgBox "Monto Modificado es diferente al Monto Original -->  " & .totalcomprobante & "  <--de la Rendicion Nro. " & .numerorendicion & " ,Verificar", vbInformation
'            .CtrAyu_Modoprovi.SetFocus
'            Exit Function
'        End If
    
    If Trim(.CtrAyu_moneda.xclave) = "02" Then
       montosoles = IIf(.TxTotTotal.valor = "", 0, .TxTotTotal.valor * montosoles)
     Else
       montosoles = IIf(.TxTotTotal.valor = "", 0, .TxTotTotal.valor)
    End If
    .Emiteretencion = "0"
    If Not .buencontribuyente = True And .comprainafecta = 0 Then
       If montosoles > VGParametros.minimoretencion And .ChkActCaja And .tipoinafecto <> 1 Then
           MsgBox "Monto es mayor a Minimo de retencion debe ser modo de compra PROVEEDORES", vbInformation
           .CtrAyu_Modoprovi.SetFocus
           Exit Function
         ElseIf montosoles > VGParametros.minimoretencion And .tipoinafecto <> 1 Then .Emiteretencion = "1"
       End If
    End If
    
    If .documentoinafecto = 0 Or .tipoinafecto = 1 Or .comprainafecta = 1 Then .Emiteretencion = "2"
    'Validando Que el Documento no se repita para un proveedor
    Dim DocAct As String
    DocAct = Trim(FrmValorizacionArticulos.CtrAyu_Proveedor.xclave) & "-" & Trim(FrmValorizacionArticulos.CtrAyu_TipDoc.xclave) & "-" & Trim(FrmValorizacionArticulos.TxSerie.text) & IIf(Trim(FrmValorizacionArticulos.TxSerie.text) = "", "", "-") & Trim(FrmValorizacionArticulos.TxNdoc.text)
    
    If (FrmValorizacionArticulos.IMant = 1) Or (FrmValorizacionArticulos.VlDocAnt <> DocAct) Then
       SQL = " Select * From dbo.co_cabeceraprovisiones Where isnull(proveedorcodigo,'')+'-'+isnull(documetocodigo,'')+'-'+cabprovinumdoc='" & DocAct & "'"
       If ExisteSQL(VGCNx, SQL) Then
           MsgBox "Este Documento ya ha sido ingresado para este proveedor ", vbExclamation
           FrmValorizacionArticulos.TxNdoc.SetFocus
           Exit Function
        End If
    End If
    If .ChkActCaja = 1 And Trim(.Ctr_AyudaCaja.xclave) = "" Then
        MsgBox "No se ha ingresado codigo de caja , ingrese por favor ", vbExclamation
          .Ctr_AyudaCaja.SetFocus
          Exit Function
    
    End If
    If .ChkActCaja = 1 And Trim(.CtrAyu_moneda.xclave) <> VGParametros.monedabase And Trim(VGParametros.monedabase) <> "" Then
        MsgBox "Documento en Moneda " & Trim(.CtrAyu_moneda.xnombre) & " y por Caja Chica, ingrese por favor Modo de Compra PROVEEDORES ", vbExclamation
          .Ctr_AyudaCaja.SetFocus
          Exit Function
    End If
    If VGParametros.sistemamultiempresas = True And Trim(.Ctr_Ayuempresa.xclave) = "" Then
       MsgBox "Tiene que Seleccionar un codigo de empresa", vbInformation
       .Ctr_Ayuempresa.SetFocus
            Exit Function
        End If
    If VGParametros.sistemabancarizacion = True And .ChkActCaja = 1 Then
       If .CtrAyu_moneda.xclave = "01" And .TxTotTotal.valor > VGParametros.sistemabancarizacion01 Then
           MsgBox "Tiene que Seleccionar Modo Proveedores por Bancarizacion de Soles Mayor a " & VGParametros.sistemabancarizacion01, vbInformation
          .CtrAyu_Modoprovi.SetFocus
            Exit Function
        End If
       If .CtrAyu_moneda.xclave = "02" And .TxTotTotal.valor > VGParametros.sistemabancarizacion02 Then
           MsgBox "Tiene que Seleccionar Modo Proveedores por Bancarizacion de dolares  Mayor a " & VGParametros.sistemabancarizacion02, vbInformation
          .CtrAyu_Modoprovi.SetFocus
            Exit Function
        End If
    End If
    ValidarGrabarCabecera = True
End With
End Function


Public Sub GrabarCabecera(ByVal op As Integer, Optional ByVal numero As Long, Optional ByVal NumeroAux As String, Optional ByVal Numerotesor As String)
On Error GoTo ErrorGrabaCabecera
    Set VGCommandoSP = New ADODB.Command
    Set VGvardllgen = New dllgeneral.dll_general
    VGCommandoSP.ActiveConnection = VGgeneral
    VGCommandoSP.CommandType = adCmdStoredProc
    VGCommandoSP.CommandText = "co_grabacabprovi"
    VGCommandoSP.Parameters.Refresh
     With VGCommandoSP
        .Parameters("@base") = VGParamSistem.BDEmpresa
        .Parameters("@tabla") = VGParamSistem.TablaCabcomprob
        .Parameters("@op") = op
        If Month(FrmValorizacionArticulos.DTPFechaContab) <> Val(VGParamSistem.mesproceso) Then
           .Parameters("@cabproviano") = Year(FrmValorizacionArticulos.DTPFechaContab)
           .Parameters("@cabprovimes") = Month(FrmValorizacionArticulos.DTPFechaContab)
         Else
            .Parameters("@cabproviano") = VGParamSistem.AnoProceso
            .Parameters("@cabprovimes") = Val(VGParamSistem.mesproceso)
        End If
        .Parameters("@cabprovinumero") = numero
        'Este para que al eliminar no utilizar estos parametros
        If op <= 2 Then
            .Parameters("@proveedorcodigo") = Trim(FrmValorizacionArticulos.CtrAyu_Proveedor.xclave)
            .Parameters("@cabprovirznsoc") = Left(Trim(FrmValorizacionArticulos.CtrAyu_Proveedor.xnombre), 50)
            .Parameters("@cabproviruc") = FrmValorizacionArticulos.TxRuc.text
            .Parameters("@monedacodigo") = FrmValorizacionArticulos.CtrAyu_moneda.xclave
            .Parameters("@modoprovicod") = FrmValorizacionArticulos.CtrAyu_Modoprovi.xclave
            .Parameters("@documetocodigo") = FrmValorizacionArticulos.CtrAyu_TipDoc.xclave
            .Parameters("@cabprovictacte") = FrmValorizacionArticulos.ChkCtaCte.Value
            .Parameters("@cabproviregcom") = FrmValorizacionArticulos.ChkRegComp.Value
            '.Parameters("@cabproviestado") = 0
            .Parameters("@cabprovinumdoc") = Trim(FrmValorizacionArticulos.TxSerie.text) & IIf(Trim(FrmValorizacionArticulos.TxSerie.text) = "", "", "-") & Trim(FrmValorizacionArticulos.TxNdoc.text)
            '.Parameters("@cabprovinumord") = 0
            .Parameters("@cabprovifchdoc") = FrmValorizacionArticulos.Dtp_FechaDoc.Value
            .Parameters("@cabprovifchven") = FrmValorizacionArticulos.DtpFech_Ven.Value
            '.Parameters("@cabproviitem") = 0
            .Parameters("@tipocompracodigo") = FrmValorizacionArticulos.CtrAyu_TipCompra.xclave
            .Parameters("@cabprovitotbru") = CDbl(FrmValorizacionArticulos.TxTotBruto.valor)
            '.Parameters("@cabprovitotdcto") = 0
            '.Parameters("@cabprovitotven") = 0
            .Parameters("@cabprovitotigv") = CDbl(FrmValorizacionArticulos.TxTotIGV.valor)
            .Parameters("@cabprovitotinaf") = CDbl(FrmValorizacionArticulos.TxTotInafecto.valor)
            .Parameters("@cabprovitotal") = CDbl(FrmValorizacionArticulos.TxTotTotal.valor)
            '.Parameters("@cabprovitotcxp") = 0
            '.Parameters("@cabprovitipigv") = 0
            .Parameters("@cabprovifchconta") = FrmValorizacionArticulos.Dtp_FechaDoc.Value
            '.Parameters("@cabprovifchcancel") = 0
            '.Parameters("@cabprovifnupol") = 0
            If CDbl(VGvardllgen.ESNULO(FrmValorizacionArticulos.lb_vcambio.Caption, 0)) = 0 Then
               .Parameters("@cabprovitipcambio") = 1
              Else
              .Parameters("@cabprovitipcambio") = CDbl(VGvardllgen.ESNULO(FrmValorizacionArticulos.lb_vcambio.Caption, 0))
            End If
            .Parameters("@cabprovinumaux") = NumeroAux
'            .Parameters("@cabprovinumtes") = 0
             .Parameters("@usuariocodigo") = VGusuario
            .Parameters("@fechaact") = Now
            '.Parameters("@tiposubasicodigo") = FrmValorizacionArticulos.CtrAyu_TipSubAsi.xclave
            .Parameters("@tiposubasicodigo") = "00"
            .Parameters("@cabprovitipdocref") = FrmValorizacionArticulos.TDBNota.Columns(0)
            .Parameters("@cabprovinref") = FrmValorizacionArticulos.TDBNota.Columns(1) + FrmValorizacionArticulos.TDBNota.Columns(2)
            .Parameters("@cabprovifechdocref") = FrmValorizacionArticulos.TDBNota.Columns(3)
            .Parameters("@cabproviopergrab") = FrmValorizacionArticulos.ChkOperGrab.Value
            '@cabprovioficina,@cabprovicaja,@cabprovifechcaja
            
            .Parameters("@cabprovioficina") = FrmValorizacionArticulos.Ctr_AyudaOficina.xclave
            .Parameters("@cabprovicaja") = IIf(FrmValorizacionArticulos.Ctr_AyudaCaja.Visible, FrmValorizacionArticulos.Ctr_AyudaCaja.xclave, "")
            .Parameters("@cabprovifechcaja") = IIf(FrmValorizacionArticulos.DTPFechaCaja.Visible, FrmValorizacionArticulos.DTPFechaCaja.Value, Null)
            If FrmValorizacionArticulos.IMant = 1 Then
                .Parameters("@cabproviflagmodi") = 0
              Else
                .Parameters("@cabproviflagmodi") = 1
            End If
            .Parameters("@cabprovinumtesor") = Trim(Numerotesor)
           If VGParametros.sistemamultiempresas = True Then
              .Parameters("@empresacodigo") = Trim(FrmValorizacionArticulos.Ctr_Ayuempresa.xclave)
            Else
            .Parameters("@empresacodigo") = "01"
           End If
            
            If op > 1 Then
              numero = Trim(FrmValorizacionArticulos.lbNumComprobCab)
             .Parameters("@cabprovinumero") = numero
            End If
        End If
    End With
    VGCommandoSP.Execute
    Exit Sub
ErrorGrabaCabecera:
    VGvarVerifica = False
  '  VGErrorString
    MsgBox ("Error en Grabar Cabecera " & Chr(13) & Err.Description)
    Exit Sub
    Resume
End Sub

Public Sub GrabarCP_Cargo(op As Integer, Optional numero As Long = 0)
On Error GoTo ErrorGrabaCP
Dim rsaux As ADODB.Recordset
Dim numerocomprobante As String
'@base, @tipo, @tabla, @tipodocu, @numero, @cliente, @vendedor, @zona,
'@apefecemi, @moneda, @apeimppag, @usuario, @tipocambio, @fechaact, @flagcancel, @cargoabono, @concepto
Set VGvardllgen = New dllgeneral.dll_general
    Set VGCommandoSP = New ADODB.Command
    Set VGvardllgen = New dllgeneral.dll_general
    VGCommandoSP.ActiveConnection = VGgeneral
    VGCommandoSP.CommandType = adCmdStoredProc
    VGCommandoSP.CommandText = "co_ingresacargo_pro"
    VGCommandoSP.Parameters.Refresh
    With VGCommandoSP
        .Parameters("@base") = VGParamSistem.BDEmpresa
        .Parameters("@tabla") = "CP_Cargo"
        .Parameters("@tipo") = op
        If op > 1 Then
            numero = Trim(FrmValorizacionArticulos.lbNumComprobCab)
        End If
        .Parameters("@abonotipoplanilla") = VGParametros.CpTiplan
        .Parameters("@abononumplanilla") = Format(numero, "000000")
        .Parameters("@cliente") = Trim(FrmValorizacionArticulos.CtrAyu_Proveedor.xclave)
        .Parameters("@tipodocu") = Trim(FrmValorizacionArticulos.CtrAyu_TipDoc.xclave)
        .Parameters("@numero") = Format(Trim(FrmValorizacionArticulos.TxSerie.text), "000") & Format(Left(Trim(FrmValorizacionArticulos.TxNdoc.text), 8), "00000000")
        'Este para que al eliminar no utilizar estos parametros
        If op <= 2 Then
            If op = 2 Then
               .Parameters("@oldnumero") = Format(Trim(FrmValorizacionArticulos.TxSerie.Tag), "000") & Format(Left(Trim(FrmValorizacionArticulos.TxNdoc.Tag), 8), "00000000")
               .Parameters("@oldtipodocu") = FrmValorizacionArticulos.CtrAyu_TipDoc.Tag
               .Parameters("@oldcliente") = FrmValorizacionArticulos.CtrAyu_Proveedor.Tag
            End If
            .Parameters("@vendedor") = VGParametros.CpOficina
            .Parameters("@zona") = Null
            .Parameters("@apefecemi") = FrmValorizacionArticulos.Dtp_FechaDoc.Value
            .Parameters("@moneda") = FrmValorizacionArticulos.CtrAyu_moneda.xclave
            .Parameters("@apeimppag") = CDbl(FrmValorizacionArticulos.TxTotTotal.valor)
            .Parameters("@usuario") = VGusuario
            .Parameters("@tipocambio") = CDbl(FrmValorizacionArticulos.lb_vcambio.Caption)
            .Parameters("@fechaact") = Now
            .Parameters("@flagcancel") = 0
            .Parameters("@cargoabono") = FrmValorizacionArticulos.VlDocNota
            .Parameters("@concepto") = "01"
            .Parameters("@glosa") = " Valorizacion de almacenes"
            .Parameters("@cargoapetiporefe") = ""
            .Parameters("@cargoapenrorefe") = ""
            .Parameters("@cargoapefecvct") = Format(FrmValorizacionArticulos.DtpFech_Ven.Value, "dd/mm/yyyy")
            .Parameters("@cargoemiteretencion") = FrmValorizacionArticulos.Emiteretencion
            .Parameters("@cargoemitedetraccion") = FrmValorizacionArticulos.emitedetraccion
           If VGParametros.sistemamultiempresas = True Then
              .Parameters("@empresacodigo") = Trim(FrmValorizacionArticulos.Ctr_Ayuempresa.xclave)
            Else
            .Parameters("@empresacodigo") = "01"
           End If
        End If
        .Execute
    End With
    Exit Sub
ErrorGrabaCP:
End Sub
Public Sub GrabarDetalle(ByVal rs As Recordset, Optional ByVal numero As Long = 0)
On Error GoTo ErrorGrabaDetalle
Dim rsaux As ADODB.Recordset
Dim numerocomprobante As String
Set VGvardllgen = New dllgeneral.dll_general
    Set rsaux = rs.Clone(adLockReadOnly)
 '   RSAUX.Filter = "(impbruto<>0 or impcompra <> 0)"
    Set VGCommandoSP = New ADODB.Command
    Set VGvardllgen = New dllgeneral.dll_general
    'Elimar los Detalle antes de grabar
    VGCommandoSP.ActiveConnection = VGgeneral
    VGCommandoSP.CommandType = adCmdStoredProc
    VGCommandoSP.CommandText = "co_grabadetprovi"
    VGCommandoSP.Parameters.Refresh
    With VGCommandoSP
        .Parameters("@base") = VGParamSistem.BDEmpresa
        .Parameters("@tabla") = VGParamSistem.tabladetcomprob
        .Parameters("@cabprovinumero") = numero
        If Month(FrmValorizacionArticulos.DTPFechaContab) <> Val(VGParamSistem.mesproceso) Then
           .Parameters("@cabproviano") = Year(FrmValorizacionArticulos.DTPFechaContab)
           .Parameters("@cabprovimes") = Month(FrmValorizacionArticulos.DTPFechaContab)
         Else
          .Parameters("@cabproviano") = VGParamSistem.AnoProceso
          .Parameters("@cabprovimes") = CInt(VGParamSistem.mesproceso)
        End If
        If VGParametros.sistemamultiempresas = True Then
              .Parameters("@empresa") = Trim(FrmValorizacionArticulos.Ctr_Ayuempresa.xclave)
            Else
            .Parameters("@empresa") = "01"
           End If
        .Parameters("@op") = 2
        .Execute
    End With
    rsaux.MoveFirst
    While Not rsaux.EOF
        With VGCommandoSP
            .Parameters("@cabprovinumero") = numero
            If Month(FrmValorizacionArticulos.DTPFechaContab) <> Val(VGParamSistem.mesproceso) Then
               .Parameters("@cabprovimes") = Month(FrmValorizacionArticulos.DTPFechaContab)
             Else
               .Parameters("@cabprovimes") = CInt(VGParamSistem.mesproceso)
            End If
            .Parameters("@op") = 1
            .Parameters("@detproviitem") = "001"  'RSAUX!item
            '.Parameters("@detprovicod1") = 0
            '.Parameters("@detprovicod2") = 0
            '.Parameters("@detprovicod3") = 0
            '.Parameters("@detprovicod4") = 0
            .Parameters("@gastoscodigo") = ESNULO(rsaux!gastos, "0213")
            .Parameters("@cuentacodigo") = "" 'RSAUX!Cuentacodigo
            .Parameters("@detprovimon") = FrmValorizacionArticulos.CtrAyu_moneda.xclave
            '.Parameters("@detproviestado") = 0
            .Parameters("@detproviimpbru") = rsaux!bruto
            .Parameters("@detproviimpigv") = rsaux!Igv
            .Parameters("@detproviimpina") = rsaux!inafecto
            .Parameters("@detprovitotal") = rsaux!Total
            '.Parameters("@detprovidscto") = 0
            '.Parameters("@detproviimpdct") = 0
            '.Parameters("@detproviimpven") = 0
            '.Parameters("@detproviigv") = 0
            .Parameters("@detproviformcamb") = IIf(FrmValorizacionArticulos.lb_vcambio.Visible, Format(FrmValorizacionArticulos.CmbTcambio.ListIndex + 1, "00"), "00")
            .Parameters("@detprovitipcam") = IIf(FrmValorizacionArticulos.lb_vcambio.Visible, CDbl(VGvardllgen.ESNULO(FrmValorizacionArticulos.lb_vcambio.Caption, 0)), 0)
            .Parameters("@usuariocodigo") = VGusuario
            .Parameters("@fechaact") = Now
            .Parameters("@detproviglosa") = " Valorizacion de almacenes " 'RSAUX!glosa
            .Parameters("@detproviccosto") = "00" 'Trim(VGvardllgen.ESNULO(RSAUX!cCosto, "00"))
            .Parameters("@analitico") = "00" ' Trim(VGvardllgen.ESNULO(RSAUX!analitico, "00"))
           .Execute
        End With
        rsaux.MoveNext
    Wend
    Exit Sub
ErrorGrabaDetalle:
    VGvarVerifica = False
    VGErrorString = "Error en Grabar Detalle " & Chr(13) & Err.Description
    MsgBox (VGErrorString)
    Exit Sub
    Resume
End Sub

Public Sub GeneraAsientoenLine(ByVal op As Integer, ByVal Nprovi As String, ByVal Comprob_Contable As String)
On Error GoTo genasiento
    Screen.MousePointer = 11
    'Generando los Analiticos que no Esten en contabilidad
    Set VGCommandoSP = New ADODB.Command
    Set VGvardllgen = New dllgeneral.dll_general
    VGCommandoSP.ActiveConnection = VGgeneral
    VGCommandoSP.CommandType = adCmdStoredProc
    VGCommandoSP.CommandText = "co_VerificaAnaliticoenLinea"
    VGCommandoSP.Parameters.Refresh
    With VGCommandoSP
        .Parameters("@BaseConta") = VGParamSistem.BDEmpresaCT
        .Parameters("@BaseCompra") = VGParamSistem.BDEmpresa
        .Parameters("@Ano") = VGParamSistem.AnoProceso
        .Parameters("@Mes") = VGParamSistem.mesproceso
        .Parameters("@tipanal") = VGParametros.xTipAnal
        .Parameters("@User") = VGParamSistem.Usuario
        .Parameters("@Nprovi") = Nprovi
        .Execute
    End With
    'Generando el Asiento en contabilidad
    Set VGCommandoSP = New ADODB.Command
    Set VGvardllgen = New dllgeneral.dll_general
    
    VGCommandoSP.ActiveConnection = VGgeneral
    VGCommandoSP.CommandType = adCmdStoredProc
    VGCommandoSP.CommandText = "co_generaasientolinea_pro"
    VGCommandoSP.Parameters.Refresh
    With VGCommandoSP
        .Parameters("@BaseConta") = VGParamSistem.BDEmpresaCT
        .Parameters("@BaseCompra") = VGParamSistem.BDEmpresa
        .Parameters("@SubAsiento") = VGParametros.xsubasiento
        .Parameters("@Libro") = VGParametros.xLibro
        If Month(FrmValorizacionArticulos.DTPFechaContab) <> Val(VGParamSistem.mesproceso) Then
           .Parameters("@mes") = Format(Month(FrmValorizacionArticulos.DTPFechaContab), "00")
         Else
           .Parameters("@mes") = Format(VGParamSistem.mesproceso, "00")
        End If
        .Parameters("@Ano") = VGParamSistem.AnoProceso
        .Parameters("@ctatotal") = VGParametros.xCtaTotal
        .Parameters("@ctaIGV") = VGParametros.xCtaIGV
        .Parameters("@ctaIES") = VGParametros.xCtaIES
        .Parameters("@ctaRTA") = VGParametros.xCtaRTA
        .Parameters("@tipanal") = VGParametros.xTipAnal
        .Parameters("@Compu") = VGComputer
        .Parameters("@Usuario") = VGParamSistem.Usuario
        .Parameters("@Oficina") = VGParametros.CpOficina
        .Parameters("@Nprovi") = Nprovi
        .Parameters("@op") = op
        .Parameters("@comprobconta") = Comprob_Contable
        .Execute
    End With
    
    'Actualizando las Glosas de Cabecera y Detalle
    Set VGCommandoSP = New ADODB.Command
    Set VGvardllgen = New dllgeneral.dll_general
    VGCommandoSP.ActiveConnection = VGgeneral
    VGCommandoSP.CommandType = adCmdStoredProc
    VGCommandoSP.CommandText = "co_GrabaGlosasProvisionLinea_pro"
    VGCommandoSP.Parameters.Refresh
    With VGCommandoSP
        .Parameters("@BaseConta") = VGParamSistem.BDEmpresaCT
        .Parameters("@BaseCompra") = VGParamSistem.BDEmpresa
        If Month(FrmValorizacionArticulos.DTPFechaContab) <> Val(VGParamSistem.mesproceso) Then
           .Parameters("@mes") = Format(Month(FrmValorizacionArticulos.DTPFechaContab), "00")
         Else
           .Parameters("@mes") = Format(VGParamSistem.mesproceso, "00")
        End If
        .Parameters("@Ano") = VGParamSistem.AnoProceso
        .Parameters("@Nprovi") = Nprovi
        .Execute
    End With
    
    'Actualizando los Registros que no se incluyen en el Reg. Compras
    Set VGCommandoSP = New ADODB.Command
    Set VGvardllgen = New dllgeneral.dll_general
    VGCommandoSP.ActiveConnection = VGgeneral
    VGCommandoSP.CommandType = adCmdStoredProc
    VGCommandoSP.CommandText = "co_RegComprasNoIncluyenenLinea_pro"
    VGCommandoSP.Parameters.Refresh
    With VGCommandoSP
        .Parameters("@BaseCompra") = VGParamSistem.BDEmpresa
        .Parameters("@BaseConta") = VGParamSistem.BDEmpresaCT
        .Parameters("@Asiento") = "081"
        If Month(FrmValorizacionArticulos.DTPFechaContab) <> Val(VGParamSistem.mesproceso) Then
           .Parameters("@mes") = Format(Month(FrmValorizacionArticulos.DTPFechaContab), "00")
         Else
           .Parameters("@mes") = Format(VGParamSistem.mesproceso, "00")
        End If
        .Parameters("@Ano") = VGParamSistem.AnoProceso
        .Parameters("@Nprovi") = Nprovi
        .Execute
    End With
    
    MsgBox "Se Realizo la Operacion Satisfactoriamente"
    Screen.MousePointer = 1
    Exit Sub
genasiento:
    Screen.MousePointer = 1
    VGvarVerifica = False
    VGErrorString = "Error en Grabar Cabecera " & Chr(13) & Err.Description
End Sub
Public Sub CargarParametrosCompras()
Dim rsaux As ADODB.Recordset
    
Set rsaux = New ADODB.Recordset
SQL = " select * from co_sistema"
Set rsaux = Nothing
Set rsaux = VGCNx.Execute(SQL)
If rsaux.RecordCount = 0 Then Exit Sub
Set VGvardllgen = New dllgeneral.dll_general

VGParametros.monedabase = Trim(rsaux!monedacodigo)
VGParametros.NomEmpresa = Trim(rsaux!sistemadescripcionempresa)
VGParametros.direccionempresa = Trim(rsaux!sistemadireccionempresa)
'VGparametros.RucEmpresa = Trim(RSAUX!sistemaempresaruc)
VGParametros.Igv = rsaux!sistemaigv
   
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
VGoficina = VGvardllgen.ESNULO(rsaux!sistemaoficina, "00")

VGParametros.xCtaTotal = rsaux!sistemactatotal
VGParametros.permite_tc = IIf(VGvardllgen.ESNULO(rsaux!permite_tc, 0) = 0, False, True)
VGParametros.sistemaactivaccostos = IIf(VGvardllgen.ESNULO(rsaux!sistemaactivaccostos, 0) = 0, False, True)
VGParametros.sistemaasientoenlinea = IIf(VGvardllgen.ESNULO(rsaux!sistemaasientoenlinea, 0) = 0, False, True)
VGParametros.sistemactrlgastos = IIf(VGvardllgen.ESNULO(rsaux!sistemactrlgastos, 0) = 0, False, True)
VGParametros.sistemamultiempresas = IIf(VGvardllgen.ESNULO(rsaux!sistemamultiempresas, 0) = 0, False, True)
VGParametros.minimoretencion = IIf(VGvardllgen.ESNULO(rsaux!sistemaminimoretencion, 0) = 0, 99999, rsaux!sistemaminimoretencion)
VGParametros.sistemabancarizacion = IIf(VGvardllgen.ESNULO(rsaux!bancarizacion, 0) = 0, 0, rsaux!bancarizacion)
VGParametros.sistemabancarizacion01 = IIf(VGvardllgen.ESNULO(rsaux!minimobancarizacion01, 0) = 0, 9999999, rsaux!minimobancarizacion01)
VGParametros.sistemabancarizacion02 = IIf(VGvardllgen.ESNULO(rsaux!minimobancarizacion02, 0) = 0, 9999999, rsaux!minimobancarizacion02)


SQL = " select * from vt_sistema"
Set rsaux = Nothing
Set rsaux = VGCNx.Execute(SQL)
If rsaux.RecordCount = 0 Then Exit Sub
VGParamSistem.tipoanaliticocodigo = rsaux!tipoanaliticocodigo
    
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
Set rsaux = VGCNx.Execute("select * from  al_sistema")
If rsaux.RecordCount = 0 Then
   VGParametros.tipocreacioncodigo = "L"
   VGParametros.tipogeneracioncodigo = 1
  Else
   VGParametros.tipocreacioncodigo = rsaux!tipocreacionarticulo
   VGParametros.tipogeneracioncodigo = rsaux!tipogeneracioncodigo
End If
End Sub

Public Function ValidarRsDetalle(ByRef rs As Recordset, ByRef rs1 As Recordset) As Boolean
Dim Doc As String, docaux As String
    Set VGvardllgen = New dllgeneral.dll_general
    rs.UpdateBatch adAffectAllChapters
    Doc = "select gastos,sum(impbruto) as bruto,sum(impigv) as igv,sum(impinafecto) as inafecto"
    Doc = Doc & ",sum(imptotal) as total from " & FrmValorizacionArticulos.sqltabla1 & " group by gastos "
    rs1.Open (Doc), VGCNx, adOpenDynamic, adLockBatchOptimistic
   If VGParametros.sistemactrlgastos Then
      With FrmValorizacionArticulos
           If .tipodetraccion = 1 Then .Emiteretencion = 2
           If .tipodetraccion = 1 Then .emitedetraccion = 1
         End With
    End If
    ValidarRsDetalle = True
End Function

