Attribute VB_Name = "FuncionesAlmacenes"
'***************************************************
'  Declaración API para Escribir y Leer un (*.INI)
'***************************************************

Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpAplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpAplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
        
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
        
Sub ACTUALIZACION2001()
  
  If Not ExisteElem(0, VGCNx, "AL_CIERRESMENSUALES") Then
        SQL = " Create Table AL_CIERRESMENSUALES (CIERRMES Text(6),CIERRFECH DATETIME, CIERROPER TEXT(15) , " & _
        " CONSTRAINT Clave PRIMARY KEY (CIERRMES))"
        VGCNx.Execute SQL
  End If


End Sub
Public Function NumeroAuxiliar(mes As Integer, Optional ByRef numero As Long, Optional ByRef AnoProceso As String) As String
On Error GoTo Errnum
Dim RSAUX As ADODB.Recordset
Dim cad As String
    Set RSAUX = New ADODB.Recordset
    cad = "Select isnull(mes" & Trim(Format(mes, "00")) & ",0)+1 as numcorrelativo   From co_correlames " & _
          "Where Ano='" & AnoProceso & "'"
          
    RSAUX.Open cad, VGCNx, adOpenKeyset, adLockReadOnly
               
    If RSAUX.RecordCount > 0 Then
       NumeroAuxiliar = Trim(Format(RSAUX!numcorrelativo, "00000"))
       numero = RSAUX!numcorrelativo
       Else
        NumeroAuxiliar = "00"
        numero = 0
    End If
    Exit Function
Errnum:
    VGvarVerifica = False
    vgerrorstring = "Error en Numero de Comprobante " & Chr(13) & Err.Description
End Function
Public Sub ActualizaCorrelComprob(ByVal numero As Double, ByVal Fecha As Date)
On Error GoTo Actualizacorre
    Set VGCommandoSP = New ADODB.Command
    VGCommandoSP.ActiveConnection = VGGeneral
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
    vgerrorstring = "Error en Actualizar el Numero de Comprobante Auxiliar " & Chr(13) & Err.Description
End Sub

Public Sub ActualizaCorrelAuxiliar(ByVal numero As Double, ByVal Fecha As Date)
On Error GoTo Actualizacorre
    Set VGCommandoSP = New ADODB.Command
    VGCommandoSP.ActiveConnection = VGGeneral
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
    vgerrorstring = "Error en Actualizar el Numero de Comprobante Auxiliar " & Chr(13) & Err.Description
End Sub
Public Function UltNumeroAuto(Tabla As String, op As String, cnx As ADODB.Connection) As Long
Dim RSAUX As ADODB.Recordset
On Error GoTo errornum
    Set RSAUX = New ADODB.Recordset
    Select Case op
        Case 1
'            rsaux.Open "SELECT Numx=isnull(IDENT_CURRENT('" & TABLA & "'),0)", cnx, adOpenKeyset, adLockReadOnly
            RSAUX.Open "SELECT top 1 Numx=isnull(cabprovinumero,1) from co_sistema ", cnx, adOpenKeyset, adLockReadOnly
    End Select
    If RSAUX.EOF Or RSAUX.BOF Then
      UltNumeroAuto = 1
      Exit Function
    Else
      UltNumeroAuto = RSAUX!Numx
      Set RSAUX = New ADODB.Recordset
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
    VGCommandoSP.ActiveConnection = VGGeneral
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
            .Parameters("@cabproviruc") = FrmValorizacionArticulos.txRuc.text
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
             .Parameters("@usuariocodigo") = VGUsuario
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
Dim RSAUX As ADODB.Recordset
Dim numerocomprobante As String
'@base, @tipo, @tabla, @tipodocu, @numero, @cliente, @vendedor, @zona,
'@apefecemi, @moneda, @apeimppag, @usuario, @tipocambio, @fechaact, @flagcancel, @cargoabono, @concepto
Set VGvardllgen = New dllgeneral.dll_general
    Set VGCommandoSP = New ADODB.Command
    Set VGvardllgen = New dllgeneral.dll_general
    VGCommandoSP.ActiveConnection = VGGeneral
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
            .Parameters("@usuario") = VGUsuario
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
Dim RSAUX As ADODB.Recordset
Dim numerocomprobante As String
Set VGvardllgen = New dllgeneral.dll_general
    Set RSAUX = rs.Clone(adLockReadOnly)
 '   RSAUX.Filter = "(impbruto<>0 or impcompra <> 0)"
    Set VGCommandoSP = New ADODB.Command
    Set VGvardllgen = New dllgeneral.dll_general
    'Elimar los Detalle antes de grabar
    VGCommandoSP.ActiveConnection = VGGeneral
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
    RSAUX.MoveFirst
    While Not RSAUX.EOF
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
            .Parameters("@gastoscodigo") = ESNULO(RSAUX!gastos, "0213")
            .Parameters("@cuentacodigo") = "" 'RSAUX!Cuentacodigo
            .Parameters("@detprovimon") = FrmValorizacionArticulos.CtrAyu_moneda.xclave
            '.Parameters("@detproviestado") = 0
            .Parameters("@detproviimpbru") = RSAUX!bruto
            .Parameters("@detproviimpigv") = RSAUX!Igv
            .Parameters("@detproviimpina") = RSAUX!inafecto
            .Parameters("@detprovitotal") = RSAUX!Total
            '.Parameters("@detprovidscto") = 0
            '.Parameters("@detproviimpdct") = 0
            '.Parameters("@detproviimpven") = 0
            '.Parameters("@detproviigv") = 0
            .Parameters("@detproviformcamb") = IIf(FrmValorizacionArticulos.lb_vcambio.Visible, Format(FrmValorizacionArticulos.CmbTcambio.ListIndex + 1, "00"), "00")
            .Parameters("@detprovitipcam") = IIf(FrmValorizacionArticulos.lb_vcambio.Visible, CDbl(VGvardllgen.ESNULO(FrmValorizacionArticulos.lb_vcambio.Caption, 0)), 0)
            .Parameters("@usuariocodigo") = VGUsuario
            .Parameters("@fechaact") = Now
            .Parameters("@detproviglosa") = " Valorizacion de almacenes " 'RSAUX!glosa
            .Parameters("@detproviccosto") = "00" 'Trim(VGvardllgen.ESNULO(RSAUX!cCosto, "00"))
            .Parameters("@analitico") = "00" ' Trim(VGvardllgen.ESNULO(RSAUX!analitico, "00"))
           .Execute
        End With
        RSAUX.MoveNext
    Wend
    Exit Sub
ErrorGrabaDetalle:
    VGvarVerifica = False
    vgerrorstring = "Error en Grabar Detalle " & Chr(13) & Err.Description
    MsgBox (vgerrorstring)
    Exit Sub
    Resume
End Sub

Public Sub GeneraAsientoenLine(ByVal op As Integer, ByVal Nprovi As String, ByVal Comprob_Contable As String)
On Error GoTo genasiento
    Screen.MousePointer = 11
    'Generando los Analiticos que no Esten en contabilidad
    Set VGCommandoSP = New ADODB.Command
    Set VGvardllgen = New dllgeneral.dll_general
    VGCommandoSP.ActiveConnection = VGGeneral
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
    
    VGCommandoSP.ActiveConnection = VGGeneral
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
        .Parameters("@Compu") = VGcomputer
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
    VGCommandoSP.ActiveConnection = VGGeneral
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
    VGCommandoSP.ActiveConnection = VGGeneral
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
    vgerrorstring = "Error en Grabar Cabecera " & Chr(13) & Err.Description
End Sub
Public Sub CargarParametros()
Dim RSAUX As ADODB.Recordset
    
Set RSAUX = New ADODB.Recordset
SQL = " select * from co_sistema"
Set RSAUX = Nothing
Set RSAUX = VGCNx.Execute(SQL)
If RSAUX.RecordCount = 0 Then Exit Sub
Set VGvardllgen = New dllgeneral.dll_general

VGParametros.monedabase = Trim(RSAUX!monedacodigo)
VGParametros.NomEmpresa = Trim(RSAUX!sistemadescripcionempresa)
VGParametros.direccionempresa = Trim(RSAUX!sistemadireccionempresa)
'VGparametros.RucEmpresa = Trim(RSAUX!sistemaempresaruc)
VGParametros.Igv = RSAUX!sistemaigv
   
'Parametros Exclusivos para la generacion de asientos a contabilidad
VGParametros.xLibro = VGvardllgen.ESNULO(RSAUX!sistemalibro, "")
VGParametros.xTipAnal = VGvardllgen.ESNULO(RSAUX!sistematipanal, "00")
VGParametros.xsubasiento = VGvardllgen.ESNULO(RSAUX!sistemasubasiento, "00")
VGParametros.xCtaIGV = VGvardllgen.ESNULO(RSAUX!sistemactaIGV, "00")
VGParametros.xCtaIES = VGvardllgen.ESNULO(RSAUX!sistemactaIES, "00")
VGParametros.xCtaRTA = VGvardllgen.ESNULO(RSAUX!sistemactaRTA, "00")
VGParametros.auxaut = True ' Se tiene que crear el campo para controlar auxiliar automatico
    
'Cargar parametros para pasar a cuentas por cobrar
VGParametros.CpTiplan = VGvardllgen.ESNULO(RSAUX!sistematipoplan, "00")
VGParametros.CpOficina = VGvardllgen.ESNULO(RSAUX!sistemaoficina, "00")
VGoficina = VGvardllgen.ESNULO(RSAUX!sistemaoficina, "00")

VGParametros.xCtaTotal = RSAUX!sistemactatotal
VGParametros.permite_tc = IIf(VGvardllgen.ESNULO(RSAUX!permite_tc, 0) = 0, False, True)
VGParametros.sistemaactivaccostos = IIf(VGvardllgen.ESNULO(RSAUX!sistemaactivaccostos, 0) = 0, False, True)
VGParametros.sistemaasientoenlinea = IIf(VGvardllgen.ESNULO(RSAUX!sistemaasientoenlinea, 0) = 0, False, True)
VGParametros.sistemactrlgastos = IIf(VGvardllgen.ESNULO(RSAUX!sistemactrlgastos, 0) = 0, False, True)
VGParametros.sistemamultiempresas = IIf(VGvardllgen.ESNULO(RSAUX!sistemamultiempresas, 0) = 0, False, True)
VGParametros.minimoretencion = IIf(VGvardllgen.ESNULO(RSAUX!sistemaminimoretencion, 0) = 0, 99999, RSAUX!sistemaminimoretencion)
VGParametros.sistemabancarizacion = IIf(VGvardllgen.ESNULO(RSAUX!bancarizacion, 0) = 0, 0, RSAUX!bancarizacion)
VGParametros.sistemabancarizacion01 = IIf(VGvardllgen.ESNULO(RSAUX!minimobancarizacion01, 0) = 0, 9999999, RSAUX!minimobancarizacion01)
VGParametros.sistemabancarizacion02 = IIf(VGvardllgen.ESNULO(RSAUX!minimobancarizacion02, 0) = 0, 9999999, RSAUX!minimobancarizacion02)


SQL = " select * from vt_sistema"
Set RSAUX = Nothing
Set RSAUX = VGCNx.Execute(SQL)
If RSAUX.RecordCount = 0 Then Exit Sub
VGParamSistem.tipoanaliticocodigo = RSAUX!tipoanaliticocodigo
    
Set RSAUX = New ADODB.Recordset
RSAUX.Open "select sistemaultimonivel,sistemaultimonivelcostos from  ct_sistema", VGCnxCT, adOpenKeyset, adLockReadOnly
If RSAUX.RecordCount = 0 Then Exit Sub
VGnumniveles = RSAUX!sistemaultimonivel
VGnumnivcos = ESNULO(RSAUX!sistemaultimonivelcostos, 1)

Set RSAUX = New ADODB.Recordset
RSAUX.Open "select sistemaultimonivel from  co_sistema", VGCNx, adOpenKeyset, adLockReadOnly
If RSAUX.RecordCount = 0 Then Exit Sub
VGnumnivgas = RSAUX!sistemaultimonivel
  
Set RSAUX = New ADODB.Recordset
Set RSAUX = VGCNx.Execute("select * from  al_sistema")
If RSAUX.RecordCount = 0 Then
   VGParametros.tipocreacioncodigo = "L"
   VGParametros.tipogeneracioncodigo = 1
  Else
   VGParametros.tipocreacioncodigo = RSAUX!tipocreacionarticulo
   VGParametros.tipogeneracioncodigo = RSAUX!tipogeneracioncodigo
End If
End Sub

Public Function ValidarRsDetalle(ByRef rs As Recordset, ByRef rs1 As Recordset) As Boolean
Dim doc As String, docaux As String
    Set VGvardllgen = New dllgeneral.dll_general
    rs.UpdateBatch adAffectAllChapters
    doc = "select gastos,sum(impbruto) as bruto,sum(impigv) as igv,sum(impinafecto) as inafecto"
    doc = doc & ",sum(imptotal) as total from " & FrmValorizacionArticulos.sqltabla1 & " group by gastos "
    rs1.Open (doc), VGCNx, adOpenDynamic, adLockBatchOptimistic
   If VGParametros.sistemactrlgastos Then
      With FrmValorizacionArticulos
           If .tipodetraccion = 1 Then .Emiteretencion = 2
           If .tipodetraccion = 1 Then .emitedetraccion = 1
         End With
    End If
    ValidarRsDetalle = True
End Function
Public Sub Init_ControlDBGrid(EsteGrid As DBGrid)
 With EsteGrid
      .MarqueeStyle = dbgHighlightRow
 End With
End Sub
Public Sub AlinearAyuda(f As Form)
f.Left = MDIPrincipal.Left + MDIPrincipal.Width - f.Width
' f.Top = MDIPrincipal.Height - MDIPrincipal.ScaleHeight
f.Top = (Screen.Height - f.Height) / 2
End Sub

Public Sub AlinearFrm(f As Form)
 f.Left = MDIPrincipal.Left + 50
 f.Top = MDIPrincipal.Top + 50
End Sub

Function ValidFecha(vText As String) As String
Dim cTxtNew As String, ncnt As Integer
Dim cTxt As String, cTxtDig As String

cTxtDig = "": cTxtNew = ""
For ncnt = 1 To Len(vText)
      cTxt = Mid(vText, ncnt, 1)
      If cTxt = "/" Then
         cTxtNew = cTxtNew & Str(Val(cTxtDig)) & "/"
         cTxtDig = ""
      Else
         If cTxt <> "_" Then cTxtDig = cTxtDig & cTxt
      End If
Next
If cTxtDig <> "" Then cTxtNew = cTxtNew & Str(Val(cTxtDig))

If IsDate(cTxtNew) Then
   ValidFecha = Format(CDate(cTxtNew), "dd/mm/yyyy")
End If
End Function

Function ValidHora(vText As String) As String
Dim cTxtNew As String

cTxtNew = "01/01/74 " & vText
If IsDate(cTxtNew) Then
   ValidHora = Format(CDate(cTxtNew), "hh:mm")
Else
    ValidHora = "00:00"
End If
End Function

Function FValidFec(vText As String) As Boolean
Dim cTxtNew As String, ncnt As Integer
Dim cTxt As String, cTxtDig As String

If Day(vText) = Null Then
   FValidFec = False
Else
  If IsNull(Day(CDate(vText))) Then
     FValidFec = False
    Exit Function
  End If
  FValidFec = True
End If
'If IsDate(cTxtNew) Then
'   ValidFecha = Format(CDate(cTxtNew), "dd/mm/yyyy")
'Else
'
'End If
End Function


Public Function codigo(cCod As String) As Boolean  ' Codigo del ARTICULO
Dim cSelC As ADODB.Recordset, csql As String
If Trim(cCod) = "" Then
    codigo = False
    Exit Function
End If
csql = "Select ACODIGO,adescri from MaeART where ACODIGO = '" & SupCadSQL(Trim(cCod)) & "'"
Set cSelC = New ADODB.Recordset
cSelC.Open csql, VGCNx, adOpenStatic
If cSelC.RecordCount > 0 Then
    codigo = False: cSelC.Close
    Exit Function
End If
codigo = True: cSelC.Close
End Function

Public Function Codigo2(cCod As String) As Boolean  'Codigo del FABRICANTE
Dim cSelC As ADODB.Recordset, csql As String
If Trim(cCod) = "" Then
    MsgBox "Falta Codigo", vbInformation, "Mensaje"
    Codigo2 = False
    Exit Function
End If
csql = "Select ACODIGO2 from MaeART where ACODIGO2 = '" & SupCadSQL(Trim(cCod)) & "'"
Set cSelC = New ADODB.Recordset
cSelC.Open csql, VGCNx, adOpenStatic
If cSelC.RecordCount > 0 Then
    Codigo2 = False: cSelC.Close
    Exit Function
End If
Codigo2 = True: cSelC.Close
End Function

Public Function CodigoC(cCod As String) As Boolean     'Codigo del Cliente
Dim cSelC As ADODB.Recordset, csql As String
If Trim(cCod) = "" Then
    MsgBox "Falta Codigo", vbInformation, "Mensaje"
    CodigoC = False
    Exit Function
End If
csql = "Select CLIENTECODIGO from VT_CLIENTE where CLIENTECODIGO = '" & SupCadSQL(cCod) & "'"
Set cSelC = New ADODB.Recordset
cSelC.Open csql, VGCNx, adOpenStatic
If cSelC.RecordCount > 0 Then
    CodigoC = False: cSelC.Close
    Exit Function
End If
CodigoC = True: cSelC.Close
End Function


Public Sub Ubi_Tab(oT As CrystalReport)
Dim nI As Integer, nN As Integer
'nN = oT.RetrieveDataFiles
For nI = 0 To nN
    If InStr(UCase(oT.DataFiles(nI)), "BDCOMUN") > 0 Then
        oT.DataFiles(nI) = cRuta2
    
    ElseIf InStr(UCase(oT.DataFiles(nI)), "BDAUXCOM") > 0 Then
        oT.DataFiles(nI) = App.Path & "\BDAUXCOM.MDB"
        
    ElseIf InStr(UCase(oT.DataFiles(nI)), "BDWENCO") > 0 Then     'Configuración
        oT.DataFiles(nI) = sName & "\BDWENCO.MDB"
    ElseIf InStr(UCase(oT.DataFiles(nI)), VGNameCont) > 0 Then
        If UCase(Dir$(cRuta4)) = UCase(VGNameCont & ".MDB") Then
             oT.DataFiles(nI) = VGParamSistem.RutaReport & "\" & VGNameCont & ".MDB"
        Else
             oT.DataFiles(nI) = cRuta2
        End If
    End If
Next nI
End Sub

'Posicionar la barra en el DataGrid
Public Function Pos_Dato(Adc As ADODB.Recordset) As Integer
Dim nN As Integer

Adc.MoveNext
If Not Adc.EOF Then
          nN = Adc.Bookmark - 1
Else
    Adc.MovePrevious
    Adc.MovePrevious
    If Not Adc.BOF Then
          nN = Adc.Bookmark
    End If
End If

Pos_Dato = nN
End Function

Public Function Pos_Dato1(Adc1 As Recordset, cCampo As String) As String
Dim cCodigo As String
Dim cCodigo1 As String

        
    cCodigo = Adc1(cCampo)
    Adc1.Delete
    Adc1.MoveNext
    If Adc1.EOF Then
       Adc1.MoveFirst
       If Adc1.BOF Then
       Else
         cCodigo1 = Adc1(cCampo)
       End If
    Else
      cCodigo1 = Adc1(cCampo)
    End If
 Pos_Dato1 = cCodigo1
End Function

Public Function fFam(cFam As String) As String        'FAMILIA
Dim cSqlA As String, cSelA As ADODB.Recordset
If Trim(cFam) = "" Then
    fFam = ""
    Exit Function
End If
cSqlA = "Select * FROM FAMILIA WHERE FAM_CODIGO = '" & Trim(cFam) & "' ORDER BY FAM_CODIGO "
Set cSelA = New ADODB.Recordset
cSelA.Open cSqlA, VGCNx, adOpenStatic
If cSelA.RecordCount = 0 Then
    fFam = "": cSelA.Close
    Exit Function
Else
    fFam = cSelA("FAM_NOMBRE")
End If
cSelA.Close
End Function

Public Function Val_Ayu(cAyu As String, cCodayu As String) As String
Dim cSqlA As String, cSelA As ADODB.Recordset

If Trim(cAyu) = "" Then
    Val_Ayu = ""
    Exit Function
End If

cSqlA = "Select * FROM TABAYU WHERE TCOD='" & cCodayu & "' And tClave = '" & Trim(cAyu) & "' ORDER BY TCLAVE "
Set cSelA = New ADODB.Recordset
cSelA.Open cSqlA, VGCNx, adOpenStatic
If cSelA.RecordCount = 0 Then
    Val_Ayu = "": cSelA.Close
    Exit Function
Else
    Val_Ayu = cSelA("tdescri")
End If
cSelA.Close
End Function

Public Function fPre(cPre As String) As String 'Precio
Dim cSqlA As String, cSelA As ADODB.Recordset
If Trim(cPre) = "" Then
    fPre = ""
    Exit Function
End If
cSqlA = "SELECT Cod_LisPre,Des_LisPre FROM TIPO_PRECIO where Cod_LisPre= '" & Trim(cPre) & "' ORDER BY Cod_LisPre"
Set cSelA = New ADODB.Recordset
cSelA.Open cSqlA, VGCNx, adOpenStatic
If cSelA.RecordCount = 0 Then
    fPre = "": cSelA.Close
    Exit Function
Else
    fPre = cSelA("Des_LisPre")
End If
cSelA.Close
End Function


Public Function fDis(cDis As String) As String 'Distrito
Dim cSqlA As String, cSelA As ADODB.Recordset
If Trim(cDis) = "" Then
    fDis = ""
    Exit Function
End If
cSqlA = "Select * FROM TABAYU WHERE TCOD='13' And tClave = '" & Trim(cDis) & "' ORDER BY TCLAVE "
Set cSelA = New ADODB.Recordset
cSelA.Open cSqlA, VGCNx, adOpenStatic
If cSelA.RecordCount = 0 Then
    fDis = "": cSelA.Close
    Exit Function
Else
    fDis = cSelA("tdescri")
End If
cSelA.Close
End Function

Public Function fGir(cGir As String) As String 'Giro
Dim cSqlA As String, cSelA As ADODB.Recordset
If Trim(cGir) = "" Then
    fGir = ""
    Exit Function
End If
cSqlA = "Select * FROM TABAYU WHERE TCOD='62' And tClave = '" & Trim(cGir) & "' ORDER BY TCLAVE "
Set cSelA = New ADODB.Recordset
cSelA.Open cSqlA, VGCNx, adOpenStatic
If cSelA.RecordCount = 0 Then
    fGir = "": cSelA.Close
    Exit Function
Else
    fGir = cSelA("tdescri")
End If
cSelA.Close
End Function

Public Function fMost(cMost As String) As String                        'corregir
Dim cSqlA As String, cSelA As ADODB.Recordset
If Trim(cMost) = "" Then
    fMost = ""
    Exit Function
End If
cSqlA = "SELECT f.FAM_NOMBRE, l.LIN_NOMBRE"
cSqlA = cSqlA & " FROM MAEART m INNER JOIN (FAMILIA f INNER JOIN LINEAS l ON f.FAM_CODIGO=l.FAM_CODIGO) ON (m.AFAMILIA = f.FAM_CODIGO) AND (m.AMODELO=l.LIN_CODIGO)"
cSqlA = cSqlA & " WHERE m.ACODIGO='7895'"
Set cSelA = New ADODB.Recordset
cSelA.Open cSqlA, VGCNx, adOpenStatic
If cSelA.RecordCount = 0 Then
    fMost = "": cSelA.Close
    Exit Function
Else
    fMost = cSelA("f.FAM_NOMBRE") + ", " + cSelA("l.LIN_NOMBRE") + ", " + cSelA("g.GRU_NOMBRE") + ", " + cSelA("m.ACOLOR") + ", " + cSelA("m.AMARCA")
End If
cSelA.Close
End Function

Public Function NumPto(cKey As Integer) As Boolean
If (cKey < 48 Or cKey > 57) And cKey <> 46 And cKey <> 13 And cKey <> 8 Then
    NumPto = False
Else
    NumPto = True
End If
End Function

'Numeros sin pto. decimal
Public Function NumSpto(cKey As Integer) As Boolean
If (cKey < 48 Or cKey > 57) And cKey <> 13 And cKey <> 8 Then
    NumSpto = False
Else
    NumSpto = True
End If
End Function
Public Function fEqui(cEqui As String, cUni) As String       'UNIDAD DE EQUIVALENCIA
Dim cSqlA As String, cSelA As ADODB.Recordset
If Trim(cEqui) = "" Then
    fEqui = ""
    Exit Function
End If
cSqlA = "Select * FROM TABEQUI WHERE EQUNIEQUI = '" & Trim(cEqui) & "' AND EQUNIPRI = '" & Trim(cUni) & "' ORDER BY EQUNIEQUI"
Set cSelA = New ADODB.Recordset
cSelA.Open cSqlA, VGCNx, adOpenStatic
If cSelA.RecordCount = 0 Then
    fEqui = "": cSelA.Close
    Exit Function
Else
    fEqui = "WWWW"
End If
cSelA.Close
End Function

Public Function Last_Day(mes As Integer, Aa As Integer) As Integer
Dim dia As Integer
Last_Day = 0
 If mes > 0 And mes < 13 Then
  If Aa > 1000 Then
    Select Case mes
     Case 1, 3, 5, 7, 8, 10, 12:
        dia = 31
     Case 4, 6, 9, 11:
        dia = 30
     Case 2:
        If (Aa Mod 4) = 0 Then
         dia = 29
        Else
         dia = 28
        End If
    End Select
    Last_Day = dia
  End If
End If
End Function


Public Function prove(txt As TextBox) As String
 Dim rs As New ADODB.Recordset
 Dim RSQL As String
   RSQL = "select PRVCNOMBRE FROM maeprov where PRVCCODIGO= '" & txt & "'" '
   
   'Set RS = VGBaseDatos.OpenRecordset(RSQL, dbOpenSnapshot)
   Set rs = VGCNx.Execute(RSQL)
   If Not rs.EOF Then
      prove = rs(0)
   Else
     MsgBox "El codigo del proveedor no existe !", vbExclamation, "Error"
     prove = ""
  End If
  rs.Close
End Function


Public Sub HabilitarMenu_Usuarios(Optional tipo As String, Optional Emp As String, Optional Nivel As String)
Dim ADOMenu As ADODB.Recordset
Dim ADOUsuA As ADODB.Recordset
Dim nArr() As Boolean
Dim nNum As Integer
Dim nUn As Integer

Set ADOMenu = New ADODB.Recordset
Set ADOUsuA = New ADODB.Recordset
If Not vGAdmLog Then
  ADOUsuA.Open "Select * From Men_Usu_Inv Where usuariocodigo = '" & tipo & "' and EMP_CODIGO = '" & Emp & "' order by MEN_CODIGO", VGConfig, adOpenStatic
End If
ADOMenu.Open "Select * From Menu_Inv Order by Men_Codigo", VGConfig, adOpenStatic

nNum = 67  'Contiene la cantidad de opciones en el Menu (Si hay cambios aumentar o disminuir en número)

If nNum > ADOMenu.RecordCount Or nNum < ADOMenu.RecordCount Then
   ' MsgBox "La cantidad de opciones registradas en el programa no es igual a las de la tabla", vbInformation, "Verificar"
    'ADOUsuA.Close: ADOMenu.Close: Exit Sub
End If

ReDim nArr(1 To ADOMenu.RecordCount, 1 To 2)

For nUn = 1 To ADOMenu.RecordCount
            nArr(nUn, 1) = False
             nArr(nUn, 2) = False
Next nUn

If Nivel = "A1" Then   'administrador solo conf archivo salir
        nUn = 1
        Do While Not ADOMenu.EOF
                If ADOMenu("Men_Codigo") = "01" Or ADOMenu("Men_Codigo") = "0109" Or ADOMenu("Men_Codigo") = "07" Or ADOMenu("Men_Codigo") = "0701" _
                                    Or ADOMenu("Men_Codigo") = "070101" Or ADOMenu("Men_Codigo") = "070102" Or ADOMenu("Men_Codigo") = "070103" Or ADOMenu("Men_Codigo") = "0703" Then
                                            nArr(nUn, 1) = True
                                            nArr(nUn, 2) = ADOMenu("Men_Visible")
                Else
                                            nArr(nUn, 1) = False
                                            nArr(nUn, 2) = ADOMenu("Men_Visible")
                End If
                nUn = nUn + 1
                ADOMenu.MoveNext
                If ADOMenu.EOF Then Exit Do
        Loop
        ADOMenu.Close
        If ADOUsuA.State <> 0 Then
          ADOUsuA.Close
       End If
ElseIf Nivel = "A2" Or Nivel = "M" Then   'adminnistrador todas las opciones
        nUn = 1
        ADOMenu.MoveFirst
        Do While Not ADOMenu.EOF
                 nArr(nUn, 1) = True
                 nArr(nUn, 2) = ADOMenu("Men_Visible")
                 ADOMenu.MoveNext
                 If ADOMenu.EOF Then Exit Do
                 nUn = nUn + 1
        Loop
Else
        nUn = 1
        If ADOMenu.RecordCount > 0 Then
                Do While Not ADOMenu.EOF
                            If ADOUsuA.RecordCount > 0 Then
                                    ADOUsuA.MoveFirst
                                    ADOUsuA.Filter = "Men_Codigo = '" & ADOMenu("Men_Codigo") & "'"
                                    If Not ADOUsuA.EOF Then
                                            nArr(nUn, 1) = ADOUsuA("Men_Hab")
                                            nArr(nUn, 2) = ADOMenu("Men_Visible")
                                    Else
                                            nArr(nUn, 1) = False
                                            nArr(nUn, 2) = ADOMenu("Men_Visible")
                                    End If
                                    ADOUsuA.Filter = ""
                            Else
                                    nArr(nUn, 1) = False
                                    nArr(nUn, 2) = ADOMenu("Men_Visible")
                            End If
                            ADOMenu.MoveNext
                            If ADOMenu.EOF Then Exit Do
                            nUn = nUn + 1
                Loop
                
        Else
                For nUn = 1 To ADOMenu.RecordCount
                        nArr(nUn, 1) = False
                        nArr(nUn, 2) = ADOMenu("Men_Visible")
            Next nUn
        End If
        ADOMenu.Close
        If ADOUsuA.State <> 0 Then
           ADOUsuA.Close
       End If
End If

End Sub
Public Function FechS(Fecha As Variant, tipo As TIPFECHA) As Variant
Dim h As Date
Dim fechaAux As Double
On Error GoTo ErrorFecha
   h = CDate(Fecha)
   Select Case tipo
      Case Sqlf: 'Para transformar al sql
         fechaAux = DateSerial(Year(Fecha), Month(Fecha), Day(Fecha)) - 2
      Case Adof: 'Para transformar al ado
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


Function ExisteTabla(ByVal NombreTabla As String, ByVal ADOConnection As ADODB.Connection) As Boolean
    Dim RsTbls As New ADODB.Recordset
    Set RsTbls = ADOConnection.OpenSchema(adSchemaTables)
    RsTbls.Find "[Table_Name]='" & NombreTabla & "'"
    If RsTbls.EOF Then ExisteTabla = False Else ExisteTabla = True
End Function

Public Function AGREGARBASE(codigo As String) As Boolean
On Error GoTo errfil
    AGREGARBASE = True
    Screen.MousePointer = 11
    If UCase(Dir$(sName & "Data\" & codigo & "\" & "BdComun.mdb")) <> "BDCOMUN.MDB" Then
        FileCopy sName & "Bdplanti.mdk", sName & "Data\" & codigo & "\" & "BDComun.mdb"
        FileCopy sName & "BdTransf.mdk", sName & "Data\" & codigo & "\" & "BDTransf.mdb"
        MsgBox "Proceso Terminado", vbInformation, "Información"
    Else
        MsgBox "Proceso ya ha sido Generado", vbInformation, "Información"
    End If
    Screen.MousePointer = 1
Exit Function
errfil:
    AGREGARBASE = False
    Select Case Err.Number
        Case 53
            MsgBox "No se encontraron las plantillas(BDPLANTI ó BDTRANSF) en la ruta especificada en archivo Invetarios.Ini" & Chr(13) & _
                   "Ruta especificada: " & sName
        Case 76
            MsgBox "No se encuentra la carpeta """ & codigo & """ de la empresa especificada" & Chr(13) & _
                   "en la ruta: " & sName & "DATA\"
        Case Else
           MsgBox Err.Description
    End Select
End Function
Public Sub ActualizaBD()
 Dim SQL As String
 
  If Not ExisteElem(0, VGCNx, "KITS") Then
      SQL = " Create Table KITS (CODART Text(20),CODKIT Text(20), " & _
      "  CANART double)"
      VGCNx.Execute SQL
  End If
  If Not ExisteElem(1, VGCNx, "MAEART", "AMARCA") Then
      VGCNx.Execute "ALTER TABLE MAEART ADD COLUMN   AMARCA  TEXT(20)" '
  End If
  If Not ExisteElem(1, VGCNx, "MAEART", "ACOLOR") Then
      VGCNx.Execute "ALTER TABLE  MAEART  ADD COLUMN   ACOLOR  TEXT(20)" '
  End If
 '*****************************************************************
 '*** ULTIMA ACTUALIZACION 28/06/2001    ROBERTO M.M.
 '*****************************************************************
   If Not ExisteElem(0, VGCNx, "AL_CIERRESMENSUALES") Then
        SQL = " Create Table AL_CIERRESMENSUALES (CIERRMES Text(6),CIERRALMA Text(2),CIERRFECH DATETIME, CIERROPER TEXT(15) , " & _
        " CONSTRAINT Clave PRIMARY KEY (CIERRMES))"
        VGCNx.Execute SQL
  Else
       If Not ExisteElem(1, VGCNx, "AL_CIERRESMENSUALES", "CIERRALMA") Then
           VGCNx.Execute "ALTER TABLE  AL_CIERRESMENSUALES  ADD COLUMN  CIERRALMA text(2) " '
       End If
  End If
  
  If Not ExisteElem(1, VGCNx, "MORESMES", "SMSALDOINI") Then
      VGCNx.Execute "ALTER TABLE  MORESMES  ADD COLUMN  SMSALDOINI  DOUBLE " '
  End If
   
  If Not ExisteElem(0, VGCNx, "COSPROFECH") Then
        SQL = " Create Table COSPROFECH ( AUXALMA Text(3),AUXTD Text(3),AUXNUMDOC Text(10),AUXCODART Text(20) ,AUXFECDOC DATETIME,AUXCANT DOUBLE,AUXPRECIO DOUBLE,AUXPRECOS DOUBLE   )" '(AUXTD , AUXNUMDOC , AUXCODART , AUXFECDOC )
        VGCNx.Execute SQL
  End If
  
  
  If Not ExisteElem(1, VGCNx, "KARDEXAUX", "TIPDOCRF") Then
     VGCNx.Execute "ALTER TABLE  KARDEXAUX  ADD COLUMN  TIPDOCRF text(2) " '
  End If
  
  If Not ExisteElem(1, VGCNx, "KARDEXAUX", "NUMDOCRF") Then
     VGCNx.Execute "ALTER TABLE  KARDEXAUX  ADD COLUMN  NUMDOCRF text(10) " '
  End If
  
  If Not ExisteElem(1, VGCNx, "KARDEXAUX", "NOMREFE") Then
     VGCNx.Execute "ALTER TABLE  KARDEXAUX  ADD COLUMN  NOMREFE text(50) " '
  End If
  
  Call ADOConectar
  
  If Not ExisteElem(1, VGCNx, "al_Kardex_Val", "ING_SAL") Then
     cConexAux.Execute "ALTER TABLE  al_Kardex_Val  ADD COLUMN  ING_SAL TEXT(20) " '
  End If
  
  cConexAux.Close
  
  If Not ExisteElem(0, VGCNx, "InveFisiCab") Then
        SQL = " Create Table InveFisiCab ( AUXNUMINVE TEXT(10), AUXALMA Text(3),AUXFECH DATETIME ,AUXRESPON TEXT(15),AUXOBSER TEXT(255)" & _
        ", CONSTRAINT Clave PRIMARY KEY ( AUXALMA,AUXNUMINVE )  )"
        VGCNx.Execute SQL
  Else
       If Not ExisteElem(1, VGCNx, "InveFisiCab", "AUXESTADO") Then
          VGCNx.Execute "ALTER TABLE  InveFisiCab  ADD COLUMN  AUXESTADO TEXT(2) " '*****INDICA SI YA SE INGRESO O ESTA PENDIENTE EN INVENTARIO FISICO
       End If
  End If

  
  If Not ExisteElem(0, VGCNx, "InveFisiDet") Then
        SQL = " Create Table InveFisiDet ( AUXNUMINVE TEXT(10), AUXALMA Text(3), AUXFAMIL Text(8),AUXCODART Text(20) ,AUXSTOCK DOUBLE,AUXINGR DOUBLE,AUXDIFE DOUBLE " & _
        ", CONSTRAINT Clave PRIMARY KEY ( AUXALMA , AUXNUMINVE,AUXCODART )  )"
        VGCNx.Execute SQL
  Else
       If Not ExisteElem(1, VGCNx, "InveFisiDet", "AUXFAMIL") Then
          VGCNx.Execute "ALTER TABLE  InveFisiDet  ADD COLUMN  AUXFAMIL Text(8) " '*****INDICA SI YA SE INGRESO O ESTA PENDIENTE EN INVENTARIO FISICO
       End If
        
  End If

  If Not ExisteElem(1, VGCNx, "CONFIGURACION", "conf_codigoIng") Then
     VGCNx.Execute "ALTER TABLE CONFIGURACION  ADD COLUMN  conf_codigoIng Text(8) "
  End If
  
  
 '*****************************************************************
  If Not ExisteElem(0, VGCNx, "MAECOLOR") Then
        SQL = " Create Table MAECOLOR (COD_COLOR Text(20),DESCRI_COLOR Text(20), " & _
        " CONSTRAINT Clave PRIMARY KEY (COD_COLOR))"
        VGCNx.Execute SQL
  End If
  
  If Not ExisteElem(0, VGCNx, "MAEMARCA") Then
        SQL = " Create Table MAEMARCA (COD_MARCA Text(20),DESCRI_MARCA Text(20), " & _
        " CONSTRAINT Clave PRIMARY KEY (COD_MARCA))"
        VGCNx.Execute SQL
  End If
  
   If ExisteIndice(cRuta2, "InveFisiDet", "clave") Then
     EliminaIndice cRuta2, "InveFisiDet", "clave"
     CreaIndice cRuta2, "InveFisiDet", "PRIMARYKEY", True, "AUXNUMINVE", "AUXALMA", "AUXCODART"
  End If
     

 
End Sub

Public Function EncMes(FF As Date) As Boolean       'CORREGIR
 Dim rs As ADODB.Recordset
 Dim RSQL As String
    If Month(FF) < 12 Then
        RSQL = "SELECT CACIERRE FROM MOVALMCAB WHERE CAFECDOC >='01/" & Format(Month(FF), "00") & "/" & Format(Year(FF), "0000") & "' AND CAFECDOC < '" & Format(Val(Month(FF)) + 1, "00") & "/01/" & (Format(Year(FF), "0000")) & "'"
    Else
        RSQL = "SELECT CACIERRE FROM MOVALMCAB WHERE CAFECDOC >='01/" & Format(Month(FF), "00") & "/" & Format(Year(FF), "0000") & "' AND CAFECDOC < '01/01/" & (Format(Val(Year(FF)) + 1, "0000")) & "'"
    End If
    Set rs = New ADODB.Recordset
    rs.Open RSQL, VGCNx, adOpenStatic, adLockReadOnly
    If rs.EOF Then EncMes = False: rs.Close:  Exit Function
    If rs(0) = True Then EncMes = True
    rs.Close
End Function
'*************** FUNCIONES ADICIONADAS A 28 JUNIO DEL 2001
'*************** ROBERTO M.M.
Function AnioMesAnterior(ByVal arAnioMes As String) As String
Dim LMes, LAnio As String
   If Val(Mid(arAnioMes, 5, 2)) = 1 Then
      LAnio = Val(Left(arAnioMes, 4)) - 1
      LMes = 12
   Else
      LAnio = Val(Left(arAnioMes, 4))
      LMes = Val(Mid(arAnioMes, 5, 2)) - 1
   End If
   AnioMesAnterior = Format(LAnio, "0000") & Format(LMes, "00")
End Function

Function AnioMesSiguiente(ByVal arAnioMes As String) As String
Dim LMes, LAnio As String
   If Val(Mid(arAnioMes, 5, 2)) = 12 Then
      LAnio = Val(Left(arAnioMes, 4)) + 1
      LMes = 1
   Else
      LAnio = Val(Left(arAnioMes, 4))
      LMes = Val(Mid(arAnioMes, 5, 2)) + 1
   End If
   AnioMesSiguiente = Format(LAnio, "#000") & Format(LMes, "0#")
End Function
'**************************************************************
'**************************************************************
Function DevolverTCambio(ByVal arDate As Date) As Double
On Error Resume Next
          If UCase(Dir$(cRuta4)) = UCase(cNomBd4) Then
              DevolverTCambio = Val(Devolver_Dato(3, CDate(arDate), "Tipo_Cambio", "TIPOCAMB_FECHA", True, "TIPOCAMB_VENTA", "02", "TIPOMON_CODIGO"))
          ElseIf UCase(Dir$(cRuta2)) = "BDCOMUN.MDB" Then
              DevolverTCambio = Val(Devolver_Dato(1, CDate(arDate), "Tipo_Cambio", "TIPOCAMB_FECHA", True, "TIPOCAMB_VENTA", "02", "TIPOMON_CODIGO"))
          End If
End Function


Function UltimoCierre() As String
Dim rs As New ADODB.Recordset
rs.Open "Select max(CierrMes) as Tot From AL_CIERRESMENSUALES where empresacodigo='" & VGParametros.empresacodigo & "'", VGCNx, adOpenForwardOnly, adLockReadOnly
If Not rs.EOF Then
   If Not (IsNull(rs!tot) Or rs!tot = "") Then
      UltimoCierre = rs!tot
   Else
      UltimoCierre = "" 'IIf(IsNull(rs!Min) Or rs!Min = "", "", AnioMesAnterior(Format(Year(rs!Min), "0000") & Format(Month(rs!Min), "00")))
   End If
End If
rs.Close
End Function

Private Sub ADOConectar()
Dim cRt As String
cRt = App.Path & "\BdAuxCom.Mdb"
Set cConexAux = New ADODB.Connection
cConexAux.CursorLocation = adUseClient
cConexAux.ConnectionString = "PROVIDER=Microsoft.Jet.OLEDB.3.51;Data Source=" & cRt & ";"
cConexAux.Open
End Sub

Function cNull(ByVal arNulo As Variant) As Variant
         cNull = IIf(IsNull(arNulo), "", arNulo)
End Function


Public Function ExisteIndice(BaseDeDatos As String, NombreTabla As String, NombreIndice As String) As Boolean
'**********************************************
'*                                            *
'*   Verifica la existencia de un índice      *
'*   Funcion creada por Julio Calderón        *
'*                                            *
'**********************************************

'   ExisteIndice = False
    
'   'Dim dbs As DAO.Database
'   Dim Tdf As DAO.TableDef
'   Dim idxBucle As DAO.Index
'
'   Set dbs = OpenDatabase(BaseDeDatos)
'   Set Tdf = dbs(nombretabla)
'   With Tdf
'      ' Recorre la colección Indexes de la tabla.
'      For Each idxBucle In .Indexes
'         If StrConv(idxBucle.name, vbUpperCase) = StrConv(NombreIndice, vbUpperCase) Then
'           ExisteIndice = True
'         End If
'      Next idxBucle
'   End With
'   dbs.Close
End Function

Public Sub CreaIndice(BaseDeDatos As String, _
                      NombreTabla As String, _
                      NombreIndice As String, _
                      ClavePrimaria As Boolean, _
                      CampoIndice1 As String, _
                      Optional CampoIndice2 As String, _
                      Optional CampoIndice3 As String, _
                      Optional CampoIndice4 As String)
'************************************************
'*                                              *
'*   Crea un índice compuesto de hasta 4 campos *
'*   Funcion creada por Julio Calderón          *
'*                                              *
'************************************************
   
'   Dim dbs As DAO.Database
'   Dim Tdf As DAO.TableDef
'   Dim NuevoIndice As DAO.Index
'   Dim idxNombre As DAO.Index
'   Dim idxBucle As DAO.Index
'
'   Set dbs = OpenDatabase(BaseDeDatos)
'   Set Tdf = dbs(nombretabla)
'   With Tdf
'      ' Primero crea objeto Index, crea y agrega los
'      ' objetos Field al objeto Index y después agrega
'      ' el objeto Index a la colección Indexes de TableDef.
'      Set NuevoIndice = .CreateIndex(NombreIndice)
'      With NuevoIndice
'         .Fields.Append .CreateField(CampoIndice1)
'         If CampoIndice2 <> "" Then
'           .Fields.Append .CreateField(CampoIndice2)
'         End If
'         If CampoIndice3 <> "" Then
'           .Fields.Append .CreateField(CampoIndice3)
'         End If
'         If CampoIndice4 <> "" Then
'           .Fields.Append .CreateField(CampoIndice4)
'         End If
'      End With
'
'      NuevoIndice.Primary = ClavePrimaria
'
'      .Indexes.Append NuevoIndice
'      ' Actualiza la colección para que pueda tener
'      ' acceso a los objetos Index nuevos.
'      .Indexes.Refresh
'   End With
'   dbs.Close
End Sub

Public Sub EliminaIndice(ByVal BaseDeDatos As String, ByVal NombreTabla As String, ByVal NombreIndice As String)
'**********************************************
'*                                            *
'*   Elimina un índice                        *
'*   Sub creada por Julio Calderón            *
'*                                            *
'**********************************************

'   Dim dbs As DAO.Database
'   Dim Tdf As DAO.TableDef
'   Dim idxBucle As DAO.Index
'
'   Set dbs = OpenDatabase(BaseDeDatos)
'   Set Tdf = dbs(nombretabla)
'
'   With Tdf
'   'Recorre la colección Indexes de la tabla.
'     For Each idxBucle In .Indexes
'         If StrConv(idxBucle.name, vbUpperCase) = StrConv(NombreIndice, vbUpperCase) Then
'          .Indexes.Delete (NombreIndice)
'          Exit For
'         End If
'     Next idxBucle
'   End With
'   dbs.Close
End Sub

Function UltimoCierreFech(ByVal xfecha As Date) As Date
Dim Rs2 As New ADODB.Recordset
Dim RSQL As String
Dim CIERRE, Anio, mes As String
Dim lFecha As Date
CIERRE = UltimoCierre
If CIERRE = "" Then
   UltimoCierreFech = CDate(Format(xfecha, "dd/MM/yyyy"))
   Exit Function
Else
   Anio = Left(CIERRE, 4)
   mes = Mid(CIERRE, 5, 2)
   lFecha = CDate("01/" & mes & "/" & Anio)
   If xfecha <= Fin_MES(lFecha) Then
       UltimoCierreFech = CDate(Format(Fin_MES(lFecha) + 1, "dd/MM/yyyy"))
    Else
       UltimoCierreFech = CDate(Format(xfecha, "dd/MM/yyyy"))
    End If
End If
       
End Function

Sub Enteros_Positivos(k As Integer, t As TextBox)
    If k = 8 Then Exit Sub
    If k < 48 Or k > 57 Then
        k = 0
    End If
End Sub

Sub Reales_Positivos(k As Integer, t As TextBox)
Dim t1 As String
    k = Asc(UCase(Chr(k)))
    If k = 8 Then Exit Sub
    If k <> 45 And k <> 44 And k <> 32 And k <> 69 And k <> 43 Then
        t1 = Left(t, t.SelStart)
        t1 = t1 & Chr(k) & Right(t, Len(t) - Len(t1))
        If IsNumeric(t1) Then Exit Sub
    End If
    k = 0
    
End Sub

Public Function DateSQL2000(ByVal Fecha As Variant) As Variant
    If Len(Trim(Fecha)) > 0 And IsDate(Fecha) Then
      ' DateSQL2000 = "'" & Month(Fecha) & "/" & Day(Fecha) & "/" & Year(Fecha) & "'"
        DateSQL2000 = "'" & Fecha & "'"
    Else
       'DateSQL2000 = "'" & Month(0) & "/" & Day(0) & "/" & Year(0) & "'"
       'DateSQL2000 = "'" & "00/00/0000" & "'"
       DateSQL2000 = "Null"
    End If
End Function

Function nNull(ByVal arNulo As Variant) As Variant
         nNull = IIf(IsNull(arNulo), 0, arNulo)
End Function

'*****************************************************************
'EL ULTIMO INGRESO CUYO PRECIO SEA MAYOR QUE CERO  RMM 09/06/2001
'*****************************************************************
Function UltimoPrecio(ByVal arCodigo As String, ByVal moneda As String) As Double
 Dim rs As New ADODB.Recordset
 Dim RSQL As String
 RSQL = "SELECT DEprecio,CACODMON,CATIPCAM FROM MovAlmDet AS A INNER JOIN MovAlmCab AS B ON (B.CAALMA = A.DEALMA) AND " & _
         "(B.CATD = A.DETD) AND (B.CANUMDOC = A.DENUMDOC) WHERE CAALMA = '" & VGAlma & "'  And CASITGUI<>'A' and " & _
         "catipmov='I'  and decodigo='" & arCodigo & "' and  a.deprecio<>0 AND  cafecdoc= ( SELECT max(cafecdoc) FROM MovAlmDet AS A " & _
         "INNER JOIN MovAlmCab AS B ON (B.CAALMA = A.DEALMA) AND (B.CATD = A.DETD) AND (B.CANUMDOC = " & _
         "A.DENUMDOC) WHERE CAALMA = '" & VGAlma & "'  And CASITGUI<>'A'  and a.deprecio<>0 and catipmov='I' and decodigo='" & arCodigo & "') "
 rs.Open RSQL, VGCNx, adOpenStatic
 UltimoPrecio = 0
 If Not rs.EOF Then
    If rs(1) = "02" And moneda = "01" And nNull(rs(2)) <> 0 Then UltimoPrecio = Format(rs(0) * rs(2), "###,##0.00")
    If rs(1) = "02" And moneda = "02" And nNull(rs(2)) <> 0 Then UltimoPrecio = Format(rs(0), "###,##0.00")
    If rs(1) = "01" And moneda = "02" And nNull(rs(2)) <> 0 Then UltimoPrecio = Format(rs(0) / rs(2), "###,##0.00")
    If rs(1) = "01" And moneda = "01" And nNull(rs(2)) <> 0 Then UltimoPrecio = Format(rs(0), "###,##0.00")
 Else
     UltimoPrecio = 0
 End If
 
 rs.Close
End Function



Function Fin_MES(ByVal Afech As Date) As Date
Dim mes, Anio, lastday As String
mes = Format(Month(Afech), "0#")
Anio = CStr(Year(Afech))
lastday = "31/" & mes & "/" + Anio

If IsDate(lastday) Then
   Fin_MES = Format(lastday, "dd/mm/yyyy")
   Exit Function
Else
   lastday = "30/" & mes & "/" + Anio
   If IsDate(lastday) Then
      Fin_MES = Format(lastday, "dd/mm/yyyy")
      Exit Function
   Else
       lastday = "29/" & mes & "/" + Anio
       If IsDate(lastday) Then
          Fin_MES = Format(lastday, "dd/mm/yyyy")
          Exit Function
       Else
           lastday = "28/" & mes & "/" + Anio
           If IsDate(lastday) Then
              Fin_MES = Format(lastday, "dd/mm/yyyy")
              Exit Function
           Else
              MsgBox "Existe errores en la function Fin_mes...!"
           End If
       End If
   End If
End If
End Function

Function Ini_MES(ByVal Afech As Date) As Date
Dim mes, Anio As String
mes = Format(Month(Afech), "0#")
Anio = CStr(Year(Afech))
Ini_MES = CDate("01/" & mes & "/" + Anio)
End Function

Sub Tabula(ByVal key As Long)
    If key = 13 Then SendKeys "{tab}"
End Sub


Function FechMask(ByVal arFech As Variant) As Variant
If IsNull(arFech) Then
   FechMask = "__/__/____"
   Exit Function
End If
If Year(arFech) < 1901 Or Not IsDate(arFech) Then
   FechMask = "__/__/____"
Else
   FechMask = arFech
End If

End Function

Public Sub ModiFieldDef(ByVal sDataBase As String, ByVal sTable As String, ByVal sField As String, _
 Optional ByVal sType As String, _
 Optional ByVal Decimales As Variant, _
 Optional ByVal AllowZeroLen As Variant, _
 Optional ByVal FRequired As Variant, _
 Optional ByVal DefVal As Variant)
 
End Sub
Public Sub DEMORA(TIEMPO As Double)
'Fernando: 06/08/2001:
Dim HORA As Double
    Screen.MousePointer = 11
    HORA = Time()
    Do While Format(TimeSerial(0, 0, TIEMPO) + HORA, "HH:MM:SS") <> Format(Time(), "HH:MM:SS")
       ' DoEvents
    Loop
    Screen.MousePointer = 1
End Sub

Sub impresion(cNombreReporte As String)
On Error GoTo X
  With MDIPrincipal.CryRptProc
        .Reset
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .ReportFileName = VGParamSistem.RutaReport & cNombreReporte
      '  .LogOnServer "pdssql.dll", VGParamSistem.Servidor, VGParamSistem.BDEmpresa, VGParamSistem.UsuarioReporte, VGParamSistem.Pwd
        .Connect = VGCadenaReport2
        .DiscardSavedData = True
        .Action = 1
  End With
  Exit Sub
X:
  MsgBox "Error inesperado: " & Err.Number & "  " & Err.Description, vbExclamation
End Sub


Public Sub Init_ControlDataGrid(EsteGrid As DataGrid)
 With EsteGrid
  .AllowAddNew = False
  .AllowDelete = False
  .AllowUpdate = False
  .AllowRowSizing = False
  .TabAction = dbgControlNavigation
  .MarqueeStyle = dbgHighlightRow
 ' .Font =
 End With
End Sub
Public Sub ParametrosdeAlmacenes()
  Dim RSAUX As New ADODB.Recordset
  RSAUX.Open "select * from empresa where emp_codigo='" & VGCodEmpresa & "'", VGConfig
    If RSAUX.RecordCount = 0 Then Exit Sub
    Set VGvardllgen = New dllgeneral.dll_general
    VGParametros.VGLongCodigo = VGvardllgen.ESNULO(RSAUX!digitoscodigo, 10)
End Sub


Public Sub importarExcel(Tabla As Recordset, hojadecalculo As String, DataGrid1 As DataGrid)
    'Variable de tipo Aplicación de Excel
    Dim objExcel As Excel.Application
 
    'Una variable de tipo Libro de Excel
    Dim xLibro As Excel.Workbook
    Dim fila As Integer
    Dim Columna As Integer
        'creamos un nuevo objeto excel
    Set objExcel = New Excel.Application
    objExcel.SheetsInNewWorkbook = 1
  
    'Usamos el método open para abrir el archivo que está _
     en el directorio del programa llamado archivo.xls
    Set xLibro = objExcel.Workbooks.Open("" & hojadecalculo & "")
  
    'Hacemos el Excel Visible
    objExcel.Visible = True
    Set DataGrid1.DataSource = Tabla
    With xLibro
  
        ' Hacemos referencia a la Hoja
        With .Sheets(1)

            'Recorremos la fila desde la 1 hasta la 7
            For fila = 5 To 30000
                If .Cells(fila, 1) = "" Then Exit For
                Tabla.AddNew
                For Columna = 1 To 12
                'Agregamos el valor de la fila que _
                 corresponde a la columna 2
                Tabla.Fields(Columna - 1) = .Cells(fila, Columna)
                Next
                Tabla.Update
                DataGrid1.Refresh
            Next
         
        End With
    End With
    Tabla.Update
    Tabla.UpdateBatch adAffectAllChapters
 
  
    'Eliminamos los objetos si ya no los usamos
    Set objExcel = Nothing
    Set xLibro = Nothing
  
End Sub

Public Function sGetIni(sIniFile As String, sSection As String, sKey As String, sDefault As String) As String
 Dim sTemp As String * 256
 Dim nLength As Integer
 sTemp = Space$(256)
 nLength = GetPrivateProfileString(sSection, sKey, sDefault, sTemp, 255, sIniFile)
 sGetIni = Left$(sTemp, nLength)
End Function

Public Sub WriteIni(sIniFile As String, sSection As String, sKey _
                    As String, sValue As String)
 Dim sTemp As String
 Dim n As Integer
 
 sTemp = sValue
 For n = 1 To Len(sValue)
  If Mid$(sValue, n, 1) = vbCr Or Mid$(sValue, n, 1) = vbLf Then
   Mid$(sValue, n) = " "
  End If
 Next n
 n = WritePrivateProfileString(sSection, sKey, sTemp, sIniFile)
End Sub



Sub ImpresionRptCad(Reporte As CrystalReport, cNombreReporte As String, PFormulas(), Param(), Optional ORDEN As String, Optional titulo As String, Optional Seleccion As String)
Dim i As Integer
Dim sServer As String
Dim sBase As String
Dim sUsuario As String
Dim sPwd As String
Dim sRutaReportes As String
On Error GoTo procImpresionRptError
    'Leo ini  sección de proc(s) almacenados
    sServer = sGetIni(App.Path & "\wenco.ini", "Bstore", "dserver", "?")
    If Trim(sServer) = "?" Then sServer = "(local)"
        
    sBase = sGetIni(App.Path & "\wenco.ini", "Bstore", "dbase", "?")
    If Trim(sBase) = "?" Then sBase = "BDMAIN"
    
    sUsuario = sGetIni(App.Path & "\wenco.ini", "Bstore", "duser", "?")
    If Trim(sUsuario) = "?" Then sUsuario = "sa"
    
    sPwd = sGetIni(App.Path & "\wenco.ini", "Bstore", "dpass", "?")
    If Trim(sPwd) = "?" Then sPwd = ""

    'Leo la ruta en donde se encuentra los archivos de reportes
    sRutaReportes = sGetIni(App.Path & "\wenco.ini", "CONFIG", "rpt ", "?")
    If Trim(sRutaReportes) = "?" Then sRutaReportes = "C:\WENCO\REPORTES\"
    Screen.MousePointer = 11
    With Reporte
        .Reset
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .WindowTitle = titulo
        .WindowShowPrintSetupBtn = True
        .WindowShowRefreshBtn = True
        .WindowShowSearchBtn = True
        .ReportFileName = sRutaReportes & cNombreReporte
'        .LogOnServer "pdssql.dll", sServer, sBase, sUsuario, sPwd
        .Connect = "DSN=" & sServer & ";DSQ=" & sBase & ";UID=" & sUsuario & ";PWD=" & sPwd
        If UBound(PFormulas) > 0 Then
            For i = 0 To UBound(PFormulas) - 1
                .formulas(2 + i) = PFormulas(i)
            Next
        End If
        .DiscardSavedData = True
        If UBound(Param) > 0 Then
            For i = 0 To UBound(Param) - 1
                .StoredProcParam(i) = Param(i)
            Next
        End If
   '     If orden <> "" Then Call CrystOrden(MDIPrincipal.CryRptProcProc, orden)
        If Seleccion <> "" Then .SelectionFormula = Seleccion
        If .Status <> 2 Then .Action = 1
    End With
    Screen.MousePointer = 1
    Exit Sub
    
procImpresionRptError:
  If Err.Number = 9 Then Resume Next
  Screen.MousePointer = 1
  MsgBox "Error inesperado: " & Err.Number & "  " & Err.Description, vbExclamation
End Sub
Sub ImpresionRptdefault(Reporte As CrystalReport, titulo)
Dim i As Integer
Dim sServer As String
Dim sBase As String
Dim sUsuario As String
Dim sPwd As String
Dim sRutaReportes As String
Screen.MousePointer = 11
With Reporte
     .Reset
     .Destination = crptToWindow
     .WindowState = crptMaximized
     .WindowTitle = titulo
     .WindowShowPrintSetupBtn = True
     .WindowShowRefreshBtn = True
     .WindowShowSearchBtn = True
End With
End Sub

Sub Captura_error()
    If Err.Number <> 0 Then
        MsgBox Str(Err.Number) + "," + Err.Description, vbCritical, "Mensaje"
        
    End If
End Sub


