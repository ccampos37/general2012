VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsMantMov1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Enum Campos
    cuenta = 1
    gastos = 2
    glosa = 3
    analitico = 4
    ccosto = 5
    impbruto = 6
    Igv = 7
    Inafecto = 8
    ImpCompra = 9

End Enum
Public Enum Tipocambio
    Compra = "01"
    Venta = "02"
    Promedio = "03"
End Enum
Public Sub CargarAyudas()
With frmMantprovision
    Call .CtrAyu_Cuenta.Conexion(VGcnxCT): .CtrAyu_Cuenta.Filtro = "(cuentanivel=" & VGnumniveles & " and cuentacodigo <>'00') and (" & VGParametros.ctascompra & ")"
    Call .CtrAyu_gastos.Conexion(VGcnx): .CtrAyu_gastos.Filtro = "(gastosnivel=" & VGnumnivgas & " and gastoscodigo <>'00') "
    Call .CtrAyu_TipDoc.Conexion(VGcnx): .CtrAyu_TipDoc.Filtro = "tdocumentocodigo<>'00'"
    Call .CtrAyu_TipRef.Conexion(VGcnx): .CtrAyu_TipRef.Filtro = "tdocumentocodigo<>'00'"
    Call .CtrAyu_Moneda.Conexion(VGcnxCT): .CtrAyu_Moneda.Filtro = "monedacodigo<>'00'"
    Call .CtrAyu_Modoprovi.Conexion(VGcnx): .CtrAyu_Modoprovi.Filtro = "modoprovicod <>'00'"
    Call .CtrAyu_TipCompra.Conexion(VGcnx)
    Call .CtrAyu_TipSubAsi.Conexion(VGcnx)
    Call .Ctr_AyuAnalitico.Conexion(VGcnx)
    Call .CtrAyu_Proveedor.Conexion(VGcnx): .CtrAyu_Proveedor.Filtro = "clientecodigo <>'00'"
    Call .Ctr_AyudaCaja.Conexion(VGcnx)
    Call .Ctr_AyudaOficina.Conexion(VGcnx)
    Call .CtrAyu_Ccosto.Conexion(VGcnxCT)
    .CtrAyu_Ccosto.Filtro = "centrocostotipo=" & VGnumnivcos & " and centrocostocodigo<>'00' "
    Call .Ctr_Ayuempresa.Conexion(VGcnx): .Ctr_Ayuempresa.Filtro = "empresacodigo<>'00'"
End With
End Sub
Public Sub CreaRsTempDetalle(ByRef rs As ADODB.Recordset)
    With rs.Fields
        Call .Append("Item", adVarChar, 3)
        Call .Append("gastoscodigo", adVarChar, 20, adFldIsNullable)
        Call .Append("cuentacodigo", adVarChar, 20, adFldIsNullable)
        Call .Append("CuentaDes", adVarChar, 100, adFldIsNullable)
        Call .Append("impbruto", adDouble)
        Call .Append("Igv", adDouble)
        Call .Append("Inafecto", adDouble)
        Call .Append("Impcompra", adDouble)
        Call .Append("glosa", adVarChar, 100)
        Call .Append("Ccosto", adVarChar, 10)
        Call .Append("analitico", adVarChar, 11)
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
Public Sub Limpia()
    VGflaglimpia = False
    With frmMantprovision
        .CtrAyu_Cuenta.xclave = "": .CtrAyu_Cuenta.xnombre = ""
        .CtrAyu_gastos.xclave = "": .CtrAyu_gastos.xnombre = ""
        If VGParametros.sistemactrlgastos = True Then
           .CtrAyu_gastos.Visible = True
         Else
           .CtrAyu_Cuenta.Visible = True
        End If
        .TxImpBruto.Text = ""
        .TxIGV.Text = ""
        .TxInafecto.Text = ""
        .TxImpCompra.Text = ""
        .Txtglosa.Text = ""
        .CtrAyu_Ccosto.xclave = "": .CtrAyu_Ccosto.xnombre = ""
    End With
    VGflaglimpia = True
End Sub
Public Sub ActualizarDetalle(ByRef rs As ADODB.Recordset, OP As Campos)
    Dim vardllgen As New dllgeneral.dll_general
    If rs Is Nothing Then Exit Sub
    If rs.RecordCount = 0 Then Exit Sub
       VGactulizodoc = False
    If VGMoverRegistro Then Exit Sub
    Screen.MousePointer = 11
    With rs
        '!detcomprobitem = Format(Val(vardllgen.ESNULO(MaxItem(rs, "detcomprobitem"), 0) + 1), "00000")
        Select Case OP
            Case cuenta
                !Cuentacodigo = frmMantprovision.CtrAyu_Cuenta.xclave
                !cuentades = frmMantprovision.CtrAyu_Cuenta.xnombre
                !ccosto = "00"
                !gastoscodigo = "00"
           Case gastos
                    !gastoscodigo = frmMantprovision.CtrAyu_gastos.xclave
                !cuentades = frmMantprovision.CtrAyu_gastos.xnombre
                !ccosto = "00"
                !Cuentacodigo = Cuentacodigo
            Case glosa: !glosa = Trim(frmMantprovision.Txtglosa.Text)
            Case analitico: !analitico = frmMantprovision.Ctr_AyuAnalitico.xclave
            Case ccosto: !ccosto = frmMantprovision.CtrAyu_Ccosto.xclave
            Case impbruto: !impbruto = vardllgen.ESNULO(Espunto(frmMantprovision.TxImpBruto.valor), 0)
            Case Igv: !Igv = vardllgen.ESNULO(Espunto(frmMantprovision.TxIGV.valor), 0)
            Case Inafecto: !Inafecto = vardllgen.ESNULO(Espunto(frmMantprovision.TxInafecto.valor), 0)
            Case ImpCompra: !ImpCompra = vardllgen.ESNULO(Espunto(frmMantprovision.TxImpCompra.valor), 0)
            

        
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
Public Sub GeneraAsientoenLine(ByVal OP As Integer, ByVal Nprovi As String, ByVal Comprob_Contable As String)
On Error GoTo genasiento
    Screen.MousePointer = 11
    'Generando los Analticos que no Esten en contabilidad
    Set VGCommandoSP = New ADODB.Command
    Set VGvardllgen = New dllgeneral.dll_general
    VGCommandoSP.ActiveConnection = VGcnxMarfice
    VGCommandoSP.CommandType = adCmdStoredProc
    VGCommandoSP.CommandText = "co_VerificaAnaliticoenLinea"
    VGCommandoSP.Parameters.Refresh
    With VGCommandoSP
        .Parameters("@BaseConta") = VGParamSistem.BDEmpresaCT
        .Parameters("@BaseCompra") = VGParamSistem.BDEmpresa
        .Parameters("@Ano") = VGParamSistem.Anoproceso
        .Parameters("@Mes") = VGParamSistem.Mesproceso
        .Parameters("@tipanal") = VGParametros.xTipAnal
        .Parameters("@User") = VGParamSistem.Usuario
        .Parameters("@Nprovi") = Nprovi
        .Execute
    End With
    'Generando el Asiento en contabilidad
    Set VGCommandoSP = New ADODB.Command
    Set VGvardllgen = New dllgeneral.dll_general
    
    VGCommandoSP.ActiveConnection = VGcnxMarfice
    VGCommandoSP.CommandType = adCmdStoredProc
    VGCommandoSP.CommandText = "co_generaasientolinea_pro"
    VGCommandoSP.Parameters.Refresh
    With VGCommandoSP
        .Parameters("@BaseConta") = VGParamSistem.BDEmpresaCT
        .Parameters("@BaseCompra") = VGParamSistem.BDEmpresa
        .Parameters("@SubAsiento") = VGParametros.xsubasiento
        .Parameters("@Libro") = VGParametros.xLibro
        If Month(frmMantprovision.DTPFechaContab) <> Val(VGParamSistem.Mesproceso) Then
           .Parameters("@mes") = Format(Month(frmMantprovision.DTPFechaContab), "00")
         Else
           .Parameters("@mes") = Format(VGParamSistem.Mesproceso, "00")
        End If
        .Parameters("@Ano") = VGParamSistem.Anoproceso
        .Parameters("@ctatotal") = VGParametros.xCtaTotal
        .Parameters("@ctaIGV") = VGParametros.xCtaIGV
        .Parameters("@ctaIES") = VGParametros.xCtaIES
        .Parameters("@ctaRTA") = VGParametros.xCtaRTA
        .Parameters("@tipanal") = VGParametros.xTipAnal
        .Parameters("@Compu") = VGComputer
        .Parameters("@Usuario") = VGParamSistem.Usuario
        .Parameters("@Oficina") = Format(VGParametros.CpOficina, "00")
        .Parameters("@Nprovi") = Nprovi
        .Parameters("@op") = OP
        .Parameters("@comprobconta") = Comprob_Contable
        .Execute
    End With
    
    'Actualizando las Glosas de Cabecera y Detalle
    Set VGCommandoSP = New ADODB.Command
    Set VGvardllgen = New dllgeneral.dll_general
    VGCommandoSP.ActiveConnection = VGcnxMarfice
    VGCommandoSP.CommandType = adCmdStoredProc
    VGCommandoSP.CommandText = "co_GrabaGlosasProvisionLinea_pro"
    VGCommandoSP.Parameters.Refresh
    With VGCommandoSP
        .Parameters("@BaseConta") = VGParamSistem.BDEmpresaCT
        .Parameters("@BaseCompra") = VGParamSistem.BDEmpresa
        If Month(frmMantprovision.DTPFechaContab) <> Val(VGParamSistem.Mesproceso) Then
           .Parameters("@mes") = Format(Month(frmMantprovision.DTPFechaContab), "00")
         Else
           .Parameters("@mes") = Format(VGParamSistem.Mesproceso, "00")
        End If
        .Parameters("@Ano") = VGParamSistem.Anoproceso
        .Parameters("@Nprovi") = Nprovi
        .Execute
    End With
    
    'Actualizando los Registros que no se incluyen en el Reg. Compras
    Set VGCommandoSP = New ADODB.Command
    Set VGvardllgen = New dllgeneral.dll_general
    VGCommandoSP.ActiveConnection = VGcnxMarfice
    VGCommandoSP.CommandType = adCmdStoredProc
    VGCommandoSP.CommandText = "co_RegComprasNoIncluyenenLinea_pro"
    VGCommandoSP.Parameters.Refresh
    With VGCommandoSP
        .Parameters("@BaseCompra") = VGParamSistem.BDEmpresa
        .Parameters("@BaseConta") = VGParamSistem.BDEmpresaCT
        .Parameters("@Asiento") = "081"
        If Month(frmMantprovision.DTPFechaContab) <> Val(VGParamSistem.Mesproceso) Then
           .Parameters("@mes") = Format(Month(frmMantprovision.DTPFechaContab), "00")
         Else
           .Parameters("@mes") = Format(VGParamSistem.Mesproceso, "00")
        End If
        .Parameters("@Ano") = VGParamSistem.Anoproceso
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
Public Sub GrabarCabecera(ByVal OP As Integer, Optional ByVal numero As Long, Optional ByVal NumeroAux As String, Optional ByVal Numerotesor As String)
On Error GoTo ErrorGrabaCabecera
    Set VGCommandoSP = New ADODB.Command
    Set VGvardllgen = New dllgeneral.dll_general
    VGCommandoSP.ActiveConnection = VGcnxMarfice
    VGCommandoSP.CommandType = adCmdStoredProc
    VGCommandoSP.CommandText = "co_grabacabprovi"
    VGCommandoSP.Parameters.Refresh
     With VGCommandoSP
        .Parameters("@base") = VGParamSistem.BDEmpresa
        .Parameters("@tabla") = VGParamSistem.TablaCabcomprob
        .Parameters("@op") = OP
        If Month(frmMantprovision.DTPFechaContab) <> Val(VGParamSistem.Mesproceso) Then
            .Parameters("@cabprovimes") = Month(frmMantprovision.DTPFechaContab)
         Else
            .Parameters("@cabprovimes") = Val(VGParamSistem.Mesproceso)
        End If
        .Parameters("@cabprovinumero") = numero
        'Este para que al eliminar no utilizar estos parametros
        If OP <= 2 Then
            .Parameters("@proveedorcodigo") = Trim(frmMantprovision.CtrAyu_Proveedor.xclave)
            .Parameters("@cabprovirznsoc") = Left(Trim(frmMantprovision.CtrAyu_Proveedor.xnombre), 50)
            .Parameters("@cabproviruc") = frmMantprovision.txRuc.Text
            .Parameters("@monedacodigo") = frmMantprovision.CtrAyu_Moneda.xclave
            .Parameters("@modoprovicod") = frmMantprovision.CtrAyu_Modoprovi.xclave
            .Parameters("@documetocodigo") = frmMantprovision.CtrAyu_TipDoc.xclave
            .Parameters("@cabprovictacte") = frmMantprovision.ChkCtaCte.Value
            .Parameters("@cabproviregcom") = frmMantprovision.ChkRegComp.Value
            '.Parameters("@cabproviestado") = 0
            .Parameters("@cabprovinumdoc") = Trim(frmMantprovision.TxSerie.Text) & IIf(Trim(frmMantprovision.TxSerie.Text) = "", "", "-") & Trim(frmMantprovision.TxNdoc.Text)
            '.Parameters("@cabprovinumord") = 0
            .Parameters("@cabprovifchdoc") = frmMantprovision.Dtp_FechaDoc.Value
            .Parameters("@cabprovifchven") = frmMantprovision.DtpFech_Ven.Value
            '.Parameters("@cabproviitem") = 0
            .Parameters("@tipocompracodigo") = frmMantprovision.CtrAyu_TipCompra.xclave
            .Parameters("@cabprovitotbru") = CDbl(frmMantprovision.TxTotBruto.valor)
            '.Parameters("@cabprovitotdcto") = 0
            '.Parameters("@cabprovitotven") = 0
            .Parameters("@cabprovitotigv") = CDbl(frmMantprovision.TxTotIGV.valor)
            .Parameters("@cabprovitotinaf") = CDbl(frmMantprovision.TxTotInafecto.valor)
            .Parameters("@cabprovitotal") = CDbl(frmMantprovision.TxTotImpCompra.valor)
            '.Parameters("@cabprovitotcxp") = 0
            '.Parameters("@cabprovitipigv") = 0
            .Parameters("@cabprovifchconta") = frmMantprovision.Dtp_FechaDoc.Value
            '.Parameters("@cabprovifchcancel") = 0
            '.Parameters("@cabprovifnupol") = 0
            If CDbl(VGvardllgen.ESNULO(frmMantprovision.lb_vcambio.Caption, 0)) = 0 Then
               .Parameters("@cabprovitipcambio") = 1
              Else
              .Parameters("@cabprovitipcambio") = CDbl(VGvardllgen.ESNULO(frmMantprovision.lb_vcambio.Caption, 0))
            End If
            .Parameters("@cabprovinumaux") = NumeroAux
'            .Parameters("@cabprovinumtes") = 0
             .Parameters("@usuariocodigo") = VGusuario
            .Parameters("@fechaact") = Now
            '.Parameters("@tiposubasicodigo") = frmMantprovision.CtrAyu_TipSubAsi.xclave
            .Parameters("@tiposubasicodigo") = "00"
            .Parameters("@cabprovitipdocref") = frmMantprovision.CtrAyu_TipRef.xclave
            .Parameters("@cabprovinref") = Trim(frmMantprovision.TxNref.Text)
            .Parameters("@cabprovifechdocref") = frmMantprovision.Dtp_FechaDocRef.Value
            .Parameters("@cabproviopergrab") = frmMantprovision.ChkOperGrab.Value
            '@cabprovioficina,@cabprovicaja,@cabprovifechcaja
            
            .Parameters("@cabprovioficina") = frmMantprovision.Ctr_AyudaOficina.xclave
            .Parameters("@cabprovicaja") = IIf(frmMantprovision.Ctr_AyudaCaja.Visible, frmMantprovision.Ctr_AyudaCaja.xclave, "")
            .Parameters("@cabprovifechcaja") = IIf(frmMantprovision.DTPFechaCaja.Visible, frmMantprovision.DTPFechaCaja.Value, Null)
            If frmMantprovision.IMant = 1 Then
                .Parameters("@cabproviflagmodi") = 0
              Else
                .Parameters("@cabproviflagmodi") = 1
            End If
            .Parameters("@cabprovinumtesor") = Trim(Numerotesor)
           If VGParametros.sistemamultiempresas = True Then
              .Parameters("@empresacodigo") = Trim(frmMantprovision.Ctr_Ayuempresa.xclave)
            Else
            .Parameters("@empresacodigo") = "00"
           End If
            
            If OP > 1 Then
              numero = Trim(frmMantprovision.lbNumComprobCab)
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
    Resume Next
End Sub
Public Function NumeroAuxiliar(mes As Integer, Optional ByRef numero As Long) As String
On Error GoTo Errnum
Dim rsaux As ADODB.Recordset
Dim cad As String
    Set rsaux = New ADODB.Recordset
    cad = "Select isnull(mes" & Trim(Format(mes, "00")) & ",0)+1 as numcorrelativo   From co_correlames " & _
          "Where Ano='" & VGParamSistem.Anoproceso & "'"
          
    rsaux.Open cad, VGcnx, adOpenKeyset, adLockReadOnly
               
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
Public Sub ActualizaCorrelComprob(ByVal numero As Double)
On Error GoTo Actualizacorre
    Set VGCommandoSP = New ADODB.Command
    VGCommandoSP.ActiveConnection = VGcnxMarfice
    VGCommandoSP.CommandType = adCmdStoredProc
    VGCommandoSP.CommandText = "co_actcorraux_pro"
    VGCommandoSP.Parameters.Refresh
    With VGCommandoSP
        .Parameters("@Base") = VGParamSistem.BDEmpresa
        .Parameters("@Ano") = VGParamSistem.Anoproceso
        .Parameters("@Mes") = Format(IIf(Month(frmMantprovision.DTPFechaContab) < Val(VGParamSistem.Mesproceso), Month(frmMantprovision.DTPFechaContab), CInt(VGParamSistem.Mesproceso)), "00")
        .Parameters("@numero") = numero
        .Execute
    End With
    Exit Sub
Actualizacorre:
    VGvarVerifica = False
    VGErrorString = "Error en Actualizar el Numero de Comprobante Auxiliar " & Chr(13) & Err.Description
End Sub

Public Sub ActualizaCorrelAuxiliar(ByVal numero As Double)
On Error GoTo Actualizacorre
    Set VGCommandoSP = New ADODB.Command
    VGCommandoSP.ActiveConnection = VGcnxMarfice
    VGCommandoSP.CommandType = adCmdStoredProc
    VGCommandoSP.CommandText = "co_actcorraux_pro"
    VGCommandoSP.Parameters.Refresh
    With VGCommandoSP
        .Parameters("@Base") = VGParamSistem.BDEmpresa
        .Parameters("@Ano") = VGParamSistem.Anoproceso
        .Parameters("@Mes") = Format(IIf(Month(frmMantprovision.DTPFechaContab) < Val(VGParamSistem.Mesproceso), Month(frmMantprovision.DTPFechaContab), CInt(VGParamSistem.Mesproceso)), "00")
        .Parameters("@numero") = numero
        .Execute
    End With
    Exit Sub
Actualizacorre:
    VGvarVerifica = False
    VGErrorString = "Error en Actualizar el Numero de Comprobante Auxiliar " & Chr(13) & Err.Description
End Sub
Public Sub GrabarDetalle(ByVal rs As Recordset, Optional ByVal numero As Long = 0)
On Error GoTo ErrorGrabaDetalle
Dim rsaux As ADODB.Recordset
Dim numerocomprobante As String
Set VGvardllgen = New dllgeneral.dll_general
    Set rsaux = rs.Clone(adLockReadOnly)
    rsaux.Filter = "(impbruto<>0 or impcompra <> 0)"
    Set VGCommandoSP = New ADODB.Command
    Set VGvardllgen = New dllgeneral.dll_general
    'Elimar los Detalle antes de grabar
    VGCommandoSP.ActiveConnection = VGcnxMarfice
    VGCommandoSP.CommandType = adCmdStoredProc
    VGCommandoSP.CommandText = "co_grabadetprovi"
    VGCommandoSP.Parameters.Refresh
    With VGCommandoSP
        .Parameters("@base") = VGParamSistem.BDEmpresa
        .Parameters("@tabla") = VGParamSistem.TablaDetcomprob
        .Parameters("@cabprovinumero") = numero
        If Month(frmMantprovision.DTPFechaContab) <> Val(VGParamSistem.Mesproceso) Then
           .Parameters("@cabprovimes") = Month(frmMantprovision.DTPFechaContab)
         Else
           .Parameters("@cabprovimes") = CInt(VGParamSistem.Mesproceso)
        End If
        .Parameters("@op") = 2
        .Execute
    End With
    rsaux.MoveFirst
    While Not rsaux.EOF
        With VGCommandoSP
            .Parameters("@cabprovinumero") = numero
            If Month(frmMantprovision.DTPFechaContab) <> Val(VGParamSistem.Mesproceso) Then
               .Parameters("@cabprovimes") = Month(frmMantprovision.DTPFechaContab)
             Else
               .Parameters("@cabprovimes") = CInt(VGParamSistem.Mesproceso)
            End If
            .Parameters("@op") = 1
            .Parameters("@detproviitem") = rsaux!Item
            '.Parameters("@detprovicod1") = 0
            '.Parameters("@detprovicod2") = 0
            '.Parameters("@detprovicod3") = 0
            '.Parameters("@detprovicod4") = 0
            .Parameters("@gastoscodigo") = rsaux!gastoscodigo
            .Parameters("@cuentacodigo") = rsaux!Cuentacodigo
            .Parameters("@detprovimon") = frmMantprovision.CtrAyu_Moneda.xclave
            '.Parameters("@detproviestado") = 0
            .Parameters("@detproviimpbru") = rsaux!impbruto
            .Parameters("@detproviimpigv") = rsaux!Igv
            .Parameters("@detproviimpina") = rsaux!Inafecto
            .Parameters("@detprovitotal") = rsaux!ImpCompra
            '.Parameters("@detprovidscto") = 0
            '.Parameters("@detproviimpdct") = 0
            '.Parameters("@detproviimpven") = 0
            '.Parameters("@detproviigv") = 0
            .Parameters("@detproviformcamb") = IIf(frmMantprovision.lb_vcambio.Visible, Format(frmMantprovision.CmbTcambio.ListIndex + 1, "00"), "00")
            .Parameters("@detprovitipcam") = IIf(frmMantprovision.lb_vcambio.Visible, CDbl(VGvardllgen.ESNULO(frmMantprovision.lb_vcambio.Caption, 0)), 0)
            .Parameters("@usuariocodigo") = VGusuario
            .Parameters("@fechaact") = Now
            .Parameters("@detproviglosa") = rsaux!glosa
            .Parameters("@detproviccosto") = Trim(VGvardllgen.ESNULO(rsaux!ccosto, "00"))
            .Parameters("@analitico") = Trim(VGvardllgen.ESNULO(rsaux!analitico, "00"))
            .Execute
        End With
        rsaux.MoveNext
    Wend
    Exit Sub
ErrorGrabaDetalle:
    VGvarVerifica = False
    VGErrorString = "Error en Grabar Detalle " & Chr(13) & Err.Description
 '   MsgBox (VGErrorString)
End Sub
Public Sub GrabarCP_Cargo(OP As Integer, Optional numero As Long = 0)
On Error GoTo ErrorGrabaCP
Dim rsaux As ADODB.Recordset
Dim numerocomprobante As String
'@base, @tipo, @tabla, @tipodocu, @numero, @cliente, @vendedor, @zona,
'@apefecemi, @moneda, @apeimppag, @usuario, @tipocambio, @fechaact, @flagcancel, @cargoabono, @concepto
Set VGvardllgen = New dllgeneral.dll_general
    Set VGCommandoSP = New ADODB.Command
    Set VGvardllgen = New dllgeneral.dll_general
    VGCommandoSP.ActiveConnection = VGcnxMarfice
    VGCommandoSP.CommandType = adCmdStoredProc
    VGCommandoSP.CommandText = "co_ingresacargo_pro"
    VGCommandoSP.Parameters.Refresh
    With VGCommandoSP
        .Parameters("@base") = VGParamSistem.BDEmpresa
        .Parameters("@tabla") = "CP_Cargo"
        .Parameters("@tipo") = OP
        If OP > 1 Then
            numero = Trim(frmMantprovision.lbNumComprobCab)
        End If
        .Parameters("@abonotipoplanilla") = VGParametros.CpTiplan
        .Parameters("@abononumplanilla") = Format(numero, "000000")
        .Parameters("@cliente") = Trim(frmMantprovision.CtrAyu_Proveedor.xclave)
        .Parameters("@tipodocu") = Trim(frmMantprovision.CtrAyu_TipDoc.xclave)
        .Parameters("@numero") = Format(Trim(frmMantprovision.TxSerie.Text), "000") & Format(Left(Trim(frmMantprovision.TxNdoc.Text), 8), "00000000")
        'Este para que al eliminar no utilizar estos parametros
        If OP <= 2 Then
            If OP = 2 Then
               .Parameters("@oldnumero") = Format(Trim(frmMantprovision.TxSerie.Tag), "000") & Format(Left(Trim(frmMantprovision.TxNdoc.Tag), 8), "00000000")
               .Parameters("@oldtipodocu") = frmMantprovision.CtrAyu_TipDoc.Tag
               .Parameters("@oldcliente") = frmMantprovision.CtrAyu_Proveedor.Tag
            End If
            .Parameters("@vendedor") = VGParametros.CpOficina
            .Parameters("@zona") = Null
            .Parameters("@apefecemi") = frmMantprovision.Dtp_FechaDoc.Value
            .Parameters("@moneda") = frmMantprovision.CtrAyu_Moneda.xclave
            .Parameters("@apeimppag") = CDbl(frmMantprovision.TxTotImpCompra.valor)
            .Parameters("@usuario") = VGusuario
            .Parameters("@tipocambio") = CDbl(frmMantprovision.lb_vcambio.Caption)
            .Parameters("@fechaact") = Now
            .Parameters("@flagcancel") = 0
            .Parameters("@cargoabono") = frmMantprovision.VlDocNota
            .Parameters("@concepto") = "01"
            .Parameters("@glosa") = Left(Trim(frmMantprovision.Txtglosa.valor), 80)
            If Trim(frmMantprovision.CtrAyu_TipRef.xclave) = Empty Then
               .Parameters("@cargoapetiporefe") = Trim(frmMantprovision.CtrAyu_TipRef.xclave)
               .Parameters("@cargoapenrorefe") = Trim(frmMantprovision.TxNref.Text)   'Trim(Left(frmMantprovision.TxNref.Text, 3)) & Format(Right(Trim(frmMantprovision.TxNref), Len(Trim(frmMantprovision.TxNref.Text) - 4)), "00000000") 'Format(Left(Trim(frmMantprovision.TxNref.Text), 8), "00000000")
            Else
               .Parameters("@cargoapetiporefe") = Trim(frmMantprovision.CtrAyu_TipRef.xclave)
               .Parameters("@cargoapenrorefe") = Trim(Left(frmMantprovision.TxNref.Text, 3)) & Format(Right(Trim(Trim(frmMantprovision.TxNref.Text)), Len(Trim(Trim(frmMantprovision.TxNref.Text))) - 4), "00000000")
            End If
            .Parameters("@cargoapefecvct") = Format(frmMantprovision.DtpFech_Ven.Value, "dd/mm/yyyy")
        End If
        .Execute
    End With
    Exit Sub
ErrorGrabaCP:
    VGvarVerifica = False
    VGErrorString = "Error en Grabar en Cuentas por Pagar " & Chr(13) & Err.Description
End Sub
Public Sub Grabaren_Tesoreria(ByVal OP As Integer, Optional ByVal Numeroprovi As Long = 0, Optional ByVal rs As Recordset, Optional ByRef Xnum As String)
'On Error GoTo ErrorGrabaTesore
Dim numero As String
Set VGvardllgen = New dllgeneral.dll_general
Dim rb As ADODB.Recordset
Dim Item As Integer
   'Obtener el Ultimo Numero Correlativo de las cajas
    Dim opaux As Integer
    opaux = OP
    Set VGCommandoSP = New ADODB.Command
    Set VGvardllgen = New dllgeneral.dll_general
    VGCommandoSP.ActiveConnection = VGcnxMarfice
    VGCommandoSP.CommandType = adCmdStoredProc
    VGCommandoSP.CommandText = "te_abonadocumento_pro"
    VGCommandoSP.Parameters.Refresh
    If OP = 2 Then
        If Trim(frmMantprovision.TxTesor.Text) = "" Then
            OP = 1
        End If
    End If
    With VGCommandoSP
        .Parameters("@base") = VGParamSistem.BDEmpresa
        .Parameters("@tipo") = IIf(OP = 3, 2, OP)
        If OP = 2 Or OP = 3 Then
            'Set rb = VGcnx.Execute("Select cabprovinumtesor  From " & VGParamSistem.TablaCabcomprob & " Where cabprovinumero=" & frmMantprovision.lbNumComprobCab.Caption)
            numero = frmMantprovision.TxTesor.Text
            
          Else
            Set rb = VGcnx.Execute("select * from te_parametroempresa where empresacodigo='01'")
            If rb.RecordCount > 0 Then
                numero = Format(CDbl(VGvardllgen.ESNULO(rb!empresanumegreso, "0")) + 1, "000000")
                VGcnx.Execute "Update te_parametroempresa Set empresanumegreso='" & numero & "' where empresacodigo='01'"
                'VGcnx.Execute "Update " & VGParamSistem.TablaCabcomprob & " Set cabprovinumtesor='" & Numero & _
                '              "' Where cabprovinumero=" & Numeroprovi
            End If
        End If
        Xnum = numero
        .Parameters("@estadoreg") = ""
        .Parameters("@numrecibo") = numero
        If OP = 3 Then
            .Execute
        End If
        If OP = 2 Then
            .Execute
            .Parameters("@tipo") = 1
           OP = 1
        End If
        
        'Este para que al eliminar no utilizar estos parametros
        If OP = 1 Then
            .Parameters("@controlctacte") = "N"
            .Parameters("@vendedorcodigo") = frmMantprovision.Ctr_AyudaOficina.xclave
            .Parameters("@cajacodigo") = frmMantprovision.Ctr_AyudaCaja.xclave
            .Parameters("@clientecodigo") = frmMantprovision.CtrAyu_Proveedor.xclave
            .Parameters("@descripcion") = ""
            .Parameters("@operacion") = frmMantprovision.CtrAyu_Modoprovi.xclave
            .Parameters("@monedacodigo") = frmMantprovision.CtrAyu_Moneda.xclave
            .Parameters("@ingsal") = "E"
            .Parameters("@tipocambio") = CDbl(frmMantprovision.lb_vcambio.Caption)
            .Parameters("@totsoles") = IIf(frmMantprovision.CtrAyu_Moneda.xclave = "01", CDbl(frmMantprovision.TxTotImpCompra.valor), Round(CDbl(frmMantprovision.TxTotImpCompra.valor) * CDbl(frmMantprovision.lb_vcambio.Caption), 2))
            .Parameters("@totdolares") = IIf(frmMantprovision.CtrAyu_Moneda.xclave <> "01", CDbl(frmMantprovision.TxTotImpCompra.valor), Round(CDbl(frmMantprovision.TxTotImpCompra.valor) / CDbl(frmMantprovision.lb_vcambio.Caption), 2))
            .Parameters("@fechadocumento") = frmMantprovision.DTPFechaCaja.Value
            .Parameters("@observa") = ""
            .Parameters("@transferauto") = ""
            .Parameters("@numreciboegreso") = ""
            .Parameters("@usuario") = VGusuario
            .Parameters("@fechaact") = Now
            .Parameters("@empresacodigo") = frmMantprovision.Ctr_Ayuempresa.xclave
            .Execute
        End If
    End With
OP = opaux
Set VGCommandoSP = New ADODB.Command
VGCommandoSP.ActiveConnection = VGcnxMarfice
VGCommandoSP.CommandType = adCmdStoredProc
VGCommandoSP.CommandText = "te_abonadetalledocumento_pro"
VGCommandoSP.Parameters.Refresh
If OP = 3 Then
   With VGCommandoSP
        .Parameters("@base") = VGParamSistem.BDEmpresa
        .Parameters("@tipo") = IIf(OP = 3, 2, OP)
        .Parameters("@numrecibo") = numero
        If OP = 3 Then
            .Execute
        End If
   End With
Else
  rs.MoveFirst
  Item = 1
  While Not rs.EOF
      With VGCommandoSP
        .Parameters("@base") = VGParamSistem.BDEmpresa
        .Parameters("@tipo") = IIf(OP = 3, 2, OP)
        .Parameters("@numrecibo") = numero
        If OP = 3 Then
            .Execute
        End If
        If OP = 2 Then
            If Item = 1 Then .Execute
            .Parameters("@tipo") = 1
            OP = 1
        End If
        If OP = 1 Then
            .Parameters("@estadoreg") = ""
            .Parameters("@item") = Item
            .Parameters("@emisioncheque") = "C"
            .Parameters("@tipodocconcepto") = frmMantprovision.CtrAyu_TipDoc.xclave
            .Parameters("@numdocumento") = Format(Trim(frmMantprovision.TxSerie.Text), "000") & Format(Left(Trim(frmMantprovision.TxNdoc.Text), 8), "00000000")
            .Parameters("@carabo") = frmMantprovision.VlDocNota
            .Parameters("@formacan") = ""
            .Parameters("@tdqc") = ""
            .Parameters("@ndqc") = ""
            .Parameters("@tipocajabanco") = "C"
            .Parameters("@cajabanco") = frmMantprovision.Ctr_AyudaCaja.xclave
            .Parameters("@numctacte") = ""
            .Parameters("@adicionactacte") = "P"
            .Parameters("@monedadocumento") = frmMantprovision.CtrAyu_Moneda.xclave
            .Parameters("@monedacancela") = frmMantprovision.CtrAyu_Moneda.xclave
            
            
            .Parameters("@importesoles") = IIf(frmMantprovision.CtrAyu_Moneda.xclave = "01", CDbl(rs.Fields("impcompra")), Round(CDbl(rs.Fields("ImpCompra")) * CDbl(frmMantprovision.lb_vcambio.Caption), 2))
            .Parameters("@importedolares") = IIf(frmMantprovision.CtrAyu_Moneda.xclave <> "01", CDbl(rs.Fields("ImpCompra")), Round(CDbl(rs.Fields("ImpCompra") / CDbl(frmMantprovision.lb_vcambio.Caption)), 2))
            .Parameters("@contabledisponi") = "S"
            .Parameters("@fechacancela") = frmMantprovision.DTPFechaCaja.Value
            .Parameters("@observacion") = ""
            .Parameters("@gastos") = rs.Fields("gastoscodigo")
            .Parameters("@usuario") = VGusuario
            .Parameters("@fechaact") = Now
            .Parameters("@entidad") = rs.Fields("analitico")
            .Parameters("@centrocosto") = rs.Fields("ccosto")
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
Public Sub MostrarCabecera(ByVal rs As ADODB.Fields)
Dim rsaux As ADODB.Fields, pos As Integer
    Set rsaux = rs
    Set VGvardllgen = New dll_general
    With frmMantprovision
        .lbNumComprobCab.Caption = rsaux!cabprovinumero
        .CtrAyu_Proveedor.xclave = rsaux!proveedorcodigo: .CtrAyu_Proveedor.Ejecutar
        .CtrAyu_Moneda.xclave = rsaux!monedacodigo: .CtrAyu_Moneda.Ejecutar
        .CtrAyu_Modoprovi.xclave = rsaux!modoprovicod: .CtrAyu_Modoprovi.Ejecutar
        .CtrAyu_TipDoc.xclave = rsaux!documetocodigo: .CtrAyu_TipDoc.Ejecutar
        .ChkCtaCte.Value = IIf(rsaux!cabprovictacte = 0, 0, 1)
        .ChkRegComp.Value = IIf(rsaux!cabproviregcom = 0, 0, 1)
        .ChkOperGrab.Value = IIf(rsaux!cabproviopergrab = 0, 0, 1)
        'La Serie y El Numero de Documento
         pos = InStr(1, rsaux!cabprovinumdoc, "-", vbTextCompare)
        .TxSerie.Text = Trim(Left(rsaux!cabprovinumdoc, IIf(pos = 0, 0, pos - 1)))
        .TxNdoc.Text = Right(rsaux!cabprovinumdoc, Len(rsaux!cabprovinumdoc) - pos)
        
        'Leer valores anteriores de Tipo Doc, Nº Doc y Cod. Cliente
        .CtrAyu_TipDoc.Tag = Trim(rsaux!documetocodigo)
        .TxSerie.Tag = Trim(Left(rsaux!cabprovinumdoc, IIf(pos = 0, 0, pos - 1)))
        .TxNdoc.Tag = Right(rsaux!cabprovinumdoc, Len(rsaux!cabprovinumdoc) - pos)
        .CtrAyu_Proveedor.Tag = Trim(rsaux!proveedorcodigo)
                
        .Dtp_FechaDoc.Value = rsaux!cabprovifchdoc
        .DtpFech_Ven.Value = IIf(IsNull(rsaux!cabprovifchven), rsaux!cabprovifchdoc, rsaux!cabprovifchven)
        .CtrAyu_TipCompra.xclave = rsaux!tipocompracodigo: .CtrAyu_TipCompra.Ejecutar
        .CtrAyu_TipSubAsi.xclave = rsaux!tiposubasicodigo: .CtrAyu_TipSubAsi.Ejecutar
        .CtrAyu_TipRef.xclave = rsaux!cabprovitipdocref: .CtrAyu_TipRef.Ejecutar
        .TxNref.Text = VGvardllgen.ESNULO(rsaux!cabprovinref, "")
        .Dtp_FechaDocRef.Value = VGvardllgen.ESNULO(rsaux!cabprovifechdocref, Null)
        .TxTotBruto.Text = Format(rsaux!cabprovitotbru, "###,###,###.00"): .TxTotBruto.valor = Format(rsaux!cabprovitotbru, "#0.00")
        .TxTotIGV.Text = Format(rsaux!cabprovitotigv, "###,###,###.00"): .TxTotIGV.valor = Format(rsaux!cabprovitotigv, "#0.00")
        .TxTotInafecto.Text = Format(rsaux!cabprovitotinaf, "###,###,###.00"): .TxTotInafecto.valor = Format(rsaux!cabprovitotinaf, "#0.00")
        .TxTotImpCompra.Text = Format(rsaux!cabprovitotal, "###,###,###.00"): .TxTotImpCompra.valor = Format(rsaux!cabprovitotal, "#0.00")
        .Dtp_FechaDoc.Value = rsaux!cabprovifchconta
        .TxNAux.Text = Right(Trim(rsaux!cabprovinumaux), 5)
        .lb_vcambio.Caption = Format(rsaux!cabprovitipcambio, "#0.000")
        
        .Ctr_AyudaOficina.xclave = VGvardllgen.ESNULO(rsaux!cabprovioficina, "00"): .Ctr_AyudaOficina.Ejecutar
        .Ctr_AyudaCaja.xclave = VGvardllgen.ESNULO(rsaux!cabprovicaja, "00"): .Ctr_AyudaCaja.Ejecutar
        .DTPFechaCaja.Value = IIf(.DTPFechaCaja.Visible, VGvardllgen.ESNULO(rsaux!cabprovifechcaja, Date), Date)
        .TxTesor.Text = VGvardllgen.ESNULO(rsaux!cabprovinumtesor, "")
        .VlComprob_Conta = VGvardllgen.ESNULO(rsaux!cabprovinconta, "''''")
        .Ctr_Ayuempresa.xclave = VGvardllgen.ESNULO(rsaux!empresacodigo, "00")
    End With
End Sub
Public Sub MostrarDetalle(ByRef rs As ADODB.Recordset)
Dim rsaux As ADODB.Recordset, i As Integer
    Set rsaux = New ADODB.Recordset
    Set VGCommandoSP = New ADODB.Command
    Set VGvardllgen = New dllgeneral.dll_general
    'Elimar los Detalle antes de grabar
    VGCommandoSP.ActiveConnection = VGcnxMarfice
    VGCommandoSP.CommandType = adCmdStoredProc
    VGCommandoSP.CommandText = "co_grabadetprovi"
    VGCommandoSP.Parameters.Refresh
    With VGCommandoSP
        .Parameters("@base") = VGParamSistem.BDEmpresa
        .Parameters("@tabla") = VGParamSistem.TablaDetcomprob
        .Parameters("@cabprovinumero") = frmMantprovision.lbNumComprobCab
        .Parameters("@op") = 3
        Set rsaux = .Execute
    End With
    If rsaux.RecordCount() > 0 Then
       rsaux.MoveFirst
    End If
    While Not rsaux.EOF
        rs.AddNew
        For i = 0 To rsaux.Fields.Count - 1
            rs.Fields(rsaux.Fields(i).Name) = IIf(IsNull(rsaux.Fields(i).Value), "0", rsaux.Fields(i).Value)
        Next
        rs.Update
        rsaux.MoveNext
    Wend
    If rs.RecordCount > 0 Then
        rs.AbsolutePosition = 1
        frmMantprovision.TDBG_Det.SetFocus
        frmMantprovision.lbnregdetalle.Caption = Format(rs.RecordCount, "0 ")
      Else
        frmMantprovision.lbnregdetalle.Caption = Format(0, "0 ")
    End If
End Sub
Public Function ValidarGrabarCabecera(Nr As Long) As Boolean
    ValidarGrabarCabecera = False
With frmMantprovision
    'Validando Que exista por lo menos un registro en el detalle
    Set VGvardllgen = New dllgeneral.dll_general
    If Nr = 0 Then
        MsgBox "Por lo menos debe haber ingresado un registro de detalle", vbInformation
        frmMantprovision.CtrAyu_Modoprovi.SetFocus
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
    If Trim(.TxSerie.Text) = "" Then
        MsgBox "Tiene que ingresar la Serie del Documento", vbInformation
        .TxSerie.SetFocus
        Exit Function
    End If
    If Trim(.TxNdoc.Text) = "" Then
        MsgBox "Tiene que ingresar el Numero de Documento", vbInformation
        .TxNdoc.SetFocus
        Exit Function
    End If
    If Trim(.CtrAyu_Moneda.xclave) = "" Or .CtrAyu_Moneda.xclave = "00" Then
        MsgBox "Tiene que seleccionar una moneda ", vbInformation
        .CtrAyu_Moneda.SetFocus
        Exit Function
    End If
    If Trim(.CtrAyu_TipCompra.xclave) = "" Or .CtrAyu_TipCompra.xclave = "00" Then
        MsgBox "Tiene que seleccionar una moneda ", vbInformation
        .CtrAyu_TipCompra.SetFocus
        Exit Function
    End If
    
    If VGvardllgen.ESNULO(.lb_vcambio.Caption, 0) = 0 Then
        MsgBox "Debe de escoger una fecha que exista tipo de cambio", vbInformation
        Exit Function
    End If
    
    If .VlDocNota = "A" Then
        If Trim(.CtrAyu_TipRef.xclave) = "" Or .CtrAyu_TipRef.xclave = "00" Then
            MsgBox "Tiene que Seleccionar un tipo de documento de referencia", vbInformation
            .CtrAyu_TipRef.SetFocus
            Exit Function
        End If
        If Trim(.TxNref.Text) = "" Then
            MsgBox "Tiene que Ingresar el Numero de Documento ", vbInformation
            .TxNref.SetFocus
            Exit Function
        End If
        If IsNull(.Dtp_FechaDocRef.Value) Then
            MsgBox "Tiene que Ingresar una fecha de referencia ", vbInformation
            .Dtp_FechaDocRef.SetFocus
            Exit Function
        End If
    End If
    
    
    
    
    'Validando Que el Documento no se repita para un proveedor
    Dim DocAct As String
    DocAct = Trim(frmMantprovision.CtrAyu_Proveedor.xclave) & "-" & Trim(frmMantprovision.CtrAyu_TipDoc.xclave) & "-" & Trim(frmMantprovision.TxSerie.Text) & IIf(Trim(frmMantprovision.TxSerie.Text) = "", "", "-") & Trim(frmMantprovision.TxNdoc.Text)
    If (frmMantprovision.IMant = 1) Or (frmMantprovision.VlDocAnt <> DocAct) Then
        If ExisteSQL(VGcnx, " Select * From dbo.co_cabprovi" & VGParamSistem.Anoproceso & _
                     " Where " & _
                     " isnull(proveedorcodigo,'')+'-'+isnull(documetocodigo,'')+'-'+cabprovinumdoc='" & DocAct & "'") Then
           MsgBox "Este Documento ya ha sido ingresado para este proveedor ", vbExclamation
           frmMantprovision.TxNdoc.SetFocus
           Exit Function
        End If
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
With frmMantprovision
    If VGParametros.sistemactrlgastos Then
       If (.CtrAyu_gastos.Enabled And .CtrAyu_gastos.Visible) And (Trim(.CtrAyu_gastos.xclave) = "" Or Trim(.CtrAyu_gastos.xclave) = "00") Then
           MsgBox "Debe ingresar la Cuenta de gastos ", vbInformation
          .CtrAyu_gastos.SetFocus
          Exit Function
       End If
    Else
       If (.CtrAyu_Cuenta.Enabled And .CtrAyu_Cuenta.Visible) And (Trim(.CtrAyu_Cuenta.xclave) = "" Or Trim(.CtrAyu_Cuenta.xclave) = "00") Then
           MsgBox "Debe ingresar la Cuenta", vbInformation
          .CtrAyu_Cuenta.SetFocus
          Exit Function
       End If
    End If
    If CDbl(VGvardllgen.ESNULO(Espunto(.TxImpBruto.valor), 0)) = 0 And CDbl(VGvardllgen.ESNULO(Espunto(.TxInafecto.valor), 0)) = 0 Then
        MsgBox "Por lo menos el Importe Bruto o el Importe Inafecto tiene que ser mayor a cero"
        .TxImpBruto.SetFocus
        Exit Function
    End If
    If CDbl(VGvardllgen.ESNULO(Espunto(.TxImpCompra.valor), 0)) = 0 Then
        MsgBox "El Importe de Compra no debe ser Cero"
        .TxImpCompra.SetFocus
        Exit Function
    End If
    If (.CtrAyu_Ccosto.Enabled And .CtrAyu_Ccosto.Visible) And (Trim(.CtrAyu_Ccosto.xclave) = "" Or Trim(.CtrAyu_Ccosto.xclave) = "00") Then
        MsgBox "Debe ingresar El cento de costo", vbInformation
        .CtrAyu_Ccosto.SetFocus
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
    With frmMantprovision
        .lbNumComprobCab.Caption = ""
        .CtrAyu_Modoprovi.xclave = "": .CtrAyu_Modoprovi.xnombre = ""
        .CtrAyu_Proveedor.xclave = "": .CtrAyu_Proveedor.xnombre = ""
        .CtrAyu_TipDoc.xclave = "": .CtrAyu_TipDoc.xnombre = ""
        .CtrAyu_TipRef.xclave = "": .CtrAyu_TipRef.xnombre = ""
        .txRuc.Text = "": .lbTelef.Caption = ""
        .TxNref.Text = "": .Dtp_FechaDocRef.Value = Null
        .CtrAyu_Moneda.xclave = "": .CtrAyu_Moneda.xnombre = ""
        .CtrAyu_TipCompra.xclave = "": .CtrAyu_TipCompra.xnombre = ""
        .ChkCtaCte.Value = 0: .ChkRegComp.Value = 0: .ChkActCaja.Value = 0
        .CtrAyu_TipSubAsi.xclave = "": .CtrAyu_TipSubAsi.xnombre = ""
        .Ctr_AyudaCaja.xclave = "": .Ctr_AyudaCaja.xnombre = ""
        .Ctr_AyudaOficina.xclave = "": .Ctr_AyudaOficina.xnombre = ""
        'Limpiar Totales
        .TxTotBruto.Text = "": .TxTotBruto.valor = ""
        .TxTotIGV.Text = "": .TxTotIGV.valor = ""
        .TxTotInafecto.Text = "": .TxTotInafecto.valor = ""
        .TxTotImpCompra.Text = "": .TxTotImpCompra.valor = ""
        .ChkOperGrab.Value = 1
        .Ctr_Ayuempresa.xclave = "": .Ctr_Ayuempresa.xnombre = ""
    End With
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

