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
    Ccosto = 5
    impbruto = 6
    igv = 7
    inafecto = 8
    Impcompra = 9

End Enum

Public Sub CargarAyudas()
With frmMantprovision
    Call .CtrAyu_Cuenta.conexion(VGCnxCT): .CtrAyu_Cuenta.filtro = "(cuentanivel=" & VGnumniveles & " and cuentacodigo <>'00') and (" & VGParametros.ctascompra & ")"
    Call .CtrAyu_gastos.conexion(VGCNx): .CtrAyu_gastos.filtro = "(gastosnivel=" & VGnumnivgas & " and gastoscodigo <>'00' and gastosestado<>1 ) "
    Call .CtrAyu_TipDoc.conexion(VGCNx): .CtrAyu_TipDoc.filtro = "tdocumentocodigo<>'00'"
    Call .CtrAyu_TipRef.conexion(VGCNx): .CtrAyu_TipRef.filtro = "tdocumentocodigo<>'00'"
    Call .CtrAyu_Moneda.conexion(VGCnxCT): .CtrAyu_Moneda.filtro = "monedacodigo<>'00'"
    Call .CtrAyu_Modoprovi.conexion(VGCNx): .CtrAyu_Modoprovi.filtro = "modoprovicod <>'00'"
    Call .CtrAyu_TipCompra.conexion(VGCNx)
    Call .CtrAyu_TipSubAsi.conexion(VGCNx)
    Call .Ctr_AyuAnalitico.conexion(VGCNx)
    
    Call .Ctr_AyuEmpresaPlanillas.conexion(VGCNx)
    Call .Ctr_AyuEmpresaactivos.conexion(VGCNx)
    
    
    Call .CtrAyu_Proveedor.conexion(VGCNx): .CtrAyu_Proveedor.filtro = "clientecodigo <>'00'"
    
    Call .Ctr_AyudaCaja.conexion(VGCNx)
    .Ctr_AyudaCaja.filtro = " isnull(CajaCuentaxRendir,0)=" & .m_cuentasxrendir & " and isnull(Cajafondofijo,0)=" & .m_fondofijo
    
    Call .Ctr_Ayutransf.conexion(VGCNx)
    
    Call .Ctr_AyudaOficina.conexion(VGCNx)
    Call .CtrAyu_Ccosto.conexion(VGCnxCT)
    Call .Ctr_Ayuempresa.conexion(VGCNx): .Ctr_Ayuempresa.filtro = "empresacodigo<>'00'"
    If VGParametros.sistemamultiempresas Then
       .CtrAyu_Ccosto.filtro = "empresacodigo='" & .Ctr_Ayuempresa.xclave & "' and centrocostonivel='" & VGnumnivcos & "' and centrocostocodigo<>'00' "
      Else
        .CtrAyu_Ccosto.filtro = "empresacodigo='" & VGParametros.empresacodigo & "' and centrocostonivel='" & VGnumnivcos & "' and centrocostocodigo<>'00' "
    End If
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
Public Sub ActualizarDetalle(ByRef rs As ADODB.Recordset, op As Campos)
    Dim vardllgen As New dllgeneral.dll_general
    If rs Is Nothing Then Exit Sub
    If rs.RecordCount = 0 Then Exit Sub
       VGactulizodoc = False
    If VGMoverRegistro Then Exit Sub
    Screen.MousePointer = 11
    With rs
        '!detcomprobitem = Format(Val(vardllgen.ESNULO(MaxItem(rs, "detcomprobitem"), 0) + 1), "00000")
        Select Case op
            Case cuenta
                !Cuentacodigo = frmMantprovision.CtrAyu_Cuenta.xclave
                !cuentades = frmMantprovision.CtrAyu_Cuenta.xnombre
                !Ccosto = "00"
                !gastoscodigo = "00"
           Case gastos
                    !gastoscodigo = frmMantprovision.CtrAyu_gastos.xclave
                !cuentades = frmMantprovision.CtrAyu_gastos.xnombre
                !Ccosto = "00"
                !Cuentacodigo = Cuentacodigo
            Case glosa: !glosa = Trim(frmMantprovision.Txtglosa.Text)
            Case analitico: !analitico = frmMantprovision.Ctr_AyuAnalitico.xclave
            Case Ccosto: !Ccosto = frmMantprovision.CtrAyu_Ccosto.xclave
            Case impbruto: !impbruto = vardllgen.ESNULO(Espunto(frmMantprovision.TxImpBruto.valor), 0)
            Case igv: !igv = vardllgen.ESNULO(Espunto(frmMantprovision.TxIGV.valor), 0)
            Case inafecto: !inafecto = vardllgen.ESNULO(Espunto(frmMantprovision.TxInafecto.valor), 0)
            Case Impcompra: !Impcompra = vardllgen.ESNULO(Espunto(frmMantprovision.TxImpCompra.valor), 0)
            

        
        End Select
    End With
    
    Screen.MousePointer = 0
End Sub
Private Function MaxItem(ByVal rs As ADODB.Recordset, ByVal campo As String) As Variant
Dim rsauxiliar As New ADODB.Recordset
Set rsauxiliar = rs.Clone(adLockReadOnly)
On Error GoTo errormaxitem
    rsauxiliar.Sort = campo & " desc"
    rsauxiliar.MoveFirst
    MaxItem = rsauxiliar.Fields(campo).Value
    Exit Function
errormaxitem:
    MaxItem = Null
End Function
Public Sub GeneraAsientoenLine(ByVal op As Integer, ByVal Nprovi As String, ByVal Comprob_Contable As String)
On Error GoTo genasiento
    VGvarVerifica = True
    VGCNx.BeginTrans

    SQL = " Update " & VGParamSistem.TablaCabcomprob & " Set cabprovinconta=''"
    SQL = SQL & " Where cabproviano='" & VGParamSistem.AnoProceso & "' and cabprovimes=" & VGParamSistem.MesProceso
    SQL = SQL & " and cabprovinumero=" & Nprovi & ""
    VGCNx.Execute SQL
    
    SQL = " Delete from ct_cabcomprob" & VGParamSistem.AnoProceso & " Where cabcomprobmes=" & VGParamSistem.MesProceso
    SQL = SQL & " and cabcomprobnprovi='COM" & Nprovi & "'"
    SQL = SQL & " and empresacodigo='" & frmMantprovision.Ctr_Ayuempresa.xclave & "'"
    VGCnxCT.Execute SQL
    
    SQL = " Update " & VGParamSistem.TablaCabcomprob & " Set cabproviflagmodi = 0 "
    SQL = SQL & " Where cabproviano='" & VGParamSistem.AnoProceso & "' and cabprovimes=" & VGParamSistem.MesProceso
    SQL = SQL & " and cabprovinumero=" & Nprovi & ""
    VGCNx.Execute SQL

    VGCNx.CommitTrans
    VGCNx.BeginTrans

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
        .Parameters("@Mes") = VGParamSistem.MesProceso
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
    VGCommandoSP.CommandText = "co_generaasientocomprasenlinea_pro"
    VGCommandoSP.Parameters.Refresh
    With VGCommandoSP
        .Parameters("@BaseConta") = VGParamSistem.BDEmpresaCT
        .Parameters("@BaseCompra") = VGParamSistem.BDEmpresa
        If VGParametros.sistemamultiempresas Then
          .Parameters("@empresa") = frmMantprovision.Ctr_Ayuempresa.xclave
          Else
            .Parameters("@empresa") = VGParametros.empresacodigo
        End If
        .Parameters("@SubAsiento") = VGParametros.xsubasiento
        .Parameters("@Libro") = VGParametros.xLibro
        If Month(frmMantprovision.DTPFechaContab) <> Val(VGParamSistem.MesProceso) Then
           .Parameters("@mes") = Format(Month(frmMantprovision.DTPFechaContab), "00")
         Else
           .Parameters("@mes") = Format(VGParamSistem.MesProceso, "00")
        End If
        .Parameters("@Ano") = VGParamSistem.AnoProceso
        .Parameters("@ctatotal") = VGParametros.xCtaTotal
        .Parameters("@ctaIGV") = VGParametros.xCtaIGV
        .Parameters("@ctaIES") = VGParametros.xCtaIES
        .Parameters("@ctaRTA") = VGParametros.xCtaRTA
        .Parameters("@tipanal") = VGParametros.xTipAnal
        .Parameters("@Compu") = VGcomputer
        .Parameters("@Usuario") = VGParamSistem.Usuario
        .Parameters("@Oficina") = Format(VGParametros.CpOficina, "00")
        .Parameters("@numcomprob") = Nprovi
        .Parameters("@tipo") = 1
        .Parameters("@tipoasientoAuto") = VGParametros.AsientoAutoxCCostos
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
        If Month(frmMantprovision.DTPFechaContab) <> Val(VGParamSistem.MesProceso) Then
           .Parameters("@mes") = Format(Month(frmMantprovision.DTPFechaContab), "00")
         Else
           .Parameters("@mes") = Format(VGParamSistem.MesProceso, "00")
        End If
        .Parameters("@Ano") = VGParamSistem.AnoProceso
        .Parameters("@Nprovi") = Nprovi
        .Execute
    End With
    VGCNx.CommitTrans
    Exit Sub
genasiento:

    Screen.MousePointer = 1
    VGvarVerifica = False
    VGErrorString = "Error en Grabar Cabecera " & Chr(13) & Err.Description
End Sub
Public Sub asientotesoreriaenlinea(ByVal op As Integer, recibo As String, comprobconta As String, tipocajabanco As String, moneda As String)
Dim Comando As ADODB.Command
Dim rsparimpo As New ADODB.Recordset
VGvarVerifica = True
On Error GoTo genasiento
Screen.MousePointer = 11
Set rsparimpo = New ADODB.Recordset
Set rsparimpo = VGCNx.Execute("Select * From  ct_importartesoreria Where Left(Upper(tipooperacion),1) <>'T' AND Left(Upper(tipocajabanco),1) ='" & tipocajabanco & "' And monedacodigo='" & moneda & "' ")
If rsparimpo.RecordCount() = 0 Then
   Screen.MousePointer = 1
   MsgBox "Verifique el parametro del asiento de " & recibo & " en ct_importartesoreria ", vbInformation, "Sistema de Tesoreria"
   VGvarVerifica = False
   Exit Sub
End If
VGCNx.BeginTrans
Set VGCommandoSP = New ADODB.Command
Set VGvardllgen = New dllgeneral.dll_general
    Set Comando = New ADODB.Command
    With Comando
         .CommandType = adCmdStoredProc
         .CommandText = "te_GeneraAsientosTesoreriaLinea_pro"
         .CommandTimeout = 0
         .ActiveConnection = VGGeneral
         .Parameters.Refresh
         .Parameters("@BaseConta") = VGCnxCT.DefaultDatabase
         .Parameters("@BaseVenta") = VGCNx.DefaultDatabase
         If VGParametros.sistemamultiempresas = True Then
            .Parameters("@empresa") = frmMantprovision.Ctr_Ayuempresa.xclave
           Else
           .Parameters("@empresa") = VGParametros.empresacodigo
         End If
         .Parameters("@Asiento") = rsparimpo!Asiento
         .Parameters("@SubAsiento") = rsparimpo!SubAsiento
         .Parameters("@Libro") = rsparimpo!libro
         .Parameters("@Mes") = Format(VGParamSistem.MesProceso, "00")
         .Parameters("@Ano") = VGParamSistem.AnoProceso
         .Parameters("@tipanal") = "002"
         .Parameters("@Compu") = VGcomputer
         .Parameters("@Usuario") = VGUsuario
         .Parameters("@TipoMov") = "E"
         .Parameters("@Nrecibo") = recibo
         .Parameters("@op") = op
         .Parameters("@comprobconta") = comprobconta
         .Parameters("@ajustehaber") = VGParametros.sistemactaajustehab
         .Parameters("@ajustedebe") = VGParametros.sistemactaajustedeb
         .Execute
     End With
VGCNx.CommitTrans
Exit Sub
genasiento:
    op = 1
    VGvarVerifica = False
    Screen.MousePointer = 1
    MsgBox "Hubo Errores al momento que se genero el recibo " & recibo & Chr(13) & Err.Description
End Sub
Public Sub GrabarCabecera(ByVal op As Integer, Optional ByVal numero As Long, Optional ByVal NumeroAux As String, Optional ByVal Numerotesor As String)
On Error GoTo ErrorGrabaCabecera
VGvarVerifica = True
VGCNx.BeginTrans
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
        If Month(frmMantprovision.DTPFechaContab) <> Val(VGParamSistem.MesProceso) Then
            .Parameters("@cabprovimes") = Month(frmMantprovision.DTPFechaContab)
            .Parameters("@cabproviano") = Year(frmMantprovision.DTPFechaContab)
         Else
            .Parameters("@cabproviano") = VGParamSistem.AnoProceso
            .Parameters("@cabprovimes") = Val(VGParamSistem.MesProceso)
        End If

        .Parameters("@cabprovinumero") = numero
        'Este para que al eliminar no utilizar estos parametros
        If op <= 2 Then
            .Parameters("@proveedorcodigo") = Trim(frmMantprovision.CtrAyu_Proveedor.xclave)
            .Parameters("@cabprovirznsoc") = Left(Trim(frmMantprovision.CtrAyu_Proveedor.xnombre), 50)
            .Parameters("@cabproviruc") = frmMantprovision.txRuc.Text
            .Parameters("@monedacodigo") = frmMantprovision.CtrAyu_Moneda.xclave
            .Parameters("@modoprovicod") = frmMantprovision.CtrAyu_Modoprovi.xclave
            .Parameters("@documetocodigo") = frmMantprovision.CtrAyu_TipDoc.xclave
            .Parameters("@cabprovictacte") = frmMantprovision.ChkCtaCte.Value
            .Parameters("@cabproviregcom") = frmMantprovision.ChkRegComp.Value
            '.Parameters("@cabproviestado") = 0
            .Parameters("@cabprovinumdoc") = Format(frmMantprovision.TxSerie.Text, "0000") & Format(Trim(frmMantprovision.TxNdoc.Text), "0000000000")
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
            .Parameters("@cabprovifchconta") = frmMantprovision.DTPFechaContab.Value
            '.Parameters("@cabprovifchcancel") = 0
            '.Parameters("@cabprovifnupol") = 0
            If CDbl(VGvardllgen.ESNULO(frmMantprovision.lb_vcambio.Caption, 0)) = 0 Then
               .Parameters("@cabprovitipcambio") = 1
              Else
              .Parameters("@cabprovitipcambio") = CDbl(VGvardllgen.ESNULO(frmMantprovision.lb_vcambio.Caption, 0))
            End If
            .Parameters("@cabprovinumaux") = NumeroAux
'            .Parameters("@cabprovinumtes") = 0
             .Parameters("@usuariocodigo") = VGUsuario
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
             .Parameters("@empresacodigo") = VGParametros.empresacodigo
           End If
            
            If op > 1 Then
              numero = Trim(frmMantprovision.lbNumComprobCab)
             .Parameters("@cabprovinumero") = numero
            End If
        End If
    End With
    VGCommandoSP.Execute
VGCNx.CommitTrans
    Exit Sub
ErrorGrabaCabecera:
    VGvarVerifica = False
    MsgBox ("Error en Grabar Cabecera " & Chr(13) & Err.Description)
    Exit Sub
    Resume Next
End Sub
Public Function NumeroAuxiliar(ByRef empresa As String, ByRef libro As String, ByRef anno As String, ByRef mes As String, ByRef numero As Long) As String
On Error GoTo Errnum
Dim rsaux As ADODB.Recordset
Dim cad As String
    Set rsaux = New ADODB.Recordset
    cad = "Select isnull(libronumcorr" & Format(mes, "00") & ",0)+1 as numcorrelativo   From ct_librocorre " & _
          "Where empresacodigo='" & empresa & "' and  librocodigo='" & libro & "' and libroAnno='" & anno & "'"
    Set rsaux = VGCnxCT.Execute(cad)
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
VGvarVerifica = True
VGCNx.BeginTrans
On Error GoTo Actualizacorre
    Set VGCommandoSP = New ADODB.Command
    VGCommandoSP.ActiveConnection = VGGeneral
    VGCommandoSP.CommandType = adCmdStoredProc
    VGCommandoSP.CommandText = "co_actcorraux_pro"
    VGCommandoSP.Parameters.Refresh
    With VGCommandoSP
        .Parameters("@Base") = VGParamSistem.BDEmpresa
        .Parameters("@Ano") = VGParamSistem.AnoProceso
        .Parameters("@Mes") = Format(IIf(Month(frmMantprovision.DTPFechaContab) < Val(VGParamSistem.MesProceso), Month(frmMantprovision.DTPFechaContab), CInt(VGParamSistem.MesProceso)), "00")
        .Parameters("@numero") = numero
        .Execute
    End With
VGCNx.CommitTrans
Exit Sub
Actualizacorre:
    VGvarVerifica = False
    VGErrorString = "Error en Actualizar el Numero de Comprobante Auxiliar " & Chr(13) & Err.Description
End Sub

Public Sub ActualizaCorrelAuxiliar(ByRef empresa As String, ByRef libro As String, ByRef anno As String, ByVal mes As Integer, ByVal numero As Double)
VGvarVerifica = True
On Error GoTo Actualizacorre
    Set VGCommandoSP = New ADODB.Command
    VGCommandoSP.ActiveConnection = VGGeneral
    VGCommandoSP.CommandType = adCmdStoredProc
    VGCommandoSP.CommandText = "co_actcorraux_pro"
    VGCommandoSP.Parameters.Refresh
    With VGCommandoSP
        .Parameters("@Base") = VGParamSistem.BDEmpresa
        .Parameters("@empresa") = empresa
        .Parameters("@libro") = libro
        .Parameters("@Ano") = anno
        .Parameters("@Mes") = Format(mes, "00")
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
VGvarVerifica = True
VGCNx.BeginTrans
Set VGvardllgen = New dllgeneral.dll_general
    Set rsaux = rs.Clone(adLockReadOnly)
    rsaux.Filter = "(impbruto<>0 or impcompra <> 0)"
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
        If VGParametros.sistemamultiempresas = True Then
              .Parameters("@empresa") = Trim(frmMantprovision.Ctr_Ayuempresa.xclave)
        End If
        .Parameters("@cabprovinumero") = numero
        If Month(frmMantprovision.DTPFechaContab) <> Val(VGParamSistem.MesProceso) Then
           .Parameters("@cabproviano") = Year(frmMantprovision.DTPFechaContab)
           .Parameters("@cabprovimes") = Month(frmMantprovision.DTPFechaContab)
         Else
           .Parameters("@cabproviano") = VGParamSistem.AnoProceso
           .Parameters("@cabprovimes") = CInt(VGParamSistem.MesProceso)
        End If
       If VGParametros.sistemamultiempresas = True Then
          .Parameters("@empresa") = Trim(frmMantprovision.Ctr_Ayuempresa.xclave)
        Else
          .Parameters("@empresa") = VGParametros.empresacodigo
       End If
        .Parameters("@op") = 2
        .Execute
    End With
    rsaux.MoveFirst
    While Not rsaux.EOF
        With VGCommandoSP
            .Parameters("@cabprovinumero") = numero
            If Month(frmMantprovision.DTPFechaContab) <> Val(VGParamSistem.MesProceso) Then
               .Parameters("@cabprovimes") = Month(frmMantprovision.DTPFechaContab)
             Else
               .Parameters("@cabprovimes") = CInt(VGParamSistem.MesProceso)
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
            .Parameters("@detproviimpigv") = rsaux!igv
            .Parameters("@detproviimpina") = rsaux!inafecto
            .Parameters("@detprovitotal") = rsaux!Impcompra
            '.Parameters("@detprovidscto") = 0
            '.Parameters("@detproviimpdct") = 0
            '.Parameters("@detproviimpven") = 0
            '.Parameters("@detproviigv") = 0
            .Parameters("@detproviformcamb") = IIf(frmMantprovision.lb_vcambio.Visible, Format(frmMantprovision.CmbTcambio.ListIndex + 1, "00"), "00")
            .Parameters("@detprovitipcam") = IIf(frmMantprovision.lb_vcambio.Visible, CDbl(VGvardllgen.ESNULO(frmMantprovision.lb_vcambio.Caption, 0)), 0)
            .Parameters("@usuariocodigo") = VGUsuario
            .Parameters("@fechaact") = Now
            .Parameters("@detproviglosa") = rsaux!glosa
            .Parameters("@detproviccosto") = Trim(VGvardllgen.ESNULO(rsaux!Ccosto, "00"))
            .Parameters("@analitico") = Trim(VGvardllgen.ESNULO(rsaux!analitico, "00"))
            .Execute
        End With
        rsaux.MoveNext
    Wend
VGCNx.CommitTrans
    Exit Sub
ErrorGrabaDetalle:
    VGvarVerifica = False
    VGErrorString = "Error en Grabar Detalle " & Chr(13) & Err.Description
End Sub
Public Sub GrabarCP_Cargo(op As Integer, Optional numero As Long = 0, Optional codigo As String, Optional monto As Double, Optional tipodoc As String)
On Error GoTo ErrorGrabaCP
Dim rsaux As ADODB.Recordset
Dim numerocomprobante As String
VGvarVerifica = True
VGCNx.BeginTrans

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
            numero = Trim(frmMantprovision.lbNumComprobCab)
        End If
        .Parameters("@abonotipoplanilla") = VGParametros.CpTiplan
        .Parameters("@abononumplanilla") = Format(numero, "000000")
        .Parameters("@numero") = Format(Trim(frmMantprovision.TxSerie.Text), "0000") & Format(Left(Trim(frmMantprovision.TxNdoc.Text), 10), "0000000000")
        'Este para que al eliminar no utilizar estos parametros
        If op <= 2 Then
              .Parameters("@cliente") = Trim(frmMantprovision.CtrAyu_Proveedor.xclave)
              .Parameters("@apeimppag") = CDbl(frmMantprovision.TxTotImpCompra.valor)
              .Parameters("@tipodocu") = Trim(frmMantprovision.CtrAyu_TipDoc.xclave)
            If op = 2 Then
               .Parameters("@oldnumero") = Format(Right(Trim(frmMantprovision.TxNdoc.Tag), 14), "00000000000000")
                             
               .Parameters("@oldtipodocu") = frmMantprovision.CtrAyu_TipDoc.Tag
               .Parameters("@oldcliente") = frmMantprovision.CtrAyu_Proveedor.Tag
            End If
            If codigo <> "" Then
              .Parameters("@cliente") = codigo
              .Parameters("@apeimppag") = CDbl(monto)
              .Parameters("@tipodocu") = tipodoc
            End If
            .Parameters("@vendedor") = VGParametros.CpOficina
            .Parameters("@zona") = Null
            .Parameters("@apefecemi") = frmMantprovision.Dtp_FechaDoc.Value
            .Parameters("@moneda") = frmMantprovision.CtrAyu_Moneda.xclave
            .Parameters("@usuario") = VGUsuario
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
            .Parameters("@cargoemiteretencion") = frmMantprovision.Emiteretencion
            .Parameters("@cargoemitedetraccion") = frmMantprovision.emitedetraccion
           If VGParametros.sistemamultiempresas = True Then
             .Parameters("@empresacodigo") = Trim(frmMantprovision.Ctr_Ayuempresa.xclave)
           Else
             .Parameters("@empresacodigo") = VGParametros.empresacodigo
          End If
        End If
        .Execute
    End With
VGCNx.CommitTrans
    Exit Sub
ErrorGrabaCP:
    VGvarVerifica = False
    VGErrorString = "Error en Grabar en Cuentas por Pagar " & Chr(13) & Err.Description
End Sub
Public Sub Grabaren_Tesoreria(ByVal op As Integer, Optional ByVal Numeroprovi As Long = 0, Optional ByVal rs As Recordset, Optional ByRef Xnum As String)
On Error GoTo ErrorGrabaTesore
Dim numero As String
VGvarVerifica = True
Set VGvardllgen = New dllgeneral.dll_general
Dim rb As ADODB.Recordset
Dim Item As Integer
   'Obtener el Ultimo Numero Correlativo de las cajas
    Dim opaux As Integer
    opaux = op
    Set VGCommandoSP = New ADODB.Command
    Set VGvardllgen = New dllgeneral.dll_general
    VGCommandoSP.ActiveConnection = VGGeneral
    VGCommandoSP.CommandType = adCmdStoredProc
    VGCommandoSP.CommandText = "te_abonadocumento_pro"
    VGCommandoSP.Parameters.Refresh
    If op = 2 Then
        If Trim(frmMantprovision.TxTesor.Text) = "" Then
            op = 1
        End If
    End If
    With VGCommandoSP
        .Parameters("@base") = VGParamSistem.BDEmpresa
        .Parameters("@tipo") = IIf(op = 3, 2, op)
        If op = 2 Or op = 3 Then
            numero = frmMantprovision.TxTesor.Text
            
          Else
            Set rb = VGCNx.Execute("select * from te_parametroempresa where empresacodigo='01'")
            VGCNx.BeginTrans
            If rb.RecordCount > 0 Then
                numero = Format(CDbl(VGvardllgen.ESNULO(rb!empresanumegreso, "0")) + 1, "000000")
                VGCNx.Execute "Update te_parametroempresa Set empresanumegreso='" & numero & "' where empresacodigo='01'"
            End If
            VGCNx.CommitTrans
        End If
        Xnum = numero
        .Parameters("@estadoreg") = ""
        .Parameters("@numrecibo") = numero
        If op = 3 Then
            .Execute
        End If
        If op = 2 Or frmMantprovision.modoproviold = 1 Then
            .Execute
           If frmMantprovision.ChkActCaja = 1 Then
              .Parameters("@tipo") = 1
              op = 1
            Else
              Xnum = ""
           End If
        End If
        
        'Este para que al eliminar no utilizar estos parametros
        If op = 1 Then
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
            .Parameters("@numreciboegreso") = frmMantprovision.Ctr_Ayutransf.xclave
            .Parameters("@usuario") = VGUsuario
            .Parameters("@fechaact") = Now
            If VGParametros.sistemamultiempresas = True Then
              .Parameters("@empresa") = Trim(frmMantprovision.Ctr_Ayuempresa.xclave)
            Else
              .Parameters("@empresa") = VGParametros.empresacodigo
            End If
            .Parameters("@cabprovinumero") = Numeroprovi
            .Parameters("@saldodocxrendir") = IIf(frmMantprovision.CtrAyu_Moneda.xclave = "01", CDbl(frmMantprovision.TxTotImpCompra.valor), Round(CDbl(frmMantprovision.TxTotImpCompra.valor) * CDbl(frmMantprovision.lb_vcambio.Caption), 2))
            .Parameters("@NumeroDocXRendir") = frmMantprovision.Ctr_Ayutransf.xclave
             If frmMantprovision.Ctr_Ayutransf.Visible = True Then .Parameters("@responsablectasxrendir") = frmMantprovision.clientecodigo
            .Execute
        End If
    End With
op = opaux
VGCNx.BeginTrans
Set VGCommandoSP = New ADODB.Command
VGCommandoSP.ActiveConnection = VGGeneral
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
  While Not rs.EOF
      With VGCommandoSP
        .Parameters("@base") = VGParamSistem.BDEmpresa
        .Parameters("@tipo") = IIf(op = 3, 2, op)
        .Parameters("@numrecibo") = numero
        If op = 3 Then
            .Execute
        End If
        If op = 2 Or frmMantprovision.modoproviold = 1 Then
            If Item = 1 Then .Execute
            If frmMantprovision.ChkActCaja = 1 Then
               .Parameters("@tipo") = 1
               op = 1
            End If
        End If
        If op = 1 Then
            .Parameters("@estadoreg") = ""
            .Parameters("@item") = Item
            .Parameters("@emisioncheque") = "C"
            .Parameters("@tipodocconcepto") = frmMantprovision.CtrAyu_TipDoc.xclave
            .Parameters("@numdocumento") = Format(Trim(frmMantprovision.TxSerie.Text), "0000") & Format(Left(Trim(frmMantprovision.TxNdoc.Text), 10), "0000000000")
            .Parameters("@carabo") = frmMantprovision.VlDocNota
            .Parameters("@formacan") = ""
            .Parameters("@tdqc") = ""
            .Parameters("@ndqc") = ""
            If frmMantprovision.Ctr_Ayutransf.Visible = True Then
               .Parameters("@tdqc") = "20"
               .Parameters("@ndqc") = frmMantprovision.Ctr_Ayutransf.xclave
            End If
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
            .Parameters("@usuario") = VGUsuario
            .Parameters("@fechaact") = Now
            .Parameters("@entidad") = rs.Fields("analitico")
            .Parameters("@centrocosto") = rs.Fields("ccosto")
            .Parameters("@cliente") = frmMantprovision.CtrAyu_Proveedor.xclave
            .Execute
         End If
    End With
    Item = Item + 1
    rs.MoveNext
  Wend
End If
VGCNx.CommitTrans
Exit Sub

ErrorGrabaTesore:
    VGvarVerifica = False
    VGErrorString = "Error en Grabar en Tesoreria " & Chr(13) & Err.Description
    Exit Sub
    Resume
End Sub
Public Sub MostrarCabecera(ByVal rs As ADODB.Fields)
Dim rsaux As ADODB.Fields, pos As Integer
Dim sqltes As ADODB.Recordset
    Set rsaux = rs
    Set VGvardllgen = New dll_general
    With frmMantprovision
        .estadorendicion = 0
        If rsaux!cabprovinumtesor <> "" Then
           Set sqltes = VGCNx.Execute("select top 1 * from te_detallerecibos where cabrec_numrecibo='" & rsaux!cabprovinumtesor & "' and isnull(detalle_no_saldos,0)<>1")
           If sqltes.RecordCount > 0 Then
             If ESNULO(sqltes!chkconcil, 0) = 1 Then
                .totalcomprobante = rsaux!cabprovitotal
                .estadorendicion = IIf(sqltes!chkconcil, 1, 0)
                If .estadorendicion = 1 Then
                   .estadorendicion = IIf(ESNULO(sqltes!rendicionnumero, 0) = 0, 0, 1)
                End If
                .numerorendicion = ESNULO(sqltes!rendicionnumero, 0)
                .fecharendicion = sqltes!fechconcil
             End If
           End If
           .numerorecibo = rsaux!cabprovinumtesor
        End If
        .lbNumComprobCab.Caption = rsaux!cabprovinumero
        .CtrAyu_Proveedor.xclave = rsaux!proveedorcodigo: .CtrAyu_Proveedor.Ejecutar
        .CtrAyu_Moneda.xclave = rsaux!monedacodigo: .CtrAyu_Moneda.Ejecutar
        .CtrAyu_Modoprovi.xclave = rsaux!modoprovicod: .CtrAyu_Modoprovi.Ejecutar
        .CtrAyu_TipDoc.xclave = rsaux!documetocodigo: .CtrAyu_TipDoc.Ejecutar
        .ChkCtaCte.Value = IIf(rsaux!cabprovictacte = 0, 0, 1)
        .ChkRegComp.Value = IIf(rsaux!cabproviregcom = 0, 0, 1)
'        .ChkActCaja.Value = IIf(rsaux!cabprotesor = 0, 0, 1)
        .modoproviold = .ChkActCaja.Value
        .ChkOperGrab.Value = IIf(rsaux!cabproviopergrab = 0, 0, 1)
        .Ctr_Ayutransf.xclave = ESNULO(rsaux!NumeroDocXRendir, "")
        'La Serie y El Numero de Documento
         pos = InStr(1, rsaux!CABPROVINUMDOC, "-", vbTextCompare)
         If pos > 0 Then
             .TxSerie.Text = Trim(Left(rsaux!CABPROVINUMDOC, IIf(pos = 0, 0, pos - 1)))
             .TxNdoc.Text = Right(rsaux!CABPROVINUMDOC, Len(rsaux!CABPROVINUMDOC) - pos)
           Else
             .TxSerie.Text = Left(rsaux!CABPROVINUMDOC, .TxSerie.MaxLength)
             .TxNdoc.Text = Right(rsaux!CABPROVINUMDOC, Len(rsaux!CABPROVINUMDOC) - .TxSerie.MaxLength)
             
         End If
        'Leer valores anteriores de Tipo Doc, Nº Doc y Cod. Cliente
        .CtrAyu_TipDoc.Tag = Trim(rsaux!documetocodigo)
        .TxSerie.Tag = .TxSerie.Text
        .TxNdoc.Tag = .TxNdoc.Text
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
        .TxNAux.Text = Right(Trim(rsaux!cabprovinumaux), 5)
        .lb_vcambio.Caption = Format(rsaux!cabprovitipcambio, "#0.000")
        
        .Ctr_AyudaOficina.xclave = VGvardllgen.ESNULO(rsaux!cabprovioficina, "00"): .Ctr_AyudaOficina.Ejecutar
        .Ctr_AyudaCaja.xclave = VGvardllgen.ESNULO(rsaux!cabprovicaja, "00"): .Ctr_AyudaCaja.Ejecutar
        
        .DTPFechaCaja.Value = IIf(.DTPFechaCaja.Visible, VGvardllgen.ESNULO(rsaux!cabprovifechcaja, Date), Date)
        .DTPFechaCaja.Tag = .DTPFechaCaja
        .DTPFechaContab.Value = rsaux!cabprovifchconta
        .TxTesor.Text = VGvardllgen.ESNULO(rsaux!cabprovinumtesor, "")
        .VlComprob_Conta = VGvardllgen.ESNULO(rsaux!cabprovinconta, "''''")
        .Ctr_Ayuempresa.xclave = VGvardllgen.ESNULO(rsaux!empresacodigo, "00"): .Ctr_Ayuempresa.Ejecutar
        If VGParametros.sistemamultiempresas Then
          .CtrAyu_Ccosto.filtro = "empresacodigo='" & .Ctr_Ayuempresa.xclave & "' and centrocostonivel='" & VGnumnivcos & "' and centrocostocodigo<>'00' "
         Else
          .CtrAyu_Ccosto.filtro = "empresacodigo='" & VGParametros.empresacodigo & "' and centrocostonivel='" & VGnumnivcos & "' and centrocostocodigo<>'00' "
        End If

    End With
End Sub
Public Sub MostrarDetalle(ByRef rs As ADODB.Recordset)
Dim rsaux As ADODB.Recordset, I As Integer
    Set rsaux = New ADODB.Recordset
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
        .Parameters("@empresa") = frmMantprovision.Ctr_Ayuempresa.xclave
        .Parameters("@op") = 3
        .Parameters("@cabprovinumero") = frmMantprovision.lbNumComprobCab
        .Parameters("@cabproviano") = VGParamSistem.AnoProceso
        Set rsaux = .Execute
    End With
    If rsaux.RecordCount() > 0 Then
       rsaux.MoveFirst
    End If
    While Not rsaux.EOF
        rs.AddNew
        For I = 0 To rsaux.Fields.Count - 1
            rs.Fields(rsaux.Fields(I).Name) = IIf(IsNull(rsaux.Fields(I).Value), "0", rsaux.Fields(I).Value)
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
Public Function ValidarGrabarCabecera(NR As Long) As Boolean
ValidarGrabarCabecera = False
Dim rsaux As New ADODB.Recordset
Dim montosoles As Double
Dim xrendicion As String
With frmMantprovision
    'Validando Que exista por lo menos un registro en el detalle
    Set VGvardllgen = New dllgeneral.dll_general
    If NR = 0 Then
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
        MsgBox "Tiene que seleccionar Tipo de provision ", vbInformation
        .CtrAyu_TipCompra.SetFocus
        Exit Function
    End If
    montosoles = VGvardllgen.ESNULO(.lb_vcambio.Caption, 0)
    If VGvardllgen.ESNULO(.lb_vcambio.Caption, 0) = 0 Then
        MsgBox "Debe de escoger una fecha que exista tipo de cambio", vbInformation
        Exit Function
    End If
    If Format(Month(.DTPFechaContab), "00") <> Format(VGParamSistem.MesProceso, "00") Then
       MsgBox "MES de Fecha Contable diferente a MES de fecha de proceso, VERIFIQUE", vbInformation
           .DTPFechaContab.SetFocus
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
    If .totalcomprobante <> .TxTotImpCompra.Text And .estadorendicion = 1 Then
            MsgBox "Monto Modificado es diferente al Monto Original -->  " & .totalcomprobante & "  <--de la Rendicion Nro. " & .numerorendicion & " ,Verificar", vbInformation
            .CtrAyu_Modoprovi.SetFocus
            Exit Function
        End If
    
    If Trim(.CtrAyu_Moneda.xclave) = "02" Then
       montosoles = IIf(.TxTotImpCompra.valor = "", 0, .TxTotImpCompra.valor * montosoles)
     Else
       montosoles = IIf(.TxTotImpCompra.valor = "", 0, .TxTotImpCompra.valor)
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
    DocAct = Trim(frmMantprovision.CtrAyu_Proveedor.xclave) & Trim(frmMantprovision.CtrAyu_TipDoc.xclave)
    DocAct = DocAct & Format(Trim(frmMantprovision.TxSerie.Text), "0000")
    DocAct = DocAct & Format(Trim(frmMantprovision.TxNdoc.Text), "0000000000")
    If (frmMantprovision.IMant = 1) Or (frmMantprovision.VlDocAnt <> DocAct) Then
        Set rsaux = VGCNx.Execute(" Select * From " & VGParamSistem.TablaCabcomprob & _
                     " Where " & _
                     " isnull(proveedorcodigo,'')+isnull(documetocodigo,'')+cabprovinumdoc='" & DocAct & "'")
        If rsaux.RecordCount > 0 Then
           MsgBox "Documento ya ingresado Nro provision : " & rsaux!cabprovinumero & "  Mes : " & rsaux!cabprovimes, vbExclamation
           frmMantprovision.TxNdoc.SetFocus
           Exit Function
        End If
    End If
    If .ChkActCaja = 1 Then
       If Trim(.Ctr_AyudaCaja.xclave) = "" Then
          MsgBox "No se ha ingresado codigo de caja , ingrese por favor ", vbExclamation
          .Ctr_AyudaCaja.SetFocus
          Exit Function
        ElseIf Trim(.CtrAyu_Moneda.xclave) <> VGParametros.monedabase And Trim(VGParametros.monedabase) <> "" Then
               MsgBox "Documento en Moneda " & Trim(.CtrAyu_Moneda.xnombre) & " y por Caja Chica, ingrese por favor Modo de Compra PROVEEDORES ", vbExclamation
               .Ctr_AyudaCaja.SetFocus
               Exit Function
        ElseIf .Dtp_FechaDoc.Value > .DTPFechaCaja.Value Then
               MsgBox "Fecha de Caja es MENOR a Fecha del documento " & " Corrija por favor ", vbExclamation
               .DTPFechaCaja.SetFocus
                Exit Function
         ElseIf Year(.DTPFechaContab) <> Year(.DTPFechaCaja) Then
                MsgBox " Fecha de Caja es Diferente a de Provision( FECHA CONTABLE ), Corrija por favor ", vbExclamation
                .DTPFechaCaja.SetFocus
                Exit Function
         ElseIf Month(.DTPFechaContab) <> Month(.DTPFechaCaja) Then
                MsgBox "MES de Fecha de Caja es Diferente a MES de Provision( FECHA CONTABLE ), ingrese por Tipo Modo de Cta Cte  " & " Corriga por favor ", vbExclamation
                .DTPFechaCaja.SetFocus
                Exit Function
       End If
    End If
    If VGParametros.sistemamultiempresas = True And Trim(.Ctr_Ayuempresa.xclave) = "" Then
       MsgBox "Tiene que Seleccionar un codigo de empresa", vbInformation
       .Ctr_Ayuempresa.SetFocus
            Exit Function
        End If
    If VGParametros.sistemabancarizacion = True And .ChkActCaja = 1 Then
       If .CtrAyu_Moneda.xclave = "01" And .TxTotImpCompra.valor > VGParametros.sistemabancarizacion01 Then
           MsgBox "Tiene que Seleccionar Modo Proveedores por Bancarizacion de Soles Mayor a " & VGParametros.sistemabancarizacion01, vbInformation
          .CtrAyu_Modoprovi.SetFocus
            Exit Function
        End If
       If .CtrAyu_Moneda.xclave = "02" And .TxTotImpCompra.valor > VGParametros.sistemabancarizacion02 Then
           MsgBox "Tiene que Seleccionar Modo Proveedores por Bancarizacion de dolares  Mayor a " & VGParametros.sistemabancarizacion02, vbInformation
          .CtrAyu_Modoprovi.SetFocus
            Exit Function
        End If
    End If
    If .Ctr_Ayutransf.Visible = True And .TxImpCompra.Text > .saldodocxrendir Then
       MsgBox "Monto del documento ingresado excede al saldo de doc. a rendir, que es  -- > " & .saldodocxrendir & " , verifique ", vbInformation
      .Ctr_Ayutransf.SetFocus
  '    Exit Function
    End If
    If VGParametros.controlaestadosrendicion Then
       If .controlarendicion Then
          SQL = "select numero=max(rendicionnumero) from te_rendiciones where oficinacodigo='" & .Ctr_AyudaOficina.xclave & "'"
          SQL = SQL & " and codigocaja='" & .Ctr_AyudaCaja.xclave & "'"
          Set rsaux = VGCNx.Execute(SQL)
          xrendicion = numero(rsaux!numero)
          If xrendicion > 0 Then
             SQL = " select rendicionfecha,rendicionsaldofinal=rendicionsaldoinicial+rendicioningresos-rendicionegresos + isnull(saldoacumuladoxrendir,0)"
             SQL = SQL & " from te_rendiciones where oficinacodigo='" & .Ctr_AyudaOficina.xclave & "'"
             SQL = SQL & " and codigocaja='" & .Ctr_AyudaCaja.xclave & "' and rendicionnumero='" & xrendicion & "'"
             Set rsaux = VGCNx.Execute(SQL)
             If .IMant = 1 Then
                .fecharendicion = rsaux!rendicionfecha - VGParametros.diasatrazorendicion
                If .DTPFechaCaja.Value < .fecharendicion Then
                    MsgBox "fecha de Pago -- > " & .DTPFechaCaja & " ,menor a fecha Permitida por la Ultima rendicion -->  " & .fecharendicion, vbInformation
                   .DTPFechaCaja.SetFocus
                   Exit Function
                End If
                If numero(rsaux!rendicionsaldofinal) - .TxTotImpCompra.Text < 0 Then
                   MsgBox " Monto de cancelacion Origina Saldos Negativos, Monto permitido es de  -- > " & numero(rsaux!rendicionsaldofinal), vbInformation
                   .Ctr_AyudaCaja.SetFocus
                   Exit Function
                End If
             End If
           Else
             If .DTPFechaCaja.Value <> .DTPFechaCaja.Tag And .IMant <> 1 Then
                 MsgBox " Modificacion de Fecha de Caja no esta permitido ", vbInformation
                .DTPFechaCaja.SetFocus
                Exit Function
             End If
          End If
       End If
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
    If .CheckActivo.Value = 0 Then
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
    End If
    If (.CtrAyu_Ccosto.Enabled And .CtrAyu_Ccosto.Visible) And (Trim(.CtrAyu_Ccosto.xclave) = "" Or Trim(.CtrAyu_Ccosto.xclave) = "00") Then
        MsgBox "Debe ingresar El cento de costo", vbInformation
        .CtrAyu_Ccosto.SetFocus
        Exit Function
    End If
    If (.Ctr_AyuAnalitico.Enabled And .Ctr_AyuAnalitico.Visible) And (Trim(.Ctr_AyuAnalitico.xclave) = "" Or Trim(.Ctr_AyuAnalitico.xclave) = "00") Then
        MsgBox "Debe ingresar Codigo de analitico", vbInformation
        .Ctr_AyuAnalitico.SetFocus
        Exit Function
    End If
    If .tipodetraccion = 1 Then .Emiteretencion = 2
    If .tipodetraccion = 1 Then .emitedetraccion = 1
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
        .txRuc.Text = "":
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
        If VGParametros.sistemamultiempresas = True Then
          .Ctr_Ayuempresa.xclave = ""
          .Ctr_Ayuempresa.xnombre = ""
        End If
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

Public Sub Actualizasaldorendicion()
Dim rsaux As New ADODB.Recordset
Dim xrendicion As String
Dim xsaldo As Double
On Error GoTo actualizasaldo
VGvarVerifica = True
VGCNx.BeginTrans
With frmMantprovision
    SQL = "select numero=max(rendicionnumero) from te_rendiciones where oficinacodigo='" & .Ctr_AyudaOficina.xclave & "'"
    SQL = SQL & " and codigocaja='" & .Ctr_AyudaCaja.xclave & "'"
    Set rsaux = VGCNx.Execute(SQL)
    xrendicion = numero(rsaux!numero)
    If .CtrAyu_Moneda.xclave = "02" Then
       xsaldo = .TxTotImpCompra.Text * .lb_vcambio
     Else
       xsaldo = .TxTotImpCompra.Text
    End If
    SQL = " update te_rendiciones set saldoacumuladoxrendir=isnull(saldoacumuladoxrendir,0)-" & xsaldo
    SQL = SQL & " where oficinacodigo='" & .Ctr_AyudaOficina.xclave & "'"
    SQL = SQL & " and codigocaja='" & .Ctr_AyudaCaja.xclave & "' and rendicionnumero='" & xrendicion & "'"
    Set rsaux = VGCNx.Execute(SQL)
End With
VGCNx.CommitTrans
Exit Sub
actualizasaldo:
Screen.MousePointer = 1
VGvarVerifica = False
VGErrorString = "Error en Grabar saldo de rendiciones  " & Chr(13) & Err.Description
End Sub

Public Sub GrabaAsientoAuto(ByVal comp As String, ByVal numconta As String)
Dim rsql As New ADODB.Recordset
SQL = "select cabcomprobnumero,asientocodigo,subasientocodigo,cabcomprobmes from ct_cabcomprob" & VGParamSistem.AnoProceso
SQL = SQL & " where cabcomprobnprovi='" & comp & "'"
Set rsql = VGCNx.Execute(SQL)
If rsql.RecordCount > 0 Then
    On Error GoTo GrabaAuto
    Set VGCommandoSP = New ADODB.Command
    VGCommandoSP.ActiveConnection = VGGeneral
    VGCommandoSP.CommandType = adCmdStoredProc
    VGCommandoSP.CommandText = "ct_grabaautomatico_pro"
    VGCommandoSP.Parameters.Refresh
    With VGCommandoSP
        .Parameters("@base") = VGParamSistem.BDEmpresaCT
        .Parameters("@tabla") = "ct_detcomprob" & VGParamSistem.AnoProceso
        .Parameters("@comp") = comp
        .Parameters("@mes") = CInt(VGParamSistem.MesProceso)
        .Parameters("@asiento") = rsql!asientocodigo
        .Parameters("@subasiento") = rsql!subasientocodigo
        .Execute
    End With
End If
    Exit Sub
GrabaAuto:
    VGvarVerifica = False
    VGErrorString = "Error en Grabar Asientos Automaticos " & Chr(13) & Err.Description
End Sub



