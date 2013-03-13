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
    operacioncodigo = 1
    cuentacodigo = 2
    centrocostocodigo = 3
    tipoanaliticocodigo = 4
    analiticocodigo = 5
    detcomprobruc = 6
    documentocodigo = 7
    detcomprobnumdocumento = 8
    detcomprobfechavencimiento = 9
    detcomprobglosa = 10
    indicador = 11
    monedacodigo = 12
    Montos = 13
    detcomprobfecharef = 14
    detcomprobfechaemision = 15
    plantillaasientoinafecto = 16
End Enum

Public Sub CargarAyudas()
With frmantcomprobantes
    Call .CtrAyu_Analitico.conexion(VGCNx): .CtrAyu_Analitico.Filtro = "analiticocodigo<>'00'"
    Call .CtrAyu_Asiento.conexion(VGCNx): .CtrAyu_Asiento.Filtro = "asientocodigo <>'00'"
    Call .CtrAyu_CCosto.conexion(VGCNx): .CtrAyu_CCosto.Filtro = "empresacodigo='" & VGParametros.empresacodigo & "' and centrocostocodigo<>'00'"
    Call .CtrAyu_Cuenta.conexion(VGCNx): .CtrAyu_Cuenta.Filtro = "empresacodigo='" & VGParametros.empresacodigo & "' and cuentanivel=" & VGnumnivelescuenta & " and cuentacodigo <>'00'"
    Call .CtrAyu_Opera.conexion(VGCNx)
    Call .CtrAyu_SubAsiento.conexion(VGCNx): .CtrAyu_SubAsiento.Filtro = "subasientocodigo <>'00'"
    Call .CtrAyu_TipAnal.conexion(VGCNx): .CtrAyu_TipAnal.Filtro = "tipoanaliticocodigo<>'00'"
    Call .CtrAyu_TipDoc.conexion(VGCNx): .CtrAyu_TipDoc.Filtro = "documentocodigo<>'00'"
    Call .CtrAyu_TipRef.conexion(VGCNx): .CtrAyu_TipRef.Filtro = "documentocodigo<>'00' and isnull(documentoref,0)<>0 "
    Call .CtrAyu_Moneda.conexion(VGCNx): .CtrAyu_Moneda.Filtro = "monedacodigo<>'00'"
End With
End Sub
Public Sub CreaRsTempDetalle(ByRef rs As ADODB.Recordset)
    With rs.Fields
        Call .Append("detcomprobitem", adVarChar, 5)
        Call .Append("operacioncodigo", adVarChar, 2)
        Call .Append("cuentacodigo", adVarChar, 20)
        Call .Append("centrocostocodigo", adVarChar, 6)
        Call .Append("tipoanaliticocodigo", adVarChar, 3)
        Call .Append("analiticocodigo", adVarChar, 15)
        Call .Append("detcomprobruc", adVarChar, 13, adFldIsNullable)
        Call .Append("documentocodigo", adVarChar, 2)
        Call .Append("detcomprobnumdocumento", adVarChar, 20, adFldIsNullable)
        Call .Append("detcomprobfechaemision", adDate)
        Call .Append("detcomprobfechavencimiento", adDate, , adFldIsNullable)
        Call .Append("detcomprobglosa", adVarChar, 50)
        Call .Append("indicador", adVarChar, 1)
        Call .Append("monedacodigo", adVarChar, 2)
        Call .Append("tcambio", adVarChar, 3)
        Call .Append("valcambio", adDouble)
        Call .Append("montosol", adDouble)
        Call .Append("montouss", adDouble)
        Call .Append("detcomprobauto", adBoolean)
        Call .Append("detcomprobajusteuser", adBoolean)
        Call .Append("plantillaasientoinafecto", adBoolean)
        Call .Append("tipdocref", adVarChar, 2, adFldIsNullable)
        Call .Append("detcomprobnumref", adVarChar, 20, adFldIsNullable)
        Call .Append("comodin", adVarChar, 250, adFldIsNullable)
        Call .Append("tdaux", adVarChar, 2, adFldIsNullable)
        Call .Append("ndaux", adVarChar, 50, adFldIsNullable)
        Call .Append("index", adInteger)
        Call .Append("detcomprobitemaux", adVarChar, 5)
        Call .Append("NumPlantilla", adInteger)
        Call .Append("detcomprobitemres", adVarChar, 5)
        Call .Append("detcomprobnlibro", adVarChar, 10, adFldIsNullable)
        Call .Append("detcomprobfecharef", adDate, , adFldIsNullable)
    End With
    
End Sub
Public Sub AñadiralDetalle(ByRef rs As ADODB.Recordset)
Dim vardllgen As New dllgeneral.dll_general
'FIXIT: Declare 'fila' con un tipo de datos de enlace en tiempo de compilación             FixIT90210ae-R1672-R1B8ZE
Dim fila As Variant, CamposAux As ADODB.Fields, i As Integer
'FIXIT: Declare 'Posant' con un tipo de datos de enlace en tiempo de compilación           FixIT90210ae-R1672-R1B8ZE
Dim Rsaux As ADODB.Recordset, Posant As Variant
    Set Rsaux = rs.Clone(adLockReadOnly)
    Rsaux.Filter = 0
    Call Limpia
    fila = 1
    If rs.RecordCount > 1 Then
        'fila = CDbl(left(rs!Index, Len(rs!Index) - 1)) & CaracterProx(rs!Index)
        fila = CInt(rs!detcomprobitem) + 1
        Rsaux.Filter = "detcomprobitem='" & rs!detcomprobitem & "'"
        Set CamposAux = Rsaux.Fields
    End If
    With rs
        .AddNew
        If fila = 1 Then
            !detcomprobitem = Format(Val(vardllgen.ESNULO(MaxItem(rs, "detcomprobitem"), 0) + 1), "00000")
            !detcomprobfechaemision = Now
            !indicador = "D"
            !monedacodigo = VGMonSubAsiento
            !tcambio = "02"
            !valcambio = VGValorCambio
            !Index = 2
            !detcomprobauto = 0
           Else
            For i = 0 To CamposAux.Count - 1
                rs.Fields(i).Value = CamposAux(i).Value
            Next
           !detcomprobitem = Format(fila, "00000")
           !Index = 1
           !NumPlantilla = 0
         End If
        .Update
    End With
    If fila <> 1 Then
        Call ActualizaNumItems(rs, fila)
        Call VerfiSiEsPlantilla(rs)
        Call CalculoIGV(rs)
        Call CalculodeAjuste(rs)
        Call frmantcomprobantes.CalcularTotales(rs)
    End If
End Sub
Private Function CaracterProx(ByVal cad As String) As String
Dim pos As Integer, cadeval As String, cadaux As String
    pos = InStr(1, cad, "A", vbTextCompare)
    cadeval = Mid$(cad, pos, (Len(cad) - pos) + 1)
    cadaux = Right(cadeval, 1)
    If Asc(cadaux) = 90 Then
        cadeval = cadeval & "A"
      Else
        cadeval = Left(cadeval, Len(cadeval) - 1) & Chr(Asc(cadaux) + 1)
    End If
    CaracterProx = cadeval
End Function
Public Sub VerfiSiEsPlantilla(rs As ADODB.Recordset)
    Dim Rsaux As ADODB.Recordset
    Set Rsaux = New ADODB.Recordset
    Rsaux.Open "select * from ct_subasiento where asientocodigo='" & Trim$(frmantcomprobantes.VPAsiento) & _
               "' and subasientocodigo='" & Trim$(frmantcomprobantes.VPSubAsiento) & _
               "' and  isnull(subasientorepitedoc,0) =1 ", VGCNx, adOpenKeyset, adLockReadOnly
    
    If Rsaux.RecordCount = 0 Then Exit Sub
    If Rsaux.RecordCount <> rs.RecordCount Then
        VGRepiteDoc = False
      Else
       VGRepiteDoc = True
    End If
End Sub
'FIXIT: Declare 'criterio' con un tipo de datos de enlace en tiempo de compilación         FixIT90210ae-R1672-R1B8ZE
Public Sub ActualizaNumItems(ByRef rs As ADODB.Recordset, ByVal criterio As Variant)
'FIXIT: Declare 'Posant' con un tipo de datos de enlace en tiempo de compilación           FixIT90210ae-R1672-R1B8ZE
Dim Posant As Variant, i As Long
    VGactulizodoc = True
    rs.MoveFirst
    i = 1
    Do While Not rs.EOF
       rs!detcomprobitemres = Format(i, "00000")
       rs.Update
       i = i + 1
       rs.MoveNext
    Loop
    rs.Sort = ""
    rs.MoveFirst
    Do While Not rs.EOF
       rs!detcomprobitem = rs!detcomprobitemres
       rs!Index = 2
       rs.Update
       rs.MoveNext
    Loop
    rs.Sort = "detcomprobitem asc,index asc "
    rs.MoveFirst
    rs.Find "detcomprobitem='" & Format(criterio, "00000") & "'"
    VGactulizodoc = False
End Sub

Public Sub Limpia()
    VGflaglimpia = False
    With frmantcomprobantes
        .CtrAyu_Opera.xclave = "": .CtrAyu_Opera.xnombre = ""
        .CtrAyu_Cuenta.xclave = "": .CtrAyu_Cuenta.xnombre = ""
        .CtrAyu_CCosto.xclave = "": .CtrAyu_CCosto.xnombre = ""
        .CtrAyu_TipAnal.xclave = "": .CtrAyu_TipAnal.xnombre = ""
        .CtrAyu_Analitico.xclave = "": .CtrAyu_Analitico.xnombre = ""
        .txRuc.Text = ""
        .CtrAyu_TipDoc.xclave = "": .CtrAyu_TipDoc.xnombre = ""
        .TxSerie.Text = ""
        .TxNdoc.Text = ""
   '     .Dtp_FechaDoc = frmantcomprobantes.DTPFechaComprobCab
        .DtpFech_Ven = Null
        .TxGlosa.Text = ""
        .CmbID.ListIndex = 0
        .CtrAyu_Moneda.xclave = VGMonSubAsiento: .CtrAyu_Moneda.Ejecutar
        .CmbTcambio.ListIndex = 1
        VGValorCambio = RecuperaTipoCambio(Format(frmantcomprobantes.DTPFechaComprobCab, "dd/mm/yyyy"), Venta)
        .lb_vcambio.Caption = Format(VGValorCambio, "#.000 ")
        .TxMonto.Text = Format(0, "#.00")
        .CtrAyu_TipRef.xclave = "": .CtrAyu_TipRef.xnombre = ""
        .TxNref.Text = ""
        .DtpFech_Ven.Value = Null
        .Dtp_FechaDocRef.Value = Null
    End With
    VGflaglimpia = True
End Sub
'FIXIT: Declare 'FechaAux' con un tipo de datos de enlace en tiempo de compilación         FixIT90210ae-R1672-R1B8ZE
Public Sub ActualizarDetalle(ByRef rs As ADODB.Recordset, op As Campos, Optional flag As Boolean = False, Optional FechaAux As Variant = Null)
'Procedimiento Realizado por Fernando Cossio......
Dim vardllgen As New dllgeneral.dll_general
'FIXIT: Declare 'POSANTERIOR' con un tipo de datos de enlace en tiempo de compilación      FixIT90210ae-R1672-R1B8ZE
Dim POSANTERIOR As Variant
Dim flagajuste As Boolean
Dim swth As Integer
    If rs Is Nothing Then Exit Sub
    If rs.RecordCount = 0 Then Exit Sub
VGactulizodoc = False
    If VGMoverRegistro Then Exit Sub
    Screen.MousePointer = 11
    With rs
        '!detcomprobitem = Format(Val(vardllgen.ESNULO(MaxItem(rs, "detcomprobitem"), 0) + 1), "00000")
        Select Case op
            Case 0
                VGactulizodoc = True
                POSANTERIOR = .AbsolutePosition
                .MoveFirst
                 While Not .EOF
                    !detcomprobglosa = frmantcomprobantes.TxGlosaComprobCab.Text
                    .Update
                    .MoveNext
                 Wend
                    VGactulizodoc = False
                    .AbsolutePosition = POSANTERIOR
                    Screen.MousePointer = 1
                 Exit Sub
            Case 1: !operacioncodigo = frmantcomprobantes.CtrAyu_Opera.xclave
            Case 2: !cuentacodigo = frmantcomprobantes.CtrAyu_Cuenta.xclave
            Case 3: !centrocostocodigo = Trim$(frmantcomprobantes.CtrAyu_CCosto.xclave)
            Case 4: !tipoanaliticocodigo = frmantcomprobantes.CtrAyu_TipAnal.xclave
            Case 5
                !analiticocodigo = frmantcomprobantes.CtrAyu_Analitico.xclave
                '!detcomprobruc = Trim$(frmantcomprobantes.txRuc.Text)
            Case 6: !detcomprobruc = Trim$(frmantcomprobantes.txRuc.Text)
            Case 7, 8
                If VGRepiteDoc Then
                    VGactulizodoc = True
                    POSANTERIOR = .AbsolutePosition
                    .MoveFirst
                    While Not .EOF
                        !documentocodigo = frmantcomprobantes.CtrAyu_TipDoc.xclave
                        !detcomprobnumdocumento = Trim$(frmantcomprobantes.TxSerie.Text) & Trim$(frmantcomprobantes.TxNdoc.Text)
                        
                        !tipdocref = frmantcomprobantes.CtrAyu_TipRef.xclave
                        !detcomprobnumref = frmantcomprobantes.TxNref.Text
                        .Update
                        .MoveNext
                    Wend
                    VGactulizodoc = False
                    .AbsolutePosition = POSANTERIOR
                    Screen.MousePointer = 1
                    Exit Sub
                  Else
                    !tipdocref = frmantcomprobantes.CtrAyu_TipRef.xclave
                    !detcomprobnumref = frmantcomprobantes.TxNref.Text
                    !documentocodigo = frmantcomprobantes.CtrAyu_TipDoc.xclave
                    !detcomprobnumdocumento = Trim$(frmantcomprobantes.TxSerie.Text) & Trim$(frmantcomprobantes.TxNdoc.Text)
                End If
            Case 9: !detcomprobfechavencimiento = frmantcomprobantes.DtpFech_Ven
            Case 14: !detcomprobfecharef = frmantcomprobantes.Dtp_FechaDocRef
            Case 10: !detcomprobglosa = frmantcomprobantes.TxGlosa.Text
            Case 11: !indicador = Trim$(vardllgen.ComboDato(frmantcomprobantes.CmbID.Text))
            Case 12
                
                If (CDbl(!detcomprobitem) = 1 Or flag) And vgcont <> 2 Then
                    VGactulizodoc = False
                    Dim Fecha As Date, mon As String, camb As Double
                    
                    Fecha = frmantcomprobantes.Dtp_FechaDoc
                    mon = frmantcomprobantes.CtrAyu_Moneda.xclave
                    camb = CDbl(vardllgen.ESNULO(frmantcomprobantes.lb_vcambio, 0))
                    POSANTERIOR = .AbsolutePosition
                    '.AbsolutePosition = 1
                    .MoveFirst
                    Call frmantcomprobantes.AlMoverRegistro
                     While Not .EOF
                     swth = 1
                        !detcomprobfechaemision = Fecha
                        !valcambio = camb
                        !monedacodigo = mon
                        .Update
                        vgcont = 2
                        If Not flag Then
                            frmantcomprobantes.CtrAyu_Moneda.xclave = mon: frmantcomprobantes.CtrAyu_Moneda.Ejecutar
                          Else
     '                       frmantcomprobantes.Dtp_FechaDoc = frmantcomprobantes.DTPFechaComprobCab
                            If IsNull(FechaAux) Then FechaAux = frmantcomprobantes.Dtp_FechaDoc
                            VGValorCambio = RecuperaTipoCambio(Format(FechaAux, "dd/mm/yyyy"), frmantcomprobantes.CmbTcambio.ListIndex + 1)
                            frmantcomprobantes.lb_vcambio.Caption = Format(VGValorCambio, "#.000 ")
                            'If .AbsolutePosition = 1 Then .AbsolutePosition = 1
                            frmantcomprobantes.CtrAyu_Moneda.xclave = mon: frmantcomprobantes.CtrAyu_Moneda.Ejecutar
                            'Se ejecuta y Actuliza todo el monto de ese registro en ese momento
                        End If
                        
                        
MonedaTodos:
                        .MoveNext
                     Wend
                        VGactulizodoc = False
                        vgcont = 0
                        .AbsolutePosition = POSANTERIOR
                        
                        Screen.MousePointer = 1
                     Exit Sub
                  Else
                    swth = 2
                    GoTo ActualizaMontos
Monedasola:
                End If
                    
            Case 13
                 swth = 3
                GoTo ActualizaMontos
                    
Montos:
                
'                If cuentaigv(CInt(!detcomprobitem)) <> "" Then
'                    Call CalculoIGV(rs)
'                End If
            Case 15: !detcomprobfechaemision = frmantcomprobantes.Dtp_FechaDoc
            Case 16: !plantillaasientoinafecto = frmantcomprobantes.ChkInafecto.Value
                
        End Select
        !detcomprobauto = 0
        .Update
    End With
    Screen.MousePointer = 1
    Exit Sub
ActualizaMontos:
With rs
    !monedacodigo = frmantcomprobantes.CtrAyu_Moneda.xclave
    !tcambio = Trim$(vardllgen.ComboDato(frmantcomprobantes.CmbTcambio.Text))
    !detcomprobfechaemision = frmantcomprobantes.Dtp_FechaDoc
    !valcambio = CDbl(vardllgen.ESNULO(frmantcomprobantes.lb_vcambio, 0))
    !detcomprobajusteuser = frmantcomprobantes.ChkAjusta.Value
    If frmantcomprobantes.ChkAjusta.Value = 0 Then
        If !monedacodigo <> VGParametros.monedabase Then
            !montosol = Coversion(!monedacodigo, !valcambio, CDbl(vardllgen.ESNULO(Espunto(frmantcomprobantes.TxMonto.valor), 0)))
            !montouss = Round(CDbl(vardllgen.ESNULO(Espunto(frmantcomprobantes.TxMonto.valor), 0)), 2)
        Else
            !montouss = Coversion(!monedacodigo, !valcambio, CDbl(vardllgen.ESNULO(Espunto(frmantcomprobantes.TxMonto.valor), 0)))
            !montosol = Round(CDbl(vardllgen.ESNULO(Espunto(frmantcomprobantes.TxMonto.valor), 0)), 2)
        End If
      Else
        If !monedacodigo = VGParametros.monedabase Then
            !montosol = Round(CDbl(vardllgen.ESNULO(Espunto(frmantcomprobantes.TxMonto.valor), 0)), 2)
            !montouss = Round(CDbl(vardllgen.ESNULO(Espunto(frmantcomprobantes.TxValor.valor), 0)), 2)
        Else
            !montouss = Round(CDbl(vardllgen.ESNULO(Espunto(frmantcomprobantes.TxMonto.valor), 0)), 2)
            !montosol = Round(CDbl(vardllgen.ESNULO(Espunto(frmantcomprobantes.TxValor.valor), 0)), 2)
        End If
    End If
End With
Select Case swth
    Case 1: GoTo MonedaTodos
    Case 2: GoTo Monedasola
    Case 3: GoTo Montos
End Select
End Sub
Public Sub CalculoIGV(ByRef rs As ADODB.Recordset)
'FIXIT: Declare 'POSANTERIOR' con un tipo de datos de enlace en tiempo de compilación      FixIT90210ae-R1672-R1B8ZE
Dim POSANTERIOR As Variant
Dim Rsaux As ADODB.Recordset, monto As Double
Dim ctaigv As String, inaf As Boolean, porc As Double
    If cuentaigv(CInt(VGvardllgen.ESNULO(rs!detcomprobitemaux, "0"))) = "" Then Exit Sub
    'If rs!montosol = 0 Then Exit Sub
    VGactulizodoc = True
    Set VGvardllgen = New dllgeneral.dll_general
    Set Rsaux = New ADODB.Recordset
    Rsaux.Fields.Append "cuenta", adVarChar, 20
    Rsaux.Fields.Append "valorsol", adDouble
    Rsaux.Fields.Append "valordol", adDouble
    Rsaux.Open
With rs
    POSANTERIOR = .AbsolutePosition
    .MoveFirst
    While Not .EOF
        ctaigv = cuentaigv(CInt(VGvardllgen.ESNULO(!detcomprobitemaux, "0")), inaf, porc)
        If Trim$(ctaigv) <> "" And Not inaf Then
            Rsaux.Filter = "cuenta='" & ctaigv & "'"
            If Rsaux.RecordCount = 0 Then
                Rsaux.AddNew
                Rsaux!cuenta = ctaigv
            End If
            Rsaux!valorsol = Rsaux!valorsol + (!montosol * (porc / 100))
            Rsaux!valordol = Rsaux!valordol + (!montouss * (porc / 100))
            Rsaux.Update
        End If
        .MoveNext
    Wend
    'Actualizo recordset
    Rsaux.MoveFirst
    While Not Rsaux.EOF
        .Filter = "cuentacodigo='" & Rsaux!cuenta & "'"
        If .RecordCount > 0 Then
            !montosol = Rsaux!valorsol
            !montouss = Rsaux!valordol
            .Update
        End If
        Rsaux.MoveNext
    Wend
    .Filter = 0
    .AbsolutePosition = POSANTERIOR
    VGactulizodoc = False
End With
End Sub
Private Function cuentaigv(Item As Integer, Optional ByRef inaf As Boolean, Optional ByRef porc As Double) As String
Dim Rsaux As ADODB.Recordset
Dim SQL As String
    cuentaigv = ""
    inaf = False
    Set Rsaux = New ADODB.Recordset
    SQL = "select isnull(plantillaasientocuentaigv,'') as A,plantillaasientoinafecto as B,isnull(plantillaasientovalorigv,1) as C from ct_plantillaasiento where " & _
               "asientocodigo='" & frmantcomprobantes.VPAsiento & "' and subasientocodigo='" & frmantcomprobantes.VPSubAsiento & "' and " & _
               "plantillaasientocorrela=" & Item
    Rsaux.Open SQL, VGCNx, adOpenKeyset, adLockReadOnly
    If Rsaux.RecordCount = 0 Then Exit Function
    cuentaigv = Rsaux!a
    inaf = IIf(Rsaux!B = 0, False, True)
    porc = Rsaux!C
End Function
Private Function cuentaajuste(cuenta As String, Optional ByRef ind As String) As Integer
Dim Rsaux As ADODB.Recordset
    cuentaajuste = 0
    ind = ""
    Set Rsaux = New ADODB.Recordset
    Rsaux.Open "select isnull(plantillaasientoctaajuste,0) as A, isnull(iddebeohaber,'') as B from ct_plantillaasiento where " & _
               "asientocodigo='" & frmantcomprobantes.VPAsiento & "' and subasientocodigo='" & frmantcomprobantes.VPSubAsiento & "' and " & _
               "cuentacodigo='" & Trim$(cuenta) & "' and isnull(plantillaasientoctaajuste,0) <> 0 ", VGCNx, adOpenKeyset, adLockReadOnly
    If Rsaux.RecordCount = 0 Then Exit Function
    cuentaajuste = 1
    ind = Trim$(Rsaux!B)
End Function
'FIXIT: Declare 'MaxItem' con un tipo de datos de enlace en tiempo de compilación          FixIT90210ae-R1672-R1B8ZE
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

Public Sub GrabarCabecera(ByVal op As Integer, NumComprob As String, Optional ByVal Numlibro As String)
On Error GoTo ErrorGrabaCabecera
    Set VGCommandoSP = New ADODB.Command
    Set VGvardllgen = New dllgeneral.dll_general
    VGCommandoSP.ActiveConnection = VGGeneral
    VGCommandoSP.CommandType = adCmdStoredProc
    VGCommandoSP.CommandText = "ct_grabarcabcomprob_pro"
    VGCommandoSP.Parameters.Refresh
     With VGCommandoSP
        .Parameters("@base") = VGParamSistem.BDEmpresa
        .Parameters("@tabla") = VGParamSistem.TablaCabcomprob
        .Parameters("@empresa") = VGParametros.empresacodigo
        .Parameters("@op") = op
        .Parameters("@cabcomprobmes") = CInt(VGParamSistem.Mesproceso)
        .Parameters("@cabcomprobnumero") = Trim$(NumComprob)
        .Parameters("@asientocodigo") = frmantcomprobantes.VPAsiento
        .Parameters("@subasientocodigo") = frmantcomprobantes.VPSubAsiento
        If op <= 2 Then
            .Parameters("@cabcomprobfeccontable") = frmantcomprobantes.DTPFechaComprobCab
            .Parameters("@cabcomprobobservaciones") = UCase$(Trim$(frmantcomprobantes.TxObsComprobCab.Text))
            .Parameters("@cabcomprobglosa") = UCase$(Trim$(frmantcomprobantes.TxGlosaComprobCab.Text))
            .Parameters("@cabcomprobtotdebe") = Round(CDbl(VGvardllgen.ESNULO(frmantcomprobantes.LbTotales(0).Caption, 0)), 2)
            .Parameters("@cabcomprobtothaber") = Round(CDbl(VGvardllgen.ESNULO(frmantcomprobantes.LbTotales(1).Caption, 0)), 2)
            .Parameters("@cabcomprobtotussdebe") = Round(CDbl(VGvardllgen.ESNULO(frmantcomprobantes.LbTotales(3).Caption, 0)), 2)
            .Parameters("@cabcomprobtotusshaber") = Round(CDbl(VGvardllgen.ESNULO(frmantcomprobantes.LbTotales(4).Caption, 0)), 2)
            If frmantcomprobantes.VlGrabada Then
                .Parameters("@cabcomprobgrabada") = frmantcomprobantes.ChkGrabado.Value
              Else
                .Parameters("@cabcomprobgrabada") = 0
            End If
            If frmantcomprobantes.VlNref Then
                .Parameters("@cabcomprobnref") = Trim$(frmantcomprobantes.TxCtrNref.Text)
               Else
                .Parameters("@cabcomprobnref") = ""
            End If
            .Parameters("@usuariocodigo") = VGusuario
            .Parameters("@fechaact") = Now
            .Parameters("@estcomprobcodigo") = "01"
            .Parameters("@cabcomprobnlibro") = Numlibro
        End If
    End With
    VGCommandoSP.Execute
    Exit Sub
ErrorGrabaCabecera:
    VGvarVerifica = False
    VGErrorString = "Error en Grabar Cabecera " & Chr(13) & err.Description
    MsgBox (VGErrorString)
End Sub

Public Function NumeroComprob(mes As Integer, Optional ByRef numero As Double) As String
On Error GoTo Errnum
Dim Rsaux As ADODB.Recordset
Set Rsaux = New ADODB.Recordset

SQL = "select isnull(asientonumcorr" & Trim$(Format(mes, "00")) & ",0)+1 as numcorrelativo "
SQL = SQL & " from ct_asientocorre where empresacodigo='" & VGParametros.empresacodigo & "'"
SQL = SQL & " and asientocodigo='" & Trim$(frmantcomprobantes.VPAsiento) & "'"
SQL = SQL & " AND asientoanno='" & VGParamSistem.Anoproceso & "'"
     
Rsaux.Open SQL, VGCNx, adOpenKeyset, adLockReadOnly
    
    If Rsaux.RecordCount > 0 Then
       NumeroComprob = Trim$(Format(mes, "00")) & Trim$(frmantcomprobantes.VPAsiento) & Trim$(Format(Rsaux!numcorrelativo, "00000"))
       numero = Rsaux!numcorrelativo
       Else
        NumeroComprob = "00"
        numero = 0
    End If
    Exit Function
Errnum:
    VGvarVerifica = False
    VGErrorString = "Error en Numero de Comprobante " & Chr(13) & err.Description
End Function

Public Function NumeroComprobLibro(mes As Integer, ByVal Libro As String, Optional ByRef numero As Long) As String
On Error GoTo Errnum
    Dim Rsaux As ADODB.Recordset
    Set Rsaux = New ADODB.Recordset
    SQL = "select isnull(libronumcorr" & Trim$(Format(mes, "00")) & ",0)+1 as numcorrelativo from ct_librocorre "
    SQL = SQL & " where empresacodigo='" & VGParametros.empresacodigo & "' and librocodigo='" & Trim$(Libro) & "'"
    SQL = SQL & " AND libroanno='" & VGParamSistem.Anoproceso & "'"
    Rsaux.Open SQL, VGCNx, adOpenKeyset, adLockReadOnly
    If Rsaux.RecordCount > 0 Then
       NumeroComprobLibro = Trim$(Format(mes, "00")) & Trim$(Libro) & Trim$(Format(Rsaux!numcorrelativo, "000000"))
       numero = Rsaux!numcorrelativo
       Else
        NumeroComprobLibro = "00"
        numero = 0
    End If
    Exit Function
Errnum:
    VGvarVerifica = False
    VGErrorString = "Error en Numero de Comprobante Libro " & Chr(13) & err.Description
End Function


Public Sub CalculaComprob(ByVal NumComprob As String)
On Error GoTo ErrCalc
    Set VGCommandoSP = New ADODB.Command
    VGCommandoSP.ActiveConnection = VGGeneral
    VGCommandoSP.CommandType = adCmdStoredProc
    VGCommandoSP.CommandText = "ct_CalcComprob_pro"
    VGCommandoSP.Parameters.Refresh
    With VGCommandoSP
        .Parameters("@Servidor") = VGParamSistem.Servidor
        .Parameters("@base") = VGParamSistem.BDEmpresa
        .Parameters("@empresa") = VGParametros.empresacodigo
        .Parameters("@Ano") = VGParamSistem.Anoproceso
        .Parameters("@mes") = Format(CInt(VGParamSistem.Mesproceso), "0")
        .Parameters("@Asiento") = Trim$(frmantcomprobantes.VPAsiento)
        .Parameters("@SubAsiento") = Trim$(frmantcomprobantes.VPSubAsiento)
        .Parameters("@Comprob") = NumComprob
        .Execute
    End With
    Exit Sub
ErrCalc:
    VGvarVerifica = False
    VGErrorString = "Error en Calcular Comprobante " & Chr(13) & err.Description
End Sub

Public Sub ActualizaCorrelComprob(ByVal numero As Double)
On Error GoTo Actualizacorre
    Set VGCommandoSP = New ADODB.Command
    VGCommandoSP.ActiveConnection = VGGeneral
    VGCommandoSP.CommandType = adCmdStoredProc
    VGCommandoSP.CommandText = "ct_actcorrcomprob_pro"
    VGCommandoSP.Parameters.Refresh
    With VGCommandoSP
        .Parameters("@base") = VGParamSistem.BDEmpresa
        .Parameters("@empresa") = VGParametros.empresacodigo
        .Parameters("@anno") = VGParamSistem.Anoproceso
        .Parameters("@mes") = Format(CInt(VGParamSistem.Mesproceso), "00")
        .Parameters("@asiento") = Trim$(frmantcomprobantes.VPAsiento)
        .Parameters("@numero") = numero
        .Execute
    End With
    Exit Sub
Actualizacorre:
    VGvarVerifica = False
    VGErrorString = "Error en Actualizar Comprobante " & Chr(13) & err.Description
End Sub
Public Sub ActualizaCorrelComprobLibro(ByVal numero As Double, Libro As String)
On Error GoTo Actualizacorre
    Set VGCommandoSP = New ADODB.Command
    VGCommandoSP.ActiveConnection = VGGeneral
    VGCommandoSP.CommandType = adCmdStoredProc
    VGCommandoSP.CommandText = "ct_actcorrcomproblibro_pro"
    VGCommandoSP.Parameters.Refresh
    With VGCommandoSP
        .Parameters("@base") = VGParamSistem.BDEmpresa
        .Parameters("@anno") = VGParamSistem.Anoproceso
        .Parameters("@empresa") = VGParametros.empresacodigo
        .Parameters("@mes") = Format(CInt(VGParamSistem.Mesproceso), "00")
        .Parameters("@libro") = Trim$(Libro)
        .Parameters("@numero") = numero
        .Execute
    End With
    Exit Sub
Actualizacorre:
    VGvarVerifica = False
    VGErrorString = "Error en Actualizar Comprobante Libro " & Chr(13) & err.Description
    MsgBox (VGErrorString)
End Sub
Public Sub GrabarDetalle(ByVal rs As Recordset, NumComprob As String, Optional ByVal Numlibro As String)
On Error GoTo ErrorGrabaDetalle
Dim Rsaux As ADODB.Recordset
Dim NumeroComprobante As String
Dim numerodocumento As String
Set VGvardllgen = New dllgeneral.dll_general
    Set Rsaux = rs.Clone(adLockReadOnly)
    If Vgdocumentoanulado = "" Then
       Rsaux.Filter = "(montosol<>0 or montouss <> 0)"
     Else
         Rsaux.Filter = "operacioncodigo='" & Vgdocumentoanulado & "'"
    End If
    Set VGCommandoSP = New ADODB.Command
    Set VGvardllgen = New dllgeneral.dll_general
    'Elimar los Detalle antes de grabar
    VGCommandoSP.ActiveConnection = VGGeneral
    VGCommandoSP.CommandType = adCmdStoredProc
    VGCommandoSP.CommandText = "ct_grabardetallecomprob_pro"
    VGCommandoSP.Parameters.Refresh
    With VGCommandoSP
        .Parameters("@base") = VGParamSistem.BDEmpresa
        .Parameters("@tabla") = VGParamSistem.TablaDetcomprob
        .Parameters("@empresa") = VGParametros.empresacodigo
        .Parameters("@cabcomprobnumero") = NumComprob
        .Parameters("@detcomprobnlibro") = Numlibro
        .Parameters("@cabcomprobmes") = CInt(VGParamSistem.Mesproceso)
        .Parameters("@asientocodigo") = frmantcomprobantes.VPAsiento
        .Parameters("@subasientocodigo") = frmantcomprobantes.VPSubAsiento
        .Parameters("@op") = 2
        .Execute
    End With
    Rsaux.MoveFirst
    While Not Rsaux.EOF
        If Rsaux!detcomprobauto = 0 Then
        With VGCommandoSP
            .Parameters("@op") = 1
            .Parameters("@detcomprobitem") = Rsaux!detcomprobitem
            .Parameters("@operacioncodigo") = Rsaux!operacioncodigo
            .Parameters("@cuentacodigo") = Rsaux!cuentacodigo
            .Parameters("@centrocostocodigo") = VGvardllgen.ESNULO(Rsaux!centrocostocodigo, "00")
            .Parameters("@analiticocodigo") = Trim$(VGvardllgen.ESNULO(Rsaux!analiticocodigo, "00"))
            .Parameters("@detcomprobruc") = Rsaux!detcomprobruc
            .Parameters("@documentocodigo") = VGvardllgen.ESNULO(Rsaux!documentocodigo, "00")
             numerodocumento = ESNULO(Rsaux!detcomprobnumdocumento, "")
             If Len(numerodocumento) > 4 And IsNumeric(numerodocumento) Then
                numerodocumento = Left(numerodocumento, 4) & Right("0000000000" + Right(numerodocumento, Len(numerodocumento) - 4), 10)
             End If
            .Parameters("@detcomprobnumdocumento") = numerodocumento
            .Parameters("@detcomprobfechaemision") = Format(Rsaux!detcomprobfechaemision, "dd/mm/yyyy")
             If IsNull(Rsaux!detcomprobfechavencimiento) Then
                .Parameters("@detcomprobfechavencimiento") = Format(Rsaux!detcomprobfechaemision, "dd/mm/yyyy")
              Else
                .Parameters("@detcomprobfechavencimiento") = Format(Rsaux!detcomprobfechavencimiento, "dd/mm/yyyy")
            End If
            .Parameters("@detcomprobfecharef") = Rsaux!detcomprobfecharef
            .Parameters("@detcomprobglosa") = UCase$(Rsaux!detcomprobglosa)
            .Parameters("@monedacodigo") = Rsaux!monedacodigo
            .Parameters("@detcomprobtipocambio") = Rsaux!valcambio
            .Parameters("@detcomprobformacambio") = Rsaux!tcambio
            Select Case Rsaux!indicador
                Case "D"
                    .Parameters("@detcomprobdebe") = Round(Rsaux!montosol, 2)
                    .Parameters("@detcomprobussdebe") = Round(Rsaux!montouss, 2)
                    .Parameters("@detcomprobhaber") = 0
                    .Parameters("@detcomprobusshaber") = 0
                Case "H"
                    .Parameters("@detcomprobhaber") = Round(Rsaux!montosol, 2)
                    .Parameters("@detcomprobusshaber") = Round(Rsaux!montouss, 2)
                    .Parameters("@detcomprobdebe") = 0
                    .Parameters("@detcomprobussdebe") = 0
            End Select
            .Parameters("@detcomprobauto") = 0
            .Parameters("@detcomprobajusteuser") = Rsaux!detcomprobajusteuser
            .Parameters("@plantillaasientoinafecto") = Rsaux!plantillaasientoinafecto
            .Parameters("@tipdocref") = VGvardllgen.ESNULO(Rsaux!tipdocref, "00")
            .Parameters("@detcomprobnumref") = Rsaux!detcomprobnumref
            .Execute
        End With
        End If
        Rsaux.MoveNext
    Wend
    Exit Sub
ErrorGrabaDetalle:
    VGvarVerifica = False
    VGErrorString = "Error en Grabar Detalle " & Chr(13) & err.Description
    MsgBox (VGErrorString)
    Exit Sub
    Resume
End Sub
Public Sub GrabaAsientoAuto(Comp As String)
On Error GoTo GrabaAuto
    Set VGCommandoSP = New ADODB.Command
    VGCommandoSP.ActiveConnection = VGGeneral
    VGCommandoSP.CommandType = adCmdStoredProc
    VGCommandoSP.CommandText = "ct_grabaautomatico_pro"
    VGCommandoSP.Parameters.Refresh
    With VGCommandoSP
        .Parameters("@base") = VGParamSistem.BDEmpresa
        .Parameters("@tabla") = VGParamSistem.TablaDetcomprob
        .Parameters("@empresa") = VGParametros.empresacodigo
        .Parameters("@comp") = Comp
        .Parameters("@mes") = CInt(VGParamSistem.Mesproceso)
        .Parameters("@asiento") = frmantcomprobantes.VPAsiento
        .Parameters("@subasiento") = frmantcomprobantes.VPSubAsiento
        .Parameters("@tipo") = "0"
        .Execute
    End With
If VGParametros.AsientoAutoxCCostos = "1" Then
    Set VGCommandoSP = New ADODB.Command
    VGCommandoSP.ActiveConnection = VGGeneral
    VGCommandoSP.CommandType = adCmdStoredProc
    VGCommandoSP.CommandText = "ct_grabaautomatico_pro"
    VGCommandoSP.Parameters.Refresh
    With VGCommandoSP
        .Parameters("@base") = VGParamSistem.BDEmpresa
        .Parameters("@tabla") = VGParamSistem.TablaDetcomprob
        .Parameters("@empresa") = VGParametros.empresacodigo
        .Parameters("@comp") = Comp
        .Parameters("@mes") = CInt(VGParamSistem.Mesproceso)
        .Parameters("@asiento") = frmantcomprobantes.VPAsiento
        .Parameters("@subasiento") = frmantcomprobantes.VPSubAsiento
        .Parameters("@tipo") = "1"
        .Execute
    End With
End If
    Exit Sub
GrabaAuto:
    VGvarVerifica = False
    VGErrorString = "Error en Grabar Asientos Automaticos " & Chr(13) & err.Description
    MsgBox (VGErrorString)
End Sub
Public Sub MostrarCabecera(ByVal rs As ADODB.Fields)
    Dim Rsaux As ADODB.Fields
    Set Rsaux = rs
    Set VGvardllgen = New dll_general
    With frmantcomprobantes
        .lbNumComprobCab = Rsaux("cabcomprobnumero").Value
        .lbNumComprobCablibro = VGvardllgen.ESNULO(Rsaux("cabcomprobnlibro").Value, "")
        .DTPFechaComprobCab = Rsaux("cabcomprobfeccontable").Value
        .TxObsComprobCab.Text = Rsaux("cabcomprobobservaciones").Value
        .TxGlosaComprobCab.Text = Rsaux("cabcomprobglosa").Value
        .LbTotales(0).Caption = Format(Rsaux("cabcomprobtotdebe").Value, "###,###,###.00 ")
        .LbTotales(1).Caption = Format(Rsaux("cabcomprobtothaber").Value, "###,###,###.00 ")
        .LbTotales(2).Caption = Format(Rsaux("cabcomprobtotdebe").Value - Rsaux("cabcomprobtothaber").Value, "###,###,###.00 ")
        .LbTotales(3).Caption = Format(Rsaux("cabcomprobtotussdebe").Value, "###,###,###.00 ")
        .LbTotales(4).Caption = Format(Rsaux("cabcomprobtotusshaber").Value, "###,###,###.00 ")
        .LbTotales(5).Caption = Format(Rsaux("cabcomprobtotussdebe").Value - Rsaux("cabcomprobtotusshaber").Value, "###,###,###.00 ")
        .ChkGrabado.Value = IIf(VGvardllgen.ESNULO(Rsaux("cabcomprobgrabada").Value, False), 1, 0)
        .TxCtrNref.Text = VGvardllgen.ESNULO(Rsaux("cabcomprobnref").Value, "")
        .VPAsiento = Rsaux("asientocodigo").Value
        .VPSubAsiento = Rsaux("subasientocodigo").Value
    End With
End Sub
Public Sub MostrarDetalle(ByRef rs As ADODB.Recordset)
Dim Rsaux As ADODB.Recordset, i As Integer
On Error GoTo err
    Set Rsaux = New ADODB.Recordset
    Set VGCommandoSP = New ADODB.Command
    Set VGvardllgen = New dllgeneral.dll_general
    'Elimar los Detalle antes de grabar
    VGCommandoSP.ActiveConnection = VGGeneral
    VGCommandoSP.CommandType = adCmdStoredProc
    VGCommandoSP.CommandText = "ct_grabardetallecomprob_pro"
    VGCommandoSP.Parameters.Refresh
    With VGCommandoSP
        .Parameters("@base") = VGParamSistem.BDEmpresa
        .Parameters("@tabla") = VGParamSistem.TablaDetcomprob
        .Parameters("@empresa") = VGParametros.empresacodigo
        .Parameters("@cabcomprobnumero") = frmantcomprobantes.lbNumComprobCab
        .Parameters("@cabcomprobmes") = CInt(VGParamSistem.Mesproceso)
        .Parameters("@asientocodigo") = frmantcomprobantes.VPAsiento
        .Parameters("@subasientocodigo") = frmantcomprobantes.VPSubAsiento
        .Parameters("@op") = 3
        Set Rsaux = .Execute
    End With
If Rsaux.RecordCount > 0 Then
    Rsaux.MoveFirst
    While Not Rsaux.EOF
        rs.AddNew
        For i = 0 To Rsaux.Fields.Count - 1
               rs.Fields(Rsaux.Fields(i).Name) = Rsaux.Fields(i).Value
        Next
            rs!tdaux = Rsaux!documentocodigo
            rs!ndaux = Rsaux!detcomprobnumdocumento
            rs!Index = 2
        rs.Update
        Rsaux.MoveNext
    Wend
End If
If rs.RecordCount > 0 Then
        rs.AbsolutePosition = 1
        frmantcomprobantes.TDBG_Det.SetFocus
        frmantcomprobantes.lbnregdetalle.Caption = Format(rs.RecordCount, "0 ")
    Else
        frmantcomprobantes.lbnregdetalle.Caption = Format(0, "0 ")
End If
err:
Resume Next
End Sub
Public Function ValidarGrabarCabecera(NR As Long) As Boolean
    ValidarGrabarCabecera = False
With frmantcomprobantes
    'Validando Que exista por lo menos un registro en el detalle
    If NR = 0 Then
        MsgBox "Por lo menos debe haber ingresado un registro de detalle", vbInformation
        frmantcomprobantes.DTPFechaComprobCab.SetFocus
        Exit Function
    End If
    If Trim$(.TxGlosaComprobCab.Text) = "" Then
        MsgBox "Tiene que ingresar la glosa de la cabecera de comprobante"
        .TxGlosaComprobCab.SetFocus
        Exit Function
    End If
    ValidarGrabarCabecera = True
End With
End Function
Public Function ValidarGrabarDetalle() As Boolean
    ValidarGrabarDetalle = False
With frmantcomprobantes
    If (.CtrAyu_Opera.Enabled And .CtrAyu_Opera.Visible) And Trim$(.CtrAyu_Opera.xclave) = "" Then
        MsgBox "Debe ingresar el Codigo de Operacion", vbInformation
        .CtrAyu_Opera.SetFocus
        Exit Function
    End If
    If (.CtrAyu_Cuenta.Enabled And .CtrAyu_Cuenta.Visible) And (Trim$(.CtrAyu_Cuenta.xclave) = "" Or Trim$(.CtrAyu_Cuenta.xclave) = "00") Then
        MsgBox "Debe ingresar la Cuenta", vbInformation
        .CtrAyu_Cuenta.SetFocus
        Exit Function
    End If
    If (.CtrAyu_CCosto.Enabled And .CtrAyu_CCosto.Visible) And Trim$(.CtrAyu_CCosto.xclave) = "" Then
        MsgBox "Debe ingresar el Centro de Costo", vbInformation
        .CtrAyu_CCosto.SetFocus
        Exit Function
    End If
    If (.CtrAyu_TipAnal.Enabled And .CtrAyu_TipAnal.Visible) And Trim$(.CtrAyu_TipAnal.xclave) = "" Then
        MsgBox "Debe ingresar el Tipo de Analitico", vbInformation
        .CtrAyu_TipAnal.SetFocus
        Exit Function
    End If
    If (.CtrAyu_Analitico.Enabled And .CtrAyu_Analitico.Visible) And Trim$(.CtrAyu_Analitico.xclave) = "" Then
        MsgBox "Debe ingresar el Codigo de Analitico", vbInformation
        .CtrAyu_Analitico.SetFocus
        Exit Function
    End If
    If (.CtrAyu_TipDoc.Enabled And .CtrAyu_TipDoc.Visible) And Trim$(.CtrAyu_TipDoc.xclave) = "" Then
        MsgBox "Debe ingresar el Tipo de Documento", vbInformation
        .CtrAyu_TipDoc.SetFocus
        Exit Function
    End If
    
    If (.TxSerie.Enabled And .TxSerie.Visible) And Trim$(.TxSerie.Text & .TxNdoc.Text) = "" Then
        MsgBox "Debe Ingrear el Numero del Documento", vbInformation
        .TxSerie.SetFocus
        Exit Function
    End If
    
    If (.TxGlosa.Enabled And .TxGlosa.Visible) And Trim$(.TxGlosa.Text) = "" Then
    '     MsgBox "Debe Ingrear la Glosa del Detalle", vbInformation
    '     .TxGlosa.SetFocus
   '      Exit Function
    End If
    
    If (.CtrAyu_Moneda.Enabled And .CtrAyu_Moneda.Visible) And Trim$(.CtrAyu_Moneda.xclave) = "" Then
        MsgBox "Debe ingresar el Codigo de Moneda", vbInformation
       .CtrAyu_Moneda.SetFocus
        Exit Function
    End If
    If CDbl(.lb_vcambio.Caption) = 0 Then
        MsgBox "No existe tipo cambio para esta fecha", vbInformation
        .Dtp_FechaDoc.SetFocus
        Exit Function
    End If
    
'    If CDbl(.TxMonto.valor) = 0 Then
'        MsgBox "El monto ha ingresar debe ser mayor a 0", vbInformation
'        Exit Function
'    End If
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
        If rs!operacioncodigo = "01" Then
            doc = Trim$(VGvardllgen.ESNULO(rs!documentocodigo, "")) & Trim$(VGvardllgen.ESNULO(rs!detcomprobnumdocumento, ""))
            docaux = Trim$(VGvardllgen.ESNULO(rs!tdaux, "")) & Trim$(VGvardllgen.ESNULO(rs!ndaux, ""))
            If doc <> docaux And Trim$(rs!analiticocodigo) <> "00" Then
                If ValidaDoc(rs!cuentacodigo) = True Then
                If Not Validadocum(rs!analiticocodigo, rs!documentocodigo, rs!detcomprobnumdocumento) Then
                    MsgBox "Este documento ya ha sido registrado", vbExclamation
                    frmantcomprobantes.TxNdoc.SetFocus
                    Exit Function
                End If
                End If
            End If
        End If
        rs.MoveNext
    Wend
    frmantcomprobantes.LbTotales(2).Caption = IIf(frmantcomprobantes.LbTotales(2).Caption = "", 0, frmantcomprobantes.LbTotales(2).Caption)
    If CDbl(frmantcomprobantes.LbTotales(2).Caption) <> 0 Then
       MsgBox "El Asiento se encuentra descuadrado en Soles ", vbExclamation
           If VGParametros.CuadreAsiento And Not (frmantcomprobantes.CtrAyu_Asiento.xclave = VGParametros.asientocodigo And frmantcomprobantes.CtrAyu_SubAsiento.xclave = VGParametros.subasientocodigo) Then Exit Function
    End If
    
    ValidarRsDetalle = True
End Function
Public Function CargarPlantillaAsiento(ByRef rs As Recordset, ByVal Asiento As String, ByVal SubAsiento As String) As Boolean
Dim Rsaux As ADODB.Recordset
Dim criterio As String
CargarPlantillaAsiento = False
    Set Rsaux = New ADODB.Recordset
    Rsaux.Open "select * from ct_plantillaasiento where asientocodigo='" & Trim$(Asiento) & _
               "' and subasientocodigo='" & Trim$(SubAsiento) & "' order by plantillaasientocorrela", VGCNx, adOpenKeyset, adLockReadOnly
    If Rsaux.RecordCount = 0 Then Exit Function
    Rsaux.MoveFirst
    While Not Rsaux.EOF
      With rs
        .AddNew
        If Rsaux!cuentacodigo = "00" Then
            If Rsaux!plantillaasientocomodin <> "" Then
                criterio = ArmaCriterioComodin(Rsaux!plantillaasientocomodin)
            End If
            !comodin = criterio
        End If
        !detcomprobitem = Format(Rsaux!plantillaasientocorrela, "00000")
        !detcomprobitemaux = Format(Rsaux!plantillaasientocorrela, "00000")
        !operacioncodigo = Rsaux!operacioncodigo
        !cuentacodigo = Rsaux!cuentacodigo
        !detcomprobfechaemision = frmantcomprobantes.DTPFechaComprobCab
        !indicador = Rsaux!iddebeohaber
  '      !monedacodigo = VGParametros.monedabase
        !tcambio = "02"
        !valcambio = VGValorCambio
        !detcomprobauto = 0
        !plantillaasientoinafecto = Rsaux!plantillaasientoinafecto
        !detcomprobglosa = VGGlosa
        !monedacodigo = VGMonSubAsiento
        !Index = 2
        !NumPlantilla = Rsaux!plantillaasientocorrela
        .Update
      End With
    Rsaux.MoveNext
    Wend
    
    rs.AbsolutePosition = 1
CargarPlantillaAsiento = True
End Function
Private Function ArmaCriterioComodin(cad As String) As String
Dim pos As Integer, cadaux As String, criterio As String
Dim valor As String
    criterio = ""
    Do While True
        pos = InStr(1, cad, "%", vbTextCompare)
        If pos = 0 Then Exit Do
        valor = "'" & Left(cad, pos) & "'"
        cad = Right(cad, (Len(cad) - pos))
        criterio = criterio & "cuentacodigo like " & valor & " or "
    Loop
    ArmaCriterioComodin = Left(criterio, Len(criterio) - 3)
End Function
Public Function RecuperaTipoCambio(Fecha As String, tipo As tipocambio) As Double
Dim Rsaux As ADODB.Recordset
Set Rsaux = New ADODB.Recordset
'FIXIT: Declare 'fechaIni' con un tipo de datos de enlace en tiempo de compilación         FixIT90210ae-R1672-R1B8ZE
Dim fechaIni, fechaDato As Date
Dim Campo As String
fechaIni = "01/01/" & VGParamSistem.Anoproceso
fechaDato = Fecha
'FIXIT: Reemplazar la función 'Str' con la función 'Str$'.                                 FixIT90210ae-R9757-R1B8ZE
If fechaIni > fechaDato Then Fecha = "31/12/" & Str(VGParamSistem.Anoproceso - 1)
    RecuperaTipoCambio = 0
    Select Case tipo
        Case Compra
            Campo = "tipocambiocompra"
        Case Venta
            Campo = "tipocambioventa"
        Case Promedio
            Campo = "tipocambiopromedio"
    Case Else
            Campo = "tipocambioventa"
    End Select
    Rsaux.Open "Select Valor=isnull(" & Campo & ",0)  from ct_tipocambio where tipocambiofecha ='" & Fecha & "'", VGCNx, adOpenKeyset, adLockReadOnly
    If Rsaux.RecordCount > 0 Then
        RecuperaTipoCambio = Rsaux!valor
        frmantcomprobantes.lbnregdetalle.Caption = Format(Rsaux.RecordCount, "0 ")
    End If
End Function
Public Function Validadocum(analitico As String, tipdoc As String, ndoc As String) As Boolean
Dim sqlcad As String
Dim Rsaux As ADODB.Recordset
    Validadocum = True
    Set Rsaux = New ADODB.Recordset
    sqlcad = "select * from dbo.ct_detcomprob" & VGParamSistem.Anoproceso & _
             " where analiticocodigo='" & Trim$(analitico) & "' and " & _
             " operacioncodigo='01' and documentocodigo='" & Trim$(tipdoc) & "' and " & _
             " detcomprobnumdocumento='" & Trim$(ndoc) & "' "
    Rsaux.Open sqlcad, VGCNx, adOpenKeyset, adLockReadOnly
    If Rsaux.RecordCount = 0 Then Exit Function
    Validadocum = False
End Function
Public Function Verifidocumacancelar(Optional ByRef moneda As String, Optional Fecha As Date, _
                                     Optional montototal As Double, Optional montoacumamort As Double) As Boolean
Dim sqlcad As String
Dim Rsaux As ADODB.Recordset
Dim Ncomprob As String
    Ncomprob = Trim$(frmantcomprobantes.TxSerie.Text) & IIf(Trim$(frmantcomprobantes.TxSerie.Text) = "", "", "-") & Trim$(frmantcomprobantes.TxNdoc.Text)
    Verifidocumacancelar = False
    Set Rsaux = New ADODB.Recordset
    sqlcad = "select * from dbo.ct_detcomprob" & VGParamSistem.Anoproceso & _
             " where analiticocodigo='" & Trim$(frmantcomprobantes.CtrAyu_Analitico.xclave) & "' and " & _
             " isnull(operacioncodigo,'')='01' and documentocodigo='" & Trim$(frmantcomprobantes.CtrAyu_TipDoc.xclave) & "' and " & _
             " detcomprobnumdocumento='" & Ncomprob & "" & "' "
    Rsaux.Open sqlcad, VGCNx, adOpenKeyset, adLockReadOnly
    If Rsaux.RecordCount = 0 Then Exit Function
    Verifidocumacancelar = True
End Function

Public Sub CalculodeAjuste(ByRef rs As ADODB.Recordset, Optional flagnotipocambio As Boolean = False)
On Error GoTo ERRX
'FIXIT: Declare 'POSANTERIOR' con un tipo de datos de enlace en tiempo de compilación      FixIT90210ae-R1672-R1B8ZE
Dim POSANTERIOR As Variant
Dim Rsaux As ADODB.Recordset, monto As Double
Dim ctaajuste As String, ind As String, sicuentaajuste As Integer
Dim difsoles As Double, difdolares As Double
    Set VGvardllgen = New dll_general
    difsoles = CDbl(VGvardllgen.ESNULO(frmantcomprobantes.LbTotales(2).Caption, 0)) 'diferencia en soles
    difdolares = CDbl(VGvardllgen.ESNULO(frmantcomprobantes.LbTotales(5).Caption, 0)) 'diferencia en dolares
    VGactulizodoc = True
    Set Rsaux = New ADODB.Recordset
    Rsaux.Fields.Append "CtaAjustevar", adVarChar, 20
    Rsaux.Fields.Append "Indi", adVarChar, 1
    Rsaux.Open
With rs
    POSANTERIOR = .AbsolutePosition
    .MoveFirst
    While Not .EOF
        sicuentaajuste = cuentaajuste(!cuentacodigo, ind)
        If sicuentaajuste = 1 And ind = "D" Then
            Rsaux.AddNew
            Rsaux!CtaAjustevar = !cuentacodigo
            Rsaux!indi = "D"
            Rsaux.Update
        End If
        If sicuentaajuste = 1 And ind = "H" Then
            Rsaux.AddNew
            Rsaux!CtaAjustevar = !cuentacodigo
            Rsaux!indi = "H"
            Rsaux.Update
        End If
        .MoveNext
    Wend
    If Rsaux.RecordCount = 0 Then
        .Filter = 0
        'Modificado:03/02/2003 ICS
        If POSANTERIOR < 0 Then Exit Sub
        .AbsolutePosition = POSANTERIOR
        VGactulizodoc = False
        Exit Sub
    End If
    Call CalcularTotalessinajuste(rs, Rsaux, difsoles, difdolares)
    'En el caso que no cuadre en la moneda base que seria el sol
    'Y cuadre en la moneda en dolares
    Dim campoajustar As String
    
    If Not flagnotipocambio Then
        If difsoles <> 0 And difdolares = 0 Then
            If difsoles < 0 Then campoajustar = "D"
            If difsoles > 0 Then campoajustar = "H"
        Else
            campoajustar = "N"
        End If
      Else
        campoajustar = "N"
    End If
    'Solamente dos cuentas de ajuste deben existir
    Rsaux.MoveFirst
    .Filter = 0
    While Not Rsaux.EOF
        .Filter = 0
        .Filter = "cuentacodigo='" & Trim$(Rsaux!CtaAjustevar) & "'"
        If .RecordCount > 0 Then
            If Rsaux!indi = campoajustar Then
                !montosol = Abs(difsoles)
              Else
                !montosol = 0
            End If
            !montouss = 0
            .Update
        End If
        Rsaux.MoveNext
    Wend
         
    .Filter = 0
    .AbsolutePosition = POSANTERIOR
    VGactulizodoc = False
End With

ERRX:

End Sub

Private Sub CalcularTotalessinajuste(ByVal rs As Recordset, ByVal rsajust As Recordset, ByRef difersoles As Double, _
                            ByRef diferdolares As Double)
Dim Rsaux As ADODB.Recordset, rs2 As ADODB.Recordset
Dim vardllgen As New dllgeneral.dll_general
    Set Rsaux = rs.Clone(adLockReadOnly)
    Set rs2 = rsajust.Clone(adLockReadOnly)
Dim montosoles As Double, montodolares As Double
Dim difsoles As Double, difdolares As Double
Dim montosolesDebe As Double, montodolaresDebe As Double
Dim montosolesHaber As Double, montodolaresHaber As Double
Dim cad As String
    cad = ""
    If rs2.RecordCount = 0 Then Exit Sub
    rs2.MoveFirst
    Do While Not rs2.EOF
        cad = cad & "cuentacodigo<>'" & rs2!CtaAjustevar & "' and "
        rs2.MoveNext
    Loop
    cad = Trim$(Left(cad, Len(cad) - 4))
    
    montosolesDebe = 0: montodolaresDebe = 0:
    montosolesHaber = 0: montodolaresHaber = 0:
    difsoles = 0: difdolares = 0
    Rsaux.Filter = cad
    Rsaux.MoveFirst
    While Not Rsaux.EOF
        If Rsaux!indicador = "D" Then
            montosolesDebe = montosolesDebe + vardllgen.ESNULO(Rsaux!montosol, 0)
            montodolaresDebe = montodolaresDebe + vardllgen.ESNULO(Rsaux!montouss, 0)
          Else
            montosolesHaber = montosolesHaber + vardllgen.ESNULO(Rsaux!montosol, 0)
            montodolaresHaber = montodolaresHaber + vardllgen.ESNULO(Rsaux!montouss, 0)
        End If
        Rsaux.MoveNext
    Wend
    difsoles = montosolesDebe - montosolesHaber
    difdolares = montodolaresDebe - montodolaresHaber
    difersoles = difsoles
    diferdolares = difdolares
    Rsaux.Filter = 0
End Sub

'  *****************************JCGI
Private Function ValidaDoc(ByVal cuenta As String) As Boolean
Dim sqldoc As String
Dim rsdoc As ADODB.Recordset
Set rsdoc = New ADODB.Recordset
ValidaDoc = False
sqldoc = "Select cuentadocumento From ct_cuenta Where empresacodigo='" & VGParametros.empresacodigo & "' And cuentacodigo='" & cuenta & "' And cuentadocumento='1' "
rsdoc.Open sqldoc, VGCNx, adOpenKeyset, adLockReadOnly
    If rsdoc.RecordCount > 0 Then ValidaDoc = True
End Function
