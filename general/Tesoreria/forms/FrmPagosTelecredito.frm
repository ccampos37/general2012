VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Begin VB.Form FrmPagosTelecredito 
   Caption         =   "Form1"
   ClientHeight    =   9210
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12420
   LinkTopic       =   "Form1"
   ScaleHeight     =   9210
   ScaleWidth      =   12420
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   9135
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12285
      _ExtentX        =   21669
      _ExtentY        =   16113
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "Pagos Telecreditos"
      TabPicture(0)   =   "FrmPagosTelecredito.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).ControlCount=   0
   End
   Begin VB.Label Label5 
      Caption         =   "Label5"
      Height          =   495
      Left            =   6120
      TabIndex        =   1
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Line Line1 
      X1              =   6450
      X2              =   7665
      Y1              =   4470
      Y2              =   4965
   End
End
Attribute VB_Name = "FrmPagosTelecredito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents rsconcil As ADODB.Recordset
Attribute rsconcil.VB_VarHelpID = -1
Dim RsSaldoIni As ADODB.Recordset
Dim rsconcil1 As New ADODB.Recordset
Dim tmontosolesDebe As Double, tmontodolaresDebe As Double
Dim tmontosolesHaber As Double, tmontodolaresHaber As Double
Dim montosolesDebe As Double, montodolaresDebe As Double
Dim montosolesHaber As Double, montodolaresHaber As Double
Dim mtsoles As Double, mtdolar As Double
Public SQL As String
Public numero As String
Dim tsoles As Double, tdolar As Double
Dim montoextbanc As Double
Dim mon As String
Dim mon_descripcion As String
Dim Modificar As Integer
Dim flagcal As Boolean
Dim dllgeneral As dllgeneral.dll_general

Private Sub cmdeliminar_click()
TxtNrorendicion.Enabled = True
 Modificar = 1
 Cmdimprimir(0).Enabled = True
 Cmdimprimir(1).Enabled = False
 Cmdimprimir(2).Enabled = False
 Cmdcancelar.Enabled = True
 cmdaceptar.Enabled = True
 Modificar = 2
 Call Listar(Modificar)
 If MsgBox("desea Eliminar Rendicion", vbQuestion + vbYesNo) = vbYes Then
   rsconcil.MoveFirst
   If rsconcil.RecordCount() > 0 Then
      Do Until rsconcil.EOF
          rsconcil("chkconcil").Value = 0
          rsconcil.MoveNext
       Loop
    End If
    Call cmdaceptar_Click
 End If
 
End Sub

Private Sub cmdmodificar_click()
 TxtNrorendicion.Enabled = True
 Modificar = 1
 Cmdimprimir(0).Enabled = True
 Cmdimprimir(1).Enabled = False
 Cmdimprimir(2).Enabled = False
 Cmdcancelar.Enabled = True
 cmdaceptar.Enabled = True
 Call Listar(Modificar)
End Sub


Private Sub chkconciliado_Click()
On Error GoTo x1
Modificar = 0
Call Listar(Modificar)
    rsconcil.MoveFirst
    If rsconcil.RecordCount() > 0 Then
       Do Until rsconcil.EOF
        If chkconciliado.Value Then
            rsconcil("chkconcil").Value = 1
             rsconcil("Importepago") = rsconcil("saldo")
           Else
            rsconcil("chkconcil").Value = 0
            rsconcil("Importepago") = 0 ' rsconcil("saldo")
          End If
          rsconcil.MoveNext
       Loop
    End If
    
  Call CalcularTotal(rsconcil)
       Call CalcularTotales(rsconcil)

    Exit Sub
x1:
 ' MsgBox "No se pudo Grabar " & err.Description & " - " & err.Number, vbInformation, Caption
 Resume Next
End Sub

Private Sub Ctr_AyudaMoneda_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
Dim vardllgen As New dllgeneral.dll_general
Dim rsql As New ADODB.Recordset
    mon = ColecCampos("monedacodigo").Value
    mon_descripcion = ColecCampos("monedadescripcion").Value
    lbMon.Caption = IIf(ColecCampos("monedacodigo").Value = "01", "Moneda de origen : Soles", "Moneda de Origen : Dolares ")
    
    SQL = " select * from te_codigocaja where cajacodigo='" & Ctr_AyudaCaja.xclave & "'"
    Set rsql = VGCNx.Execute(SQL)
    Modificar = 0
    Call Listar(Modificar)
    Call Listartransferencias(Modificar)

    Select Case mon
        Case "01":
            LeDolares.Visible = False
            LbTotales(3).Visible = False
            LbTotales(4).Visible = False
            LbTotales(5).Visible = False
            
            leSoles.Visible = True
            LbTotales(0).Visible = True
            LbTotales(1).Visible = True
            LbTotales(2).Visible = True
            TDBG_concil.Columns(7).Visible = True
            TDBG_concil.Columns(9).Visible = False
            
            TxtNrorendicion.Text = rsql!rendicionnumero01
            
        Case "02"
            leSoles.Visible = False
            LbTotales(0).Visible = False
            LbTotales(1).Visible = False
            LbTotales(2).Visible = False
            TDBG_concil.Columns(9).Visible = True
            TDBG_concil.Columns(7).Visible = False
            
            LeDolares.Visible = True
            LbTotales(3).Visible = True
            LbTotales(4).Visible = True
            LbTotales(5).Visible = True
            TxtNrorendicion.Text = rsql!rendicionnumero02
    
    End Select
    cmdaceptar.Enabled = False
    Cmdeliminar.Enabled = True
    Cmdimprimir(0).Enabled = False
    Cmdimprimir(1).Enabled = False
    Cmdimprimir(2).Enabled = True
    
End Sub



Private Sub Command2_Click()
LblRazonsocial = rsconcil!razonsocial
LblDocumento = rsconcil!cargodocumento + "-" + rsconcil!cargonumdoc
TxFerImporte.Text = 0
FramePagos.Visible = True
End Sub

Private Sub Ctr_Ayudabanco_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
Ctr_AyudaCuentabanco.Filtro = "cbanco_codigo='" & Ctr_Ayudabanco.xclave & "'"
End Sub

Private Sub Ctr_AyudaCaja_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
Ctr_AyudaCuentabanco.Filtro = ""
End Sub

Private Sub Ctr_AyudaCuentabanco_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
    Call Listar(0)
    Cmdimprimir(0).Enabled = False
    Cmdimprimir(1).Enabled = False
    Cmdimprimir(2).Enabled = True
End Sub
Private Sub Form_Load()
    Call Ctr_Ayuempresa.Conexion(VGCNx)
    Call Ctr_Ayudabanco.Conexion(VGCNx)
    Call Ctr_AyudaCuentabanco.Conexion(VGCNx)
    DTPfechaini.Value = VGParamSistem.fechatrabajo
    TDBG_concil.FetchRowStyle = True
End Sub

Private Sub cmdaceptar_Click()
Dim x As Integer
Dim rsql As New ADODB.Recordset
    
Cmdimprimir(0).Enabled = True
Cmdimprimir(1).Enabled = True
Cmdimprimir(2).Enabled = True
cmdaceptar.Enabled = False
Call Grabar
cmdaceptar.Enabled = False
End Sub

Private Sub CmdCancelar_Click()
    If rsconcil Is Nothing Then
        Unload Me
        Exit Sub
    End If
    
    rsconcil.CancelBatch
    Unload Me
End Sub
Private Sub Listar(Optional OP As Integer)
Dim rs As New ADODB.Recordset
Dim vardllgen As New dllgeneral.dll_general
If ExisteElem(0, VGCNx, VGComputer & "_telec") Then VGCNx.Execute (" drop table " & VGComputer & "_telec")
SQL = " select chkconcil=0,clienterazonsocial,saldo=cargoapeimpape-cargoapeimppag,importepago=cargoapeimpape-cargoapeimppag,a.* "
SQL = SQL & " into " & VGComputer & "_telec  from cp_cargo a inner join cp_proveedor b on a.clientecodigo=b.clientecodigo"
SQL = SQL & " where empresacodigo='" & Ctr_Ayuempresa.xclave & "' and isnull(cargoapeflgreg,0)=0 and isnull(cargoapeflgcan,0)=0 "
Set rs = VGCNx.Execute(SQL)
    Set rsconcil = New ADODB.Recordset
    rsconcil.Open (" select * from " & VGComputer & "_telec"), VGCNx, adOpenDynamic, adLockBatchOptimistic
    If rsconcil.RecordCount > 0 Then
   Set TDBG_concil.DataSource = rsconcil
      TDBG_concil.Refresh
       lbnreg.Caption = Format(rsconcil.RecordCount, "0 ")
 
       Call CalcularTotal(rsconcil)
       Call CalcularTotales(rsconcil)
         TxSaldofin.valor = 0
         Txtsaldoini.valor = 0
    End If
End Sub

Private Sub CalcularTotales(ByVal rs As Recordset)
Dim rsaux As ADODB.Recordset
Dim vardllgen As New dllgeneral.dll_general
Set rsaux = rs.Clone(adLockReadOnly)


montosolesDebe = 0: montodolaresDebe = 0:
montosolesHaber = 0: montodolaresHaber = 0:
mtsoles = 0: mtdolar = 0

If rsaux.BOF = True Or rsaux.EOF = True Then Exit Sub
Dim Fecha As Double
Fecha = DTPfechaini.Value

rsaux.MoveFirst
    While Not rsaux.EOF
      If rsaux("chkconcil").Value <> 0 Then
        montosolesDebe = montosolesDebe + IIf(rsaux!monedacodigo = "01", rsaux!importepago, 0)
        montodolaresDebe = montodolaresDebe + IIf(rsaux!monedacodigo = "01", rsaux!importepago, 0)
'        montosolesHaber = montosolesHaber + IIf(rsaux!cabrec_ingsal = "E", vardllgen.ESNULO(rsaux!detrec_importesoles, 0), 0)
'        montodolaresHaber = montodolaresHaber + IIf(rsaux!cabrec_ingsal = "E", vardllgen.ESNULO(rsaux!detrec_importedolares, 0), 0)
      End If
    rsaux.MoveNext
    Wend
    'Soles
    mtsoles = ((tmontosolesDebe - montosolesDebe) - (tmontosolesHaber - montosolesHaber)) + montoextbanc
    LbTotales(0).Caption = Format(tmontosolesDebe - montosolesDebe, "###,###,###,###.00 ") ' Debe
    LbTotales(1).Caption = Format(tmontosolesHaber - montosolesHaber, "###,###,###,###.00 ") ' Haber
    LbTotales(2).Caption = Format(mtsoles, "###,###,###,###.00 ")   ' Haber
    'Dolares
    mtdolar = ((tmontodolaresDebe - montodolaresDebe) - (tmontodolaresHaber - montodolaresHaber)) + montoextbanc
    LbTotales(3).Caption = Format(tmontodolaresDebe - montodolaresDebe, "###,###,###,###.00 ") ' Debe
    LbTotales(4).Caption = Format(tmontodolaresHaber - montodolaresHaber, "###,###,###,###.00 ") ' Haber
    LbTotales(5).Caption = Format(mtdolar, "###,###,###,###.00 ") ' Haber
    
    If mon = "01" Then
        lbtot(0).Caption = Format(montosolesDebe, "###,###,###,###.00")
        lbtot(1).Caption = Format(montosolesHaber, "###,###,###,###.00")
        lbtot(2).Caption = Format(montosolesDebe - montosolesHaber, "###,###,###,###.00")
      Else
        lbtot(0).Caption = Format(montodolaresDebe, "###,###,###,###.00")
        lbtot(1).Caption = Format(montodolaresHaber, "###,###,###,###.00")
        lbtot(2).Caption = Format(montodolaresDebe - montodolaresHaber, "###,###,###,###.00")
    End If
    TxSaldofin.Text = Round(CDbl(vardllgen.ESNULO(Espunto(Txtsaldoini.Text), 0)) + CDbl(lbtot(0).Caption) - CDbl(lbtot(1).Caption), 2)
End Sub
Private Sub CalcularTotal(ByVal rs As Recordset)
Dim rsaux As ADODB.Recordset
Dim vardllgen As New dllgeneral.dll_general
Set rsaux = rs.Clone(adLockReadOnly)

    tmontosolesDebe = 0: tmontodolaresDebe = 0:
    tmontosolesHaber = 0: tmontodolaresHaber = 0:
    tsoles = 0: tdolar = 0
    If rsaux.BOF = True Or rsaux.EOF = True Then Exit Sub
    rsaux.MoveFirst
    montoextbanc = CDbl(vardllgen.ESNULO(Espunto(TxSaldofin.valor), 0))
    While Not rsaux.EOF
        tmontosolesDebe = tmontosolesDebe + IIf(rsaux!monedacodigo = "01", importepago, 0)
        tmontodolaresDebe = tmontodolaresDebe + IIf(rsaux!monedacodigo = "02", importepago, 0)
 '       tmontosolesHaber = tmontosolesHaber + IIf(rsaux!cabrec_ingsal = "E", vardllgen.ESNULO(rsaux!detrec_importesoles, 0), 0)
 '       tmontodolaresHaber = tmontodolaresHaber + IIf(rsaux!cabrec_ingsal = "E", vardllgen.ESNULO(rsaux!detrec_importedolares, 0), 0)
        rsaux.MoveNext
    Wend
    'Soles
    tsoles = tmontosolesDebe - tmontosolesHaber + montoextbanc
    LbTotales(0).Caption = Format(tmontosolesDebe, "###,###,###,###.00 ") ' Debe
    LbTotales(1).Caption = Format(tmontosolesHaber, "###,###,###,###.00 ") ' Haber
    LbTotales(2).Caption = Format(tsoles, "###,###,###,###.00 ")     ' Total
    'Dolares
    tdolar = tmontodolaresDebe - tmontodolaresHaber + montoextbanc
    LbTotales(3).Caption = Format(tmontodolaresDebe, "###,###,###,###.00 ") ' Debe
    LbTotales(4).Caption = Format(tmontodolaresHaber, "###,###,###,###.00 ") ' Haber
    LbTotales(5).Caption = Format(tdolar, "###,###,###,###.00 ") ' Haber
End Sub

Private Sub RsConcil_FieldChangeComplete(ByVal cFields As Long, ByVal Fields As Variant, ByVal pError As ADODB.error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
Static cont As Integer
   ' If flagcal Then Exit Sub
    cmdaceptar.Enabled = True
    Cmdimprimir(0).Enabled = False
    Cmdimprimir(1).Enabled = False
    Cmdimprimir(2).Enabled = False
     Cmdimprimir(0).Enabled = True
    Cmdimprimir(1).Enabled = True
    Cmdimprimir(2).Enabled = True
    TDBG_concil.Refresh
End Sub

Private Sub cmdimprimir_Click(index As Integer)
    If rsconcil.RecordCount = 0 Then Exit Sub
Dim valor As String
    Select Case index
        Case 0: valor = "1"
        Case 1: valor = "2"
        Case 2: valor = "0"
    End Select
    Call Imprimir(valor)
End Sub
Private Sub Imprimir(ValorConci As String)
Dim vardllgen As New dllgeneral.dll_general
Dim arrform(7) As Variant, arrparm(6) As Variant
Dim NombreRep As String, CadOrden As String
Dim Fecha As String
Dim fecha1 As String
    fecha1 = Format(DateSerial(DTPfechaini.Year, DTPfechaini.Month, 1), "dd/mm/yyyy")
  
    arrparm(0) = VGParamSistem.BDEmpresa
    arrparm(1) = Trim(Ctr_AyudaCaja.xclave)
    arrparm(2) = Trim(Ctr_AyudaMoneda.xclave)
    arrparm(3) = Format(DTPfechaini.Value, "dd/mm/yyyy")
    
    Select Case ValorConci
        Case "0": arrform(0) = "Todos"
        Case "1": arrform(0) = "Conciliados"
        Case "2": arrform(0) = "Pendientes"
    End Select
    If ValorConci = "2" Then
       arrparm(4) = recibosrendicion(2, rsconcil)
     ElseIf ValorConci = "1" Then
             arrparm(4) = recibosrendicion(1, rsconcil)
         Else
             arrparm(4) = "XX"
    End If
    arrparm(5) = "0"
    
    arrform(1) = "Oficina='" & Ctr_AyudaOficina.xnombre & "'"
    arrform(2) = "Caja='" & Ctr_AyudaCaja.xnombre & "'"
    arrform(3) = "mon='" & mon_descripcion & "'"
    arrform(4) = "Fecha='" & Format(DTPfechaini.Value, "dd/mm/yyyy") & "'"
    arrform(5) = "Saldoinicial=" & Txtsaldoini.valor
    arrform(6) = "Nrorendicion=" & Format(TxtNrorendicion.Text, "000000")
    NombreRep = "xx_Rendiciones.rpt"
    Call ImpresionRptProc(NombreRep, arrform, arrparm, , "Rendiciones")
End Sub


Private Sub TDB_transferencias_DblClick()
If Ctr_AyudaMoneda.xclave = "01" Then
   TxtsaldoiniDocxrendir.Text = TDB_transferencias.Columns(7)
Else
   TxtsaldoiniDocxrendir.Text = TDB_transferencias.Columns(7)
End If
Call ListarDocxRendir(Modificar)
End Sub

Private Sub TDB_transferencias_GotFocus()
Call ListarDocxRendir(Modificar)
End Sub

Private Sub TDBG_concil_DblClick()

End Sub

Private Sub TDBG_concil_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid70.StyleDisp)
Dim rsclone As New ADODB.Recordset
    On Error GoTo x
     Set rsclone = rsconcil.Clone(adLockReadOnly)
    If rsclone.RecordCount = 0 Then Exit Sub
    rsclone.Bookmark = Bookmark
    If rsclone!chkconcil = 1 Then
       RowStyle.BackColor = RGB(200, 250, 100)
    End If
    flagcal = True
    Call CalcularTotal(rsconcil)
    Call CalcularTotales(rsconcil)

    Exit Sub
x:
Resume Next

End Sub



Private Sub TDBG_concil_HeadClick(ByVal ColIndex As Integer)
 TDBG_concil.Refresh
 On Error GoTo y
 With rsconcil
    If .Sort = Empty Then
        .Sort = TDBG_concil.Columns.Item(ColIndex).DataField & " asc"
    ElseIf Right(.Sort, 3) = "asc" Then
        .Sort = TDBG_concil.Columns.Item(ColIndex).DataField & " desc"
    ElseIf Right(.Sort, 4) = "desc" Then
        .Sort = TDBG_concil.Columns.Item(ColIndex).DataField & " asc"
    End If
    TDBG_concil.Refresh
 End With
y:
End Sub

Private Sub Grabar()
Dim rsql As New ADODB.Recordset
Dim xx As String
rsconcil.MoveFirst
Do Until rsconcil.EOF()
      xx = 0
      If Escadena(rsconcil!rendicionnumero) <> "" Then
        xx = 1
      End If
   If IsNull(rsconcil!chkconcil) = False And rsconcil!chkconcil Or xx = 1 Then
      SQL = "update te_detallerecibos set fechconcil ='" & rsconcil!fechconcil & "',"
      SQL = SQL & " chkconcil=" & IIf(rsconcil!chkconcil, 1, 0) & ", rendicionnumero='" & TxtNrorendicion.Text & "'"
      SQL = SQL & " where cabrec_numrecibo='" & rsconcil!cabrec_numrecibo & "'"
      SQL = SQL & " and detrec_item='" & rsconcil!detrec_item & "'"
      Set rsql = VGCNx.Execute(SQL)
   End If
   rsconcil.MoveNext
Loop
cmdaceptar.Enabled = True
End Sub


