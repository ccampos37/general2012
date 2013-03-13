VERSION 5.00
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Object = "{4D137D9C-00A6-4458-9B46-7E95DB76D55B}#9.0#0"; "TextFer.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Begin VB.Form frmAnularBorraRecibos 
   Caption         =   "Anular Recibos"
   ClientHeight    =   4740
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9210
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4740
   ScaleWidth      =   9210
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   1935
      Left            =   120
      TabIndex        =   20
      Top             =   2640
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   3413
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame3 
      Height          =   2535
      Left            =   7320
      TabIndex        =   17
      Top             =   0
      Width           =   1815
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   750
         Left            =   240
         TabIndex        =   19
         Top             =   240
         Width           =   1350
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   750
         Left            =   120
         TabIndex        =   18
         Top             =   1440
         Width           =   1470
      End
   End
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   135
      TabIndex        =   13
      Top             =   30
      Width           =   7080
      Begin VB.CommandButton cmdMostrar 
         Caption         =   "..."
         Height          =   255
         Left            =   6510
         TabIndex        =   2
         Top             =   390
         Width           =   300
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1365
         TabIndex        =   0
         Top             =   285
         Width           =   1755
      End
      Begin TextFer.TxFer TxFer1 
         Height          =   330
         Left            =   4695
         TabIndex        =   1
         Top             =   345
         Width           =   1650
         _ExtentX        =   2910
         _ExtentY        =   582
         Object.CausesValidation=   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   ""
         ColorIlumina    =   -2147483624
         SaltarAlEnter   =   -1  'True
         Valor           =   ""
         NoCaracteres    =   "0123456789"
         MarcarTextoAlEnfoque=   -1  'True
         NoRangoCadena   =   -1  'True
      End
      Begin VB.Label Label2 
         Caption         =   "Numero Recibo"
         Height          =   390
         Left            =   3780
         TabIndex        =   15
         Top             =   345
         Width           =   825
      End
      Begin VB.Label Label1 
         Caption         =   "Seleccionar Tipo"
         Height          =   375
         Left            =   180
         TabIndex        =   14
         Top             =   300
         Width           =   870
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   7095
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Operacion 
         Height          =   390
         Left            =   3945
         TabIndex        =   16
         Top             =   240
         Width           =   2940
         _ExtentX        =   5186
         _ExtentY        =   688
         Enabled         =   0   'False
         XcodMaxLongitud =   0
         NomTabla        =   "te_operaciongeneral"
         ListaCampos     =   "operacioncodigo(1),operaciondescripcion(1)"
         XcodCampo       =   "operacioncodigo"
         XListCampo      =   "operaciondescripcion"
         ListaCamposDescrip=   "Código,Descripción"
         ListaCamposText =   "operacioncodigo,operaciondescripcion"
         Requerido       =   0   'False
      End
      Begin VB.Label lbl 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   4
         Left            =   3840
         TabIndex        =   12
         Top             =   1005
         Width           =   3000
      End
      Begin VB.Label lbl 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   3
         Left            =   1320
         TabIndex        =   11
         Top             =   1110
         Width           =   1560
      End
      Begin VB.Label Label7 
         Caption         =   "Importe Total"
         Height          =   375
         Left            =   3030
         TabIndex        =   10
         Top             =   1050
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "Moneda"
         Height          =   315
         Left            =   150
         TabIndex        =   9
         Top             =   1170
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Tipo Operación"
         Height          =   435
         Left            =   2790
         TabIndex        =   8
         Top             =   240
         Width           =   1035
      End
      Begin VB.Label lbl 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   1
         Left            =   1320
         TabIndex        =   7
         Top             =   690
         Width           =   5460
      End
      Begin VB.Label Label4 
         Caption         =   "Razón Social"
         Height          =   375
         Left            =   150
         TabIndex        =   6
         Top             =   780
         Width           =   1095
      End
      Begin VB.Label lbl 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   0
         Left            =   1320
         TabIndex        =   5
         Top             =   255
         Width           =   1230
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha"
         Height          =   255
         Left            =   150
         TabIndex        =   4
         Top             =   300
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmAnularBorraRecibos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vgdllgeneral As dll_general
Dim Vlncomprob As String
Private Sub cmdaceptar_Click()
On Error GoTo X
  Dim xCodigo As String
  Dim flag As String
  Dim rs As ADODB.Recordset
  
  Set rs = New ADODB.Recordset
  Set rs = VGCNx.Execute("Select clientecodigo,empresacodigo,comprobconta from te_cabecerarecibos where cabrec_numrecibo='" & Trim(TxFer1.Text) & "'")
  If Not rs.BOF And Not rs.EOF Then
     xCodigo = rs(0)
  End If
  Set rs = Nothing
  
  Set rs = New ADODB.Recordset
  Set rs = VGCNx.Execute("select operacioncontrolaclienteprov from te_operaciongeneral where operacioncodigo='" & Trim(Ctr_Operacion.xclave) & "'")
  If Not rs.BOF And Not rs.EOF Then
     flag = rs(0)
  Else
     flag = "ZZ"
  End If
  If ValidaAnular = True Then
     VGCNx.BeginTrans
     Call AnulaRecIngresoEgreso
     Call AnularDetalleTesoreria
     
     Dim rsc As ADODB.Recordset
     Set rsc = VGCNx.Execute("Select empresacodigo,comprobconta From te_cabecerarecibos Where cabrec_numrecibo='" & TxFer1.Text & "'")
     If rsc!comprobconta <> Empty Then Call EliminaContab(Trim(TxFer1.Text), rsc!empresacodigo, rsc!comprobconta)
     Set rsc = Nothing
     
     Select Case flag
       Case "P":
          Call EliminarDataAbonos("P")
          Call ActualizarDataCargos(xCodigo)
       Case "C":
          Call EliminarDataAbonos("C")
          Call ActualizarDataCargosClientes(xCodigo)
     End Select
     'Anulo el asiento generado en CONTABILIDAD o blanqueo el asiento
 
' o j o
 
 '    Call AnulaenConta(Vlncomprob, VGCnxCT)
     
     VGCNx.CommitTrans
     MsgBox "Recibo Nº " & TxFer1.Text & " fue Anulado Satisfactoriamente", vbInformation, Caption
  End If
  Call LimpiarData

  Exit Sub

X:
   'If Err Then
   '    Err = 0
       MsgBox "El Proceso de Anulación no se culminó ...!!!" & Chr(13) & Chr(10) & Err.Number & "-" & Err.Description, vbInformation, MsgTitle
       VGCNx.RollbackTrans
       Call LimpiarData
       Exit Sub
    Resume
   'End If

End Sub
Private Sub AnulaenConta(Comprobante As String, conex As ADODB.Connection)
    Dim sqlcad As String
    sqlcad = "" & _
    " Update dbo.ct_cabcomprob" & VGParamSistem.AnoProceso & _
    " Set cabcomprobtotdebe=0, " & _
    "     cabcomprobtothaber=0," & _
    "     cabcomprobtotussdebe=0, " & _
    "     cabcomprobtotusshaber = 0 " & _
    " Where cabcomprobnumero='" & Comprobante & "' " & Chr(13) & _
    " Update dbo.ct_detcomprob" & VGParamSistem.AnoProceso & _
    "   Set detcomprobdebe=0, " & _
    "   detcomprobhaber=0, " & _
    "   detcomprobussdebe=0, " & _
    "   detcomprobusshaber = 0 " & _
    " Where cabcomprobnumero='" & Comprobante & "' "
    VGCnxCT.Execute sqlcad

End Sub
Private Sub CmdCancelar_Click()
   Unload Me
End Sub

Private Sub cmdMostrar_Click()
  Call MuestraDatos
End Sub

Private Sub Form_Load()
   Combo1.Clear
   Combo1.AddItem "I-INGRESOS"
   Combo1.AddItem "E-EGRESOS"
   Combo1.ListIndex = 0
   Ctr_Operacion.conexion VGCNx
End Sub

Sub AnulaRecIngresoEgreso()
  Dim SQL As String
 
  SQL = "Update te_cabecerarecibos set cabrec_estadoreg='1' where "
  SQL = SQL & "cabrec_numrecibo='" & Trim(TxFer1.Text) & "'"
  VGCNx.Execute (SQL)

End Sub

Sub AnularDetalleTesoreria()
   Dim SQL As String
   SQL = "Update te_detallerecibos set detrec_estadoreg='1' where "
   SQL = SQL & "cabrec_numrecibo='" & Trim(TxFer1.Text) & "'"
   VGCNx.Execute (SQL)

End Sub

Sub EliminarDataAbonos(flagTipo As String)
Dim SQL As String
Dim rs As New ADODB.Recordset
Dim rb As New ADODB.Recordset
      
    Select Case flagTipo
      Case "C":
         SQL = "select * from te_detallerecibos a inner join te_cabecerarecibos b "
         SQL = SQL & " on a.cabrec_numrecibo=b.cabrec_numrecibo where "
         SQL = SQL & "a.cabrec_numrecibo='" & Trim(TxFer1.Text) & "'"
         Set rs = VGCNx.Execute(SQL)
         rs.MoveFirst
         Do Until rs.EOF()
            SQL = "select tdocumentocodigo,tdocumentotipo from cc_tipodocumento "
            SQL = SQL & " where tdocumentocodigo='" & rs!detrec_tipodoc_concepto & "'"
            Set rb = VGCNx.Execute(SQL)
            If (rs!cabrec_ingsal = "I" And rb!tdocumentotipo = "C") Or (rs!cabrec_ingsal = "E" And rb!tdocumentotipo = "A") Then
               SQL = "update vt_abono SET abonocanflreg=1 Where "
               SQL = SQL & " documentoabono='" & rs!detrec_tipodoc_concepto & "' and "
               SQL = SQL & " abononumdoc='" & rs!detrec_numdocumento & "' and "
               SQL = SQL & " abononumplanilla ='" & rs!cabrec_numrecibo & "'"
             Else
               SQL = "update vt_cargo Set cargoapeflgreg=1 where "
               SQL = SQL & " documentocargo='" & rs!detrec_tipodoc_concepto & "' and "
               SQL = SQL & " cargonumdoc='" & rs!detrec_numdocumento & "' and "
               SQL = SQL & " abononumplanilla ='" & rs!cabrec_numrecibo & "'"
            End If
            Call VGCNx.Execute(SQL)
            rs.MoveNext
         Loop
         
      
      Case "P":
         SQL = "select * from te_detallerecibos a inner join te_cabecerarecibos b "
         SQL = SQL & " on a.cabrec_numrecibo=b.cabrec_numrecibo where "
         SQL = SQL & "a.cabrec_numrecibo='" & Trim(TxFer1.Text) & "'"
         Set rs = VGCNx.Execute(SQL)
         rs.MoveFirst
         Do Until rs.EOF()
            SQL = "select tdocumentocodigo,tdocumentotipo from cp_tipodocumento "
            SQL = SQL & " where tdocumentocodigo='" & rs!detrec_tipodoc_concepto & "'"
            Set rb = VGCNx.Execute(SQL)
            If (rs!cabrec_ingsal = "E" And rb!tdocumentotipo = "C") Or (rs!cabrec_ingsal = "I" And rb!tdocumentotipo = "A") Then
               SQL = "update cp_abono SET abonocanflreg=1 Where "
               SQL = SQL & " documentoabono='" & rs!detrec_tipodoc_concepto & "' and "
               SQL = SQL & " abononumdoc='" & rs!detrec_numdocumento & "' and "
               SQL = SQL & " abononumplanilla ='" & rs!cabrec_numrecibo & "'"
             Else
               SQL = "update cp_cargo Set cargoapeflgreg=1 where "
               SQL = SQL & " documentocargo='" & rs!detrec_tipodoc_concepto & "' and "
               SQL = SQL & " cargonumdoc='" & rs!detrec_numdocumento & "' and "
               SQL = SQL & " abononumplanilla ='" & rs!cabrec_numrecibo & "'"
            End If
            Call VGCNx.Execute(SQL)
            rs.MoveNext
         Loop
   End Select
End Sub

Sub ActualizarDataCargos(xValor As String)
    Dim SQL As String
    Dim rb As New ADODB.Recordset
    Dim rbabono As New ADODB.Recordset
    Dim rbcargo As New ADODB.Recordset
    
    DoEvents
    ' Acumula los abonos de los documentos
      SQL = "select * "
      SQL = SQL & "FROM cp_cargo aa,"
      SQL = SQL & "(select"
      SQL = SQL & " a.cabrec_numrecibo,a.cabrec_tipocambio,"
      SQL = SQL & " a.monedacodigo,a.clientecodigo,"
      SQL = SQL & " b.detrec_tipodoc_concepto,b.detrec_numdocumento,"
      SQL = SQL & " b.detrec_monedacancela , b.detrec_importesoles, b.detrec_importedolares "
      SQL = SQL & "From "
      SQL = SQL & " te_cabecerarecibos a, te_detallerecibos b "
      SQL = SQL & "Where "
      SQL = SQL & "a.cabrec_numrecibo=b.cabrec_numrecibo and "
      SQL = SQL & "a.cabrec_numrecibo='" & Trim(TxFer1.Text) & "' and "
      SQL = SQL & " a.clientecodigo<>'' ) as ZZ "
      SQL = SQL & "Where"
      SQL = SQL & " aa.documentocargo=ZZ.detrec_tipodoc_concepto and "
      SQL = SQL & " aa.cargonumdoc=ZZ.detrec_numdocumento and "
      SQL = SQL & " aa.clientecodigo=ZZ.clientecodigo "
                              
      Set rb = VGCNx.Execute(SQL)
                        
    If rb.RecordCount > 0 Then
        rb.MoveFirst
        Do Until rb.EOF
            'Actualizar los saldos en cp_cargo
            Set rbcargo = VGCNx.Execute("select * from cp_cargo where documentocargo='" & rb!documentocargo & "' and cargonumdoc='" & rb!cargonumdoc & "' and clientecodigo like '" & xValor & "'")
            If rbcargo.RecordCount > 0 Then
                VGCNx.Execute ("Update cp_cargo set cargoapeimppag=0,cargoapeflgcan=0,cargoapefeccan=null where documentocargo='" & rb!documentocargo & "' and cargonumdoc='" & rb!cargonumdoc & "' and clientecodigo like '" & xValor & "'")
                If rbcargo.Fields("monedacodigo") = g_TipoSol Then
                    Set rbabono = VGCNx.Execute("select documentoabono,abononumdoc,abonocanfecpla," & _
                                        " round(sum( case abonocanmoncan when '02' then " & _
                                        " (isnull(abonocanimpcan,0)*isnull(abonocantipcam,1)) else 0 end),2)," & _
                                        " round(sum( case abonocanmoncan when '01' then " & _
                                        " isnull(abonocanimpcan,0) else 0 end),2) " & _
                                        " From cp_abono " & _
                                        " Inner join cp_tipodocumento " & _
                                        " on cp_abono.documentoabono=cp_tipodocumento.tdocumentocodigo " & _
                                        " where cp_abono.documentoabono='" & rbcargo!documentocargo & "' and cp_abono.abononumdoc='" & rbcargo!cargonumdoc & "' and abonocancli like '" & xValor & "' " & _
                                        " and abonocanflreg <> 1 group by documentoabono,abononumdoc,abonocanfecpla")
                ElseIf rbcargo.Fields("monedacodigo") = g_TipoDolar Then
                      Set rbabono = VGCNx.Execute("select documentoabono,abononumdoc,abonocanfecpla," & _
                                        " round(sum( case abonocanmoncan when '02' then " & _
                                        " isnull(abonocanimpcan,0) else 0 end),2)," & _
                                        " round(sum( case abonocanmoncan when '01' then " & _
                                        " (isnull(abonocanimpcan,0)/isnull(abonocantipcam,1)) else 0 end),2) " & _
                                        " From cp_abono " & _
                                        " Inner join cp_tipodocumento " & _
                                        " on cp_abono.documentoabono=cp_tipodocumento.tdocumentocodigo " & _
                                        " where cp_abono.documentoabono='" & rbcargo!documentocargo & "' and cp_abono.abononumdoc='" & rbcargo!cargonumdoc & "' and abonocancli like '" & xValor & "' " & _
                                        " and abonocanflreg<>1 group by documentoabono,abononumdoc,abonocanfecpla")
                  
                End If
                If rbabono.RecordCount > 0 Then
                    VGCNx.Execute "Update cp_cargo " & _
                                "Set cargoapeimppag=" & rbabono.Fields(3) + rbabono.Fields(4) & _
                                " Where documentocargo='" & rbcargo.Fields("documentocargo") & "' and cargonumdoc='" & rbcargo.Fields("cargonumdoc") & "' and clientecodigo like '" & xValor & "'"
                                
                    VGCNx.Execute "Update cp_cargo " & _
                                "Set cargoapeflgcan= case Round(isnull(cargoapeimpape,0),2)-Round(isnull(cargoapeimppag,0),2) when 0 then '1' else '0' end ," & _
                                " cargoapefeccan=case Round(isnull(cargoapeimpape,0),2)-Round(isnull(cargoapeimppag,0),2) when 0 then '" & rbabono.Fields(2) & "' else null end  " & _
                                "Where documentocargo='" & rbcargo.Fields("documentocargo") & "' and cargonumdoc='" & rbcargo.Fields("cargonumdoc") & "' and clientecodigo like '" & xValor & "'"
                 
                End If
            End If
            rbcargo.Close
            Set rbcargo = Nothing
            DoEvents
            rb.MoveNext
        Loop
    End If
    rb.Close
    Set rb = Nothing
    'Actualizar los Saldos del Proveedor
   Set rb = VGCNx.Execute("select  clientecodigo,round(sum( case monedacodigo when '02' then " & _
                        " (isnull(cargoapeimpape,0)-isnull(cargoapeimppag,0)) else 0 end),2)," & _
                        " round(sum( case monedacodigo when '01' then " & _
                        " isnull(cargoapeimpape,0)-isnull(cargoapeimppag,0) else 0 end),2) " & _
                        " From cp_cargo " & _
                        " Inner join cp_tipodocumento " & _
                        " on cp_cargo.documentocargo=cp_tipodocumento.tdocumentocodigo " & _
                        " where clientecodigo like '" & xValor & "'" & _
                        " group by clientecodigo")
    
    If rb.RecordCount > 0 Then
        rb.MoveFirst
        Do Until rb.EOF
            VGCNx.Execute "update cp_proveedor " & _
                       " Set clientesaldodolares=0,clientesaldosoles=0" & _
                       " Where clientecodigo='" & rb!clientecodigo & "'"
        
            VGCNx.Execute "update cp_proveedor " & _
                       " Set clientesaldodolares=isnull(clientesaldodolares,0)+" & rb.Fields(1) & ",clientesaldosoles=isnull(clientesaldosoles,0)+" & rb.Fields(2) & _
                       " Where clientecodigo='" & rb!clientecodigo & "'"
            
            DoEvents
            rb.MoveNext
        Loop
    End If
    rb.Close
    Set rb = Nothing

End Sub

Sub ActualizarDataCargosClientes(xValor As String)
    Dim SQL As String
    Dim rb As New ADODB.Recordset
    Dim rbabono As New ADODB.Recordset
    Dim rbcargo As New ADODB.Recordset
    
    DoEvents
    ' Acumula los abonos de los documentos
      SQL = "select * "
      SQL = SQL & "FROM vt_cargo aa,"
      SQL = SQL & "(select"
      SQL = SQL & " a.cabrec_numrecibo,a.cabrec_tipocambio,"
      SQL = SQL & " a.monedacodigo,a.clientecodigo,"
      SQL = SQL & " b.detrec_tipodoc_concepto,b.detrec_numdocumento,"
      SQL = SQL & " b.detrec_monedacancela , b.detrec_importesoles, b.detrec_importedolares "
      SQL = SQL & "From "
      SQL = SQL & " te_cabecerarecibos a, te_detallerecibos b "
      SQL = SQL & "Where "
      SQL = SQL & "a.cabrec_numrecibo=b.cabrec_numrecibo and "
      SQL = SQL & "a.cabrec_numrecibo='" & Trim(TxFer1.Text) & "' and "
      SQL = SQL & " a.clientecodigo<>'' ) as ZZ "
      SQL = SQL & "Where"
      SQL = SQL & " aa.documentocargo=ZZ.detrec_tipodoc_concepto and "
      SQL = SQL & " aa.cargonumdoc=ZZ.detrec_numdocumento and "
      SQL = SQL & " aa.clientecodigo=ZZ.clientecodigo "
                        
      Set rb = VGCNx.Execute(SQL)
                        
    If rb.RecordCount > 0 Then
        rb.MoveFirst
        Do Until rb.EOF
            'Actualizar los saldos en vt_cargo del documento
            Set rbcargo = VGCNx.Execute("select * from vt_cargo where documentocargo='" & rb!documentocargo & "' and cargonumdoc='" & rb!cargonumdoc & "' and clientecodigo like '" & xValor & "'")
            If rbcargo.RecordCount > 0 Then
                VGCNx.Execute ("Update vt_cargo set cargoapeimppag=0,cargoapeflgcan=0,cargoapefeccan=null where documentocargo='" & rb!documentocargo & "' and cargonumdoc='" & rb!cargonumdoc & "' and clientecodigo like '" & xValor & "'")
                If rbcargo.Fields("monedacodigo") = g_TipoSol Then
                    Set rbabono = VGCNx.Execute("select documentoabono,abononumdoc,abonocanfecpla," & _
                                        " round(sum( case abonocanmoncan when '02' then " & _
                                        " (isnull(abonocanimpcan,0)*isnull(abonocantipcam,1)) else 0 end),2)," & _
                                        " round(sum( case abonocanmoncan when '01' then " & _
                                        " isnull(abonocanimpcan,0) else 0 end),2) " & _
                                        " From vt_abono " & _
                                        " Inner join cp_tipodocumento " & _
                                        " on vt_abono.documentoabono=cp_tipodocumento.tdocumentocodigo " & _
                                        " where vt_abono.documentoabono='" & rbcargo!documentocargo & "' and vt_abono.abononumdoc='" & rbcargo!cargonumdoc & "' and abonocancli like '" & xValor & "' " & _
                                        " and isnull(vt_abono.abonocanflreg,0)<>1 group by documentoabono,abononumdoc,abonocanfecpla")
                ElseIf rbcargo.Fields("monedacodigo") = g_TipoDolar Then
                      Set rbabono = VGCNx.Execute("select documentoabono,abononumdoc,abonocanfecpla," & _
                                        " round(sum( case abonocanmoncan when '02' then " & _
                                        " isnull(abonocanimpcan,0) else 0 end),2)," & _
                                        " round(sum( case abonocanmoncan when '01' then " & _
                                        " (isnull(abonocanimpcan,0)/isnull(abonocantipcam,1)) else 0 end),2) " & _
                                        " From vt_abono " & _
                                        " Inner join cp_tipodocumento " & _
                                        " on vt_abono.documentoabono=cp_tipodocumento.tdocumentocodigo " & _
                                        " where vt_abono.documentoabono='" & rbcargo!documentocargo & "' and vt_abono.abononumdoc='" & rbcargo!cargonumdoc & "' and abonocancli like '" & xValor & "' " & _
                                        " and isnull(vt_abono.abonocanflreg,0)<>1 group by documentoabono,abononumdoc,abonocanfecpla")
                  
                End If
                If rbabono.RecordCount > 0 Then
                    VGCNx.Execute "Update vt_cargo " & _
                                "Set cargoapeimppag=" & rbabono.Fields(3) + rbabono.Fields(4) & _
                                " Where documentocargo='" & rbcargo.Fields("documentocargo") & "' and cargonumdoc='" & rbcargo.Fields("cargonumdoc") & "' and clientecodigo like '" & xValor & "'"
                                
                    VGCNx.Execute "Update vt_cargo " & _
                                "Set cargoapeflgcan= case Round(isnull(cargoapeimpape,0),2)-Round(isnull(cargoapeimppag,0),2) when 0 then '1' else '0' end ," & _
                                " cargoapefeccan=case Round(isnull(cargoapeimpape,0),2)-Round(isnull(cargoapeimppag,0),2) when 0 then '" & rbabono.Fields(2) & "' else null end  " & _
                                "Where documentocargo='" & rbcargo.Fields("documentocargo") & "' and cargonumdoc='" & rbcargo.Fields("cargonumdoc") & "' and clientecodigo like '" & xValor & "'"
                 
                End If
            End If
            rbcargo.Close
            Set rbcargo = Nothing
            DoEvents
            rb.MoveNext
        Loop
    End If
    rb.Close
    Set rb = Nothing
    'Actualizar los Saldos del Cliente
   Set rb = VGCNx.Execute("select  clientecodigo,round(sum( case monedacodigo when '02' then " & _
                        " (isnull(cargoapeimpape,0)-isnull(cargoapeimppag,0)) else 0 end),2)," & _
                        " round(sum( case monedacodigo when '01' then " & _
                        " isnull(cargoapeimpape,0)-isnull(cargoapeimppag,0) else 0 end),2) " & _
                        " From vt_cargo " & _
                        " Inner join cp_tipodocumento " & _
                        " on vt_cargo.documentocargo=cp_tipodocumento.tdocumentocodigo " & _
                        " where clientecodigo like '" & xValor & "'" & _
                        " group by clientecodigo")
    
    If rb.RecordCount > 0 Then
        rb.MoveFirst
        Do Until rb.EOF
            VGCNx.Execute "update vt_cliente " & _
                       " Set clientesaldodolares=0,clientesaldosoles=0" & _
                       " Where clientecodigo='" & rb.Fields(0) & "'"
        
            VGCNx.Execute "update vt_cliente " & _
                       " Set clientesaldodolares=isnull(clientesaldodolares,0)+" & rb.Fields(1) & ",clientesaldosoles=isnull(clientesaldosoles,0)+" & rb.Fields(2) & _
                       " Where clientecodigo='" & rb.Fields(0) & "'"
            
            DoEvents
            rb.MoveNext
        Loop
    End If
    rb.Close
    Set rb = Nothing

End Sub

Function ValidaAnular() As Boolean
 Dim SQL As String
 Dim rs As New ADODB.Recordset
  If lbl(4).Caption = Empty Then
     MsgBox "Falta seleccionar un Nº de Recibo de Ingreso/Egreso", vbInformation, Caption
     Combo1.SetFocus
     ValidaAnular = False
     Exit Function
  End If
     
  Set rs = New ADODB.Recordset
  SQL = "Select count(*) from te_cabecerarecibos where cabrec_numrecibo='" & Trim(TxFer1.Text) & "' AND "
  SQL = SQL & "cabrec_estadoreg='1'"
  Set rs = VGCNx.Execute(SQL)
  If rs(0) > 0 Then
     MsgBox "El Nº de Recibo " & TxFer1.Text & " se encuentra Anulado", vbInformation, Caption
     ValidaAnular = False
     Exit Function
  End If
  
  SQL = "Select count(*) from te_cabecerarecibos where cabrec_numrecibo='" & Trim(TxFer1.Text) & "'"
  Set rs = VGCNx.Execute(SQL)
  If rs(0) = 0 Then
     MsgBox "El Nº de Recibo " & TxFer1.Text & " no Existe en la Base", vbInformation, Caption
     ValidaAnular = False
     Exit Function
  End If
  
  SQL = "Select count(*) from te_cabecerarecibos where cabrec_numrecibo='" & Trim(TxFer1.Text) & "' and "
  SQL = SQL & "cabrec_ingsal='" & Left(Combo1.List(Combo1.ListIndex), 1) & "'"
  Set rs = VGCNx.Execute(SQL)
  If rs(0) = 0 Then
     MsgBox "El Nº de Recibo " & TxFer1.Text & " no corresponde al Tipo de Movimiento", vbInformation, Caption
     ValidaAnular = False
     Exit Function
  End If
  SQL = "Select cabrec_estadoreg from te_cabecerarecibos Where cabrec_numreciboegreso<>'' "
  SQL = SQL & " and cabrec_numrecibo='" & TxFer1.Text & "' And cabrec_transferenciaautomatico='1'"
  Set rs = VGCNx.Execute(SQL)
  If rs.RecordCount > 0 Then
     If rs(0) = 1 Then
        MsgBox "El Nº de Recibo " & TxFer1.Text & " Es de Transferencia  y esta ANULADO", vbInformation, Caption
      Else
        MsgBox "El Nº de Recibo " & TxFer1.Text & " Es de Transferencia , usar la opcion Anulacion de Transferencia", vbInformation, Caption
      End If
      ValidaAnular = False
      Exit Function
 End If
  
  
  ValidaAnular = True
End Function

Sub MuestraDatos()
 Dim rs As New ADODB.Recordset
 Dim SQL As String
   SQL = "Select cabrec_numrecibo,cabrec_fechadocumento,"
   SQL = SQL & "razonsocial=isnull("
   SQL = SQL & "case when B.operacioncontrolaclienteprov='C'"
   SQL = SQL & "then (select clienterazonsocial from vt_cliente CC where CC.clientecodigo=A.clientecodigo)"
   SQL = SQL & "else (select clienterazonsocial from cp_proveedor CC where CC.clientecodigo=A.clientecodigo) "
   SQL = SQL & "end,''),operaciondescripcion,monedacodigo,cabrec_totsoles,cabrec_totdolares,A.operacioncodigo,comprobconta=isnull(A.cabcomprobnumero,'')"
   SQL = SQL & "FROM te_cabecerarecibos A,te_operaciongeneral B "
   SQL = SQL & "where A.operacioncodigo=B.operacioncodigo and cabrec_numrecibo='" & Trim(TxFer1.Text) & "'"
   Set rs = New ADODB.Recordset
   Set rs = VGCNx.Execute(SQL)
   If Not rs.BOF And Not rs.EOF Then
      lbl(0).Caption = rs!cabrec_fechadocumento
      lbl(1).Caption = rs!razonsocial
      Ctr_Operacion.xclave = rs!operacioncodigo: Call Ctr_Operacion.Ejecutar
      lbl(3).Caption = rs!monedacodigo
      lbl(4).Caption = IIf(rs!monedacodigo = "01", rs!cabrec_totsoles, rs!cabrec_totdolares)
      Vlncomprob = rs!comprobconta
      DataGrid1.Caption = "Doc.! Contable : " + Vlncomprob
   End If
   Set DataGrid1.DataSource = rs
   DataGrid1.Refresh

End Sub

Private Sub TxFer1_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    Call cmdMostrar_Click
  End If

End Sub

Sub LimpiarData()
 Dim i As Integer
  For i = 0 To 1
    lbl(i).Caption = Empty
  Next
  For i = 3 To 4
    lbl(i).Caption = Empty
  Next
End Sub

Private Sub EliminaContab(ByVal recibo As String, ByVal empresa As String, ByVal ComprobNumero As String)
  Dim rsc As ADODB.Recordset
  Dim SQL As String
  
  Set rsc = New ADODB.Recordset
  Set rsc = VGCNx.Execute("Select Year(cabrec_fechadocumento) From te_cabecerarecibos Where cabrec_numrecibo='" & recibo & "'")
  If Not rsc.EOF Then
    SQL = "Delete From ct_cabcomprob" & rsc(0) & " Where empresacodigo='" & empresa & "'"
    SQL = SQL & " And cabcomprobnumero='" & ComprobNumero & "'"
    SQL = SQL & " And Substring(cabcomprobnprovi,4,6)='" & recibo & "'"
    
    VGCNx.Execute (SQL)
  End If
End Sub

