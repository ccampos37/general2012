VERSION 5.00
Object = "{4D137D9C-00A6-4458-9B46-7E95DB76D55B}#9.0#0"; "TextFer.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Begin VB.Form frmAnularTransferencia 
   Caption         =   "Anulación de Transferencias"
   ClientHeight    =   4035
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5370
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4035
   ScaleWidth      =   5370
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   1815
      Left            =   480
      TabIndex        =   5
      Top             =   960
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   3201
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
   Begin VB.CommandButton Command1 
      Caption         =   "...."
      Height          =   375
      Left            =   4320
      TabIndex        =   1
      Top             =   360
      Width           =   495
   End
   Begin TextFer.TxFer txtTransf 
      Height          =   345
      Left            =   1815
      TabIndex        =   0
      Tag             =   "1"
      Top             =   390
      Width           =   2325
      _ExtentX        =   4101
      _ExtentY        =   609
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
      SaltarAlEnter   =   -1  'True
      Valor           =   ""
      NoCaracteres    =   "0123456789"
      MarcarTextoAlEnfoque=   -1  'True
      NoRangoCadena   =   -1  'True
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   3300
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   2880
      TabIndex        =   4
      Top             =   3300
      Width           =   1200
   End
   Begin VB.Label Label2 
      Caption         =   "Nº Transferencia"
      Height          =   315
      Left            =   420
      TabIndex        =   2
      Top             =   420
      Width           =   1230
   End
End
Attribute VB_Name = "frmAnularTransferencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim rs1 As ADODB.Recordset

Private Sub cmdaceptar_Click()
 On Error GoTo X
  Dim rs As ADODB.Recordset
  Dim SQL As String
  
  Set rs = New ADODB.Recordset
  SQL = "Select cabrec_numrecibo,empresacodigo,comprobconta,clientecodigo from te_cabecerarecibos "
  SQL = SQL & "Where cabrec_numreciboegreso<>'' and cabrec_numreciboegreso='" & TxtTransf.Text & "' "
  SQL = SQL & "and cabrec_estadoreg<>'1'"
  
  Set rs = VGCNx.Execute(SQL)
  If Not rs.BOF And Not rs.EOF Then
     rs.MoveFirst
     VGCNx.BeginTrans
     Do Until rs.EOF
        Call AnulaRecIngresoEgreso(rs(0))
        Call EliminarDetalleTesoreria(rs(0))
        If rs(2) <> Empty Then Call EliminaContab(rs(0), rs(1), rs(2))
        Call EliminaCargoAbono(rs(0), rs(1), rs(3), TxtTransf.Text)
        rs.MoveNext
     Loop
     VGCNx.CommitTrans
     MsgBox "La Transferencia Nº " & TxtTransf.Text & " fue Anulada Satisfactoriamente", vbInformation, Caption
  Else
     SQL = "Select cabrec_numrecibo from te_cabecerarecibos "
     SQL = SQL & "Where cabrec_numreciboegreso='" & TxtTransf.Text & "' "
     SQL = SQL & "and cabrec_estadoreg='1'"
     Set rs = VGCNx.Execute(SQL)
     If Not rs.EOF And Not rs.BOF Then
        MsgBox "El Nº de Transferencia " & TxtTransf.Text & " se encuentra Anulado", vbInformation, Caption
     Else
        MsgBox "No existe el Nº de Transferencia " & TxtTransf.Text
     End If
  End If
  cmdaceptar.Visible = False
  Set rs1 = Nothing
  DataGrid1.Refresh
  TxtTransf.Text = ""
  Exit Sub
X:
  MsgBox "Error en la Anulación de la Transferencia" & Chr(13) & Err.Number & " " & Err.Description, vbInformation, Caption
  VGCNx.RollbackTrans
  
End Sub

Sub AnulaRecIngresoEgreso(xValor As String)
  Dim SQL As String
 
  SQL = "Update te_cabecerarecibos set cabrec_estadoreg='1',usuariocodigo='" & VGParamSistem.Usuario & "',fechaact=getdate() where "
  SQL = SQL & "cabrec_numrecibo='" & xValor & "'"
  VGCNx.Execute (SQL)

End Sub

Sub EliminarDetalleTesoreria(xValor As String)
   Dim SQL As String
   SQL = "update te_detallerecibos set detrec_estadoreg='1',usuariocodigo='" & VGParamSistem.Usuario & "',fechaact=getdate() where "
   SQL = SQL & "cabrec_numrecibo='" & xValor & "'"
   VGCNx.Execute (SQL)

End Sub

Private Sub CmdCancelar_Click()
  Unload Me
End Sub

Private Sub Command1_Click()
  Dim SQL As String
  
  Set rs1 = New ADODB.Recordset
  SQL = "Select cabrec_numrecibo,empresacodigo,comprobconta,clientecodigo, cabrec_totsoles,cabrec_totdolares from te_cabecerarecibos "
  SQL = SQL & "Where cabrec_numreciboegreso<>'' and cabrec_numreciboegreso='" & TxtTransf.Text & "' "
  SQL = SQL & "and cabrec_estadoreg<>'1'"
  Set rs1 = VGCNx.Execute(SQL)
  If rs1.RecordCount = 1 Then
     MsgBox (" Numero de transferencia, solo tiene un Recibos de Ingreso o Egresos , Verifique ")
     Exit Sub
  ElseIf rs1.RecordCount = 0 Then
     MsgBox (" Numero de transferencia no existe o esta ANULADO , Verifique ")
     Exit Sub
  Else
     cmdaceptar.Visible = True
  End If
  Set DataGrid1.DataSource = rs1
  DataGrid1.Refresh
End Sub

Private Sub Form_Load()
cmdaceptar.Visible = False
End Sub

Private Sub TxtTransf_LostFocus()
  If TxtTransf.Text <> Empty Then
     TxtTransf.Text = Right(Format(TxtTransf.Text, "000000"), 6)
  End If
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
Private Sub EliminaCargoAbono(ByVal recibo As String, ByVal empresa As String, ByVal cliente As String, ByVal numtransf As String)
  Dim rsca As ADODB.Recordset
  Dim SQL As String
  
  Set rsca = New ADODB.Recordset
  Set rsca = VGCNx.Execute("Select Distinct empresacodigo From te_cabecerarecibos Where cabrec_numreciboegreso='" & numtransf & "'")
  If rsca.RecordCount > 1 Then
    Set rsca = Nothing
    Set rsca = VGCNx.Execute("Select detrec_tipodoc_concepto,detrec_numdocumento From te_detallerecibos Where cabrec_numrecibo='" & recibo & "'")
    If Not rsca.EOF Then
        If Left(recibo, 2) = "10" Then
            SQL = "Delete From cp_cargo Where empresacodigo='" & empresa & "'"
            SQL = SQL & " And clientecodigo='" & cliente & "'"
            SQL = SQL & " And documentocargo='" & Trim(rsca(0)) & "'"
            SQL = SQL & " And cargonumdoc='" & Trim(rsca(1)) & "'"
        Else
            SQL = "Delete From vt_cargo Where empresacodigo='" & empresa & "'"
            SQL = SQL & " And clientecodigo='" & cliente & "'"
            SQL = SQL & " And documentocargo='" & Trim(rsca(0)) & "'"
            SQL = SQL & " And cargonumdoc='" & Trim(rsca(1)) & "'"
        End If
    
        VGCNx.Execute (SQL)
    End If
  End If
End Sub
