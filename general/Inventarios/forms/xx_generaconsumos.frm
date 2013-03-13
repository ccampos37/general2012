VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Begin VB.Form xx_generaconsumos 
   Caption         =   "Form1"
   ClientHeight    =   7080
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9810
   DrawMode        =   14  'Copy Pen
   DrawStyle       =   1  'Dash
   LinkTopic       =   "Form1"
   ScaleHeight     =   7080
   ScaleWidth      =   9810
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   120
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   6120
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Procesar"
      Height          =   495
      Left            =   1920
      TabIndex        =   3
      Top             =   6360
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Salir"
      Height          =   495
      Left            =   4440
      TabIndex        =   2
      Top             =   6360
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Registros"
      Height          =   5535
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   9375
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   4815
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   8493
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
   End
   Begin VB.Label Label2 
      Height          =   495
      Left            =   8280
      TabIndex        =   5
      Top             =   6120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Nro registros"
      Height          =   495
      Left            =   6840
      TabIndex        =   4
      Top             =   6120
      Width           =   1215
   End
End
Attribute VB_Name = "xx_generaconsumos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RSQL As New ADODB.Recordset
Public RSQL1 As New ADODB.Recordset
Public RSQL2 As New ADODB.Recordset
Public contador As Integer

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
generareceta
End Sub

Private Sub Form_Load()
SQL = "select distinct caalma,catd, canumdoc,cacodmov, cafecdoc, mesproceso, decodigo, decantid "
SQL = SQL & " from v_kardex aa inner join xx_kits_wilmer bb  on aa.DECODIGO=codkit"
SQL = SQL & " where almacenvalorizado=1 and empresacodigo='02'   and puntovtacodigo='03'"
SQL = SQL & " and CACODMOV  in ( '13','37','38','39') and mesproceso<='201210' order by mesproceso , cafecdoc, caalma, canumdoc"
Set RSQL = VGCNx.Execute(SQL)
Set DataGrid1.DataSource = RSQL
DataGrid1.Refresh
Label2.Caption = RSQL.RecordCount

End Sub
Private Sub generareceta()
RSQL.MoveFirst
Do While Not RSQL.EOF
   SQL = "select *, dif=qq-consumomes from ( Select codkit,codart, "
   SQL = SQL & " qq=canart*" & RSQL!decantid & ", consumomes=isnull(SUM(case when catipmov='S' then decantid else 0 end),0),saldo= isnull(saldo,0) "
   SQL = SQL & " from xx_kits_wilmer a , v_kardex b,"
   SQL = SQL & " ( select decodigo, saldo=isnull(SUM(case when catipmov='I' then decantid else decantid * -1 end),0) "
   SQL = SQL & " from v_kardex where mesproceso<='" & RSQL!mesproceso & "' and DEALMA='31' group by decodigo ) c "
   SQL = SQL & " where a.codart*=b.decodigo and a.codart*=c.decodigo   "
   SQL = SQL & " and b.mesproceso='" & RSQL!mesproceso & "' and DEALMA='31' and codkit='" & RSQL!decodigo & "'and cacodmov='92' "
   SQL = SQL & " group by codkit,codart, canart, saldo ) x where qq> consumomes and (qq-consumomes) <=saldo "
   Set RSQL2 = VGCNx.Execute(SQL)
   If RSQL2.RecordCount > 0 Then
      generamovimientos
   End If
   RSQL.MoveNext
Loop
End Sub
Private Sub generamovimientos()
Call grabacorrelativo
Call grabacabecera
contador = 1
RSQL2.MoveFirst
Do While Not RSQL2.EOF
   grabadetalle
   contador = contador + 1
   RSQL2.MoveNext
Loop
End Sub
Private Sub grabacabecera()
  Dim criterio As String
  Dim uusql As New ADODB.Recordset
  Dim CADENA As String
  Dim FACTOR As Double
  Dim usql As String
  Dim fechaa As Date
  Dim Data1 As New ADODB.Recordset
  Dim acmd As New ADODB.Command
  VGCNx.BeginTrans
  fechaa = Fecha(2, "01/" & Right(RSQL!mesproceso, 2) & "/" & Left(RSQL!mesproceso, 4) & "")
  Set Data1 = Nothing
        Set acmd.ActiveConnection = VGgeneral
        acmd.CommandType = adCmdStoredProc
        acmd.CommandText = "al_ingresoalma_pro"
        acmd.CommandTimeout = 0
        acmd.Prepared = True
        With acmd
            .Parameters("@base") = VGCNx.DefaultDatabase
            .Parameters("@tabla") = "movalmcab"
            .Parameters("@tipo") = "2"
            .Parameters("@numero") = Mid$(UCase$(Text4.text), 1, 11)
            .Parameters("@modoventa") = "92"
            .Parameters("@moneda") = "01"
            .Parameters("@Nroreq") = RSQL!decodigo
            .Parameters("@fecha") = fechaa
            .Parameters("@fechafactura") = fechaa
            .Parameters("@almacen") = "31"
            .Parameters("@usuario") = UCase(VGUsuario)
            .Parameters("@fechaactual") = Now
            .Parameters("@empresa") = VGparametros.empresacodigo
        End With
        acmd.Execute
        Set acmd = Nothing
        DoEvents
   VGCNx.CommitTrans
   Exit Sub
GrabErr:
       MsgBox Err.Description
       VGCNx.RollbackTrans
       Exit Sub
       Resume
End Sub

Private Sub grabadetalle()
Dim Data2 As New ADODB.Recordset
 Dim acmd As New ADODB.Command
VGCNx.BeginTrans
SQL = "select centrocostocodigo from zz_transaccion_centrocosto where cacodmov ='" & RSQL!cacodmov & "'"
Set uusql = VGCNx.Execute(SQL)
On Error GoTo Err1
        Set acmd.ActiveConnection = VGgeneral
        acmd.CommandType = adCmdStoredProc
        acmd.CommandText = "vt_ingresodetallealma_pro"
        acmd.CommandTimeout = 0
        acmd.Prepared = True
        With acmd
            .Parameters("@base") = VGCNx.DefaultDatabase
            .Parameters("@tabla") = "movalmdet"
            .Parameters("@tipo") = "2"
            .Parameters("@item") = contador
            .Parameters("@numero") = Mid$(UCase$(Text4.text), 1, 11)
            .Parameters("@producto") = RSQL2!codart
            If RSQL2!dif > 1000 Then
              MsgBox ("ok")
            End If
            .Parameters("@cantidad") = RSQL2!dif
            .Parameters("@almacen") = "31"
            .Parameters("@CENCOS") = ESNULO(uusql!centrocostocodigo, "")
        End With
        acmd.Execute
        Set acmd = Nothing
        DoEvents
   VGCNx.CommitTrans
   Exit Sub
Err1:
 '      MsgBox Err.Description
       VGCNx.RollbackTrans
       Exit Sub
       Resume
End Sub

Private Sub grabacorrelativo()
Dim rs As New ADODB.Recordset
VGCNx.BeginTrans

   rs.Open "select  TANUMENT, TANUMSAL from TabAlm  WHERE TAALMA='31'", VGCNx, adOpenDynamic, adLockOptimistic

      Text4.text = Format(rs("tanumsal"), "00000000000")
      rs("tanumsal") = rs("tanumsal") + 1
   rs.Update
   rs.Close
VGCNx.CommitTrans
End Sub

