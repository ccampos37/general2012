VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form Frmlogistica 
   Caption         =   "Reposición de Stock"
   ClientHeight    =   6240
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8805
   Icon            =   "frmlogistica.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   8805
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Height          =   696
      Left            =   36
      TabIndex        =   31
      Top             =   36
      Width           =   8688
      Begin VB.OptionButton Option2 
         Caption         =   "Por Descripción"
         Height          =   228
         Left            =   6876
         TabIndex        =   35
         Top             =   252
         Value           =   -1  'True
         Width           =   1596
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Por Codigo"
         Height          =   228
         Left            =   5256
         TabIndex        =   34
         Top             =   252
         Width           =   1272
      End
      Begin VB.TextBox Text1 
         Height          =   288
         Left            =   1224
         TabIndex        =   33
         Top             =   216
         Width           =   3540
      End
      Begin VB.Label Label2 
         Caption         =   "Buscar"
         Height          =   228
         Left            =   252
         TabIndex        =   32
         Top             =   252
         Width           =   1524
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1095
      Left            =   36
      TabIndex        =   26
      Top             =   5148
      Width           =   8652
      Begin VB.CommandButton Command3 
         Caption         =   "&Consultar"
         Height          =   675
         Left            =   1800
         Picture         =   "frmlogistica.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   240
         Width           =   775
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Modificar"
         Height          =   675
         Left            =   600
         Picture         =   "frmlogistica.frx":0D0C
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   240
         Width           =   775
      End
      Begin VB.CommandButton Command8 
         Caption         =   "&Grabar"
         Height          =   675
         Left            =   3000
         Picture         =   "frmlogistica.frx":114E
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   240
         Visible         =   0   'False
         Width           =   775
      End
      Begin VB.CommandButton Command19 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         Height          =   675
         Left            =   7596
         Picture         =   "frmlogistica.frx":1590
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   252
         Width           =   775
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   4395
      Left            =   60
      TabIndex        =   12
      Top             =   750
      Width           =   8676
      Begin VB.CommandButton Command1 
         Caption         =   "Ubicación"
         Height          =   336
         Left            =   7020
         TabIndex        =   36
         Top             =   3924
         Width           =   1380
      End
      Begin VB.TextBox TxTiemRep 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2400
         TabIndex        =   10
         Top             =   3960
         Width           =   1695
      End
      Begin VB.TextBox TxTipCom 
         Height          =   285
         Left            =   2400
         MaxLength       =   3
         TabIndex        =   5
         Top             =   2160
         Width           =   495
      End
      Begin VB.TextBox TxTipRep 
         Height          =   285
         Left            =   2400
         MaxLength       =   1
         TabIndex        =   4
         Top             =   1800
         Width           =   375
      End
      Begin VB.TextBox TxCodInt 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   2400
         MaxLength       =   20
         TabIndex        =   0
         Top             =   360
         Width           =   1815
      End
      Begin VB.TextBox TxCodFab 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   2400
         MaxLength       =   40
         TabIndex        =   1
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox TxUnidad 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2400
         MaxLength       =   6
         TabIndex        =   3
         Top             =   1440
         Width           =   855
      End
      Begin VB.TextBox TxActual 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   2400
         TabIndex        =   6
         Top             =   2520
         Width           =   1695
      End
      Begin VB.TextBox TxStkMin 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2400
         TabIndex        =   7
         Top             =   2880
         Width           =   1695
      End
      Begin VB.TextBox TxStkMax 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2400
         TabIndex        =   8
         Top             =   3240
         Width           =   1695
      End
      Begin VB.TextBox TxPedido 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2400
         MaxLength       =   12
         TabIndex        =   9
         Top             =   3600
         Width           =   1695
      End
      Begin VB.TextBox TxDescri 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2400
         TabIndex        =   2
         Top             =   1080
         Width           =   4455
      End
      Begin VB.Label Label12 
         Caption         =   "(Clasificación ABC)"
         Height          =   255
         Left            =   3000
         TabIndex        =   25
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label Label9 
         Caption         =   "Tipo de Compra"
         Height          =   255
         Left            =   360
         TabIndex        =   24
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Clasificacion"
         Height          =   255
         Left            =   360
         TabIndex        =   23
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Codigo Interno"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         TabIndex        =   22
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "Stock Minimo"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         TabIndex        =   21
         Top             =   2880
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "Unidad Medida"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         TabIndex        =   20
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label Label10 
         Caption         =   "Codigo Fab."
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         TabIndex        =   19
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label lINE 
         Caption         =   "Stock Maximo"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         TabIndex        =   18
         Top             =   3240
         Width           =   1455
      End
      Begin VB.Label Label7 
         Caption         =   "Stock Actual"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         TabIndex        =   17
         Top             =   2520
         Width           =   1455
      End
      Begin VB.Label Label8 
         Caption         =   "Punto Pedido"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         TabIndex        =   16
         Top             =   3600
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "Tiempo Reposicion"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         TabIndex        =   15
         Top             =   3960
         Width           =   1455
      End
      Begin VB.Label Label11 
         Caption         =   "Descripcion"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         TabIndex        =   14
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label lblunidad 
         Caption         =   "lblunidad"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   3360
         TabIndex        =   13
         Top             =   1440
         Width           =   4800
      End
   End
   Begin MSFlexGridLib.MSFlexGrid FG 
      Height          =   3708
      Left            =   960
      TabIndex        =   11
      Top             =   972
      Visible         =   0   'False
      Width           =   6012
      _ExtentX        =   10610
      _ExtentY        =   6535
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid DBGrid1 
      Height          =   4245
      Left            =   60
      TabIndex        =   37
      Top             =   840
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   7488
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
Attribute VB_Name = "Frmlogistica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit
 Dim Adoreg1 As ADODB.Recordset
 Dim rs As New ADODB.Recordset
 Dim Cod As String
 Dim nsql As String
 Dim consulta As Boolean

Private Sub limpia()
   TxCodInt = ""
   TxCodFab = ""
   TxDescri = ""
   TxTipRep = ""
   TxTipCom = ""
   TxStkMin = ""
   TxStkMax = ""
   TxActual = ""
   TxTiemRep = ""
   TxPedido = ""
   lblunidad = ""
End Sub

Private Sub Command1_Click()
 FormCasillero.grid.Visible = False
 'FormCasillero.Text1.SetFocus
 
 FormCasillero.limpia
 FormCasillero.Command4.Enabled = False
 FormCasillero.command5.Enabled = False
 
 FormCasillero.Command3.Visible = False
 FormCasillero.Command2.Enabled = True
 FormCasillero.cmdsalirlogistica.Visible = True
  
 FormCasillero.Command1.Enabled = True
 
 FormCasillero.Text1 = TxCodInt
 FormCasillero.Label3 = TxDescri
 FormCasillero.agregarlista
 FormCasillero.Show 1
End Sub

'Salir
Private Sub Command19_Click()
 If Frame1.Visible Then
   habilitado (True)
   bloqueado (True)
   limpia
   Frame1.Visible = False
   Command8.Enabled = False
 Else
   Unload Me
 End If
End Sub
'Modificar
Private Sub Command2_Click()
Dim citerio As String

If rs.RecordCount > 0 Then
  Frame1.Visible = True
  habilitado (False)
  Command8.Enabled = True
  TxCodInt = rs.Fields("ACODIGO")
  TxDescri = cNull(rs.Fields("ADESCRI"))
  TxUnidad = cNull(rs.Fields("Aunidad"))
  consulta = False
  llenastk
  TxTipRep.SetFocus
  Command1.Enabled = True
End If
End Sub
'Consultar
Private Sub Command3_Click()
If rs.RecordCount > 0 Then
  consulta = True
  Frame1.Visible = True
  habilitado (False)
  Command8.Enabled = False
  TxCodInt = rs.Fields("ACODIGO")
  TxDescri = cNull(rs.Fields("ADESCRI"))
  TxUnidad = cNull(rs.Fields("Aunidad"))
  llenastk
  bloqueado (False)
  Command1.Enabled = False
End If
End Sub
'Grabar
Private Sub Command8_Click()
   Dim csql As String
   Dim ngrabo As Long
 
   If Not Frame1.Visible Then Exit Sub
   
   If Trim(TxDescri = "") Then
      MsgBox "La descripcion No es Valida.....! Registre la Descripción del Articulo"
      Exit Sub
   End If
   
   If Not IsNumeric(TxStkMin) Or Trim(TxStkMin) = "" Then
          MsgBox "Ingrese el stock minimo", vbExclamation, "Aviso"
          TxStkMin.SetFocus
          Exit Sub
   End If
   If Not IsNumeric(TxStkMax) Then
        MsgBox "Ingrese el stock máximo", vbExclamation, "Aviso"
        TxStkMax.SetFocus
        Exit Sub
   End If
   TxStkMin = Val(TxStkMin)
   TxStkMax = Val(TxStkMax)
   TxPedido = Val(TxPedido)
   TxTiemRep = Val(TxTiemRep)
   
   csql = "Update StkArt set STTIPREP = '" & IIf(Trim(TxTipRep) = "", " ", TxTipRep) & "',"
   csql = csql & "STSEMREP = " & TxTiemRep & ","
   If Trim(TxTipCom) <> "" Then
        csql = csql & "STTIPCOM = '" & TxTipCom & "',"
   End If
   csql = csql & "STSKMIN = " & TxStkMin & ",STSKMAX = " & TxStkMax & " ,"
   csql = csql & "STPUNREP = " & TxPedido & " Where STCODIGO = '" & TxCodInt & "' and  STALMA = '" & VGAlma & "' "
   VGCNx.Execute csql, ngrabo
   If ngrabo = 0 Then
        If Trim(TxTipRep) = "" Then TxTipRep = "  "
        If Trim(TxTipCom) = "" Then TxTipCom = "  "
        'MsgBox "No se grabo el registro", vbInformation, "Aviso"
        
        csql = "insert into stkart (stcodigo,stalma,sttiprep,stsemrep,sttipcom,stskmin,stskmax,stpunrep) values"
        csql = csql & "('" & TxCodInt & "','" & VGAlma & "' ,'" & TxTipRep & "'," & TxTiemRep & ",'" & TxTipCom & "'," & TxStkMin & "," & TxStkMax & "," & TxPedido & ")"
        VGCNx.Execute csql, ngrabo
        If ngrabo = 0 Then
                 MsgBox "error en momento de grabar"
        Else
                MsgBox "Se grabó correctamente", vbInformation, "Aviso"
        End If
   Else
        MsgBox "Se grabó correctamente", vbInformation, "Aviso"
   End If
   Call Listado("SELECT m.ACODIGO,m.ADESCRI ,m.AUNIDAD FROM MaeArt m  order by m.ADESCRI")
   
   limpia
   habilitado (True)
   Frame1.Visible = False
   Command8.Enabled = False
End Sub

Private Sub Form_Load()
  Dim rsql As String
  
  On Error Resume Next

  limpia
  lblunidad = ""
  central Frmlogistica
  Frame1.Visible = False
  Command8.Visible = True
  Command8.Enabled = False
  nsql = "SELECT m.ACODIGO,m.ADESCRI ,m.AUNIDAD FROM MaeArt m  order by m.ADESCRI"
  
  Call Listado(nsql)
End Sub
Sub Listado(wcad)
  Set DBGrid1.DataSource = Nothing
  Set rs = Nothing
  
  Set rs = VGCNx.Execute(wcad)
  Set DBGrid1.DataSource = rs
  With DBGrid1
      .Columns(0).Caption = "Codigo"
      .Columns(0).Width = 2000
      .Columns(1).Caption = "Descripcion"
      .Columns(1).Width = 5000
      .Columns(2).Caption = "Unidad"
      .Columns(2).Width = 800
      .MarqueeStyle = dbgHighlightRow
      .Refresh
  End With
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
  Dim SQL As String
  If Option2.Value = True Then
     SQL = "SELECT m.ACODIGO,m.ADESCRI ,m.AUNIDAD FROM MaeArt m  where m.ADESCRI like '" & Text1 & "%' order by m.ADESCRI"
     'data2.Recordset.FindFirst (" Adescri like '" & Text1.text & "%'")
  Else
     'data2.Recordset.FindFirst (" ACODIGO like '" & Text1.text & "%'")
     SQL = "SELECT m.ACODIGO,m.ADESCRI ,m.AUNIDAD FROM MaeArt m  where m.ACODIGO like '" & Text1 & "%' order by m.Acodigo"
  End If
      
  Call Listado(SQL)
   
  '   .RecordSource = SQL
  '   Data2.Refresh

End Sub

Private Sub TxActual_GotFocus()
Enfoque TxActual
End Sub

Private Sub TxActual_KeyPress(KeyAscii As Integer)
If NumSpto(KeyAscii) Then
    If KeyAscii = 13 Then SendKeys "{tab}"
End If
End Sub

Private Sub TxCodFab_GotFocus()
Enfoque TxCodFab
End Sub

Private Sub TxCodFab_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
      SendKeys "{tab}"
      KeyAscii = 0
  End If
End Sub

Private Sub TxCodInt_GotFocus()
Enfoque TxCodInt
End Sub

Private Sub TxCodInt_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
      SendKeys "{tab}"
      KeyAscii = 0
  End If
End Sub

Private Sub TxDescri_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
      SendKeys "{tab}"
      KeyAscii = 0
  End If
End Sub

Private Sub TxPedido_GotFocus()
Enfoque TxPedido
End Sub

Private Sub TxPedido_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    If IsNumeric(TxStkMin) Then
       TxTiemRep.SetFocus
    Else
       MsgBox "Ingrese un dato númerico", vbExclamation, "Error"
    End If
  End If
End Sub

Private Sub TxStkMax_GotFocus()
Enfoque TxStkMax
End Sub

Private Sub TxStkMax_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 And IsNumeric(TxStkMax) Then
    TxPedido.SetFocus
  End If
End Sub

Private Sub TxStkMin_GotFocus()
Enfoque TxStkMin
End Sub

Private Sub TxStkMin_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    If IsNumeric(TxStkMin) Then
       TxStkMax.SetFocus
    Else
       MsgBox "Ingrese un dato numerico", vbExclamation, "Error"
    End If
  End If
End Sub

Private Sub TxTiemRep_GotFocus()
Enfoque TxTiemRep
End Sub

Private Sub TxTiemRep_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Not IsNumeric(TxTiemRep) Then
        MsgBox "Ingrese dato númerico", vbInformation, "Información"
        TxTiemRep.SetFocus
    Else
        Command8.SetFocus
    End If
End If
End Sub

Private Sub TxTipCom_DblClick()
  FrmAyu01.Caption = "Tipo de Compra"
  FrmAyu01.cCod = "35"
  FrmAyu01.Show 1
  TxTipCom = FrmAyu01.cC
End Sub

'Private Sub muestra()
'      If IsNull(rs("STTIPREP")) Then
'          TxEstado = "NORMAL"
'      ElseIf rs("STSKDIS") <= rs("STSKMIN") Then
'          TxEstado = "STOCK CRITICO"
'      ElseIf rs("STSKDIS") <= rs("STPUNREP") Then
'          TxEstado = "STOCK REPOSICION"
'      Else: rs("STSKDIS") = 0
'          TxEstado = "STOCK EN CERO (0)"
'      End If
'
'End Sub

Private Sub habilitado(flag As Boolean)
  Command2.Enabled = flag
  Command3.Enabled = flag
End Sub

Private Sub TxTipCom_GotFocus()
Enfoque TxTipCom
End Sub

Private Sub TxTipCom_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then TxTipCom_DblClick
End Sub

Private Sub TxTipCom_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    TxStkMin.SetFocus
  End If
End Sub

Private Sub TxTipRep_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If Trim(TxTipRep) <> "" Then
   TxTipRep = UCase(TxTipRep)
   TxTipCom.SetFocus
  Else
    MsgBox "Ingrese el tipo de Clasificación", vbExclamation, "Aviso"
  End If
End If
End Sub

Private Sub llenastk()
Dim rsql As String

   rsql = "select  stskdis, stskmin,stskmax,stpunrep,sttiprep,sttipcom,STSEMREP  from stkart  WHERE STALMA='" & VGAlma & "'  and stcodigo ='" & TxCodInt & "'"
   Set Adoreg1 = New ADODB.Recordset
   Adoreg1.Open rsql, VGCNx, adOpenStatic
   If Adoreg1.EOF Then
     Exit Sub
   End If
   'If consulta Then
     TxStkMin = IIf(IsNull(Adoreg1(1)), 0, Adoreg1(1))
     TxStkMax = IIf(IsNull(Adoreg1(2)), 0, Adoreg1(2))
     TxPedido = IIf(IsNull(Adoreg1(3)), 0, Adoreg1(3))
     TxTipRep = IIf(IsNull(Adoreg1(4)), "", Adoreg1(4))
     TxTipCom = IIf(IsNull(Adoreg1(5)), "", Adoreg1(5))
     consulta = False
     TxActual = IIf(IsNull(Adoreg1(0)), 0, Adoreg1(0))
     TxTiemRep = "" & IIf(IsNull(Adoreg1(0)), 0, Adoreg1("STSEMREP"))
End Sub

Private Sub bloqueado(flag As Boolean)
'   TxCodInt.Enabled = flag
'   TxCodFab.Enabled = flag
'   TxDescri.Enabled = flag
   TxTipRep.Enabled = flag
   TxTipCom.Enabled = flag
   TxStkMin.Enabled = flag
   TxStkMax.Enabled = flag
'   TxActual.Enabled = flag
   TxTiemRep.Enabled = flag
   TxPedido.Enabled = flag
End Sub

Private Sub Txunidad_DblClick()
Dim Adodc2 As ADODB.Recordset
Set Adodc2 = New ADODB.Recordset

Adodc2.Open "SELECT UM_ABREV,UM_NOMBRE FROM TABUNIMED", VGCNx, adOpenStatic, adLockOptimistic
frmReferencia.Conectar Adodc2, "SELECT UM_ABREV,UM_NOMBRE FROM TABUNIMED"
frmReferencia.Label1.Caption = "Unidades de Medida"
frmReferencia.Show vbModal
Adodc2.Close
If vGUtil(1) <> "" Then
  TxUnidad = vGUtil(1)
  lblunidad = vGUtil(2)
End If

End Sub

Private Sub TxUnidad_GotFocus()
Enfoque TxUnidad
End Sub

Private Sub TxUnidad_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then Txunidad_DblClick
End Sub

Private Sub TxUnidad_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Existe(1, TxUnidad, "TabUNimed", "Um_Abrev", False) Then
        SendKeys "{tab}"
    Else
        MsgBox "La no existe esta Unidad", vbInformation, "Información"
        TxUnidad.SetFocus
    End If
End If
End Sub
