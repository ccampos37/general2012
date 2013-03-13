VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmordencompra 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Emisión de Orden de Compra"
   ClientHeight    =   4656
   ClientLeft      =   1128
   ClientTop       =   2832
   ClientWidth     =   9384
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmordencompra.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4656
   ScaleWidth      =   9384
   Begin VB.Frame fraDatos 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2520
      Left            =   120
      TabIndex        =   23
      Top             =   600
      Visible         =   0   'False
      Width           =   9150
      Begin VB.TextBox txtNSol 
         Height          =   285
         Left            =   8040
         MaxLength       =   10
         TabIndex        =   55
         Top             =   210
         Width           =   945
      End
      Begin VB.TextBox txtObs 
         Height          =   285
         Left            =   1455
         TabIndex        =   10
         Top             =   2040
         Width           =   7500
      End
      Begin VB.TextBox txtCot 
         Height          =   285
         Left            =   5760
         TabIndex        =   7
         Top             =   960
         Width           =   3210
      End
      Begin VB.TextBox txtTip 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   7950
         MultiLine       =   -1  'True
         TabIndex        =   5
         Top             =   600
         Width           =   1020
      End
      Begin VB.TextBox txtEst 
         Height          =   285
         Left            =   7320
         MaxLength       =   3
         TabIndex        =   4
         Top             =   600
         Width           =   585
      End
      Begin VB.TextBox txtMon 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Left            =   5760
         MaxLength       =   2
         TabIndex        =   3
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox txtPro 
         Height          =   285
         Left            =   1440
         MaxLength       =   11
         TabIndex        =   0
         Top             =   240
         Width           =   1185
      End
      Begin VB.TextBox txtFor 
         Height          =   285
         Left            =   1440
         MaxLength       =   50
         TabIndex        =   6
         Top             =   960
         Width           =   3015
      End
      Begin VB.TextBox txtEntE 
         Height          =   285
         Left            =   1455
         MaxLength       =   50
         TabIndex        =   8
         Top             =   1320
         Width           =   5295
      End
      Begin VB.TextBox txtSol 
         Height          =   285
         Left            =   1455
         MaxLength       =   3
         TabIndex        =   9
         Top             =   1680
         Width           =   435
      End
      Begin MSComCtl2.DTPicker txtEmi 
         Height          =   285
         Left            =   1440
         TabIndex        =   1
         Top             =   600
         Width           =   1215
         _ExtentX        =   2138
         _ExtentY        =   508
         _Version        =   393216
         Format          =   60096513
         CurrentDate     =   37015
      End
      Begin MSComCtl2.DTPicker txtEnt 
         Height          =   285
         Left            =   3600
         TabIndex        =   2
         Top             =   600
         Width           =   1215
         _ExtentX        =   2138
         _ExtentY        =   508
         _Version        =   393216
         Format          =   60096513
         CurrentDate     =   37015
      End
      Begin VB.Label Label11 
         Caption         =   "N° Sol C."
         Height          =   255
         Left            =   7395
         TabIndex        =   54
         Top             =   285
         Width           =   870
      End
      Begin VB.Label Label12 
         Caption         =   "Observación :"
         Height          =   255
         Left            =   375
         TabIndex        =   38
         Top             =   2055
         Width           =   1095
      End
      Begin VB.Label lblSol 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2040
         TabIndex        =   37
         Top             =   1680
         Width           =   3135
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "T.Cambio  :"
         Height          =   195
         Left            =   6360
         TabIndex        =   36
         Top             =   615
         Width           =   810
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Moneda  :"
         Height          =   195
         Left            =   4920
         TabIndex        =   35
         Top             =   615
         Width           =   720
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Entrega   :"
         Height          =   195
         Left            =   2760
         TabIndex        =   34
         Top             =   615
         Width           =   720
      End
      Begin VB.Label lblRuc 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   6075
         TabIndex        =   32
         Top             =   225
         Width           =   1260
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "R.U.C.  :"
         Height          =   195
         Left            =   5580
         TabIndex        =   31
         Top             =   285
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Proveedor     :"
         Height          =   195
         Left            =   375
         TabIndex        =   30
         Top             =   255
         Width           =   1005
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Emisión         :"
         Height          =   195
         Left            =   375
         TabIndex        =   29
         Top             =   615
         Width           =   990
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Forma Pago  :"
         Height          =   195
         Left            =   375
         TabIndex        =   28
         Top             =   975
         Width           =   990
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Entregar en   :"
         Height          =   195
         Left            =   375
         TabIndex        =   27
         Top             =   1335
         Width           =   1005
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Solicitante     :"
         Height          =   195
         Left            =   375
         TabIndex        =   26
         Top             =   1695
         Width           =   1005
      End
      Begin VB.Label lblPro 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2625
         TabIndex        =   25
         Top             =   240
         Width           =   2910
      End
      Begin VB.Label lblCen 
         AutoSize        =   -1  'True
         Caption         =   "Cotización  :"
         Height          =   195
         Left            =   4710
         TabIndex        =   24
         Top             =   975
         Width           =   870
      End
   End
   Begin VB.Data Data2 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4440
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton cmdNue2 
      Caption         =   "&Agregar"
      Height          =   675
      Left            =   1575
      Picture         =   "frmordencompra.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5730
      Visible         =   0   'False
      Width           =   775
   End
   Begin VB.CommandButton cmdEli2 
      Caption         =   "&Quitar"
      Enabled         =   0   'False
      Height          =   675
      Left            =   4230
      Picture         =   "frmordencompra.frx":0D0C
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5730
      Visible         =   0   'False
      Width           =   775
   End
   Begin VB.CommandButton cmdEdi2 
      Caption         =   "&Editar"
      Enabled         =   0   'False
      Height          =   675
      Left            =   2895
      Picture         =   "frmordencompra.frx":114E
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5730
      Visible         =   0   'False
      Width           =   775
   End
   Begin VB.CommandButton cmdImp 
      Caption         =   "&Imprimir"
      Height          =   675
      Left            =   5535
      Picture         =   "frmordencompra.frx":1590
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   3840
      Visible         =   0   'False
      Width           =   775
   End
   Begin VB.CommandButton cmdEdi 
      Caption         =   "&Editar"
      Height          =   675
      Left            =   2910
      Picture         =   "frmordencompra.frx":19D2
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   3825
      Width           =   775
   End
   Begin VB.CommandButton CmdEli 
      Caption         =   "&Anular"
      Height          =   675
      Left            =   4230
      Picture         =   "frmordencompra.frx":1E14
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   3840
      Width           =   775
   End
   Begin VB.CommandButton cmdNue 
      Caption         =   "&Nuevo"
      Height          =   675
      Left            =   1575
      Picture         =   "frmordencompra.frx":2256
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   3810
      Width           =   775
   End
   Begin VB.CommandButton CmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   675
      Left            =   6840
      Picture         =   "frmordencompra.frx":2698
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   3825
      Width           =   775
   End
   Begin VB.CommandButton cmdGra 
      Caption         =   "&Grabar"
      Height          =   675
      Left            =   5520
      Picture         =   "frmordencompra.frx":2ADA
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5730
      Visible         =   0   'False
      Width           =   775
   End
   Begin VB.CommandButton CmdSalir2 
      Caption         =   "&Salir"
      Height          =   675
      Left            =   6840
      Picture         =   "frmordencompra.frx":2F1C
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   5730
      Visible         =   0   'False
      Width           =   775
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Flex1 
      Height          =   1515
      Left            =   120
      TabIndex        =   22
      Top             =   3120
      Visible         =   0   'False
      Width           =   9150
      _ExtentX        =   16150
      _ExtentY        =   2667
      _Version        =   393216
      Cols            =   15
      FixedCols       =   0
      RowHeightMin    =   240
      BackColorSel    =   0
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      ScrollBars      =   2
      SelectionMode   =   1
      Appearance      =   0
      FormatString    =   "^Código|Fab|Descripción|xUni|xCantidad|Uni.|Cantidad|PU|>Precio|>%Des|Igv|>Total|C1|C2"
      RowSizingMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   15
   End
   Begin VB.Frame fraCabec 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   135
      TabIndex        =   33
      Top             =   0
      Visible         =   0   'False
      Width           =   9120
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Estado  :"
         Height          =   195
         Left            =   6630
         TabIndex        =   42
         Top             =   240
         Width           =   630
      End
      Begin VB.Label lblEst 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   7380
         TabIndex        =   41
         Top             =   150
         Width           =   1590
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Número  :"
         Height          =   195
         Left            =   375
         TabIndex        =   40
         Top             =   240
         Width           =   690
      End
      Begin VB.Label lblNum 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1080
         TabIndex        =   39
         Top             =   160
         Width           =   1560
      End
   End
   Begin VB.Frame fraTotales 
      Height          =   975
      Left            =   135
      TabIndex        =   43
      Top             =   4560
      Visible         =   0   'False
      Width           =   8985
      Begin VB.Label lblCom 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "892,760.00"
         Height          =   285
         Left            =   7080
         TabIndex        =   53
         Top             =   600
         Width           =   1110
      End
      Begin VB.Label lblIgv 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "892,760.00"
         Height          =   285
         Left            =   7080
         TabIndex        =   52
         Top             =   240
         Width           =   1110
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Compra :"
         Height          =   195
         Left            =   6360
         TabIndex        =   51
         Top             =   600
         Width           =   630
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "I.G.V.   :"
         Height          =   195
         Left            =   6360
         TabIndex        =   50
         Top             =   240
         Width           =   585
      End
      Begin VB.Label lblTot 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "892,760.00"
         Height          =   285
         Left            =   4200
         TabIndex        =   49
         Top             =   240
         Width           =   1230
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Total  :"
         Height          =   195
         Left            =   3600
         TabIndex        =   48
         Top             =   240
         Width           =   495
      End
      Begin VB.Label lblDes 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "892,760.00"
         Height          =   285
         Left            =   1680
         TabIndex        =   47
         Top             =   600
         Width           =   1110
      End
      Begin VB.Label lblImp 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "892,760.00"
         Height          =   285
         Left            =   1680
         TabIndex        =   46
         Top             =   240
         Width           =   1110
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Descuento :"
         Height          =   195
         Left            =   720
         TabIndex        =   45
         Top             =   600
         Width           =   870
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Importe      :"
         Height          =   195
         Left            =   720
         TabIndex        =   44
         Top             =   240
         Width           =   840
      End
   End
   Begin Crystal.CrystalReport CrystalReport2 
      Bindings        =   "frmordencompra.frx":335E
      Left            =   0
      Top             =   3960
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2340
      Left            =   120
      TabIndex        =   16
      Top             =   600
      Width           =   8955
      _ExtentX        =   15790
      _ExtentY        =   4128
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   7
      BeginProperty Column00 
         DataField       =   "OC_CNUMORD"
         Caption         =   "        Número"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "OC_CRAZSOC"
         Caption         =   "                   Desc. Proveedor"
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
      BeginProperty Column02 
         DataField       =   "OC_DFECDOC"
         Caption         =   "    Emisión"
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
      BeginProperty Column03 
         DataField       =   "OC_CCODMON"
         Caption         =   "Mo"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "#.##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "OC_NVENTA"
         Caption         =   "     Importe"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "EST_NOMBRE"
         Caption         =   "      Estado"
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
      BeginProperty Column06 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   2
         ScrollBars      =   2
         AllowRowSizing  =   0   'False
         Size            =   273
         BeginProperty Column00 
            Alignment       =   2
            Locked          =   -1  'True
            WrapText        =   -1  'True
            ColumnWidth     =   1368
         EndProperty
         BeginProperty Column01 
            Locked          =   -1  'True
            WrapText        =   -1  'True
            ColumnWidth     =   3107.906
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
            Locked          =   -1  'True
            WrapText        =   -1  'True
            ColumnWidth     =   1116.284
         EndProperty
         BeginProperty Column03 
            Alignment       =   2
            Locked          =   -1  'True
            WrapText        =   -1  'True
            ColumnWidth     =   432
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            Locked          =   -1  'True
            WrapText        =   -1  'True
            ColumnWidth     =   1019.906
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1379.906
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   11.906
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmordencompra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Colex As New Collection
Dim Adodc1 As ADODB.Recordset
Dim cSql1 As String
Dim nT As Integer       'Ingreso,Modificación,Ficha Tecnica
Dim cCod As String
Dim nTra As Integer
Dim Mensaje As String
Dim unum As String

Sub OculObj02(nTipo As Boolean)
    cmdGra.Visible = nTipo
    CmdSalir2.Visible = nTipo
End Sub

Sub OculObj03(nTipo As Boolean)
    fraCabec.Visible = nTipo
    fraDatos.Visible = nTipo
    fraTotales.Visible = nTipo
End Sub

Sub OculObj04(nTipo As Boolean)
    cmdNue.Visible = nTipo
    cmdEdi.Visible = nTipo
    CmdEli.Visible = nTipo
    cmdImp.Visible = nTipo
    CmdSalir.Visible = nTipo
End Sub

Sub OculObj06(nTipo As Boolean)
    DataGrid1.Visible = nTipo
End Sub

Sub Abre_Tabla_OCs()
    Dim strsql As String
    
    Set Adodc1 = New ADODB.Recordset
    
    strsql = "SELECT * FROM comovc,estado_oc WHERE comovc.oc_csitord=estado_oc." & _
        "est_codigo ORDER BY oc_cnumord"
    Adodc1.Open strsql, cConexCom, adOpenDynamic, adLockOptimistic
    
    Set DataGrid1.DataSource = Adodc1
    
    If DataGrid1.Visible Then DataGrid1.SetFocus
End Sub

Private Sub cmdEdi2_Click()
On Error GoTo Err
    With frmTraEmi1
        .activado = False
        .txtCod = Flex1.TextMatrix(Flex1.Row, 0)
        .lblFab = Flex1.TextMatrix(Flex1.Row, 1)
        .txtDes = Flex1.TextMatrix(Flex1.Row, 2)
        .txtDes.Enabled = False
        .lblUni = Flex1.TextMatrix(Flex1.Row, 3)
        .txtCan = Flex1.TextMatrix(Flex1.Row, 4)
        .txtCan.Enabled = True
        .Tipo = Flex1.TextMatrix(Flex1.Row, 14)
        If Flex1.TextMatrix(Flex1.Row, 3) <> Flex1.TextMatrix(Flex1.Row, 5) Then
            .txtURe = Flex1.TextMatrix(Flex1.Row, 5)
            .txtRef = Flex1.TextMatrix(Flex1.Row, 6)
        Else
            .txtURe = ""
            .txtRef = ""
        End If
        If .txtURe <> "" Then .txtRef.Enabled = True
        .txtPUn = Flex1.TextMatrix(Flex1.Row, 7)
        .txtPDe = Flex1.TextMatrix(Flex1.Row, 9)
        .txtPIg = Flex1.TextMatrix(Flex1.Row, 10)
'        .Igv = .txtPIg
        .txtCo1 = Flex1.TextMatrix(Flex1.Row, 12)
        .txtCo2 = Flex1.TextMatrix(Flex1.Row, 13)
        .txtCod.Enabled = False
        .activado = True
        .Calculo_Automatico
        .Show 1
        
        If Not .cancelado Then
            If .Tipo = "S" Then
              .txtCan = 1
            End If
            Flex1.TextMatrix(Flex1.Row, 2) = .txtDes
            Flex1.TextMatrix(Flex1.Row, 4) = .txtCan
            If .txtURe = "" Then
                Flex1.TextMatrix(Flex1.Row, 5) = .lblUni
                Flex1.TextMatrix(Flex1.Row, 6) = .txtCan
            Else
                Flex1.TextMatrix(Flex1.Row, 5) = .txtURe
                Flex1.TextMatrix(Flex1.Row, 6) = .txtRef
            End If
            Flex1.TextMatrix(Flex1.Row, 7) = .txtPUn
'            Flex1.TextMatrix(Flex1.Row, 8) = Format(Val(.lblPNe) + Val(.lblDes), "0.00")
            Flex1.TextMatrix(Flex1.Row, 8) = .txtPUn
            Flex1.TextMatrix(Flex1.Row, 9) = .txtPDe
            Flex1.TextMatrix(Flex1.Row, 10) = .txtPIg
            Flex1.TextMatrix(Flex1.Row, 11) = Format(Flex1.TextMatrix(Flex1.Row, 6) * Flex1.TextMatrix(Flex1.Row, 8), "0.00")
            Flex1.TextMatrix(Flex1.Row, 12) = .txtCo1
            Flex1.TextMatrix(Flex1.Row, 13) = .txtCo2
            Calcula_Totales
        End If
        Flex1.SetFocus
        cmdNue2.SetFocus
    End With
 Exit Sub
Err:
    MsgBox Err.Description
 
End Sub

Private Sub CmdEli_Click()
    On Error GoTo EliErr
    
    If Adodc1("oc_csitord") <> "00" And Adodc1("oc_csitord") <> "01" Then
        Mensaje = "Imposible anular la Orden de compra en su estado actual"
        MsgBox Mensaje, vbCritical, "Mensaje"
        DataGrid1.SetFocus
        Exit Sub
    End If

    Dim strsql As String
    Dim voc As String
    
    Mensaje = "¿Está seguro que desea anular la Orden de compra?"
    If MsgBox(Mensaje, vbQuestion + vbYesNo + vbDefaultButton2, "Mensaje") = vbYes Then
        voc = Adodc1("oc_cnumord")
        
        nTra = 1
        cConexCom.BeginTrans
        
        strsql = "UPDATE comovd SET oc_cestado='06' WHERE oc_cnumord='" & voc & "'"
        cConexCom.Execute strsql
        strsql = "UPDATE comovc SET oc_csitord='06' WHERE oc_cnumord='" & voc & "'"
        cConexCom.Execute strsql

        cConexCom.CommitTrans
        nTra = 0
        
        If nTra = 0 Then
            Adodc1.Requery
            Adodc1.Find "oc_cnumord='" & voc & "'"
        End If
    End If
    DataGrid1.SetFocus
    Exit Sub
Exit Sub
    
    
    Dim Adodc2 As ADODB.Recordset
'    Dim strSQL As String
'    Dim cSql2 As String
'    Dim PrimFil As Variant
'    Dim vCodigo As String
    
'    On Error GoTo EliErr

    Mensaje = "¿Desea eliminar el documento " & Adodc1("nrorequi") & "?"
    If MsgBox(Mensaje, vbQuestion + vbYesNo, "Mensaje") = vbYes Then
        strsql = "DELETE * FROM requisd WHERE nrorequi='" & Adodc1("nrorequi") & "'"
        
        nTra = 1
        cConexCom.BeginTrans
        cConexCom.Execute strsql
        cConexCom.CommitTrans
        nTra = 0
        
        If nTra = 0 Then
            Adodc1.Delete
            Adodc1.Update
        End If
        Estado_Botones
            
' Posiblemente tenga que usar este código con SQL
' ===============================================
'            DataGrid1.SetFocus
'            Set Adodc2 = New ADODB.Recordset
'            Set Adodc2 = Adodc1.Clone
'
'            Adodc2.Bookmark = Adodc1.Bookmark
'            PrimFil = DataGrid1.FirstRow
'            Adodc2.MoveNext
'            If Adodc2.EOF Then
'                Adodc2.MovePrevious
'                Adodc2.MovePrevious
'                If Not Adodc2.BOF Then
'                    vCodigo = Adodc2("acodigo")
'                Else
'                    vCodigo = ""
'                End If
'                Adodc2.MoveNext
'            Else
'                vCodigo = Adodc2("acodigo")
'                Adodc2.MovePrevious
'            End If
'
'            csql1 = "DELETE FROM maeart WHERE acodigo='" & cCod & "'"
'            cSql2 = "DELETE FROM stkart WHERE stcodigo='" & cCod & "'"
'            nTra = 1
'            cConexCom.BeginTrans
'            cConexCom.Execute csql1
'            cConexCom.Execute cSql2
'            cConexCom.CommitTrans
'            nTra = 0
'
'            Adodc2.Requery
'
'            Adodc2.Find "acodigo LIKE '" & vCodigo & "'"
'
'            Set Adodc1 = Adodc2.Clone
'            Adodc1.Bookmark = Adodc2.Bookmark
'
'            Set DataGrid1.DataSource = Adodc1
'            DataGrid1.Refresh
'            DataGrid1.SetFocus
'            DataGrid1.SetFocus
    End If
    If Adodc1.RecordCount > 0 Then
        DataGrid1.SetFocus
    Else
        cmdNue.SetFocus
    End If
    Exit Sub

EliErr:
    MsgBox Err.Description
    If nTra = 1 Then cConexCom.RollbackTrans
End Sub

Private Sub CmdEli2_Click()
    If Tiene_Entregas Then
        Mensaje = "El artículo tiene cantidad entregada"
        MsgBox Mensaje, vbExclamation, "Advertencia"
    End If
    
    Mensaje = "¿Desea quitar el artículo seleccionado?"
    If MsgBox(Mensaje, vbQuestion + vbYesNo + vbDefaultButton2, "Mensaje") = vbYes Then
        If Flex1.Rows - 1 = 1 Then
            Dim I As Integer
            
            For I = 0 To 13
                Flex1.TextMatrix(1, I) = ""
            Next
        Else
            Flex1.RemoveItem Flex1.Row
        End If
        Calcula_Totales
        Estado_Items
    End If
End Sub

Private Sub cmdGra_Click()
    Dim SQLc As String
    Dim SQLd As String
    Dim I As Integer
    Dim vFactor As Single, vCantid As Single
    Dim vPreuni As Single, vDscpor As Single
    Dim vDescto As Single, vIgv As Single
    Dim vIgvpor As Single, vPrenet As Single
    Dim vTotven As Single, vTotnet As Single
    Dim vURef As String
    
    On Error GoTo GrabErr
    
    If nT = 1 Then
        txtPro = Trim(txtPro)
        If txtPro = "" Then
            Mensaje = "Debe ingresar Código de Proveedor"
            MsgBox Mensaje, vbExclamation, "Mensaje"
            txtPro.SetFocus
            Exit Sub
        Else
            If lblPro = "" Then
                If Not Existe(1, txtPro, "maeprov", "prvccodigo", False) Then
                    Mensaje = "El Código de Proveedor ingresado no existe"
                    MsgBox Mensaje, vbExclamation, "Mensaje"
                    txtPro.SetFocus
                    Exit Sub
                Else
                    txtPro_KeyPress 13
                    cmdGra.SetFocus
                End If
            End If
        End If
    End If
    
    'No debe ir
    'If Not ValidFecha(txtEmi) Then
    '    mensaje = "Fecha no válida"
    '    MsgBox mensaje, vbExclamation, "Error"
    '    txtEmi.SetFocus
    '    Exit Sub
    'End If
    
    'If Not ValidFecha(txtEnt) Then
    '    mensaje = "Fecha no válida"
    '    MsgBox mensaje, vbExclamation, "Error"
    '    txtEnt.SetFocus
    '    Exit Sub
    'End If
    
    If txtEmi > txtEnt Then
       MsgBox "Fecha de emision no debe ser mayor a la fecha de entrega", vbExclamation, "Error"
       Exit Sub
       txtEmi.SetFocus
    End If
       
    txtMon = Trim(txtMon)
    If txtMon = "" Then
        Mensaje = "Debe ingresar el Tipo de Moneda"
        MsgBox Mensaje, vbExclamation, "Error"
        txtMon.SetFocus
        Exit Sub
    Else
        If Not Existe(1, txtMon, "tipo_moneda", "tipomon_codigo", False) Then
            Mensaje = "El tipo de moneda ingresado no existe"
            MsgBox Mensaje, vbExclamation, "Error"
            txtMon.SetFocus
            Exit Sub
        End If
    End If
    
    txtEst = Trim(txtEst)
    If txtEst = "" Then
        Mensaje = "Debe ingresar el Tipo de conversión"
        MsgBox Mensaje, vbExclamation, "Error"
        txtEst.SetFocus
        Exit Sub
    Else
        If Not Existe(1, txtEst, "conv_moneda", "covmon_codigo", False) Then
            Mensaje = "El Tipo de conversión ingresado no existe"
            MsgBox Mensaje, vbExclamation, "Error"
            txtEst.SetFocus
            Exit Sub
        End If
    End If
    
    If Val(txtTip) = 0 Then
        Mensaje = "Debe ingresar Tipo de cambio"
        MsgBox Mensaje, vbExclamation, "Error"
        txtTip.SetFocus
        Exit Sub
    End If
        
    txtSol = Trim(txtSol)
    If txtSol = "" Then
        Mensaje = "Debe ingresar Solicitante"
        MsgBox Mensaje, vbExclamation, "Mensaje"
        txtSol.SetFocus
        Exit Sub
    Else
        If Not Existe(1, txtSol, "solicitantes", "sol_codigo", False) Then
            MsgBox "El Solicitante no existe", vbExclamation, "Mensaje"
            txtSol.SetFocus
            Exit Sub
        Else
            lblSol = Devolver_Dato(1, txtSol, "solicitantes", "sol_codigo", False, _
                "sol_nombre")
        End If
    End If
    
    If Not cmdEli2.Enabled Then
        Mensaje = "Debe especificar artículos de la Orden de Compra"
        MsgBox Mensaje, vbExclamation, "Error"
        cmdNue2.SetFocus
        Exit Sub
    End If
    
    If nT = 1 Then
        Mensaje = "¿Desea ingresar la nueva Orden de Compra?"
    Else
        Mensaje = "¿Desea guardar los cambios realizados?"
    End If
    If MsgBox(Mensaje, vbQuestion + vbYesNo, "Mensaje") = vbYes Then
        nTra = 1
        cConexCom.BeginTrans
        
        If nT = 1 Then      'Ingreso
            SQLc = "UPDATE num_documentos SET ctnnumero=" & Val(unum) & _
                " WHERE ctncodigo='OC'"
            cConexCom.Execute SQLc
            
            SQLc = "INSERT INTO comovc (oc_cnumord,oc_dfecdoc,oc_ccodpro,oc_crazsoc," & _
                "oc_cdirpro,oc_ccotiza,oc_ccodmon,oc_cforpag,oc_ntipcam,oc_dfecent," & _
                "oc_cobserv,oc_csolict,oc_centreg,oc_csitord,oc_nimport,oc_ndescue," & _
                "oc_nigv,oc_nventa,oc_dfecact,oc_chora,oc_cusuari,oc_cconver) VALUES ('" & _
                lblNum & "','" & txtEmi & "','" & txtPro & "','" & _
                lblPro & "','" & Devolver_Dato(1, txtPro, "maeprov", "prvccodigo", False, _
                "prvcdirecc") & "','" & txtCot & "','" & txtMon & "','" & txtFor & "'," & _
                Val(txtTip) & ",'" & txtEnt & "','" & _
                SupCadSQL(txtObs) & "','" & txtSol & "','" & txtEntE & "','00'," & _
                CDbl(lblImp) & "," & CDbl(lblDes) & "," & CDbl(lblIgv) & "," & CDbl(lblCom) & _
                ",'" & VG_FecTrab & "','" & Time & "','" & VGUsuario & _
                "','" & txtEst & "')"
            cConexCom.Execute SQLc
            
            For I = 1 To Flex1.Rows - 1
                vFactor = Val(Flex1.TextMatrix(I, 6))
                vCantid = Val(Flex1.TextMatrix(I, 4))
                If vCantid = 0 Then
                   vCantid = 1
                End If
                vPreuni = Val(Flex1.TextMatrix(I, 7))
                vDscpor = Val(Flex1.TextMatrix(I, 9))
                vDescto = IIf(vFactor > 0, vFactor / vCantid, 1) * vPreuni * vCantid * _
                    vDscpor / 100
                vIgvpor = Val(Flex1.TextMatrix(I, 10))
                vTotven = Val(Flex1.TextMatrix(I, 11))
                vIgv = (vTotven - vDescto) * vIgvpor / 100
                vPrenet = Val(Flex1.TextMatrix(I, 8)) * (1 - vDscpor / 100)
                SQLd = "INSERT INTO comovd (oc_cnumord,oc_ccodpro,oc_dfecdoc,oc_citem," & _
                    "oc_ccodigo,oc_ccodref,oc_cdesref,oc_cunidad,oc_cuniref,oc_nfactor," & _
                    "oc_ncantid,oc_npreuni,oc_ndscpor,oc_ndescto,oc_nigv,oc_nigvpor," & _
                    "oc_nprenet,oc_ntotven,oc_ntotnet,oc_cestado,oc_ccomen1,oc_ccomen2, tipord) " & _
                    "VALUES ('" & lblNum & "','" & txtPro & "','" & txtEmi _
                    & "','" & Format(I, "000") & "','" & _
                    Flex1.TextMatrix(I, 0) & "','" & Flex1.TextMatrix(I, 1) & "','" & _
                    Flex1.TextMatrix(I, 2) & "','" & Flex1.TextMatrix(I, 3) & "','" & _
                    Flex1.TextMatrix(I, 5) & "'," & vFactor & "," & vCantid & "," & _
                    vPreuni & "," & vDscpor & "," & vDescto & "," & vIgv & "," & _
                    vIgvpor & "," & vPrenet & "," & vTotven & "," & vTotven - vDescto + _
                    vIgv & ",'00','" & Flex1.TextMatrix(I, 12) & "','" & _
                    Flex1.TextMatrix(I, 13) & "','" & Flex1.TextMatrix(I, 14) & "')"
                cConexCom.Execute SQLd
                
                SQLd = "UPDATE maeart SET aprecom=" & Val(Flex1.TextMatrix(I, 8)) & _
                    ",acodpro='" & txtPro & "',afecven='" & txtEmi _
                    & "' WHERE acodigo='" & Flex1.TextMatrix(I, 0) & "'"
                cConexCom.Execute SQLd
            Next
        ElseIf nT = 2 Then     'Modificar
            SQLc = "UPDATE comovc SET oc_dfecdoc='" & txtEmi & _
                "',oc_ccotiza='" & txtCot & "',oc_ccodmon='" & txtMon & "',oc_cforpag='" & _
                txtFor & "',oc_ntipcam=" & Val(txtTip) & ",oc_dfecent='" & _
                txtEnt & "',oc_cobserv='" & SupCadSQL(txtObs) & _
                "',oc_csolict='" & txtSol & "',oc_centreg='" & txtEntE & "',oc_nimport=" & _
                CDbl(lblImp) & ",oc_ndescue=" & CDbl(lblDes) & ",oc_nigv=" & CDbl(lblIgv) & _
                ",oc_nventa=" & CDbl(lblCom) & ",oc_dfecact='" & _
                VG_FecTrab & "',oc_chora='" & Time & "',oc_cusuari='" & _
                VGUsuario & "',oc_cconver='" & txtEst & "' WHERE oc_cnumord='" & lblNum & "'"
            cConexCom.Execute SQLc
            
            SQLd = "DELETE * FROM comovd WHERE oc_cnumord='" & lblNum & "'"
            cConexCom.Execute SQLd
            
            For I = 1 To Flex1.Rows - 1
                vURef = ""
                vFactor = 0
                If Flex1.TextMatrix(I, 3) <> Flex1.TextMatrix(I, 5) Then
                    vURef = Flex1.TextMatrix(I, 5)
                    vFactor = Val(Flex1.TextMatrix(I, 6))
                End If
                vCantid = Val(Flex1.TextMatrix(I, 4))
                vPreuni = Val(Flex1.TextMatrix(I, 7))
                vDscpor = Val(Flex1.TextMatrix(I, 9))
                vDescto = IIf(vFactor > 0, vFactor / vCantid, 1) * vPreuni * vCantid * _
                    vDscpor / 100
                vIgvpor = Val(Flex1.TextMatrix(I, 10))
                vTotven = Val(Flex1.TextMatrix(I, 11))
                vIgv = (vTotven - vDescto) * vIgvpor / 100
                vPrenet = Val(Flex1.TextMatrix(I, 8)) * (1 - vDscpor / 100)
                SQLd = "INSERT INTO comovd (oc_cnumord,oc_ccodpro,oc_dfecdoc,oc_citem," & _
                    "oc_ccodigo,oc_ccodref,oc_cdesref,oc_cunidad,oc_cuniref,oc_nfactor," & _
                    "oc_ncantid,oc_npreuni,oc_ndscpor,oc_ndescto,oc_nigv,oc_nigvpor," & _
                    "oc_nprenet,oc_ntotven,oc_ntotnet,oc_cestado,oc_ccomen1,oc_ccomen2,tipord) " & _
                    "VALUES ('" & lblNum & "','" & txtPro & "','" & txtEmi _
                    & "','" & Format(I, "000") & "','" & _
                    Flex1.TextMatrix(I, 0) & "','" & Flex1.TextMatrix(I, 1) & "','" & _
                    Flex1.TextMatrix(I, 2) & "','" & Flex1.TextMatrix(I, 3) & "','" & _
                    vURef & "'," & vFactor & "," & vCantid & "," & _
                    vPreuni & "," & vDscpor & "," & vDescto & "," & vIgv & "," & _
                    vIgvpor & "," & vPrenet & "," & vTotven & "," & vTotven - vDescto + _
                    vIgv & ",'00','" & Flex1.TextMatrix(I, 12) & "','" & _
                    Flex1.TextMatrix(I, 13) & "', '" & Flex1.TextMatrix(I, 14) & "')"
                cConexCom.Execute SQLd
                
                SQLd = "UPDATE maeart SET aprecom=" & Val(Flex1.TextMatrix(I, 8)) & _
                    ",acodpro='" & txtPro & "',afecven='" & txtEmi _
                    & "' WHERE acodigo='" & Flex1.TextMatrix(I, 0) & "'"
                cConexCom.Execute SQLd
            Next
        End If
        
        cConexCom.CommitTrans
        nTra = 0
        Adodc1.Requery
        Adodc1.Find "oc_cnumord='" & lblNum & "'"
        
        If nT = 1 Then
            unum = Format(Val(unum) + 1, "0000000000000")
            lblNum = unum
            Limpiar
            Vacia_FlexGrid
            Estado_Items
            Calcula_Totales
            txtEmi = VG_FecTrab
            txtEnt = VG_FecTrab
            txtTip = "0.000"
            'txtEntE = VGEMP_DIREC
            txtPro.SetFocus
        Else
            CmdSalir2_Click
        End If
    End If
    Exit Sub

GrabErr:
    MsgBox Err.Description
    If nTra = 1 Then cConexCom.RollbackTrans
End Sub

'Private Sub cmdImp_Click()
'    Dim NIGV As Double
'    Dim rsI As New Recordset
'    Dim strsql As String
'
'    Set rsI = New ADODB.Recordset
'    strsql = "SELECT igv FROM ordencompra_igv WHERE oc_cnumord='" & DataGrid1.Columns(0) & _
'        "'"
'    rsI.Open strsql, cConexCom, adOpenStatic, adLockReadOnly
'    NIGV = rsI(0)
'
'    On Error GoTo FALLO
'
'    MDIPrincipal.Data1.DatabaseName = WCompPATH & "Data\" & VGEMP_CODIGO & "\" & NAMEBD & ".mdb"
'    'Data1.DatabaseName = cRuta2
'    strsql = "SELECT * FROM ordencompras WHERE oc_cnumord='" & DataGrid1.Columns(0) & _
'        "' ORDER BY oc_citem"
'    MDIPrincipal.Data1.RecordSource = strsql
'    MDIPrincipal.Data1.Refresh
'
'    MDIPrincipal.CrystalReport1.Reset
'    MDIPrincipal.CrystalReport1.ReportFileName = WCompPATH & "Reportes\comp0005.RPT"
'    ' Ubi_Tab CrystalReport2
'    MDIPrincipal.CrystalReport1.WindowTitle = "StarSoft - Compras - comp0005.rpt"
'    MDIPrincipal.CrystalReport1.Formulas(0) = "EMPRESA='" & VGEMP_REPORTE & "'"
'    MDIPrincipal.CrystalReport1.Formulas(1) = "HORA='" & Format(Time, "hh:mm:ss") & "'"
'    MDIPrincipal.CrystalReport1.Formulas(2) = "IGV=" & NIGV
'    DataGrid1.SetFocus
'    MDIPrincipal.CrystalReport1.Action = 1
'    rsI.Close
'   Exit Sub
'FALLO:
'    MsgBox Err.Description, vbExclamation, "Error"
'End Sub

Private Sub cmdNue_Click()
 Dim cSqlM As String, cSelM As ADODB.Recordset
    nT = 1
    OculObj06 False
    OculObj04 False
    OculObj02 True
    OculObj03 True
    Proceso True
    unum = ""
    If unum = "" Then
        unum = Format(Devolver_Dato(1, "OC", "num_documentos", "ctncodigo", False, _
            "ctnnumero"), "0000000000000")
        If unum = "" Then unum = 0
        unum = unum + 1
        unum = Format(unum, "0000000000000")
            
    ' inicio recien
        ' Selecciona todas las Orden de compra
        cSqlM = "SELECT oc_cnumord FROM comovc ORDER BY oc_cnumord"
        Set cSelM = New ADODB.Recordset
        cSelM.Open cSqlM, cConexCom, adOpenStatic
        Do While Not cSelM.EOF
           If cSelM("oc_cnumord") = Trim(unum) Then
              cSelM.MoveLast
              unum = Format(cSelM("oc_cnumord"), "0000000000000")
              unum = unum + 1
              unum = Format(unum, "0000000000000")
           End If
           cSelM.MoveNext
        Loop
        cSelM.Close 'Cierra el ADODB.Recorset
   ' fin
            
    End If
    lblNum = unum
    lblEst = ""
    'txtEmi = VGFecTrb
    'txtEnt = VGFecTrb
    txtTip = "0.000"
    'txtEntE = VGEMP_DIREC
    lblImp = "0.00": lblTot = "0.00": lblIgv = "0.00"
    lblDes = "0.00": lblCom = "0.00"
    
    fraDatos.Enabled = True
    cmdGra.Enabled = True
    txtPro.Enabled = True
    txtPro.SetFocus
    CmdSalir2.Cancel = True
End Sub

Private Sub cmdEdi_Click()
    If Adodc1("oc_csitord") = "06" Then
        Mensaje = "La Orden de compra ha sido anulada, no se permitirá modificaciones"
        MsgBox Mensaje, vbExclamation, "Advertencia"
        cmdNue2.Enabled = False
        cmdEdi2.Enabled = False
        cmdEli2.Enabled = False
        cmdGra.Enabled = False
        OculObj06 False
        OculObj04 False
        OculObj02 True
        Mostrar Adodc1("oc_cnumord")
        OculObj03 True
        Proceso True
        txtPro.Enabled = True
        fraDatos.Enabled = False
    Else
        nT = 2
        OculObj06 False
        OculObj04 False
        OculObj02 True
        Mostrar Adodc1("oc_cnumord")
        OculObj03 True
        Proceso True
        fraDatos.Enabled = True
        
        cmdEdi2.Enabled = True
        cmdEli2.Enabled = True
        cmdGra.Enabled = True
        
        txtPro.Enabled = False
        txtEmi.SetFocus
        CmdSalir2.Cancel = True
    End If
End Sub

Private Sub cmdNue2_Click()
    With frmTraEmi1
        .activado = False
        .txtCod = ""
        .txtDes = ""
        .txtCan = "0.00"
        .txtPUn = "0.00"
        .txtPDe = "0.00"
        .txtPIg = "18.00"
        .txtCo1 = ""
        .txtCo2 = ""
        .activado = True
       If txtNSol = "" Then
         .cmbtipo.Visible = True
         .cmbtipo.text = "Bienes"
         .lbltipo.Visible = True
       End If
       .Calculo_Automatico
       .Show 1
        
        If Not .cancelado Then
           If .Tipo = "S" Then
              .txtCan = 1
            End If
            If Flex1.Rows - 1 = 1 Then
                If Flex1.TextMatrix(1, 0) = "" Then
                    Flex1.AddItem .txtCod & vbTab & .lblFab & vbTab & .txtDes & vbTab & _
                        .lblUni & vbTab & .txtCan & vbTab & IIf(.txtURe = "", .lblUni, _
                        .txtURe) & vbTab & IIf(.txtURe = "", .txtCan, .txtRef) & vbTab & _
                        .txtPUn & vbTab & Format(Val(.lblPNe) + Val(.lblDes), "0.00") & _
                        vbTab & .txtPDe & vbTab & .txtPIg & vbTab & Format(Val(.txtCan) * _
                        (Val(.lblPNe) + Val(.lblDes)), "0.00") & vbTab & .txtCo1 & vbTab & _
                        .txtCo2 & vbTab & .Tipo, 1
                    Flex1.Rows = 2
                Else
                    Flex1.AddItem .txtCod & vbTab & .lblFab & vbTab & .txtDes & vbTab & _
                        .lblUni & vbTab & .txtCan & vbTab & IIf(.txtURe = "", .lblUni, _
                        .txtURe) & vbTab & IIf(.txtURe = "", .txtCan, .txtRef) & vbTab & _
                        .txtPUn & vbTab & Format(Val(.lblPNe) + Val(.lblDes), "0.00") & _
                        vbTab & .txtPDe & vbTab & .txtPIg & vbTab & Format(Val(.txtCan) * _
                        (Val(.lblPNe) + Val(.lblDes)), "0.00") & vbTab & .txtCo1 & vbTab & _
                        .txtCo2 & vbTab & .Tipo
                    Flex1.Row = Flex1.Rows - 1
                End If
            Else
                Flex1.AddItem .txtCod & vbTab & .lblFab & vbTab & .txtDes & vbTab & _
                    .lblUni & vbTab & .txtCan & vbTab & IIf(.txtURe = "", .lblUni, _
                    .txtURe) & vbTab & IIf(.txtURe = "", .txtCan, .txtRef) & vbTab & _
                    .txtPUn & vbTab & Format(Val(.lblPNe) + Val(.lblDes), "0.00") & _
                    vbTab & .txtPDe & vbTab & .txtPIg & vbTab & Format(Val(.txtCan) * _
                    (Val(.lblPNe) + Val(.lblDes)), "0.00") & vbTab & .txtCo1 & vbTab & _
                    .txtCo2 & vbTab & .Tipo
                Flex1.Row = Flex1.Rows - 1
            End If
            
            Calcula_Totales
            Estado_Items
            Flex1.SetFocus
         '   SendKeys "{Left}"
           cmdNue2.SetFocus
        Else
            Flex1.SetFocus
            cmdNue2.SetFocus
        End If
    End With
End Sub

Private Sub CmdSalir_Click()
    Unload frmReferencia
    Unload frmTraEmi1
    Unload Me
End Sub

Private Sub CmdSalir2_Click()
    Limpiar
    Vacia_FlexGrid
    Estado_Items
    Estado_Botones
    OculObj02 False
    OculObj03 False
    OculObj04 True
    OculObj06 True
    Proceso False
    
    If Adodc1.RecordCount > 0 Then
        DataGrid1.SetFocus
    Else
        cmdNue.SetFocus
    End If
    CmdSalir.Cancel = True
End Sub


Private Sub Form_Load()
    AlinearFrm Me
    Load frmReferencia
    Init_ControlDataGrid DataGrid1
    Formato_FlexGrid
    
    OculObj02 False
    OculObj03 False
    OculObj04 True
    OculObj06 True
    
    unum = ""
    Abre_Tabla_OCs
    Estado_Botones
    
    Load frmTraEmi1
End Sub

Sub Limpiar()
    txtPro = "": txtMon = "": txtEst = "": txtNSol = ""
    txtTip = "": txtFor = "": txtCot = ""
    txtEntE = "": txtSol = "": txtObs = ""
End Sub

Sub Mostrar(cC1 As String)
    Dim cSqlM As String, cSelM As ADODB.Recordset
    Dim k As Integer, I As Integer, vd As String
    Dim vpu As Single
    
    lblNum = cC1
    lblEst = Adodc1("est_nombre")
    txtPro = Adodc1("oc_ccodpro")
    lblPro = Adodc1("oc_crazsoc")
    lblRuc = Devolver_Dato(1, txtPro, "maeprov", "prvccodigo", False, "prvcruc")
    txtEmi = Adodc1("oc_dfecdoc")
    txtEnt = Adodc1("oc_dfecent")
    txtMon = Adodc1("oc_ccodmon")
    txtEst = Adodc1("oc_cconver")
    txtTip = Format(Adodc1("oc_ntipcam"), "0.000")
    txtFor = Adodc1("oc_cforpag")
    txtCot = Adodc1("oc_ccotiza")
    txtEntE = Adodc1("oc_centreg")
    txtSol = Adodc1("oc_csolict")
    lblSol = Devolver_Dato(1, txtSol, "solicitantes", "sol_codigo", False, "sol_nombre")
    txtObs = Adodc1("oc_cobserv")
    
    cSqlM = "SELECT * FROM comovd WHERE oc_cnumord='" & cC1 & "' ORDER BY oc_citem"
    Set cSelM = New ADODB.Recordset
    
    cSelM.Open cSqlM, cConexCom, adOpenStatic
    cSelM.MoveFirst
    
    k = 0
    Do While Not cSelM.EOF
        k = k + 1
        If k = 1 Then
            vpu = IIf(cSelM("oc_nfactor") = 0, 1, cSelM("oc_nfactor") / cSelM("oc_ncantid"))
            Flex1.AddItem cSelM("oc_ccodigo") & vbTab & cSelM("oc_ccodref") & vbTab & _
                cSelM("oc_cdesref") & vbTab & cSelM("oc_cunidad") & vbTab & _
                Format(cSelM("oc_ncantid"), "0.00") & vbTab & IIf(cSelM("oc_cuniref") = "", _
                cSelM("oc_cunidad"), cSelM("oc_cuniref")) & vbTab & _
                IIf(cSelM("oc_cuniref") = "", Format(cSelM("oc_ncantid"), "0.00"), _
                Format(cSelM("oc_nfactor"), "0.00")) & vbTab & Format(cSelM("oc_npreuni"), _
                "0.00") & vbTab & Format(cSelM("oc_npreuni"), "0.00") & vbTab & _
                Format(cSelM("oc_ndscpor"), "0.00") & vbTab & _
                Format(cSelM("oc_nigvpor"), "0.00") & vbTab & _
                Format(cSelM("oc_ntotven"), "0.00") & vbTab & cSelM("oc_ccomen1") & vbTab & _
                cSelM("oc_ccomen2") & vbTab & cSelM("tipord"), 1
            Flex1.Rows = 2
        Else
            vpu = IIf(cSelM("oc_nfactor") = 0, 1, cSelM("oc_nfactor") / cSelM("oc_ncantid"))
'            Flex1.AddItem cSelM("oc_ccodigo") & vbTab & cSelM("oc_ccodref") & vbTab & _
                cSelM("oc_cdesref") & vbTab & cSelM("oc_cunidad") & vbTab & _
                Format(cSelM("oc_ncantid"), "0.00") & vbTab & IIf(cSelM("oc_cuniref") = "", _
                cSelM("oc_cunidad"), cSelM("oc_cuniref")) & vbTab & _
                IIf(cSelM("oc_cuniref") = "", Format(cSelM("oc_ncantid"), "0.00"), _
                Format(cSelM("oc_nfactor"), "0.00")) & vbTab & Format(cSelM("oc_npreuni"), _
                "0.00") & vbTab & Format(cSelM("oc_npreuni") * vpu, "0.00") & vbTab & _
                Format(cSelM("oc_ndscpor"), "0.00") & vbTab & _
                Format(cSelM("oc_nigvpor"), "0.00") & vbTab & _
                Format(cSelM("oc_ntotven"), "0.00") & vbTab & cSelM("oc_ccomen1") & vbTab & _
                cSelM("oc_ccomen2")
            Flex1.AddItem cSelM("oc_ccodigo") & vbTab & cSelM("oc_ccodref") & vbTab & _
                cSelM("oc_cdesref") & vbTab & cSelM("oc_cunidad") & vbTab & _
                Format(cSelM("oc_ncantid"), "0.00") & vbTab & IIf(cSelM("oc_cuniref") = "", _
                cSelM("oc_cunidad"), cSelM("oc_cuniref")) & vbTab & _
                IIf(cSelM("oc_cuniref") = "", Format(cSelM("oc_ncantid"), "0.00"), _
                Format(cSelM("oc_nfactor"), "0.00")) & vbTab & Format(cSelM("oc_npreuni"), _
                "0.00") & vbTab & Format(cSelM("oc_npreuni"), "0.00") & vbTab & _
                Format(cSelM("oc_ndscpor"), "0.00") & vbTab & _
                Format(cSelM("oc_nigvpor"), "0.00") & vbTab & _
                Format(cSelM("oc_ntotven"), "0.00") & vbTab & cSelM("oc_ccomen1") & vbTab & _
                cSelM("oc_ccomen2") & vbTab & cSelM("tipord")
        End If
        cSelM.MoveNext
    Loop
    cSelM.Close
    Calcula_Totales
End Sub

Sub Estado_Botones()
    If Adodc1.RecordCount > 0 Then
        cmdEdi.Enabled = True
        CmdEli.Enabled = True
        cmdImp.Enabled = True
    Else
        cmdEdi.Enabled = False
        CmdEli.Enabled = False
        cmdImp.Enabled = False
    End If
End Sub



Private Sub txtCot_GotFocus()
    Enfoque txtCot
End Sub

Private Sub txtCot_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtEntE.SetFocus
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub


Private Sub txtEmi_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
  SendKeys "{TAB}"
End If
End Sub

Private Sub txtEmi_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Not ValidFecha(txtEmi) Then
            Mensaje = "Fecha No Válida"
            MsgBox Mensaje, vbExclamation, "Error"
            txtEmi.SetFocus
        Else
            txtEnt.SetFocus
        End If
    End If
End Sub


Private Sub txtEnt_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
  SendKeys "{TAB}"
End If
End Sub

Private Sub txtEnt_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Not ValidFecha(txtEnt) Then
            Mensaje = "Fecha No Válida"
            MsgBox Mensaje, vbExclamation, "Error"
            txtEnt.SetFocus
        Else
            txtMon.SetFocus
        End If
    End If
End Sub

Private Sub txtEntE_GotFocus()
    Enfoque txtEntE
End Sub

Private Sub txtEntE_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtSol.SetFocus
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub
'
'Private Sub txtEst_DblClick()
'    Static Adodc2 As ADODB.Recordset
'    Dim strsql As String
'
'    Set Adodc2 = New ADODB.Recordset
'
'    If UCase(Dir$(cRuta4)) = UCase(VGNameCont & ".MDB") Then
'       strsql = "SELECT covmon_codigo,covmon_descripcion FROM conversion_moneda"
'       Adodc2.Open strsql, cConexCont, adOpenStatic, adLockReadOnly
'    Else
'       strsql = "SELECT covmon_codigo,covmon_descripcion FROM conv_moneda"
'       Adodc2.Open strsql, cConexCom, adOpenStatic, adLockReadOnly
'    End If
'
'    frmReferencia.Conectar Adodc2, strsql
'    frmReferencia.lblTit = "Tipo de Cambio"
'    frmReferencia.Inicio
'    frmReferencia.show vbmodal
'    Adodc2.Close
'
'    If vGUtil(1) <> "" Then
'        txtEst = vGUtil(1)
'        txtEst_KeyPress 13
'    End If
'End Sub

Private Sub txtEst_GotFocus()
    Enfoque txtEst
End Sub

Private Sub txtEst_KeyDown(KeyCode As Integer, Shift As Integer)
    'If KeyCode = 112 Then txtEst_DblClick
End Sub

Private Sub txtEst_KeyPress(KeyAscii As Integer)
Dim strsql As String
Dim Adodc2 As New ADODB.Recordset
    
    If KeyAscii = 13 Then
        txtEst = Trim(txtEst)
        If txtEst <> "" Then
          If UCase(Dir$(cRuta4)) = VGNameCont & ".MDB" Then
             If Not Existe(1, txtEst, "conv_moneda", "covmon_codigo", False) Then
                MsgBox "El Tipo de Conversión no existe", vbExclamation, "Mensaje"
                txtEst.SetFocus
             Else
                txtTip.SetFocus
             End If
          Else
            If Not Existe(3, txtEst, "conversion_moneda", "covmon_codigo", False) Then
                MsgBox "El Tipo de Conversión no existe", vbExclamation, "Mensaje"
                txtEst.SetFocus
             Else
                txtTip.SetFocus
             End If
          End If
        Else
            txtTip.SetFocus
        End If
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
    
    'aqui *****
   Set Adodc2 = New ADODB.Recordset
   If txtEst = "COM" Then
    strsql = "Select tipocamb_compra from Tipo_cambio where tipomon_codigo='ME' and tipocamb_fecha = # " & Format(txtEmi.Value, "mm/dd/yyyy") & "#"
       If UCase(Dir$(cRuta4)) = VGNameCont & ".MDB" Then
          Adodc2.Open strsql, cConexCont, adOpenStatic, adLockReadOnly
       Else
          Adodc2.Open strsql, cConexCom, adOpenStatic, adLockReadOnly
       End If
      If Adodc2.RecordCount <> 0 Then
       txtTip.text = Format(Adodc2.Fields("tipocamb_compra"), ".000")
      End If
   Else
    If txtEst = "VTA" Then
      strsql = "Select tipocamb_fecha, tipocamb_venta from Tipo_cambio where tipomon_codigo='ME' and tipocamb_fecha = #" & Format(txtEmi.Value, "mm/dd/yyyy") & "#"
     If UCase(Dir$(cRuta4)) = VGNameCont & ".MDB" Then
       Adodc2.Open strsql, cConexCont, adOpenStatic, adLockReadOnly
     Else
       Adodc2.Open strsql, cConexCom, adOpenStatic, adLockReadOnly
     End If
     If Adodc2.RecordCount <> 0 Then
        txtTip.text = Format(Adodc2.Fields("tipocamb_venta"), ".000")
     End If
    Else
       If txtEst = "ESP" Or txtEst = "FEC" Then
           txtTip.text = Format(0, ".000")
        End If
 End If
End If
End Sub

Private Sub txtFor_GotFocus()
    Enfoque txtFor
End Sub

Private Sub txtFor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtCot.SetFocus
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

'Private Sub txtMon_DblClick()
'    Static Adodc2 As ADODB.Recordset
'    Dim strsql As String
'
'    Set Adodc2 = New ADODB.Recordset
'
'    strsql = "SELECT tipomon_codigo,tipomon_descripcion FROM tipo_moneda"
'    Adodc2.Open strsql, cConexCom, adOpenStatic, adLockReadOnly
'
'    frmReferencia.Conectar Adodc2, strsql
'    frmReferencia.lblTit = "Tipo de Moneda"
'    frmReferencia.Inicio
'    frmReferencia.show vbmodal
'    Adodc2.Close
'
'    If vGUtil(1) <> "" Then
'        txtMon = vGUtil(1)
'        txtMon_KeyPress 13
'    End If
'End Sub

Private Sub txtMon_GotFocus()
    Enfoque txtMon
End Sub

Private Sub txtMon_KeyDown(KeyCode As Integer, Shift As Integer)
    'If KeyCode = 112 Then txtMon_DblClick
End Sub

Private Sub txtMon_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtMon = Trim(txtMon)
        If txtMon <> "" Then
            If Not Existe(1, txtMon, "tipo_moneda", "tipomon_codigo", False) Then
                MsgBox "El Tipo de moneda no existe", vbExclamation, "Mensaje"
                txtMon.SetFocus
            Else
                txtEst.SetFocus
            End If
        Else
            txtEst.SetFocus
        End If
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub



'Private Sub txtNSol_DblClick()
'Static Adodc2 As ADODB.Recordset
'    Dim strsql As String
'
'    Set Adodc2 = New ADODB.Recordset
'
'    strsql = "SELECT scnumdoc,scfecdoc FROM scc001 where scprovee='" & txtPro & "'"
'    Adodc2.Open strsql, cConexCom, adOpenStatic, adLockReadOnly
'
'    frmReferencia.Conectar Adodc2, strsql
'    frmReferencia.lblTit = "Lista de Solicitudes de Cotizacion"
'    frmReferencia.Inicio
'    frmReferencia.show vbmodal
'    Adodc2.Close
'
'    If vGUtil(1) <> "" Then
'        txtNSol = vGUtil(1)
'        txtEmi.SetFocus
'    End If
'End Sub

Private Sub txtNSol_GotFocus()
    Enfoque txtNSol
End Sub

Private Sub txtNSol_KeyDown(KeyCode As Integer, Shift As Integer)
    'If KeyCode = 112 Then txtNSol_DblClick
End Sub

Private Sub txtNSol_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtNSol = Trim(txtNSol)
        If txtNSol <> "" Then
            If Not Existe(1, txtNSol, "scc001", "scnumdoc", False) Then
                Mensaje = "El Código de Solicitud de Cotizacion ingresado no existe"
                MsgBox Mensaje, vbExclamation, "Error"
                txtNSol.SetFocus
            Else
                txtEmi.SetFocus
            End If
        Else
            txtEmi.SetFocus
        End If
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Private Sub txtObs_GotFocus()
    Enfoque txtObs
End Sub

Private Sub txtObs_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If cmdEli2.Enabled Then
            Flex1.SetFocus
        Else
            cmdNue2.SetFocus
        End If
    End If
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtPro_Change()
    If lblPro <> "" Then
        lblPro = ""
        lblRuc = ""
    End If
End Sub

'Private Sub txtPro_DblClick()
'    Static Adodc2 As ADODB.Recordset
'    Dim strsql As String
'
'    Set Adodc2 = New ADODB.Recordset
'
'    strsql = "SELECT prvccodigo,prvcnombre,prvcruc FROM maeprov"
'    Adodc2.Open strsql, cConexCom, adOpenStatic, adLockReadOnly
'
'    frmReferencia.Conectar Adodc2, strsql
'    frmReferencia.lblTit = "Lista de Proveedores"
'    frmReferencia.Inicio
'    frmReferencia.show vbmodal
'    Adodc2.Close
'
'    If vGUtil(1) <> "" Then
'        txtPro = vGUtil(1)
'        lblPro = vGUtil(2)
'        lblRuc = vGUtil(3)
'        txtNSol.SetFocus
'    End If
'End Sub

Private Sub txtPro_GotFocus()
    Enfoque txtPro
End Sub

Private Sub txtPro_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = 112 Then txtPro_DblClick
End Sub

Private Sub txtPro_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtPro = Trim(txtPro)
        If txtPro <> "" Then
            If Not Existe(1, txtPro, "maeprov", "prvccodigo", False) Then
                Mensaje = "El Código de Proveedor ingresado no existe"
                MsgBox Mensaje, vbExclamation, "Error"
                txtPro.SetFocus
            Else
                lblPro = Devolver_Dato(1, txtPro, "maeprov", "prvccodigo", False, "prvcnombre")
                lblRuc = Devolver_Dato(1, txtPro, "maeprov", "prvccodigo", False, "prvcruc")
                txtEmi.SetFocus
            End If
        Else
            txtPro.SetFocus
        End If
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Private Sub txtSol_Change()
    If lblSol <> "" Then lblSol = ""
End Sub

'Private Sub txtSol_DblClick()
'    Static Adodc2 As ADODB.Recordset
'    Dim strsql As String
'
'    Set Adodc2 = New ADODB.Recordset
'
'    strsql = "SELECT sol_codigo,sol_nombre FROM solicitantes"
'    Adodc2.Open strsql, cConexCom, adOpenStatic, adLockReadOnly
'
'    frmReferencia.Conectar Adodc2, strsql
'    frmReferencia.lblTit = "Solicitantes"
'    frmReferencia.Inicio
'    frmReferencia.show vbmodal
'    Adodc2.Close
'
'    If vGUtil(1) <> "" Then
'        txtSol = vGUtil(1)
'        lblSol = vGUtil(2)
'        txtObs.SetFocus
'    End If
'End Sub

Private Sub txtSol_GotFocus()
    Enfoque txtSol
End Sub

Private Sub txtSol_KeyDown(KeyCode As Integer, Shift As Integer)
    'If KeyCode = 112 Then txtSol_DblClick
End Sub

Private Sub txtSol_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtSol = Trim(txtSol)
        If txtSol <> "" Then
            If Not Existe(1, txtSol, "solicitantes", "sol_codigo", False) Then
                MsgBox "El Solicitante no existe", vbExclamation, "Mensaje"
                txtSol.SetFocus
            Else
                lblSol = Devolver_Dato(1, txtSol, "solicitantes", "sol_codigo", False, _
                    "sol_nombre")
                txtObs.SetFocus
            End If
        Else
            txtObs.SetFocus
        End If
    Else
        If KeyAscii = 39 Or KeyAscii = 34 Then KeyAscii = 0
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Sub Proceso(Estado As Boolean)
    Flex1.Visible = Estado
    cmdNue2.Visible = Estado
    cmdEdi2.Visible = Estado
    cmdEli2.Visible = Estado
    If Estado Then
        frmTraEmi.Height = 7000
    Else
        frmTraEmi.Height = 5145
    End If
End Sub

Sub Formato_FlexGrid()
    Flex1.ColWidth(0) = 1100
    Flex1.ColWidth(1) = 0
    Flex1.ColWidth(2) = 2800
    Flex1.ColWidth(3) = 0
    Flex1.ColWidth(4) = 0
    Flex1.ColWidth(5) = 450
    Flex1.ColWidth(6) = 900
    Flex1.ColWidth(7) = 0
    Flex1.ColWidth(8) = 1200
    Flex1.ColWidth(9) = 700
    Flex1.ColWidth(10) = 0
    Flex1.ColWidth(11) = 1200
    Flex1.ColWidth(12) = 0
    Flex1.ColWidth(13) = 0
    Flex1.ColWidth(14) = 5
    Flex1.ScrollBars = flexScrollBarHorizontal
End Sub

Sub Estado_Items()
    If Flex1.Rows - 1 = 1 Then
        If Flex1.TextMatrix(1, 0) = "" Then
            cmdEdi2.Enabled = False
            cmdEli2.Enabled = False
            cmdNue2.Enabled = True
            cmdNue2.SetFocus
        Else
            cmdEdi2.Enabled = True
            cmdEli2.Enabled = True
        End If
    Else
        cmdEdi2.Enabled = True
        cmdEli2.Enabled = True
    End If
End Sub

Sub Vacia_FlexGrid()
    Dim I As Integer
    
    Do While Flex1.Rows - 1 > 1
        Flex1.RemoveItem 1
    Loop
    
    For I = 0 To 14
        Flex1.TextMatrix(1, I) = ""
    Next
End Sub

Private Sub txtTip_GotFocus()
    Enfoque txtTip
End Sub

Private Sub txtTip_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txtFor.SetFocus
    End If
    Reales_Positivos KeyAscii, txtTip
End Sub

Private Sub txtTip_LostFocus()
    txtTip = Format(Val(txtTip), "0.000")
End Sub

Sub Calcula_Totales()
    Dim I As Integer
    Dim tV As Single, VALOR As Single
    Dim tD As Single, vDesc As Single
    Dim tI As Single, vIgv As Single
    
    With Flex1
        For I = 1 To Flex1.Rows - 1
            tV = Val(.TextMatrix(I, 11))
            VALOR = VALOR + tV
            tD = tV * Val(.TextMatrix(I, 9)) / 100
            vDesc = vDesc + tD
            tI = (tV - tD) * Val(.TextMatrix(I, 10)) / 100
            vIgv = vIgv + tI
        Next
    End With
    
    lblImp = Format(VALOR, "##,##0.00")
    lblDes = Format(vDesc, "##,##0.00")
    lblTot = Format(VALOR - vDesc, "#,##0.00")
    lblIgv = Format(vIgv, "#,##0.00")
    lblCom = Format((VALOR - vDesc) + vIgv, "#,##0.00")
End Sub

Function Tiene_Entregas() As Boolean
    Dim Adodc2 As ADODB.Recordset
    
    Set Adodc2 = New ADODB.Recordset
    
    Adodc2.Open "SELECT * FROM comovd WHERE oc_cnumord='" & lblNum & "' AND oc_ccodigo='" & _
        Flex1.TextMatrix(Flex1.Row, 0) & "' AND oc_ncanten>0", cConexCom, adOpenStatic
    Tiene_Entregas = False
    If Adodc2.RecordCount > 0 Then Tiene_Entregas = True
End Function

Sub pruebita()
    Dim vcontrol As Control
    Dim I As Integer
    
    For Each vcontrol In Controls
        If TypeOf vcontrol Is TextBox Or TypeOf vcontrol Is MaskEdBox Then
            If vcontrol.Container.name = "fraDatos" Then
                Colex.Add item:=vcontrol
            End If
        End If
    Next
    
    For I = 1 To Colex.count
        Colex.item(I).BackColor = &HC0C0FF
    Next
End Sub
