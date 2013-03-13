VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmTraEmi 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Emisión de Orden de Compra"
   ClientHeight    =   6540
   ClientLeft      =   1125
   ClientTop       =   2835
   ClientWidth     =   9435
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmTraEmi.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6540
   ScaleWidth      =   9435
   Begin VB.Frame fraDatos 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2520
      Left            =   96
      TabIndex        =   23
      Top             =   576
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
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         Format          =   47382529
         CurrentDate     =   37015
      End
      Begin MSComCtl2.DTPicker txtEnt 
         Height          =   285
         Left            =   3600
         TabIndex        =   2
         Top             =   600
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         Format          =   47382529
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
         Height          =   192
         Left            =   336
         TabIndex        =   30
         Top             =   288
         Width           =   1008
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
      Picture         =   "frmTraEmi.frx":08CA
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
      Picture         =   "frmTraEmi.frx":0D0C
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
      Picture         =   "frmTraEmi.frx":114E
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
      Picture         =   "frmTraEmi.frx":1590
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
      Picture         =   "frmTraEmi.frx":19D2
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   3825
      Width           =   775
   End
   Begin VB.CommandButton CmdEli 
      Caption         =   "&Anular"
      Height          =   675
      Left            =   4230
      Picture         =   "frmTraEmi.frx":1E14
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   3840
      Width           =   775
   End
   Begin VB.CommandButton cmdNue 
      Caption         =   "&Nuevo"
      Height          =   675
      Left            =   1575
      Picture         =   "frmTraEmi.frx":2256
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
      Picture         =   "frmTraEmi.frx":2698
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   3825
      Width           =   775
   End
   Begin VB.CommandButton cmdGra 
      Caption         =   "&Grabar"
      Height          =   675
      Left            =   5520
      Picture         =   "frmTraEmi.frx":2ADA
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
      Picture         =   "frmTraEmi.frx":2F1C
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
      _ExtentX        =   16140
      _ExtentY        =   2672
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
         Size            =   8.25
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
         Size            =   8.25
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
            Size            =   8.25
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
            Size            =   8.25
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
      Bindings        =   "frmTraEmi.frx":335E
      Left            =   0
      Top             =   3960
      _ExtentX        =   741
      _ExtentY        =   741
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
      _ExtentX        =   15796
      _ExtentY        =   4128
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
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
            ColumnWidth     =   1365.165
         EndProperty
         BeginProperty Column01 
            Locked          =   -1  'True
            WrapText        =   -1  'True
            ColumnWidth     =   3105.071
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
            Locked          =   -1  'True
            WrapText        =   -1  'True
            ColumnWidth     =   1110.047
         EndProperty
         BeginProperty Column03 
            Alignment       =   2
            Locked          =   -1  'True
            WrapText        =   -1  'True
            ColumnWidth     =   434.835
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
            ColumnWidth     =   14.74
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmTraEmi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Colex As New Collection
Dim adodc1 As ADODB.Recordset
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
    Fradatos.Visible = nTipo
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
    
    Set adodc1 = New ADODB.Recordset
    
    strsql = "SELECT * FROM co_cabordcompra,co_estadoorden WHERE co_cabordcompra.oc_situacionorden =co_estadoorden." & _
        "estadooccodigo and estadoocatendido<>1 ORDER BY oc_cnumord "
    adodc1.Open strsql, VGCNx, adOpenDynamic, adLockOptimistic
    
    Set DataGrid1.DataSource = adodc1
    
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
        .tipo = Flex1.TextMatrix(Flex1.Row, 14)
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
        .txtordfab = Flex1.TextMatrix(Flex1.Row, 12)
        .txtCo1 = Flex1.TextMatrix(Flex1.Row, 13)
        .txtCod.Enabled = False
        .activado = True
        .Calculo_Automatico
        .Show 1
        
        If Not .cancelado Then
            If .tipo = "S" Then
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
            Flex1.TextMatrix(Flex1.Row, 8) = .txtPUn
            Flex1.TextMatrix(Flex1.Row, 9) = .txtPDe
            Flex1.TextMatrix(Flex1.Row, 10) = .txtPIg
            Flex1.TextMatrix(Flex1.Row, 11) = Format(Flex1.TextMatrix(Flex1.Row, 6) * Flex1.TextMatrix(Flex1.Row, 8), "0.00")
            Flex1.TextMatrix(Flex1.Row, 12) = .txtordfab
            Flex1.TextMatrix(Flex1.Row, 13) = .txtCo1
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
    
    If adodc1("oc_csitord") <> "00" And adodc1("oc_csitord") <> "01" Then
        Mensaje = "Imposible anular la Orden de compra en su estado actual"
        MsgBox Mensaje, vbCritical, "Mensaje"
        DataGrid1.SetFocus
        Exit Sub
    End If

    Dim strsql As String
    Dim voc As String
    
    Mensaje = "¿Está seguro que desea anular la Orden de compra?"
    If MsgBox(Mensaje, vbQuestion + vbYesNo + vbDefaultButton2, "Mensaje") = vbYes Then
        voc = adodc1("oc_cnumord")
        
        nTra = 1
        VGCNx.BeginTrans
        
        strsql = "UPDATE co_detordcompra SET oc_cestado='06' WHERE oc_cnumord='" & voc & "'"
        VGCNx.Execute strsql
        strsql = "UPDATE co_cabordcompra SET oc_csitord='06' WHERE oc_cnumord='" & voc & "'"
        VGCNx.Execute strsql

        VGCNx.CommitTrans
        nTra = 0
        
        If nTra = 0 Then
            adodc1.Requery
            adodc1.Find "oc_cnumord='" & voc & "'"
        End If
    End If
    DataGrid1.SetFocus
    Exit Sub
Exit Sub
    
Dim Adodc2 As ADODB.Recordset

    Mensaje = "¿Desea eliminar el documento " & adodc1("nrorequi") & "?"
    If MsgBox(Mensaje, vbQuestion + vbYesNo, "Mensaje") = vbYes Then
        strsql = "DELETE * FROM requisd WHERE nrorequi='" & adodc1("nrorequi") & "'"
        
        nTra = 1
        VGCNx.BeginTrans
        VGCNx.Execute strsql
        VGCNx.CommitTrans
        nTra = 0
        
        If nTra = 0 Then
            adodc1.Delete
            adodc1.Update
        End If
        Estado_Botones
            
    End If
    If adodc1.RecordCount > 0 Then
        DataGrid1.SetFocus
    Else
        cmdNue.SetFocus
    End If
    Exit Sub

EliErr:
    MsgBox Err.Description
    If nTra = 1 Then VGCNx.RollbackTrans
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
        
        VGCNx.BeginTrans
        unum = Format(Val(unum), "00000000000")
        lblNum = unum
        If nT = 1 Then      'Ingreso
            SQLc = "UPDATE num_documentos SET ctnnumero=" & Val(unum) & _
                " WHERE ctncodigo='OC'"
            VGCNx.Execute SQLc
            
            SQLc = "INSERT INTO co_cabordcompra (oc_cnumord,oc_dfecdoc,oc_ccodpro,oc_crazsoc," & _
                "oc_cdirpro,oc_ccotiza,oc_ccodmon,oc_cforpag,oc_ntipcam,oc_dfecent," & _
                "oc_cobserv,oc_csolict,oc_centreg,oc_estadoorden,oc_situacionorden,oc_nimport,oc_ndescue," & _
                "oc_nigv,oc_nventa,oc_dfecact,oc_chora,oc_cusuari,oc_cconver) VALUES ('" & _
                lblNum & "','" & txtEmi & "','" & txtPro & "','" & _
                lblPro & "','" & Devolver_Dato(1, txtPro, "maeprov", "prvccodigo", False, _
                "prvcdirecc") & "','" & txtCot & "','" & txtMon & "','" & txtFor & "'," & _
                Val(txtTip) & ",'" & txtEnt & "','" & _
                SupCadSQL(txtObs) & "','" & txtSol & "','" & txtEntE & "',' ','0'," & _
                CDbl(lblImp) & "," & CDbl(lblDes) & "," & CDbl(lblIgv) & "," & CDbl(lblCom) & _
                ",'" & VG_FecTrab & "','" & Format(Time, "hh.mm.ss") & "','" & VGUsuario & _
                "','" & txtEst & "')"
            VGCNx.Execute SQLc
            
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
                SQLd = "INSERT INTO co_detordcompra (oc_cnumord,oc_ccodpro,oc_dfecdoc,oc_citem," & _
                  "oc_ccodigo,oc_ccodref,oc_cdesref,oc_cunidad,oc_cuniref,oc_nfactor," & _
                  "oc_ncantid,oc_nsaldo,oc_npreuni,oc_ndscpor,oc_ndescto,oc_nigv,oc_nigvpor," & _
                  "oc_nprenet,oc_ntotven,oc_ntotnet,oc_situacionorden,ord_fabnum,oc_ccomen1, tipoarticulocodigo) " & _
                  "VALUES ('" & lblNum & "','" & txtPro & "','" & txtEmi _
                  & "','" & Format(I, "000") & "','" & _
                  Flex1.TextMatrix(I, 0) & "','" & Flex1.TextMatrix(I, 1) & "','" & _
                  Flex1.TextMatrix(I, 2) & "','" & Flex1.TextMatrix(I, 3) & "','" & _
                  Flex1.TextMatrix(I, 5) & "'," & vFactor & "," & vCantid & "," & vCantid & "," & _
                  vPreuni & "," & vDscpor & "," & vDescto & "," & vIgv & "," & _
                  vIgvpor & "," & vPrenet & "," & vTotven & "," & vTotven - vDescto + _
                  vIgv & ",'0','" & Flex1.TextMatrix(I, 12) & "','" & _
                  Flex1.TextMatrix(I, 13) & "','" & Flex1.TextMatrix(I, 14) & "')"
                VGCNx.Execute SQLd
                
                SQLd = "UPDATE maeart SET aprecom=" & Val(Flex1.TextMatrix(I, 8)) & _
                    ",acodpro='" & txtPro & "',afecven='" & txtEmi _
                    & "' WHERE acodigo='" & Flex1.TextMatrix(I, 0) & "'"
                VGCNx.Execute SQLd
            Next
        ElseIf nT = 2 Then     'Modificar
            SQLc = "UPDATE co_cabordcompra SET oc_dfecdoc='" & txtEmi & _
                "',oc_ccotiza='" & txtCot & "',oc_ccodmon='" & txtMon & "',oc_cforpag='" & _
                txtFor & "',oc_ntipcam=" & Val(txtTip) & ",oc_dfecent='" & _
                txtEnt & "',oc_cobserv='" & SupCadSQL(txtObs) & _
                "',oc_csolict='" & txtSol & "',oc_centreg='" & txtEntE & "',oc_nimport=" & _
                CDbl(lblImp) & ",oc_ndescue=" & CDbl(lblDes) & ",oc_nigv=" & CDbl(lblIgv) & _
                ",oc_nventa=" & CDbl(lblCom) & ",oc_dfecact='" & _
                VG_FecTrab & "',oc_chora='" & Format(Time, "hh.mm.ss") & "',oc_cusuari='" & _
                VGUsuario & "',oc_cconver='" & txtEst & "' WHERE oc_cnumord='" & lblNum & "'"
            VGCNx.Execute SQLc
            
            SQLd = "DELETE * FROM co_detordcompra WHERE oc_cnumord='" & lblNum & "'"
            VGCNx.Execute SQLd
            
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
                SQLd = "INSERT INTO co_detordcompra (oc_cnumord,oc_ccodpro,oc_dfecdoc,oc_citem," & _
                    "oc_ccodigo,oc_ccodref,oc_cdesref,oc_cunidad,oc_cuniref,oc_nfactor," & _
                    "oc_ncantid,oc_npreuni,oc_ndscpor,oc_ndescto,oc_nigv,oc_nigvpor," & _
                    "oc_nprenet,oc_ntotven,oc_ntotnet,oc_situacionorden,ord_fabnum,oc_ccomen1,tipoarticulocodigo) " & _
                    "VALUES ('" & lblNum & "','" & txtPro & "','" & txtEmi _
                    & "','" & Format(I, "000") & "','" & _
                    Flex1.TextMatrix(I, 0) & "','" & Flex1.TextMatrix(I, 1) & "','" & _
                    Flex1.TextMatrix(I, 2) & "','" & Flex1.TextMatrix(I, 3) & "','" & _
                    vURef & "'," & vFactor & "," & vCantid & "," & _
                    vPreuni & "," & vDscpor & "," & vDescto & "," & vIgv & "," & _
                    vIgvpor & "," & vPrenet & "," & vTotven & "," & vTotven - vDescto + _
                    vIgv & ",'0','" & Flex1.TextMatrix(I, 12) & "','" & _
                    Flex1.TextMatrix(I, 13) & "', '" & Flex1.TextMatrix(I, 14) & "')"
                VGCNx.Execute SQLd
                
                SQLd = "UPDATE maeart SET aprecom=" & Val(Flex1.TextMatrix(I, 8)) & _
                    ",acodpro='" & txtPro & "',afecven='" & txtEmi _
                    & "' WHERE acodigo='" & Flex1.TextMatrix(I, 0) & "'"
                VGCNx.Execute SQLd
            Next
        End If
        
        VGCNx.CommitTrans
        nTra = 0
        adodc1.Requery
        adodc1.Find "oc_cnumord='" & lblNum & "'"
        
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
    If nTra = 1 Then VGCNx.RollbackTrans
End Sub

Private Sub cmdImp_Click()
Dim formulas(1) As String
unum = adodc1("oc_cnumord")
CrystalReport2.Reset
CrystalReport2.WindowTitle = "rptcoordencompra -- orden de compra"
   CrystalReport2.ReportFileName = VGParamSistem.RutaReport & "rptalordencompra.rpt"
    CrystalReport2.DiscardSavedData = True
       

    CrystalReport2.Connect = VGcadenareport2
       
    CrystalReport2.Destination = crptToWindow
    CrystalReport2.WindowState = crptMaximized
    CrystalReport2.WindowShowPrintBtn = True
    CrystalReport2.WindowShowRefreshBtn = True
    CrystalReport2.WindowShowSearchBtn = True
    CrystalReport2.WindowShowPrintSetupBtn = True
    CrystalReport2.formulas(1) = "@emp ='" & VGparametros.RucEmpresa & "'"
    CrystalReport2.StoredProcParam(0) = VGCNx.DefaultDatabase
   CrystalReport2.StoredProcParam(1) = unum
   If CrystalReport2.Status <> 2 Then
      CrystalReport2.Action = 1
   End If

End Sub



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
        cSqlM = "SELECT oc_cnumord FROM co_cabordcompra ORDER BY oc_cnumord"
        Set cSelM = New ADODB.Recordset
        cSelM.Open cSqlM, VGCNx, adOpenStatic
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
    txtTip = "0.000"
    lblImp = "0.00": lblTot = "0.00": lblIgv = "0.00"
    lblDes = "0.00": lblCom = "0.00"
    
    Fradatos.Enabled = True
    cmdGra.Enabled = True
    txtPro.Enabled = True
    txtPro.SetFocus
    CmdSalir2.Cancel = True
End Sub

Private Sub cmdEdi_Click()
    If adodc1("oc_estadoorden") = "A" Then
        Mensaje = "La Orden de compra ha sido anulada, no se permitirá modificaciones"
        MsgBox Mensaje, vbExclamation, "Advertencia"
        cmdNue2.Enabled = False
        cmdEdi2.Enabled = False
        cmdEli2.Enabled = False
        cmdGra.Enabled = False
        OculObj06 False
        OculObj04 False
        OculObj02 True
        Mostrar adodc1("oc_cnumord")
        OculObj03 True
        Proceso True
        txtPro.Enabled = True
        Fradatos.Enabled = False
    Else
        nT = 2
        OculObj06 False
        OculObj04 False
        OculObj02 True
        Mostrar adodc1("oc_cnumord")
        OculObj03 True
        Proceso True
        Fradatos.Enabled = True
        
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
        .txtPIg = "19.00"
        .txtordfab = ""
        .txtCo1 = ""
        .activado = True
       If txtNSol = "" Then
         .cmbtipo.Visible = True
         .cmbtipo.text = "Bienes"
         .lbltipo.Visible = True
       End If
       .Calculo_Automatico
       .Show 1
        
        If Not .cancelado Then
           If .tipo = "S" Then
              .txtCan = 1
            End If
            If Flex1.Rows - 1 = 1 Then
                If Flex1.TextMatrix(1, 0) = "" Then
                    Flex1.AddItem .txtCod & vbTab & .lblFab & vbTab & .txtDes & vbTab & _
                        .lblUni & vbTab & .txtCan & vbTab & IIf(.txtURe = "", .lblUni, _
                        .txtURe) & vbTab & IIf(.txtURe = "", .txtCan, .txtRef) & vbTab & _
                        .txtPUn & vbTab & Format(Val(.lblPNe) + Val(.lblDes), "0.00") & _
                        vbTab & .txtPDe & vbTab & .txtPIg & vbTab & Format(Val(.txtCan) * _
                        (Val(.lblPNe) + Val(.lblDes)), "0.00") & vbTab & .txtordfab & vbTab & _
                        .txtCo1 & vbTab & .tipo, 1
                    Flex1.Rows = 2
                Else
                    Flex1.AddItem .txtCod & vbTab & .lblFab & vbTab & .txtDes & vbTab & _
                        .lblUni & vbTab & .txtCan & vbTab & IIf(.txtURe = "", .lblUni, _
                        .txtURe) & vbTab & IIf(.txtURe = "", .txtCan, .txtRef) & vbTab & _
                        .txtPUn & vbTab & Format(Val(.lblPNe) + Val(.lblDes), "0.00") & _
                        vbTab & .txtPDe & vbTab & .txtPIg & vbTab & Format(Val(.txtCan) * _
                        (Val(.lblPNe) + Val(.lblDes)), "0.00") & vbTab & .txtordfab & vbTab & _
                        .txtCo1 & vbTab & .tipo
                    Flex1.Row = Flex1.Rows - 1
                End If
            Else
                Flex1.AddItem .txtCod & vbTab & .lblFab & vbTab & .txtDes & vbTab & _
                    .lblUni & vbTab & .txtCan & vbTab & IIf(.txtURe = "", .lblUni, _
                    .txtURe) & vbTab & IIf(.txtURe = "", .txtCan, .txtRef) & vbTab & _
                    .txtPUn & vbTab & Format(Val(.lblPNe) + Val(.lblDes), "0.00") & _
                    vbTab & .txtPDe & vbTab & .txtPIg & vbTab & Format(Val(.txtCan) * _
                    (Val(.lblPNe) + Val(.lblDes)), "0.00") & vbTab & .txtordfab & vbTab & _
                    .txtCo1 & vbTab & .tipo
                Flex1.Row = Flex1.Rows - 1
            End If
            
            Calcula_Totales
            Estado_Items
            Flex1.SetFocus
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
    
    If adodc1.RecordCount > 0 Then
        DataGrid1.SetFocus
    Else
        cmdNue.SetFocus
    End If
    CmdSalir.Cancel = True
End Sub


Private Sub Form_Load()
    AlinearFrm Me
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
    lblEst = adodc1("est_nombre")
    txtPro = adodc1("oc_ccodpro")
    lblPro = adodc1("oc_crazsoc")
    lblRuc = Devolver_Dato(1, txtPro, "maeprov", "prvccodigo", False, "prvcruc")
    txtEmi = adodc1("oc_dfecdoc")
    txtEnt = adodc1("oc_dfecent")
    txtMon = adodc1("oc_ccodmon")
    txtEst = adodc1("oc_cconver")
    txtTip = Format(adodc1("oc_ntipcam"), "0.000")
    txtFor = adodc1("oc_cforpag")
    txtCot = adodc1("oc_ccotiza")
    txtEntE = adodc1("oc_centreg")
    txtSol = adodc1("oc_csolict")
    lblSol = Devolver_Dato(1, txtSol, "solicitantes", "sol_codigo", False, "sol_nombre")
    txtObs = adodc1("oc_cobserv")
    
    cSqlM = "SELECT * FROM co_detordcompra WHERE oc_cnumord='" & cC1 & "' ORDER BY oc_citem"
    Set cSelM = New ADODB.Recordset
    
    cSelM.Open cSqlM, VGCNx, adOpenStatic
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
                Format(cSelM("oc_ntotven"), "0.00") & vbTab & cSelM("ord_fabnum") & vbTab & _
                cSelM("oc_ccomen1") & vbTab & cSelM("tipoarticulocodigo"), 1
            Flex1.Rows = 2
        Else
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
                Format(cSelM("oc_ntotven"), "0.00") & vbTab & cSelM("ord_fabnum") & vbTab & _
                cSelM("oc_ccomen1") & vbTab & cSelM("tipoarticulocodigo")
        End If
        cSelM.MoveNext
    Loop
    cSelM.Close
    Calcula_Totales
End Sub

Sub Estado_Botones()
    If adodc1.RecordCount > 0 Then
      '  cmdEdi.Enabled = True
      '  CmdEli.Enabled = True
        cmdImp.Enabled = True
    Else
       ' cmdEdi.Enabled = False
      '  CmdEli.Enabled = False
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

Private Sub txtEst_GotFocus()
    Enfoque txtEst
End Sub

Private Sub txtEst_KeyPress(KeyAscii As Integer)
Dim strsql As String
Dim Adodc2 As New ADODB.Recordset
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

Private Sub txtMon_GotFocus()
    Enfoque txtMon
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

Private Sub txtNSol_GotFocus()
    Enfoque txtNSol
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

Private Sub txtPro_GotFocus()
    Enfoque txtPro
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

Private Sub txtSol_GotFocus()
    Enfoque txtSol
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
    Dim tV As Single, valor As Single
    Dim tD As Single, vDesc As Single
    Dim tI As Single, vIgv As Single
    
    With Flex1
        For I = 1 To Flex1.Rows - 1
            tV = Val(.TextMatrix(I, 11))
            valor = valor + tV
            tD = tV * Val(.TextMatrix(I, 9)) / 100
            vDesc = vDesc + tD
            tI = (tV - tD) * Val(.TextMatrix(I, 10)) / 100
            vIgv = vIgv + tI
        Next
    End With
    
    lblImp = Format(valor, "##,##0.00")
    lblDes = Format(vDesc, "##,##0.00")
    lblTot = Format(valor - vDesc, "#,##0.00")
    lblIgv = Format(vIgv, "#,##0.00")
    lblCom = Format((valor - vDesc) + vIgv, "#,##0.00")
End Sub

Function Tiene_Entregas() As Boolean
    Dim Adodc2 As ADODB.Recordset
    
    Set Adodc2 = New ADODB.Recordset
    
    Adodc2.Open "SELECT * FROM co_detordcompra WHERE oc_cnumord='" & lblNum & "' AND oc_ccodigo='" & _
        Flex1.TextMatrix(Flex1.Row, 0) & "' AND oc_ncanten>0", VGCNx, adOpenStatic
    Tiene_Entregas = False
    If Adodc2.RecordCount > 0 Then Tiene_Entregas = True
End Function
