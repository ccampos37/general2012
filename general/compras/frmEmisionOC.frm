VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Object = "{4D137D9C-00A6-4458-9B46-7E95DB76D55B}#9.0#0"; "TextFer.ocx"
Begin VB.Form frmEmisionOC 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Emisión de Orden de Compra"
   ClientHeight    =   6465
   ClientLeft      =   1125
   ClientTop       =   2790
   ClientWidth     =   9870
   ClipControls    =   0   'False
   Icon            =   "frmEmisionOC.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6465
   ScaleWidth      =   9870
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Fradatos 
      Height          =   2508
      Left            =   144
      TabIndex        =   33
      Top             =   570
      Width           =   9708
      Begin VB.TextBox txtNSol 
         Height          =   288
         Left            =   8640
         MaxLength       =   10
         TabIndex        =   2
         Top             =   240
         Width           =   945
      End
      Begin VB.TextBox txtObs 
         Height          =   288
         Left            =   1164
         TabIndex        =   10
         Top             =   2028
         Width           =   7500
      End
      Begin VB.TextBox txtCot 
         Height          =   336
         Left            =   6288
         TabIndex        =   7
         Top             =   948
         Width           =   3312
      End
      Begin VB.TextBox txtEntE 
         Height          =   288
         Left            =   1020
         MaxLength       =   50
         TabIndex        =   8
         Top             =   1308
         Width           =   5295
      End
      Begin ctrlayuda_f.Ctr_Ayuda CtrAyu_moneda 
         Height          =   348
         Left            =   6288
         TabIndex        =   5
         Top             =   576
         Width           =   3324
         _ExtentX        =   5874
         _ExtentY        =   609
         XcodMaxLongitud =   2
         xcodwith        =   200
         NomTabla        =   "gr_moneda"
         TituloAyuda     =   "Ayuda Monedas"
         ListaCampos     =   "monedacodigo(1),monedadescripcion(1)"
         XcodCampo       =   "monedacodigo"
         XListCampo      =   "monedadescripcion"
         ListaCamposDescrip=   "Codigo,Descripcion"
         ListaCamposText =   "monedacodigo,monedadescripcion"
      End
      Begin MSComCtl2.DTPicker txtEmi 
         Height          =   288
         Left            =   1008
         TabIndex        =   3
         Top             =   588
         Width           =   1212
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         Format          =   108855297
         CurrentDate     =   37015
      End
      Begin MSComCtl2.DTPicker txtEnt 
         Height          =   288
         Left            =   3648
         TabIndex        =   4
         Top             =   588
         Width           =   1212
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         Format          =   108855297
         CurrentDate     =   37015
      End
      Begin TextFer.TxFer lblRuc 
         Height          =   300
         Left            =   6240
         TabIndex        =   43
         Top             =   192
         Width           =   1308
         _ExtentX        =   2302
         _ExtentY        =   529
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
         MaxLength       =   11
         Locked          =   -1  'True
         Text            =   ""
         ColorIlumina    =   -2147483624
         SaltarAlEnter   =   -1  'True
         Valor           =   ""
         NoCaracteres    =   "0123456789"
         MarcarTextoAlEnfoque=   -1  'True
         NoRangoCadena   =   -1  'True
      End
      Begin ctrlayuda_f.Ctr_Ayuda CtrAyu_Proveedor 
         Height          =   312
         Left            =   1008
         TabIndex        =   1
         Top             =   192
         Width           =   4500
         _ExtentX        =   7938
         _ExtentY        =   556
         XcodMaxLongitud =   11
         xcodwith        =   1100
         NomTabla        =   "cp_proveedor"
         TituloAyuda     =   "Busqueda de Proveedor"
         ListaCampos     =   "clientecodigo(1),clienterazonsocial(1),clienteruc(1),clientetelefono(1)"
         XcodCampo       =   "clientecodigo"
         XListCampo      =   "clienterazonsocial"
         ListaCamposDescrip=   "Codigo,Descripcion"
         ListaCamposText =   "clientecodigo,clienterazonsocial,clienteruc,clientetelefono"
      End
      Begin ctrlayuda_f.Ctr_Ayuda CtrAyu_pago 
         Height          =   360
         Left            =   1008
         TabIndex        =   6
         Top             =   912
         Width           =   4116
         _ExtentX        =   7250
         _ExtentY        =   635
         XcodMaxLongitud =   3
         xcodwith        =   300
         NomTabla        =   "co_condicionespago"
         TituloAyuda     =   "Busqueda de Condiciones de Pago"
         ListaCampos     =   "pagocodigo(1),pagodescripcion(1)"
         XcodCampo       =   "pagocodigo"
         XListCampo      =   "Pagodescripcion"
         ListaCamposDescrip=   "Codigo,Descripcion"
         ListaCamposText =   "pagocodigo,pagodescripcion"
      End
      Begin ctrlayuda_f.Ctr_Ayuda CtrAyu_solicitante 
         Height          =   312
         Left            =   1008
         TabIndex        =   9
         Top             =   1632
         Width           =   4116
         _ExtentX        =   7250
         _ExtentY        =   556
         XcodMaxLongitud =   3
         xcodwith        =   300
         NomTabla        =   "co_solicitantes"
         TituloAyuda     =   "Busqueda de Solicitante"
         ListaCampos     =   "solicitantecodigo(1),solicitantenombre(1)"
         XcodCampo       =   "solicitantecodigo"
         XListCampo      =   "solicitantenombre"
         ListaCamposDescrip=   "Codigo,Descripcion"
         ListaCamposText =   "solicitantecodigo,solicitantenombre"
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Cond.Pago     :"
         Height          =   192
         Left            =   48
         TabIndex        =   45
         Top             =   996
         Width           =   1032
      End
      Begin VB.Label Le_Proveedor 
         Caption         =   "No. Requis."
         Height          =   252
         Left            =   7728
         TabIndex        =   44
         Top             =   288
         Width           =   1020
      End
      Begin VB.Label Label12 
         Caption         =   "Observación :"
         Height          =   252
         Left            =   84
         TabIndex        =   42
         Top             =   2040
         Width           =   1092
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Moneda  :"
         Height          =   192
         Left            =   5448
         TabIndex        =   41
         Top             =   600
         Width           =   720
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Entrega   :"
         Height          =   192
         Left            =   2808
         TabIndex        =   40
         Top             =   600
         Width           =   720
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "R.U.C.  :"
         Height          =   192
         Left            =   5616
         TabIndex        =   39
         Top             =   288
         Width           =   552
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Proveedor     :"
         Height          =   192
         Left            =   48
         TabIndex        =   38
         Top             =   276
         Width           =   1008
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Emisión         :"
         Height          =   192
         Left            =   84
         TabIndex        =   37
         Top             =   600
         Width           =   996
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Entregar en   :"
         Height          =   192
         Left            =   84
         TabIndex        =   36
         Top             =   1320
         Width           =   1008
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Solicitante     :"
         Height          =   192
         Left            =   84
         TabIndex        =   35
         Top             =   1680
         Width           =   1008
      End
      Begin VB.Label lblCen 
         AutoSize        =   -1  'True
         Caption         =   "Cotización  :"
         Height          =   192
         Left            =   5244
         TabIndex        =   34
         Top             =   960
         Width           =   876
      End
   End
   Begin VB.Frame Frame1 
      Height          =   636
      Left            =   144
      TabIndex        =   46
      Top             =   -48
      Width           =   9660
      Begin ctrlayuda_f.Ctr_Ayuda Ctrayu_tipoorden 
         Height          =   384
         Left            =   1056
         TabIndex        =   51
         Top             =   192
         Width           =   3396
         _ExtentX        =   5980
         _ExtentY        =   688
         XcodMaxLongitud =   11
         xcodwith        =   1100
         NomTabla        =   "co_tipodeorden"
         TituloAyuda     =   "Busqueda de Tipo de Orden"
         ListaCampos     =   "tipoordencodigo(1),tipoordendescripcion(1),tipoordennumeracion(2)"
         XcodCampo       =   "tipoordencodigo"
         XListCampo      =   "tipoordendescripcion"
         ListaCamposDescrip=   "Codigo,Descripcion"
         ListaCamposText =   "tipoordencodigo,tipoordendescripcion,tipoordennumeracion"
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         Caption         =   "Tipo Orden     :"
         Height          =   192
         Left            =   96
         TabIndex        =   52
         Top             =   276
         Width           =   1032
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
         Height          =   348
         Left            =   5340
         TabIndex        =   50
         Top             =   192
         Width           =   1560
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Número  :"
         Height          =   192
         Left            =   4656
         TabIndex        =   49
         Top             =   288
         Width           =   696
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
         Height          =   348
         Left            =   7728
         TabIndex        =   48
         Top             =   204
         Width           =   1644
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Estado  :"
         Height          =   192
         Left            =   7080
         TabIndex        =   47
         Top             =   288
         Width           =   636
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
      Left            =   1152
      Picture         =   "frmEmisionOC.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5616
      Visible         =   0   'False
      Width           =   775
   End
   Begin VB.CommandButton cmdEli2 
      Caption         =   "&Quitar"
      Enabled         =   0   'False
      Height          =   675
      Left            =   4080
      Picture         =   "frmEmisionOC.frx":0D0C
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5616
      Visible         =   0   'False
      Width           =   775
   End
   Begin VB.CommandButton cmdEdi2 
      Caption         =   "&Editar"
      Enabled         =   0   'False
      Height          =   675
      Left            =   2736
      Picture         =   "frmEmisionOC.frx":114E
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5616
      Visible         =   0   'False
      Width           =   775
   End
   Begin VB.CommandButton cmdImp 
      Caption         =   "&Imprimir"
      Height          =   675
      Left            =   5520
      Picture         =   "frmEmisionOC.frx":1590
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   3840
      Visible         =   0   'False
      Width           =   775
   End
   Begin VB.CommandButton cmdEdi 
      Caption         =   "&Editar"
      Height          =   675
      Left            =   2910
      Picture         =   "frmEmisionOC.frx":19D2
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   3825
      Width           =   775
   End
   Begin VB.CommandButton CmdEli 
      Caption         =   "&Anular"
      Height          =   675
      Left            =   4230
      Picture         =   "frmEmisionOC.frx":1E14
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   3840
      Width           =   775
   End
   Begin VB.CommandButton cmdNue 
      Caption         =   "&Nuevo"
      Height          =   675
      Left            =   1575
      Picture         =   "frmEmisionOC.frx":2256
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3810
      Width           =   775
   End
   Begin VB.CommandButton CmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   675
      Left            =   6840
      Picture         =   "frmEmisionOC.frx":2698
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   3825
      Width           =   775
   End
   Begin VB.CommandButton cmdGra 
      Caption         =   "&Grabar"
      Height          =   675
      Left            =   5400
      Picture         =   "frmEmisionOC.frx":2ADA
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5616
      Visible         =   0   'False
      Width           =   775
   End
   Begin VB.CommandButton CmdSalir2 
      Caption         =   "&Salir"
      Height          =   675
      Left            =   6864
      Picture         =   "frmEmisionOC.frx":2F1C
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   5616
      Visible         =   0   'False
      Width           =   775
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Flex1 
      Height          =   1515
      Left            =   120
      TabIndex        =   21
      Top             =   3120
      Visible         =   0   'False
      Width           =   9735
      _ExtentX        =   17171
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
   Begin VB.Frame fraTotales 
      Height          =   975
      Left            =   135
      TabIndex        =   22
      Top             =   4560
      Visible         =   0   'False
      Width           =   9708
      Begin VB.Label lblCom 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "892,760.00"
         Height          =   285
         Left            =   7080
         TabIndex        =   32
         Top             =   600
         Width           =   1110
      End
      Begin VB.Label lblIgv 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "892,760.00"
         Height          =   285
         Left            =   7080
         TabIndex        =   31
         Top             =   240
         Width           =   1110
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Compra :"
         Height          =   195
         Left            =   6360
         TabIndex        =   30
         Top             =   600
         Width           =   630
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "I.G.V.   :"
         Height          =   195
         Left            =   6360
         TabIndex        =   29
         Top             =   240
         Width           =   585
      End
      Begin VB.Label lblTot 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "892,760.00"
         Height          =   285
         Left            =   4200
         TabIndex        =   28
         Top             =   240
         Width           =   1230
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Total  :"
         Height          =   195
         Left            =   3600
         TabIndex        =   27
         Top             =   240
         Width           =   495
      End
      Begin VB.Label lblDes 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "892,760.00"
         Height          =   285
         Left            =   1680
         TabIndex        =   26
         Top             =   600
         Width           =   1110
      End
      Begin VB.Label lblImp 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "892,760.00"
         Height          =   285
         Left            =   1680
         TabIndex        =   25
         Top             =   240
         Width           =   1110
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Descuento :"
         Height          =   195
         Left            =   720
         TabIndex        =   24
         Top             =   600
         Width           =   870
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Importe      :"
         Height          =   195
         Left            =   720
         TabIndex        =   23
         Top             =   240
         Width           =   840
      End
   End
   Begin Crystal.CrystalReport CrystalReport2 
      Bindings        =   "frmEmisionOC.frx":335E
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
      Left            =   30
      TabIndex        =   16
      Top             =   720
      Width           =   9480
      _ExtentX        =   16722
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
      Caption         =   "Ordenes pendientes por atender"
      ColumnCount     =   8
      BeginProperty Column00 
         DataField       =   "tipoordencodigo"
         Caption         =   "T.Orden"
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
      BeginProperty Column02 
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
      BeginProperty Column03 
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
      BeginProperty Column04 
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
      BeginProperty Column05 
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
      BeginProperty Column06 
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
      BeginProperty Column07 
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
            ColumnWidth     =   764.787
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            Locked          =   -1  'True
            WrapText        =   -1  'True
            ColumnWidth     =   1365.165
         EndProperty
         BeginProperty Column02 
            Locked          =   -1  'True
            WrapText        =   -1  'True
            ColumnWidth     =   3105.071
         EndProperty
         BeginProperty Column03 
            Alignment       =   2
            Locked          =   -1  'True
            WrapText        =   -1  'True
            ColumnWidth     =   1110.047
         EndProperty
         BeginProperty Column04 
            Alignment       =   2
            Locked          =   -1  'True
            WrapText        =   -1  'True
            ColumnWidth     =   434.835
         EndProperty
         BeginProperty Column05 
            Alignment       =   1
            Locked          =   -1  'True
            WrapText        =   -1  'True
            ColumnWidth     =   1019.906
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1379.906
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   14.74
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmEmisionOC"
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


Sub OculObj02(ntipo As Boolean)
    cmdGra.Visible = ntipo
    CmdSalir2.Visible = ntipo
End Sub

Sub OculObj03(ntipo As Boolean)
    Fradatos.Visible = ntipo
    fraTotales.Visible = ntipo
End Sub

Sub OculObj04(ntipo As Boolean)
    cmdNue.Visible = ntipo
    cmdEdi.Visible = ntipo
    CmdEli.Visible = ntipo
    cmdImp.Visible = ntipo
    CmdSalir.Visible = ntipo
End Sub

Sub OculObj06(ntipo As Boolean)
    DataGrid1.Visible = ntipo
End Sub

Sub Abre_Tabla_OCs()
    Dim strsql As String
    
    Set Adodc1 = New ADODB.Recordset
    
    strsql = "SELECT * FROM co_cabordcompra,co_estadoorden WHERE co_cabordcompra.oc_situacionorden =co_estadoorden." & _
        "estadooccodigo and estadoocatendido<>1 ORDER BY oc_cnumord "
    Adodc1.Open strsql, VGCNx, adOpenDynamic, adLockOptimistic
    
    Set DataGrid1.DataSource = Adodc1
    
End Sub

Private Sub cmdEdi2_Click()
On Error GoTo err
    With frmemisionOCdetalle
        .activado = False
        .CtrAyu_articulo.xclave = Flex1.TextMatrix(Flex1.Row, 0)
        .lblFab = Flex1.TextMatrix(Flex1.Row, 1)
        .CtrAyu_articulo.xnombre = Flex1.TextMatrix(Flex1.Row, 2)
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
        .CtrAyu_articulo.Enabled = False
        .activado = True
        .Calculo_Automatico
        .Show 1
        
        If Not .cancelado Then
            If .tipo = "S" Then
              .txtCan = 1
            End If
            Flex1.TextMatrix(Flex1.Row, 2) = .CtrAyu_articulo.xnombre
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
err:
    MsgBox err.Description
 
End Sub

Private Sub CmdEli_Click()
    On Error GoTo EliErr
    
    If Adodc1("oc_estadoorden") = 1 Or Adodc1("oc_situacionorden") <> "0" Then
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
        VGCNx.BeginTrans
        
        strsql = "UPDATE co_detordcompra SET oc_situacionorden=2  WHERE oc_cnumord='" & voc & "'"
        VGCNx.Execute strsql
        strsql = "UPDATE co_cabordcompra SET oc_estadoorden=1 WHERE oc_cnumord='" & voc & "'"
        VGCNx.Execute strsql

        VGCNx.CommitTrans
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

    Mensaje = "¿Desea eliminar el documento " & Adodc1("nrorequi") & "?"
    If MsgBox(Mensaje, vbQuestion + vbYesNo, "Mensaje") = vbYes Then
        strsql = "DELETE * FROM requisd WHERE nrorequi='" & Adodc1("nrorequi") & "'"
        
        nTra = 1
        VGCNx.BeginTrans
        VGCNx.Execute strsql
        VGCNx.CommitTrans
        nTra = 0
        If nTra = 0 Then
            Adodc1.Delete
            Adodc1.Update
        End If
        Estado_Botones
            
    End If
    If Adodc1.RecordCount > 0 Then
        DataGrid1.SetFocus
    Else
        cmdNue.SetFocus
    End If
    Exit Sub

EliErr:
    MsgBox err.Description
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
            Dim i As Integer
            
            For i = 0 To 13
                Flex1.TextMatrix(1, i) = ""
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
    Dim rs2 As New ADODB.Recordset
    Dim i As Integer
    Dim vFactor As Single, vCantid As Single
    Dim vPreuni As Single, vDscpor As Single
    Dim vDescto As Single, vIgv As Single
    Dim vIgvpor As Single, vPrenet As Single
    Dim vTotven As Single, vTotnet As Single
    Dim vURef As String, txtmon As String
    Dim txtEst As String, txttip As Integer
    Dim txtpro As String, txtsol As String
    Dim lblpro As String, txtFor As String
    On Error GoTo GrabErr
    
    txttip = 0
    txtFor = Trim(CtrAyu_pago.xclave)
    
    If Trim(Ctrayu_tipoorden.xclave) = "" Then
       Mensaje = "Debe ingresar Código de Tipo de Orden"
       MsgBox Mensaje, vbExclamation, "Mensaje"
       Ctrayu_tipoorden.SetFocus
       Exit Sub
    End If
    
    txtpro = Trim(CtrAyu_Proveedor.xclave)
    If txtpro = "" Then
       Mensaje = "Debe ingresar Código de Proveedor"
       MsgBox Mensaje, vbExclamation, "Mensaje"
       CtrAyu_Proveedor.SetFocus
       Exit Sub
    End If
    
    If txtEmi > txtEnt Then
       MsgBox "Fecha de emision no debe ser mayor a la fecha de entrega", vbExclamation, "Error"
       Exit Sub
       txtEmi.SetFocus
    End If
       
    txtmon = CtrAyu_moneda.xclave
    If Trim(txtmon) = "" Then
        Mensaje = "Debe ingresar el Tipo de Moneda"
        MsgBox Mensaje, vbExclamation, "Error"
        CtrAyu_moneda.SetFocus
        Exit Sub
    End If
    
    txtEst = ""
    txtsol = Trim(CtrAyu_solicitante.xclave)
    If txtsol = "" Then
        Mensaje = "Debe ingresar Solicitante"
        MsgBox Mensaje, vbExclamation, "Mensaje"
        CtrAyu_solicitante.SetFocus
        Exit Sub
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
       unum = Format(Val(lblNum), "00000000000")

       If nT = 1 Then      'Ingreso
         'unum = Format(Devolver_Dato(1, , " & trim(ctrayu_tipoordencodigo) & ", "tipoordencodigo", False,
         '      "ctnnumero"), "00000000000")
         SQLc = "select tipoordennumeracion from co_tipodeorden where tipoordencodigo='" & Trim(Ctrayu_tipoorden.xclave) & "' "
         Set rs2 = New ADODB.Recordset
         rs2.Open SQLc, VGCnxCT, adOpenKeyset, adLockReadOnly
         unum = rs2!tipoordennumeracion + 1
          
          SQLc = "UPDATE co_tipodeorden SET tipoordennumeracion=" & unum & _
                " WHERE tipoordencodigo='" & Trim(Ctrayu_tipoorden.xclave) & "' "
            VGCNx.Execute SQLc
           unum = Format(Val(unum), "00000000000")
           lblNum = unum
            SQLc = "INSERT INTO co_cabordcompra (tipoordencodigo,oc_cnumord,oc_dfecdoc,oc_ccodpro," & _
                "oc_crazsoc,oc_ccotiza,oc_ccodmon,oc_cforpag,oc_dfecent," & _
                "oc_cobserv,oc_csolict,oc_centreg,oc_estadoorden,oc_situacionorden,oc_nimport,oc_ndescue," & _
                "oc_nigv,oc_nventa,oc_dfecact,oc_chora,oc_cusuari,oc_cconver) VALUES ('" & _
                Ctrayu_tipoorden.xclave & "','" & lblNum & "','" & txtEmi & "','" & txtpro & "','" & _
                CtrAyu_Proveedor.xnombre & "','" & txtCot & "','" & txtmon & "','" & txtFor & "','" & _
                txtEnt & "','" & _
                SupCadSQL(txtObs) & "','" & txtsol & "','" & txtEntE & "',' ','0'," & _
                CDbl(lblImp) & "," & CDbl(lblDes) & "," & CDbl(lblIgv) & "," & CDbl(lblCom) & _
                ",'" & VGParamSistem.FechaTrabajo & "','" & Format(Time, "hh.mm.ss") & "','" & VGUsuario & _
                "','" & txtEst & "')"
            VGCNx.Execute SQLc
            
            For i = 1 To Flex1.Rows - 1
                vFactor = Val(Flex1.TextMatrix(i, 6))
                vCantid = Val(Flex1.TextMatrix(i, 4))
                If vCantid = 0 Then
                   vCantid = 1
                End If
                vPreuni = Val(Flex1.TextMatrix(i, 7))
                vDscpor = Val(Flex1.TextMatrix(i, 9))
                vDescto = IIf(vFactor > 0, vFactor / vCantid, 1) * vPreuni * vCantid * _
                    vDscpor / 100
                vIgvpor = Val(Flex1.TextMatrix(i, 10))
                vTotven = Val(Flex1.TextMatrix(i, 11))
                vIgv = (vTotven - vDescto) * vIgvpor / 100
                vPrenet = Val(Flex1.TextMatrix(i, 8)) * (1 - vDscpor / 100)
                SQLd = "INSERT INTO co_detordcompra (tipoordencodigo,oc_cnumord,oc_ccodpro,oc_dfecdoc,oc_citem," & _
                  "oc_ccodigo,oc_ccodref,oc_cdesref,oc_cunidad,oc_cuniref,oc_nfactor," & _
                  "oc_ncantid,oc_nsaldo,oc_npreuni,oc_ndscpor,oc_ndescto,oc_nigv,oc_nigvpor," & _
                  "oc_nprenet,oc_ntotven,oc_ntotnet,oc_situacionorden,ord_fabnum,oc_ccomen1, tipoarticulocodigo) " & _
                  "VALUES ('" & Ctrayu_tipoorden.xclave & "','" & lblNum & "','" & txtpro & "','" & txtEmi _
                  & "','" & Format(i, "000") & "','" & _
                  Flex1.TextMatrix(i, 0) & "','" & Flex1.TextMatrix(i, 1) & "','" & _
                  Flex1.TextMatrix(i, 2) & "','" & Flex1.TextMatrix(i, 3) & "','" & _
                  Flex1.TextMatrix(i, 5) & "'," & vFactor & "," & vCantid & "," & vCantid & "," & _
                  vPreuni & "," & vDscpor & "," & vDescto & "," & vIgv & "," & _
                  vIgvpor & "," & vPrenet & "," & vTotven & "," & vTotven - vDescto + _
                  vIgv & ",'0','" & Flex1.TextMatrix(i, 12) & "','" & _
                  Flex1.TextMatrix(i, 13) & "','" & Flex1.TextMatrix(i, 14) & "')"
                VGCNx.Execute SQLd
                
                SQLd = "UPDATE maeart SET aprecom=" & Val(Flex1.TextMatrix(i, 8)) & _
                    ",acodpro='" & txtpro & "',afecven='" & txtEmi _
                    & "' WHERE acodigo='" & Flex1.TextMatrix(i, 0) & "'"
                VGCNx.Execute SQLd
            Next
        ElseIf nT = 2 Then     'Modificar
            SQLc = "UPDATE co_cabordcompra SET oc_dfecdoc='" & txtEmi & _
                "',oc_ccotiza='" & txtCot & "',oc_ccodmon='" & txtmon & "',oc_cforpag='" & _
                txtFor & "',oc_ntipcam=" & Val(txttip) & ",oc_dfecent='" & _
                txtEnt & "',oc_cobserv='" & SupCadSQL(txtObs) & _
                "',oc_csolict='" & txtsol & "',oc_centreg='" & txtEntE & "',oc_nimport=" & _
                CDbl(lblImp) & ",oc_ndescue=" & CDbl(lblDes) & ",oc_nigv=" & CDbl(lblIgv) & _
                ",oc_nventa=" & CDbl(lblCom) & ",oc_dfecact='" & _
                VGParamSistem.FechaTrabajo & "',oc_chora='" & Format(Time, "hh.mm.ss") & "',oc_cusuari='" & _
                VGUsuario & "',oc_cconver='" & txtEst & "' WHERE oc_cnumord='" & lblNum & "'"
            VGCNx.Execute SQLc
            
            SQLd = "DELETE * FROM co_detordcompra WHERE oc_cnumord='" & lblNum & "'"
            VGCNx.Execute SQLd
            
            For i = 1 To Flex1.Rows - 1
                vURef = ""
                vFactor = 0
                If Flex1.TextMatrix(i, 3) <> Flex1.TextMatrix(i, 5) Then
                    vURef = Flex1.TextMatrix(i, 5)
                    vFactor = Val(Flex1.TextMatrix(i, 6))
                End If
                vCantid = Val(Flex1.TextMatrix(i, 4))
                vPreuni = Val(Flex1.TextMatrix(i, 7))
                vDscpor = Val(Flex1.TextMatrix(i, 9))
                vDescto = IIf(vFactor > 0, vFactor / vCantid, 1) * vPreuni * vCantid * _
                    vDscpor / 100
                vIgvpor = Val(Flex1.TextMatrix(i, 10))
                vTotven = Val(Flex1.TextMatrix(i, 11))
                vIgv = (vTotven - vDescto) * vIgvpor / 100
                vPrenet = Val(Flex1.TextMatrix(i, 8)) * (1 - vDscpor / 100)
                SQLd = "INSERT INTO co_detordcompra (oc_cnumord,oc_ccodpro,oc_dfecdoc,oc_citem," & _
                    "oc_ccodigo,oc_ccodref,oc_cdesref,oc_cunidad,oc_cuniref,oc_nfactor," & _
                    "oc_ncantid,oc_npreuni,oc_ndscpor,oc_ndescto,oc_nigv,oc_nigvpor," & _
                    "oc_nprenet,oc_ntotven,oc_ntotnet,oc_situacionorden,ord_fabnum,oc_ccomen1,tipoarticulocodigo) " & _
                    "VALUES ('" & lblNum & "','" & txtpro & "','" & txtEmi _
                    & "','" & Format(i, "000") & "','" & _
                    Flex1.TextMatrix(i, 0) & "','" & Flex1.TextMatrix(i, 1) & "','" & _
                    Flex1.TextMatrix(i, 2) & "','" & Flex1.TextMatrix(i, 3) & "','" & _
                    vURef & "'," & vFactor & "," & vCantid & "," & _
                    vPreuni & "," & vDscpor & "," & vDescto & "," & vIgv & "," & _
                    vIgvpor & "," & vPrenet & "," & vTotven & "," & vTotven - vDescto + _
                    vIgv & ",'0','" & Flex1.TextMatrix(i, 12) & "','" & _
                    Flex1.TextMatrix(i, 13) & "', '" & Flex1.TextMatrix(i, 14) & "')"
                VGCNx.Execute SQLd
                
                SQLd = "UPDATE maeart SET aprecom=" & Val(Flex1.TextMatrix(i, 8)) & _
                    ",acodpro='" & txtpro & "',afecven='" & txtEmi _
                    & "' WHERE acodigo='" & Flex1.TextMatrix(i, 0) & "'"
                VGCNx.Execute SQLd
            Next
        End If
        
        VGCNx.CommitTrans
        nTra = 0
        Adodc1.Requery
        Adodc1.Find "oc_cnumord='" & lblNum & "'"
        
        If nT = 1 Then
            unum = Format(Val(unum) + 1, "00000000000")
            lblNum = unum
            Limpiar
            Vacia_FlexGrid
            Estado_Items
            Calcula_Totales
            txtEmi = VGParamSistem.FechaTrabajo
            txtEnt = VGParamSistem.FechaTrabajo
            txttip = "0.000"
                        
        Else
            CmdSalir2_Click
        End If
    Exit Sub
End If
GrabErr:
    MsgBox err.Description
    If nTra = 1 Then VGCNx.RollbackTrans
End Sub

Private Sub cmdImp_Click()
Dim formulas(3) As String
unum = Adodc1("oc_cnumord")
CrystalReport2.Reset
CrystalReport2.WindowTitle = "rptcoordencompra -- orden de compra"
   CrystalReport2.ReportFileName = VGParamSistem.RutaReport & "\" & VGParamSistem.carpetareportes & "\" & "rptcoordencompra.rpt"
    CrystalReport2.DiscardSavedData = True
       
    CrystalReport2.LogOnServer "pdssql.dll", _
                                VGParamSistem.ServidorGEN, _
                                VGParamSistem.BDEmpresaGEN, _
                                VGParamSistem.UsuarioGEN, _
                                ""
    CrystalReport2.Connect = VGCadenaReport2
       
    CrystalReport2.Destination = crptToWindow
    CrystalReport2.WindowState = crptMaximized
    CrystalReport2.WindowShowPrintBtn = True
    CrystalReport2.WindowShowRefreshBtn = True
    CrystalReport2.WindowShowSearchBtn = True
    CrystalReport2.WindowShowPrintSetupBtn = True
    CrystalReport2.formulas(0) = "@emp ='" & VGParametros.NomEmpresa & "'"
    CrystalReport2.formulas(1) = "@ruc ='" & VGParametros.RucEmpresa & "'"
    CrystalReport2.formulas(2) = "@direccion ='" & VGParametros.direccionempresa & "'"
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
    lblImp = "0.00": lblTot = "0.00": lblIgv = "0.00"
    lblDes = "0.00": lblCom = "0.00"
    Frame1.Visible = True
    Fradatos.Visible = True
    Fradatos.Enabled = True
    cmdGra.Enabled = True
    CmdSalir2.Cancel = True
End Sub

Private Sub cmdEdi_Click()
    If Adodc1("oc_estadoorden") = "A" Then
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
        Fradatos.Enabled = False
    Else
        nT = 2
        OculObj06 False
        OculObj04 False
        OculObj02 True
        Mostrar Adodc1("oc_cnumord")
        OculObj03 True
        Proceso True
        Fradatos.Enabled = True
        
        cmdEdi2.Enabled = True
        cmdEli2.Enabled = True
        cmdGra.Enabled = True
        
        txtEmi.SetFocus
        CmdSalir2.Cancel = True
    End If
End Sub

Private Sub cmdNue2_Click()
    With frmemisionOCdetalle
        .activado = False
        .CtrAyu_articulo.xclave = ""
        .txtCan = "0.00"
        .txtPUn = "0.00"
        .txtPDe = "0.00"
        .txtPIg = "19.00"
        .txtordfab = ""
        .lblFab.Caption = ""
        .txtCo1 = ""
        .activado = True
       .Show 1
        
        If Not .cancelado Then
           If .tipo = "S" Then
              .txtCan = 1
            End If
            
            If Flex1.Rows - 1 = 1 Then
                If Flex1.TextMatrix(1, 0) = "" Then
                    Flex1.AddItem Trim(.CtrAyu_articulo.xclave) & vbTab & .lblFab.Caption & vbTab & Trim(.CtrAyu_articulo.xnombre) & vbTab & _
                        .lblUni & vbTab & .txtCan & vbTab & IIf(.txtURe = "", .lblUni, _
                        .txtURe) & vbTab & IIf(.txtURe = "", .txtCan, .txtRef) & vbTab & _
                        .txtPUn & vbTab & Format(Val(.lblPNe) + Val(.lblDes), "0.00") & _
                        vbTab & .txtPDe & vbTab & .txtPIg & vbTab & Format(IIf(.txtURe = "", .txtCan, .txtRef) * _
                        (Val(.lblPNe) + Val(.lblDes)), "0.00") & vbTab & .txtordfab & vbTab & _
                        .txtCo1 & vbTab & .tipo, 1
                    Flex1.Rows = 2
                Else
                    Flex1.AddItem Trim(.CtrAyu_articulo.xclave) & vbTab & .lblFab & vbTab & Trim(.CtrAyu_articulo.xnombre) & vbTab & _
                        .lblUni & vbTab & .txtCan & vbTab & IIf(.txtURe = "", .lblUni, _
                        .txtURe) & vbTab & IIf(.txtURe = "", .txtCan, .txtRef) & vbTab & _
                        .txtPUn & vbTab & Format(Val(.lblPNe) + Val(.lblDes), "0.00") & _
                        vbTab & .txtPDe & vbTab & .txtPIg & vbTab & Format(IIf(.txtURe = "", .txtCan, .txtRef) * _
                        (Val(.lblPNe) + Val(.lblDes)), "0.00") & vbTab & .txtordfab & vbTab & _
                        .txtCo1 & vbTab & .tipo
                    Flex1.Row = Flex1.Rows - 1
                End If
            Else
                Flex1.AddItem Trim(.CtrAyu_articulo.xclave) & vbTab & .lblFab & vbTab & Trim(.CtrAyu_articulo.xnombre) & vbTab & _
                    .lblUni & vbTab & .txtCan & vbTab & IIf(.txtURe = "", .lblUni, _
                    .txtURe) & vbTab & IIf(.txtURe = "", .txtCan, .txtRef) & vbTab & _
                    .txtPUn & vbTab & Format(Val(.lblPNe) + Val(.lblDes), "0.00") & _
                    vbTab & .txtPDe & vbTab & .txtPIg & vbTab & Format(IIf(.txtURe = "", .txtCan, .txtRef) * _
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
    Unload frmemisionOCdetalle
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
    Frame1.Visible = False
    If Adodc1.RecordCount > 0 Then
        DataGrid1.SetFocus
    Else
        cmdNue.SetFocus
    End If
    CmdSalir.Cancel = True
End Sub
Public Function SupCadSQL(S As String) As String
 Dim Aux As String
 If Not IsNull(S) Then
     Aux = Replace(S, "'", "''")
 End If
 SupCadSQL = Aux
 
End Function

Private Sub CtrAyu_tipoorden_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
    Dim unum As String
    Set VGvardllgen = New dllgeneral.dll_general
    unum = VGvardllgen.ESNULO(ColecCampos("tipoordennumeracion").Value, "")
    unum = Format(Val(unum) + 1, "00000000000")
    lblNum = unum
    
End Sub


Private Sub CtrAyu_Proveedor_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
    Set VGvardllgen = New dllgeneral.dll_general
    lblRuc.Text = VGvardllgen.ESNULO(ColecCampos("clienteruc").Value, "")
End Sub
Private Sub CtrAyu_Proveedor_AlNoDevolverNada()
    lblRuc.Text = ""
End Sub

Private Sub Form_Load()
    Formato_FlexGrid
    Call CtrAyu_moneda.conexion(VGCnxCT): CtrAyu_moneda.Filtro = "(monedacodigo <>'00') "
    Call Ctrayu_tipoorden.conexion(VGCNx)
    Call CtrAyu_Proveedor.conexion(VGCNx): CtrAyu_Proveedor.Filtro = "(clientecodigo <>'00') "
    Call CtrAyu_pago.conexion(VGCNx)
    Call CtrAyu_solicitante.conexion(VGCNx)
    
    OculObj02 False
    OculObj03 False
    OculObj04 True
    OculObj06 True
    
    unum = ""
    Abre_Tabla_OCs
    Estado_Botones
    Frame1.Visible = False
    Load frmemisionOCdetalle
End Sub
Private Sub Reales_Positivos(k As Integer, t As TextBox)
Dim t1 As String
    k = Asc(UCase(Chr(k)))
    If k = 8 Then Exit Sub
    If k <> 45 And k <> 44 And k <> 32 And k <> 69 And k <> 43 Then
        t1 = Left(t, t.SelStart)
        t1 = t1 & Chr(k) & Right(t, Len(t) - Len(t1))
        If IsNumeric(t1) Then Exit Sub
    End If
    k = 0
    
End Sub

Public Function Existe(tipo As Integer, Cod As String, Tabla As String, Campo As String, Fecha As Boolean, Optional Cod2 As String, Optional cCampo2 As String, Optional Cod3 As String, Optional cCampo3 As String, Optional Cod4 As Boolean, Optional cCampo4 As String, Optional Cod5 As String, Optional cCampo5 As String) As Boolean
Dim cSel1 As ADODB.Recordset, cSL As String
Set cSel1 = New ADODB.Recordset

 If Fecha Then
        cSL = "Select * from " & Tabla & "  Where " & Campo & " =  '" & Cod & "'"
 Else
       If UCase(Tabla) = "PUNTO_VENTA" Then
                cSL = "Select * from " & Tabla & "  Where " & Campo & " =  '" & Cod & "'"
       Else
                cSL = "Select * from " & Tabla & "  Where " & Campo & " =  '" & Cod & "'"
       End If
       If Trim(Cod2) <> "" Then
            cSL = cSL & " And  " & cCampo2 & " =  '" & SupCadSQL(Cod2) & "'"
       End If
       If Trim(Cod3) <> "" Then
            cSL = cSL & " And  " & cCampo3 & " =  '" & SupCadSQL(Cod3) & "'"
       End If
       If Trim(cCampo4) <> "" Then
            If Cod4 = True Then
                cSL = cSL & " And  " & cCampo4
            Else
                cSL = cSL & " And  " & Not cCampo4
            End If
        End If
        If Trim(Cod5) <> "" Then
            cSL = cSL & " And  " & cCampo5 & " =  '" & Cod5 & "'"
        End If
 End If
 
Select Case tipo
Case 1 'Bd. Comun
            cSel1.Open cSL, VGCNx, adOpenStatic
Case 2 'Bd. Config
            cSel1.Open cSL, VGCNx, adOpenStatic
Case 3 'Bd. Contab
            cSel1.Open cSL, VGCnxCT, adOpenStatic
End Select

If cSel1.RecordCount > 0 Then
     Existe = True
Else
     Existe = False
End If
'csel1.Close
End Function

Sub Limpiar()

txtNSol = ""
txtCot = ""
Ctrayu_tipoorden.xclave = ""
Ctrayu_tipoorden.xnombre = ""
CtrAyu_Proveedor.xclave = ""
CtrAyu_Proveedor.xnombre = ""
CtrAyu_pago.xclave = ""
CtrAyu_pago.xnombre = ""
CtrAyu_solicitante.xclave = ""
CtrAyu_solicitante.xnombre = ""
CtrAyu_moneda.xclave = ""
CtrAyu_moneda.xnombre = ""
txtEntE = "": txtObs = ""
End Sub

Sub Mostrar(cC1 As String)
    Dim cSqlM As String, cSelM As ADODB.Recordset
    Dim k As Integer, i As Integer, vd As String
    Dim vpu As Single, txtpro As String
    Dim txtsol As String
    
    lblNum = cC1
   ' lblEst = Adodc1("est_nombre")
    CtrAyu_Proveedor.xclave = Adodc1("oc_ccodpro")
    txtpro = CtrAyu_Proveedor.xclave
    CtrAyu_Proveedor.xnombre = Devolver_Dato(1, txtpro, "cp_proveedor", "clientecodigo", False, "clienterazonsocial")
 '   lblRuc = Devolver_Dato(1, txtpro, "cp_proveedor", "clientecodigo", False, "clienteruc")
    txtEmi = Adodc1("oc_dfecdoc")
    txtEnt = Adodc1("oc_dfecent")
    CtrAyu_moneda.xclave = Adodc1("oc_ccodmon")
    CtrAyu_pago.xclave = Adodc1("oc_cforpag")
    txtCot = Adodc1("oc_ccotiza")
    txtEntE = Adodc1("oc_centreg")
    CtrAyu_solicitante.xclave = Adodc1("oc_csolict")
    txtsol = CtrAyu_solicitante.xclave
    CtrAyu_solicitante.xnombre = Devolver_Dato(1, txtsol, "co_solicitantes", "solicitantecodigo", False, "solicitantenombre")
    txtObs = Adodc1("oc_cobserv")
    
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
    If Adodc1.RecordCount > 0 Then
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

Function ValidFecha(vText As String) As String
Dim cTxtNew As String, ncnt As Integer
Dim cTxt As String, cTxtDig As String

cTxtDig = "": cTxtNew = ""
For ncnt = 1 To Len(vText)
      cTxt = Mid(vText, ncnt, 1)
      If cTxt = "/" Then
         cTxtNew = cTxtNew & Str(Val(cTxtDig)) & "/"
         cTxtDig = ""
      Else
         If cTxt <> "_" Then cTxtDig = cTxtDig & cTxt
      End If
Next
If cTxtDig <> "" Then cTxtNew = cTxtNew & Str(Val(cTxtDig))

If IsDate(cTxtNew) Then
   ValidFecha = Format(CDate(cTxtNew), "dd/mm/yyyy")
End If
End Function


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
        End If
    End If
End Sub

Private Sub txtEntE_GotFocus()
    Enfoque txtEntE
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


Private Sub Enfoque(OBJ As Object)
  OBJ.SelStart = 0
  OBJ.SelLength = Len(OBJ)
End Sub

Sub Proceso(Estado As Boolean)
    Flex1.Visible = Estado
    cmdNue2.Visible = Estado
    cmdEdi2.Visible = Estado
    cmdEli2.Visible = Estado
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
    Dim i As Integer
    
    Do While Flex1.Rows - 1 > 1
        Flex1.RemoveItem 1
    Loop
    
    For i = 0 To 14
        Flex1.TextMatrix(1, i) = ""
    Next
End Sub

Sub Calcula_Totales()
    Dim i As Integer
    Dim tV As Single, valor As Single
    Dim tD As Single, vDesc As Single
    Dim tI As Single, vIgv As Single
    
    With Flex1
        For i = 1 To Flex1.Rows - 1
            tV = Val(.TextMatrix(i, 11))
            valor = valor + tV
            tD = tV * Val(.TextMatrix(i, 9)) / 100
            vDesc = vDesc + tD
            tI = (tV - tD) * Val(.TextMatrix(i, 10)) / 100
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
