VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Object = "{4D137D9C-00A6-4458-9B46-7E95DB76D55B}#9.0#0"; "textfer.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmTraEmi 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Emisi�n de Orden de Compra"
   ClientHeight    =   6468
   ClientLeft      =   1128
   ClientTop       =   2832
   ClientWidth     =   9864
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmTraEmi.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6468
   ScaleWidth      =   9864
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2508
      Left            =   144
      TabIndex        =   28
      Top             =   576
      Width           =   9708
      Begin VB.TextBox txtNSol 
         Height          =   288
         Left            =   8640
         MaxLength       =   10
         TabIndex        =   47
         Top             =   240
         Width           =   945
      End
      Begin VB.TextBox txtObs 
         Height          =   288
         Left            =   1164
         TabIndex        =   31
         Top             =   2028
         Width           =   7500
      End
      Begin VB.TextBox txtCot 
         Height          =   288
         Left            =   6288
         TabIndex        =   30
         Top             =   948
         Width           =   3312
      End
      Begin VB.TextBox txtEntE 
         Height          =   288
         Left            =   1020
         MaxLength       =   50
         TabIndex        =   29
         Top             =   1308
         Width           =   5295
      End
      Begin ctrlayuda_f.Ctr_Ayuda CtrAyu_moneda 
         Height          =   348
         Left            =   6288
         TabIndex        =   32
         Top             =   576
         Width           =   3324
         _ExtentX        =   5863
         _ExtentY        =   614
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
         TabIndex        =   33
         Top             =   588
         Width           =   1212
         _ExtentX        =   2138
         _ExtentY        =   508
         _Version        =   393216
         Format          =   19791873
         CurrentDate     =   37015
      End
      Begin MSComCtl2.DTPicker txtEnt 
         Height          =   288
         Left            =   3648
         TabIndex        =   34
         Top             =   588
         Width           =   1212
         _ExtentX        =   2138
         _ExtentY        =   508
         _Version        =   393216
         Format          =   19791873
         CurrentDate     =   37015
      End
      Begin TextFer.TxFer lblRuc 
         Height          =   300
         Left            =   6240
         TabIndex        =   44
         Top             =   192
         Width           =   1308
         _ExtentX        =   2307
         _ExtentY        =   529
         Object.CausesValidation=   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
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
         TabIndex        =   45
         Top             =   192
         Width           =   4500
         _ExtentX        =   7938
         _ExtentY        =   550
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
         Height          =   312
         Left            =   1008
         TabIndex        =   49
         Top             =   912
         Width           =   4116
         _ExtentX        =   7260
         _ExtentY        =   550
         XcodMaxLongitud =   3
         xcodwith        =   300
         NomTabla        =   "cp_proveedor"
         TituloAyuda     =   "Busqueda de Proveedor"
         ListaCampos     =   "clientecodigo(1),clienterazonsocial(1),clienteruc(1),clientetelefono(1)"
         XcodCampo       =   "clientecodigo"
         XListCampo      =   "clienterazonsocial"
         ListaCamposDescrip=   "Codigo,Descripcion"
         ListaCamposText =   "clientecodigo,clienterazonsocial,clienteruc,clientetelefono"
      End
      Begin ctrlayuda_f.Ctr_Ayuda CtrAyu_solicitante 
         Height          =   312
         Left            =   1008
         TabIndex        =   50
         Top             =   1632
         Width           =   4116
         _ExtentX        =   7260
         _ExtentY        =   550
         XcodMaxLongitud =   3
         xcodwith        =   300
         NomTabla        =   "cp_proveedor"
         TituloAyuda     =   "Busqueda de Proveedor"
         ListaCampos     =   "clientecodigo(1),clienterazonsocial(1),clienteruc(1),clientetelefono(1)"
         XcodCampo       =   "clientecodigo"
         XListCampo      =   "clienterazonsocial"
         ListaCamposDescrip=   "Codigo,Descripcion"
         ListaCamposText =   "clientecodigo,clienterazonsocial,clienteruc,clientetelefono"
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Cond.Pago     :"
         Height          =   192
         Left            =   48
         TabIndex        =   48
         Top             =   996
         Width           =   1032
      End
      Begin VB.Label Le_Proveedor 
         Caption         =   "No. Requis."
         Height          =   252
         Left            =   7728
         TabIndex        =   46
         Top             =   288
         Width           =   1020
      End
      Begin VB.Label Label12 
         Caption         =   "Observaci�n :"
         Height          =   252
         Left            =   84
         TabIndex        =   43
         Top             =   2040
         Width           =   1092
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Moneda  :"
         Height          =   192
         Left            =   5448
         TabIndex        =   42
         Top             =   600
         Width           =   720
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Entrega   :"
         Height          =   192
         Left            =   2808
         TabIndex        =   41
         Top             =   600
         Width           =   720
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "R.U.C.  :"
         Height          =   192
         Left            =   5616
         TabIndex        =   40
         Top             =   288
         Width           =   552
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Proveedor     :"
         Height          =   192
         Left            =   48
         TabIndex        =   39
         Top             =   276
         Width           =   1008
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Emisi�n         :"
         Height          =   192
         Left            =   84
         TabIndex        =   38
         Top             =   600
         Width           =   996
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Entregar en   :"
         Height          =   192
         Left            =   84
         TabIndex        =   37
         Top             =   1320
         Width           =   1008
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Solicitante     :"
         Height          =   192
         Left            =   84
         TabIndex        =   36
         Top             =   1680
         Width           =   1008
      End
      Begin VB.Label lblCen 
         AutoSize        =   -1  'True
         Caption         =   "Cotizaci�n  :"
         Height          =   192
         Left            =   5244
         TabIndex        =   35
         Top             =   960
         Width           =   876
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
      Picture         =   "frmTraEmi.frx":08CA
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
      Picture         =   "frmTraEmi.frx":0D0C
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5616
      Visible         =   0   'False
      Width           =   775
   End
   Begin VB.CommandButton cmdEdi2 
      Caption         =   "&Editar"
      Enabled         =   0   'False
      Height          =   675
      Left            =   2736
      Picture         =   "frmTraEmi.frx":114E
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5616
      Visible         =   0   'False
      Width           =   775
   End
   Begin VB.CommandButton cmdImp 
      Caption         =   "&Imprimir"
      Height          =   675
      Left            =   5535
      Picture         =   "frmTraEmi.frx":1590
      Style           =   1  'Graphical
      TabIndex        =   9
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
      TabIndex        =   7
      Top             =   3825
      Width           =   775
   End
   Begin VB.CommandButton CmdEli 
      Caption         =   "&Anular"
      Height          =   675
      Left            =   4230
      Picture         =   "frmTraEmi.frx":1E14
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3840
      Width           =   775
   End
   Begin VB.CommandButton cmdNue 
      Caption         =   "&Nuevo"
      Height          =   675
      Left            =   1575
      Picture         =   "frmTraEmi.frx":2256
      Style           =   1  'Graphical
      TabIndex        =   6
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
      TabIndex        =   10
      Top             =   3825
      Width           =   775
   End
   Begin VB.CommandButton cmdGra 
      Caption         =   "&Grabar"
      Height          =   675
      Left            =   5424
      Picture         =   "frmTraEmi.frx":2ADA
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5616
      Visible         =   0   'False
      Width           =   775
   End
   Begin VB.CommandButton CmdSalir2 
      Caption         =   "&Salir"
      Height          =   675
      Left            =   6864
      Picture         =   "frmTraEmi.frx":2F1C
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5616
      Visible         =   0   'False
      Width           =   775
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Flex1 
      Height          =   1512
      Left            =   120
      TabIndex        =   11
      Top             =   3120
      Visible         =   0   'False
      Width           =   9732
      _ExtentX        =   17166
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
      FormatString    =   "^C�digo|Fab|Descripci�n|xUni|xCantidad|Uni.|Cantidad|PU|>Precio|>%Des|Igv|>Total|C1|C2"
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
      TabIndex        =   12
      Top             =   0
      Visible         =   0   'False
      Width           =   9696
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Estado  :"
         Height          =   192
         Left            =   7068
         TabIndex        =   16
         Top             =   240
         Width           =   636
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
         Height          =   348
         Left            =   7812
         TabIndex        =   15
         Top             =   156
         Width           =   1644
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "N�mero  :"
         Height          =   195
         Left            =   375
         TabIndex        =   14
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
         TabIndex        =   13
         Top             =   160
         Width           =   1560
      End
   End
   Begin VB.Frame fraTotales 
      Height          =   975
      Left            =   135
      TabIndex        =   17
      Top             =   4560
      Visible         =   0   'False
      Width           =   9708
      Begin VB.Label lblCom 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "892,760.00"
         Height          =   285
         Left            =   7080
         TabIndex        =   27
         Top             =   600
         Width           =   1110
      End
      Begin VB.Label lblIgv 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "892,760.00"
         Height          =   285
         Left            =   7080
         TabIndex        =   26
         Top             =   240
         Width           =   1110
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Compra :"
         Height          =   195
         Left            =   6360
         TabIndex        =   25
         Top             =   600
         Width           =   630
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "I.G.V.   :"
         Height          =   195
         Left            =   6360
         TabIndex        =   24
         Top             =   240
         Width           =   585
      End
      Begin VB.Label lblTot 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "892,760.00"
         Height          =   285
         Left            =   4200
         TabIndex        =   23
         Top             =   240
         Width           =   1230
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Total  :"
         Height          =   195
         Left            =   3600
         TabIndex        =   22
         Top             =   240
         Width           =   495
      End
      Begin VB.Label lblDes 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "892,760.00"
         Height          =   285
         Left            =   1680
         TabIndex        =   21
         Top             =   600
         Width           =   1110
      End
      Begin VB.Label lblImp 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "892,760.00"
         Height          =   285
         Left            =   1680
         TabIndex        =   20
         Top             =   240
         Width           =   1110
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Descuento :"
         Height          =   195
         Left            =   720
         TabIndex        =   19
         Top             =   600
         Width           =   870
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Importe      :"
         Height          =   195
         Left            =   720
         TabIndex        =   18
         Top             =   240
         Width           =   840
      End
   End
   Begin Crystal.CrystalReport CrystalReport2 
      Bindings        =   "frmTraEmi.frx":335E
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
      Left            =   144
      TabIndex        =   5
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
         Caption         =   "        N�mero"
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
         Caption         =   "    Emisi�n"
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
Attribute VB_Name = "frmTraEmi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Colex As New Collection
Dim Adodc1 As ADODB.Recordset
Dim cSql1 As String
Dim nT As Integer       'Ingreso,Modificaci�n,Ficha Tecnica
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
    
    Set Adodc1 = New ADODB.Recordset
    
    strsql = "SELECT * FROM co_cabordcompra,co_estadoorden WHERE co_cabordcompra.oc_situacionorden =co_estadoorden." & _
        "estadooccodigo and estadoocatendido<>1 ORDER BY oc_cnumord "
    Adodc1.Open strsql, VGcnx, adOpenDynamic, adLockOptimistic
    
    Set DataGrid1.DataSource = Adodc1
    
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
        .txtordfab = Flex1.TextMatrix(Flex1.Row, 12)
        .txtCo1 = Flex1.TextMatrix(Flex1.Row, 13)
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
    
    If Adodc1("oc_csitord") <> "00" And Adodc1("oc_csitord") <> "01" Then
        Mensaje = "Imposible anular la Orden de compra en su estado actual"
        MsgBox Mensaje, vbCritical, "Mensaje"
        DataGrid1.SetFocus
        Exit Sub
    End If

    Dim strsql As String
    Dim voc As String
    
    Mensaje = "�Est� seguro que desea anular la Orden de compra?"
    If MsgBox(Mensaje, vbQuestion + vbYesNo + vbDefaultButton2, "Mensaje") = vbYes Then
        voc = Adodc1("oc_cnumord")
        
        nTra = 1
        VGcnx.BeginTrans
        
        strsql = "UPDATE co_detordcompra SET oc_cestado='06' WHERE oc_cnumord='" & voc & "'"
        VGcnx.Execute strsql
        strsql = "UPDATE co_cabordcompra SET oc_csitord='06' WHERE oc_cnumord='" & voc & "'"
        VGcnx.Execute strsql

        VGcnx.CommitTrans
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

    Mensaje = "�Desea eliminar el documento " & Adodc1("nrorequi") & "?"
    If MsgBox(Mensaje, vbQuestion + vbYesNo, "Mensaje") = vbYes Then
        strsql = "DELETE * FROM requisd WHERE nrorequi='" & Adodc1("nrorequi") & "'"
        
        nTra = 1
        VGcnx.BeginTrans
        VGcnx.Execute strsql
        VGcnx.CommitTrans
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
    MsgBox Err.Description
    If nTra = 1 Then VGcnx.RollbackTrans
End Sub

Private Sub CmdEli2_Click()
    If Tiene_Entregas Then
        Mensaje = "El art�culo tiene cantidad entregada"
        MsgBox Mensaje, vbExclamation, "Advertencia"
    End If
    
    Mensaje = "�Desea quitar el art�culo seleccionado?"
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
    Dim vURef As String, txtmon As String
    Dim txtEst As String, txttip As Integer
    On Error GoTo GrabErr
    
    txttip = 0
    If nT = 1 Then
        txtPro = Trim(txtPro)
        If txtPro = "" Then
            Mensaje = "Debe ingresar C�digo de Proveedor"
            MsgBox Mensaje, vbExclamation, "Mensaje"
            txtPro.SetFocus
            Exit Sub
        Else
            If lblPro = "" Then
                If Not Existe(1, txtPro, "maeprov", "prvccodigo", False) Then
                    Mensaje = "El C�digo de Proveedor ingresado no existe"
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
       
    txtmon = CtrAyu_moneda.xclave
    If Trim(txtmon) = "" Then
        Mensaje = "Debe ingresar el Tipo de Moneda"
        MsgBox Mensaje, vbExclamation, "Error"
        CtrAyu_moneda.SetFocus
        Exit Sub
    End If
    
    txtEst = ""
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
        Mensaje = "Debe especificar art�culos de la Orden de Compra"
        MsgBox Mensaje, vbExclamation, "Error"
        cmdNue2.SetFocus
        Exit Sub
    End If
    
    If nT = 1 Then
        Mensaje = "�Desea ingresar la nueva Orden de Compra?"
    Else
        Mensaje = "�Desea guardar los cambios realizados?"
    End If
    If MsgBox(Mensaje, vbQuestion + vbYesNo, "Mensaje") = vbYes Then
        nTra = 1
        
        VGcnx.BeginTrans
        unum = Format(Val(unum), "00000000000")
        lblNum = unum
        If nT = 1 Then      'Ingreso
          unum = Format(Devolver_Dato(1, "OC", "num_documentos", "ctncodigo", False, _
                "ctnnumero"), "00000000000")
          If unum = "" Then unum = 0
             unum = unum + 1
             unum = Format(unum, "00000000000")
          End If
          SQLc = "UPDATE num_documentos SET ctnnumero=" & Val(unum) & _
                " WHERE ctncodigo='OC'"
            VGcnx.Execute SQLc
            
            SQLc = "INSERT INTO co_cabordcompra (oc_cnumord,oc_dfecdoc,oc_ccodpro,oc_crazsoc," & _
                "oc_cdirpro,oc_ccotiza,oc_ccodmon,oc_cforpag,oc_ntipcam,oc_dfecent," & _
                "oc_cobserv,oc_csolict,oc_centreg,oc_estadoorden,oc_situacionorden,oc_nimport,oc_ndescue," & _
                "oc_nigv,oc_nventa,oc_dfecact,oc_chora,oc_cusuari,oc_cconver) VALUES ('" & _
                lblNum & "','" & txtEmi & "','" & txtPro & "','" & _
                lblPro & "','" & Devolver_Dato(1, txtPro, "maeprov", "prvccodigo", False, _
                "prvcdirecc") & "','" & txtCot & "','" & txtmon & "','" & txtFor & "'," & _
                Val(txttip) & ",'" & txtEnt & "','" & _
                SupCadSQL(txtObs) & "','" & txtSol & "','" & txtEntE & "',' ','0'," & _
                CDbl(lblImp) & "," & CDbl(lblDes) & "," & CDbl(lblIgv) & "," & CDbl(lblCom) & _
                ",'" & VGParamSistem.FechaTrabajo & "','" & Format(Time, "hh.mm.ss") & "','" & VGusuario & _
                "','" & txtEst & "')"
            VGcnx.Execute SQLc
            
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
                VGcnx.Execute SQLd
                
                SQLd = "UPDATE maeart SET aprecom=" & Val(Flex1.TextMatrix(I, 8)) & _
                    ",acodpro='" & txtPro & "',afecven='" & txtEmi _
                    & "' WHERE acodigo='" & Flex1.TextMatrix(I, 0) & "'"
                VGcnx.Execute SQLd
            Next
        ElseIf nT = 2 Then     'Modificar
            SQLc = "UPDATE co_cabordcompra SET oc_dfecdoc='" & txtEmi & _
                "',oc_ccotiza='" & txtCot & "',oc_ccodmon='" & txtmon & "',oc_cforpag='" & _
                txtFor & "',oc_ntipcam=" & Val(txttip) & ",oc_dfecent='" & _
                txtEnt & "',oc_cobserv='" & SupCadSQL(txtObs) & _
                "',oc_csolict='" & txtSol & "',oc_centreg='" & txtEntE & "',oc_nimport=" & _
                CDbl(lblImp) & ",oc_ndescue=" & CDbl(lblDes) & ",oc_nigv=" & CDbl(lblIgv) & _
                ",oc_nventa=" & CDbl(lblCom) & ",oc_dfecact='" & _
                VGParamSistem.FechaTrabajo & "',oc_chora='" & Format(Time, "hh.mm.ss") & "',oc_cusuari='" & _
                VGusuario & "',oc_cconver='" & txtEst & "' WHERE oc_cnumord='" & lblNum & "'"
            VGcnx.Execute SQLc
            
            SQLd = "DELETE * FROM co_detordcompra WHERE oc_cnumord='" & lblNum & "'"
            VGcnx.Execute SQLd
            
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
                VGcnx.Execute SQLd
                
                SQLd = "UPDATE maeart SET aprecom=" & Val(Flex1.TextMatrix(I, 8)) & _
                    ",acodpro='" & txtPro & "',afecven='" & txtEmi _
                    & "' WHERE acodigo='" & Flex1.TextMatrix(I, 0) & "'"
                VGcnx.Execute SQLd
            Next
        End If
        
        VGcnx.CommitTrans
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
            'txtEntE = VGEMP_DIREC
            txtPro.SetFocus
        Else
            CmdSalir2_Click
        End If
    Exit Sub

GrabErr:
    MsgBox Err.Description
    If nTra = 1 Then VGcnx.RollbackTrans
End Sub

Private Sub cmdImp_Click()
Dim formulas(1) As String
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
  '  CrystalReport2.Connect =
  '                              "DSN=" & VGServer & ";" & _
  '                              "DSQ=" & VGBase3 & ";" & _
  '                              "UID=" & VGBUsuario2 & ";" & _
  '                              "PWD=''"
       
    CrystalReport2.Destination = crptToWindow
    CrystalReport2.WindowState = crptMaximized
    CrystalReport2.WindowShowPrintBtn = True
    CrystalReport2.WindowShowRefreshBtn = True
    CrystalReport2.WindowShowSearchBtn = True
    CrystalReport2.WindowShowPrintSetupBtn = True
    CrystalReport2.formulas(1) = "@emp ='" & VGParamCompra.NomEmpresa & "'"
    CrystalReport2.StoredProcParam(0) = VGcnx.DefaultDatabase
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
            "ctnnumero"), "00000000000")
        If unum = "" Then unum = 0
        unum = unum + 1
        unum = Format(unum, "00000000000")
            
    ' inicio recien
    ' Selecciona todas las Orden de compra
        cSqlM = "SELECT oc_cnumord FROM co_cabordcompra ORDER BY oc_cnumord"
        Set cSelM = New ADODB.Recordset
        cSelM.Open cSqlM, VGcnx, adOpenStatic
        Do While Not cSelM.EOF
           If cSelM("oc_cnumord") = Trim(unum) Then
              cSelM.MoveLast
              unum = Format(cSelM("oc_cnumord"), "00000000000")
              unum = unum + 1
              unum = Format(unum, "00000000000")
           End If
           cSelM.MoveNext
        Loop
        cSelM.Close 'Cierra el ADODB.Recorset
   ' fin
            
    End If
    lblNum = unum
    lblEst = ""
'    txtTip = "0.000"
    lblImp = "0.00": lblTot = "0.00": lblIgv = "0.00"
    lblDes = "0.00": lblCom = "0.00"
    
    Fradatos.Enabled = True
    cmdGra.Enabled = True
    txtPro.Enabled = True
    txtPro.SetFocus
    CmdSalir2.Cancel = True
End Sub

Private Sub cmdEdi_Click()
    If Adodc1("oc_estadoorden") = "A" Then
        Mensaje = "La Orden de compra ha sido anulada, no se permitir� modificaciones"
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
         .cmbtipo.Text = "Bienes"
         .lbltipo.Visible = True
       End If
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
                        (Val(.lblPNe) + Val(.lblDes)), "0.00") & vbTab & .txtordfab & vbTab & _
                        .txtCo1 & vbTab & .Tipo, 1
                    Flex1.Rows = 2
                Else
                    Flex1.AddItem .txtCod & vbTab & .lblFab & vbTab & .txtDes & vbTab & _
                        .lblUni & vbTab & .txtCan & vbTab & IIf(.txtURe = "", .lblUni, _
                        .txtURe) & vbTab & IIf(.txtURe = "", .txtCan, .txtRef) & vbTab & _
                        .txtPUn & vbTab & Format(Val(.lblPNe) + Val(.lblDes), "0.00") & _
                        vbTab & .txtPDe & vbTab & .txtPIg & vbTab & Format(Val(.txtCan) * _
                        (Val(.lblPNe) + Val(.lblDes)), "0.00") & vbTab & .txtordfab & vbTab & _
                        .txtCo1 & vbTab & .Tipo
                    Flex1.Row = Flex1.Rows - 1
                End If
            Else
                Flex1.AddItem .txtCod & vbTab & .lblFab & vbTab & .txtDes & vbTab & _
                    .lblUni & vbTab & .txtCan & vbTab & IIf(.txtURe = "", .lblUni, _
                    .txtURe) & vbTab & IIf(.txtURe = "", .txtCan, .txtRef) & vbTab & _
                    .txtPUn & vbTab & Format(Val(.lblPNe) + Val(.lblDes), "0.00") & _
                    vbTab & .txtPDe & vbTab & .txtPIg & vbTab & Format(Val(.txtCan) * _
                    (Val(.lblPNe) + Val(.lblDes)), "0.00") & vbTab & .txtordfab & vbTab & _
                    .txtCo1 & vbTab & .Tipo
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

Private Sub Form_Load()
   ' AlinearFrm Me
    'Init_ControlDataGrid DataGrid1
    Formato_FlexGrid
    Call CtrAyu_moneda.conexion(VGcnxCT): CtrAyu_moneda.Filtro = "(monedacodigo <>'00') "
    OculObj02 False
    OculObj03 False
    OculObj04 True
    OculObj06 True
    
    unum = ""
    Abre_Tabla_OCs
    Estado_Botones
    
    Load frmTraEmi1
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

Public Function Existe(Tipo As Integer, Cod As String, Tabla As String, Campo As String, Fecha As Boolean, Optional Cod2 As String, Optional cCampo2 As String, Optional Cod3 As String, Optional cCampo3 As String, Optional Cod4 As Boolean, Optional cCampo4 As String, Optional Cod5 As String, Optional cCampo5 As String) As Boolean
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
 
Select Case Tipo
Case 1 'Bd. Comun
            cSel1.Open cSL, VGcnx, adOpenStatic
Case 2 'Bd. Config
            cSel1.Open cSL, VGcnx, adOpenStatic
Case 3 'Bd. Contab
            cSel1.Open cSL, VGcnxCT, adOpenStatic
End Select

If cSel1.RecordCount > 0 Then
     Existe = True
Else
     Existe = False
End If
'csel1.Close
End Function

Public Sub Init_ControlDataGrid(EsteGrid As DataGrid)
 With EsteGrid
 ' .AllowAddNew = False
 ' .AllowDelete = False
 ' .AllowUpdate = False
 ' .AllowRowSizing = False
 ' .TabAction = dbgControlNavigation
 ' .MarqueeStyle = dbgHighlightRow
 ' .Font =
 End With
End Sub

Sub Limpiar()
    txtPro = "":  CtrAyu_moneda.xclave = "": txtNSol = ""
: txtFor = "": txtCot = ""
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
    CtrAyu_moneda.xclave = Adodc1("oc_ccodmon")
    txtFor = Adodc1("oc_cforpag")
    txtCot = Adodc1("oc_ccotiza")
    txtEntE = Adodc1("oc_centreg")
    txtSol = Adodc1("oc_csolict")
    lblSol = Devolver_Dato(1, txtSol, "solicitantes", "sol_codigo", False, "sol_nombre")
    txtObs = Adodc1("oc_cobserv")
    
    cSqlM = "SELECT * FROM co_detordcompra WHERE oc_cnumord='" & cC1 & "' ORDER BY oc_citem"
    Set cSelM = New ADODB.Recordset
    
    cSelM.Open cSqlM, VGcnx, adOpenStatic
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
            Mensaje = "Fecha No V�lida"
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
            Mensaje = "Fecha No V�lida"
            MsgBox Mensaje, vbExclamation, "Error"
            txtEnt.SetFocus
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

Private Sub Enfoque(OBJ As Object)
  OBJ.SelStart = 0
  OBJ.SelLength = Len(OBJ)
End Sub

Private Sub txtPro_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtPro = Trim(txtPro)
        If txtPro <> "" Then
            If Not Existe(1, txtPro, "maeprov", "prvccodigo", False) Then
                Mensaje = "El C�digo de Proveedor ingresado no existe"
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

Sub Calcula_Totales()
    Dim I As Integer
    Dim tV As Single, Valor As Single
    Dim tD As Single, vDesc As Single
    Dim tI As Single, vIgv As Single
    
    With Flex1
        For I = 1 To Flex1.Rows - 1
            tV = Val(.TextMatrix(I, 11))
            Valor = Valor + tV
            tD = tV * Val(.TextMatrix(I, 9)) / 100
            vDesc = vDesc + tD
            tI = (tV - tD) * Val(.TextMatrix(I, 10)) / 100
            vIgv = vIgv + tI
        Next
    End With
    
    lblImp = Format(Valor, "##,##0.00")
    lblDes = Format(vDesc, "##,##0.00")
    lblTot = Format(Valor - vDesc, "#,##0.00")
    lblIgv = Format(vIgv, "#,##0.00")
    lblCom = Format((Valor - vDesc) + vIgv, "#,##0.00")
End Sub

Function Tiene_Entregas() As Boolean
    Dim Adodc2 As ADODB.Recordset
    
    Set Adodc2 = New ADODB.Recordset
    
    Adodc2.Open "SELECT * FROM co_detordcompra WHERE oc_cnumord='" & lblNum & "' AND oc_ccodigo='" & _
        Flex1.TextMatrix(Flex1.Row, 0) & "' AND oc_ncanten>0", VGcnx, adOpenStatic
    Tiene_Entregas = False
    If Adodc2.RecordCount > 0 Then Tiene_Entregas = True
End Function
