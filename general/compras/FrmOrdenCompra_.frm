VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Begin VB.Form FrmOrdenCompra 
   Caption         =   "Orden de Compra"
   ClientHeight    =   7620
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   10080
   ForeColor       =   &H80000016&
   Icon            =   "FrmOrdenCompra_.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   NegotiateMenus  =   0   'False
   ScaleHeight     =   7620
   ScaleWidth      =   10080
   Begin VB.Frame FComando 
      Height          =   570
      Left            =   2265
      TabIndex        =   51
      Top             =   7020
      Width           =   6660
      Begin VB.CommandButton CmdNuevo 
         Caption         =   "&Nuevo"
         Height          =   375
         Left            =   90
         TabIndex        =   0
         Top             =   135
         Width           =   915
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "&Buscar"
         Height          =   375
         Left            =   1005
         TabIndex        =   1
         Top             =   135
         Width           =   915
      End
      Begin VB.CommandButton CmdModificar 
         Caption         =   "&Modificar"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1920
         TabIndex        =   21
         Top             =   135
         Width           =   915
      End
      Begin VB.CommandButton CmdEliminar 
         Caption         =   "&Eliminar"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2835
         TabIndex        =   22
         Top             =   135
         Width           =   915
      End
      Begin VB.CommandButton CmdGuardar 
         Caption         =   "&Guardar"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3750
         TabIndex        =   23
         Top             =   135
         Width           =   915
      End
      Begin VB.CommandButton CmdSalir 
         Caption         =   "&Salir"
         Height          =   375
         Left            =   5580
         TabIndex        =   25
         Top             =   135
         Width           =   915
      End
      Begin VB.CommandButton CmdImprimir 
         Caption         =   "&Imprimir"
         Enabled         =   0   'False
         Height          =   375
         Left            =   4665
         TabIndex        =   24
         Top             =   135
         Width           =   915
      End
   End
   Begin VB.Frame FCrono 
      Height          =   2220
      Left            =   585
      TabIndex        =   81
      Top             =   4665
      Visible         =   0   'False
      Width           =   3660
      Begin VB.CommandButton CmdCronSalir 
         Caption         =   "Sali&r"
         Height          =   285
         Left            =   2565
         TabIndex        =   97
         Top             =   360
         Width           =   1005
      End
      Begin VB.CommandButton CmdConfirmar 
         Caption         =   "Con&firmar"
         Height          =   285
         Left            =   1305
         TabIndex        =   83
         Top             =   360
         Width           =   1275
      End
      Begin VB.CheckBox ChkCronUm 
         BackColor       =   &H00FF0000&
         Caption         =   "Unid Almacén"
         ForeColor       =   &H8000000E&
         Height          =   285
         Left            =   90
         MaskColor       =   &H0080C0FF&
         Style           =   1  'Graphical
         TabIndex        =   98
         Top             =   360
         Width           =   1230
      End
      Begin VB.CommandButton CmdProbar 
         Caption         =   "&Probar"
         Height          =   285
         Left            =   7290
         TabIndex        =   82
         Top             =   2250
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox TxtCronCant 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   6795
         TabIndex        =   84
         Text            =   "fsdfdsfds"
         Top             =   2520
         Visible         =   0   'False
         Width           =   1500
      End
      Begin MSMask.MaskEdBox TxtCronFech 
         Height          =   285
         Left            =   6840
         TabIndex        =   85
         Top             =   2835
         Visible         =   0   'False
         Width           =   1185
         _ExtentX        =   2096
         _ExtentY        =   508
         _Version        =   393216
         AutoTab         =   -1  'True
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSFlexGridLib.MSFlexGrid FxgIndi 
         Height          =   1230
         Left            =   90
         TabIndex        =   86
         Top             =   945
         Width           =   3525
         _ExtentX        =   6202
         _ExtentY        =   2159
         _Version        =   393216
         Rows            =   10
         Cols            =   3
         BackColorFixed  =   12417118
         ForeColorFixed  =   -2147483639
         FormatString    =   ".Item.|.      Fecha      .|.      Cantidad      ."
      End
      Begin MSFlexGridLib.MSFlexGrid FxgCron 
         Height          =   915
         Left            =   3735
         TabIndex        =   87
         Top             =   2205
         Visible         =   0   'False
         Width           =   3480
         _ExtentX        =   6138
         _ExtentY        =   1609
         _Version        =   393216
         Rows            =   100
         Cols            =   3
         BackColorFixed  =   12417118
         ForeColorFixed  =   -2147483639
         FormatString    =   ".Item.|.      Fecha      .|.      Cantidad      ."
      End
      Begin VB.Label LblCronUmFact 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C00000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H8000000E&
         Height          =   240
         Left            =   2565
         TabIndex        =   99
         Top             =   135
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Label Label23 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "Cronograma"
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   90
         TabIndex        =   94
         Top             =   135
         Width           =   3480
      End
      Begin VB.Label Label24 
         Caption         =   "Item"
         Height          =   240
         Left            =   180
         TabIndex        =   96
         Top             =   1305
         Width           =   510
      End
      Begin VB.Label LblCronItem 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   240
         Left            =   1215
         TabIndex        =   95
         Top             =   1350
         Width           =   960
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
         Height          =   240
         Left            =   90
         TabIndex        =   93
         Top             =   675
         Width           =   555
      End
      Begin VB.Label LblCronTota 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   240
         Left            =   540
         TabIndex        =   92
         Top             =   675
         Width           =   960
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Programado"
         Height          =   240
         Left            =   1665
         TabIndex        =   91
         Top             =   675
         Width           =   960
      End
      Begin VB.Label LblCronProg 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         Height          =   240
         Left            =   2610
         TabIndex        =   90
         Top             =   675
         Width           =   960
      End
      Begin VB.Label LblCronPosi 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         Height          =   240
         Left            =   1215
         TabIndex        =   89
         Top             =   1575
         Width           =   960
      End
      Begin VB.Label Label13 
         Caption         =   "Posicion"
         Height          =   240
         Left            =   180
         TabIndex        =   88
         Top             =   1575
         Width           =   960
      End
   End
   Begin VB.Frame FLista 
      Height          =   2355
      Left            =   585
      TabIndex        =   62
      Top             =   4545
      Width           =   5055
      Begin VB.ListBox LXCodi 
         Height          =   1968
         Left            =   4275
         TabIndex        =   63
         Top             =   225
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.ListBox LEnvi 
         Height          =   1968
         Left            =   90
         TabIndex        =   70
         Top             =   225
         Visible         =   0   'False
         Width           =   4830
      End
      Begin VB.ListBox LPago 
         Height          =   1968
         Left            =   90
         TabIndex        =   69
         Top             =   225
         Visible         =   0   'False
         Width           =   4830
      End
      Begin VB.ListBox LMerc 
         Height          =   1968
         Left            =   90
         TabIndex        =   68
         Top             =   225
         Visible         =   0   'False
         Width           =   4830
      End
      Begin VB.ListBox LUnid 
         Height          =   1968
         Left            =   90
         TabIndex        =   72
         Top             =   225
         Visible         =   0   'False
         Width           =   2445
      End
      Begin VB.ListBox LUnid2 
         Height          =   1968
         Left            =   2520
         TabIndex        =   73
         Top             =   225
         Visible         =   0   'False
         Width           =   2400
      End
      Begin VB.ListBox LCost 
         Height          =   1968
         Left            =   90
         TabIndex        =   67
         Top             =   225
         Visible         =   0   'False
         Width           =   4830
      End
      Begin VB.ListBox LComp 
         Height          =   1968
         Left            =   90
         TabIndex        =   66
         Top             =   225
         Visible         =   0   'False
         Width           =   4830
      End
      Begin VB.ListBox LProv 
         Height          =   1968
         Left            =   90
         TabIndex        =   65
         Top             =   225
         Visible         =   0   'False
         Width           =   4830
      End
      Begin VB.ListBox LCodi 
         Height          =   1968
         Left            =   90
         TabIndex        =   64
         Top             =   225
         Visible         =   0   'False
         Width           =   4830
      End
   End
   Begin VB.Frame FGene 
      Enabled         =   0   'False
      ForeColor       =   &H8000000F&
      Height          =   2760
      Left            =   60
      TabIndex        =   26
      Top             =   0
      Width           =   10005
      Begin VB.Frame Frame1 
         Caption         =   "Estado"
         Height          =   1110
         Left            =   7770
         TabIndex        =   101
         Top             =   1590
         Width           =   2175
         Begin VB.OptionButton OptEstado 
            Caption         =   "Atendido"
            Height          =   270
            Index           =   2
            Left            =   165
            TabIndex        =   104
            Top             =   810
            Width           =   1845
         End
         Begin VB.OptionButton OptEstado 
            Caption         =   "Parcialm. Atendido"
            Height          =   270
            Index           =   1
            Left            =   165
            TabIndex        =   103
            Top             =   525
            Width           =   1845
         End
         Begin VB.OptionButton OptEstado 
            Caption         =   "Pendiente"
            Height          =   270
            Index           =   0
            Left            =   165
            TabIndex        =   102
            Top             =   240
            Width           =   1845
         End
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ccostos 
         Height          =   315
         Left            =   1500
         TabIndex        =   5
         Top             =   1605
         Width           =   3315
         _ExtentX        =   5842
         _ExtentY        =   550
         XcodMaxLongitud =   10
         xcodwith        =   1000
         NomTabla        =   "ct_centrocosto"
         TituloAyuda     =   "Ayuda de Centro de Costos"
         ListaCampos     =   "centrocostocodigo(1), centrocostodescripcion(2)"
         XcodCampo       =   "centrocostocodigo"
         XListCampo      =   "centrocostodescripcion"
         ListaCamposDescrip=   "Codigo,Descripcion"
         ListaCamposText =   "centrocostocodigo,centrocostodescripcion"
         Requerido       =   0   'False
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_TipoArt 
         Height          =   315
         Left            =   1470
         TabIndex        =   2
         Top             =   495
         Width           =   3330
         _ExtentX        =   5884
         _ExtentY        =   550
         XcodMaxLongitud =   2
         xcodwith        =   100
         NomTabla        =   "al_tipoarticulo"
         TituloAyuda     =   "Tipo de Articulo"
         ListaCampos     =   "Tipoarticulocodigo(1),tipoarticuloDescripcion(2)"
         XcodCampo       =   "tipoarticulocodigo"
         XListCampo      =   "tipoarticulodescripcion"
         ListaCamposDescrip=   "Código,Descripción"
         ListaCamposText =   "Tipoarticulocodigo,tipoarticuloDescripcion"
         Requerido       =   0   'False
      End
      Begin VB.CheckBox ChkMone 
         BackColor       =   &H00008000&
         Caption         =   "&Dólares"
         ForeColor       =   &H8000000E&
         Height          =   285
         Left            =   5130
         MaskColor       =   &H0080C0FF&
         Style           =   1  'Graphical
         TabIndex        =   74
         Top             =   1215
         Width           =   1275
      End
      Begin MSMask.MaskEdBox TxtFechEmis 
         Height          =   285
         Left            =   6525
         TabIndex        =   8
         Top             =   180
         Width           =   1185
         _ExtentX        =   2096
         _ExtentY        =   508
         _Version        =   393216
         AutoTab         =   -1  'True
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.TextBox TxtObse 
         Height          =   690
         Left            =   6525
         MaxLength       =   140
         MultiLine       =   -1  'True
         TabIndex        =   10
         Top             =   900
         Width           =   3345
      End
      Begin VB.TextBox TxtMerc 
         Height          =   285
         Left            =   1500
         TabIndex        =   6
         Top             =   1950
         Width           =   375
      End
      Begin VB.CheckBox ChkCond 
         Alignment       =   1  'Right Justify
         Caption         =   "Liquidar"
         Enabled         =   0   'False
         Height          =   285
         Left            =   5160
         TabIndex        =   13
         Top             =   2280
         Width           =   1995
      End
      Begin VB.TextBox TxtCont 
         Height          =   285
         Left            =   1485
         TabIndex        =   4
         Text            =   "sfdsfgdg"
         Top             =   1260
         Width           =   3300
      End
      Begin MSMask.MaskEdBox TxtFechEntr 
         Height          =   285
         Left            =   6525
         TabIndex        =   9
         Top             =   540
         Width           =   1185
         _ExtentX        =   2096
         _ExtentY        =   508
         _Version        =   393216
         AutoTab         =   -1  'True
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox TxtIgv 
         Height          =   285
         Left            =   6525
         TabIndex        =   11
         Top             =   1620
         Width           =   1185
         _ExtentX        =   2096
         _ExtentY        =   508
         _Version        =   393216
         AutoTab         =   -1  'True
         Format          =   "#0.00"
         PromptChar      =   "0"
      End
      Begin MSMask.MaskEdBox TxtDscto 
         Height          =   285
         Left            =   6525
         TabIndex        =   12
         Top             =   1935
         Width           =   1185
         _ExtentX        =   2096
         _ExtentY        =   508
         _Version        =   393216
         AutoTab         =   -1  'True
         MaxLength       =   5
         Format          =   "#0.00"
         PromptChar      =   "0"
      End
      Begin ctrlayuda_f.Ctr_Ayuda CtrAyu_Proveedor 
         Height          =   315
         Left            =   1470
         TabIndex        =   3
         Top             =   840
         Width           =   3345
         _ExtentX        =   5906
         _ExtentY        =   550
         XcodMaxLongitud =   11
         xcodwith        =   900
         NomTabla        =   "cp_proveedor"
         TituloAyuda     =   "Busqueda de Proveedor"
         ListaCampos     =   "clientecodigo(1),clienterazonsocial(1),clienteruc(1),clientetelefono(1)"
         XcodCampo       =   "clientecodigo"
         XListCampo      =   "clienterazonsocial"
         ListaCamposDescrip=   "Codigo,Descripcion"
         ListaCamposText =   "clientecodigo,clienterazonsocial,clienteruc,clientetelefono"
         Requerido       =   0   'False
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_CondPag 
         Height          =   315
         Left            =   1500
         TabIndex        =   7
         Top             =   2250
         Width           =   3315
         _ExtentX        =   5842
         _ExtentY        =   550
         XcodMaxLongitud =   3
         xcodwith        =   200
         NomTabla        =   "co_condicionespago"
         TituloAyuda     =   "Ayuda de Condición de Pago"
         ListaCampos     =   "PagoCodigo(1),Pagodescripcion(1)"
         XcodCampo       =   "PagoCodigo"
         XListCampo      =   "Pagodescripcion"
         ListaCamposDescrip=   "Código,Descripción"
         ListaCamposText =   "PagoCodigo,Pagodescripcion"
         Requerido       =   0   'False
      End
      Begin VB.Label LblMerc 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1890
         TabIndex        =   100
         Top             =   1950
         Width           =   1770
      End
      Begin VB.Label Label3 
         Caption         =   "Descuento %"
         Height          =   285
         Left            =   5175
         TabIndex        =   76
         Top             =   1935
         Width           =   1365
      End
      Begin VB.Label Label19 
         Caption         =   "Factor IGV"
         Height          =   285
         Left            =   5175
         TabIndex        =   38
         Top             =   1620
         Width           =   1365
      End
      Begin VB.Label Label22 
         Caption         =   "Observaciones"
         Height          =   285
         Left            =   5175
         TabIndex        =   37
         Top             =   900
         Width           =   1185
      End
      Begin VB.Label Label20 
         Caption         =   "Fecha de Entrega"
         Height          =   285
         Left            =   5175
         TabIndex        =   36
         Top             =   540
         Width           =   1365
      End
      Begin VB.Label Label18 
         Caption         =   "Fecha de Emisión"
         Height          =   285
         Left            =   5175
         TabIndex        =   35
         Top             =   180
         Width           =   1365
      End
      Begin VB.Label Label16 
         Caption         =   "Condición Pago"
         Height          =   285
         Left            =   135
         TabIndex        =   34
         Top             =   2340
         Width           =   1320
      End
      Begin VB.Label Label12 
         Caption         =   "Tipo de Compra"
         Height          =   285
         Left            =   135
         TabIndex        =   33
         Top             =   1980
         Width           =   1320
      End
      Begin VB.Label Label10 
         Caption         =   "Centro de Costos"
         Height          =   285
         Left            =   135
         TabIndex        =   32
         Top             =   1620
         Width           =   1320
      End
      Begin VB.Label Label9 
         Caption         =   "Representante"
         Height          =   285
         Left            =   135
         TabIndex        =   31
         Top             =   1260
         Width           =   1185
      End
      Begin VB.Label Label7 
         Caption         =   "Proveedor"
         Height          =   285
         Left            =   135
         TabIndex        =   30
         Top             =   900
         Width           =   1185
      End
      Begin VB.Label Label5 
         Caption         =   "Tipo de Artículo"
         Height          =   285
         Left            =   135
         TabIndex        =   29
         Top             =   540
         Width           =   1185
      End
      Begin VB.Label LblParte 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H0080C0FF&
         Height          =   285
         Left            =   1485
         TabIndex        =   28
         Top             =   180
         Width           =   1185
      End
      Begin VB.Label Label1 
         Caption         =   "Cod Compra"
         Height          =   240
         Left            =   135
         TabIndex        =   27
         Top             =   180
         Width           =   1005
      End
   End
   Begin VB.Frame FProg 
      ForeColor       =   &H0000FFFF&
      Height          =   555
      Left            =   45
      TabIndex        =   40
      Top             =   3690
      Width           =   8850
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Desc."
         Height          =   285
         Left            =   4140
         TabIndex        =   78
         Top             =   210
         Width           =   495
      End
      Begin VB.Label LblDscto 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   285
         Left            =   4620
         TabIndex        =   77
         Top             =   180
         Width           =   1320
      End
      Begin VB.Label LblImp 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   285
         Left            =   855
         TabIndex        =   46
         Top             =   180
         Width           =   1320
      End
      Begin VB.Label Label25 
         BackStyle       =   0  'Transparent
         Caption         =   "Imp. Bruto"
         Height          =   285
         Left            =   75
         TabIndex        =   45
         Top             =   180
         Width           =   795
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "I.G.V."
         Height          =   285
         Left            =   2220
         TabIndex        =   44
         Top             =   225
         Width           =   450
      End
      Begin VB.Label LblIgv 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   285
         Left            =   2670
         TabIndex        =   43
         Top             =   180
         Width           =   1320
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Total a Pagar"
         Height          =   285
         Left            =   6060
         TabIndex        =   42
         Top             =   195
         Width           =   990
      End
      Begin VB.Label LblTota 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   285
         Left            =   7125
         TabIndex        =   41
         Top             =   195
         Width           =   1320
      End
   End
   Begin VB.Frame FDeta 
      Enabled         =   0   'False
      Height          =   960
      Left            =   45
      TabIndex        =   52
      Top             =   2745
      Width           =   10005
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Uni 
         CausesValidation=   0   'False
         Height          =   315
         Left            =   2130
         TabIndex        =   17
         Top             =   585
         Visible         =   0   'False
         Width           =   1770
         _ExtentX        =   3112
         _ExtentY        =   550
         Enabled         =   0   'False
         XcodMaxLongitud =   0
         xcodwith        =   100
         NomTabla        =   "v_unidad"
         TituloAyuda     =   "Ayuda de Unidades"
         ListaCampos     =   "Codigo(1),Conver(1),Factor(1)"
         XcodCampo       =   "Codigo"
         XListCampo      =   "Conver"
         ListaCamposDescrip=   "Codigo,Descripcion,Factor"
         ListaCamposText =   "Codigo,Conver,Factor"
         Requerido       =   0   'False
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Art 
         Height          =   345
         Left            =   675
         TabIndex        =   14
         Top             =   225
         Width           =   5700
         _ExtentX        =   10054
         _ExtentY        =   614
         XcodMaxLongitud =   0
         xcodwith        =   1500
         NomTabla        =   "maeart"
         ListaCampos     =   "acodigo(1),adescri(1)"
         XcodCampo       =   "acodigo"
         XListCampo      =   "adescri"
         ListaCamposDescrip=   "Codigo,descripcion"
         ListaCamposText =   "acodigo,adescri"
         Requerido       =   0   'False
      End
      Begin VB.CheckBox ChkUm 
         BackColor       =   &H000000FF&
         Caption         =   "Ualma Cant"
         ForeColor       =   &H8000000E&
         Height          =   285
         Left            =   4080
         MaskColor       =   &H0080C0FF&
         Style           =   1  'Graphical
         TabIndex        =   80
         Top             =   585
         Width           =   1230
      End
      Begin VB.TextBox TxtPedi 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   675
         TabIndex        =   16
         Text            =   "7115.00"
         Top             =   600
         Width           =   780
      End
      Begin VB.TextBox TxtCantPedi 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   5355
         TabIndex        =   18
         Text            =   "fgsd"
         Top             =   600
         Width           =   1005
      End
      Begin VB.TextBox TxtCome 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   7380
         TabIndex        =   15
         Text            =   "gsfdgfs"
         Top             =   180
         Width           =   2445
      End
      Begin VB.TextBox TxtPrec 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   7065
         TabIndex        =   19
         Text            =   "sdfgsfgfsdgfdsg"
         Top             =   585
         Width           =   1230
      End
      Begin VB.TextBox TxtDscto2 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   9045
         TabIndex        =   20
         Text            =   "sdfgsfgfsdgfdsg"
         Top             =   585
         Width           =   780
      End
      Begin VB.Label LblUmFact 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Kg"
         Height          =   285
         Left            =   4110
         TabIndex        =   71
         Top             =   585
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.Label Label222 
         Caption         =   "Artículo"
         Height          =   255
         Left            =   90
         TabIndex        =   61
         Top             =   285
         Width           =   555
      End
      Begin VB.Label Label17 
         Caption         =   "Pedido"
         Height          =   285
         Left            =   75
         TabIndex        =   60
         Top             =   645
         Width           =   555
      End
      Begin VB.Label LPNeto 
         Alignment       =   2  'Center
         Caption         =   "Comentario"
         Height          =   255
         Left            =   6480
         TabIndex        =   58
         Top             =   180
         Width           =   840
      End
      Begin VB.Label Label21 
         Caption         =   "U Med"
         Height          =   240
         Left            =   1560
         TabIndex        =   57
         Top             =   645
         Width           =   480
      End
      Begin VB.Label LblPosi 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   285
         Left            =   4125
         TabIndex        =   56
         Top             =   540
         Width           =   375
      End
      Begin VB.Label LblPosiEdit 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   285
         Left            =   4110
         TabIndex        =   55
         Top             =   540
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label LblDocu 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   285
         Left            =   4125
         TabIndex        =   54
         Top             =   585
         Width           =   570
      End
      Begin VB.Label LblLote 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   4080
         TabIndex        =   53
         Top             =   540
         Width           =   570
      End
      Begin VB.Label Label6 
         Caption         =   "Dscto %"
         Height          =   180
         Left            =   8370
         TabIndex        =   75
         Top             =   630
         Width           =   660
      End
      Begin VB.Label Label8 
         Caption         =   "Precio"
         Height          =   180
         Left            =   6480
         TabIndex        =   59
         Top             =   630
         Width           =   525
      End
   End
   Begin VB.Frame FComandoDeta 
      Enabled         =   0   'False
      Height          =   1620
      Left            =   9000
      TabIndex        =   47
      Top             =   4215
      Width           =   1095
      Begin VB.CommandButton CmdNuevoDeta 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Agregar"
         Height          =   405
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   50
         ToolTipText     =   "Nuevo Item"
         Top             =   165
         Width           =   915
      End
      Begin VB.CommandButton CmdActualiza 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Insertar"
         Height          =   420
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   49
         ToolTipText     =   "Actualizar item"
         Top             =   600
         Width           =   915
      End
      Begin VB.CommandButton CmdBorrar 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Bor&rar"
         Height          =   435
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   48
         ToolTipText     =   "Eliminar Item"
         Top             =   1065
         Width           =   915
      End
   End
   Begin MSFlexGridLib.MSFlexGrid FxgDeta 
      Height          =   2715
      Left            =   45
      TabIndex        =   39
      Top             =   4275
      Width           =   8880
      _ExtentX        =   15663
      _ExtentY        =   4805
      _Version        =   393216
      Rows            =   15
      Cols            =   13
      BackColorFixed  =   128
      ForeColorFixed  =   -2147483639
      BackColorSel    =   128
      ForeColorSel    =   -2147483643
      FormatString    =   $"FrmOrdenCompra_.frx":1272
   End
   Begin MSComctlLib.ProgressBar PBar 
      Height          =   150
      Left            =   60
      TabIndex        =   79
      Top             =   7005
      Visible         =   0   'False
      Width           =   2025
      _ExtentX        =   3577
      _ExtentY        =   275
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Menu m_crono 
      Caption         =   "&Cronograma"
      Visible         =   0   'False
      WindowList      =   -1  'True
      Begin VB.Menu m_crono_actualiza 
         Caption         =   "&Actualizar Cronograma"
      End
      Begin VB.Menu m_crono_eliminar 
         Caption         =   "&Eliminar Cronograma"
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "FrmOrdenCompra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cCodi, cLCodi, cClas, cLClas As String
Dim cComp, cLComp, cProv, cLProv As String
Dim cCost, cLCost, cMerc, cLMerc As String
Dim cPago, cLPago As String
Dim cUnid, cLUnid, cL2Unid As String
Dim Veri As Integer
Dim gParte, gTipo, gProv, gCost, gMerc, gCond As Integer
Dim gIgv, gIgvF, gDsctoF, gDscto, gImpo, gNeto As Double
Dim gRepr, gObse As String
Dim gFech, gFechEntr As Date
Dim xDeta(18) As String
Dim t_Grid As Double
Dim CmdComando As ADODB.Command
Public Tipcom As String
Dim vlestado As Integer
'El Centro de Costo cuando es blanco por defecto es el codigo 10110
Private Sub ChkCronUm_Click()
On Error Resume Next
a = Val(LblCronTota.Caption)
b = Val(LblCronProg.Caption)
c = Val(LblCronUmFact.Caption)
    
With ChkCronUm
    If .Value = 1 Then
        .BackColor = RGB(0, 150, 150)
        .Caption = "Unid Compra"
        LblCronTota.Caption = a * c
        LblCronProg.Caption = b * c
        FxgIndi.Col = 2
        For i = 1 To FxgIndi.Rows - 1
            FxgIndi.Row = i
            d = Val(FxgIndi.Text)
            If d > 0 Then FxgIndi.Text = d * c
        Next
    Else
        .BackColor = RGB(0, 0, 250)
        .Caption = "Unid Almacén"
        LblCronTota.Caption = a / c
        LblCronProg.Caption = b / c
        FxgIndi.Col = 2
        For i = 1 To FxgIndi.Rows - 1
            FxgIndi.Row = i
            d = Val(FxgIndi.Text)
            If d > 0 Then FxgIndi.Text = d / c
        Next
    End If
End With
FxgIndi.Col = 1
FxgIndi.Row = 1
FxgIndi.SetFocus

End Sub

Private Sub ChkMone_Click()
On Error Resume Next
With ChkMone
    If .Value = 0 Then
        .BackColor = RGB(0, 150, 0)
        .Caption = "&Dólares"
    Else
        .BackColor = RGB(0, 0, 150)
        .Caption = "&Soles"
    End If
End With
TxtObse.SetFocus
End Sub

Private Sub ChkUm_Click()
On Error Resume Next

a = Val(TxtCantPedi.Text)
b = Val(TxtPrec.Text)
c = Val(LblUmFact.Caption)
    
With ChkUm
    If .Value = 1 Then
        .BackColor = RGB(0, 150, 150)
        .Caption = "UCompr Cant"
        TxtCantPedi.Text = a * c
        TxtPrec.Text = b / c
    Else
        .BackColor = RGB(0, 0, 250)
        .Caption = "UAlmac Cant "
        TxtCantPedi.Text = a / c
        TxtPrec.Text = b * c
    End If
End With
TxtCantPedi.SetFocus
End Sub

Private Sub CmdActualiza_Click()
AgregaDetalle
End Sub

Private Sub CmdBorrar_Click()
On Error Resume Next

xPos = Val(LblPosiEdit.Caption)
If Val(xPos) = 0 Then Exit Sub
FxgDeta.Row = xPos
For i = 0 To FxgDeta.Cols - 1
    FxgDeta.Col = i
    FxgDeta.Text = ""
Next
BorraCronograma Val(xPos)
LimpiaDetalle
Ctr_Art.SetFocus

Exit Sub
ErrBorra:
MsgBox "El Sistema no permite la eliminación del registro seleccionado", vbCritical
End Sub

Private Sub CmdBuscar_Click()
BuscaDocumento
End Sub

Private Sub CmdCronSalir_Click()
Me.ChkCronUm.Value = 0
LblCronItem.Caption = 0
FDeta.Enabled = True
FComandoDeta.Enabled = True
FxgDeta.Enabled = True
FCrono.Visible = False
End Sub

Private Sub CmdGuardar_Click()
GuardaDocumento
End Sub

Private Sub CmdImprimir_Click()
    'On Error Resume Next
    'FrmRptOcEmisión.Show
    Call imprimir
End Sub
Public Sub imprimir()
Dim arrform(0) As Variant, arrparm(13) As Variant
On Error GoTo Imprime
    Screen.MousePointer = 11
    'Val(FrmOrdenCompra.LblParte.Caption) & ", " & Val(FrmOrdenCompra.Tipcom)
    Dim rsdate As New ADODB.Recordset
    Set VGvardllgen = New dllgeneral.dll_general
    arrparm(0) = Val(FrmOrdenCompra.LblParte.Caption)
    arrparm(1) = Val(FrmOrdenCompra.Tipcom)
    Call ImpresionRptbase("rptcoestaordencompra.rpt", arrform, arrparm, , "Listado de orden de compra pendientes ")
    Screen.MousePointer = 1
    Exit Sub
Imprime:
Screen.MousePointer = 1
MsgBox Err.Description
End Sub


Private Sub CmdModificar_Click()
On Error Resume Next
CmdModificar.Enabled = False
CmdEliminar.Enabled = False
FGene.Enabled = True
FDeta.Enabled = True
FxgDeta.Enabled = True
FComandoDeta.Enabled = True
CmdGuardar.Enabled = True
'TxtComp.SetFocus
Ctr_TipoArt.SetFocus
End Sub

Private Sub CmdNuevo_Click()
On Error Resume Next
LimpiaForm
TxtFechEmis.Text = "  /  /    "
TxtFechEntr.Text = "  /  /    "
CmdModificar.Enabled = False
CmdEliminar.Enabled = False
FGene.Enabled = True
FDeta.Enabled = True
FxgDeta.Enabled = True
FComandoDeta.Enabled = True
CmdGuardar.Enabled = True
'TxtComp.SetFocus
Ctr_TipoArt.SetFocus
OptEstado(0).Value = True
End Sub

Private Sub CmdNuevoDeta_Click()
On Error Resume Next
LimpiaDetalle
Ctr_Art.SetFocus
End Sub

Private Sub CmdSalir_Click()
Unload Me
End Sub

Private Sub Ctr_TipoArt_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
    Tipcom = ColecCampos(0).Value
'    Select Case ColecCampos(0).Value
'        Case "1":
'            Ctr_Art.Enabled = True
'            Ctr_Art.TituloAyuda = "Ayuda de Maestro de Hilos"
'            Ctr_Art.NomTabla = "[Maestro Hilos]"
'            Ctr_Art.ListaCampos = "HiloCodigo(1),HiloDescripcion(1)"
'            Ctr_Art.ListaCamposDescrip = "Código,Descripción"
'            Ctr_Art.ListaCamposText = "HiloCodigo,HiloDescripcion"
'            Ctr_Art.XcodCampo = "HiloCodigo"
'            Ctr_Art.XListCampo = "HiloDescripcion"
'            Call Ctr_Art.conexion(VGcnx)
'        Case "2":
'            Ctr_Art.Enabled = True
'            Ctr_Art.TituloAyuda = "Ayuda de Maestro Tela Cruda"
'            Ctr_Art.NomTabla = "[Maestro Tela Cruda]"
'            Ctr_Art.ListaCampos = "TelaCrudaID(1),TelaCrudaDescripcion(1)"
'            Ctr_Art.ListaCamposDescrip = "Código,Descripción"
'            Ctr_Art.ListaCamposText = "TelaCrudaID,TelaCrudaDescripcion"
'            Ctr_Art.XcodCampo = "TelaCrudaID"
'            Ctr_Art.XListCampo = "TelaCrudaDescripcion"
'            Call Ctr_Art.conexion(VGcnx)
'        Case "5":
'            Ctr_Art.Enabled = True
'            Ctr_Art.TituloAyuda = "Ayuda de Maestro Quimicos"
'            Ctr_Art.NomTabla = "[Maestro Quimicos]"
'            Ctr_Art.ListaCampos = "QuimicoId(1),QuimicoDescripcion(1)"
'            Ctr_Art.ListaCamposDescrip = "Código,Descripción"
'            Ctr_Art.ListaCamposText = "QuimicoId,QuimicoDescripcion"
'            Ctr_Art.XcodCampo = "QuimicoId"
'            Ctr_Art.XListCampo = "QuimicoDescripcion"
'            Call Ctr_Art.conexion(VGcnx)
'       Case Else:
'            Ctr_Art.Enabled = True
'    End Select
End Sub

Private Sub Ctr_TipoArt_AlNoDevolverNada()
    Tipcom = ""
End Sub

Private Sub Ctr_Uni_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
    LblUmFact.Caption = ColecCampos("Factor").Value
End Sub

Private Sub Ctr_Uni_AlNoDevolverNada()
    LblUmFact.Caption = ""
End Sub

Private Sub Form_Load()
    LimpiaForm
    Call ConectarAyudas
End Sub
Private Sub ConectarAyudas()
    Call Ctr_TipoArt.conexion(VGcnx)
    Call Ctr_Art.conexion(VGcnx)
    Call CtrAyu_Proveedor.conexion(VGcnxCP)
    Call Ctr_Ccostos.conexion(VGcnx)
    Call Ctr_CondPag.conexion(VGcnx)
    Call Ctr_Uni.conexion(VGcnx)
End Sub
Private Sub Form_Resize()
If Me.Height = 360 Or Me.Width = 2400 Then Exit Sub
Me.Height = 8025
Me.Width = 10200
End Sub
Public Sub LimpiaControl(Formulario As Form, ControlType As String)
    Dim ctlControl As Control, strTipoControl
        
    For Each ctlControl In Formulario.Controls
        
        strTipoControl = TypeName(ctlControl)
        
        If InStr(ControlType, strTipoControl) Then
            Select Case strTipoControl
                Case "TextBox"
                    ctlControl.Text = Empty
                
                Case "CheckBox"
                    ctlControl.Value = False
                    
                Case "MaskEdBox"
                    ctlControl.Text = Empty
                
                Case "DataCombo"
                    ctlControl.Text = Empty
            End Select
        End If
    Next
End Sub
Public Sub LimpiaForm()
t_Grid = 0
ChkMone.Value = 0
LblParte.Caption = 0
LimpiaControl Me, "TextBox"
TxtFechEmis.Text = "  /  /    "
TxtFechEntr.Text = "  /  /    "
TxtIgv.Text = ""
TxtDscto.Text = ""
FGene.Enabled = False
FDeta.Enabled = False
FxgDeta.Enabled = False
FComandoDeta.Enabled = False
CmdEliminar.Enabled = False
CmdModificar.Enabled = False
CmdImprimir.Enabled = False
LimpiaVariable
LimpiaLista
FxgDeta.Clear
FxgDeta.FormatString = ".Item.|.   Codigo   .|.Pedido.|.                              Descripción                              .|.Conversión.|.  Cant  .|.    Precio    .|.    Dscto    .|.    Importe    .|.Tipo.|. UM .|.    Factor    .|.                    Observaciones                    ."
a = FxgCron.FormatString
FxgCron.Clear
FxgCron.FormatString = a
LimpiaDetalle
Ctr_Ccostos.xclave = "": Ctr_Ccostos.xnombre = ""
Ctr_CondPag.xclave = "": Ctr_CondPag.xnombre = ""
Ctr_TipoArt.xclave = "": Ctr_TipoArt.xnombre = ""
CtrAyu_Proveedor.xclave = "": CtrAyu_Proveedor.xnombre = ""
OptEstado(0).Value = False: OptEstado(1).Value = False: OptEstado(2).Value = False

'Me.Height = 8025
End Sub

Private Sub FxgDeta_DblClick()
On Error Resume Next
LblPosiEdit.Caption = 0
LblLote.Caption = ""
ChkUm.Value = 0
With FxgDeta
    .Col = 0
    LblPosiEdit.Caption = .Row
    .Col = 1    'Codigo
    Ctr_Art.xclave = .Text: Ctr_Art.Ejecutar
    .Col = 2    'Pedido
    TxtPedi.Text = .Text
    .Col = 3    'Descripción
    LblCodi.Caption = .Text
    .Col = 10   'Um
    Ctr_Uni.xclave = .Text: Ctr_Uni.Ejecutar
    .Col = 4    'Conversión
    LblUm.Caption = .Text
    .Col = 5    'Cantidad
    TxtCantPedi.Text = .Text
    .Col = 6    'Precio
    TxtPrec.Text = .Text
    .Col = 7    'Descuento Unitario
    TxtDscto2.Text = .Text
    .Col = 8    'Importe
    .Col = 9    'Tipo

    .Col = 11   'Factor de Conversión
    LblUmFact.Caption = .Text
    .Col = 12   'Comentario
    TxtCome.Text = .Text
End With
FLista.Visible = False
End Sub

Private Sub FxgDeta_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
With FxgDeta
    .Col = 0
    If Trim(.Text) = "" Then Exit Sub
    If Button = 2 Then PopupMenu Me.m_crono
End With
End Sub

Private Sub LblComp_Change()
On Error Resume Next
FxgDeta.Clear
FxgDeta.FormatString = ".Item.|.   Codigo   .|.Pedido.|.                              Descripción                              .|.Conversión.|.  Cant  .|.    Precio    .|.    Dscto    .|.    Importe    .|.Tipo.|. UM .|.    Factor    .|.                    Observaciones                    ."
End Sub
Private Sub LComp_DblClick()
On Error Resume Next
With LComp
    If .ListCount <= 0 Then Exit Sub
    cComp = LXCodi.List(.ListIndex)
    cLComp = .Text
    TxtComp.Text = cComp
    LblComp.Caption = cLComp
    LimpiaLista
    TxtComp.SetFocus
End With
End Sub

Private Sub LComp_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
LComp_DblClick
End Sub
Private Sub LMerc_DblClick()
On Error Resume Next
With LMerc
    If .ListCount <= 0 Then Exit Sub
    cMerc = LXCodi.List(.ListIndex)
    cLMerc = .Text
    TxtMerc.Text = cMerc
    LblMerc.Caption = cLMerc
    LimpiaLista
    TxtMerc.SetFocus
End With
End Sub

Private Sub LMerc_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
LMerc_DblClick
End Sub
Private Sub m_crono_actualiza_Click()
FDeta.Enabled = False
FComandoDeta.Enabled = False
LblCronItem.Caption = Val(FxgDeta.Row)
FxgDeta.Col = 5
LblCronTota.Caption = Val(FxgDeta.Text)
FxgDeta.Col = 11
LblCronUmFact.Caption = Val(FxgDeta.Text)
FCrono.Visible = True
FxgDeta.Enabled = False
FxgIndi.Col = 1
FxgIndi.Row = 1
FxgIndi.SetFocus
End Sub

Private Sub OptEstado_Click(Index As Integer)
    vlestado = Index + 1
End Sub

Private Sub TxtCantPedi_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
TxtPrec.SetFocus
End Sub

Private Sub TxtCome_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub TxtCont_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub TxtDscto_Change()
Totalizar
End Sub

Private Sub TxtDscto_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii <> 13 Then Exit Sub
Ctr_Art.SetFocus
End Sub

Private Sub TxtDscto2_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii <> 13 Then Exit Sub
CmdActualiza.Value = True
End Sub

Private Sub TxtFechEmis_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii <> 13 Then Exit Sub
If Not IsDate(TxtFechEmis) Then TxtFechEmis = Format(Date, "dd/mm/yyyy")
TxtFechEntr.SetFocus
End Sub

Private Sub TxtFechEmis_LostFocus()
If Not IsDate(TxtFechEmis.Text) Then TxtFechEmis.Text = "  /  /    "
End Sub

Private Sub TxtFechEntr_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
If Not IsDate(TxtFechEntr) Then TxtFechEntr = Format(Date, "dd/mm/yyyy")
TxtObse.SetFocus
End Sub

Private Sub TxtFechEntr_LostFocus()
If Not IsDate(TxtFechEntr.Text) Then TxtFechEntr.Text = "  /  /    "
End Sub

Private Sub TxtIgv_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii <> 13 Then Exit Sub
TxtDscto.SetFocus
End Sub

Private Sub TxtIgv_LostFocus()
a = TxtIgv.Text
If Not IsNumeric(a) Or Val(a) < 0 Or Val(a) >= 100 Then TxtIgv.Text = 18
Totalizar
End Sub

Private Sub TxtMerc_Change()
If Trim(TxtMerc.Text) = "" Then BuscaMercado
LblMerc.Caption = ""
End Sub

Private Sub TxtMerc_KeyPress(KeyAscii As Integer)
On Error Resume Next
'If KeyAscii <> 13 Then Exit Sub

If IsNumeric(TxtMerc.Text) Then
    If Val(TxtMerc.Text) > 2 Or Val(TxtMerc.Text) = 0 Then
       BuscaMercado
     Else
      Ctr_CondPag.SetFocus
    End If
 Else
    MsgBox "Ingreso valor menor a 3", vbCritical
    TxtMerc.SetFocus
End If
End Sub

Private Sub TxtObse_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
If Trim(TxtObse.Text) = "" Then TxtObse.Text = "Despachar en Calle San Ernesto 6326 - Los Olivos."
TxtIgv.SetFocus
End Sub

Private Sub TxtPedi_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
    'Ctr_Uni.Visible = True
    'Ctr_Uni.Enabled = True
    TxtCantPedi.SetFocus
    'Ctr_Uni.SetFocus
End Sub

Private Sub TxtPrec_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
TxtDscto2.SetFocus
End Sub
Public Sub LimpiaLista()
On Error Resume Next
LXCodi.Clear: FLista.Visible = False

LCodi.Clear: LCodi.Visible = False
LClas.Clear: LClas.Visible = False
LComp.Clear: LComp.Visible = False
LProv.Clear: LProv.Visible = False
LCost.Clear: LCost.Visible = False
LMerc.Clear: LMerc.Visible = False
LPago.Clear: LPago.Visible = False
LUnid.Clear: LUnid.Visible = False
LUnid2.Clear: LUnid2.Visible = False

End Sub

Private Sub LimpiaVariable()
t_Grid = 0
cCodi = "": cLCodi = ""
cClas = "": cLClas = ""
cComp = "": cLComp = ""
cProv = "": cLProv = ""
cCost = "": cLCost = ""
cMerc = "": cLMerc = ""
cPago = "": cLPago = ""
cEnvi = "": cLEnvi = ""
cUnid = "": cLUnid = "": cL2Unid = ""
End Sub

Public Sub BuscaMercado()
On Error Resume Next
LimpiaVariable
LimpiaLista
LXCodi.AddItem 1
LMerc.AddItem "Nacional"
LXCodi.AddItem 2
LMerc.AddItem "Extranjero"
LMerc.Visible = True
FLista.Visible = True
LMerc.SetFocus
End Sub

Public Sub BuscaClaseCompra()
On Error Resume Next
LimpiaVariable
LimpiaLista
LXCodi.AddItem 0
LClas.AddItem "Bienes"
LXCodi.AddItem 1
LClas.AddItem "Todos"
LClas.Visible = True
FLista.Visible = True
LClas.SetFocus
End Sub

Public Sub BuscaTipoCompra()
On Error Resume Next

Dim Rs As New ADODB.Recordset
Set Rs = DEEmpresas.cnnEmpresas.Execute("select * from bienes")

With Rs
    LimpiaVariable
    LimpiaLista
    a = 0
    Do While Not .EOF
       a = a + 1
       If a = 1 Then cComp = .Fields(0): cLComp = .Fields(1)
       LXCodi.AddItem .Fields(0)
       LComp.AddItem LCase(.Fields(1))
        .MoveNext
    Loop
    If a <> 0 Then FLista.Visible = True: LComp.Visible = True: LComp.SetFocus
End With
End Sub
Public Sub BuscaCentroCostos()
On Error Resume Next
resp = InputBox$("Ingrese parte del nombre a Buscar")
If Trim(resp = "") Then Exit Sub
Dim rFiltro As String
rFiltro = resp

With DEEmpresas
    If .rsCmdBuscaCentroCostos.State = adStateOpen Then .rsCmdBuscaCentroCostos.Close
    .CmdBuscaCentroCostos rFiltro, 2
    Dim Rs As New ADODB.Recordset
    Set Rs = .rsCmdBuscaCentroCostos
    LimpiaVariable
    LimpiaLista
    With Rs
        Do While Not .EOF
            a = a + 1
            If a = 1 Then cCost = .Fields(0): cLCost = .Fields(1)
            LXCodi.AddItem .Fields(0)
            LCost.AddItem LCase(.Fields(1))
            .MoveNext
        Loop
        If a <> 0 Then FLista.Visible = True: LCost.Visible = True: LCost.SetFocus
    End With
End With
End Sub
Public Sub LimpiaDetalle()
On Error Resume Next
a = FxgIndi.FormatString
FxgIndi.Clear
FxgIndi.FormatString = 1
Ctr_Art.xclave = ""
Ctr_Art.xnombre = ""
ChkUm.Value = 0
LblPosiEdit.Caption = 0
LblCodi.Caption = ""
TxtPedi.Text = ""
Ctr_Uni.xclave = ""
Ctr_Uni.xnombre = ""
TxtPrec.Text = ""
TxtCantPedi.Text = ""
TxtDscto2.Text = ""
TxtCome.Text = ""
Totalizar
End Sub

Public Sub AgregaDetalle()
On Error GoTo ErrDeta
VerificaDetalle
If Veri <> 0 Then Exit Sub

If LblPosiEdit.Caption = 0 Then
    BuscaPos
    xPos = Val(LblPosi.Caption)
Else
    xPos = Val(LblPosiEdit.Caption)
End If

With FxgDeta
    If .Row = .Rows - 1 Then Exit Sub
    .Row = xPos
    .Col = 0    'Items
    .Text = xPos
    .Col = 1    'Código
    .Text = Trim(Ctr_Art.xclave)
    .Col = 2    'Pedido
    .Text = Trim(TxtPedi.Text)
    .Col = 3    'Descripción
    .Text = " " & Trim$(Ctr_Art.xnombre)
    .Col = 4    'Conversión
    .Text = Ctr_Uni.xnombre
    a = Val(TxtCantPedi.Text)
    b = Val(TxtPrec.Text)
    c = Val(TxtDscto2.Text)
    e = Val(LblUmFact.Caption)
    
    If ChkUm.Value = 1 Then
        'Si está utilizando Unid Compra, convertir a Unid Almac
        a = a / e
        b = b / e
    End If
    
    d = (a * b) * (1 - c / 100)
    
    .Col = 5    'Cantidad
    .Text = Format(a, "#0.00")
    .Col = 6    'Precio
    .Text = Format(b, "#0.00")
    .Col = 7    'Descuento
    .Text = Format(c, "#0.00")
    .Col = 8    'Importe
    .Text = Format(d, "#0.00")
    .Col = 9    'Tipo
    .Text = Val(Ctr_TipoArt.xclave)
    .Col = 10   'Um
    .Text = Val(Ctr_Uni.xclave)
    .Col = 11   'Factor
    .Text = e
    .Col = 12   'Comentario
    .Text = Trim(TxtCome.Text)
End With

LimpiaDetalle
LimpiaVariable
Ctr_Art.SetFocus

Exit Sub
ErrDeta:
MsgBox "Se produjeron errores al realizar el procedimiento", vbCritical
'Resume
End Sub

Public Sub VerificaDetalle()
On Error GoTo ErrVeri
Veri = 0


If Not IsNumeric(TxtPrec.Text) Then TxtPrec.Text = 0
If Not IsNumeric(TxtCantPedi.Text) Then TxtCantPedi.Text = 0
If Not IsNumeric(TxtDscto2.Text) Then TxtDscto2.Text = 0

If Val(LblPosiEdit.Caption) <> 0 Then
    SumarSiGrid Me.FxgCron, 2, 0, Val(LblPosiEdit.Caption)
    If t_Grid > Val(TxtCantPedi.Text) Then Veri = 7: xMsg = "La cantidad que está especificando es menor al que figura en el cronograma" & vbNewLine & "Corrija su cronograma a fin de permitir este cambio"
End If

If Trim(Ctr_TipoArt.xclave) = "" Then Veri = 6: xMsg = "No ha definido el tipo de artículo": Ctr_TipoArt.SetFocus
If Val(TxtDscto2.Text) > 100 Or Val(TxtDscto2.Text) < 0 Then Veri = 5: xMsg = "No es un porcentaje de descunto válido": TxtDscto2.SetFocus
If Val(TxtCantPedi.Text) <= 0 Then Veri = 4: xMsg = "La cantidad pedida no es válida": TxtCantPedi.SetFocus
If Val(TxtPrec.Text) < 0 Then Veri = 3: xMsg = "No es un precio de compra válido": TxtPrec.SetFocus
'If Ctr_Uni.xclave = "" Then Veri = 2: xMsg = "Falta indicar la unidad de transacción": Ctr_Uni.SetFocus
If Ctr_Art.xclave = "" Then Veri = 1: xMsg = "Falta colocar el código del artículo": Ctr_Art.SetFocus

mVeri:
If Veri <> 0 Then MsgBox xMsg, vbCritical, "Error en Verificación de Detalle"

Exit Sub
ErrVeri:
Veri = 1: xMsg = "Error grave al intentar verificar el detalle"
GoTo mVeri
End Sub
Public Function SumarSiGrid(grid As MSFlexGrid, ColuSuma, SiColu As Integer, SiValo As Double)
On Error Resume Next
t_Grid = 0
a = 0
With grid
    For i = 1 To .Rows - 1
        .Row = i
        .Col = Val(SiColu)
        If Val(.Text) = SiValo Then .Col = ColuSuma: a = a + Val(.Text)
    Next
    t_Grid = a
End With
End Function
Public Sub BuscaPago()
On Error Resume Next
Dim Rs As New ADODB.Recordset

Set Rs = DEEmpresas.cnnEmpresas.Execute("Select * from PagoCondicion")
LimpiaVariable
LimpiaLista
a = 0
With Rs
    Do While Not .EOF
        a = a + 1
        cPago = .Fields(0): cLPago = .Fields(1)
        LXCodi.AddItem .Fields(0)
        LPago.AddItem .Fields(1)
        .MoveNext
    Loop
End With
If a <> 0 Then FLista.Visible = True: LPago.Visible = True: LPago.SetFocus
End Sub
Public Sub LlenaMercado()
On Error Resume Next
If cMerc <> "" Then TxtMerc.Text = cMerc
If cLMerc <> "" Then LblMerc.Caption = cLMerc
End Sub

Public Sub BuscaPos()
FxgDeta.Col = 0
X = 0
Do While Y = 0
    X = X + 1
    FxgDeta.Row = X
    If Trim(FxgDeta.Text) = "" Then Y = 1
Loop
If X = 1 Then LblPosi.Caption = 1 Else LblPosi.Caption = X
End Sub

Public Sub Totalizar()
TotalizarGrid Me.FxgDeta, 8
LblImp.Caption = Format(t_Grid, "#0.00")
LblDscto.Caption = Format((t_Grid * (Val(TxtDscto.Text) / 100)), "#0.00")
a = Val(LblImp.Caption) - Val(LblDscto.Caption)
b = a * (Val(TxtIgv.Text) / 100)
LblIgv.Caption = Format(b, "#0.00")
LblTota.Caption = Format(a + b, "#0.00")
End Sub
Public Function TotalizarGrid(grid As MSFlexGrid, Colu As Integer)
On Error Resume Next
t_Grid = 0
a = 0
grid.Col = Colu
For i = 1 To grid.Rows - 1
    grid.Row = i
    a = a + Val(grid.Text)
Next
t_Grid = a
End Function
Public Sub VerificaData()
On Error GoTo ErrVeri
Veri = 0

'Verifica Cronograma
'Si alguna de las fechas del cronograma estan fuera del
'intervalo de Fechas no continuar

If Not IsNumeric(TxtIgv.Text) Then TxtIgv.Text = 18
If Not IsNumeric(TxtDscto.Text) Then TxtDscto.Text = 0

If IsDate(TxtFechEmis.Text) And IsDate(TxtFechEntr.Text) Then
    a = CDate(TxtFechEmis.Text)
    b = CDate((TxtFechEntr.Text))
    xa = DateDiff("d", a, b)
    If xa < 1 Then Veri = 8: xMsg = "La fecha de entrega no puede ser anterior o igual al de emisión": TxtFechEmis.SetFocus
End If
If Not IsDate(TxtFechEntr.Text) Then Veri = 7: xmg = "No es una fecha de entrega válida": TxtFechEntr.SetFocus
If Not IsDate(TxtFechEmis.Text) Then Veri = 6: xmg = "No es una fecha de emisión válida": TxtFechEmis.SetFocus
If Ctr_CondPag.xclave = "" Then Veri = 5: xMsg = "No ha colocado la condición de pago": Ctr_CondPag.SetFocus
If LblMerc.Caption = "" Then Veri = 4: xMsg = "Falta el tipo de mercado al cual va dirigido la Orden de Compra": TxtMerc.SetFocus
If Trim(TxtCont.Text) = "" Then Veri = 3: xMsg = "Falta definir el representante del proveedor": TxtCont.SetFocus

If CtrAyu_Proveedor.xclave = "" Then Veri = 2: xMsg = "Falta colocar el proveedor": CtrAyu_Proveedor.SetFocus

If Ctr_TipoArt.xclave = "" Then Veri = 1: xMsg = "No ha definido el tipo de artículo a solicitar": Ctr_TipoArt.SetFocus

If Veri <> 0 Then MsgBox xMsg, vbCritical, "Error " & Veri
Exit Sub
ErrVeri:
Veri = 99: MsgBox "No se pudo completar la verificación de los datos", vbCritical
End Sub

Public Sub GuardaDocumento()
'On Error GoTo ErrGuar
VerificaData
If Veri <> 0 Then Exit Sub

'Capturar Valores del Encabezado
gTipo = Val(Ctr_TipoArt.xclave)
gProv = CtrAyu_Proveedor.xclave
gRepr = Trim$(TxtCont.Text)
gCost = Val(Ctr_Ccostos.xclave)
gMerc = Val(TxtMerc.Text)
gCond = Val(Ctr_CondPag.xclave)
gFech = CDate(TxtFechEmis.Text)
gFechEntr = CDate(TxtFechEntr.Text)
gMone = ChkMone.Value
gObse = Trim$(TxtObse.Text)
gIgvF = Val(TxtIgv.Text)
gIgv = Val(LblIgv.Caption)
gDsctoF = Val(TxtDscto.Text)
gDscto = Val(LblDscto.Caption)
gImpo = Val(LblImp.Caption)
gNeto = Val(LblTota.Caption)

DEData.CnxVg.BeginTrans

If LblParte.Caption <> 0 Then
    CADENA = "Delete from OrdenCompra where OrdenNro=" & Val(LblParte.Caption): VGcnx.Execute (CADENA)
    CADENA = "Delete from OrdenCompraDetalle where OrdenNro=" & Val(LblParte.Caption): VGcnx.Execute (CADENA)
    CADENA = "Delete from OrdenCompraControl where TranTipo=0 and OrdenNro=" & Val(LblParte.Caption): VGcnx.Execute (CADENA)
    CADENA = "Delete from OrdenCompraCronograma where OrdenNro=" & Val(LblParte.Caption): VGcnx.Execute (CADENA)
End If

'Captura número de Operación
Dim NumDoc As Integer


If Val(LblParte.Caption) = 0 Then resp = DEData.cmdParametrosUpdate(1, "CompraNro", NumDoc): LblParte.Caption = Val(NumDoc)
gParte = Val(LblParte.Caption)

'Inserta el documento
DEData.CmdOcInserta gParte, gTipo, gProv, gRepr, gCost, gMerc, 1, gCond, Date, gFech, gFechEntr, gMone, gIgvF, gImpo, gDsctoF, gDscto, gNeto, gIgv, gObse, vlestado
InsertaDetalle
DEData.CnxVg.CommitTrans

For i = 0 To 100
    PBar.Value = i
Next



xMsg = "Se ha generado la Orden de Compra " & LblParte.Caption
MsgBox xMsg, vbInformation + vbOKOnly
LimpiaForm
CmdNuevo.SetFocus


Exit Sub
ErrGuar:
MsgBox "No se pudo completar la transacción. No se ha actualizado la base de datos principal", vbCritical
End Sub

Public Sub InsertaDetalle()
'On Error GoTo ErrDeta
With FxgDeta
    .Row = 1
    Do While Not .Rows - 1
        .Col = 0
        For i = 0 To .Cols - 1
            .Col = i
            xDeta(i) = .Text
        Next
        .Col = 0
        If Trim(.Text) <> "" Then
            'Registra el detalle del documento
            DEData.CmdOcInsertaDetalle Val(gParte), Val(xDeta(0)), Trim(xDeta(1)), xDeta(2), Val(xDeta(5)), xDeta(6), xDeta(7), gTipo, Val(xDeta(10)), Trim$(xDeta(12))
            'Insertar datos del Control de Orden de Compra
            DEData.CmdOccInsertar Val(gParte), Val(gTipo), Trim(xDeta(1)), xDeta(2), " ", 0, " ", gFech, 0, Val(xDeta(5)), 0, xDeta(6)
            'Insertar Cronograma Individual
            k = 0
        End If
        For i = 0 To 100
            PBar.Value = i
        Next i
        If .Row = .Rows - 1 Then Exit Do
        .Row = .Row + 1
    Loop
End With
InsertaCronograma Val(gParte)
For i = 0 To 100
    PBar.Value = i
Next i
PBar.Visible = False
        
Exit Sub
ErrDeta:
MsgBox "Error en detalle", vbCritical
End Sub

Public Sub BuscaDocumento()
On Error GoTo ErrBusc
CmdModificar.Enabled = False
CmdEliminar.Enabled = False
CmdImprimir.Enabled = False
CmdGuardar.Enabled = True

resp = InputBox("Ingrese Nro de Orden de Compra")
If Not IsNumeric(resp) Then Exit Sub
Dim Rs As New ADODB.Recordset
Set Rs = VGcnx.Execute("Select * from OrdenCompra where OrdenNro=" & Val(resp))
a = 0
LimpiaForm

With Rs
    
    Do While Not .EOF
        'Datos de Cabecera
        a = a + 1
        LblParte.Caption = .Fields(0)       'Parte
        
        Ctr_TipoArt.xclave = .Fields(1)      'Tipo de Artículo
        Call Ctr_TipoArt.Ejecutar
        
        CtrAyu_Proveedor.xclave = .Fields(2) 'Proveedor
        Call CtrAyu_Proveedor.Ejecutar
        
        TxtCont.Text = Trim$(.Fields(3))    'Representante
        
        Ctr_Ccostos.xclave = .Fields(4)     'Centro de Costos
        Call Ctr_Ccostos.Ejecutar
                
        TxtMerc.Text = .Fields(5)           'Mercado
        TxtMerc_KeyPress (13)
        
        Ctr_CondPag.xclave = .Fields(7)     'Condiciones de Pago
        Call Ctr_CondPag.Ejecutar
              
        TxtFechEmis.Text = Format(.Fields(9), "dd/mm/yyyy")     'Fecha de Emisión
        TxtFechEntr.Text = Format(.Fields(10), "dd/mm/yyyy")    'Fecha de Entrega
        ChkMone.Value = .Fields(11)         'Moneda
        TxtObse.Text = .Fields(18)          'Observaciones
        Select Case .Fields(19)
            Case 1:
                OptEstado(0).Value = True: OptEstado(1).Value = False: OptEstado(2).Value = False
            Case 2:
                OptEstado(0).Value = False: OptEstado(1).Value = True: OptEstado(2).Value = False
            Case 3:
                OptEstado(0).Value = False: OptEstado(1).Value = False: OptEstado(2).Value = True
        End Select
        TxtIgv.Text = .Fields(12)           'Igv
        TxtDscto.Text = .Fields(15)         'Dscto
        .MoveNext
    Loop
End With

If a = 0 Then MsgBox "No se ha ubicado el documento": Exit Sub     'Si no ubicó el documento obviar el detalle

If DEData.rsCmdOccBuscaDetalle.State = 1 Then DEData.rsCmdOccBuscaDetalle.Close

DEData.CmdOccBuscaDetalle Val(LblParte.Caption), Val(Tipcom)
Set Rs = DEData.rsCmdOccBuscaDetalle
Z = 0

Dim Ip, It, It2 As Integer
'On Error Resume Next
With Rs
    Do While Not .EOF
        'Datos del detalle
        Z = Z + 1
        With FxgDeta
            .Row = Z
            
            Ip = Rs.Fields(0) 'OrdenNro
            It = Rs.Fields(1) 'Item
            It2 = Val(Z)
            Me.BuscaCronograma Ip, It, It2
            
            .Col = 0        'Items
            .Text = Z
            .Col = 1        'Código
            .Text = Rs.Fields(2)
            .Col = 2        'Pedido
            .Text = Rs.Fields(3)
            .Col = 3        'Descripción
            .Text = " " & Rs.Fields(10)
            .Col = 4        'Conversión
            .Text = Rs.Fields(11)
            
            a = Val(Rs.Fields(4))
            b = Val(Rs.Fields(5))
            c = Val(Rs.Fields(6))
            d = (a * b) * (1 - c / 100)
            
            .Col = 5        'Cantidad
            .Text = Format(a, "#0.00")
            .Col = 6        'Precio
            .Text = Format(b, "#0.00")
            .Col = 7        'Descuento
            .Text = Format(c, "#0.00")
            .Col = 8        'Importe
            .Text = Format(d, "#0.00")
            .Col = 9        'Tipo
            .Text = Val(Rs.Fields(7))
            .Col = 10       'Um
            .Text = Val(Rs.Fields(8))
            .Col = 11       'Factor
            .Text = Val(Rs.Fields(12))
            .Col = 12       'Comentario
            .Text = Trim(Rs.Fields(9))
        End With
        .MoveNext
    Loop
End With

CmdModificar.Enabled = True
CmdEliminar.Enabled = True
CmdImprimir.Enabled = True
CmdGuardar.Enabled = False

Totalizar

Exit Sub
ErrBusc:
MsgBox "No se pudo realizar la búsqueda", vbCritical
End Sub

'====================================================
'<<<<<<<< Cronograma de Entregas Parciales >>>>>>>>>>
'====================================================

Private Sub CmdConfirmar_Click()
VerificarCronograma
If Veri <> 0 Then Exit Sub

c = Val(LblCronUmFact.Caption)
With ChkCronUm
    If .Value = 1 Then
        .BackColor = RGB(0, 150, 150)
        .Caption = "Unid Compra"
        FxgIndi.Col = 2
        For i = 1 To FxgIndi.Rows - 1
            FxgIndi.Row = i
            d = Val(FxgIndi.Text)
            If d > 0 Then FxgIndi.Text = d / c
        Next
    End If
End With

'Eliminar Datos Anteriores en destino
EliminarIndividual Val(LblCronItem.Caption)

'Ordenar Registros en destino
OrdenarCronograma

'Transferir nuevos datos
PasarInidividual

'Eliminar Datos Individuales
LimpiaIndividual

CmdCronSalir.Value = True
End Sub

Private Sub CmdProbar_Click()
resp = InputBox("Ingrese Nro de Item")
LblCronItem.Caption = resp

resp = InputBox("Ingrese Cantidad x Item")
LblCronTota.Caption = resp
End Sub

Private Sub FxgIndi_EnterCell()
On Error Resume Next
With FxgIndi
    If .Col = 0 Then Exit Sub
    .CellBackColor = RGB(24, 92, 168)
    .CellForeColor = RGB(255, 255, 255)
End With
End Sub

Private Sub FxgIndi_KeyPress(KeyAscii As Integer)
On Error Resume Next
'If KeyAscii <> 13 Then Exit Sub
TxtCronFech.Text = "  /  /    "
With FxgIndi
    If .Col = 1 Then TxtCronFech.Text = .Text: TxtCronFech.Visible = True: TxtCronFech.SetFocus
    If .Col = 2 Then TxtCronCant.Text = .Text: TxtCronCant.Visible = True: TxtCronCant.SetFocus
End With
End Sub

Private Sub FxgIndi_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode <> vbKeyDelete Then Exit Sub
With FxgIndi
    a = .Row
    .Col = 1
    .Text = ""
    .Col = 2
    .Text = ""
    Me.TotalCrono
    .Row = a
End With
End Sub

Private Sub FxgIndi_LeaveCell()
With FxgIndi
    If .Col = 0 Then Exit Sub
    .CellBackColor = RGB(255, 255, 255)
    .CellForeColor = RGB(0, 0, 0)
End With
End Sub

Private Sub FxgIndi_RowColChange()
With FxgIndi
    If .Col = 1 Then
        TxtCronFech.Width = .CellWidth
        TxtCronFech.Height = .CellHeight
        TxtCronFech.Top = .Top + .CellTop
        TxtCronFech.Left = .Left + .CellLeft
    End If
    
    If .Col = 2 Then
        TxtCronCant.Width = .CellWidth
        TxtCronCant.Height = .CellHeight
        TxtCronCant.Top = .Top + .CellTop
        TxtCronCant.Left = .Left + .CellLeft
    End If
End With
End Sub

Private Sub LblCronItem_Change()
If Val(LblCronItem.Caption) = 0 Then Exit Sub
a = FxgIndi.FormatString
FxgIndi.Clear
FxgIndi.FormatString = a
RecuperarCronograma Val(LblCronItem.Caption)
TotalCrono
End Sub

Private Sub TxtCronFech_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
a = Trim$(TxtCronFech.Text)
If Not IsDate(a) Then
    d = ""
Else
    d = Format(a, "dd/mm/yyyy")
    a = DateDiff("d", TxtFechEmis.Text, d)
    b = DateDiff("d", d, TxtFechEntr.Text)
    c = DateDiff("d", TxtFechEmis.Text, TxtFechEntr.Text)
    If (a < 0) Or (a > c) Then d = ""
End If
FxgIndi.Text = d
TxtCronFech.Visible = False
FxgIndi.Col = 2
FxgIndi.SetFocus
End Sub

Private Sub TxtCronCant_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
a = Trim(TxtCronCant.Text)
If Not IsNumeric(a) Then a = ""
FxgIndi.Text = Format(a, "#0.00")
TxtCronCant.Visible = False
b = FxgIndi.Row
TotalCrono
FxgIndi.Row = b
FxgIndi.SetFocus
End Sub

Public Sub TotalCrono()
TotalizarGrid Me.FxgIndi, 2
LblCronProg.Caption = t_Grid
End Sub

Public Sub VerificarCronograma()
On Error GoTo ErrCrono
Veri = 0

'Verifica elementos registrados
With FxgIndi
    For i = 1 To .Rows - 1
        .Row = i
        .Col = 1
        If Not IsDate(.Text) Then
Eliminar:
            .Text = "": .Col = 2: .Text = "": .Col = 0: .Text = ""
        Else
            '.Col = 0: .Text = LblCronItem.Caption
            a = DateDiff("d", TxtFechEmis.Text, .Text)
            b = DateDiff("d", .Text, TxtFechEntr.Text)
            c = DateDiff("d", TxtFechEmis.Text, TxtFechEntr.Text)
            If (a < 0) Or (a > c) Then GoTo Eliminar
        End If
        .Col = 2
        If Val(.Text) <= 0 Then .Text = "": .Col = 1: .Text = ""
    Next
    TotalCrono
    .Row = 1
End With

If Val(LblCronTota.Caption) < Val(LblCronProg.Caption) Then Veri = 1: xMsg = "El Total Programado no puede ser mayor al Pedido"

If Veri <> 0 Then MsgBox "Error en Cronograma" & vbNewLine & xMsg, vbCritical
Exit Sub
ErrCrono:
Veri = 1
MsgBox "No se pudo terminar la verificación del cronograma", vbCritical
End Sub

Public Sub PasarInidividual()
With FxgIndi
    For i = 1 To .Rows - 1
        .Row = i
        .Col = 1
        If Trim(.Text) <> "" Then
            a = .Text: .Col = 2: b = .Text
            LblCronPosi.Caption = Val(LblCronPosi.Caption) + 1
            With FxgCron
                Z = Val(LblCronPosi.Caption)
                If Z = .Rows Then Z = .Rows - 1
                .Row = Z
                .Col = 0
                .Text = LblCronItem.Caption
                .Col = 1
                .Text = a
                .Col = 2
                .Text = b
            End With
        End If
    Next
End With
End Sub

Public Sub LimpiaIndividual()
FxgIndi.Clear
FxgIndi.FormatString = ".Item.|.      Fecha      .|.      Cantidad      ."
LblCronItem.Caption = 0
LblCronTota.Caption = 0
LblCronProg.Caption = 0
TxtCronFech.Text = "  /  /    "
TxtCronCant.Text = ""
End Sub

Public Function EliminarIndividual(fOrden As Integer)
With FxgCron
    For i = 1 To .Rows - 1
        .Row = i
        .Col = 0
        If Trim(.Text) = fOrden Then
            .Col = 0
            .Text = ""
            .Col = 1
            .Text = ""
            .Col = 2
            .Text = ""
        End If
    Next
End With
End Function

Public Sub OrdenarCronograma()
Dim Rec0(100) As String
Dim Rec1(100) As String
Dim Rec2(100) As String

'Pasar solo datos Completos al Arreglo
d = 0
With FxgCron
    For i = 1 To .Rows - 1
        .Row = i
        .Col = 0
        If Trim(.Text) <> "" Then
            d = d + 1 'Indice del Arreglo
            .Col = 0
            Rec0(d) = .Text
            .Col = 1
            Rec1(d) = .Text
            .Col = 2
            Rec2(d) = .Text
        End If
    Next

    'Limpiar Grid
    b = .FormatString
    .Clear
    .FormatString = b

    'Pasar Datos del Arreglo al Grid
    c = 0
    For i = 1 To FxgCron.Rows - 1
        .Row = i
        .Col = 0
        If Rec0(i) = "" Then LblCronPosi.Caption = i - 1: Exit Sub
        .Text = Rec0(i)
        .Col = 1
        .Text = Rec1(i)
        .Col = 2
        .Text = Rec2(i)
    Next
End With
End Sub

Public Sub RecuperarCronograma(xItem As Integer)
b = FxgIndi.FormatString: FxgIndi.Clear: FxgIndi.FormatString = b
a = 0
With FxgCron
    For i = 1 To .Rows - 1
        .Row = i
        .Col = 0
        If Val(.Text) = xItem Then
            a = a + 1
            FxgIndi.Row = a
            .Col = 0
            FxgIndi.Col = 0: FxgIndi.Text = .Text
            .Col = 1
            FxgIndi.Col = 1: FxgIndi.Text = .Text
            .Col = 2
            FxgIndi.Col = 2: FxgIndi.Text = .Text
        End If
    Next
End With
End Sub

Public Function InsertaCronograma(crOrden As Integer)
'On Error GoTo ErrCron
Dim crItem As Integer
Dim crFech As String
Dim crCant As Double
With FxgCron
    For i = 1 To .Rows - 1
        .Row = i
        .Col = 0
        crItem = Val(.Text)
        .Col = 1
        crFech = .Text
        .Col = 2
        crCant = Val(.Text)
        If crItem <> 0 Then DEData.CmdOcCrInserta crOrden, crItem, crFech, crCant
    Next
End With

Exit Function
ErrCron:
MsgBox "No se pudo insertar el cronograma", vbCritical
End Function

Public Function BuscaCronograma(gOrden, gItem, gItemDest As Integer)
Dim Rs As New ADODB.Recordset
CADENA = "Select * from OrdenCompraCronograma where OrdenNro=" & gOrden & " and Item=" & gItem & " order by Fecha"
Set Rs = VGcnx.Execute(CADENA)
'Insertar en el Grid Crono la relación de Cronogramama
With FxgCron
    Do While Not Rs.EOF
        a = Val(LblCronPosi.Caption) + 1
        .Row = a
        .Col = 0
        .Text = gItemDest
        .Col = 1
        .Text = Rs.Fields(2)
        .Col = 2
        .Text = Rs.Fields(3)
        Rs.MoveNext
        LblCronPosi.Caption = a
    Loop
End With

End Function

Public Function BorraCronograma(gItem As Integer)
With FxgCron
    For i = 1 To .Rows - 1
        .Row = i
        .Col = 0
        If Val(.Text) = gItem Then
            .Text = ""
            .Col = 1
            .Text = ""
            .Col = 2
            .Text = ""
        End If
    Next
End With
End Function


