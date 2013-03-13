VERSION 5.00
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Begin VB.Form FrmCierreDiario 
   Caption         =   "Cierre Diario"
   ClientHeight    =   6990
   ClientLeft      =   60
   ClientTop       =   555
   ClientWidth     =   10995
   ForeColor       =   &H00FF0000&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6990
   ScaleWidth      =   10995
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Caption         =   "Todos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   3150
      TabIndex        =   20
      Top             =   900
      Width           =   1005
   End
   Begin VB.Frame Frame1 
      Height          =   4440
      Left            =   90
      TabIndex        =   19
      Top             =   2340
      Width           =   10755
      Begin VB.CommandButton CmdImprimir 
         Caption         =   "Imprimir Saldos"
         Height          =   465
         Left            =   270
         TabIndex        =   24
         Top             =   3780
         Width           =   2625
      End
      Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
         Height          =   3225
         Left            =   120
         TabIndex        =   21
         Top             =   360
         Width           =   10530
         _ExtentX        =   18574
         _ExtentY        =   5689
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "C/B"
         Columns(0).DataField=   "tipocajabanco"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Empresa"
         Columns(1).DataField=   "empresadescripcion"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Codigo"
         Columns(2).DataField=   "cajabanco"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Descripcion"
         Columns(3).DataField=   "descripcion"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Moneda/Cta. Banco"
         Columns(4).DataField=   "moneda"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "S. inicial"
         Columns(5).DataField=   "saldoinicial"
         Columns(5).EditMask=   "####,###.##"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "Ingresos Mes"
         Columns(6).DataField=   "totalingresos"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "Egresos Mes"
         Columns(7).DataField=   "totalegresos"
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(8)._VlistStyle=   0
         Columns(8)._MaxComboItems=   5
         Columns(8).Caption=   "Saldo Final"
         Columns(8).DataField=   "saldofinal"
         Columns(8).NumberFormat=   "Currency"
         Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   9
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   13160660
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=9"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=688"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=609"
         Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(5)=   "Column(1).Width=2725"
         Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2646"
         Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(9)=   "Column(2).Width=1032"
         Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=953"
         Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(13)=   "Column(3).Width=2752"
         Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=2672"
         Splits(0)._ColumnProps(16)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(17)=   "Column(4).Width=2725"
         Splits(0)._ColumnProps(18)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(19)=   "Column(4)._WidthInPix=2646"
         Splits(0)._ColumnProps(20)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(21)=   "Column(5).Width=1667"
         Splits(0)._ColumnProps(22)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(23)=   "Column(5)._WidthInPix=1588"
         Splits(0)._ColumnProps(24)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(25)=   "Column(6).Width=1826"
         Splits(0)._ColumnProps(26)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(27)=   "Column(6)._WidthInPix=1746"
         Splits(0)._ColumnProps(28)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(29)=   "Column(7).Width=2302"
         Splits(0)._ColumnProps(30)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(31)=   "Column(7)._WidthInPix=2223"
         Splits(0)._ColumnProps(32)=   "Column(7).Order=8"
         Splits(0)._ColumnProps(33)=   "Column(8).Width=2725"
         Splits(0)._ColumnProps(34)=   "Column(8).DividerColor=0"
         Splits(0)._ColumnProps(35)=   "Column(8)._WidthInPix=2646"
         Splits(0)._ColumnProps(36)=   "Column(8).Order=9"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   3
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         DefColWidth     =   0
         HeadLines       =   1
         FootLines       =   1
         MultipleLines   =   0
         CellTipsWidth   =   0
         DeadAreaBackColor=   13160660
         RowDividerColor =   13160660
         RowSubDividerColor=   13160660
         DirectionAfterEnter=   1
         MaxRows         =   250000
         ViewColumnCaptionWidth=   0
         ViewColumnWidth =   0
         _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
         _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
         _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(15)  =   ":id=3,.fontname=MS Sans Serif"
         _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
         _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1"
         _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
         _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
         _StyleDefs(27)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(30)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=32,.parent=13"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=28,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
         _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=50,.parent=13"
         _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
         _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
         _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=54,.parent=13"
         _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
         _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
         _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
         _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=58,.parent=13"
         _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14"
         _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15"
         _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
         _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=62,.parent=13"
         _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=59,.parent=14"
         _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=60,.parent=15"
         _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=61,.parent=17"
         _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=66,.parent=13"
         _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=63,.parent=14"
         _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=64,.parent=15"
         _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=65,.parent=17"
         _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=70,.parent=13"
         _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=67,.parent=14"
         _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=68,.parent=15"
         _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=69,.parent=17"
         _StyleDefs(72)  =   "Named:id=33:Normal"
         _StyleDefs(73)  =   ":id=33,.parent=0"
         _StyleDefs(74)  =   "Named:id=34:Heading"
         _StyleDefs(75)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(76)  =   ":id=34,.wraptext=-1"
         _StyleDefs(77)  =   "Named:id=35:Footing"
         _StyleDefs(78)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(79)  =   "Named:id=36:Selected"
         _StyleDefs(80)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(81)  =   "Named:id=37:Caption"
         _StyleDefs(82)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(83)  =   "Named:id=38:HighlightRow"
         _StyleDefs(84)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(85)  =   "Named:id=39:EvenRow"
         _StyleDefs(86)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(87)  =   "Named:id=40:OddRow"
         _StyleDefs(88)  =   ":id=40,.parent=33"
         _StyleDefs(89)  =   "Named:id=41:RecordSelector"
         _StyleDefs(90)  =   ":id=41,.parent=34"
         _StyleDefs(91)  =   "Named:id=42:FilterBar"
         _StyleDefs(92)  =   ":id=42,.parent=33"
      End
   End
   Begin VB.Frame fraDetallado 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   180
      TabIndex        =   7
      Top             =   900
      Width           =   8910
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyudaBancoCuenta 
         Height          =   315
         Left            =   2880
         TabIndex        =   8
         Top             =   855
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   556
         XcodMaxLongitud =   4
         xcodwith        =   800
         NomTabla        =   "ct_subasiento"
         ListaCampos     =   "subasientocodigo(1),subasientodescripcion(1)"
         XcodCampo       =   "subasientocodigo"
         XListCampo      =   "subasientodescripcion"
         ListaCamposDescrip=   "Código,Descripción"
         ListaCamposText =   "subasientocodigo,subasientodescripcion"
         Requerido       =   0   'False
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyudaBanco 
         Height          =   300
         Left            =   2880
         TabIndex        =   9
         Top             =   360
         Width           =   5565
         _ExtentX        =   9816
         _ExtentY        =   529
         XcodMaxLongitud =   3
         xcodwith        =   800
         NomTabla        =   "ct_asiento"
         ListaCampos     =   "asientocodigo(1),asientodescripcion(1)"
         XcodCampo       =   "asientocodigo"
         XListCampo      =   "asientodescripcion"
         ListaCamposDescrip=   "Código,Descripción"
         ListaCamposText =   "asientocodigo,asientodescripcion"
         Requerido       =   0   'False
      End
      Begin MSComCtl2.DTPicker DTPickerFecFinal 
         Height          =   300
         Left            =   675
         TabIndex        =   10
         Top             =   810
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   529
         _Version        =   393216
         Format          =   20971521
         CurrentDate     =   37474
      End
      Begin MSComCtl2.DTPicker DTPickerFecInicio 
         Height          =   300
         Left            =   675
         TabIndex        =   11
         Top             =   315
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   529
         _Version        =   393216
         Format          =   20971521
         CurrentDate     =   37474
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyuMoneda 
         Height          =   360
         Left            =   2880
         TabIndex        =   12
         Top             =   855
         Width           =   4725
         _ExtentX        =   8334
         _ExtentY        =   635
         XcodMaxLongitud =   0
         xcodwith        =   800
         NomTabla        =   "ct_cuenta"
         ListaCampos     =   "cuentacodigo(1),cuentadescripcion(1)"
         XcodCampo       =   "cuentacodigo"
         XListCampo      =   "cuentadescripcion"
         ListaCamposDescrip=   "Código,Descripción"
         ListaCamposText =   "cuentacodigo,cuentadescripcion"
         Requerido       =   0   'False
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyuCaja 
         Height          =   360
         Left            =   2880
         TabIndex        =   1
         Top             =   270
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   635
         XcodMaxLongitud =   0
         xcodwith        =   800
         NomTabla        =   "ct_cuenta"
         ListaCampos     =   "cuentacodigo(1),cuentadescripcion(1)"
         XcodCampo       =   "cuentacodigo"
         XListCampo      =   "cuentadescripcion"
         ListaCamposDescrip=   "Código,Descripción"
         ListaCamposText =   "cuentacodigo,cuentadescripcion"
         Requerido       =   0   'False
      End
      Begin VB.Label lmon 
         Caption         =   "Moneda"
         Height          =   225
         Left            =   1995
         TabIndex        =   18
         Top             =   870
         Width           =   885
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha Final"
         Height          =   375
         Left            =   90
         TabIndex        =   17
         Top             =   765
         Width           =   750
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Inicial"
         Enabled         =   0   'False
         Height          =   480
         Left            =   75
         TabIndex        =   16
         Top             =   270
         Width           =   840
      End
      Begin VB.Label lcaja 
         Caption         =   "Caja"
         Height          =   255
         Left            =   2040
         TabIndex        =   15
         Top             =   345
         Width           =   885
      End
      Begin VB.Label lban 
         Caption         =   "Banco"
         Height          =   255
         Left            =   2040
         TabIndex        =   14
         Top             =   345
         Width           =   885
      End
      Begin VB.Label lcta 
         Caption         =   "Cuenta"
         Height          =   285
         Left            =   1950
         TabIndex        =   13
         Top             =   870
         Width           =   885
      End
   End
   Begin VB.Frame FrameCajaBancos 
      Height          =   795
      Left            =   225
      TabIndex        =   4
      Top             =   90
      Width           =   8835
      Begin VB.OptionButton Opt 
         Caption         =   "BANCO"
         Height          =   210
         Index           =   1
         Left            =   5265
         TabIndex        =   6
         Top             =   360
         Width           =   915
      End
      Begin VB.OptionButton Opt 
         Caption         =   "CAJA"
         Height          =   300
         Index           =   0
         Left            =   2880
         TabIndex        =   5
         Top             =   315
         Width           =   825
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuempresa 
         Height          =   315
         Left            =   2970
         TabIndex        =   22
         Top             =   315
         Visible         =   0   'False
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   556
         XcodMaxLongitud =   3
         xcodwith        =   300
         NomTabla        =   "co_multiempresas"
         TituloAyuda     =   "Busqueda de Empresas"
         ListaCampos     =   "empresacodigo(1),empresadescripcion(1)"
         XcodCampo       =   "empresacodigo"
         XListCampo      =   "empresadescripcion"
         ListaCamposDescrip=   "Codigo,Descripcion"
         ListaCamposText =   "empresacodigo,empresadescripcion"
         Requerido       =   0   'False
      End
      Begin VB.Label Lblempresa 
         AutoSize        =   -1  'True
         Caption         =   "Empresa :"
         Height          =   195
         Left            =   2115
         TabIndex        =   23
         Top             =   360
         Visible         =   0   'False
         Width           =   705
      End
   End
   Begin VB.Frame Frame6 
      Height          =   1350
      Left            =   9240
      TabIndex        =   0
      Top             =   960
      Width           =   1335
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   480
         Left            =   135
         TabIndex        =   3
         Top             =   765
         Width           =   1035
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   405
         Left            =   90
         TabIndex        =   2
         Top             =   240
         Width           =   1080
      End
   End
End
Attribute VB_Name = "FrmCierreDiario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim valorop As String
Dim valoroptext As String
Dim rsql As New ADODB.Recordset

Private Sub Check1_Click()
    If Check1.Value = 1 Then
        Call mostrardatos("X")
    Else
        TDBGrid1.DataSource = Nothing
    End If
 End Sub

Private Sub cmdaceptar_Click()
    If Ctr_AyuCaja.xclave = Empty Then
        MsgBox "Debe seleccionar caja a cerrar", vbExclamation, "Cierre Diario de Caja"
        Ctr_AyuCaja.SetFocus
        Exit Sub
    End If
      Call Grabar
      MsgBox "Saldos generados correctamente", vbInformation, "Cierre Diario de Caja"
End Sub

Private Sub CmdCancelar_Click()
Unload Me
End Sub

Private Sub cmdimprimir_Click()
Dim Param(4) As Variant
Dim formulas(4) As Variant


Param(0) = VGParamSistem.BDEmpresa
Param(1) = DTPickerFecInicio.Value
Param(2) = DTPickerFecFinal.Value
Param(3) = "##tmpCierremes" & VGComputer
formulas(0) = "@empresa='" & VGParametros.empresacodigo & "'"
formulas(1) = "@ruc='" & VGParametros.RucEmpresa & "'"
formulas(2) = "FechaIni='" & DTPickerFecInicio & "'"
'formulas(3) = "@FecFin='" & DTPickerFecFinal & "'"

Call ImpresionRptProc("te_imprimircierrediario.rpt", formulas, Param, , "Cierre Mensual")
End Sub

Private Sub Ctr_AyuCaja_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
Call mostrardatos("X")
End Sub

Private Sub Ctr_AyudaBanco_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
  Ctr_AyudaBancoCuenta.Filtro = "cbanco_codigo='" & ColecCampos("bancocodigo").Value & "'"
End Sub

Private Sub Ctr_Ayuempresa_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
Call mostrardatos("X")
End Sub

Private Sub Ctr_AyuMoneda_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
Call mostrardatos("X")
End Sub

Private Sub Form_Load()
cmdAceptar.Enabled = False
  Dim cFecha As Date
  Opt(0).Value = True
  Check1.Value = 1
  DTPickerFecInicio.Value = VGParamSistem.fechatrabajo
  DTPickerFecFinal.Value = VGParamSistem.fechatrabajo
  Call Ctr_AyuCaja.Conexion(VGCNx)
  Call Ctr_AyuMoneda.Conexion(VGCNx)
  Call Ctr_AyudaBanco.Conexion(VGCNx)
  Call Ctr_AyudaBancoCuenta.Conexion(VGCNx)
  Call Ctr_Ayuempresa.Conexion(VGCNx)
  If VGParametros.sistemamultiempresas Then
     Ctr_Ayuempresa.Enabled = True
   Else
     Ctr_Ayuempresa.xclave = VGParametros.empresacodigo: Ctr_Ayuempresa.Ejecutar
     Ctr_Ayuempresa.Enabled = False
  End If
  cmdAceptar.Enabled = False
  Check1_Click
  End Sub
Private Sub mostrardatos(index As String)
Set rsql = Nothing
Dim rsaux As New ADODB.Recordset
Dim fecha1 As Date
Dim yyyyi As String, mmi As String, ddi As String
Dim yyyyf As String, mmf As String, ddf As String
Dim n As Integer
fecha1 = DTPickerFecFinal + 1
yyyyf = Format(Year(fecha1), "0000")
mmf = Format(Month(fecha1), "00")
ddf = Format(Day(fecha1), "00")
yyyyi = Format(Year(DTPickerFecInicio.Value), "00")
mmi = Format(Month(DTPickerFecInicio.Value), "00")
ddi = Format(Day(DTPickerFecInicio.Value), "00")

Set VGCommandoSP = New ADODB.Command
VGCommandoSP.ActiveConnection = VGgeneral
VGCommandoSP.CommandType = adCmdStoredProc
VGCommandoSP.CommandText = "te_cierrecajabancosdiario_pro"
VGCommandoSP.Parameters.Refresh
With VGCommandoSP
    .Parameters("@base") = VGParamSistem.BDEmpresa
    If Ctr_Ayuempresa.xclave = "" Then
       .Parameters("@empresa") = "%%"
    Else
       .Parameters("@empresa") = Ctr_Ayuempresa.xclave
    End If
    'if Ctr_AyuCaja.xclave
    .Parameters("@cajabanco") = "%%"
    .Parameters("@ctamoneda") = "%%"
        If Opt(0).Value Then
       .Parameters("@tipo") = "C"
        If Check1.Value = 1 Then
           If Ctr_AyuCaja.xclave <> Empty Then .Parameters("@cajabanco") = Ctr_AyuCaja.xclave
           If Ctr_AyuMoneda.xclave <> Empty Then .Parameters("@ctamoneda") = Ctr_AyuMoneda.xclave
        End If
     Else
       .Parameters("@tipo") = "B"
        If Check1.Value = 0 Then
           .Parameters("@cajabanco") = Ctr_AyudaBanco.xclave
           .Parameters("@ctamoneda") = Ctr_AyudaBancoCuenta.xclave
        End If
   End If
   .Parameters("@diaactual") = yyyyi + mmi + ddi
   .Parameters("@dianuevo") = yyyyf + mmf + ddf
   .Parameters("@computer") = "##tmpCierremes" & VGComputer
   .Parameters("@fechaini") = DTPickerFecInicio.Value
   .Parameters("@fechafin") = DTPickerFecFinal.Value
   .Parameters("@cierre") = index
   .Execute
End With
 Set rsql = VGCNx.Execute("select * from ##tmpCierremes" & VGComputer)
 TDBGrid1.DataSource = rsql
 n = 0
 Do While Not rsql.EOF
    If (rsql!Saldoinicial + rsql!totalingresos - rsql!totalegresos) < 0 Then n = n + 1
    rsql.MoveNext
 Loop
 If n > 0 Then
    MsgBox (" existe  " & Str(n) & "  Regstros con saldos Negativos ")
    cmdAceptar.Enabled = False
  Else
     If Ctr_AyuCaja.xclave <> Empty Then cmdAceptar.Enabled = True
 End If
End Sub

Sub Grabar()
If MsgBox("Esta seguro de procesar el cierre del Mes ", vbYesNo, "AVISO") = vbYes Then
   If Check1.Value = 1 Then
      Call mostrardatos(1)
    Else
      Call mostrardatos(0)
   End If
   Call mostrardatos(3)
End If
End Sub

Sub ConfiguraCajaBanco(valor As Boolean)
  Ctr_AyuCaja.Enabled = valor
  Ctr_AyuMoneda.Enabled = valor
  Ctr_AyuCaja.Visible = valor
  Ctr_AyuMoneda.Visible = valor
  
  lcaja.Visible = valor
  lmon.Visible = valor
  
  lban.Visible = Not valor
  lcta.Visible = Not valor
  Ctr_AyudaBanco.Enabled = Not valor
  Ctr_AyudaBanco.Visible = Not valor
  Ctr_AyudaBancoCuenta.Enabled = Not valor
  Ctr_AyudaBancoCuenta.Visible = Not valor
  
  If valor = True Then
     Ctr_AyuCaja.ListaCampos = "cajacodigo(1),cajadescripcion(1)"
     Ctr_AyuCaja.ListaCamposDescrip = "Código,Descripción"
     Ctr_AyuCaja.ListaCamposText = "cajacodigo,cajadescripcion"
     Ctr_AyuCaja.NomTabla = "te_codigocaja"
     Ctr_AyuCaja.XcodCampo = "cajacodigo"
     Ctr_AyuCaja.XListCampo = "cajadescripcion"
     Ctr_AyuCaja.Conexion VGCNx
  Else
     Ctr_AyudaBanco.ListaCampos = "bancocodigo(1),bancodescripcion(1)"
     Ctr_AyudaBanco.ListaCamposDescrip = "Código,Descripción"
     Ctr_AyudaBanco.ListaCamposText = "bancocodigo,bancodescripcion"
     Ctr_AyudaBanco.NomTabla = "gr_banco"
     Ctr_AyudaBanco.XcodCampo = "bancocodigo"
     Ctr_AyudaBanco.XListCampo = "bancodescripcion"
     Ctr_AyudaBanco.Conexion VGCNx
  End If
  
  If valor = True Then
      Ctr_AyuMoneda.ListaCampos = "monedacodigo(1),monedadescripcion(1)"
      Ctr_AyuMoneda.ListaCamposDescrip = "Código,Descripción"
      Ctr_AyuMoneda.ListaCamposText = "monedacodigo,monedadescripcion"
      Ctr_AyuMoneda.NomTabla = "gr_moneda"
      Ctr_AyuMoneda.XcodCampo = "monedacodigo"
      Ctr_AyuMoneda.XListCampo = "monedadescripcion"
      Ctr_AyuMoneda.Conexion VGCNx
  Else
      Ctr_AyudaBancoCuenta.ListaCampos = "cbanco_codigo(1),cbanco_numero(1),monedasimbolo(1),cbanco_referenciacta(1),cbanco_nrocheque(1),monedacodigo(1)"
      Ctr_AyudaBancoCuenta.ListaCamposDescrip = "Código,Descripción,Mon,Ref,NCheque,MonCod"
      Ctr_AyudaBancoCuenta.ListaCamposText = "cbanco_codigo,cbanco_numero,monedasimbolo,cbanco_referenciacta,cbanco_nrocheque,monedacodigo"
      Ctr_AyudaBancoCuenta.NomTabla = "v_bancomoneda"
      Ctr_AyudaBancoCuenta.XcodCampo = "cbanco_codigo"
      Ctr_AyudaBancoCuenta.XListCampo = "cbanco_numero"
      Ctr_AyudaBancoCuenta.Conexion VGCNx
  End If
  
End Sub
Sub ConfiguraBanco(valor As Boolean)
  Ctr_AyudaBanco.Enabled = valor
  Ctr_AyudaBancoCuenta.Enabled = valor
End Sub

Private Sub Opt_Click(index As Integer)
  Select Case index
    Case 0:
       Call ConfiguraCajaBanco(True)
    Case 1:
       Call ConfiguraCajaBanco(False)
  End Select
Check1.Value = 0
Call mostrardatos("X")
End Sub

Private Sub Opt2_Click(index As Integer)
    Select Case index
        Case 0:
            valorop = "1"
        Case 1:
            valorop = "0"
        Case 2:
            valorop = "%%"
    End Select
End Sub


