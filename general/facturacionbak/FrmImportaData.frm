VERSION 5.00
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FrmImportaData 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Exportacion de Datos"
   ClientHeight    =   7710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13065
   Icon            =   "FrmImportaData.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7710
   ScaleWidth      =   13065
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Criterio de Migración "
      Height          =   1230
      Left            =   105
      TabIndex        =   10
      Top             =   60
      Width           =   12840
      Begin ctrlayuda_f.Ctr_Ayuda CtrAyu_Pventa 
         Height          =   390
         Left            =   135
         TabIndex        =   21
         Top             =   750
         Width           =   2550
         _ExtentX        =   4498
         _ExtentY        =   688
         XcodMaxLongitud =   3
         xcodwith        =   100
         NomTabla        =   "vt_puntoventa"
         TituloAyuda     =   "Puntos de Ventas"
         ListaCampos     =   "puntovtacodigo(1),puntovtadescripcion(1)"
         XcodCampo       =   "puntovtacodigo"
         XListCampo      =   "puntovtadescripcion"
         ListaCamposDescrip=   "Codigo,Descripcion"
         ListaCamposText =   "puntovtacodigo,puntovtadescripcion"
         Requerido       =   0   'False
      End
      Begin VB.CommandButton CmdProceso 
         Caption         =   "Procesar Info"
         Enabled         =   0   'False
         Height          =   315
         Left            =   10995
         TabIndex        =   19
         Top             =   555
         Width           =   1785
      End
      Begin VB.CommandButton Command1 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   315
         Left            =   10620
         TabIndex        =   18
         Top             =   555
         Width           =   330
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   11805
         Top             =   1050
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.TextBox txDir 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   7725
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   570
         Width           =   2865
      End
      Begin VB.CommandButton CmdConsultar 
         Caption         =   "&Consultar"
         Height          =   315
         Left            =   5115
         TabIndex        =   15
         Top             =   555
         Width           =   1230
      End
      Begin MSComCtl2.DTPicker DTPFechaFin 
         Height          =   315
         Left            =   3720
         TabIndex        =   14
         Top             =   765
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   556
         _Version        =   393216
         Format          =   62521345
         CurrentDate     =   37655
      End
      Begin MSComCtl2.DTPicker DTPFechaIni 
         Height          =   315
         Left            =   3735
         TabIndex        =   12
         Top             =   360
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   556
         _Version        =   393216
         Format          =   62521345
         CurrentDate     =   37655
      End
      Begin VB.Label Label8 
         Caption         =   "Punto de Venta :"
         Height          =   255
         Left            =   135
         TabIndex        =   20
         Top             =   450
         Width           =   1275
      End
      Begin VB.Label Label5 
         Caption         =   "Ruta de Archivo a Generar en el Servidor : "
         Height          =   615
         Left            =   6405
         TabIndex        =   16
         Top             =   540
         Width           =   1395
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha Fin  :"
         Height          =   300
         Left            =   2760
         TabIndex        =   13
         Top             =   810
         Width           =   1035
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha inicio :"
         Height          =   300
         Left            =   2745
         TabIndex        =   11
         Top             =   390
         Width           =   1035
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00808080&
      Height          =   2550
      Left            =   105
      ScaleHeight     =   2490
      ScaleWidth      =   12795
      TabIndex        =   4
      Top             =   1710
      Width           =   12855
      Begin TrueOleDBGrid70.TDBGrid TDBG_Cabped 
         Height          =   1845
         Left            =   105
         TabIndex        =   5
         Top             =   105
         Width           =   12585
         _ExtentX        =   22199
         _ExtentY        =   3254
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Nº de Pedido"
         Columns(0).DataField=   "pedidonumero"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Fecha Pedido"
         Columns(1).DataField=   "pedidofecha"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Fecha Docu."
         Columns(2).DataField=   "pedidofechafact"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "T/D"
         Columns(3).DataField=   "pedidotipofac"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Nro. Doc"
         Columns(4).DataField=   "pedidonrofact"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "Cod. Cliente."
         Columns(5).DataField=   "clientecodigo"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "Razon Social"
         Columns(6).DataField=   "clienterazonsocial"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "Total Bruto"
         Columns(7).DataField=   "pedidototbruto"
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(8)._VlistStyle=   0
         Columns(8)._MaxComboItems=   5
         Columns(8).Caption=   "I.G.V."
         Columns(8).DataField=   "pedidototimpuesto"
         Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(9)._VlistStyle=   0
         Columns(9)._MaxComboItems=   5
         Columns(9).Caption=   "Precio de Venta"
         Columns(9).DataField=   "pedidototneto"
         Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   10
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   12632256
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=10"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=1984"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1905"
         Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=1"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=1958"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=1879"
         Splits(0)._ColumnProps(9)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(10)=   "Column(2).Width=1931"
         Splits(0)._ColumnProps(11)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(12)=   "Column(2)._WidthInPix=1852"
         Splits(0)._ColumnProps(13)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(14)=   "Column(3).Width=714"
         Splits(0)._ColumnProps(15)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(16)=   "Column(3)._WidthInPix=635"
         Splits(0)._ColumnProps(17)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(18)=   "Column(4).Width=1852"
         Splits(0)._ColumnProps(19)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(20)=   "Column(4)._WidthInPix=1773"
         Splits(0)._ColumnProps(21)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(22)=   "Column(5).Width=2037"
         Splits(0)._ColumnProps(23)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(24)=   "Column(5)._WidthInPix=1958"
         Splits(0)._ColumnProps(25)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(26)=   "Column(6).Width=3201"
         Splits(0)._ColumnProps(27)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(28)=   "Column(6)._WidthInPix=3122"
         Splits(0)._ColumnProps(29)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(30)=   "Column(7).Width=2302"
         Splits(0)._ColumnProps(31)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(32)=   "Column(7)._WidthInPix=2223"
         Splits(0)._ColumnProps(33)=   "Column(7).Order=8"
         Splits(0)._ColumnProps(34)=   "Column(8).Width=2328"
         Splits(0)._ColumnProps(35)=   "Column(8).DividerColor=0"
         Splits(0)._ColumnProps(36)=   "Column(8)._WidthInPix=2249"
         Splits(0)._ColumnProps(37)=   "Column(8).Order=9"
         Splits(0)._ColumnProps(38)=   "Column(9).Width=2725"
         Splits(0)._ColumnProps(39)=   "Column(9).DividerColor=0"
         Splits(0)._ColumnProps(40)=   "Column(9)._WidthInPix=2646"
         Splits(0)._ColumnProps(41)=   "Column(9).Order=10"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   3
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         Appearance      =   2
         DefColWidth     =   0
         HeadLines       =   1
         FootLines       =   1
         MultipleLines   =   0
         CellTipsWidth   =   0
         MultiSelect     =   2
         DeadAreaBackColor=   12632256
         RowDividerColor =   12632256
         RowSubDividerColor=   12632256
         DirectionAfterEnter=   1
         MaxRows         =   250000
         ViewColumnCaptionWidth=   0
         ViewColumnWidth =   0
         _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=13,.bold=0,.fontsize=825,.italic=0"
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
         _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36,.bgcolor=&H344A87&"
         _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
         _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1,.bgcolor=&HFFFFFF&"
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
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=78,.parent=13,.alignment=2,.bgcolor=&HFFFFB7&"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=75,.parent=14,.bgcolor=&HFFFF00&"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=76,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=77,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=28,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=32,.parent=13"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
         _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=46,.parent=13"
         _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=43,.parent=14"
         _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=44,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=45,.parent=17"
         _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=50,.parent=13"
         _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=47,.parent=14"
         _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=48,.parent=15"
         _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=49,.parent=17"
         _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=54,.parent=13"
         _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=51,.parent=14"
         _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=52,.parent=15"
         _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=53,.parent=17"
         _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=58,.parent=13"
         _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=55,.parent=14"
         _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=56,.parent=15"
         _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=57,.parent=17"
         _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=62,.parent=13"
         _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=59,.parent=14"
         _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=60,.parent=15"
         _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=61,.parent=17"
         _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=66,.parent=13"
         _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=63,.parent=14"
         _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=64,.parent=15"
         _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=65,.parent=17"
         _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=70,.parent=13"
         _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=67,.parent=14"
         _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=68,.parent=15"
         _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=69,.parent=17"
         _StyleDefs(76)  =   "Named:id=33:Normal"
         _StyleDefs(77)  =   ":id=33,.parent=0"
         _StyleDefs(78)  =   "Named:id=34:Heading"
         _StyleDefs(79)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(80)  =   ":id=34,.wraptext=-1"
         _StyleDefs(81)  =   "Named:id=35:Footing"
         _StyleDefs(82)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(83)  =   "Named:id=36:Selected"
         _StyleDefs(84)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(85)  =   "Named:id=37:Caption"
         _StyleDefs(86)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(87)  =   "Named:id=38:HighlightRow"
         _StyleDefs(88)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(89)  =   "Named:id=39:EvenRow"
         _StyleDefs(90)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(91)  =   "Named:id=40:OddRow"
         _StyleDefs(92)  =   ":id=40,.parent=33"
         _StyleDefs(93)  =   "Named:id=41:RecordSelector"
         _StyleDefs(94)  =   ":id=41,.parent=34"
         _StyleDefs(95)  =   "Named:id=42:FilterBar"
         _StyleDefs(96)  =   ":id=42,.parent=33"
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Nro. Reg :"
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   10485
         TabIndex        =   7
         Top             =   2145
         Width           =   930
      End
      Begin VB.Label lbreg1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "0 "
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   11535
         TabIndex        =   6
         Top             =   2145
         Width           =   1140
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00808080&
      Height          =   2970
      Left            =   105
      ScaleHeight     =   2910
      ScaleWidth      =   12780
      TabIndex        =   0
      Top             =   4680
      Width           =   12840
      Begin TrueOleDBGrid70.TDBGrid TDBG_Detped 
         Height          =   2370
         Left            =   120
         TabIndex        =   1
         Top             =   165
         Width           =   12585
         _ExtentX        =   22199
         _ExtentY        =   4180
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Item"
         Columns(0).DataField=   "detpeditem"
         Columns(0).NumberFormat=   "###,###,###.00"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Unidad"
         Columns(1).DataField=   "unidadcodigo"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Cantidad "
         Columns(2).DataField=   "detpedcantpedida"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Codigo Producto "
         Columns(3).DataField=   "productocodigo"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Precio Articulo "
         Columns(4).DataField=   "detpedpreciopact"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "Importe Bruto"
         Columns(5).DataField=   "detpedimpbruto"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "I.G.V."
         Columns(6).DataField=   "detpedmontoimpto"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "Precio de Venta"
         Columns(7).DataField=   "detpedmontoprecvta"
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   8
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   12632256
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=8"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=741"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=661"
         Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=2"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=1667"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=1588"
         Splits(0)._ColumnProps(9)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(10)=   "Column(2).Width=2064"
         Splits(0)._ColumnProps(11)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(12)=   "Column(2)._WidthInPix=1984"
         Splits(0)._ColumnProps(13)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(14)=   "Column(3).Width=2752"
         Splits(0)._ColumnProps(15)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(16)=   "Column(3)._WidthInPix=2672"
         Splits(0)._ColumnProps(17)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(18)=   "Column(4).Width=2725"
         Splits(0)._ColumnProps(19)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(20)=   "Column(4)._WidthInPix=2646"
         Splits(0)._ColumnProps(21)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(22)=   "Column(5).Width=2725"
         Splits(0)._ColumnProps(23)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(24)=   "Column(5)._WidthInPix=2646"
         Splits(0)._ColumnProps(25)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(26)=   "Column(6).Width=2725"
         Splits(0)._ColumnProps(27)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(28)=   "Column(6)._WidthInPix=2646"
         Splits(0)._ColumnProps(29)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(30)=   "Column(7).Width=2725"
         Splits(0)._ColumnProps(31)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(32)=   "Column(7)._WidthInPix=2646"
         Splits(0)._ColumnProps(33)=   "Column(7).Order=8"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   3
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         Appearance      =   2
         DefColWidth     =   0
         HeadLines       =   1
         FootLines       =   1
         MultipleLines   =   0
         CellTipsWidth   =   0
         MultiSelect     =   2
         DeadAreaBackColor=   12632256
         RowDividerColor =   12632256
         RowSubDividerColor=   12632256
         DirectionAfterEnter=   1
         MaxRows         =   250000
         ViewColumnCaptionWidth=   0
         ViewColumnWidth =   0
         _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=13,.bold=0,.fontsize=825,.italic=0"
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
         _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36,.bgcolor=&H344A87&"
         _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
         _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1,.bgcolor=&HFFFFFF&"
         _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
         _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2,.bgcolor=&HC0C0C0&"
         _StyleDefs(27)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(30)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=74,.parent=13,.alignment=1"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=71,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=72,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=73,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=28,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=32,.parent=13"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
         _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=46,.parent=13"
         _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=43,.parent=14"
         _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=44,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=45,.parent=17"
         _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=50,.parent=13"
         _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=47,.parent=14"
         _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=48,.parent=15"
         _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=49,.parent=17"
         _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=54,.parent=13"
         _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=51,.parent=14"
         _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=52,.parent=15"
         _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=53,.parent=17"
         _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=58,.parent=13"
         _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=55,.parent=14"
         _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=56,.parent=15"
         _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=57,.parent=17"
         _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=62,.parent=13"
         _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=59,.parent=14"
         _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=60,.parent=15"
         _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=61,.parent=17"
         _StyleDefs(68)  =   "Named:id=33:Normal"
         _StyleDefs(69)  =   ":id=33,.parent=0"
         _StyleDefs(70)  =   "Named:id=34:Heading"
         _StyleDefs(71)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(72)  =   ":id=34,.wraptext=-1"
         _StyleDefs(73)  =   "Named:id=35:Footing"
         _StyleDefs(74)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(75)  =   "Named:id=36:Selected"
         _StyleDefs(76)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(77)  =   "Named:id=37:Caption"
         _StyleDefs(78)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(79)  =   "Named:id=38:HighlightRow"
         _StyleDefs(80)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(81)  =   "Named:id=39:EvenRow"
         _StyleDefs(82)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(83)  =   "Named:id=40:OddRow"
         _StyleDefs(84)  =   ":id=40,.parent=33"
         _StyleDefs(85)  =   "Named:id=41:RecordSelector"
         _StyleDefs(86)  =   ":id=41,.parent=34"
         _StyleDefs(87)  =   "Named:id=42:FilterBar"
         _StyleDefs(88)  =   ":id=42,.parent=33"
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Nro. Reg :"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   10515
         TabIndex        =   3
         Top             =   2610
         Width           =   930
      End
      Begin VB.Label lbreg2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "0 "
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   11565
         TabIndex        =   2
         Top             =   2595
         Width           =   1140
      End
   End
   Begin VB.Label Label1 
      BackColor       =   &H00DDF7F9&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Provisiones no Contabilizadas en el Mes"
      Height          =   315
      Left            =   90
      TabIndex        =   9
      Top             =   1335
      Width           =   12855
   End
   Begin VB.Label Label2 
      BackColor       =   &H00DDF7F9&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Provisiones Contabilizadas en el Mes"
      Height          =   315
      Left            =   105
      TabIndex        =   8
      Top             =   4305
      Width           =   12825
   End
End
Attribute VB_Name = "FrmImportaData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsdata As ADODB.Recordset
Dim rsdetalle As ADODB.Recordset
Dim SqlCad As String
Dim Fechaini As Long, Fechafin As Long
Dim Comando As ADODB.Command
Dim Directory As String
Private Sub CmdConsultar_Click()
    If Trim(CtrAyu_Pventa.xclave) = "" Then
        MsgBox "Tiene que seleccionar un punto de venta", vbExclamation, "Mensaje del Sistema"
        CtrAyu_Pventa.SetFocus
        Exit Sub
    End If

    Set rsdata = New ADODB.Recordset
    Set rsdetalle = New ADODB.Recordset
    Fechaini = FechS(DTPFechaIni.Value, Sqlf)
    Fechafin = FechS(DTPFechaFin.Value, Sqlf)
    SqlCad = "Select pedidonumero,pedidofecha,pedidofechafact,clienteruc,clientecodigo,clienterazonsocial,pedidotipofac,pedidonrofact, " & _
             "pedidototbruto,pedidototalotros,pedidototalflete,pedidodsctoglobal,pedidototimpuesto, " & _
             "pedidototneto , pedidofechasunat " & _
             "From vt_pedido A " & _
             "Where floor(cast(A.pedidofecha as real)) between " & Fechaini & " And " & Fechafin & _
             " and puntovtacodigo='" & CtrAyu_Pventa.xclave & "'"
    rsdata.Open SqlCad, VGcnx, adOpenKeyset, adLockReadOnly
    Set TDBG_Cabped.DataSource = rsdata
    lbreg1.Caption = Format(rsdata.RecordCount, "0 ")
    If rsdata.RecordCount = 0 Then
        Command1.Enabled = False
        CmdProceso.Enabled = False
      Else
        Command1.Enabled = True
        CmdProceso.Enabled = True
    End If
End Sub

Private Sub CmdProceso_Click()
On Error GoTo Proceso
If Trim(txDir.Text) = "" Then
    MsgBox "Tiene que seleccionar un directorio donde se almacenara el archivo", vbExclamation
    Command1.SetFocus
    Exit Sub
End If
Set Comando = New ADODB.Command
    Screen.MousePointer = 11
    
    With Comando
        .CommandType = adCmdStoredProc
        .CommandText = "vt_exportadata_pro"
        .ActiveConnection = VGgeneral
        .Parameters.Refresh
        .Parameters("@Basetransfer") = "Transfer"
        .Parameters("@BaseVenta") = VGcnx.DefaultDatabase
        .Parameters("@FechaIni") = Format(Fechaini, "0")
        .Parameters("@FechaFin") = Format(Fechafin, "0")
        .Parameters("@PtoVta") = CtrAyu_Pventa.xclave
        .Execute
    End With
    MsgBox "Se procede a generar el archivo"
    Call Backup
    Screen.MousePointer = 1
    MsgBox "La Operacion se Realizo Satisfactoriamente", vbInformation, "Sistema de Ventas"
    Unload Me
    Exit Sub
Proceso:
    Screen.MousePointer = 1
    MsgBox Err.Description
    
End Sub
Private Sub Backup()
Dim NomArchivo As String, compara As String
On Error GoTo ErrorBackup

    NomArchivo = Directory & "\EXPO_P" & CtrAyu_Pventa.xclave & "_" & Format(Day(DTPFechaIni), "00") & _
                 "_" & Format(Day(DTPFechaFin), "00") & "_" & Format(Month(DTPFechaFin), "00") & _
                 Format(Year(DTPFechaFin), "0000") & ".EX"
    
    compara = "EXPO_P" & CtrAyu_Pventa.xclave & "_" & Format(Day(DTPFechaIni), "00") & _
              "_" & Format(Day(DTPFechaFin), "00") & "_" & Format(Month(DTPFechaFin), "00") & _
                 Format(Year(DTPFechaFin), "0000") & ".EX"
    
    If Trim(Dir$(NomArchivo, vbArchive)) = compara Then
        Call VBA.Kill(compara)
    End If
    'VGgeneral.BeginTrans
    'MsgBox "Se procede a comprimir"
    DoEvents
    SqlCad = "DBCC SHRINKDATABASE (TRANSFER)  " & Chr(13) & _
           "BACKUP DATABASE TRANSFER " & _
           "TO DISK ='" & NomArchivo & "' " & _
           "With Format "
    VGgeneral.Execute SqlCad
    'VGgeneral.CommitTrans
    Exit Sub
ErrorBackup:
    'VGgeneral.RollbackTrans
    MsgBox "Error al Hacer el Backup " & Err.Description
End Sub

Private Sub Command1_Click()
Dim pos As Integer
    CommonDialog1.ShowOpen
    Directory = CommonDialog1.FileName
    Directory = StrReverse(Directory)
    pos = InStr(Directory, "\")
    Directory = Right(Directory, Len(Directory) - pos)
    Directory = StrReverse(Directory)
    txDir.Text = Directory
End Sub

Private Sub DTPFechaFin_Change()
    DTPFechaFin.Month = DTPFechaIni.Month
    DTPFechaFin.Year = DTPFechaIni.Year
End Sub

Private Sub DTPFechaIni_Change()
    DTPFechaFin.Month = DTPFechaIni.Month
    DTPFechaFin.Year = DTPFechaIni.Year
End Sub

Private Sub Form_Load()
    Me.Height = 8085
    Me.Width = 13155
    Call CtrAyu_Pventa.conexion(VGcnx)
End Sub

Private Sub TDBG_Cabped_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error Resume Next
    SqlCad = "Select pedidonumero, detpeditem, detpedcantpedida,detpedpreciopact, " & _
             "productocodigo, unidadcodigo, detpedimpbruto, detpedmontoprecvta, " & _
             "detpedestado , detpedmontoimpto " & _
             "From vt_detallepedido Where pedidonumero='" & TDBG_Cabped.Columns(0).Text & "'"
    Set rsdetalle = New ADODB.Recordset
    rsdetalle.Open SqlCad, VGcnx, adOpenKeyset, adLockReadOnly
    Set TDBG_Detped.DataSource = rsdetalle
    lbreg2.Caption = Format(rsdetalle.RecordCount, "0 ")
End Sub
