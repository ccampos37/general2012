VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Begin VB.Form FrmDistribucion_1 
   Caption         =   "Desdobles"
   ClientHeight    =   7845
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11535
   LinkTopic       =   "Form2"
   ScaleHeight     =   7845
   ScaleWidth      =   11535
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame8 
      Caption         =   "Producto Origen"
      Height          =   1155
      Left            =   0
      TabIndex        =   36
      Top             =   1200
      Width           =   11145
      Begin VB.TextBox Textlote 
         Height          =   285
         Left            =   615
         MaxLength       =   20
         TabIndex        =   7
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox txtdensidad1 
         BackColor       =   &H80000004&
         Height          =   330
         Left            =   8040
         TabIndex        =   48
         Top             =   240
         Width           =   900
      End
      Begin VB.TextBox txtcanti 
         Height          =   330
         Index           =   5
         Left            =   3030
         TabIndex        =   8
         Top             =   690
         Width           =   1005
      End
      Begin VB.CommandButton cAyuda 
         Caption         =   "..."
         Height          =   285
         Index           =   0
         Left            =   1635
         TabIndex        =   10
         Top             =   240
         Width           =   285
      End
      Begin VB.TextBox txtcanti 
         Enabled         =   0   'False
         Height          =   330
         Index           =   3
         Left            =   9840
         TabIndex        =   11
         Top             =   180
         Width           =   900
      End
      Begin MSMask.MaskEdBox MBox2 
         Height          =   330
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   582
         _Version        =   393216
         MaxLength       =   8
         PromptChar      =   "_"
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayumaquina1 
         Height          =   375
         Left            =   5520
         TabIndex        =   9
         Top             =   720
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   661
         XcodMaxLongitud =   0
         xcodwith        =   100
         NomTabla        =   "al_tipomaquina"
         TituloAyuda     =   "Tipo de Maquina"
         ListaCampos     =   "codigomaquina(1),descripcionmaquina(1),factormaquina(1)"
         XcodCampo       =   "codigomaquina"
         XListCampo      =   "descripcionmaquina"
         ListaCamposDescrip=   "Codigo,Descripcion,factor"
         ListaCamposText =   "codigomaquina,descripcionmaquina,factormaquina"
      End
      Begin VB.Label Label4 
         Caption         =   "Lote"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   53
         Top             =   750
         Width           =   525
      End
      Begin VB.Label Label2 
         Caption         =   "Densidad"
         Height          =   285
         Index           =   2
         Left            =   7080
         TabIndex        =   49
         Top             =   240
         Width           =   840
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo de Sistema"
         Height          =   195
         Index           =   1
         Left            =   4080
         TabIndex        =   46
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Numero Sistemas"
         Height          =   510
         Index           =   5
         Left            =   2160
         TabIndex        =   39
         Top             =   660
         Width           =   645
      End
      Begin VB.Label Label8 
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   2010
         TabIndex        =   38
         Top             =   240
         Width           =   4830
      End
      Begin VB.Label Label2 
         Caption         =   "Saldo"
         Height          =   285
         Index           =   3
         Left            =   9240
         TabIndex        =   37
         Top             =   225
         Width           =   480
      End
   End
   Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
      Height          =   1620
      Left            =   0
      TabIndex        =   22
      Top             =   5040
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   2858
      _LayoutType     =   0
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   7
      Splits(0)._UserFlags=   0
      Splits(0).MarqueeStyle=   3
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).DividerColor=   12632256
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=7"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=2725"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2646"
      Splits(0)._ColumnProps(9)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(11)=   "Column(2).Width=2725"
      Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=2646"
      Splits(0)._ColumnProps(14)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(16)=   "Column(3).Width=2725"
      Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=2646"
      Splits(0)._ColumnProps(19)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(21)=   "Column(4).Width=2725"
      Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=2646"
      Splits(0)._ColumnProps(24)=   "Column(4)._EditAlways=0"
      Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(26)=   "Column(5).Width=2725"
      Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=2646"
      Splits(0)._ColumnProps(29)=   "Column(5)._EditAlways=0"
      Splits(0)._ColumnProps(30)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(31)=   "Column(6).Width=2725"
      Splits(0)._ColumnProps(32)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(33)=   "Column(6)._WidthInPix=2646"
      Splits(0)._ColumnProps(34)=   "Column(6)._EditAlways=0"
      Splits(0)._ColumnProps(35)=   "Column(6).Order=7"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=9.75,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=System"
      PrintInfos(0).PageFooterFont=   "Size=9.75,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=System"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowUpdate     =   0   'False
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      MultipleLines   =   0
      CellTipsWidth   =   0
      InsertMode      =   0   'False
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
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=18,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HFFFFFF&,.bold=0,.fontsize=825"
      _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
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
      _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1,.bgcolor=&HC0C0C0&"
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
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
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
      _StyleDefs(64)  =   "Named:id=33:Normal"
      _StyleDefs(65)  =   ":id=33,.parent=0"
      _StyleDefs(66)  =   "Named:id=34:Heading"
      _StyleDefs(67)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(68)  =   ":id=34,.wraptext=-1"
      _StyleDefs(69)  =   "Named:id=35:Footing"
      _StyleDefs(70)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(71)  =   "Named:id=36:Selected"
      _StyleDefs(72)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(73)  =   "Named:id=37:Caption"
      _StyleDefs(74)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(75)  =   "Named:id=38:HighlightRow"
      _StyleDefs(76)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(77)  =   "Named:id=39:EvenRow"
      _StyleDefs(78)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(79)  =   "Named:id=40:OddRow"
      _StyleDefs(80)  =   ":id=40,.parent=33"
      _StyleDefs(81)  =   "Named:id=41:RecordSelector"
      _StyleDefs(82)  =   ":id=41,.parent=34"
      _StyleDefs(83)  =   "Named:id=42:FilterBar"
      _StyleDefs(84)  =   ":id=42,.parent=33"
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   2880
      Top             =   6720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame2 
      Height          =   555
      Left            =   1935
      TabIndex        =   30
      Top             =   6915
      Width           =   6030
      Begin VB.Label Label3 
         Caption         =   "Nota Salida"
         Height          =   165
         Index           =   0
         Left            =   210
         TabIndex        =   34
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Nota Ingreso"
         Height          =   195
         Index           =   1
         Left            =   3390
         TabIndex        =   33
         Top             =   240
         Width           =   945
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
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
         Height          =   285
         Index           =   0
         Left            =   1320
         TabIndex        =   32
         Top             =   180
         Width           =   1245
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
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
         Height          =   285
         Index           =   1
         Left            =   4560
         TabIndex        =   31
         Top             =   180
         Width           =   1245
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1140
      Left            =   0
      TabIndex        =   27
      Top             =   0
      Width           =   11310
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   9855
         MaxLength       =   10
         TabIndex        =   5
         Top             =   750
         Visible         =   0   'False
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   9840
         TabIndex        =   4
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   98828289
         CurrentDate     =   38623
         MinDate         =   37987
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuda1 
         Height          =   375
         Left            =   1080
         TabIndex        =   0
         Top             =   240
         Width           =   3780
         _ExtentX        =   6668
         _ExtentY        =   661
         XcodMaxLongitud =   0
         xcodwith        =   100
         NomTabla        =   "tabalm"
         TituloAyuda     =   "Almacenes"
         ListaCampos     =   "TAALMA(1),TADESCRI(1)"
         XcodCampo       =   "TAALMA"
         XListCampo      =   "TADESCRI"
         ListaCamposDescrip=   "Codigo,Descripcion"
         ListaCamposText =   "TAALMA,TADESCRI"
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayusalida 
         Height          =   375
         Left            =   1080
         TabIndex        =   1
         Top             =   600
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   661
         XcodMaxLongitud =   0
         xcodwith        =   100
         NomTabla        =   "tabtransa"
         TituloAyuda     =   "Transaciones"
         ListaCampos     =   "tt_codmov(1),tt_descri(1),tt_dr(1),tt_codtrans_auto(1),tt_codtrans_merma(1)"
         XcodCampo       =   "tt_codmov"
         XListCampo      =   "tt_descri"
         ListaCamposDescrip=   "Codigo,Descripcion,doc.ref.,trans.auto,trans merma"
         ListaCamposText =   "tt_codmov,tt_descri,tt_dr,tt_codtrans_auto,tt_codtrans_merma"
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuingreso 
         Height          =   375
         Left            =   6000
         TabIndex        =   2
         Top             =   240
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   661
         Enabled         =   0   'False
         XcodMaxLongitud =   0
         xcodwith        =   100
         NomTabla        =   "tabtransa"
         TituloAyuda     =   "Transaciones"
         ListaCampos     =   "tt_codmov(1),tt_descri(1),tt_dr(1)"
         XcodCampo       =   "tt_codmov"
         XListCampo      =   "tt_descri"
         ListaCamposDescrip=   "Codigo,Descripcion,doc.ref."
         ListaCamposText =   "tt_codmov,tt_descri,tt_dr"
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayumerma 
         Height          =   375
         Left            =   6000
         TabIndex        =   3
         Top             =   720
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   661
         Enabled         =   0   'False
         XcodMaxLongitud =   0
         xcodwith        =   100
         NomTabla        =   "tabtransa"
         TituloAyuda     =   "Transaciones"
         ListaCampos     =   "tt_codmov(1),tt_descri(1),tt_dr(1)"
         XcodCampo       =   "tt_codmov"
         XListCampo      =   "tt_descri"
         ListaCamposDescrip=   "Codigo,Descripcion,doc.ref."
         ListaCamposText =   "tt_codmov,tt_descri,tt_dr"
      End
      Begin VB.Label Label1 
         Caption         =   "Trans. Sal. Merma"
         Height          =   435
         Index           =   6
         Left            =   5040
         TabIndex        =   45
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Trans. Ingreso"
         Height          =   315
         Index           =   5
         Left            =   4920
         TabIndex        =   44
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Lote"
         Height          =   255
         Index           =   2
         Left            =   9360
         TabIndex        =   43
         Top             =   780
         Width           =   525
      End
      Begin VB.Label Label1 
         Caption         =   "Trans. salida"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   42
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Alm.Origen"
         Height          =   195
         Index           =   0
         Left            =   75
         TabIndex        =   29
         Top             =   300
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha "
         Height          =   285
         Index           =   4
         Left            =   9270
         TabIndex        =   28
         Top             =   225
         Width           =   825
      End
   End
   Begin VB.Frame Frame4 
      Height          =   930
      Left            =   8340
      TabIndex        =   26
      Top             =   6780
      Width           =   2880
      Begin VB.CommandButton Cmdbotones 
         Caption         =   "&Ingreso"
         Height          =   720
         Index           =   0
         Left            =   1935
         Picture         =   "FrmDistribucion_1.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   120
         Width           =   870
      End
      Begin VB.CommandButton Cmdbotones 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         Height          =   735
         Index           =   12
         Left            =   990
         Picture         =   "FrmDistribucion_1.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   135
         Width           =   855
      End
      Begin VB.CommandButton Cmdbotones 
         Caption         =   "&Grabar"
         Height          =   735
         Index           =   11
         Left            =   90
         Picture         =   "FrmDistribucion_1.frx":0884
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   135
         Width           =   810
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Nota"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   975
      Left            =   60
      TabIndex        =   23
      Top             =   6735
      Width           =   1800
      Begin VB.Label Label6 
         Caption         =   "[DEL]"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   810
         TabIndex        =   25
         Top             =   255
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   150
         Top             =   195
         Width           =   480
      End
      Begin VB.Label Label4 
         Caption         =   "Eliminar Item"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Index           =   3
         Left            =   645
         TabIndex        =   24
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Producto Destino"
      Height          =   1275
      Left            =   120
      TabIndex        =   21
      Top             =   3720
      Width           =   11265
      Begin VB.TextBox Txtdensidad2 
         BackColor       =   &H80000004&
         Height          =   330
         Left            =   8040
         TabIndex        =   51
         Top             =   240
         Width           =   900
      End
      Begin VB.TextBox txtcanti 
         Height          =   330
         Index           =   1
         Left            =   1290
         TabIndex        =   15
         Top             =   780
         Width           =   960
      End
      Begin VB.TextBox txtcanti 
         Enabled         =   0   'False
         Height          =   330
         Index           =   2
         Left            =   9960
         TabIndex        =   14
         Top             =   180
         Width           =   900
      End
      Begin VB.CommandButton cAyuda 
         Caption         =   "..."
         Height          =   285
         Index           =   3
         Left            =   1395
         TabIndex        =   13
         Top             =   240
         Width           =   285
      End
      Begin MSMask.MaskEdBox MBox2 
         Height          =   330
         Index           =   1
         Left            =   45
         TabIndex        =   12
         Top             =   240
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   582
         _Version        =   393216
         MaxLength       =   8
         PromptChar      =   "_"
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayumaquina2 
         Height          =   375
         Left            =   5640
         TabIndex        =   16
         Top             =   840
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   661
         XcodMaxLongitud =   0
         xcodwith        =   100
         NomTabla        =   "al_tipomaquina"
         TituloAyuda     =   "Tipo de Maquina"
         ListaCampos     =   "codigomaquina(1),descripcionmaquina(1),factormaquina(1)"
         XcodCampo       =   "codigomaquina"
         XListCampo      =   "descripcionmaquina"
         ListaCamposDescrip=   "Codigo,Descripcion,factor"
         ListaCamposText =   "codigomaquina,descripcionmaquina,factormaquina"
      End
      Begin VB.Label Label2 
         Caption         =   "Densidad"
         Height          =   285
         Index           =   6
         Left            =   7080
         TabIndex        =   52
         Top             =   240
         Width           =   840
      End
      Begin VB.Label Label2 
         Caption         =   "Cantidad Sistema"
         Height          =   405
         Index           =   1
         Left            =   240
         TabIndex        =   50
         Top             =   720
         Width           =   945
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo de Sistema"
         Height          =   195
         Index           =   4
         Left            =   4200
         TabIndex        =   47
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Saldo"
         Height          =   285
         Index           =   0
         Left            =   9330
         TabIndex        =   35
         Top             =   225
         Width           =   480
      End
      Begin VB.Label Label5 
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   1770
         TabIndex        =   20
         Top             =   240
         Width           =   4230
      End
   End
   Begin TrueOleDBGrid70.TDBGrid TDBGrid0 
      Height          =   1260
      Left            =   0
      TabIndex        =   40
      Top             =   2400
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   2223
      _LayoutType     =   0
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   7
      Splits(0)._UserFlags=   0
      Splits(0).MarqueeStyle=   3
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).DividerColor=   12632256
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=7"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=2725"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2646"
      Splits(0)._ColumnProps(9)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(11)=   "Column(2).Width=2725"
      Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=2646"
      Splits(0)._ColumnProps(14)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(16)=   "Column(3).Width=2725"
      Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=2646"
      Splits(0)._ColumnProps(19)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(21)=   "Column(4).Width=2725"
      Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=2646"
      Splits(0)._ColumnProps(24)=   "Column(4)._EditAlways=0"
      Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(26)=   "Column(5).Width=2725"
      Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=2646"
      Splits(0)._ColumnProps(29)=   "Column(5)._EditAlways=0"
      Splits(0)._ColumnProps(30)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(31)=   "Column(6).Width=2725"
      Splits(0)._ColumnProps(32)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(33)=   "Column(6)._WidthInPix=2646"
      Splits(0)._ColumnProps(34)=   "Column(6)._EditAlways=0"
      Splits(0)._ColumnProps(35)=   "Column(6).Order=7"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=9.75,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=System"
      PrintInfos(0).PageFooterFont=   "Size=9.75,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=System"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowUpdate     =   0   'False
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      MultipleLines   =   0
      CellTipsWidth   =   0
      InsertMode      =   0   'False
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
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=18,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HFFFFFF&,.bold=0,.fontsize=825"
      _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
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
      _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1,.bgcolor=&HC0C0C0&"
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
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
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
      _StyleDefs(64)  =   "Named:id=33:Normal"
      _StyleDefs(65)  =   ":id=33,.parent=0"
      _StyleDefs(66)  =   "Named:id=34:Heading"
      _StyleDefs(67)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(68)  =   ":id=34,.wraptext=-1"
      _StyleDefs(69)  =   "Named:id=35:Footing"
      _StyleDefs(70)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(71)  =   "Named:id=36:Selected"
      _StyleDefs(72)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(73)  =   "Named:id=37:Caption"
      _StyleDefs(74)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(75)  =   "Named:id=38:HighlightRow"
      _StyleDefs(76)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(77)  =   "Named:id=39:EvenRow"
      _StyleDefs(78)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(79)  =   "Named:id=40:OddRow"
      _StyleDefs(80)  =   ":id=40,.parent=33"
      _StyleDefs(81)  =   "Named:id=41:RecordSelector"
      _StyleDefs(82)  =   ":id=41,.parent=34"
      _StyleDefs(83)  =   "Named:id=42:FilterBar"
      _StyleDefs(84)  =   ":id=42,.parent=33"
   End
   Begin VB.Label Label1 
      Caption         =   "Alm.Destino"
      Height          =   195
      Index           =   2
      Left            =   15000
      TabIndex        =   41
      Top             =   600
      Width           =   975
   End
End
Attribute VB_Name = "FrmDistribucion_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dllgeneral As New dllgeneral.dll_general
Dim rsdeta As New ADODB.Recordset
Dim rsdeta1 As New ADODB.Recordset
Dim rssaldo As New ADODB.Recordset
Dim FACTOR As Double
Dim flag As Integer
Dim posi As Integer
Dim ruc As String


Public Function CargaGrilla()
   Set rsdeta1 = Nothing
   Call rsdeta1.Fields.Append("Item", adInteger)
   Call rsdeta1.Fields.Append("Codigo", adChar, 20)
   Call rsdeta1.Fields.Append("Descripcion", adChar, 100)
   Call rsdeta1.Fields.Append("UM", adChar, 3)
   Call rsdeta1.Fields.Append("Cant", adDouble)
   Call rsdeta1.Fields.Append("CantRef", adDouble)
   Call rsdeta1.Fields.Append("CantSistema", adDouble)
   Call rsdeta1.Fields.Append("TipoSistema", adChar, 4)
   Call rsdeta1.Fields.Append("Lote", adChar, 20)
   
   Set rsdeta = Nothing
   Call rsdeta.Fields.Append("Item", adInteger)
   Call rsdeta.Fields.Append("Codigo", adChar, 20)
   Call rsdeta.Fields.Append("Descripcion", adChar, 100)
   Call rsdeta.Fields.Append("UM", adChar, 3)
   Call rsdeta.Fields.Append("Cant", adDouble)
   Call rsdeta.Fields.Append("CantRef", adDouble)
   Call rsdeta.Fields.Append("CantSistema", adDouble)
   Call rsdeta.Fields.Append("TipoSistema", adChar, 4)
   Call rsdeta.Fields.Append("Lote", adChar, 20)

   rsdeta1.Open
   rsdeta.Open
   
   ConfigGrid

End Function

Public Function ConfigGrid()
   
Set TDBGrid0.DataSource = Nothing
   
   Set TDBGrid0.DataSource = rsdeta1
   With TDBGrid0
      .Columns(0).Width = 400
      .Columns(0).Caption = "Item"
      .Columns(1).Width = 1200
      .Columns(1).Caption = "Codigo"
      .Columns(2).Width = 4500
      .Columns(2).Caption = "Descripcion"
      .Columns(3).Width = 400
      .Columns(3).Caption = "UM"
      .Columns(4).Width = 1000
      .Columns(4).Caption = "Cant"
      .Columns(4).NumberFormat = "##,###,##0.00"
      .Columns(5).Width = 1000
      .Columns(5).Caption = "Cant.Ref"
      .Columns(5).NumberFormat = "##,###,##0.00"
      .Columns(6).Width = 1000
      .Columns(6).Caption = "Cant.Sistema"
      .Columns(6).NumberFormat = "##,###,##0.00"
      .Columns(7).Width = 1000
      .Columns(7).Caption = "Tip.Sist."
      .Columns(8).Width = 800
      .Columns(8).Caption = "Lote"
    End With
   TDBGrid0.Refresh
   
   '--------
   
   Set TDBGrid1.DataSource = Nothing
   
   Set TDBGrid1.DataSource = rsdeta
   With TDBGrid1
     .Columns(0).Width = 400
      .Columns(0).Caption = "Item"
      .Columns(1).Width = 1200
      .Columns(1).Caption = "Codigo"
      .Columns(2).Width = 4500
      .Columns(2).Caption = "Descripcion"
      .Columns(3).Width = 400
      .Columns(3).Caption = "UM"
      .Columns(4).Width = 1000
      .Columns(4).Caption = "Cant"
      .Columns(4).NumberFormat = "##,###,##0.00"
      .Columns(5).Width = 1000
      .Columns(5).Caption = "Cant.Ref"
      .Columns(5).NumberFormat = "##,###,##0.00"
      .Columns(6).Width = 1000
      .Columns(6).Caption = "Cant.Sistema"
      .Columns(6).NumberFormat = "##,###,##0.00"
      .Columns(7).Width = 1000
      .Columns(7).Caption = "Tip.Sist."
      .Columns(8).Width = 800
      .Columns(8).Caption = "Lote"
   End With
   TDBGrid1.Refresh
End Function

Private Sub cAyuda_Click(Index As Integer)
   Dim sfiltra(1 To 2, 1 To 2) As String
    If Index = 0 Then
        If Len(Label8) > 0 Then
          SendKeys "{tab}"
          Exit Sub
        End If
        sfiltra(2, 1) = "Codigo": sfiltra(2, 2) = "acodigo"
        sfiltra(1, 1) = "Descripcion": sfiltra(2, 2) = "adescri"
        FrmAyuda2.TipoForma = 1
        
        FrmAyuda2.BConexion = VGCNx
        FrmAyuda2.BTabla = "[" & VGCNx.DefaultDatabase & "].dbo.maeart "
        FrmAyuda2.bdata = "5"
        FrmAyuda2.bdato = Escadena(Trim(MBox2(1).ClipText))
        FrmAyuda2.BCampos = "acodigo as Codigo,adescri as Descripcion,apeso as Densidad"
        FrmAyuda2.BCondi = " apeso> 0"
        FrmAyuda2.BOrden = "adescri"
        FrmAyuda2.BFiltro = sfiltra
        FrmAyuda2.Show 1
        MBox2(0) = Escadena(nAyuda):   Label8 = Escadena(nDetalle)
        txtcanti(3) = Escadena(nSaldo)
        txtdensidad1 = numero(ndensidad)
        txtcanti(5).SetFocus
     Else
        If Len(Label5) > 0 Then
          SendKeys "{tab}"
          Exit Sub
        End If
        sfiltra(2, 1) = "Codigo": sfiltra(2, 2) = "acodigo"
        sfiltra(1, 1) = "Descripcion": sfiltra(2, 2) = "adescri"
        FrmAyuda2.TipoForma = 1
        FrmAyuda2.BConexion = VGCNx
        FrmAyuda2.BTabla = "[" & VGCNx.DefaultDatabase & "].dbo.maeart "
        FrmAyuda2.bdata = "5"
        FrmAyuda2.bdato = Escadena(Trim(MBox2(1).ClipText))
        FrmAyuda2.BCampos = "acodigo as Codigo,adescri as Descripcion, apeso as densidad "
        FrmAyuda2.BCondi = "apeso> 0"
        FrmAyuda2.BOrden = "adescri"
        FrmAyuda2.BFiltro = sfiltra
        FrmAyuda2.Show 1
        MBox2(1) = Escadena(nAyuda):   Label5 = Escadena(nDetalle)
        Txtdensidad2 = numero(ndensidad)
        txtcanti(1).SetFocus
     
     End If
End Sub
Private Sub CmbSerie_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
   SendKeys "{tab}"
   KeyAscii = 0
 End If
End Sub

Private Sub cmdBotones_Click(Index As Integer)
  Select Case Index
    Case 0
       MBox2(1) = ""
       txtcanti(0) = "": txtcanti(1) = "": Label5 = ""
       Call CargaGrilla
       Ctr_Ayuda1.SetFocus
    Case 11
      If rsdeta1.RecordCount > 0 Then
        GrabarData
      Else
        MsgBox "Debe ingresar productos...verifique!!!", vbInformation, "AVISO"
        Exit Sub
      End If
    Case Else
      Set rsdeta = Nothing
      Unload Me
  End Select
End Sub


Private Sub Ctr_Ayutransa_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
 If ColecCampos("tt_dr") = "S" Then
 
 End If
End Sub

Private Sub Ctr_Ayumaquina1_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
FACTOR = ColecCampos("factormaquina") * numero(txtcanti(5)) * txtdensidad1
If rsdeta1.RecordCount > 0 Then
   rsdeta1.MoveLast
   posi = IIf(IsNull(rsdeta1.Fields("item")), 0, rsdeta1.Fields("item"))
 Else
   posi = 0
End If
 Set rssaldo = VGCNx.Execute("select saldo=stsLkdis*1 from stkLOTE where stSalma='" & Ctr_Ayuda1.xclave & "' and stScodigo='" & MBox2(0) & "' and stslote='" & Textlote.text & "' and round(stsLkdis,2) >= " & Format(FACTOR, "0.00") & "")
If rssaldo.RecordCount <= 0 Then
   MsgBox " No existe saldo disponible...!!", vbInformation, "AVISO"
   Exit Sub
End If
If numero(txtcanti(5)) > 0 And Label8 <> "" Then
   rsdeta1.AddNew
   rsdeta1.Fields(0) = posi + 1
   rsdeta1.Fields(1) = Escadena(MBox2(0))
   rsdeta1.Fields(2) = Left(Escadena(Label8) & Space(65), 65)
   rsdeta1.Fields(3) = ""
   rsdeta1.Fields(5) = 0  'Format(numero(txtcanti(4)), "##,###,##0.00")
   rsdeta1.Fields(4) = Format(numero(FACTOR), "##,###,##0.00")
   rsdeta1.Fields(6) = Format(numero(txtcanti(5)), "##,###,##0.00")
   rsdeta1.Fields(7) = Ctr_Ayumaquina1.xclave
   rsdeta1.Fields(8) = Textlote.text
   rsdeta1.Update
   ConfigGrid
   MBox2(0) = "": txtcanti(3) = 0
   txtcanti(5) = 0: Label8 = ""
End If
MBox2(0).SetFocus
End Sub

Private Sub Ctr_Ayumaquina2_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
FACTOR = ColecCampos("factormaquina") * numero(txtcanti(1)) * numero(Txtdensidad2)
If rsdeta.RecordCount > 0 Then
   rsdeta.MoveLast
   posi = IIf(IsNull(rsdeta1.Fields("item")), 0, rsdeta1.Fields("item"))
 Else
   posi = 0
End If
If numero(txtcanti(1)) > 0 And Label5 <> "" Then
   rsdeta.AddNew
   rsdeta.Fields(0) = posi + 1
   rsdeta.Fields(1) = Escadena(MBox2(1))
   rsdeta.Fields(2) = Left(Escadena(Label5) & Space(65), 65)
   rsdeta.Fields(3) = ""
   rsdeta.Fields(5) = 0  'Format(numero(txtcanti(4)), "##,###,##0.00")
   rsdeta.Fields(4) = Format(numero(FACTOR), "##,###,##0.00")
   rsdeta.Fields(6) = Format(numero(txtcanti(1)), "##,###,##0.00")
   rsdeta.Fields(7) = Ctr_Ayumaquina2.xclave
   rsdeta.Fields(8) = Textlote.text
   rsdeta.Update
   ConfigGrid
   MBox2(1) = "": txtcanti(3) = 0
   txtcanti(1) = 0: Label5 = ""
End If
MBox2(1).SetFocus
End Sub

Private Sub Ctr_Ayusalida_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
 Ctr_Ayuingreso.xclave = ColecCampos("tt_codtrans_auto").Value
 Ctr_Ayuingreso.Ejecutar
 Ctr_Ayumerma.xclave = ColecCampos("tt_codtrans_merma").Value
 Ctr_Ayumerma.Ejecutar
 
End Sub

Private Sub Form_Load()
    Dim SQL As String
    central Me
    
    DTPicker1 = Date
    Call Ctr_Ayuda1.Conexion(VGCNx)
    Call Ctr_Ayusalida.Conexion(VGCNx): Ctr_Ayusalida.filtro = "tt_tipmov='S' and rtrim(tt_codtrans_auto)<>''"
    Call Ctr_Ayuingreso.Conexion(VGCNx)
    Call Ctr_Ayumerma.Conexion(VGCNx)
    Call Ctr_Ayumaquina1.Conexion(VGCNx)
    Call Ctr_Ayumaquina2.Conexion(VGCNx)
    If Not ExisteElem(0, VGCNx, "al_tipomaquina") Then
        SQL = " Create Table al_tipomaquina (codigomaquina nvarchar(4),descripcionmaquina nvarchar(30),factormaquina float " & _
        " CONSTRAINT Clave PRIMARY KEY (codigomaquina))"
        VGCNx.Execute SQL
       MsgBox " No existe Registros de tipo de amquina...Verifique!!", vbInformation, "AVISO"
       Exit Sub
  End If
  Call CargaGrilla
    
End Sub

Public Function GrabarData() As Integer
    Dim J As Integer
    Dim nsql As String
    Dim ltipo As String
    Dim lzona As String
    Dim xserie As String * 3
    Dim xfactu As Double
    Dim xtipofac As String * 2
    Dim ndato As String
    
    Dim ACMD As New ADODB.Command
    Dim asql As New ADODB.Recordset
    Dim arbusca As New ADODB.Recordset
    Dim wCabe(40)
    Dim nroreg As Integer
       
   On Error GoTo error
   
    GrabarData = 0
    
    
    '******** CABECERA DE MOVIMIENTO *****************
    For J = 1 To 29
        wCabe(J) = ""
    Next J
    Label4(0) = "": Label4(1) = ""
    
   
    Set asql = VGCNx.Execute("select * from  num_documentos where ctncodigo='TR'")
    If asql.RecordCount > 0 Then
        ndato = Right("00000000000" & Trim(CStr(asql!ctnnumero + 1)), 11)                  'nro pedido"
    Else
       MsgBox " No existe documentos de transacciones...Verifique!!", vbInformation, "AVISO"
       asql.Close
       Set asql = Nothing
       Exit Function
    End If
    asql.Close
    Set asql = Nothing

    VGCNx.Execute "update num_documentos " & _
                    " set ctnnumero=ctnnumero+1 " & _
                    " where ctncodigo='TR'"

    
    For J = 1 To 3
        wCabe(1) = g_ptoventa                        'Pto Venta
        Set asql = Nothing
        If J = 1 Or J = 3 Then
            ' de Almacen origen
           Set asql = VGCNx.Execute("select * from tabalm where taalma='" & Ctr_Ayuda1.xclave & "'")
           If asql.RecordCount > 0 Then
               wCabe(2) = Right("00000000000" & Trim(CStr(asql!tanumsal + 1)), 11)                     'nro pedido"
           End If
           asql.Close
           Set asql = Nothing
           VGCNx.Execute "update tabalm " & _
                           " set tanumsal=tanumsal+1 " & _
                           " where taalma='" & Ctr_Ayuda1.xclave & "'"
                           
           Label4(0) = wCabe(2)
        Else
            ' al almacen destino
           Set asql = VGCNx.Execute("select * from tabalm where taalma='" & Ctr_Ayuda1.xclave & "'")
           If asql.RecordCount > 0 Then
               wCabe(2) = Right("00000000000" & Trim(CStr(asql!tanument + 1)), 11)                     'nro pedido"
           End If
           asql.Close
           Set asql = Nothing
           VGCNx.Execute "update tabalm " & _
                           " set tanument=tanument+1 " & _
                           " where taalma='" & Ctr_Ayuda1.xclave & "'"
           Label4(1) = wCabe(2)
        End If
        wCabe(3) = ndato                      'nro factura
        wCabe(4) = "TR"                      'nro boleta
        wCabe(5) = ""                      'nro guia
        wCabe(6) = 0                       'dscto gral
        
        wCabe(7) = "DD"                    'tipo documento
        wCabe(8) = Text4.text           'nro de guia
         
        wCabe(9) = g_tiposol               'moneda
        wCabe(10) = 0                      'tipo de cambio
        wCabe(11) = 0                      'lista de precios
        wCabe(12) = ""                'mensajes
        If J = 1 Then
           wCabe(13) = Ctr_Ayusalida.xclave                      'modo de venta
         ElseIf J = 2 Then
             wCabe(13) = Ctr_Ayuingreso.xclave                      'modo de venta
           Else
              wCabe(13) = Ctr_Ayumerma.xclave                      'modo de venta
        End If
        wCabe(14) = DTPicker1               'fecha de atencion
        wCabe(15) = "00"                   'forma de pago
        wCabe(16) = ""                     'cliente
        wCabe(17) = ""                     'vendedor
        wCabe(18) = 0                      'comision
        wCabe(19) = Ctr_Ayuda1.xclave           'almacen
        wCabe(20) = 0                     'otros gastos
        wCabe(21) = 0                     'nota pedido
        wCabe(22) = 0                     'orden de compra
        wCabe(23) = 0                     'autorizacion
        wCabe(24) = 0                     'dias pago
        wCabe(25) = 0                     'Total Cantidad
        wCabe(26) = 0                     'Total Bruto
        wCabe(27) = 0                     'total fletes --T.D.
        wCabe(28) = 0                     'Total Igv
        wCabe(29) = 0         'Neto a Facturar
        wCabe(30) = ""             'entrega pedido
        wCabe(31) = ""                    'nombre cliente
        wCabe(32) = ""                    'direccion
        wCabe(33) = ""                    'ruc
        wCabe(34) = DTPicker1                           'fechafactura
        wCabe(35) = 0                     'Total Descuentos Globales
        wCabe(36) = 0                    'Total Descuentos Cliente
        wCabe(37) = 0                  'Total Descuentos Oficina
        wCabe(38) = 0                       'Total Descuentos Item
        wCabe(39) = 0                      'Total Descuentos Linea
        wCabe(40) = 0                      'Total Descuentos x Promocion
        
        Set ACMD.ActiveConnection = VGgeneral
        ACMD.CommandType = adCmdStoredProc
        ACMD.CommandText = "al_ingresoalma_pro"
        ACMD.CommandTimeout = 0
        ACMD.Prepared = True
        With ACMD
            .Parameters("@base") = VGCNx.DefaultDatabase
            .Parameters("@tabla") = "movalmcab"
            If J = 1 Or J = 3 Then
              .Parameters("@tipo") = "2"
            Else
              .Parameters("@tipo") = "3"
            End If
            .Parameters("@puntovta") = wCabe(1)
            .Parameters("@numero") = wCabe(2)
            .Parameters("@factura") = wCabe(8)
            .Parameters("@boleta") = wCabe(7)
            .Parameters("@nrotransf") = wCabe(3)
            .Parameters("@tipotransf") = wCabe(4)
            .Parameters("@guia") = wCabe(5)
            .Parameters("@dsctoglobal") = wCabe(6)
            .Parameters("@dsctoppago") = wCabe(6)
            .Parameters("@dsctovtaofi") = wCabe(6)
            .Parameters("@moneda") = IIf(wCabe(9) = g_tiposol, "S", "D")
            .Parameters("@tipocambio") = wCabe(10)
            .Parameters("@listaprecio") = wCabe(11)
            .Parameters("@mensaje") = wCabe(12)
            .Parameters("@modoventa") = wCabe(13)
            .Parameters("@fecha") = wCabe(14)
            .Parameters("@formapago") = wCabe(15)
            .Parameters("@cliente") = wCabe(16)
            .Parameters("@vendedor") = wCabe(17)
            .Parameters("@porcomision") = wCabe(18)
            .Parameters("@almacen") = wCabe(19)
            .Parameters("@totalotros") = wCabe(20)
            .Parameters("@notaped") = wCabe(21)
            .Parameters("@ordencompra") = wCabe(22)
            .Parameters("@autoriza") = wCabe(23)
            .Parameters("@diaspago") = wCabe(24)
            .Parameters("@totalitem") = wCabe(25)
            .Parameters("@totalbruto") = wCabe(26)
            .Parameters("@totalflete") = wCabe(27)
            .Parameters("@totalimpuesto") = wCabe(28)
            .Parameters("@totalneto") = wCabe(29)
            .Parameters("@usuario") = "star"
            .Parameters("@fechaactual") = Date
            .Parameters("@totaldsctoxlinea") = wCabe(39)
            .Parameters("@montodsctoppago") = 0
            .Parameters("@entregapedido") = wCabe(30)
            .Parameters("@razon") = wCabe(31)
            .Parameters("@direccion") = wCabe(32)
            .Parameters("@ruc") = wCabe(33)
            .Parameters("@fechafactura") = wCabe(34)
            .Parameters("@TDGlobal") = wCabe(35)
            .Parameters("@TDCliente") = wCabe(36)
            .Parameters("@TDOficina") = wCabe(37)
            .Parameters("@TDItem") = wCabe(38)
            .Parameters("@TDPromo") = wCabe(40)
        End With
        ACMD.Execute
        Set ACMD = Nothing
        DoEvents
          
       '** Actualizamos detalle
       
   
        If rsdeta1.RecordCount > 0 Then
            If J = 1 Or J = 3 Then
               rsdeta1.MoveFirst
               nroreg = 0
               Do Until rsdeta1.EOF
                 nroreg = nroreg + 1
                 If J = 2 Then
                          If dllgeneral.VerificaDatoExistente(VGCNx, "select * from stklote where stsalma='" & Ctr_Ayuda1.xclave & "' and stscodigo='" & Trim(rsdeta.Fields(1)) & "' and stslote='" & Textlote.text & "'") = 0 Then
                                          VGCNx.Execute "insert into stklote " & _
                                            "(stsalma,stscodigo,stslote,stskdis)" & _
                                            " Values ('" & Ctr_Ayuda1.xclave & "','" & Trim(rsdeta.Fields(1)) & "','" & Trim(rsdeta.Fields(8)) & "','" & Trim(rsdeta.Fields(4)) & "',0)"
                        End If
                End If
                
                Set ACMD.ActiveConnection = VGgeneral
                ACMD.CommandType = adCmdStoredProc
                ACMD.CommandTimeout = 0
                ACMD.CommandText = "vt_ingresodetallealma_pro"
                ACMD.Prepared = True
                With ACMD
                    .Parameters("@base") = VGCNx.DefaultDatabase
                    .Parameters("@tabla") = "movalmdet" ' nsql
                    .Parameters("@tipo") = "2"
                    .Parameters("@item") = nroreg
                    .Parameters("@numero") = wCabe(2)
                    .Parameters("@almacen") = Trim(Ctr_Ayuda1.xclave)
                    .Parameters("@producto") = Trim(rsdeta1.Fields(1))   'Trim(MBox2(1).Text)
                    .Parameters("@unidad") = ""
                    If J = 1 Then
                       .Parameters("@cantidad") = Trim(rsdeta1.Fields(4))   'Trim(txtcanti(1).Text)
                     Else
                       .Parameters("@cantidad") = Trim(rsdeta1.Fields(5))   'Trim(txtcanti(1).Text)
                     End If
                    .Parameters("@preciopacto") = 0
                    .Parameters("@dsctoxitem") = 0
                    .Parameters("@importebruto") = 0
                    .Parameters("@porcomision") = 0
                    .Parameters("@mdsctoitem") = 0
                    .Parameters("@mdsctoxlinea") = 0
                    .Parameters("@mdsctoxprom") = 0
                    .Parameters("@mimpor") = 0
                    .Parameters("@unidadref") = Trim(rsdeta1.Fields(5))   'rtxtcanti(1)
                End With
                ACMD.Execute
                Set ACMD = Nothing
                            
                Set ACMD.ActiveConnection = VGgeneral
                ACMD.CommandType = adCmdStoredProc
                ACMD.CommandTimeout = 0
                ACMD.CommandText = "al_actualizoalma_lote_pro"
                ACMD.Prepared = True
                With ACMD
                    .Parameters("@basedes") = CStr(VGCNx.DefaultDatabase)
                    .Parameters("@almacen") = wCabe(19)
                    .Parameters("@tipo") = "1"
                    .Parameters("@articulo") = Trim(rsdeta1.Fields(1))   'Trim(MBox2(1).Text)
                    .Parameters("@lote") = Trim(rsdeta1.Fields(8))   'Trim(MBox2(1).Text)
                    If J = 1 Then
                      .Parameters("@cantidad") = Trim(rsdeta1.Fields(4))   'txtcanti(1)
                     Else
                      .Parameters("@cantidad") = Trim(rsdeta1.Fields(5))   'txtcanti(1)
                    End If
                End With
                ACMD.Execute
                Set ACMD = Nothing
                rsdeta1.MoveNext
            
           Loop
                
          Else
               rsdeta.MoveFirst
               nroreg = 0
               Do Until rsdeta.EOF
                 nroreg = nroreg + 1
                 If J = 2 Then
                          If dllgeneral.VerificaDatoExistente(VGCNx, "select * from stklote where stsalma='" & Ctr_Ayuda1.xclave & "' and stslote='" & Trim(rsdeta.Fields(8)) & "' and stscodigo='" & Trim(rsdeta.Fields(1)) & "'") = 0 Then
                                          VGCNx.Execute "insert into stkLOTE " & _
                                            "(stsalma,stscodigo,stslote,stsLkdis,stscodprov)" & _
                                            " Values ('" & Ctr_Ayuda1.xclave & "','" & Trim(rsdeta.Fields(1)) & "','" & Trim(rsdeta.Fields(8)) & "',0,'')"
                        End If
                End If
                
                Set ACMD.ActiveConnection = VGgeneral
                ACMD.CommandType = adCmdStoredProc
                ACMD.CommandTimeout = 0
                ACMD.CommandText = "vt_ingresodetallealma_pro"
                ACMD.Prepared = True
                With ACMD
                    .Parameters("@base") = VGCNx.DefaultDatabase
                    .Parameters("@tabla") = "movalmdet" ' nsql
                    .Parameters("@tipo") = "3"
                    .Parameters("@item") = nroreg
                    .Parameters("@numero") = wCabe(2)
                    .Parameters("@almacen") = Trim(Ctr_Ayuda1.xclave)
                    .Parameters("@producto") = Trim(rsdeta.Fields(1))   'Trim(MBox2(1).Text)
                    .Parameters("@unidad") = ""
                    .Parameters("@cantidad") = Trim(rsdeta.Fields(4))   'Trim(txtcanti(1).Text)
                    .Parameters("@preciopacto") = 0
                    .Parameters("@dsctoxitem") = 0
                    .Parameters("@importebruto") = 0
                    .Parameters("@porcomision") = 0
                    .Parameters("@mdsctoitem") = 0
                    .Parameters("@mdsctoxlinea") = 0
                    .Parameters("@mdsctoxprom") = 0
                    .Parameters("@mimpor") = 0
                    .Parameters("@unidadref") = Trim(rsdeta.Fields(5))   'rtxtcanti(1)
                End With
                ACMD.Execute
                Set ACMD = Nothing
                            
                Set ACMD.ActiveConnection = VGgeneral
                ACMD.CommandType = adCmdStoredProc
                ACMD.CommandTimeout = 0
                ACMD.CommandText = "al_actualizoalma_lote_pro"
                ACMD.Prepared = True
                With ACMD
                    .Parameters("@basedes") = CStr(VGCNx.DefaultDatabase)
                    .Parameters("@almacen") = wCabe(19)
                    .Parameters("@tipo") = "2"
                    .Parameters("@articulo") = Trim(rsdeta.Fields(1))   'Trim(MBox2(1).Text)
                    .Parameters("@lote") = Trim(rsdeta.Fields(8))
                    .Parameters("@cantidad") = Trim(rsdeta.Fields(4))   'txtcanti(1)
                End With
                ACMD.Execute
                Set ACMD = Nothing
                If J = 1 Or J = 3 Then
                    rsdeta1.MoveNext
                 Else
                    rsdeta.MoveNext
                End If
          Loop
                
          End If
                
       
       End If
    Next
   
    GrabarData = 1
    MsgBox "Traslado de almacen satisfactorio...!!", vbInformation, "AVISO"
 '   If MsgBox("¿Desea Imprimir?", vbYesNo + vbQuestion, "Aviso") Then
 '          imprimirguias
 '   End If
 Exit Function
error:
   If Err Then
      MsgBox "Comunicarse con Sistemas ...!!" & Chr(13) & Chr(10) & VGCNx.Errors(0).Number & "-" & VGCNx.Errors(0).Description
    
    Resume Next
      Exit Function
   End If
 End Function

Private Sub MBox_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then
    MBox2(1).SetFocus
  End If
End Sub

Private Sub MBox2_Change(Index As Integer)
  If Len(Trim(MBox2(1).ClipText)) = 0 Then
    Label5 = ""
  End If
  If Len(Trim(MBox2(0).ClipText)) = 0 Then
    Label8 = ""
  End If
  
End Sub

Private Sub MBox2_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  Dim nsql As String
  Dim rabusca As New ADODB.Recordset
  
  If KeyCode = 13 Then
     If dllgeneral.ValidaCadena(Trim(MBox2(1).ClipText), "N") = False Then
        MBox2(1).MaxLength = 64
        Call cAyuda_Click(3)
        MBox2(1).MaxLength = 20
     '   SendKeys "{tab}"
     Else
        MBox2(1).MaxLength = 20
        nsql = "select * from maeart inner join stkart on acodigo=stcodigo  where stcodigo='" & MBox2(1).ClipText & "' and stalma='" & Ctr_Ayuda1.xclave & "' "
        Set rabusca = VGCNx.Execute(nsql)
        If rabusca.RecordCount > 0 Then
          Label5 = Escadena(rabusca!ADESCRI)
          If stockcomp Then
             txtcanti(1) = Round(rabusca!STSKDIS, 3) - Round(rabusca!STSKcom, 3)
           Else
             txtcanti(1) = Round(rabusca!STSKDIS, 3)
          End If
          txtcanti(2) = txtcanti(1)
        Else
          MsgBox "No existe articulo...!!", vbInformation, "AVISO"
          rabusca.Close
          Set rabusca = Nothing
          Exit Sub
        End If
        rabusca.Close
        txtcanti(1).SetFocus
       ' cmdBotones(11).SetFocus
     End If
 End If
 Set rabusca = Nothing
 
End Sub
Private Sub TDBGrid0_Click()
   If rsdeta1.RecordCount > 0 Then
      TDBGrid0.SetFocus
   End If
End Sub
Private Sub TDBGrid1_Click()
   If rsdeta.RecordCount > 0 Then
      TDBGrid1.SetFocus
   End If
End Sub

Private Sub TDBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
  Dim nvalor As String
  If KeyCode = 46 Then
     If rsdeta.RecordCount <= 0 Then
        MBox2(1) = ""
        txtcanti(0) = "": txtcanti(1) = "": Label5 = ""
        Exit Sub
     End If
     nvalor = TDBGrid1.Columns(0).text
     If rsdeta.RecordCount > 0 Then
        rsdeta.MoveFirst
        Do Until rsdeta.EOF
          If rsdeta.Fields(0) = nvalor Then
            rsdeta.Delete adAffectCurrent
            rsdeta.Update
            Exit Do
          End If
          rsdeta.MoveNext
        Loop
     End If
     ConfigGrid
     MBox2(1).SetFocus
  End If
End Sub
Private Sub TDBGrid0_keydown(KeyCode As Integer, Shift As Integer)
  Dim nvalor As String
  If KeyCode = 46 Then
     If rsdeta1.RecordCount <= 0 Then
        MBox2(1) = ""
        txtcanti(0) = "": txtcanti(1) = "": Label5 = ""
        Exit Sub
     End If
     nvalor = TDBGrid0.Columns(0).text
     If rsdeta1.RecordCount > 0 Then
        rsdeta1.MoveFirst
        Do Until rsdeta1.EOF
          If rsdeta1.Fields(0) = nvalor Then
            rsdeta1.Delete adAffectCurrent
            rsdeta1.Update
            Exit Do
          End If
          rsdeta1.MoveNext
        Loop
     End If
     ConfigGrid
     MBox2(1).SetFocus
  End If
End Sub
Private Sub Text3_DblClick()
Dim Adodc3 As ADODB.Recordset
Set Adodc3 = New ADODB.Recordset
Adodc3.Open "SELECT TDO_TIPDOC,TDO_DESCRI  FROM TIPO_DOCU", VGCNx, adOpenStatic, adLockOptimistic
frmReferencia.Conectar Adodc3, "SELECT TDO_TIPDOC,TDO_DESCRI  FROM TIPO_DOCU"
frmReferencia.Label1.Caption = "Tipo de Documentos"
frmReferencia.Show vbModal
Adodc3.Close
'If vGUtil(1) <> "" Then Text3 = (vGUtil(1))
'If Text3 <> "" Then
'        Text4.SetFocus
'End If
End Sub
Private Sub Text3_LostFocus()
   On Error Resume Next
   Call Text3_KeyDown(13, 0)
End Sub
Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)
 Dim rst As New ADODB.Recordset
 If KeyCode = 112 Then
    Text3_DblClick
 ElseIf KeyCode = 13 Then
 '  If Text3.text = "GR" Then
 '     Set rst = VGcnx.Execute("select * from vt_puntovtadocumento where puntovtacodigo='" & GPunto & "' and documentocodigo='" & Trim(Text3) & "'")
 '     If rst.RecordCount > 0 Then
 '        CmbSerie.Clear
 '        Do Until rst.EOF
 '           CmbSerie.AddItem rst!puntovtadocserie
 '           Text4.text = Trim(rst!puntovtadoccorr)
 '           rst.MoveNext
 '        Loop
 '        CmbSerie.ListIndex = 0
 '
 '     End If
 '     rst.Close
 '   End If
 End If
 
 Set rst = Nothing
End Sub

Private Sub TextLOTE_DblClick()
 VGcod = MBox2(0)
 VGAlma = Ctr_Ayuda1.xclave
 FrmAyuLote.Show 1
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
     SendKeys "{tab}"
End If
End Sub

Private Sub txtcanti_KeyPress(Index As Integer, KeyAscii As Integer)
  Dim posi As Integer
  Dim rssaldo As New ADODB.Recordset
  Dim qq As String
  
  If KeyAscii = 13 Then
      If Index = 5 Then
         txtcanti(5) = Format(txtcanti(5), "#####,##0.00")
          If Val(txtcanti(5)) <= 0 Then
              MsgBox "Cantidad debe ser mayor a Cero..!!", vbInformation, "AVISO"
              Exit Sub
          End If
       End If
       If Index = 1 Then
          If Val(txtcanti(1)) <= 0 Then
              MsgBox "Cantidad debe ser mayor a Cero..!!", vbInformation, "AVISO"
              Exit Sub
          End If
       End If
      End If
  End Sub


Public Function Escadena(pdato) As String
   If IsNull(pdato) Or Len(Trim(pdato)) = 0 Then
     Escadena = ""
   Else
     Escadena = Trim(pdato)
   End If
End Function
