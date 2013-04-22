VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Begin VB.Form FrmTraslado 
   Caption         =   "Traslado entre Almacenes"
   ClientHeight    =   7065
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11355
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   7065
   ScaleWidth      =   11355
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FrmReq 
      BackColor       =   &H00C9955A&
      BorderStyle     =   0  'None
      Caption         =   "Pendientes"
      Height          =   3555
      Left            =   660
      TabIndex        =   56
      Top             =   2565
      Visible         =   0   'False
      Width           =   10380
      Begin VB.CommandButton CmdCan 
         Caption         =   "Cancelar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8910
         Style           =   1  'Graphical
         TabIndex        =   58
         Top             =   180
         Width           =   1275
      End
      Begin TrueOleDBGrid70.TDBGrid TDBGrid 
         Height          =   2850
         Left            =   135
         TabIndex        =   57
         Top             =   540
         Width           =   10050
         _ExtentX        =   17727
         _ExtentY        =   5027
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "fecha"
         Columns(0).DataField=   "fecha"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "tipo Doc"
         Columns(1).DataField=   "tipoordencodigo"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Nro orden"
         Columns(2).DataField=   "nroorden"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "razon Social"
         Columns(3).DataField=   "oc_crazsoc"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Articulo"
         Columns(4).DataField=   "producto"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "Est_Ord,"
         Columns(5).DataField=   "Est_Ord,"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "Destino"
         Columns(6).DataField=   "Destino"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "tipodeorden"
         Columns(7).DataField=   "tipodeorden"
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   8
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   688
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   14215660
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=8"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=1482"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1402"
         Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(5)=   "Column(1).Width=2725"
         Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2646"
         Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(9)=   "Column(2).Width=2249"
         Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=2170"
         Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(13)=   "Column(3).Width=3440"
         Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=3360"
         Splits(0)._ColumnProps(16)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(17)=   "Column(4).Width=2725"
         Splits(0)._ColumnProps(18)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(19)=   "Column(4)._WidthInPix=2646"
         Splits(0)._ColumnProps(20)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(21)=   "Column(5).Width=1191"
         Splits(0)._ColumnProps(22)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(23)=   "Column(5)._WidthInPix=1111"
         Splits(0)._ColumnProps(24)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(25)=   "Column(6).Width=2725"
         Splits(0)._ColumnProps(26)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(27)=   "Column(6)._WidthInPix=2646"
         Splits(0)._ColumnProps(28)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(29)=   "Column(7).Width=2725"
         Splits(0)._ColumnProps(30)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(31)=   "Column(7)._WidthInPix=2646"
         Splits(0)._ColumnProps(32)=   "Column(7).Order=8"
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
         DeadAreaBackColor=   14215660
         RowDividerColor =   14215660
         RowSubDividerColor=   14215660
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
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=66,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=63,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=64,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=65,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=32,.parent=13"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
         _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=50,.parent=13"
         _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
         _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
         _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=46,.parent=13"
         _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=43,.parent=14"
         _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=44,.parent=15"
         _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=45,.parent=17"
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
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "REQUERIMIENTOS INTERNOS PENDIENTES"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   195
         Left            =   195
         TabIndex        =   59
         Top             =   210
         Width           =   3390
      End
   End
   Begin VB.CommandButton Cmdbotones 
      Caption         =   "&Grabar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1005
      Index           =   11
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   5940
      Width           =   1125
   End
   Begin VB.CommandButton Cmdbotones 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1005
      Index           =   12
      Left            =   8820
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   5940
      Width           =   1125
   End
   Begin VB.CommandButton Cmdbotones 
      Caption         =   "&Ingreso"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1005
      Index           =   0
      Left            =   10080
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   5940
      Width           =   1125
   End
   Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
      Height          =   1980
      Left            =   45
      TabIndex        =   24
      Top             =   3735
      Width           =   11190
      _ExtentX        =   19738
      _ExtentY        =   3493
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
   Begin VB.Frame Frame2 
      Height          =   1035
      Left            =   1935
      TabIndex        =   33
      Top             =   5970
      Width           =   5490
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   4
         Left            =   4080
         TabIndex        =   64
         Top             =   600
         Width           =   1245
      End
      Begin VB.Label Label3 
         Caption         =   "Nro."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   3
         Left            =   3480
         TabIndex        =   63
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label3 
         Caption         =   "Tipo Transf  TR"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   2
         Left            =   3720
         TabIndex        =   62
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Nota Salida :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   0
         Left            =   210
         TabIndex        =   37
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Nota Ingreso :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   165
         TabIndex        =   36
         Top             =   720
         Width           =   1185
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   0
         Left            =   1440
         TabIndex        =   35
         Top             =   180
         Width           =   1245
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   1
         Left            =   1455
         TabIndex        =   34
         Top             =   660
         Width           =   1245
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1425
      Left            =   45
      TabIndex        =   29
      Top             =   45
      Width           =   11190
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5160
         MaxLength       =   10
         TabIndex        =   6
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3000
         MaxLength       =   4
         TabIndex        =   5
         Top             =   1080
         Width           =   615
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1515
         MaxLength       =   2
         TabIndex        =   4
         Top             =   1035
         Width           =   495
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuda2 
         Height          =   375
         Left            =   6975
         TabIndex        =   3
         Top             =   615
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   661
         XcodMaxLongitud =   0
         xcodwith        =   100
         NomTabla        =   "tabalm"
         TituloAyuda     =   "Almacenes"
         ListaCampos     =   "TAALMA(1),TADESCRI(1),empresacodigo(1)"
         XcodCampo       =   "TAALMA"
         XListCampo      =   "TADESCRI"
         ListaCamposDescrip=   "Codigo,Descripcion,empresa"
         ListaCamposText =   "TAALMA,TADESCRI,empresacodigo"
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayusalida 
         Height          =   375
         Left            =   1515
         TabIndex        =   0
         Top             =   240
         Width           =   3960
         _ExtentX        =   6985
         _ExtentY        =   661
         XcodMaxLongitud =   0
         xcodwith        =   100
         NomTabla        =   "tabtransa"
         TituloAyuda     =   "Transaciones"
         ListaCampos     =   "tt_codmov(1),tt_descri(1),tt_dr(1),tt_codtrans_auto(1),tt_clie(2),tt_dr(2),intercompanias(1),tt_equip(1)"
         XcodCampo       =   "tt_codmov"
         XListCampo      =   "tt_descri"
         ListaCamposDescrip=   "Codigo,Descripcion,doc.ref.,trans.auto,Ctrl.Cliente,Doc.ref.Proyectos"
         ListaCamposText =   "tt_codmov,tt_descri,tt_dr,tt_codtrans_auto,tt_clie,tt_dr,intercompanias,tt_equip"
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuingreso 
         Height          =   375
         Left            =   6960
         TabIndex        =   1
         Top             =   240
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   661
         XcodMaxLongitud =   0
         xcodwith        =   100
         NomTabla        =   "tabtransa"
         TituloAyuda     =   "Transaciones"
         ListaCampos     =   "tt_codmov(1),tt_descri(1),tt_dr(1),tt_codtrans_auto(1)"
         XcodCampo       =   "tt_codmov"
         XListCampo      =   "tt_descri"
         ListaCamposDescrip=   "Codigo,Descripcion,doc.ref.,trans.auto"
         ListaCamposText =   "tt_codmov,tt_descri,tt_dr,tt_codtrans_auto"
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuda1 
         Height          =   375
         Left            =   1515
         TabIndex        =   2
         Top             =   645
         Width           =   3960
         _ExtentX        =   6985
         _ExtentY        =   661
         XcodMaxLongitud =   0
         xcodwith        =   100
         NomTabla        =   "tabalm"
         TituloAyuda     =   "Almacenes"
         ListaCampos     =   "TAALMA(1),TADESCRI(1),empresacodigo(1)"
         XcodCampo       =   "TAALMA"
         XListCampo      =   "TADESCRI"
         ListaCamposDescrip=   "Codigo,Descripcion,empresa"
         ListaCamposText =   "TAALMA,TADESCRI,empresacodigo"
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   255
         Left            =   7560
         TabIndex        =   7
         Top             =   1080
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   450
         _Version        =   393216
         Format          =   97452033
         CurrentDate     =   41125
         MinDate         =   39814
      End
      Begin MSMask.MaskEdBox MBox 
         Height          =   210
         Index           =   10
         Left            =   9600
         TabIndex        =   65
         Top             =   1080
         Visible         =   0   'False
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   370
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   8
         PromptChar      =   "_"
      End
      Begin VB.Label lblSerie 
         AutoSize        =   -1  'True
         Caption         =   "Serie :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2370
         TabIndex        =   55
         Top             =   1065
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Doc. Referencia :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   165
         TabIndex        =   54
         Top             =   1065
         Width           =   1260
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Num. Doc :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   4245
         TabIndex        =   53
         Top             =   1065
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Trans. Ingreso :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   5715
         TabIndex        =   49
         Top             =   240
         Width           =   1170
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Trans. Salida :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   165
         TabIndex        =   48
         Top             =   285
         Width           =   1035
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Almacen Origen :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   165
         TabIndex        =   32
         Top             =   675
         Width           =   1230
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Almacen Destino :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   5580
         TabIndex        =   31
         Top             =   645
         Width           =   1290
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   6960
         TabIndex        =   30
         Top             =   1080
         Width           =   615
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
      TabIndex        =   26
      Top             =   5790
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
         TabIndex        =   28
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
         TabIndex        =   27
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.Frame Frame6 
      Height          =   1035
      Left            =   45
      TabIndex        =   38
      Top             =   1470
      Width           =   11175
      Begin VB.TextBox Text7 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         TabIndex        =   10
         Top             =   600
         Width           =   5775
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4170
         TabIndex        =   9
         Top             =   195
         Width           =   5385
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1395
         TabIndex        =   8
         Top             =   195
         Width           =   1440
      End
      Begin VB.CommandButton Bdire 
         Caption         =   "..."
         Height          =   285
         Left            =   10485
         TabIndex        =   39
         Top             =   1035
         Width           =   375
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyuAnalitico 
         Height          =   390
         Left            =   6960
         TabIndex        =   11
         Top             =   600
         Visible         =   0   'False
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   688
         XcodMaxLongitud =   11
         xcodwith        =   900
         NomTabla        =   "gr_proyectos"
         TituloAyuda     =   "Busqueda de Proyectos"
         ListaCampos     =   "proyectocodigo(1),proyectodescripcion(1)"
         XcodCampo       =   "proyectocodigo"
         XListCampo      =   "proyectodescripcion"
         ListaCamposDescrip=   "Código,Descripción"
         ListaCamposText =   "proyectocodigo,proyectodescripcion"
         Requerido       =   0   'False
      End
      Begin VB.Label Lblanalitico 
         AutoSize        =   -1  'True
         Caption         =   "Analitico"
         Height          =   195
         Left            =   9840
         TabIndex        =   66
         Top             =   285
         Width           =   600
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Dirección :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   165
         TabIndex        =   42
         Top             =   645
         Width           =   750
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Razon Social :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   3060
         TabIndex        =   41
         Top             =   225
         Width           =   1005
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Cod. Ciente :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   180
         TabIndex        =   40
         Top             =   240
         Width           =   960
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   " Producto     "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   45
      TabIndex        =   22
      Top             =   2520
      Width           =   11190
      Begin VB.TextBox txtcanti 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   0
         Left            =   3480
         TabIndex        =   14
         Top             =   735
         Width           =   1125
      End
      Begin VB.TextBox txtcanti 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   1
         Left            =   7680
         TabIndex        =   16
         Top             =   735
         Width           =   1080
      End
      Begin VB.TextBox txtuni 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   1
         Left            =   5790
         TabIndex        =   15
         Top             =   735
         Width           =   780
      End
      Begin VB.TextBox txtuni 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   1470
         TabIndex        =   13
         Top             =   735
         Width           =   840
      End
      Begin VB.TextBox txtcanti 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Index           =   2
         Left            =   9270
         TabIndex        =   21
         Top             =   270
         Width           =   900
      End
      Begin VB.CommandButton cAyuda 
         Caption         =   "..."
         Height          =   285
         Index           =   3
         Left            =   1530
         TabIndex        =   20
         Top             =   270
         Width           =   285
      End
      Begin MSMask.MaskEdBox MBox2 
         Height          =   330
         Index           =   1
         Left            =   180
         TabIndex        =   12
         Top             =   270
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   8
         PromptChar      =   "_"
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   6885
         TabIndex        =   52
         Top             =   780
         Width           =   750
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Un. Refer :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   4920
         TabIndex        =   51
         Top             =   780
         Width           =   810
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Un. Medida :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   225
         TabIndex        =   50
         Top             =   780
         Width           =   915
      End
      Begin VB.Label Label2 
         Caption         =   "Saldo :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   8730
         TabIndex        =   47
         Top             =   360
         Width           =   480
      End
      Begin VB.Label Label5 
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   1905
         TabIndex        =   23
         Top             =   270
         Width           =   6600
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   2655
         TabIndex        =   25
         Top             =   780
         Width           =   750
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Direcciones"
      Height          =   2760
      Left            =   2430
      TabIndex        =   43
      Top             =   2970
      Visible         =   0   'False
      Width           =   7065
      Begin VB.CommandButton cCerrar 
         Caption         =   "Cerrar"
         Height          =   345
         Left            =   5550
         TabIndex        =   46
         Top             =   2160
         Width           =   1365
      End
      Begin VB.CommandButton cAcepta 
         Caption         =   "&Acepta"
         Height          =   345
         Left            =   4260
         TabIndex        =   45
         Top             =   2160
         Width           =   1155
      End
      Begin MSDataGridLib.DataGrid dbGrid1 
         Height          =   1605
         Left            =   210
         TabIndex        =   44
         Top             =   120
         Width           =   6645
         _ExtentX        =   11721
         _ExtentY        =   2831
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
            MarqueeStyle    =   3
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   180
      Top             =   6300
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label LblTipOrd 
      Height          =   285
      Left            =   4005
      TabIndex        =   61
      Top             =   6660
      Visible         =   0   'False
      Width           =   1860
   End
   Begin VB.Label LblNroOrd 
      Height          =   285
      Left            =   1980
      TabIndex        =   60
      Top             =   6660
      Visible         =   0   'False
      Width           =   1860
   End
End
Attribute VB_Name = "FrmTraslado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dllgeneral As New dllgeneral.dll_general
Dim rsdeta As New ADODB.Recordset
Dim flag As Integer
Dim nropedido As String
Dim empresaorigen As String
Dim empresadestino As String
Dim intercompanias As String
Dim Nroreq As String
Dim Interestablecimientos As String
Dim ruc As String, empresa As String

Dim analitico As Integer
Dim RsEmpresa As ADODB.Recordset
Dim wCabe(40)
Dim RsRq As ADODB.Recordset
Dim RsRq2 As ADODB.Recordset

Public Function CargaGrilla()
   Set rsdeta = Nothing
   Call rsdeta.Fields.Append("Item", adInteger)
   Call rsdeta.Fields.Append("Codigo", adChar, 20)
   Call rsdeta.Fields.Append("Descripcion", adChar, 100)
   Call rsdeta.Fields.Append("UM", adChar, 3)
   Call rsdeta.Fields.Append("Cant", adDouble)
   Call rsdeta.Fields.Append("CantRef", adDouble)
   
   rsdeta.Open
   ConfigGrid

End Function

Public Function ConfigGrid()
   Set TDBGrid1.DataSource = Nothing
   
   Set TDBGrid1.DataSource = rsdeta
   With TDBGrid1
      .Columns(0).Width = 600
      .Columns(0).Caption = "Item"
      .Columns(1).Width = 1700
      .Columns(1).Caption = "Codigo"
      .Columns(2).Width = 5600
      .Columns(2).Caption = "Descripcion"
      .Columns(3).Width = 800
      .Columns(3).Caption = "UM"
      .Columns(4).Width = 1200
      .Columns(4).Caption = "Cant"
      .Columns(4).NumberFormat = "##,###,##0.00"
      .Columns(5).Width = 1200
      .Columns(5).Caption = "Cant.Ref"
      .Columns(5).NumberFormat = "##,###,##0.00"
   End With
   TDBGrid1.Refresh
End Function

Private Sub Bdire_Click()
  Dim rg As New ADODB.Recordset
  On Error Resume Next
   Set rg = Nothing
   Set DbGrid1.DataSource = Nothing
   
   Set rg = VGCNx.Execute("select cliedirnumero as Nro,cliedirdireccion as Direccion from vt_clientedireccion where clientecodigo='" & Text5 & "'")
   If rg.RecordCount > 0 Then
       Frame7.Visible = True
       Set DbGrid1.DataSource = rg
       DbGrid1.Refresh
   Else
      Frame7.Visible = False
   End If
   
End Sub
Private Sub cAcepta_Click()
   Text7 = IIf(IsNull(DbGrid1.Columns(1).text), "", DbGrid1.Columns(1).text)
   Frame7.Visible = False
   Text7.SetFocus

End Sub

Private Sub cAyuda_Click(Index As Integer)
Dim RSQL As New ADODB.Recordset
    If Index = 3 Then
        If Len(Label5) > 0 Then
          SendKeys "{tab}"
          Exit Sub
        End If
        Dim sfiltra(1 To 2, 1 To 2) As String
   '     sfiltra(1, 1) = "Codigo": sfiltra(1, 2) = "acodigo"
   '     sfiltra(2, 1) = "Descripcion": sfiltra(2, 2) = "adescri"
        sfiltra(2, 1) = "Codigo": sfiltra(2, 2) = "acodigo"
        sfiltra(1, 1) = "Descripcion": sfiltra(1, 2) = "adescri"
        
        FrmAyuda2.TipoForma = 1
        
        FrmAyuda2.BConexion = VGCNx
        If Ctr_Ayuda1.xclave <> Ctr_Ayuda2.xclave Then
           FrmAyuda2.BTabla = "[" & VGCNx.DefaultDatabase & "].dbo.maeart inner join [" & _
                            VGCNx.DefaultDatabase & "].dbo.stkart " & _
                            " ON acodigo=stcodigo"
           FrmAyuda2.bdata = "4"
           FrmAyuda2.bdato = Escadena(Trim(MBox2(1).ClipText))
           If stockcomp Then
              FrmAyuda2.BCampos = "acodigo as Codigo,adescri as Descripcion,(stskdis-stskcom) as stock"
              FrmAyuda2.BCondi = "stalma='" & Ctr_Ayuda1.xclave & "' and (stskdis-stskcom)>0"
            Else
              FrmAyuda2.BCampos = "acodigo as Codigo,adescri as Descripcion,stskdis as stock"
              FrmAyuda2.BCondi = "stalma='" & Ctr_Ayuda1.xclave & "' and stskdis>0"
           End If
         Else
           FrmAyuda2.BTabla = "[" & VGCNx.DefaultDatabase & "].dbo.maeart"
           FrmAyuda2.bdata = "4"
           FrmAyuda2.bdato = Escadena(Trim(MBox2(1).ClipText))
           FrmAyuda2.BCampos = "acodigo as Codigo,adescri as Descripcion,0 as stock"
           FrmAyuda2.BCondi = "acodigo <>'xxxx'"
         End If
        FrmAyuda2.BOrden = "adescri"
        FrmAyuda2.BFiltro = sfiltra
        FrmAyuda2.Show 1
        MBox2(1) = Escadena(nAyuda):   Label5 = Escadena(nDetalle)
        txtcanti(2) = nsaldo
        Set RSQL = VGCNx.Execute("select * from maeart where acodigo ='" & MBox2(1) & "'")
        If RSQL.RecordCount() > 0 Then txtuni(0) = ESNULO(RSQL!aunidad, "")
        txtcanti(0).SetFocus
     End If
End Sub

Private Sub cCerrar_Click()
Frame7.Visible = False
End Sub

Private Sub CmdCan_Click()
FrmReq.Visible = False
Interestablecimientos = ""
End Sub

Private Sub Ctr_Ayuda1_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
empresaorigen = ColecCampos("empresacodigo")
Ctr_Ayuda2.filtro = "taalma<>'" & Ctr_Ayuda1.xclave & "'"
End Sub

Private Sub Ctr_Ayuda2_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)

Interestablecimientos = ""
SQL = "select tipo=1,pedidonumero=space(11),oc_dfecdoc as Fecha,tipoordencodigo,oc_cnumord as NroOrden,oc_ccodpro as Producto,"
SQL = SQL & " oc_crazsoc=c.empresadescripcion,estadooccodigo as Est_Ord,destino=b.tadescri  from co_cabordcompra a"
SQL = SQL & " left join tabalm b on a.almacendestino=b.taalma left join co_multiempresas c on b.empresacodigo=c.empresacodigo"
SQL = SQL & " where estadooccodigo <>'5' and "
SQL = SQL & " almacenorigen ='" & Ctr_Ayuda1.xclave & "' and almacendestino ='" & Ctr_Ayuda2.xclave & "'"

Set RsRq = VGCNx.Execute(SQL)
If RsRq.RecordCount > 0 Then
   FrmReq.Visible = True
Else
   FrmReq.Visible = False
End If
empresa = Ctr_Ayuda2.xclave
empresadestino = ColecCampos("empresacodigo")
TDBGrid.DataSource = RsRq
TDBGrid.Refresh


End Sub

Private Sub DTPicker1_Change()
DTPicker1.Value = UltimoCierreFech(DTPicker1.Value)
MBox(10) = DTPicker1.Value
End Sub

Private Sub TDBGrid_DblClick()
Dim n As Integer
Nroreq = RsRq!tipoordencodigo + RsRq!NroOrden
If RsRq!tipo = 1 Then

   SQL = "select item,Codigo,referenciaItem as Descripcion,'' as UM, "
   SQL = SQL & " (cantid-sum(isnull(decantid,0)))  as Cantidad,'' as CantRF from v_ordenes a "
   SQL = SQL & " left join v_kardex b on tipoordencodigo+numeroorden=b.canumord and a.codigo=b.decodigo "
   SQL = SQL & " where  tipoordencodigo+numeroorden= '" & Nroreq & "' and isnull(catipmov,'I')='I'"
   SQL = SQL & " group by tipoordencodigo,numeroorden,item,codigo,referenciaItem, cantid having a.cantid > sum(isnull(decantid,0)) "
   Set RsRq2 = VGCNx.Execute(SQL)
   Do While Not RsRq2.EOF
     rsdeta.AddNew
     rsdeta.Fields(0) = RsRq2!item
     rsdeta.Fields(1) = RsRq2!codigo
     rsdeta.Fields(2) = RsRq2!descripcion
     rsdeta.Fields(3) = ""
     rsdeta.Fields(4) = RsRq2!CANTIDAD
     rsdeta.Fields(5) = 0
     rsdeta.Update
     RsRq2.MoveNext
   Loop
Else
   Set RsRq2 = VGCNx.Execute("select pedidonumero, productocodigo as Codigo,adescri as Descripcion,'' as UM, " _
   & " detpedcantpedida- sum(decantid) as Cantidad,'' as CantRF from v_almacenyventas " _
   & " where pedidotipofac='" & RsRq!tipoordencodigo & "' and pedidonrofact='" & RsRq!NroOrden & "' and isnull(catd,'NS')<>'NI'" _
   & " group by pedidonumero, productocodigo ,adescri,detpedcantpedida ")
   n = 0
   nropedido = RsRq!pedidonumero
   Do While Not RsRq2.EOF
     n = n = 1
     rsdeta.AddNew
     rsdeta.Fields(0) = Format(n, "000")
     rsdeta.Fields(1) = RsRq2!codigo
     rsdeta.Fields(2) = RsRq2!descripcion
     rsdeta.Fields(3) = ""
     rsdeta.Fields(4) = RsRq2!CANTIDAD
     rsdeta.Fields(5) = 0
     rsdeta.Update
     RsRq2.MoveNext
   Loop
End If
Interestablecimientos = "S"
TDBGrid1.DataSource = RsRq2
TDBGrid1.Refresh

FrmReq.Visible = False
LblNroOrd.Caption = RsRq!NroOrden
LblTipOrd.Caption = RsRq!tipoordencodigo

End Sub


Private Sub Text1_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
   Text4.SetFocus
   KeyAscii = 0
 End If
End Sub

Private Sub cmdBotones_Click(Index As Integer)
Dim RSQL As New ADODB.Recordset
Dim rrsql As ADODB.Recordset
Dim totreq As String
totreq = "4"
Dim ok As Integer
Select Case Index
Case 0
     MBox2(1) = ""
     txtcanti(0) = "": txtcanti(1) = "": Label5 = ""
     Call CargaGrilla
     Ctr_Ayusalida.SetFocus

Case 11
    'If Ctr_Ayusalida.xclave = "51" Then
        If rsdeta.RecordCount > 0 Then
'            MsgBox "Debe ingresar productos...verifique!!!", vbInformation, "AVISO"
'            Exit Sub
'        End If
    'Else
'        If RsRq2.RecordCount <= 0 Then
'            MsgBox "Debe ingresar productos...verifique!!!", vbInformation, "AVISO"
'            Exit Sub
'        End If
'    End If

        If Len(Trim(Text3.text)) = 0 Then
            MsgBox "Falta seleccionar tipo de documento", vbInformation, "Sistema"
            Text3.SetFocus
            Exit Sub
        End If
        ok = 1
        rsdeta.MoveFirst
        Do While Not rsdeta.EOF
           SQL = " select * from stkart where stalma='" & Ctr_Ayuda1.xclave & "' and stcodigo='" & RTrim(rsdeta.Fields(1)) & "'"
           Set rrsql = VGCNx.Execute(SQL)
           If rrsql.RecordCount() = 0 Then
              MsgBox (" producto " & rsdeta.Fields(2) & " No tiene saldo disponible , Saldo Actual 0 ")
              ok = 0
            ElseIf rrsql!STSKDIS - rsdeta.Fields(4) < 0 Then
              MsgBox (" producto " & RTrim(rsdeta.Fields(1)) & " - " & RTrim(rsdeta.Fields(2)) & " , No tiene saldo disponible , Saldo Actual " & rrsql!STSKDIS & "")
              ok = 0
           End If
           rsdeta.MoveNext
        Loop
        If ok = 0 Then
           Exit Sub
        End If
        GrabarData
        If intercompanias = "S" Or Interestablecimientos = "S" Then
           SQL = " select tipoordencodigo,numeroorden,codigo, cantid, dd=sum(isnull(decantid,0)) from v_ordenes a left join v_kardex b  "
           SQL = SQL & " on tipoordencodigo+numeroorden=b.canumord and a.codigo=b.decodigo "
           SQL = SQL & " where  tipoordencodigo+numeroorden= '" & Nroreq & "' and isnull(catipmov,'I')='I'"
           SQL = SQL & " group by tipoordencodigo,numeroorden,codigo, cantid having a.cantid > sum(isnull(decantid,0)) "
           Set RSQL = VGCNx.Execute(SQL)
           If RSQL.RecordCount = 0 Then
              totreq = "5"
            Else
              totreq = "4"
           End If
           VGCNx.Execute "update co_cabordcompra set estadooccodigo='" & totreq & "'  where oc_cnumord='" & LblNroOrd.Caption & "' " _
              & " and tipoordencodigo='" & LblTipOrd.Caption & "'"
        
            LblNroOrd.Caption = Empty
            LblTipOrd.Caption = Empty
        End If
        Interestablecimientos = ""
        intercompanias = ""
        If MsgBox("Desea Imprimir Notas de Almacen  ", vbYesNo) = vbYes Then Call imprimirNotas
        Ctr_Ayusalida.xclave = Empty: Ctr_Ayusalida.Ejecutar
        Ctr_Ayuingreso.xclave = Empty: Ctr_Ayuingreso.Ejecutar
        Ctr_Ayuda1.xclave = Empty: Ctr_Ayuda1.Ejecutar
        Ctr_Ayuda2.xclave = Empty: Ctr_Ayuda2.Ejecutar
        Text3.text = Empty
        Text1.text = Empty
        Text4.text = Empty
        Text5.text = Empty
        Text6.text = Empty
        Text7.text = Empty
        Ctr_Ayusalida.SetFocus
        
        Else
        
          MsgBox "Debe ingresar productos...verifique!!!", vbInformation, "AVISO"
            Exit Sub
        End If
Case Else
    Set rsdeta = Nothing
    Unload Me
End Select

End Sub

Private Sub Ctr_Ayusalida_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
analitico = 0
Ctr_Ayuingreso.xclave = ColecCampos("tt_codtrans_auto"): Ctr_Ayuingreso.Ejecutar
Ctr_Ayuingreso.Enabled = False

If ColecCampos("tt_clie") = "S" Then
    Frame6.Visible = True
    Text5 = Empty
    Text6 = Empty
    Text7 = Empty
Else
    Frame6.Visible = False
End If

If ColecCampos("tt_dr") = "S" Then
  Text4.Enabled = True
  Text5.Enabled = True
  Text1.Enabled = True
 Else
  Text4.Enabled = False
  Text5.Enabled = False
  Text1.Enabled = False
End If
Ctr_Ayuda1.filtro = "empresacodigo='" & VGParametros.empresacodigo & "' and puntovtacodigo='" & VGParametros.puntovta & "'"
If ESNULO(ColecCampos("intercompanias"), "N") = "S" Then
   Ctr_Ayuda2.filtro = "empresacodigo<>'" & VGParametros.empresacodigo & "' "
   intercompanias = ESNULO(ColecCampos("intercompanias"), "N")
 Else
   Ctr_Ayuda2.filtro = "empresacodigo='" & VGParametros.empresacodigo & "' "
   intercompanias = ESNULO(ColecCampos("intercompanias"), "N")
 End If
Ctr_Ayuda1.xclave = Empty: Ctr_Ayuda1.Ejecutar
Ctr_Ayuda2.xclave = Empty: Ctr_Ayuda2.Ejecutar
If ESNULO(ColecCampos("tt_equip"), "N") = "S" Then
   Lblanalitico.Visible = True
   Ctr_AyuAnalitico.Visible = True
   Ctr_AyuAnalitico.filtro = " tipoanaliticocodigo='" & VGParamSistem.tipoanaliticocodigo & "'"
   analitico = 1
 Else
   Lblanalitico.Visible = False
   Ctr_AyuAnalitico.Visible = False
End If
End Sub


Private Sub Form_Load()
    
central Me
DTPicker1.MaxDate = VGParamSistem.fechatrabajo
DTPicker1.Value = UltimoCierreFech(CDate(Format(VGParamSistem.fechatrabajo, "dd/MM/yyyy")))
MBox(10) = DTPicker1.Value
Call Ctr_Ayuda1.Conexion(VGCNx)
Call Ctr_Ayuda2.Conexion(VGCNx)
Call Ctr_Ayusalida.Conexion(VGCNx): Ctr_Ayusalida.filtro = "tt_tipmov='S' and rtrim(tt_codtrans_auto)<>''"
Call Ctr_Ayuingreso.Conexion(VGCNx)
Call Ctr_AyuAnalitico.Conexion(VGCNx)
Call CargaGrilla
    
cmdBotones(11).Picture = MDIPrincipal.ImageList2.ListImages.item("Grabar").Picture
cmdBotones(12).Picture = MDIPrincipal.ImageList2.ListImages.item("Retornar").Picture
cmdBotones(0).Picture = MDIPrincipal.ImageList2.ListImages.item("Insertar").Picture

End Sub

Public Function GrabarData() As Integer
Dim J As Integer
Dim nsql As String
Dim ltipo As String
Dim lzona As String
Dim xserie As String * 3
Dim xfactu As Double  'String * 8
Dim xtipofac As String * 2
Dim ndato As String

Dim acmd As New ADODB.Command
Dim asql As New ADODB.Recordset
Dim arbusca As New ADODB.Recordset
Dim nroreg As Integer
   
On Error GoTo error

GrabarData = 0
    
'******** CABECERA DE MOVIMIENTO *****************
For J = 1 To 29
    wCabe(J) = ""
Next J
Label4(0) = "": Label4(1) = ""

Set asql = New ADODB.Recordset
VGCNx.BeginTrans
asql.Open "select * from  num_documentos where ctncodigo='TR'", VGCNx, adOpenDynamic, adLockOptimistic
If asql.RecordCount > 0 Then
    ndato = Right("00000000000" & Trim(CStr(ESNULO(asql!ctnnumero, 0))), 11)                   'nro pedido"
Else
   MsgBox " No existe documentos de transacciones...Verifique!!", vbInformation, "AVISO"
   asql.Close
   Set asql = Nothing
   Exit Function
End If
asql.Close
Set asql = Nothing

    VGCNx.Execute "update num_documentos set ctnnumero=ctnnumero+1  where ctncodigo='TR'"
VGCNx.CommitTrans
    For J = 1 To 2
        wCabe(1) = VGParametros.puntovta                        'Pto Venta
        Set asql = Nothing
        VGCNx.BeginTrans
        If J = 1 Then
            ' de Almacen origen
           Set asql = Nothing
           asql.Open "select * from tabalm where taalma='" & Ctr_Ayuda1.xclave & "'", VGCNx, adOpenDynamic, adLockOptimistic
           empresaorigen = asql!empresacodigo
           If asql.RecordCount > 0 Then
               wCabe(2) = Right("00000000000" & Trim(CStr(asql!tanumsal)), 11)                       'nro pedido"
           End If
           VGCNx.Execute "update tabalm set tanumsal=tanumsal+1 where taalma='" & Ctr_Ayuda1.xclave & "'"
           Label4(0) = wCabe(2)
           wCabe(13) = Ctr_Ayusalida.xclave
           asql.Close
           Set asql = Nothing
         Else
            ' al almacen destino
           Set asql = VGCNx.Execute("select * from tabalm where taalma='" & Ctr_Ayuda2.xclave & "'")
           empresadestino = asql!empresacodigo
           If asql.RecordCount > 0 Then
               wCabe(2) = Right("00000000000" & Trim(CStr(asql!tanument)), 11)                       'nro pedido"
           End If
           asql.Close
           Set asql = Nothing
           VGCNx.Execute "update tabalm " & _
                           " set tanument=tanument+1 " & _
                           " where taalma='" & Ctr_Ayuda2.xclave & "'"
           Label4(1) = wCabe(2)
           wCabe(13) = Ctr_Ayuingreso.xclave
        End If
        VGCNx.CommitTrans
       wCabe(3) = ndato                      'nro factura
       Label4(4).Caption = wCabe(3)
        wCabe(4) = "TR"                      'nro boleta
        wCabe(5) = ""                      'nro guia
        wCabe(6) = 0                       'dscto gral
'        If UCase(Text3) = "GR" Then
          wCabe(7) = UCase(Text3) 'Text3                       'tipo documento
          wCabe(8) = Text1 & Text4              'nro de guia SOLO ACEPTA 11 CARACTERES
'        End If
        wCabe(9) = g_tiposol               'moneda
        wCabe(10) = 0                      'tipo de cambio
        wCabe(11) = 0                      'lista de precios
        wCabe(12) = ""                     'mensajes
        wCabe(14) = MBox(10)               'fecha de atencion
        wCabe(15) = "00"                   'forma de pago
        wCabe(16) = ""                     'cliente
        wCabe(17) = ""                     'vendedor
        wCabe(18) = 0                      'comision
        If J = 1 Then
           wCabe(19) = Ctr_Ayuda1.xclave           'almacen
        Else
           wCabe(19) = Ctr_Ayuda2.xclave           'almacen
        End If
        wCabe(20) = 0                     'otros gastos
        wCabe(21) = nropedido                    'nota pedido
        wCabe(22) = 0                     'orden de compra
        wCabe(23) = 0                     'autorizacion
        wCabe(24) = 0                     'dias pago
        wCabe(25) = 0                     'Total Cantidad
        wCabe(26) = 0                     'Total Bruto
        wCabe(27) = 0                     'total fletes --T.D.
        wCabe(28) = 0                     'Total Igv
        wCabe(29) = 0         'Neto a Facturar
        wCabe(30) = Ctr_Ayuda2.xclave            'entrega pedido
        wCabe(31) = Trim(Text6.text)                    'nombre cliente
        wCabe(32) = Trim(Text7.text)                    'direccion
        wCabe(33) = Trim(Text5.text)                   'ruc
        wCabe(34) = MBox(10)                           'fechafactura
        wCabe(35) = 0                     'Total Descuentos Globales
        wCabe(36) = 0                    'Total Descuentos Cliente
        wCabe(37) = 0                  'Total Descuentos Oficina
        wCabe(38) = 0                       'Total Descuentos Item
        wCabe(39) = 0                      'Total Descuentos Linea
        wCabe(40) = 0                      'Total Descuentos x Promocion
        
        Set acmd.ActiveConnection = VGGeneral
        acmd.CommandType = adCmdStoredProc
        acmd.CommandText = "al_ingresoalma_pro"
        acmd.CommandTimeout = 0
        acmd.Prepared = True
        With acmd
            .Parameters("@base") = VGCNx.DefaultDatabase
            .Parameters("@tabla") = "movalmcab"
            If J = 1 Then
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
            .Parameters("@moneda") = IIf(wCabe(9) = g_tiposol, "01", "02")
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
            .Parameters("@usuario") = UCase(VGUsuario)
            .Parameters("@fechaactual") = Now
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
            If J = 1 Then
                .Parameters("@empresa") = empresaorigen
            Else
                .Parameters("@empresa") = empresadestino
            End If
            If Interestablecimientos = "S" Then .Parameters("@Nroreq") = Nroreq
           
        End With
        acmd.Execute
        Set acmd = Nothing
        DoEvents
          
       '** Actualizamos detalle
       

        If rsdeta.RecordCount > 0 Then
            rsdeta.MoveFirst
            nroreg = 0
            Do Until rsdeta.EOF
                nroreg = nroreg + 1
                If J = 1 Then
                   Set asql = VGCNx.Execute("select * from stkart where stalma='" & Ctr_Ayuda1.xclave & "' and stcodigo='" & Trim(rsdeta.Fields(1)) & "'")
                   If asql.RecordCount = 0 Then
                      VGCNx.Execute " insert into stkart (stalma,stcodigo,stskdis)" & _
                      " Values ('" & Ctr_Ayuda1.xclave & "','" & Trim(rsdeta.Fields(1)) & "',0)"
                   End If
                 Else
                   Set asql = VGCNx.Execute("select * from stkart where stalma='" & Ctr_Ayuda2.xclave & "' and stcodigo='" & Trim(rsdeta.Fields(1)) & "'")
                 End If
                If asql.RecordCount = 0 Then
                          VGCNx.Execute "insert into stkart (stalma,stcodigo,stskdis)" & _
                                  " Values ('" & Ctr_Ayuda2.xclave & "','" & Trim(rsdeta.Fields(1)) & "',0)"
                End If
                asql.Close
                Set acmd.ActiveConnection = VGGeneral
                acmd.CommandType = adCmdStoredProc
                acmd.CommandTimeout = 0
                acmd.CommandText = "vt_ingresodetallealma_pro"
                acmd.Prepared = True
                With acmd
                    .Parameters("@base") = VGCNx.DefaultDatabase
                    .Parameters("@tabla") = "movalmdet" ' nsql
                    If J = 1 Then
                      .Parameters("@tipo") = "2"
                    Else
                      .Parameters("@tipo") = "3"
                    End If
                    .Parameters("@item") = nroreg
                    .Parameters("@numero") = wCabe(2)
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
                     If J = 1 Then
                       .Parameters("@almacen") = Trim(Ctr_Ayuda1.xclave)
                     Else
                       .Parameters("@almacen") = Trim(Ctr_Ayuda2.xclave)
                     End If
                     .Parameters("@equipo") = Ctr_AyuAnalitico.xclave
                End With
                acmd.Execute
                Set acmd = Nothing
                            
                Set acmd.ActiveConnection = VGGeneral
                acmd.CommandType = adCmdStoredProc
                acmd.CommandTimeout = 0
                acmd.CommandText = "vt_actualizoalma_pro"
                acmd.Prepared = True
                With acmd
                    .Parameters("@basedes") = CStr(VGCNx.DefaultDatabase)
                    .Parameters("@almacen") = wCabe(19)
                    If J = 1 Then
                      .Parameters("@tipo") = "1"
                    Else
                      .Parameters("@tipo") = "2"
                    End If
                    .Parameters("@articulo") = Trim(rsdeta.Fields(1))   'Trim(MBox2(1).Text)
                    .Parameters("@cantidad") = Trim(rsdeta.Fields(4))   'txtcanti(1)
                End With
                acmd.Execute
                Set acmd = Nothing
                rsdeta.MoveNext
          Loop
       End If
    '------------------------------------------------------------------------------------------------------------
    
    
    '------------------------------------------------------------------------------------------------------------
    
    Next

    GrabarData = 1
    MsgBox "Traslado de almacen satisfactorio...!!", vbInformation, "AVISO"
    If Text3 = "GR" Then     'Text3 = "GR" And Frame6.Visible
        If MsgBox("¿Desea Imprimir?", vbYesNo + vbQuestion, "Aviso") = vbYes Then
           imprimirguias
        End If
   End If
   cmdBotones_Click (0)
 Exit Function
error:
   If Err Then
      MsgBox "Comunicarse con Sistemas ...!!" & Chr(13) & Chr(10) & VGCNx.Errors(0).Number & "-" & VGCNx.Errors(0).Description
      Exit Function
      Resume Next
   End If
 End Function

Private Sub imprimirguias()
Dim nguia As String, ReporteNombre As String
Dim nflag As Integer
Dim i As Integer
Dim arrparam(4) As Variant
Dim arrform(2) As Variant
Dim Serie As String
Dim RSQL As New ADODB.Recordset

arrparam(0) = VGParamSistem.BDEmpresa
arrparam(1) = Text1.text & Text4.text
arrparam(2) = VGParametros.empresacodigo
arrparam(3) = Ctr_Ayuda1.xclave

ReporteNombre = "Repguiaimpresa" & VGParametros.empresacodigo & ".rpt"

If Text3.text = "GR" Then

   SQL = "select puntovtadoccorr from vt_puntovtadocumento where empresacodigo='" & VGParametros.empresacodigo & "' and documentocodigo='GR' "
   SQL = SQL & " and  puntovtacodigo='" & VGParametros.puntovta & "' "
   SQL = SQL & " and puntovtadocserie='" & Text1 & "'"
     Set RSQL = VGCNx.Execute(SQL)
   If RSQL.RecordCount > 0 Then
      nguia = Right("00000000" & TraeDataSerie(SQL, VGCNx), 8)
      VGCNx.Execute "Update vt_puntovtadocumento " & _
       " set puntovtadoccorr='" & CStr(nguia) + 1 & "'" & _
       " Where empresacodigo='" & VGParametros.empresacodigo & "' and documentocodigo='GR' and puntovtacodigo='" & VGParametros.puntovta & "' and puntovtadocserie='" & Text1 & "'"
       arrparam(1) = Text1.text & nguia
      SQL = " update movalmcab set CARFTDOC='GR' , CARFNDOC='" & Text1.text & nguia & "'"
      SQL = SQL & " where catipotransf='TR' and canrotransf='" & wCabe(3) & "'"
      Set RSQL = Nothing
      Set RSQL = VGCNx.Execute(SQL)
   End If
End If

Call ImpresionRptProc(ReporteNombre, arrform, arrparam, "", "Impresion de guia")
Text1.text = Empty
Text3.text = Empty
Text4.text = Empty
Text5.text = Empty
Text6.text = Empty
Text7.text = Empty

For i = 0 To 1
    Label4(i).Caption = Empty
    txtuni(i).text = Empty
    txtcanti(i).text = Empty
Next i

txtcanti(2).text = Empty

nerror:
   If Err Then
      If nflag = 1 Then
         VGCNx.RollbackTrans
      End If
      MsgBox "Comunicarse con Sistemas ...!!" & Chr(13) & Chr(10) & Err.Number & "-" & Err.Description
      Exit Sub
   End If
  
End Sub


Private Sub MBox_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then
    MBox2(1).SetFocus
  End If
End Sub

Private Sub MBox2_Change(Index As Integer)
  If Len(Trim(MBox2(1).ClipText)) = 0 Then
    Label5 = ""
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
             txtcanti(0) = Round(rabusca!STSKDIS, 3) - Round(rabusca!STSKcom, 3)
           Else
             txtcanti(0) = Round(rabusca!STSKDIS, 3)
          End If
'          txtcanti(1) = txtcanti(0)
        Else
          MsgBox "No existe articulo...!!", vbInformation, "AVISO"
          rabusca.Close
          Set rabusca = Nothing
          Exit Sub
        End If
        txtuni(0).text = rabusca!aunidad
        rabusca.Close
        txtcanti(0).SetFocus
       ' cmdBotones(11).SetFocus
     End If
 End If
 Set rabusca = Nothing
 
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

Private Sub Text3_DblClick()
Dim Adodc3 As ADODB.Recordset
Set Adodc3 = New ADODB.Recordset
Adodc3.Open "SELECT TDO_TIPDOC,TDO_DESCRI  FROM TIPO_DOCU", VGCNx, adOpenStatic, adLockOptimistic
frmReferencia.Conectar Adodc3, "SELECT TDO_TIPDOC,TDO_DESCRI  FROM TIPO_DOCU"
frmReferencia.Label1.Caption = "Tipo de Documentos"
frmReferencia.Show vbModal
If vGUtil(1) <> "" Then
   Text3 = (vGUtil(1))
   Call Text3_KeyDown(13, 0)
End If
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
   Text3.text = UCase(Text3.text)
   If Text3.text = "GR" Then
      Set rst = VGCNx.Execute("select * from vt_puntovtadocumento where empresacodigo='" & VGParametros.empresacodigo & "' and puntovtacodigo='" & VGParametros.puntovta & "' and documentocodigo='" & Trim(UCase(Text3)) & "'")
      If rst.RecordCount > 0 Then
          Text1.text = rst!puntovtadocserie
          Text4.text = Trim(rst!puntovtadoccorr)
      End If
      rst.Close
   End If
       SendKeys "{tab}"
 End If
 Set rst = Nothing
End Sub
Private Sub Text4_DblClick()
If Text4 <> "" Then
     Text4_KeyPress (13)
End If
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)

 If KeyAscii = 13 And Text3 = "GR" Then
    If Text4 = "" Then
        'Text3 = "FT"
        MsgBox "Ingrese el número de Guia", vbInformation, "Aviso"
       Text4.SetFocus
     End If
 End If
 If KeyAscii = 13 Then
    SendKeys "{tab}"
     KeyAscii = 0
 End If
End Sub

Private Sub Text5_DblClick()
Dim acliente As New ADODB.Recordset
Text5 = ""
Text6 = ""
Text7 = ""
FrmAyuCliente.Show 1
Text5 = FrmAyuCliente.cCod
Text6 = FrmAyuCliente.cNom
Text7 = FrmAyuCliente.cDir
ruc = FrmAyuCliente.cRuc
If analitico = 1 Then
   Ctr_AyuAnalitico.Visible = True
   If Text5 <> "" Then
      SQL = " clientecodigo='" & Text5 & "' and proyectocierre=0 and tipoanaliticocodigo='" & VGParamSistem.tipoanaliticocodigo & "'"
      Set acliente = VGCNx.Execute(" select * from gr_proyectos where " & SQL)
      If acliente.RecordCount = 0 Then
         MsgBox ("No existe proyectos activos para este cliente ")
         Text5.SetFocus
         Ctr_AyuAnalitico.Visible = False
         Exit Sub
       Else
         Ctr_AyuAnalitico.filtro = SQL
      End If
   End If
End If
End Sub
Private Sub Text6_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 And Trim(Text7.text) <> "" Then
       Text7.SetFocus
  End If
End Sub

Private Sub txtcanti_KeyPress(Index As Integer, KeyAscii As Integer)
  Dim posi As Integer
 If KeyAscii = 13 Then
    txtcanti(1) = Format(txtcanti(1), "#####,##0.000")
    txtcanti(0) = Format(txtcanti(0), "#####,##0.000")
    If Index = 0 Then
      txtcanti(1) = txtcanti(0)
      txtuni(1) = ""
      SendKeys "{tab}"
    Else
      If rsdeta.RecordCount > 0 Then
        rsdeta.MoveLast
        posi = IIf(IsNull(rsdeta.Fields("item")), 0, rsdeta.Fields("item"))
      Else
        posi = 0
      End If
      txtcanti(0) = Format(txtcanti(0), "#######0.0000")
      txtcanti(1) = Format(txtcanti(1), "#######0.0000")
      If Val(txtcanti(1)) <= 0 Then
          MsgBox "Cantidad debe ser mayor a Cero..!!", vbInformation, "AVISO"
          Exit Sub
      End If
      Dim rssaldo As New ADODB.Recordset
     If stockcomp Then
        Set rssaldo = VGCNx.Execute("select saldo=(stskdis - stskcom ) from stkart where stalma='" & Ctr_Ayuda1.xclave & "' and stcodigo='" & MBox2(1) & "' and (round(stskdis,2) -round(stskcom,2)) >= " & txtcanti(1) & "")
      Else
        Set rssaldo = VGCNx.Execute("select stskdis from stkart where stalma='" & Ctr_Ayuda1.xclave & "' and stcodigo='" & MBox2(1) & "' and round(stskdis,2) >= " & txtcanti(1) & "")
     End If
     If rssaldo.RecordCount <= 0 And Ctr_Ayuda1.xclave <> Ctr_Ayuda2.xclave Then
         MsgBox " No existe saldo disponible...!!", vbInformation, "AVISO"
     Else
        rsdeta.AddNew
        rsdeta.Fields(0) = posi + 1
        rsdeta.Fields(1) = Escadena(MBox2(1))
        rsdeta.Fields(2) = Left(Escadena(Label5) & Space(65), 65)
        rsdeta.Fields(3) = ""
        rsdeta.Fields(4) = Format(txtcanti(1), "##,###,##0.000")
        rsdeta.Fields(5) = Format(numero(txtcanti(0)), "##,###,##0.000")
        rsdeta.Update
        ConfigGrid
        MBox2(1) = ""
        txtcanti(0) = "": txtcanti(1) = "": Label5 = ""
        txtuni(0) = "": txtcanti(1) = ""
        MBox2(1).SetFocus
        txtuni(1) = ""
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


Private Sub txtuni_DblClick(Index As Integer)
Dim rs As New ADODB.Recordset
Dim RSQL As String
Dim FACTOR As Double
VGabrev = txtuni(0).text
FACTOR = 1
If Index = 1 Then
    Frmayuunidades.Show 1
    txtuni(1) = VGabrev
    If Trim(txtuni(0)) <> Trim(txtuni(1)) Then                          'CONSULTA POR DEFECTO MODIFICAR
        RSQL = "select  p.EQCANTEQUI from TabEqui p where p.EQUNIPRI = '" & VGabrev & "'   and p.EQUNIEQUI = '" & txtuni(0).text & "'"
        Set rs = VGCNx.Execute(RSQL)
        If rs.RecordCount = 0 Then
            MsgBox "la unidad de referencia no tiene unidad equivalente"
            Exit Sub
        End If
        rs.MoveFirst
        FACTOR = rs.Fields("EQCANTEQUI")
        rs.Close
      Else
        FACTOR = 1
     End If
     SendKeys "{tab}"
End If

txtcanti(1) = Round(Val(Format(txtcanti(0), "#######0.000")) * FACTOR, 0) 'VGcant

End Sub

Private Sub txtuni_KeyPress(Index As Integer, KeyAscii As Integer)
If Index = 1 And KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txtuni_LostFocus(Index As Integer)
If Index = 1 And Trim(txtuni(1)) = "" Then txtcanti(1) = txtcanti(0)
End Sub
Private Sub imprimirNotas()
Dim aparam(4) As Variant
Dim aform(12) As Variant
Dim n As Integer
Dim rrsql As New ADODB.Recordset
n = 0
Do While n < 2
    aparam(0) = VGParamSistem.BDEmpresa
                                
    aform(0) = "fecha='" & MBox(10).text & "'"
    aform(1) = "xtrans = '" & Ctr_Ayusalida.xclave & "' "
    aform(2) = "xtd = '" & Text3.text & "' "
    aform(3) = "xndoc = '" & RTrim(Text1.text) & Text4.text & "' "
                                
     aform(8) = "NRef = '" & RTrim(Text1.text) & Text4.text & "' "
     aform(9) = "DocRef = '" & Text3.text & "' "
     aform(10) = "TTrans = '" & Ctr_Ayusalida.xclave & "' "
     aform(11) = "emp = '" & VGParametros.NomEmpresa & "'"
  
        If n = 0 Then
           aparam(1) = Ctr_Ayuda1.xclave
           aparam(2) = "NS"
           aparam(3) = Label4(0).Caption
           
           aform(0) = "ctitulo = ' NOTA DE SALIDA ' "
           aform(4) = "Xnalma = '" & Ctr_Ayuda1.xclave & "' "
           aform(5) = "Dalma = '" & Ctr_Ayuda1.xnombre & "' "
           aform(6) = "AlmaDes = '" & Ctr_Ayuda2.xclave & "' "
           aform(7) = "Dalmades = '" & Ctr_Ayuda2.xnombre & "' "
        Else
           aparam(1) = Ctr_Ayuda2.xclave
           aparam(2) = "NI"
           aparam(3) = Label4(1).Caption
           
           aform(0) = "ctitulo = ' NOTA DE INGRESO ' "
           aform(4) = "Xnalma = '" & Ctr_Ayuda1.xclave & "' "
           aform(5) = "Dalma = '" & Ctr_Ayuda1.xnombre & "' "
           aform(6) = "AlmaDes = '" & Ctr_Ayuda2.xclave & "' "
           aform(7) = "Dalmades = '" & Ctr_Ayuda2.xnombre & "' "
        End If
  
     If Not (Text3.text = "GR" And n = 0) Then Call ImpresionRptProc("al_notasAlmacen.rpt", aform, aparam, , "Impresion de Notas de Almacen ")
     n = n + 1
Loop

End Sub
