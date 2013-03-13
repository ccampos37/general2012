VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Object = "{4D137D9C-00A6-4458-9B46-7E95DB76D55B}#9.0#0"; "TextFer.ocx"
Begin VB.Form FrmOrdenes_Requerimientos 
   Caption         =   "Form1"
   ClientHeight    =   8370
   ClientLeft      =   -6210
   ClientTop       =   450
   ClientWidth     =   11175
   LinkTopic       =   "Form1"
   ScaleHeight     =   8370
   ScaleWidth      =   11175
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   7935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10965
      _ExtentX        =   19341
      _ExtentY        =   13996
      _Version        =   393216
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "Pendientes de Ingreso a Almacen"
      TabPicture(0)   =   "FrmOrdenes_Requerimientos.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "CmdSalir"
      Tab(0).Control(1)=   "cmdNue"
      Tab(0).Control(2)=   "CmdEli"
      Tab(0).Control(3)=   "cmdEdi"
      Tab(0).Control(4)=   "cmdImp"
      Tab(0).Control(5)=   "DataGrid1"
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Generacion de Ordenes"
      TabPicture(1)   =   "FrmOrdenes_Requerimientos.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "TDBGrid2"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "TDBGrid1"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Command1"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Emision de Ordenes"
      TabPicture(2)   =   "FrmOrdenes_Requerimientos.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "CrystalReport1"
      Tab(2).Control(1)=   "Flex1"
      Tab(2).Control(2)=   "Frame1"
      Tab(2).Control(3)=   "Data2"
      Tab(2).Control(4)=   "cmdNue2"
      Tab(2).Control(5)=   "cmdEli2"
      Tab(2).Control(6)=   "cmdEdi2"
      Tab(2).Control(7)=   "cmdGra"
      Tab(2).Control(8)=   "CmdSalir2"
      Tab(2).Control(9)=   "fraTotales"
      Tab(2).Control(10)=   "Fradatos"
      Tab(2).ControlCount=   11
      Begin VB.CommandButton Command1 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         Height          =   675
         Left            =   9120
         Picture         =   "FrmOrdenes_Requerimientos.frx":0054
         Style           =   1  'Graphical
         TabIndex        =   63
         Top             =   720
         Width           =   775
      End
      Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
         Height          =   2895
         Left            =   240
         TabIndex        =   54
         Top             =   1680
         Width           =   10455
         _ExtentX        =   18441
         _ExtentY        =   5106
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "familia"
         Columns(0).DataField=   "fam_nombre"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Nro.Req."
         Columns(1).DataField=   "OC_CNUMORD"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Item"
         Columns(2).DataField=   "oc_citem"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Codigo"
         Columns(3).DataField=   "OC_cCodigo"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Descripcion"
         Columns(4).DataField=   "descripcion"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "Unidad"
         Columns(5).DataField=   "oc_cunidad"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "Cantidad"
         Columns(6).DataField=   "oc_ncantid"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "Solicitante"
         Columns(7).DataField=   "solicitantenombre"
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   8
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   13160660
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=8"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2170"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2090"
         Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(5)=   "Column(1).Width=2011"
         Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=1931"
         Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(9)=   "Column(2).Width=688"
         Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=609"
         Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(13)=   "Column(3).Width=1640"
         Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=1561"
         Splits(0)._ColumnProps(16)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(17)=   "Column(4).Width=3519"
         Splits(0)._ColumnProps(18)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(19)=   "Column(4)._WidthInPix=3440"
         Splits(0)._ColumnProps(20)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(21)=   "Column(5).Width=1111"
         Splits(0)._ColumnProps(22)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(23)=   "Column(5)._WidthInPix=1032"
         Splits(0)._ColumnProps(24)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(25)=   "Column(6).Width=1376"
         Splits(0)._ColumnProps(26)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(27)=   "Column(6)._WidthInPix=1296"
         Splits(0)._ColumnProps(28)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(29)=   "Column(7).Width=5186"
         Splits(0)._ColumnProps(30)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(31)=   "Column(7)._WidthInPix=5106"
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
         _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=66,.parent=13"
         _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=63,.parent=14"
         _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=64,.parent=15"
         _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=65,.parent=17"
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
      Begin TrueOleDBGrid70.TDBGrid TDBGrid2 
         Height          =   2895
         Left            =   240
         TabIndex        =   60
         Top             =   4680
         Width           =   10455
         _ExtentX        =   18441
         _ExtentY        =   5106
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "familia"
         Columns(0).DataField=   "fam_nombre"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Nro.Req."
         Columns(1).DataField=   "OC_CNUMORD"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Item"
         Columns(2).DataField=   "oc_citem"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Codigo"
         Columns(3).DataField=   "OC_cCodigo"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Descripcion"
         Columns(4).DataField=   "descripcion"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "Unidad"
         Columns(5).DataField=   "oc_cunidad"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "Cantidad"
         Columns(6).DataField=   "oc_ncantid"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "Solicitante"
         Columns(7).DataField=   "solicitantenombre"
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   8
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   13160660
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=8"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2170"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2090"
         Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(5)=   "Column(1).Width=1984"
         Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=1905"
         Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(9)=   "Column(2).Width=688"
         Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=609"
         Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(13)=   "Column(3).Width=1640"
         Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=1561"
         Splits(0)._ColumnProps(16)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(17)=   "Column(4).Width=3519"
         Splits(0)._ColumnProps(18)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(19)=   "Column(4)._WidthInPix=3440"
         Splits(0)._ColumnProps(20)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(21)=   "Column(5).Width=1111"
         Splits(0)._ColumnProps(22)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(23)=   "Column(5)._WidthInPix=1032"
         Splits(0)._ColumnProps(24)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(25)=   "Column(6).Width=1376"
         Splits(0)._ColumnProps(26)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(27)=   "Column(6)._WidthInPix=1296"
         Splits(0)._ColumnProps(28)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(29)=   "Column(7).Width=5186"
         Splits(0)._ColumnProps(30)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(31)=   "Column(7)._WidthInPix=5106"
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
         _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=66,.parent=13"
         _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=63,.parent=14"
         _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=64,.parent=15"
         _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=65,.parent=17"
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
      Begin VB.Frame Frame2 
         Height          =   990
         Left            =   240
         TabIndex        =   55
         Top             =   480
         Width           =   10020
         Begin VB.CommandButton Command2 
            Caption         =   "&Eliminar"
            Height          =   675
            Left            =   8040
            Picture         =   "FrmOrdenes_Requerimientos.frx":0496
            Style           =   1  'Graphical
            TabIndex        =   64
            Top             =   240
            Width           =   775
         End
         Begin VB.CommandButton cmdOK 
            Caption         =   "&Aceptar"
            Height          =   675
            Left            =   7200
            Picture         =   "FrmOrdenes_Requerimientos.frx":08D8
            Style           =   1  'Graphical
            TabIndex        =   62
            Top             =   240
            Width           =   780
         End
         Begin VB.CommandButton Cmdsalirorden 
            Caption         =   "&Salir"
            Height          =   675
            Left            =   8160
            Picture         =   "FrmOrdenes_Requerimientos.frx":0D1A
            Style           =   1  'Graphical
            TabIndex        =   61
            Top             =   1440
            Width           =   775
         End
         Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyutipoOrden1 
            Height          =   270
            Left            =   960
            TabIndex        =   56
            Top             =   240
            Width           =   3390
            _ExtentX        =   5980
            _ExtentY        =   476
            XcodMaxLongitud =   11
            xcodwith        =   1100
            NomTabla        =   "co_tipodeorden"
            TituloAyuda     =   "Busqueda de Tipo de Orden"
            ListaCampos     =   "tipoordencodigo(1),tipoordendescripcion(1),tipoordennumeracion(2),ordendebienes(2)"
            XcodCampo       =   "tipoordencodigo"
            XListCampo      =   "tipoordendescripcion"
            ListaCamposDescrip=   "Codigo,Descripcion,numeracion,tipo de orden"
            ListaCamposText =   "tipoordencodigo,tipoordendescripcion,tipoordennumeracion,ordendebienes"
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Número  :"
            Height          =   195
            Left            =   4530
            TabIndex        =   59
            Top             =   285
            Width           =   690
         End
         Begin VB.Label Label16 
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
            Left            =   5220
            TabIndex        =   58
            Top             =   195
            Width           =   1560
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            Caption         =   "Tipo Orden     :"
            Height          =   192
            Left            =   96
            TabIndex        =   57
            Top             =   276
            Width           =   1032
         End
      End
      Begin VB.CommandButton CmdSalir 
         Caption         =   "&Salir"
         Height          =   675
         Left            =   -67920
         Picture         =   "FrmOrdenes_Requerimientos.frx":115C
         Style           =   1  'Graphical
         TabIndex        =   52
         Top             =   7065
         Width           =   775
      End
      Begin VB.CommandButton cmdNue 
         Caption         =   "&Nuevo"
         Height          =   675
         Left            =   -73185
         Picture         =   "FrmOrdenes_Requerimientos.frx":159E
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   7050
         Width           =   775
      End
      Begin VB.CommandButton CmdEli 
         Caption         =   "&Anular"
         Height          =   675
         Left            =   -70530
         Picture         =   "FrmOrdenes_Requerimientos.frx":19E0
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   7080
         Width           =   775
      End
      Begin VB.CommandButton cmdEdi 
         Caption         =   "&Editar"
         Height          =   675
         Left            =   -71850
         Picture         =   "FrmOrdenes_Requerimientos.frx":1E22
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   7065
         Width           =   775
      End
      Begin VB.CommandButton cmdImp 
         Caption         =   "&Imprimir"
         Height          =   675
         Left            =   -69240
         Picture         =   "FrmOrdenes_Requerimientos.frx":2264
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   7080
         Width           =   775
      End
      Begin VB.Frame Fradatos 
         Height          =   2508
         Left            =   -74730
         TabIndex        =   25
         Top             =   975
         Width           =   9708
         Begin VB.TextBox txtEntE 
            Height          =   288
            Left            =   1020
            MaxLength       =   80
            TabIndex        =   29
            Top             =   1308
            Width           =   8535
         End
         Begin VB.TextBox txtCot 
            Height          =   336
            Left            =   6288
            TabIndex        =   28
            Top             =   948
            Width           =   3312
         End
         Begin VB.TextBox txtObs 
            Height          =   288
            Left            =   1164
            MaxLength       =   80
            TabIndex        =   27
            Top             =   2028
            Width           =   8340
         End
         Begin VB.TextBox txtNSol 
            Height          =   288
            Left            =   8640
            MaxLength       =   10
            TabIndex        =   26
            Top             =   240
            Width           =   945
         End
         Begin ctrlayuda_f.Ctr_Ayuda CtrAyu_moneda 
            Height          =   348
            Left            =   6288
            TabIndex        =   30
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
            Height          =   285
            Left            =   1005
            TabIndex        =   31
            Top             =   585
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   503
            _Version        =   393216
            Format          =   31588353
            CurrentDate     =   37015
         End
         Begin MSComCtl2.DTPicker txtEnt 
            Height          =   285
            Left            =   3645
            TabIndex        =   32
            Top             =   585
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   503
            _Version        =   393216
            Format          =   31588353
            CurrentDate     =   37015
         End
         Begin TextFer.TxFer lblRuc 
            Height          =   300
            Left            =   6240
            TabIndex        =   33
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
            TabIndex        =   34
            Top             =   192
            Width           =   4500
            _ExtentX        =   7938
            _ExtentY        =   556
            XcodMaxLongitud =   11
            xcodwith        =   1100
            NomTabla        =   "cp_proveedor"
            TituloAyuda     =   "Busqueda de Proveedor"
            ListaCampos     =   "clientecodigo(1),clienterazonsocial(1),clienteruc(1)"
            XcodCampo       =   "clientecodigo"
            XListCampo      =   "clienterazonsocial"
            ListaCamposDescrip=   "Codigo,Descripcion,Ruc"
            ListaCamposText =   "clientecodigo,clienterazonsocial,clienteruc"
         End
         Begin ctrlayuda_f.Ctr_Ayuda CtrAyu_pago 
            Height          =   360
            Left            =   1008
            TabIndex        =   35
            Top             =   912
            Width           =   4116
            _ExtentX        =   7250
            _ExtentY        =   635
            XcodMaxLongitud =   3
            xcodwith        =   300
            NomTabla        =   "vt_formapago"
            TituloAyuda     =   "Busqueda de Condiciones de Pago"
            ListaCampos     =   "formapagocodigo(1),formapagodescripcion(1)"
            XcodCampo       =   "formapagocodigo"
            XListCampo      =   "formaPagodescripcion"
            ListaCamposDescrip=   "Codigo,Descripcion"
            ListaCamposText =   "formapagocodigo,formapagodescripcion"
         End
         Begin ctrlayuda_f.Ctr_Ayuda CtrAyu_solicitante 
            Height          =   312
            Left            =   1008
            TabIndex        =   36
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
         Begin VB.Label lblCen 
            AutoSize        =   -1  'True
            Caption         =   "Cotización  :"
            Height          =   192
            Left            =   5244
            TabIndex        =   47
            Top             =   960
            Width           =   876
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Solicitante     :"
            Height          =   192
            Left            =   84
            TabIndex        =   46
            Top             =   1680
            Width           =   1008
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Entregar en   :"
            Height          =   192
            Left            =   84
            TabIndex        =   45
            Top             =   1320
            Width           =   1008
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Emisión         :"
            Height          =   192
            Left            =   84
            TabIndex        =   44
            Top             =   600
            Width           =   996
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Proveedor     :"
            Height          =   192
            Left            =   48
            TabIndex        =   43
            Top             =   276
            Width           =   1008
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "R.U.C.  :"
            Height          =   192
            Left            =   5616
            TabIndex        =   42
            Top             =   288
            Width           =   552
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
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Moneda  :"
            Height          =   192
            Left            =   5448
            TabIndex        =   40
            Top             =   600
            Width           =   720
         End
         Begin VB.Label Label12 
            Caption         =   "Observación :"
            Height          =   252
            Left            =   84
            TabIndex        =   39
            Top             =   2040
            Width           =   1092
         End
         Begin VB.Label Le_Proveedor 
            Caption         =   "No. Requis."
            Height          =   252
            Left            =   7728
            TabIndex        =   38
            Top             =   288
            Width           =   1020
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Cond.Pago     :"
            Height          =   192
            Left            =   48
            TabIndex        =   37
            Top             =   996
            Width           =   1032
         End
      End
      Begin VB.Frame fraTotales 
         Height          =   975
         Left            =   -74745
         TabIndex        =   14
         Top             =   6045
         Visible         =   0   'False
         Width           =   9708
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Importe      :"
            Height          =   195
            Left            =   720
            TabIndex        =   24
            Top             =   240
            Width           =   840
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Descuento :"
            Height          =   195
            Left            =   720
            TabIndex        =   23
            Top             =   600
            Width           =   870
         End
         Begin VB.Label lblImp 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "892,760.00"
            Height          =   285
            Left            =   1680
            TabIndex        =   22
            Top             =   240
            Width           =   1110
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
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Total  :"
            Height          =   195
            Left            =   3600
            TabIndex        =   20
            Top             =   240
            Width           =   495
         End
         Begin VB.Label lblTot 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "892,760.00"
            Height          =   285
            Left            =   4200
            TabIndex        =   19
            Top             =   240
            Width           =   1230
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "I.G.V.   :"
            Height          =   195
            Left            =   6360
            TabIndex        =   18
            Top             =   240
            Width           =   585
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "Compra :"
            Height          =   195
            Left            =   6360
            TabIndex        =   17
            Top             =   600
            Width           =   630
         End
         Begin VB.Label lblIgv 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "892,760.00"
            Height          =   285
            Left            =   7080
            TabIndex        =   16
            Top             =   240
            Width           =   1110
         End
         Begin VB.Label lblCom 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "892,760.00"
            Height          =   285
            Left            =   7080
            TabIndex        =   15
            Top             =   600
            Width           =   1110
         End
      End
      Begin VB.CommandButton CmdSalir2 
         Caption         =   "&Salir"
         Height          =   675
         Left            =   -68010
         Picture         =   "FrmOrdenes_Requerimientos.frx":26A6
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   7110
         Visible         =   0   'False
         Width           =   775
      End
      Begin VB.CommandButton cmdGra 
         Caption         =   "&Grabar"
         Height          =   675
         Left            =   -69480
         Picture         =   "FrmOrdenes_Requerimientos.frx":2AE8
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   7110
         Visible         =   0   'False
         Width           =   775
      End
      Begin VB.CommandButton cmdEdi2 
         Caption         =   "&Editar"
         Enabled         =   0   'False
         Height          =   675
         Left            =   -72150
         Picture         =   "FrmOrdenes_Requerimientos.frx":2F2A
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   7110
         Visible         =   0   'False
         Width           =   775
      End
      Begin VB.CommandButton cmdEli2 
         Caption         =   "&Quitar"
         Enabled         =   0   'False
         Height          =   675
         Left            =   -70800
         Picture         =   "FrmOrdenes_Requerimientos.frx":336C
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   7110
         Visible         =   0   'False
         Width           =   775
      End
      Begin VB.CommandButton cmdNue2 
         Caption         =   "&Agregar"
         Height          =   675
         Left            =   -73725
         Picture         =   "FrmOrdenes_Requerimientos.frx":37AE
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   7110
         Visible         =   0   'False
         Width           =   775
      End
      Begin VB.Data Data2 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   -74880
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   4845
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Frame Frame1 
         Height          =   636
         Left            =   -74730
         TabIndex        =   2
         Top             =   360
         Width           =   9660
         Begin ctrlayuda_f.Ctr_Ayuda Ctrayu_tipoorden 
            Height          =   270
            Left            =   960
            TabIndex        =   3
            Top             =   240
            Width           =   3390
            _ExtentX        =   5980
            _ExtentY        =   476
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
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Estado  :"
            Height          =   192
            Left            =   7080
            TabIndex        =   8
            Top             =   288
            Width           =   636
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
            TabIndex        =   7
            Top             =   204
            Width           =   1644
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Número  :"
            Height          =   192
            Left            =   4656
            TabIndex        =   6
            Top             =   288
            Width           =   696
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
            TabIndex        =   5
            Top             =   192
            Width           =   1560
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            Caption         =   "Tipo Orden     :"
            Height          =   192
            Left            =   96
            TabIndex        =   4
            Top             =   276
            Width           =   1032
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid Flex1 
         Height          =   2475
         Left            =   -74760
         TabIndex        =   1
         Top             =   3525
         Visible         =   0   'False
         Width           =   9735
         _ExtentX        =   17171
         _ExtentY        =   4366
         _Version        =   393216
         Cols            =   15
         FixedCols       =   0
         RowHeightMin    =   240
         BackColorSel    =   0
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         Appearance      =   0
         FormatString    =   "^Código|Fab|Descripción|xUni|xCantidad|Uni.|Cantidad|PU|>Precio|>%Des|Igv|>Total|C1|C2"
         BandDisplay     =   1
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
      Begin Crystal.CrystalReport CrystalReport1 
         Bindings        =   "FrmOrdenes_Requerimientos.frx":3BF0
         Left            =   -74880
         Top             =   4365
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         PrintFileLinesPerPage=   60
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   6015
         Left            =   -74400
         TabIndex        =   53
         Top             =   600
         Width           =   9975
         _ExtentX        =   17595
         _ExtentY        =   10610
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
            DataField       =   "estadoocdescripcion"
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
               ColumnWidth     =   675.213
            EndProperty
            BeginProperty Column01 
               Alignment       =   2
               Locked          =   -1  'True
               WrapText        =   -1  'True
               ColumnWidth     =   1289.764
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
               ColumnWidth     =   1665.071
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   14.74
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "FrmOrdenes_Requerimientos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents adodc1 As ADODB.Recordset
Attribute adodc1.VB_VarHelpID = -1
Dim WithEvents rs As ADODB.Recordset
Attribute rs.VB_VarHelpID = -1
Dim Colex As New Collection
Public VGvardllgen As dllgeneral.dll_general
Public dllgeneral As New dllgeneral.dll_general
Public tabla1 As String
Dim cSql1 As String
Dim nT As Integer       'Ingreso,Modificación,Ficha Tecnica
Dim cCod As String
Dim nTra As Integer
Dim Mensaje As String
Dim tipodebienes As String
Dim unum As String


Sub OculObj02(nTipo As Boolean)
    cmdEdi2.Visible = nTipo
    cmdGra.Visible = nTipo
    CmdSalir2.Visible = nTipo
End Sub

Sub OculObj03(nTipo As Boolean)
    Fradatos.Visible = nTipo
    fraTotales.Visible = nTipo
End Sub

Sub OculObj04(nTipo As Boolean)
    cmdNue.Visible = nTipo
    cmdEdi.Visible = nTipo
    CmdEli.Visible = nTipo
    cmdImp.Visible = nTipo
    cmdsalir.Visible = nTipo
End Sub

Sub OculObj06(nTipo As Boolean)
    DataGrid1.Visible = nTipo
End Sub

Sub Abre_Tabla_OCs()
    Dim SQL As String
    Set VGvardllgen = New dllgeneral.dll_general
    Set adodc1 = New ADODB.Recordset
    
    SQL = "SELECT * FROM co_cabordcompra a inner join co_estadoorden b on "
    SQL = SQL & " a.estadooccodigo= b.estadooccodigo"
    SQL = SQL & " inner join co_tipodeorden c on a.tipoordencodigo=c.tipoordencodigo "
    SQL = SQL & " where b.estadoocatendido<>1 and c.flagrequerimientosordenes<>'1' "
    SQL = SQL & " and oc_estadoorden <> 1 ORDER BY oc_cnumord "
    adodc1.Open SQL, VGCNx, adOpenDynamic, adLockOptimistic
    
    Set DataGrid1.DataSource = adodc1
    
End Sub

Private Sub cmdEdi2_Click()
On Error GoTo Err
    With FrmOrdenes_requerimientosdetalle
        .activado = False
        .CtrAyu_articulo.xclave = Flex1.TextMatrix(Flex1.Row, 0)
        .lblFab = Flex1.TextMatrix(Flex1.Row, 1)
        .CtrAyu_articulo.xnombre = Flex1.TextMatrix(Flex1.Row, 2)
        .lblUni = Flex1.TextMatrix(Flex1.Row, 3)
        .txtCan = Flex1.TextMatrix(Flex1.Row, 4)
        .txtCan.Enabled = True
        .Checkigv.Value = 0
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
        .TxtOrdfab = Flex1.TextMatrix(Flex1.Row, 12)
        .Txtco1 = Flex1.TextMatrix(Flex1.Row, 13)
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
            Flex1.TextMatrix(Flex1.Row, 11) = Format(Flex1.TextMatrix(Flex1.Row, 6) * Flex1.TextMatrix(Flex1.Row, 8), "0.0000")
            Flex1.TextMatrix(Flex1.Row, 12) = .TxtOrdfab
            Flex1.TextMatrix(Flex1.Row, 13) = .Txtco1
            Calcula_Totales
        End If
        Flex1.SetFocus
    End With
 Exit Sub
Err:
    MsgBox Err.Description
    Exit Sub
    Resume
End Sub

Private Sub CmdEli_Click()
    On Error GoTo EliErr
    
    If adodc1("oc_estadoorden") = 1 Or ESNULO(adodc1("oc_situacionorden"), 0) <> "0" Then
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
        
        strsql = "UPDATE co_detordcompra SET oc_situacionorden=2  WHERE oc_cnumord='" & voc & "'"
        VGCNx.Execute strsql
        strsql = "UPDATE co_cabordcompra SET oc_estadoorden=1 WHERE oc_cnumord='" & voc & "'"
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
    Dim vURef As String, txtMon As String
    Dim txtEst As String, txtTip As Integer
    Dim txtPro As String, txtSol As String
    Dim LblPro As String, txtFor As String
    On Error GoTo GrabErr
    
    txtTip = 0
    txtFor = Trim(CtrAyu_pago.xclave)
    
    If Trim(Ctrayu_tipoorden.xclave) = "" Then
       Mensaje = "Debe ingresar Código de Tipo de Orden"
       MsgBox Mensaje, vbExclamation, "Mensaje"
       Ctrayu_tipoorden.SetFocus
       Exit Sub
    End If
    
    txtPro = Trim(CtrAyu_Proveedor.xclave)
    If txtPro = "" Then
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
       
    txtMon = CtrAyu_moneda.xclave
    If Trim(txtMon) = "" Then
        Mensaje = "Debe ingresar el Tipo de Moneda"
        MsgBox Mensaje, vbExclamation, "Error"
        CtrAyu_moneda.SetFocus
        Exit Sub
    End If
    
    txtEst = ""
    txtSol = Trim(CtrAyu_solicitante.xclave)
    If txtSol = "" Then
        Mensaje = "Debe ingresar Solicitante"
        MsgBox Mensaje, vbExclamation, "Mensaje"
        CtrAyu_solicitante.SetFocus
        Exit Sub
    End If
    
    If nT = 1 Then
        Mensaje = "¿Desea ingresar la nueva Orden de Compra?"
    Else
        Mensaje = "¿Desea guardar los cambios realizados?"
    End If
    
    If MsgBox(Mensaje, vbQuestion + vbYesNo, "Mensaje") = vbYes Then
 '      nTra = 1
       VGCNx.BeginTrans
       unum = Format(Val(lblNum), "00000000000")

       If nT = 1 Then      'Ingreso
         'unum = Format(Devolver_Dato(1, , " & trim(ctrayu_tipoordencodigo) & ", "tipoordencodigo", False,
         '      "ctnnumero"), "00000000000")
         SQLc = "select tipoordennumeracion from co_tipodeorden where tipoordencodigo='" & Trim(Ctrayu_tipoorden.xclave) & "' "
         Set rs2 = New ADODB.Recordset
         rs2.Open SQLc, VGCNx, adOpenKeyset, adLockReadOnly
         unum = rs2!tipoordennumeracion + 1
          
          SQLc = "UPDATE co_tipodeorden SET tipoordennumeracion=" & unum & _
                " WHERE tipoordencodigo='" & Trim(Ctrayu_tipoorden.xclave) & "' "
            VGCNx.Execute SQLc
           unum = Format(Val(unum), "00000000000")
           lblNum = unum
            SQLc = "INSERT INTO co_cabordcompra (tipoordencodigo,oc_cnumord,oc_dfecdoc,oc_ccodpro," & _
                "oc_crazsoc,oc_ccotiza,oc_ccodmon,oc_cforpag,oc_dfecent," & _
                "oc_cobserv,oc_csolict,oc_centreg,oc_estadoorden,estadooccodigo,oc_nimport,oc_ndescue," & _
                "oc_nigv,oc_nventa,oc_dfecact,oc_chora,oc_cusuari,oc_cconver) VALUES ('" & _
                Ctrayu_tipoorden.xclave & "','" & lblNum & "','" & txtEmi & "','" & txtPro & "','" & _
                CtrAyu_Proveedor.xnombre & "','" & txtCot & "','" & txtMon & "','" & txtFor & "','" & _
                txtEnt & "','" & _
                SupCadSQL(txtObs) & "','" & txtSol & "','" & txtEntE & "',' ','0'," & _
                CDbl(lblImp) & "," & CDbl(lblDes) & "," & CDbl(lblIgv) & "," & CDbl(lblCom) & _
                ",'" & txtEmi.Value & "','" & Format(Time, "hh.mm.ss") & "','" & VGUsuario & _
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
                  "oc_nprenet,oc_ntotven,oc_ntotnet,estadooccodigo,ord_fabnum,oc_ccomen1, tipoarticulocodigo, " & _
                  "oc_ncanten)" & _
                  "VALUES ('" & Ctrayu_tipoorden.xclave & "','" & lblNum & "','" & txtPro & "','" & txtEmi _
                  & "','" & Format(i, "000") & "','" & _
                  Flex1.TextMatrix(i, 0) & "','" & Flex1.TextMatrix(i, 1) & "','" & _
                  Left(Flex1.TextMatrix(i, 2), 65) & "','" & Flex1.TextMatrix(i, 3) & "','" & _
                  Flex1.TextMatrix(i, 5) & "'," & vFactor & "," & vCantid & "," & vCantid & "," & _
                  vPreuni & "," & vDscpor & "," & vDescto & "," & vIgv & "," & _
                  vIgvpor & "," & vPrenet & "," & vTotven & "," & vTotven - vDescto + _
                  vIgv & ",'0','" & Flex1.TextMatrix(i, 12) & "','" & _
                  Flex1.TextMatrix(i, 13) & "','" & Flex1.TextMatrix(i, 14) & "',0)"
                VGCNx.Execute SQLd
                
                SQLd = "UPDATE maeart SET aprecom=" & Val(Flex1.TextMatrix(i, 8)) & _
                    ",acodpro='" & txtPro & "',afecven='" & txtEmi _
                    & "' WHERE acodigo='" & Flex1.TextMatrix(i, 0) & "'"
                VGCNx.Execute SQLd
            Next
        ElseIf nT = 2 Then     'Modificar
            SQLc = "UPDATE co_cabordcompra SET oc_dfecdoc='" & txtEmi & _
                "',oc_ccodpro='" & txtPro & "',oc_crazsoc='" & Trim(CtrAyu_Proveedor.xnombre) & _
                "',oc_ccotiza='" & txtCot & "',oc_ccodmon='" & txtMon & "',oc_cforpag='" & _
                txtFor & "',oc_ntipcam=" & Val(txtTip) & ",oc_dfecent='" & _
                txtEnt & "',oc_cobserv='" & SupCadSQL(txtObs) & _
                "',oc_csolict='" & txtSol & "',oc_centreg='" & txtEntE & "',oc_nimport=" & _
                CDbl(lblImp) & ",oc_ndescue=" & CDbl(lblDes) & ",oc_nigv=" & CDbl(lblIgv) & _
                ",oc_nventa=" & CDbl(lblCom) & ",oc_dfecact='" & _
                txtEmi.Value & "',oc_chora='" & Format(Time, "hh.mm.ss") & "',oc_cusuari='" & _
                VGUsuario & "',oc_cconver='" & txtEst & "' WHERE tipoordencodigo='" & Ctrayu_tipoorden.xclave & "' and oc_cnumord='" & lblNum & "'"
            VGCNx.Execute SQLc
            
            SQLd = "DELETE co_detordcompra WHERE tipoordencodigo='" & Ctrayu_tipoorden.xclave & "' and oc_cnumord='" & lblNum & "'"
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
                SQLd = "INSERT INTO co_detordcompra (tipoordencodigo,oc_cnumord,oc_ccodpro,oc_dfecdoc,oc_citem," & _
                    "oc_ccodigo,oc_ccodref,oc_cdesref,oc_cunidad,oc_cuniref,oc_nfactor," & _
                    "oc_ncantid,oc_npreuni,oc_ndscpor,oc_ndescto,oc_nigv,oc_nigvpor," & _
                    "oc_nprenet,oc_ntotven,oc_ntotnet,estadooccodigo,ord_fabnum,oc_ccomen1,tipoarticulocodigo, " & _
                    "oc_ncanten,oc_nsaldo)" & _
                    "VALUES ('" & Ctrayu_tipoorden.xclave & "','" & lblNum & "','" & txtPro & "','" & txtEmi _
                    & "','" & Format(i, "000") & "','" & _
                    Flex1.TextMatrix(i, 0) & "','" & Flex1.TextMatrix(i, 1) & "','" & _
                    Flex1.TextMatrix(i, 2) & "','" & Flex1.TextMatrix(i, 3) & "','" & _
                    vURef & "'," & vFactor & "," & vCantid & "," & _
                    vPreuni & "," & vDscpor & "," & vDescto & "," & vIgv & "," & _
                    vIgvpor & "," & vPrenet & "," & vTotven & "," & vTotven - vDescto + _
                    vIgv & ",'0','" & Flex1.TextMatrix(i, 12) & "','" & _
                    Flex1.TextMatrix(i, 13) & "', '" & Flex1.TextMatrix(i, 14) & "',0,0)"
                VGCNx.Execute SQLd
                
                SQLd = "UPDATE maeart SET aprecom=" & Val(Flex1.TextMatrix(i, 8)) & _
                    ",acodpro='" & txtPro & "',afecven='" & txtEmi _
                    & "' WHERE acodigo='" & Flex1.TextMatrix(i, 0) & "'"
                VGCNx.Execute SQLd
            Next
        End If
        
        VGCNx.CommitTrans
        nTra = 0
        adodc1.Requery
        adodc1.Find "oc_cnumord='" & lblNum & "'"
        
        If nT = 1 Then
            unum = Format(Val(unum) + 1, "00000000000")
            lblNum = unum
            Limpiar
            Vacia_FlexGrid
            Estado_Items
            Calcula_Totales
            txtEmi = Date
            txtEnt = Date
            txtTip = "0.000"
                        
        Else
            CmdSalir2_Click
        End If
    
End If
Call CmdSalir2_Click
actualiza_requerimiento
Abre_Tabla_OCs
Exit Sub
GrabErr:
    MsgBox Err.Description
    If nTra = 1 Then VGCNx.RollbackTrans
    Exit Sub
    Resume
    
End Sub

Private Sub cmdImp_Click()
Dim formulas(3) As String
Dim tipoorden As String
unum = adodc1("oc_cnumord")
tipoorden = adodc1("tipoordencodigo")
CrystalReport1.Reset
CrystalReport1.WindowTitle = "al_rptordencompra.rpt -- orden de compra"
   CrystalReport1.ReportFileName = VGParamSistem.RutaReport & "al_rptordencompra" & Trim(VGCNx.DefaultDatabase) & ".rpt"
    CrystalReport1.DiscardSavedData = True
 
    CrystalReport1.Connect = VGCadenaReport2
    
    CrystalReport1.Destination = crptToWindow
    CrystalReport1.WindowState = crptMaximized
    CrystalReport1.WindowShowPrintBtn = True
    CrystalReport1.WindowShowRefreshBtn = True
    CrystalReport1.WindowShowSearchBtn = True
    CrystalReport1.WindowShowPrintSetupBtn = True
    Dim letras As String
    letras = NUMLET(adodc1("oc_nventa"))
    If adodc1("oc_ccodmon") = "01" Then
      letras = letras + " Nuevos Soles "
     Else
      letras = letras + " Dolares Americanos "
    End If
    CrystalReport1.formulas(0) = "@emp ='" & VGParametros.NomEmpresa & "'"
    CrystalReport1.formulas(1) = "@ruc ='" & VGParametros.RucEmpresa & "'"
    CrystalReport1.formulas(2) = "@letras ='" & letras & "'"
    CrystalReport1.StoredProcParam(0) = VGCNx.DefaultDatabase
    CrystalReport1.StoredProcParam(1) = tipoorden
   CrystalReport1.StoredProcParam(2) = unum
   If CrystalReport1.Status <> 2 Then
      CrystalReport1.Action = 1
   End If

End Sub

Private Sub cmdNue_Click()
Dim dllgeneral As New dllgeneral.dll_general
Call dllgeneral.ActivaTab(1, 1, SSTab1)
Set rs = New ADODB.Recordset
Set adodc1 = New ADODB.Recordset
TDBGrid1.FetchRowStyle = True
TDBGrid2.FetchRowStyle = True
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
        Frame1.Visible = True
        cmdEdi2.Enabled = True
        cmdEli2.Enabled = True
        cmdGra.Enabled = True
        
        txtEmi.SetFocus
        CmdSalir2.Cancel = True
    End If
End Sub

Private Sub cmdNue2_Click()
    With FrmOrdenes_requerimientosdetalle
        .activado = False
        .CtrAyu_articulo.xclave = ""
        .txtCan = "0.00"
        .txtPUn = "0.00"
        .txtPDe = "0.00"
        .txtPIg = "19.00"
        .TxtOrdfab = ""
        .lblFab.Caption = ""
        .Txtco1 = ""
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
                        (Val(.lblPNe) + Val(.lblDes)), "0.00") & vbTab & .TxtOrdfab & vbTab & _
                        .Txtco1 & vbTab & .tipo, 1
                    Flex1.Rows = 2
                Else
                    Flex1.AddItem Trim(.CtrAyu_articulo.xclave) & vbTab & .lblFab & vbTab & Trim(.CtrAyu_articulo.xnombre) & vbTab & _
                        .lblUni & vbTab & .txtCan & vbTab & IIf(.txtURe = "", .lblUni, _
                        .txtURe) & vbTab & IIf(.txtURe = "", .txtCan, .txtRef) & vbTab & _
                        .txtPUn & vbTab & Format(Val(.lblPNe) + Val(.lblDes), "0.00") & _
                        vbTab & .txtPDe & vbTab & .txtPIg & vbTab & Format(IIf(.txtURe = "", .txtCan, .txtRef) * _
                        (Val(.lblPNe) + Val(.lblDes)), "0.00") & vbTab & .TxtOrdfab & vbTab & _
                        .Txtco1 & vbTab & .tipo
                    Flex1.Row = Flex1.Rows - 1
                End If
            Else
                Flex1.AddItem Trim(.CtrAyu_articulo.xclave) & vbTab & .lblFab & vbTab & Trim(.CtrAyu_articulo.xnombre) & vbTab & _
                    .lblUni & vbTab & .txtCan & vbTab & IIf(.txtURe = "", .lblUni, _
                    .txtURe) & vbTab & IIf(.txtURe = "", .txtCan, .txtRef) & vbTab & _
                    .txtPUn & vbTab & Format(Val(.lblPNe) + Val(.lblDes), "0.00") & _
                    vbTab & .txtPDe & vbTab & .txtPIg & vbTab & Format(IIf(.txtURe = "", .txtCan, .txtRef) * _
                    (Val(.lblPNe) + Val(.lblDes)), "0.00") & vbTab & .TxtOrdfab & vbTab & _
                    .Txtco1 & vbTab & .tipo
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

Private Sub cmdOK_Click()
Dim SQL As String
Dim rstot As New ADODB.Recordset
Call dllgeneral.ActivaTab(2, 1, SSTab1)
TDBGrid1.FetchRowStyle = False
TDBGrid2.FetchRowStyle = False
rs.UpdateBatch adAffectAllChapters
adodc1.Update
adodc1.UpdateBatch adAffectAllChapters
SQL = " select oc_ccodigo,DESCRIPCION,oc_cunidad,oc_ncantid=sum(oc_ncantid) from " & tabla1 & " left join maeart "
SQL = SQL & " on oc_ccodigo=acodigo group by oc_ccodigo,DESCRIPCION,oc_cunidad "
Set rstot = VGCNx.Execute(SQL)
Call cargarcabecera(rs)
Call cargardetalle(rstot)
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
    cmdEdi2.Enabled = True

End Sub

Private Sub cmdsalir_Click()
    Unload frmReferencia
    Unload FrmOrdenes_requerimientosdetalle
    Unload Me
End Sub

Private Sub CmdSalir2_Click()
    Call dllgeneral.ActivaTab(0, 1, SSTab1)
    Limpiar
    Vacia_FlexGrid
    Estado_Items
    OculObj02 False
    OculObj03 False
    OculObj04 True
    OculObj06 True
    Proceso False
    Frame1.Visible = False
    If adodc1.RecordCount > 0 Then
        DataGrid1.SetFocus
    Else
        cmdNue.SetFocus
    End If
    cmdsalir.Cancel = True
End Sub
Public Function SupCadSQL(s As String) As String
 Dim Aux As String
 If Not IsNull(s) Then
     Aux = Replace(s, "'", "''")
 End If
 SupCadSQL = Aux
 
End Function

Private Sub Cmdsalirorden_Click()
Dim dllgeneral As New dllgeneral.dll_general
Call dllgeneral.ActivaTab(0, 1, SSTab1)
End Sub

Private Sub Command1_Click()
  Dim dllgeneral As New dllgeneral.dll_general
  Call dllgeneral.ActivaTab(0, 1, SSTab1)
End Sub

Private Sub Command2_Click()
Dim i As Integer
If rs.RecordCount() <= 0 Then Exit Sub
adodc1.AddNew
For i = 0 To rs.Fields.Count - 1
    adodc1.Collect(rs.Fields(i).Name) = rs.Collect(i)
Next
rs.Delete
TDBGrid2.Refresh
End Sub

Private Sub Ctr_AyutipoOrden1_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
Dim i As Integer
Dim Tabla As String
Tabla = "##relacionrequerimientos" + ComputerName
tabla1 = "##relacionrequerimientos1" + ComputerName
Set adodc1 = New ADODB.Recordset
Set rs = New ADODB.Recordset
If ExisteElem(0, VGCNx, "" & Tabla & "") Then VGCNx.Execute ("drop table " & Tabla)
If ExisteElem(0, VGCNx, "" & tabla1 & "") Then VGCNx.Execute ("drop table " & tabla1)
Set VGCommandoSP = New ADODB.Command
VGCommandoSP.ActiveConnection = VGGeneral
VGCommandoSP.CommandType = adCmdStoredProc
VGCommandoSP.CommandText = "al_relacionrequerimientosOrdenes_pro"
VGCommandoSP.Parameters.Refresh
With VGCommandoSP
    .Parameters("@base") = VGParamSistem.BDEmpresa
    .Parameters("@solicitante") = "%%"
    .Parameters("@orden") = Ctr_AyutipoOrden1.xclave
    .Parameters("@estado") = "3"
    .Parameters("@tipo") = 0
    .Parameters("@fechaini") = Date
    .Parameters("@fechafin") = Date
    .Parameters("@tabla") = Tabla
    Set adodc1 = .Execute
End With
adodc1.Open ("SELECT * FROM " & Tabla), VGCNx, adOpenDynamic, adLockBatchOptimistic
Set rs = New ADODB.Recordset
Set rs = VGCNx.Execute("SELECT top 0 * into " & tabla1 & " FROM " & Tabla)
rs.Open ("SELECT * FROM " & tabla1), VGCNx, adOpenDynamic, adLockBatchOptimistic
Set TDBGrid1.DataSource = adodc1
Set TDBGrid2.DataSource = rs
tipodebienes = ColecCampos("ordendebienes")
End Sub
Private Sub Ctrayu_tipoorden_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
    Dim unum As String
    Set VGvardllgen = New dllgeneral.dll_general
    unum = VGvardllgen.ESNULO(ColecCampos("tipoordennumeracion").Value, "")
    unum = Format(Val(unum) + 1, "00000000000")
    lblNum = unum
    
End Sub


Private Sub CtrAyu_Proveedor_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
    Set VGvardllgen = New dllgeneral.dll_general
    lblRuc.text = VGvardllgen.ESNULO(ColecCampos("clienteruc").Value, "")
End Sub
Private Sub CtrAyu_Proveedor_AlNoDevolverNada()
    lblRuc.text = ""
End Sub

Private Sub Form_Load()

    Formato_FlexGrid
    Call CtrAyu_moneda.Conexion(VGCNx): CtrAyu_moneda.filtro = "(monedacodigo <>'00') "
    Call Ctrayu_tipoorden.Conexion(VGCNx): Ctrayu_tipoorden.filtro = "(flagrequerimientosordenes <> 1 )"
    Call Ctr_AyutipoOrden1.Conexion(VGCNx): Ctr_AyutipoOrden1.filtro = "(flagrequerimientosordenes=1 )"
    Call CtrAyu_Proveedor.Conexion(VGCNx)
    Call CtrAyu_pago.Conexion(VGCNx)
    Call CtrAyu_solicitante.Conexion(VGCNx)
    
    Call dllgeneral.ActivaTab(0, 1, SSTab1)
    txtEmi.Value = Date
    txtEnt.Value = Date
    unum = ""
    Abre_Tabla_OCs
    Frame1.Visible = False
    Load FrmOrdenes_requerimientosdetalle
    TDBGrid1.FetchRowStyle = True
    TDBGrid2.FetchRowStyle = True
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
            cSel1.Open cSL, VGcnxCT, adOpenStatic
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
Ctrayu_tipoorden.xnombre = ""
CtrAyu_Proveedor.xnombre = ""
CtrAyu_pago.xnombre = ""
CtrAyu_solicitante.xnombre = ""
CtrAyu_moneda.xnombre = ""
txtEntE = "": txtObs = ""
End Sub

Sub Mostrar(cC1 As String)
    Dim cSqlM As String, cSelM As ADODB.Recordset
    Dim k As Integer, i As Integer, vd As String
    Dim vpu As Single, txtPro As String
    Dim txtSol As String
    
    lblNum = cC1
   ' lblEst = Adodc1("est_nombre")
    CtrAyu_Proveedor.xclave = adodc1("oc_ccodpro")
    txtPro = CtrAyu_Proveedor.xclave
    CtrAyu_Proveedor.xnombre = Devolver_Dato(1, txtPro, "cp_proveedor", "clientecodigo", False, "clienterazonsocial")
 '   lblRuc = Devolver_Dato(1, txtpro, "cp_proveedor", "clientecodigo", False, "clienteruc")
    txtEmi = adodc1("oc_dfecdoc")
    txtEnt = adodc1("oc_dfecent")
    CtrAyu_moneda.xclave = adodc1("oc_ccodmon")
    CtrAyu_pago.xclave = adodc1("oc_cforpag")
    txtCot = adodc1("oc_ccotiza")
    txtEntE = adodc1("oc_centreg")
    CtrAyu_solicitante.xclave = adodc1("oc_csolict")
    txtSol = CtrAyu_solicitante.xclave
    CtrAyu_solicitante.xnombre = Devolver_Dato(1, txtSol, "co_solicitantes", "solicitantecodigo", False, "solicitantenombre")
    txtObs = adodc1("oc_cobserv")
    Ctrayu_tipoorden.xclave = adodc1("tipoordencodigo")
    
    cSqlM = "SELECT * FROM co_detordcompra WHERE oc_cnumord='" & cC1 & "' ORDER BY oc_citem"
    Set cSelM = New ADODB.Recordset
    
    cSelM.Open cSqlM, VGCNx, adOpenStatic
    cSelM.MoveFirst
    
    k = 0
    Do While Not cSelM.EOF
        k = k + 1
        If k = 1 Then
            vpu = IIf(cSelM("oc_nfactor") = 0, 1, ESNULO(cSelM("oc_nfactor"), 0) / cSelM("oc_ncantid"))
            Flex1.AddItem cSelM("oc_ccodigo") & vbTab & cSelM("oc_ccodref") & vbTab & _
                cSelM("oc_cdesref") & vbTab & cSelM("oc_cunidad") & vbTab & _
                Format(cSelM("oc_ncantid"), "0.00") & vbTab & IIf(cSelM("oc_cuniref") = "", _
                cSelM("oc_cunidad"), cSelM("oc_cuniref")) & vbTab & _
                IIf(cSelM("oc_cuniref") = "", Format(cSelM("oc_ncantid"), "0.00"), _
                Format(cSelM("oc_nfactor"), "0.00")) & vbTab & Format(cSelM("oc_npreuni"), _
                "0.0000") & vbTab & Format(cSelM("oc_npreuni"), "0.0000") & vbTab & _
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
                "0.0000") & vbTab & Format(cSelM("oc_npreuni"), "0.0000") & vbTab & _
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



Private Sub TDBGrid1_DblClick()
Dim i As Integer
On Error GoTo Err
If adodc1.RecordCount() <= 0 Then Exit Sub
rs.AddNew
For i = 0 To adodc1.Fields.Count - 1
    rs.Collect(adodc1.Fields(i).Name) = adodc1.Collect(i)
Next
adodc1.Delete
adodc1.MoveNext
adodc1.Update
rs.Update
TDBGrid2.Refresh
TDBGrid1.Refresh
Err:
Exit Sub
End Sub

Private Sub TDBGrid1_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid70.StyleDisp)
Dim rsclone As New ADODB.Recordset
    On Error Resume Next
    Set rsclone = adodc1.Clone(adLockReadOnly)
    If rsclone.RecordCount = 0 Then Exit Sub
    rsclone.Bookmark = Bookmark
'    If rsclone!anno = Year(DTPFechaIni) And rsclone!mes = Month(DTPFechaIni) Then
'       RowStyle.BackColor = RGB(254, 251, 218)
'    End If
'    If rsclone!fechconcil > DateSerial(DTPFechaIni.Year, DTPFechaIni.Month, 1) Then
'       RowStyle.BackColor = RGB(200, 250, 100)
'    End If
'    flagcal = True
End Sub

Private Sub TDBGrid1_HeadClick(ByVal ColIndex As Integer)
 With adodc1
    If .Sort = Empty Then
        .Sort = TDBGrid1.Columns.item(ColIndex).DataField & " asc"
    ElseIf Right(.Sort, 3) = "asc" Then
        .Sort = TDBGrid1.Columns.item(ColIndex).DataField & " desc"
    ElseIf Right(.Sort, 4) = "desc" Then
        .Sort = TDBGrid1.Columns.item(ColIndex).DataField & " asc"
    End If
    TDBGrid1.Refresh
 End With
End Sub

Private Sub TDBGrid2_DblClick()
Dim i As Integer
On Error GoTo Err
If rs.RecordCount() <= 0 Then Exit Sub
adodc1.AddNew
For i = 0 To rs.Fields.Count - 1
    adodc1.Collect(rs.Fields(i).Name) = rs.Collect(i)
Next
rs.Delete
rs.MoveNext
adodc1.Update
rs.Update
TDBGrid2.Refresh
TDBGrid1.Refresh
Err:
Exit Sub
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
    cmdEdi2.Visible = Estado
    cmdEli2.Visible = Estado
    cmdGra.Visible = Estado
    CmdSalir2.Visible = Estado
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
    Dim tV As Single, Valor As Single
    Dim tD As Single, vDesc As Single
    Dim tI As Single, vIgv As Single
    
    With Flex1
        For i = 1 To Flex1.Rows - 1
            tV = Val(.TextMatrix(i, 11))
            Valor = Valor + tV
            tD = tV * Val(.TextMatrix(i, 9)) / 100
            vDesc = vDesc + tD
            tI = (tV - tD) * Val(.TextMatrix(i, 10)) / 100
            vIgv = vIgv + tI
        Next
    End With
    
    lblImp = Format(Valor, "##,##0.0000")
    lblDes = Format(vDesc, "##,##0.0000")
    lblTot = Format(Valor - vDesc, "#,##0.0000")
    lblIgv = Format(vIgv, "#,##0.00")
    lblCom = Format((Valor - vDesc) + vIgv, "#,##0.00")
End Sub

Function Tiene_Entregas() As Boolean
    Dim Adodc2 As ADODB.Recordset
    
    Set Adodc2 = New ADODB.Recordset
    
    Adodc2.Open "SELECT * FROM co_detordcompra WHERE oc_cnumord='" & lblNum & "' AND oc_ccodigo='" & _
        Flex1.TextMatrix(Flex1.Row, 0) & "' AND oc_ncanten>0", VGCNx, adOpenStatic
    Tiene_Entregas = False
    If Adodc2.RecordCount > 0 Then Tiene_Entregas = True
End Function

Sub cargarcabecera(rs1 As ADODB.Recordset)
Dim codpro As String
Dim solicitante As String
Dim i As Integer
Dim Y As Integer
Dim rs2 As New ADODB.Recordset
Set rs2 = VGCNx.Execute("select * from co_tipodeorden where flagrequerimientosOrdenes=0 and ordendebienes='" & tipodebienes & "'")
'Ctrayu_tipoorden.filtro = "": Ctrayu_tipoorden.Ejecutar
Ctrayu_tipoorden.xclave = rs2!tipoordencodigo: Ctrayu_tipoorden.Ejecutar
lblNum = Format(rs2!tipoordennumeracion, "00000000")
Ctrayu_tipoorden.filtro = "(flagrequerimientosordenes <> 1 )"
rs1.MoveFirst
i = 0
Y = 0
codpro = rs1!codpro
solicitante = rs1!solicitantecodigo
Do Until rs1.EOF()
   If codpro <> rs1!codpro Then i = 1
   If solicitante <> rs1!solicitantecodigo Then Y = 1
   rs1.MoveNext
Loop
If i = 0 Then
   CtrAyu_Proveedor.xclave = codpro: CtrAyu_Proveedor.Ejecutar
End If
If Y = 0 Then
   CtrAyu_solicitante.xclave = solicitante: CtrAyu_solicitante.Ejecutar
End If

End Sub

Sub cargardetalle(rs As ADODB.Recordset)
Dim k As Integer
rs.MoveFirst
k = 0
Do Until rs.EOF()
   k = k + 1
   Flex1.AddItem Trim(rs!oc_ccodigo) & vbTab & "" & vbTab & Trim(Escadena(rs!descripcion)) & vbTab & _
   rs!oc_cunidad & vbTab & rs!oc_ncantid & vbTab & rs!oc_cunidad & vbTab & rs!oc_cunidad & vbTab & _
   0# & vbTab & 0# & _
   vbTab & 0# & vbTab & VGParametros.Igv & vbTab & 0# & vbTab & "" & vbTab & _
   "" & vbTab & ""
   If k = 1 Then
     Flex1.Row = 1
     Flex1.RemoveItem 1
   End If
   rs.MoveNext
Loop
End Sub
Sub actualiza_requerimiento()
Dim clave As String
clave = Ctrayu_tipoorden.xclave + lblNum
Set rs = New ADODB.Recordset
Set rs = VGCNx.Execute("SELECT * FROM " & tabla1)
If rs.RecordCount() > 0 Then
   Do Until rs.EOF()
      SQL = " update co_detordcompra set ordenreferencia ='" & clave & "'"
      SQL = SQL & " where tipoordencodigo='" & rs!tipoordencodigo & "' and "
      SQL = SQL & " oc_cnumord='" & rs!oc_cnumord & "' and oc_citem='" & rs!oc_citem & "'"
      VGCNx.Execute (SQL)
      rs.MoveNext
   Loop
End If
End Sub
