VERSION 5.00
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Begin VB.Form FrmGenAsiento 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Generar Asientos a Contabilidad"
   ClientHeight    =   8805
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12975
   Icon            =   "FrmGenAsiento.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8805
   ScaleWidth      =   12975
   Begin VB.CheckBox Check1 
      BackColor       =   &H80000005&
      Caption         =   "Todas"
      ForeColor       =   &H00004080&
      Height          =   255
      Left            =   4680
      MaskColor       =   &H00004080&
      TabIndex        =   17
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Cmd_ElimConta 
      Caption         =   "&Eliminar de Contabilidad"
      Height          =   375
      Left            =   2385
      TabIndex        =   15
      Top             =   8370
      Width           =   2385
   End
   Begin VB.CommandButton Cmd_PasaConta 
      Caption         =   "&Traspasar a Contabilidad"
      Height          =   360
      Left            =   60
      TabIndex        =   13
      Top             =   8370
      Width           =   2235
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00808080&
      Height          =   3330
      Left            =   60
      ScaleHeight     =   3270
      ScaleWidth      =   12780
      TabIndex        =   2
      Top             =   4980
      Width           =   12840
      Begin TrueOleDBGrid70.TDBGrid TDBG_Siconta 
         Height          =   2715
         Left            =   120
         TabIndex        =   14
         Top             =   165
         Width           =   12585
         _ExtentX        =   22199
         _ExtentY        =   4789
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Nº Prov."
         Columns(0).DataField=   "cabprovinumero"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Mon"
         Columns(1).DataField=   "monedacodigo"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Razón Social"
         Columns(2).DataField=   "cabprovirznsoc"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Fech Doc."
         Columns(3).DataField=   "cabprovifchdoc"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "TD"
         Columns(4).DataField=   "documetocodigo"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "Nº Doc."
         Columns(5).DataField=   "cabprovinumdoc"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "Total Bruto"
         Columns(6).DataField=   "cabprovitotbru"
         Columns(6).NumberFormat=   "###,###,###.00"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "Total IGV"
         Columns(7).DataField=   "cabprovitotigv"
         Columns(7).NumberFormat=   "###,###,###.00"
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(8)._VlistStyle=   0
         Columns(8)._MaxComboItems=   5
         Columns(8).Caption=   "Total Inafecto"
         Columns(8).DataField=   "cabprovitotinaf"
         Columns(8).NumberFormat=   "###,###,###.00"
         Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(9)._VlistStyle=   0
         Columns(9)._MaxComboItems=   5
         Columns(9).Caption=   "Total Compra"
         Columns(9).DataField=   "cabprovitotal"
         Columns(9).NumberFormat=   "###,###,###.00"
         Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(10)._VlistStyle=   0
         Columns(10)._MaxComboItems=   5
         Columns(10).Caption=   "Nº C. Contable"
         Columns(10).DataField=   "cabprovinconta"
         Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   11
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   12632256
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=11"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=1376"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1296"
         Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(5)=   "Column(1).Width=609"
         Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=529"
         Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(9)=   "Column(2).Width=3704"
         Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=3625"
         Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(13)=   "Column(3).Width=1773"
         Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=1693"
         Splits(0)._ColumnProps(16)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(17)=   "Column(4).Width=529"
         Splits(0)._ColumnProps(18)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(19)=   "Column(4)._WidthInPix=450"
         Splits(0)._ColumnProps(20)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(21)=   "Column(5).Width=2302"
         Splits(0)._ColumnProps(22)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(23)=   "Column(5)._WidthInPix=2223"
         Splits(0)._ColumnProps(24)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(25)=   "Column(6).Width=2090"
         Splits(0)._ColumnProps(26)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(27)=   "Column(6)._WidthInPix=2011"
         Splits(0)._ColumnProps(28)=   "Column(6)._ColStyle=2"
         Splits(0)._ColumnProps(29)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(30)=   "Column(7).Width=1958"
         Splits(0)._ColumnProps(31)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(32)=   "Column(7)._WidthInPix=1879"
         Splits(0)._ColumnProps(33)=   "Column(7)._ColStyle=2"
         Splits(0)._ColumnProps(34)=   "Column(7).Order=8"
         Splits(0)._ColumnProps(35)=   "Column(8).Width=2064"
         Splits(0)._ColumnProps(36)=   "Column(8).DividerColor=0"
         Splits(0)._ColumnProps(37)=   "Column(8)._WidthInPix=1984"
         Splits(0)._ColumnProps(38)=   "Column(8)._ColStyle=2"
         Splits(0)._ColumnProps(39)=   "Column(8).Order=9"
         Splits(0)._ColumnProps(40)=   "Column(9).Width=2064"
         Splits(0)._ColumnProps(41)=   "Column(9).DividerColor=0"
         Splits(0)._ColumnProps(42)=   "Column(9)._WidthInPix=1984"
         Splits(0)._ColumnProps(43)=   "Column(9)._ColStyle=2"
         Splits(0)._ColumnProps(44)=   "Column(9).Order=10"
         Splits(0)._ColumnProps(45)=   "Column(10).Width=2725"
         Splits(0)._ColumnProps(46)=   "Column(10).DividerColor=0"
         Splits(0)._ColumnProps(47)=   "Column(10)._WidthInPix=2646"
         Splits(0)._ColumnProps(48)=   "Column(10)._ColStyle=1"
         Splits(0)._ColumnProps(49)=   "Column(10).Order=11"
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
         _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=62,.parent=13,.alignment=1"
         _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=59,.parent=14"
         _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=60,.parent=15"
         _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=61,.parent=17"
         _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=66,.parent=13,.alignment=1"
         _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=63,.parent=14"
         _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=64,.parent=15"
         _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=65,.parent=17"
         _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=70,.parent=13,.alignment=1"
         _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=67,.parent=14"
         _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=68,.parent=15"
         _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=69,.parent=17"
         _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=74,.parent=13,.alignment=1"
         _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=71,.parent=14"
         _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=72,.parent=15"
         _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=73,.parent=17"
         _StyleDefs(76)  =   "Splits(0).Columns(10).Style:id=78,.parent=13,.alignment=2,.bgcolor=&HFFFFB7&"
         _StyleDefs(77)  =   "Splits(0).Columns(10).HeadingStyle:id=75,.parent=14,.bgcolor=&HFFFF00&"
         _StyleDefs(78)  =   "Splits(0).Columns(10).FooterStyle:id=76,.parent=15"
         _StyleDefs(79)  =   "Splits(0).Columns(10).EditorStyle:id=77,.parent=17,.bgcolor=&HFFFF00&"
         _StyleDefs(80)  =   "Named:id=33:Normal"
         _StyleDefs(81)  =   ":id=33,.parent=0"
         _StyleDefs(82)  =   "Named:id=34:Heading"
         _StyleDefs(83)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(84)  =   ":id=34,.wraptext=-1"
         _StyleDefs(85)  =   "Named:id=35:Footing"
         _StyleDefs(86)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(87)  =   "Named:id=36:Selected"
         _StyleDefs(88)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(89)  =   "Named:id=37:Caption"
         _StyleDefs(90)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(91)  =   "Named:id=38:HighlightRow"
         _StyleDefs(92)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(93)  =   "Named:id=39:EvenRow"
         _StyleDefs(94)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(95)  =   "Named:id=40:OddRow"
         _StyleDefs(96)  =   ":id=40,.parent=33"
         _StyleDefs(97)  =   "Named:id=41:RecordSelector"
         _StyleDefs(98)  =   ":id=41,.parent=34"
         _StyleDefs(99)  =   "Named:id=42:FilterBar"
         _StyleDefs(100) =   ":id=42,.parent=33"
      End
      Begin VB.Label lbreg2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "0 "
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   11565
         TabIndex        =   12
         Top             =   2985
         Width           =   1140
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Nro. Reg :"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   10515
         TabIndex        =   11
         Top             =   3000
         Width           =   930
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00808080&
      Height          =   3330
      Left            =   60
      ScaleHeight     =   3270
      ScaleWidth      =   12795
      TabIndex        =   1
      Top             =   1230
      Width           =   12855
      Begin TrueOleDBGrid70.TDBGrid TDBG_Noconta 
         Height          =   2715
         Left            =   90
         TabIndex        =   8
         Top             =   135
         Width           =   12585
         _ExtentX        =   22199
         _ExtentY        =   4789
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Nº Provision"
         Columns(0).DataField=   "cabprovinumero"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Mon"
         Columns(1).DataField=   "monedacodigo"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Razón Social"
         Columns(2).DataField=   "cabprovirznsoc"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Fech Doc."
         Columns(3).DataField=   "cabprovifchdoc"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "TD"
         Columns(4).DataField=   "documetocodigo"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "Nº Doc."
         Columns(5).DataField=   "cabprovinumdoc"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "Total Bruto"
         Columns(6).DataField=   "cabprovitotbru"
         Columns(6).NumberFormat=   "###,###,###.00"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "Total IGV"
         Columns(7).DataField=   "cabprovitotigv"
         Columns(7).NumberFormat=   "###,###,###.00"
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(8)._VlistStyle=   0
         Columns(8)._MaxComboItems=   5
         Columns(8).Caption=   "Total Inafecto"
         Columns(8).DataField=   "cabprovitotinaf"
         Columns(8).NumberFormat=   "###,###,###.00"
         Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(9)._VlistStyle=   0
         Columns(9)._MaxComboItems=   5
         Columns(9).Caption=   "Total Compra"
         Columns(9).DataField=   "cabprovitotal"
         Columns(9).NumberFormat=   "###,###,###.00"
         Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(10)._VlistStyle=   0
         Columns(10)._MaxComboItems=   5
         Columns(10).Caption=   "Nº C. Contable"
         Columns(10).DataField=   "cabprovinconta"
         Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   11
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   12632256
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=11"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=1402"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1323"
         Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(5)=   "Column(1).Width=635"
         Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=556"
         Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(9)=   "Column(2).Width=3625"
         Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=3545"
         Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(13)=   "Column(3).Width=1773"
         Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=1693"
         Splits(0)._ColumnProps(16)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(17)=   "Column(4).Width=582"
         Splits(0)._ColumnProps(18)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(19)=   "Column(4)._WidthInPix=503"
         Splits(0)._ColumnProps(20)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(21)=   "Column(5).Width=2355"
         Splits(0)._ColumnProps(22)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(23)=   "Column(5)._WidthInPix=2275"
         Splits(0)._ColumnProps(24)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(25)=   "Column(6).Width=2064"
         Splits(0)._ColumnProps(26)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(27)=   "Column(6)._WidthInPix=1984"
         Splits(0)._ColumnProps(28)=   "Column(6)._ColStyle=2"
         Splits(0)._ColumnProps(29)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(30)=   "Column(7).Width=1879"
         Splits(0)._ColumnProps(31)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(32)=   "Column(7)._WidthInPix=1799"
         Splits(0)._ColumnProps(33)=   "Column(7)._ColStyle=2"
         Splits(0)._ColumnProps(34)=   "Column(7).Order=8"
         Splits(0)._ColumnProps(35)=   "Column(8).Width=2011"
         Splits(0)._ColumnProps(36)=   "Column(8).DividerColor=0"
         Splits(0)._ColumnProps(37)=   "Column(8)._WidthInPix=1931"
         Splits(0)._ColumnProps(38)=   "Column(8)._ColStyle=2"
         Splits(0)._ColumnProps(39)=   "Column(8).Order=9"
         Splits(0)._ColumnProps(40)=   "Column(9).Width=2170"
         Splits(0)._ColumnProps(41)=   "Column(9).DividerColor=0"
         Splits(0)._ColumnProps(42)=   "Column(9)._WidthInPix=2090"
         Splits(0)._ColumnProps(43)=   "Column(9)._ColStyle=2"
         Splits(0)._ColumnProps(44)=   "Column(9).Order=10"
         Splits(0)._ColumnProps(45)=   "Column(10).Width=2725"
         Splits(0)._ColumnProps(46)=   "Column(10).DividerColor=0"
         Splits(0)._ColumnProps(47)=   "Column(10)._WidthInPix=2646"
         Splits(0)._ColumnProps(48)=   "Column(10)._ColStyle=1"
         Splits(0)._ColumnProps(49)=   "Column(10).Order=11"
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
         _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=62,.parent=13,.alignment=1"
         _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=59,.parent=14"
         _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=60,.parent=15"
         _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=61,.parent=17"
         _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=66,.parent=13,.alignment=1"
         _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=63,.parent=14"
         _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=64,.parent=15"
         _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=65,.parent=17"
         _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=70,.parent=13,.alignment=1"
         _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=67,.parent=14"
         _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=68,.parent=15"
         _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=69,.parent=17"
         _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=74,.parent=13,.alignment=1"
         _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=71,.parent=14"
         _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=72,.parent=15"
         _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=73,.parent=17"
         _StyleDefs(76)  =   "Splits(0).Columns(10).Style:id=78,.parent=13,.alignment=2,.bgcolor=&HFFFFB7&"
         _StyleDefs(77)  =   "Splits(0).Columns(10).HeadingStyle:id=75,.parent=14,.bgcolor=&HFFFF00&"
         _StyleDefs(78)  =   "Splits(0).Columns(10).FooterStyle:id=76,.parent=15"
         _StyleDefs(79)  =   "Splits(0).Columns(10).EditorStyle:id=77,.parent=17"
         _StyleDefs(80)  =   "Named:id=33:Normal"
         _StyleDefs(81)  =   ":id=33,.parent=0"
         _StyleDefs(82)  =   "Named:id=34:Heading"
         _StyleDefs(83)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(84)  =   ":id=34,.wraptext=-1"
         _StyleDefs(85)  =   "Named:id=35:Footing"
         _StyleDefs(86)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(87)  =   "Named:id=36:Selected"
         _StyleDefs(88)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(89)  =   "Named:id=37:Caption"
         _StyleDefs(90)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(91)  =   "Named:id=38:HighlightRow"
         _StyleDefs(92)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(93)  =   "Named:id=39:EvenRow"
         _StyleDefs(94)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(95)  =   "Named:id=40:OddRow"
         _StyleDefs(96)  =   ":id=40,.parent=33"
         _StyleDefs(97)  =   "Named:id=41:RecordSelector"
         _StyleDefs(98)  =   ":id=41,.parent=34"
         _StyleDefs(99)  =   "Named:id=42:FilterBar"
         _StyleDefs(100) =   ":id=42,.parent=33"
      End
      Begin VB.Label lbreg1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "0 "
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   11550
         TabIndex        =   10
         Top             =   2985
         Width           =   1140
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Nro. Reg :"
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   10500
         TabIndex        =   9
         Top             =   2985
         Width           =   930
      End
   End
   Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuempresa 
      Height          =   375
      Left            =   5760
      TabIndex        =   0
      Top             =   240
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   661
      XcodMaxLongitud =   3
      xcodwith        =   300
      NomTabla        =   "co_multiempresas"
      TituloAyuda     =   "Busqueda de Empresas"
      ListaCampos     =   "empresacodigo(1),empresadescripcion(1),agentederetencion(1)"
      XcodCampo       =   "empresacodigo"
      XListCampo      =   "empresadescripcion"
      ListaCamposDescrip=   "Codigo,Descripcion"
      ListaCamposText =   "empresacodigo,empresadescripcion,agentederetencion"
   End
   Begin VB.Label Leempresa 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Empresa :"
      ForeColor       =   &H00004080&
      Height          =   315
      Left            =   4680
      TabIndex        =   16
      Top             =   300
      Width           =   705
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   225
      Picture         =   "FrmGenAsiento.frx":0E42
      Top             =   105
      Width           =   480
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Mes de "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   165
      Left            =   9855
      TabIndex        =   7
      Top             =   315
      Width           =   2730
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Asistente Para Generar Asientos a Contabilidad "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404080&
      Height          =   480
      Left            =   1020
      TabIndex        =   6
      Top             =   105
      Width           =   2520
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      Height          =   15
      Left            =   30
      Top             =   795
      Width           =   12900
   End
   Begin VB.Shape Shape1 
      Height          =   15
      Left            =   30
      Top             =   780
      Width           =   12900
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Height          =   750
      Left            =   30
      TabIndex        =   5
      Top             =   0
      Width           =   12900
   End
   Begin VB.Label Label2 
      BackColor       =   &H00DDF7F9&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Provisiones Contabilizadas en el Mes"
      Height          =   315
      Left            =   60
      TabIndex        =   4
      Top             =   4620
      Width           =   12825
   End
   Begin VB.Label Label1 
      BackColor       =   &H00DDF7F9&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Provisiones no Contabilizadas en el Mes"
      Height          =   315
      Left            =   45
      TabIndex        =   3
      Top             =   855
      Width           =   12855
   End
End
Attribute VB_Name = "FrmGenAsiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsNoConta As ADODB.Recordset
Dim RsSiConta As ADODB.Recordset

Private Sub Check1_Click()
If Check1.Value = 1 Then
   Ctr_Ayuempresa.Enabled = False
   Ctr_Ayuempresa.xclave = ""
End If
End Sub

Private Sub Cmd_ElimConta_Click()
Dim sqlcad As String
Dim sqlcad1 As String
Dim RSQL As New ADODB.Recordset
Screen.MousePointer = 11
If Ctr_Ayuempresa.xclave = "" Then
   MsgBox (" Digite  odigo de empresa")
   Ctr_Ayuempresa.SetFocus
   Exit Sub
End If
sqlcad = " select eqconta from co_tipocompra"
Set RSQL = VGCNx.Execute(sqlcad)
sqlcad1 = "("
Do Until RSQL.EOF
   sqlcad1 = sqlcad1 & "'" & RSQL!eqconta & "',"
   RSQL.MoveNext
Loop
sqlcad1 = Left(sqlcad1, Len(sqlcad1) - 1) + ")"

    sqlcad = " Update " & VGParamSistem.TablaCabcomprob & _
             " Set cabprovinconta='' " & _
             " Where cabproviano='" & VGParamSistem.Anoproceso & "' and cabprovimes=" & VGParamSistem.Mesproceso
    If Ctr_Ayuempresa.xclave <> "" Then
       sqlcad = sqlcad & " and empresacodigo='" & Ctr_Ayuempresa.xclave & "'"
    End If
    VGCNx.Execute sqlcad
    
    sqlcad = " Delete " & _
             " From ct_cabcomprob" & VGParamSistem.Anoproceso & _
             " Where cabcomprobmes=" & VGParamSistem.Mesproceso & " and (asientocodigo In " & sqlcad1 & ") and subasientocodigo='0099'"
             If Ctr_Ayuempresa.xclave <> "" Then
       sqlcad = sqlcad & " and empresacodigo='" & Ctr_Ayuempresa.xclave & "'"
    End If
    VGcnxCT.Execute sqlcad
    
    sqlcad = " Update " & VGParamSistem.TablaCabcomprob & _
             " Set cabproviflagmodi = 0 " & _
             " where cabproviano='" & VGParamSistem.Anoproceso & "' and cabprovimes = " & VGParamSistem.Mesproceso
    If Ctr_Ayuempresa.xclave <> "" Then
       sqlcad = sqlcad & " and empresacodigo='" & Ctr_Ayuempresa.xclave & "'"
    End If
    VGCNx.Execute sqlcad
    
    'Reinicializando el número correlativo para el comprobante del asiento
'    sqlcad = " Update ct_asientocorre "
'    If VGParamSistem.Mesproceso < 10 Then
'        sqlcad = sqlcad & " Set asientonumcorr0" & VGParamSistem.Mesproceso & " = 0 "
'    Else
'        sqlcad = sqlcad & " Set asientonumcorr" & VGParamSistem.Mesproceso & " = 0 "
'    End If
'    sqlcad = sqlcad & " Where (asientocodigo In " & sqlcad1 & ")"
'    If Ctr_Ayuempresa.xclave <> "" Then
'       sqlcad = sqlcad & " and empresacodigo='" & Ctr_Ayuempresa.xclave & "'"
'    End If
'    VGCNx.Execute sqlcad
    ' ---
    Call CargarDatos
 Screen.MousePointer = 1
End Sub
Private Sub Cmd_PasaConta_Click()
On Error GoTo genasiento

    Screen.MousePointer = 11
    VGCNx.BeginTrans
    
    'Generando los Analticos que no Esten en contabilidad
    Set VGCommandoSP = New ADODB.Command
    Set VGvardllgen = New dllgeneral.dll_general
    VGCommandoSP.ActiveConnection = VGGeneral
    VGCommandoSP.CommandType = adCmdStoredProc
    VGCommandoSP.CommandText = "co_VerificaAnalitico"
    VGCommandoSP.Parameters.Refresh
    With VGCommandoSP
        .Parameters("@BaseConta") = VGParamSistem.BDEmpresaCT
        .Parameters("@BaseCompra") = VGParamSistem.BDEmpresa
        .Parameters("@Ano") = VGParamSistem.Anoproceso
        .Parameters("@Mes") = VGParamSistem.Mesproceso
        .Parameters("@tipanal") = VGParametros.xTipAnal
        .Parameters("@User") = VGParamSistem.Usuario
        .Execute
    End With
RsNoConta.MoveFirst
If RsNoConta.RecordCount > 0 Then
   Do Until RsNoConta.EOF
      'Generando el Asiento en contabilidad
      Set VGCommandoSP = New ADODB.Command
      Set VGvardllgen = New dllgeneral.dll_general
    
      VGCommandoSP.ActiveConnection = VGGeneral
      VGCommandoSP.CommandType = adCmdStoredProc
      VGCommandoSP.CommandText = "co_generaasientoComprasenlinea_pro"
      VGCommandoSP.Parameters.Refresh
      With VGCommandoSP
          .Parameters("@BaseConta") = VGParamSistem.BDEmpresaCT
          .Parameters("@BaseCompra") = VGParamSistem.BDEmpresa
          .Parameters("@BaseCompra") = VGParamSistem.BDEmpresa
          .Parameters("@empresa") = RsNoConta!empresacodigo
          .Parameters("@SubAsiento") = VGParametros.xsubasiento
          .Parameters("@Libro") = VGParametros.xLibro
          .Parameters("@Mes") = Format(VGParamSistem.Mesproceso, "00")
          .Parameters("@Ano") = VGParamSistem.Anoproceso
          .Parameters("@ctatotal") = VGParametros.xCtaTotal
          .Parameters("@ctaIGV") = VGParametros.xCtaIGV
          .Parameters("@ctaIES") = VGParametros.xCtaIES
          .Parameters("@ctaRTA") = VGParametros.xCtaRTA
          .Parameters("@tipanal") = VGParametros.xTipAnal
          .Parameters("@Compu") = VGComputer
          .Parameters("@Usuario") = VGParamSistem.Usuario
          .Parameters("@Oficina") = Format(VGParametros.CpOficina, "00")
          .Parameters("@numcomprob") = RsNoConta!cabprovinumero
          If VGParamSistem.Anoproceso & Format(VGParamSistem.Mesproceso, "00") <= "200906" Then
            .Parameters("@tipo") = 0
           Else
           .Parameters("@tipo") = 1
          End If
          .Execute
    End With
    RsNoConta.MoveNext
 Loop
    
    'Actualizando las Glosas de Cabecera y Detalle
    
    Set VGCommandoSP = New ADODB.Command
    Set VGvardllgen = New dllgeneral.dll_general
    VGCommandoSP.ActiveConnection = VGGeneral
    VGCommandoSP.CommandType = adCmdStoredProc
    VGCommandoSP.CommandText = "co_GrabaGlosasProvision_pro"
    VGCommandoSP.Parameters.Refresh
    With VGCommandoSP
        .Parameters("@BaseConta") = VGParamSistem.BDEmpresaCT
        .Parameters("@BaseCompra") = VGParamSistem.BDEmpresa
        .Parameters("@Mes") = Format(VGParamSistem.Mesproceso, "00")
        .Parameters("@Ano") = VGParamSistem.Anoproceso
        .Execute
    End With
    
    VGCNx.CommitTrans
    
    MsgBox "Se Realizo la Operacion Satisfactoriamente"
    Call CargarDatos
    Screen.MousePointer = 1
End If
Exit Sub
genasiento:
    Screen.MousePointer = 1
    VGCNx.RollbackTrans
    MsgBox "Hubo Errores al momento que se genero los asientos " & Chr(13) & err.Description
End Sub

Private Sub Ctr_Ayuempresa_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
Call CargarDatos
Cmd_PasaConta.Enabled = True  ' Valida("Provisiones")
Cmd_ElimConta.Enabled = True  ' Valida("Provisiones")
End Sub

Private Sub Form_Load()
  Call Ctr_Ayuempresa.Conexion(VGCNx)
  If VGParametros.sistemamultiempresas = False Then
     Ctr_Ayuempresa.xclave = VGParametros.empresacodigo: Ctr_Ayuempresa.Ejecutar
     Ctr_Ayuempresa.Enabled = False
  End If
  Label5.Caption = DesMes(VGParamSistem.Mesproceso)
  Call CargarDatos
End Sub
Private Sub CargarDatos()
Dim SQL1 As String
Set RsNoConta = New ADODB.Recordset
Set RsSiConta = New ADODB.Recordset

SQL = "Select empresacodigo,cabprovinumero,monedacodigo,cabprovirznsoc, cabprovifchdoc,"
SQL = SQL & "documetocodigo, cabprovinumdoc,cabprovitotbru,cabprovitotigv,"
SQL = SQL & " cabprovitotinaf, cabprovitotal,cabprovinconta from " & VGParamSistem.TablaCabcomprob
SQL = SQL & " where cabproviano='" & VGParamSistem.Anoproceso & "' And cabprovimes = " & VGParamSistem.Mesproceso
If VGParametros.sistemamultiempresas Then
   If Ctr_Ayuempresa.xclave = "" Then
      SQL1 = SQL & " and ltrim(rtrim(isnull(cabprovinconta,'')))='' "
    Else
      SQL1 = SQL & " and empresacodigo='" & Ctr_Ayuempresa.xclave & "' and  ltrim(rtrim(isnull(cabprovinconta,'')))='' "
   End If
 Else
      SQL1 = SQL & " and ltrim(rtrim(isnull(cabprovinconta,'')))='' "
End If

Set RsNoConta = VGCNx.Execute(SQL1)
Set TDBG_Noconta.DataSource = RsNoConta
lbreg1.Caption = Format(RsNoConta.RecordCount, "0 ")
If VGParametros.sistemamultiempresas Then
   If Ctr_Ayuempresa.xclave = "" Then
      SQL1 = SQL & " and ltrim(rtrim(isnull(cabprovinconta,'')))<>''"
    Else
      SQL1 = SQL & " and empresacodigo='" & Ctr_Ayuempresa.xclave & "' and ltrim(rtrim(isnull(cabprovinconta,'')))<>''"
   End If
 Else
      SQL1 = SQL & " and ltrim(rtrim(isnull(cabprovinconta,'')))<>''"
End If
Set RsSiConta = VGCNx.Execute(SQL1)
Set TDBG_Siconta.DataSource = RsSiConta
lbreg2.Caption = Format(RsSiConta.RecordCount, "0 ")
TDBG_Noconta.FetchRowStyle = True

End Sub

Private Sub TDBG_Noconta_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid70.StyleDisp)
Dim rsclone As New ADODB.Recordset
On Error Resume Next
    Set rsclone = RsNoConta.Clone(adLockReadOnly)
    If rsclone.RecordCount = 0 Then Exit Sub
    rsclone.Bookmark = Bookmark
    If rsclone!cabproviflagmodi Then
       RowStyle.BackColor = RGB(249, 247, 221)
    End If
End Sub
