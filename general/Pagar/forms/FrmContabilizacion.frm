VERSION 5.00
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form FrmContabilizacion 
   Caption         =   "Contabilizacion"
   ClientHeight    =   8235
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12750
   LinkTopic       =   "Form1"
   ScaleHeight     =   8235
   ScaleWidth      =   12750
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Cmd_PasaConta 
      Caption         =   "&Traspasar a Contabilidad"
      Height          =   360
      Left            =   240
      TabIndex        =   10
      Top             =   7800
      Width           =   2235
   End
   Begin VB.CommandButton Cmd_ElimConta 
      Caption         =   "&Eliminar de Contabilidad"
      Height          =   375
      Left            =   2565
      TabIndex        =   9
      Top             =   7800
      Width           =   2385
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00808080&
      Height          =   3330
      Index           =   1
      Left            =   0
      ScaleHeight     =   3270
      ScaleWidth      =   12795
      TabIndex        =   6
      Top             =   4320
      Width           =   12855
      Begin TrueOleDBGrid70.TDBGrid TDBG_siconta 
         Height          =   2955
         Left            =   0
         TabIndex        =   11
         Top             =   0
         Width           =   12585
         _ExtentX        =   22199
         _ExtentY        =   5212
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Tipo"
         Columns(0).DataField=   "abonotipoplanilla"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Nro.Planilla"
         Columns(1).DataField=   "abononumplanilla"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Fecha"
         Columns(2).DataField=   "abonocanfecpla"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Cliente"
         Columns(3).DataField=   "Abonocancli"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Mon Doc"
         Columns(4).DataField=   "abonocanmoneda"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "T,Doc"
         Columns(5).DataField=   "documentoabono"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "Nro.Doc"
         Columns(6).DataField=   "abononumdoc"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "Importe"
         Columns(7).DataField=   "abonocanimcan"
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(8)._VlistStyle=   0
         Columns(8)._MaxComboItems=   5
         Columns(8).Caption=   "Mon Canc"
         Columns(8).DataField=   "abonocanmoncan"
         Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(9)._VlistStyle=   0
         Columns(9)._MaxComboItems=   5
         Columns(9).Caption=   "TDQC"
         Columns(9).DataField=   "abonocantdqc"
         Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(10)._VlistStyle=   0
         Columns(10)._MaxComboItems=   5
         Columns(10).Caption=   "Nro.DQC"
         Columns(10).DataField=   "abonocanndqc"
         Columns(10).DataWidth=   11
         Columns(10).NumberFormat=   "FormatText Event"
         Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(11)._VlistStyle=   0
         Columns(11)._MaxComboItems=   5
         Columns(11).Caption=   "Nº C. Contable"
         Columns(11).DataField=   "Comprobconta"
         Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   12
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   12632256
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=12"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=741"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=661"
         Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(5)=   "Column(1).Width=1482"
         Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=1402"
         Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(9)=   "Column(2).Width=1720"
         Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=1640"
         Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(13)=   "Column(3).Width=2223"
         Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=2143"
         Splits(0)._ColumnProps(16)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(17)=   "Column(4).Width=1349"
         Splits(0)._ColumnProps(18)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(19)=   "Column(4)._WidthInPix=1270"
         Splits(0)._ColumnProps(20)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(21)=   "Column(5).Width=1032"
         Splits(0)._ColumnProps(22)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(23)=   "Column(5)._WidthInPix=953"
         Splits(0)._ColumnProps(24)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(25)=   "Column(6).Width=2143"
         Splits(0)._ColumnProps(26)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(27)=   "Column(6)._WidthInPix=2064"
         Splits(0)._ColumnProps(28)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(29)=   "Column(7).Width=2170"
         Splits(0)._ColumnProps(30)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(31)=   "Column(7)._WidthInPix=2090"
         Splits(0)._ColumnProps(32)=   "Column(7).Order=8"
         Splits(0)._ColumnProps(33)=   "Column(8).Width=1482"
         Splits(0)._ColumnProps(34)=   "Column(8).DividerColor=0"
         Splits(0)._ColumnProps(35)=   "Column(8)._WidthInPix=1402"
         Splits(0)._ColumnProps(36)=   "Column(8).Order=9"
         Splits(0)._ColumnProps(37)=   "Column(9).Width=582"
         Splits(0)._ColumnProps(38)=   "Column(9).DividerColor=0"
         Splits(0)._ColumnProps(39)=   "Column(9)._WidthInPix=503"
         Splits(0)._ColumnProps(40)=   "Column(9).Order=10"
         Splits(0)._ColumnProps(41)=   "Column(10).Width=2778"
         Splits(0)._ColumnProps(42)=   "Column(10).DividerColor=0"
         Splits(0)._ColumnProps(43)=   "Column(10)._WidthInPix=2699"
         Splits(0)._ColumnProps(44)=   "Column(10)._ColStyle=2"
         Splits(0)._ColumnProps(45)=   "Column(10).Order=11"
         Splits(0)._ColumnProps(46)=   "Column(11).Width=2725"
         Splits(0)._ColumnProps(47)=   "Column(11).DividerColor=0"
         Splits(0)._ColumnProps(48)=   "Column(11)._WidthInPix=2646"
         Splits(0)._ColumnProps(49)=   "Column(11)._ColStyle=1"
         Splits(0)._ColumnProps(50)=   "Column(11).Order=12"
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
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=70,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=67,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=68,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=69,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=58,.parent=13"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=55,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=56,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=57,.parent=17"
         _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=32,.parent=13"
         _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=29,.parent=14"
         _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=30,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=31,.parent=17"
         _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=86,.parent=13"
         _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=83,.parent=14"
         _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=84,.parent=15"
         _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=85,.parent=17"
         _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=82,.parent=13"
         _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=79,.parent=14"
         _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=80,.parent=15"
         _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=81,.parent=17"
         _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=46,.parent=13"
         _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=43,.parent=14"
         _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=44,.parent=15"
         _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=45,.parent=17"
         _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=50,.parent=13"
         _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=47,.parent=14"
         _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=48,.parent=15"
         _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=49,.parent=17"
         _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=74,.parent=13"
         _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=71,.parent=14"
         _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=72,.parent=15"
         _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=73,.parent=17"
         _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=54,.parent=13"
         _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=51,.parent=14"
         _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=52,.parent=15"
         _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=53,.parent=17"
         _StyleDefs(76)  =   "Splits(0).Columns(10).Style:id=62,.parent=13,.alignment=1"
         _StyleDefs(77)  =   "Splits(0).Columns(10).HeadingStyle:id=59,.parent=14"
         _StyleDefs(78)  =   "Splits(0).Columns(10).FooterStyle:id=60,.parent=15"
         _StyleDefs(79)  =   "Splits(0).Columns(10).EditorStyle:id=61,.parent=17"
         _StyleDefs(80)  =   "Splits(0).Columns(11).Style:id=78,.parent=13,.alignment=2,.bgcolor=&HFFFFB7&"
         _StyleDefs(81)  =   "Splits(0).Columns(11).HeadingStyle:id=75,.parent=14,.bgcolor=&HFFFF00&"
         _StyleDefs(82)  =   "Splits(0).Columns(11).FooterStyle:id=76,.parent=15"
         _StyleDefs(83)  =   "Splits(0).Columns(11).EditorStyle:id=77,.parent=17"
         _StyleDefs(84)  =   "Named:id=33:Normal"
         _StyleDefs(85)  =   ":id=33,.parent=0"
         _StyleDefs(86)  =   "Named:id=34:Heading"
         _StyleDefs(87)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(88)  =   ":id=34,.wraptext=-1"
         _StyleDefs(89)  =   "Named:id=35:Footing"
         _StyleDefs(90)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(91)  =   "Named:id=36:Selected"
         _StyleDefs(92)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(93)  =   "Named:id=37:Caption"
         _StyleDefs(94)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(95)  =   "Named:id=38:HighlightRow"
         _StyleDefs(96)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(97)  =   "Named:id=39:EvenRow"
         _StyleDefs(98)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(99)  =   "Named:id=40:OddRow"
         _StyleDefs(100) =   ":id=40,.parent=33"
         _StyleDefs(101) =   "Named:id=41:RecordSelector"
         _StyleDefs(102) =   ":id=41,.parent=34"
         _StyleDefs(103) =   "Named:id=42:FilterBar"
         _StyleDefs(104) =   ":id=42,.parent=33"
      End
      Begin VB.Label lbreg2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "0 "
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   11550
         TabIndex        =   8
         Top             =   2985
         Width           =   1140
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Nro. Reg :"
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   10500
         TabIndex        =   7
         Top             =   2985
         Width           =   930
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00808080&
      Height          =   3570
      Index           =   0
      Left            =   0
      ScaleHeight     =   3510
      ScaleWidth      =   12795
      TabIndex        =   0
      Top             =   480
      Width           =   12855
      Begin TrueOleDBGrid70.TDBGrid TDBG_Noconta 
         Height          =   2955
         Left            =   90
         TabIndex        =   1
         Top             =   135
         Width           =   12585
         _ExtentX        =   22199
         _ExtentY        =   5212
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Tipo"
         Columns(0).DataField=   "abonotipoplanilla"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Nro.Planilla"
         Columns(1).DataField=   "abononumplanilla"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Fecha"
         Columns(2).DataField=   "abonocanfecpla"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Cliente"
         Columns(3).DataField=   "Abonocancli"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Mon Doc"
         Columns(4).DataField=   "abonocanmoneda"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "T,Doc"
         Columns(5).DataField=   "documentoabono"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "Nro.Doc"
         Columns(6).DataField=   "abononumdoc"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "Importe"
         Columns(7).DataField=   "abonocanimcan"
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(8)._VlistStyle=   0
         Columns(8)._MaxComboItems=   5
         Columns(8).Caption=   "Mon Canc"
         Columns(8).DataField=   "abonocanmoncan"
         Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(9)._VlistStyle=   0
         Columns(9)._MaxComboItems=   5
         Columns(9).Caption=   "TDQC"
         Columns(9).DataField=   "abonocantdqc"
         Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(10)._VlistStyle=   0
         Columns(10)._MaxComboItems=   5
         Columns(10).Caption=   "Nro.DQC"
         Columns(10).DataField=   "abonocanndqc"
         Columns(10).DataWidth=   11
         Columns(10).NumberFormat=   "FormatText Event"
         Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(11)._VlistStyle=   0
         Columns(11)._MaxComboItems=   5
         Columns(11).Caption=   "Nº C. Contable"
         Columns(11).DataField=   "Comprobconta"
         Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   12
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   12632256
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=12"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=741"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=661"
         Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(5)=   "Column(1).Width=1402"
         Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=1323"
         Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(9)=   "Column(2).Width=1482"
         Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=1402"
         Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(13)=   "Column(3).Width=2434"
         Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=2355"
         Splits(0)._ColumnProps(16)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(17)=   "Column(4).Width=1349"
         Splits(0)._ColumnProps(18)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(19)=   "Column(4)._WidthInPix=1270"
         Splits(0)._ColumnProps(20)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(21)=   "Column(5).Width=1032"
         Splits(0)._ColumnProps(22)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(23)=   "Column(5)._WidthInPix=953"
         Splits(0)._ColumnProps(24)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(25)=   "Column(6).Width=2143"
         Splits(0)._ColumnProps(26)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(27)=   "Column(6)._WidthInPix=2064"
         Splits(0)._ColumnProps(28)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(29)=   "Column(7).Width=2170"
         Splits(0)._ColumnProps(30)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(31)=   "Column(7)._WidthInPix=2090"
         Splits(0)._ColumnProps(32)=   "Column(7).Order=8"
         Splits(0)._ColumnProps(33)=   "Column(8).Width=1482"
         Splits(0)._ColumnProps(34)=   "Column(8).DividerColor=0"
         Splits(0)._ColumnProps(35)=   "Column(8)._WidthInPix=1402"
         Splits(0)._ColumnProps(36)=   "Column(8).Order=9"
         Splits(0)._ColumnProps(37)=   "Column(9).Width=582"
         Splits(0)._ColumnProps(38)=   "Column(9).DividerColor=0"
         Splits(0)._ColumnProps(39)=   "Column(9)._WidthInPix=503"
         Splits(0)._ColumnProps(40)=   "Column(9).Order=10"
         Splits(0)._ColumnProps(41)=   "Column(10).Width=2699"
         Splits(0)._ColumnProps(42)=   "Column(10).DividerColor=0"
         Splits(0)._ColumnProps(43)=   "Column(10)._WidthInPix=2619"
         Splits(0)._ColumnProps(44)=   "Column(10)._ColStyle=2"
         Splits(0)._ColumnProps(45)=   "Column(10).Order=11"
         Splits(0)._ColumnProps(46)=   "Column(11).Width=2725"
         Splits(0)._ColumnProps(47)=   "Column(11).DividerColor=0"
         Splits(0)._ColumnProps(48)=   "Column(11)._WidthInPix=2646"
         Splits(0)._ColumnProps(49)=   "Column(11)._ColStyle=1"
         Splits(0)._ColumnProps(50)=   "Column(11).Order=12"
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
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=70,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=67,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=68,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=69,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=58,.parent=13"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=55,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=56,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=57,.parent=17"
         _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=32,.parent=13"
         _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=29,.parent=14"
         _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=30,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=31,.parent=17"
         _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=86,.parent=13"
         _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=83,.parent=14"
         _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=84,.parent=15"
         _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=85,.parent=17"
         _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=82,.parent=13"
         _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=79,.parent=14"
         _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=80,.parent=15"
         _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=81,.parent=17"
         _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=46,.parent=13"
         _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=43,.parent=14"
         _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=44,.parent=15"
         _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=45,.parent=17"
         _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=50,.parent=13"
         _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=47,.parent=14"
         _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=48,.parent=15"
         _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=49,.parent=17"
         _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=74,.parent=13"
         _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=71,.parent=14"
         _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=72,.parent=15"
         _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=73,.parent=17"
         _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=54,.parent=13"
         _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=51,.parent=14"
         _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=52,.parent=15"
         _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=53,.parent=17"
         _StyleDefs(76)  =   "Splits(0).Columns(10).Style:id=62,.parent=13,.alignment=1"
         _StyleDefs(77)  =   "Splits(0).Columns(10).HeadingStyle:id=59,.parent=14"
         _StyleDefs(78)  =   "Splits(0).Columns(10).FooterStyle:id=60,.parent=15"
         _StyleDefs(79)  =   "Splits(0).Columns(10).EditorStyle:id=61,.parent=17"
         _StyleDefs(80)  =   "Splits(0).Columns(11).Style:id=78,.parent=13,.alignment=2,.bgcolor=&HFFFFB7&"
         _StyleDefs(81)  =   "Splits(0).Columns(11).HeadingStyle:id=75,.parent=14,.bgcolor=&HFFFF00&"
         _StyleDefs(82)  =   "Splits(0).Columns(11).FooterStyle:id=76,.parent=15"
         _StyleDefs(83)  =   "Splits(0).Columns(11).EditorStyle:id=77,.parent=17"
         _StyleDefs(84)  =   "Named:id=33:Normal"
         _StyleDefs(85)  =   ":id=33,.parent=0"
         _StyleDefs(86)  =   "Named:id=34:Heading"
         _StyleDefs(87)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(88)  =   ":id=34,.wraptext=-1"
         _StyleDefs(89)  =   "Named:id=35:Footing"
         _StyleDefs(90)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(91)  =   "Named:id=36:Selected"
         _StyleDefs(92)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(93)  =   "Named:id=37:Caption"
         _StyleDefs(94)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(95)  =   "Named:id=38:HighlightRow"
         _StyleDefs(96)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(97)  =   "Named:id=39:EvenRow"
         _StyleDefs(98)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(99)  =   "Named:id=40:OddRow"
         _StyleDefs(100) =   ":id=40,.parent=33"
         _StyleDefs(101) =   "Named:id=41:RecordSelector"
         _StyleDefs(102) =   ":id=41,.parent=34"
         _StyleDefs(103) =   "Named:id=42:FilterBar"
         _StyleDefs(104) =   ":id=42,.parent=33"
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Nro. Reg :"
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   10500
         TabIndex        =   3
         Top             =   3225
         Width           =   930
      End
      Begin VB.Label lbreg1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "0 "
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   11550
         TabIndex        =   2
         Top             =   3225
         Width           =   1140
      End
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   0
      Picture         =   "FrmContabilizacion.frx":0000
      Top             =   0
      Width           =   480
   End
   Begin VB.Label Label1 
      BackColor       =   &H00DDF7F9&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Provisiones no Contabilizadas en el Mes"
      Height          =   555
      Left            =   480
      TabIndex        =   5
      Top             =   0
      Width           =   12015
   End
   Begin VB.Label Label2 
      BackColor       =   &H00DDF7F9&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Provisiones Contabilizadas en el Mes"
      Height          =   315
      Left            =   0
      TabIndex        =   4
      Top             =   3990
      Width           =   12825
   End
End
Attribute VB_Name = "FrmContabilizacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsNoConta As ADODB.Recordset
Dim RsSiConta As ADODB.Recordset
Dim rsparimpo As ADODB.Recordset
Dim tabla As String

Private Sub Cmd_ElimConta_Click()
Screen.MousePointer = 11
SQL = " Update cp_abono Set comprobconta='' Where month(abonocanfecpla)=" & VGparamsistem.Mesproceso
SQL = SQL & "  and year(abonocanfecpla)=" & VGparamsistem.Anoproceso
VGcnx.Execute SQL
    
SQL = " Delete ct_cabcomprob" & VGparamsistem.Anoproceso & " Where cabcomprobmes=" & VGparamsistem.Mesproceso
SQL = SQL & " and asientocodigo like '08%'   and subasientocodigo='0099'"
VGcnxCT.Execute SQL
Call CargarDatos
Screen.MousePointer = 1
End Sub

Private Sub Cmd_PasaConta_Click()
Set RsNoConta = VGcnx.Execute("select distinct abonotipoplanilla,abononumplanilla from " & tabla)
RsNoConta.MoveFirst
Dim nn As Integer
Dim n As Integer
nn = 0

Do While Not RsNoConta.EOF
    n = 0
    nn = nn + 1
    Call graba(nn)
    RsNoConta.MoveNext
 Loop
 Screen.MousePointer = 1
 MsgBox "La Operacion se Realizo Satisfactoriamente", vbInformation, "Sistema de Tesoreria"
 Call CargarDatos
End Sub
Private Sub graba(nn As Integer)
Dim Comando As ADODB.Command
On Error GoTo genasiento
Screen.MousePointer = 11
Set rsparimpo = New ADODB.Recordset
'rsparimpo.Open "Select * From  ct_importartesoreria Where Left(Upper(tipooperacion),1) ='" & RsNoConta!cabrec_ingsal & "'", VGcnxCT, adOpenKeyset, adLockReadOnly
'If rsparimpo.RecordCount() = 0 Then
'   Screen.MousePointer = 1
'   MsgBox "Verifique el parametro del asiento de " & RsNoConta!cabrec_ingsal & " en ct_importartesoreria ", vbInformation, "Sistema de Tesoreria"
'   Exit Sub
'End If
VGcnx.BeginTrans
Set VGCommandoSP = New ADODB.Command
Set VGvardllgen = New dllgeneral.dll_general
    comprobconta = Format(VGparamsistem.Mesproceso, "00") + "080" + Format(nn, "00000")
    Set Comando = New ADODB.Command
    
    
    With Comando
         .CommandType = adCmdStoredProc
         .CommandText = "cp_GeneraasientoxPagarenLinea_pro"
         .CommandTimeout = 0
         .ActiveConnection = VGgeneral
         .Parameters.Refresh
         .Parameters("@BaseConta") = VGcnxCT.DefaultDatabase
         .Parameters("@BaseVenta") = VGcnx.DefaultDatabase
         .Parameters("@Asiento") = "080"
         .Parameters("@SubAsiento") = "0099"
         .Parameters("@Libro") = "01"
         .Parameters("@Mes") = Format(VGparamsistem.Mesproceso, "00")
         .Parameters("@Ano") = VGparamsistem.Anoproceso
         .Parameters("@tipanal") = "001"
         .Parameters("@ctaingrfinanc") = "771100"
         .Parameters("@ctaegresofinanc") = "671100"
         .Parameters("@Compu") = VGComputer
         .Parameters("@Usuario") = VGparamsistem.Usuario
         .Parameters("@op") = 1
         .Parameters("@tipo") = RsNoConta!abonotipoplanilla
         .Parameters("@numero") = RsNoConta!abononumplanilla
         .Parameters("@comprobconta") = comprobconta
         .Execute
     End With
If n = 0 Then VGcnx.CommitTrans
Exit Sub
genasiento:
    n = 1
    Screen.MousePointer = 1
    VGcnx.RollbackTrans
    MsgBox "Hubo Errores al momento que se genero el recibo " & RsNoConta!abononumplanilla & Chr(13) & Err.Description
    Resume Next
End Sub

Private Sub Form_Load()
    Call CargarDatos
End Sub
Private Sub CargarDatos()
    Set RsNoConta = New ADODB.Recordset
    Set RsSiConta = New ADODB.Recordset
    tabla = VGComputer & "_Si"
    If ExisteElem(0, VGcnx, tabla) Then Set RsNoConta = VGcnx.Execute("drop table " & tabla)
    SQL = "Select a.* into " & tabla & " from cp_abono a inner join cp_tipoplanilla b "
    SQL = SQL & " on a.abonotipoplanilla=b.tplanillacodigo Where month(abonocanfecpla)=" & CInt(VGparamsistem.Mesproceso)
    SQL = SQL & " and ltrim(rtrim(isnull(comprobConta,'')))='' and year(abonocanfecpla)=" & VGparamsistem.Anoproceso
    SQL = SQL & " and isnull(abonocanflreg,0)<>1 and ( b.tplanillacobranza=1 or tplanillacanjes=1"
    SQL = SQL & " or tplanillarenovar=1 ) order by abononumplanilla,abonocancli"
    Set RsNoConta = VGcnx.Execute(SQL)
    Set RsNoConta = New ADODB.Recordset
    SQL = " select * from " & VGComputer & "_Si"
    RsNoConta.Open (SQL), VGcnx, adOpenKeyset, adLockReadOnly
    Set TDBG_Noconta.DataSource = RsNoConta
    lbreg1.Caption = Format(RsNoConta.RecordCount, "0 ")
    
    SQL = "Select a.* from cp_abono a inner join cp_tipoplanilla b "
    SQL = SQL & " on a.abonotipoplanilla=b.tplanillacodigo Where month(abonocanfecpla)=" & CInt(VGparamsistem.Mesproceso)
    SQL = SQL & " and ltrim(rtrim(isnull(comprobConta,'')))<>'' and year(abonocanfecpla)=" & VGparamsistem.Anoproceso
    SQL = SQL & " and isnull(abonocanflreg,0)<>1 and ( b.tplanillacobranza=1 or tplanillacanjes=1"
    SQL = SQL & " or tplanillarenovar=1) order by abononumplanilla,abonocancli"
    RsSiConta.Open SQL, VGcnx, adOpenKeyset, adLockReadOnly
    Set TDBG_siconta.DataSource = RsSiConta
    lbreg2.Caption = Format(RsSiConta.RecordCount, "0 ")
    
    
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


Private Sub SSTab1_DblClick()

End Sub

