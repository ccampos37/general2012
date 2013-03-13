VERSION 5.00
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Begin VB.Form FrmContabilizaTesoreria 
   Caption         =   "Contabilizacion de Tesoreria"
   ClientHeight    =   10095
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13470
   LinkTopic       =   "Form1"
   ScaleHeight     =   10095
   ScaleWidth      =   13470
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   9855
      Left            =   120
      TabIndex        =   1
      Top             =   510
      Width           =   13125
      _ExtentX        =   23151
      _ExtentY        =   17383
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "Contabilizacion Ingresos / egresos"
      TabPicture(0)   =   "FrmContabilizaTesoreria.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(3)=   "Shape1"
      Tab(0).Control(4)=   "Shape2"
      Tab(0).Control(5)=   "Label5"
      Tab(0).Control(6)=   "Image1"
      Tab(0).Control(7)=   "Picture1"
      Tab(0).Control(8)=   "Picture2"
      Tab(0).Control(9)=   "Cmd_PasaConta"
      Tab(0).Control(10)=   "Cmd_ElimConta"
      Tab(0).ControlCount=   11
      TabCaption(1)   =   "ContabilizacionTransferencias"
      TabPicture(1)   =   "FrmContabilizaTesoreria.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Image2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label11"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label12"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Picture3"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Picture4"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Cmd_ElimContaTransf"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Cmd_PasaContaTransf"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).ControlCount=   7
      Begin VB.CommandButton Cmd_PasaContaTransf 
         Caption         =   "&Traspasar a Contabilidad"
         Height          =   360
         Left            =   150
         TabIndex        =   25
         Top             =   8520
         Width           =   2235
      End
      Begin VB.CommandButton Cmd_ElimContaTransf 
         Caption         =   "&Eliminar de Contabilidad"
         Height          =   375
         Left            =   2475
         TabIndex        =   24
         Top             =   8490
         Width           =   2385
      End
      Begin VB.PictureBox Picture4 
         BackColor       =   &H00808080&
         Height          =   3330
         Left            =   120
         ScaleHeight     =   3270
         ScaleWidth      =   12795
         TabIndex        =   20
         Top             =   1320
         Width           =   12855
         Begin TrueOleDBGrid70.TDBGrid TDBG_NoContaTransf 
            Height          =   2715
            Left            =   90
            TabIndex        =   21
            Top             =   135
            Width           =   12585
            _ExtentX        =   22199
            _ExtentY        =   4789
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Nº Recibo"
            Columns(0).DataField=   "cabrec_numrecibo"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "I/E"
            Columns(1).DataField=   "cabrec_ingsal"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Nro Transferencia"
            Columns(2).DataField=   "cabrec_numreciboegreso"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Cod. Caja"
            Columns(3).DataField=   "cajacodigo"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Razón Social"
            Columns(4).DataField=   "clientecodigo"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Fech Doc."
            Columns(5).DataField=   "cabrec_fechadocumento"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "T.OP"
            Columns(6).DataField=   "operacioncodigo"
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(7)._VlistStyle=   0
            Columns(7)._MaxComboItems=   5
            Columns(7).Caption=   "Moneda"
            Columns(7).DataField=   "monedacodigo"
            Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(8)._VlistStyle=   0
            Columns(8)._MaxComboItems=   5
            Columns(8).Caption=   "Soles"
            Columns(8).DataField=   "cabrec_totsoles"
            Columns(8).NumberFormat=   "###,###,###.00"
            Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(9)._VlistStyle=   0
            Columns(9)._MaxComboItems=   5
            Columns(9).Caption=   "Dolares"
            Columns(9).DataField=   "cabrec_totdolares"
            Columns(9).NumberFormat=   "###,###,###.00"
            Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(10)._VlistStyle=   0
            Columns(10)._MaxComboItems=   5
            Columns(10).Caption=   "Nº C. Contable"
            Columns(10).DataField=   "Comprobconta"
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
            Splits(0)._ColumnProps(5)=   "Column(1).Width=900"
            Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=820"
            Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(9)=   "Column(2).Width=2487"
            Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=2408"
            Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(13)=   "Column(3).Width=1693"
            Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=1614"
            Splits(0)._ColumnProps(16)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(17)=   "Column(4).Width=3625"
            Splits(0)._ColumnProps(18)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(19)=   "Column(4)._WidthInPix=3545"
            Splits(0)._ColumnProps(20)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(21)=   "Column(5).Width=1773"
            Splits(0)._ColumnProps(22)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(23)=   "Column(5)._WidthInPix=1693"
            Splits(0)._ColumnProps(24)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(25)=   "Column(6).Width=582"
            Splits(0)._ColumnProps(26)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(27)=   "Column(6)._WidthInPix=503"
            Splits(0)._ColumnProps(28)=   "Column(6).Order=7"
            Splits(0)._ColumnProps(29)=   "Column(7).Width=2355"
            Splits(0)._ColumnProps(30)=   "Column(7).DividerColor=0"
            Splits(0)._ColumnProps(31)=   "Column(7)._WidthInPix=2275"
            Splits(0)._ColumnProps(32)=   "Column(7).Order=8"
            Splits(0)._ColumnProps(33)=   "Column(8).Width=2064"
            Splits(0)._ColumnProps(34)=   "Column(8).DividerColor=0"
            Splits(0)._ColumnProps(35)=   "Column(8)._WidthInPix=1984"
            Splits(0)._ColumnProps(36)=   "Column(8)._ColStyle=2"
            Splits(0)._ColumnProps(37)=   "Column(8).Order=9"
            Splits(0)._ColumnProps(38)=   "Column(9).Width=1879"
            Splits(0)._ColumnProps(39)=   "Column(9).DividerColor=0"
            Splits(0)._ColumnProps(40)=   "Column(9)._WidthInPix=1799"
            Splits(0)._ColumnProps(41)=   "Column(9)._ColStyle=2"
            Splits(0)._ColumnProps(42)=   "Column(9).Order=10"
            Splits(0)._ColumnProps(43)=   "Column(10).Width=2725"
            Splits(0)._ColumnProps(44)=   "Column(10).DividerColor=0"
            Splits(0)._ColumnProps(45)=   "Column(10)._WidthInPix=2646"
            Splits(0)._ColumnProps(46)=   "Column(10)._ColStyle=1"
            Splits(0)._ColumnProps(47)=   "Column(10).Order=11"
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
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=74,.parent=13"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=71,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=72,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=73,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=32,.parent=13"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=29,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=30,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=31,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=46,.parent=13"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=43,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=44,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=45,.parent=17"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=50,.parent=13"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=47,.parent=14"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=48,.parent=15"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=49,.parent=17"
            _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=54,.parent=13"
            _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=51,.parent=14"
            _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=52,.parent=15"
            _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=53,.parent=17"
            _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=58,.parent=13"
            _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=55,.parent=14"
            _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=56,.parent=15"
            _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=57,.parent=17"
            _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=62,.parent=13,.alignment=1"
            _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=59,.parent=14"
            _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=60,.parent=15"
            _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=61,.parent=17"
            _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=66,.parent=13,.alignment=1"
            _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=63,.parent=14"
            _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=64,.parent=15"
            _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=65,.parent=17"
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
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Nro. Reg :"
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Left            =   10500
            TabIndex        =   23
            Top             =   2985
            Width           =   930
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "0 "
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   11550
            TabIndex        =   22
            Top             =   2985
            Width           =   1140
         End
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00808080&
         Height          =   3330
         Left            =   120
         ScaleHeight     =   3270
         ScaleWidth      =   12780
         TabIndex        =   16
         Top             =   5070
         Width           =   12840
         Begin TrueOleDBGrid70.TDBGrid TDBG_SiContaTransf 
            Height          =   2715
            Left            =   120
            TabIndex        =   17
            Top             =   120
            Width           =   12585
            _ExtentX        =   22199
            _ExtentY        =   4789
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Nº Recibo"
            Columns(0).DataField=   "cabrec_numrecibo"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "I/E"
            Columns(1).DataField=   "cabrec_ingsal"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Nro.Transferencia"
            Columns(2).DataField=   "cabrec_numreciboegreso"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Cod. Caja"
            Columns(3).DataField=   "cajacodigo"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Razón Social"
            Columns(4).DataField=   "clientecodigo"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Fech Doc."
            Columns(5).DataField=   "cabrec_fechadocumento"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "T.OP"
            Columns(6).DataField=   "operacioncodigo"
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(7)._VlistStyle=   0
            Columns(7)._MaxComboItems=   5
            Columns(7).Caption=   "Moneda"
            Columns(7).DataField=   "monedacodigo"
            Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(8)._VlistStyle=   0
            Columns(8)._MaxComboItems=   5
            Columns(8).Caption=   "Soles"
            Columns(8).DataField=   "cabrec_totsoles"
            Columns(8).NumberFormat=   "###,###,###.00"
            Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(9)._VlistStyle=   0
            Columns(9)._MaxComboItems=   5
            Columns(9).Caption=   "Dolares"
            Columns(9).DataField=   "cabrec_totdolares"
            Columns(9).NumberFormat=   "###,###,###.00"
            Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(10)._VlistStyle=   0
            Columns(10)._MaxComboItems=   5
            Columns(10).Caption=   "Nº C. Contable"
            Columns(10).DataField=   "Comprobconta"
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
            Splits(0)._ColumnProps(5)=   "Column(1).Width=900"
            Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=820"
            Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(9)=   "Column(2).Width=2461"
            Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=2381"
            Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(13)=   "Column(3).Width=1693"
            Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=1614"
            Splits(0)._ColumnProps(16)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(17)=   "Column(4).Width=3625"
            Splits(0)._ColumnProps(18)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(19)=   "Column(4)._WidthInPix=3545"
            Splits(0)._ColumnProps(20)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(21)=   "Column(5).Width=1773"
            Splits(0)._ColumnProps(22)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(23)=   "Column(5)._WidthInPix=1693"
            Splits(0)._ColumnProps(24)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(25)=   "Column(6).Width=582"
            Splits(0)._ColumnProps(26)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(27)=   "Column(6)._WidthInPix=503"
            Splits(0)._ColumnProps(28)=   "Column(6).Order=7"
            Splits(0)._ColumnProps(29)=   "Column(7).Width=2355"
            Splits(0)._ColumnProps(30)=   "Column(7).DividerColor=0"
            Splits(0)._ColumnProps(31)=   "Column(7)._WidthInPix=2275"
            Splits(0)._ColumnProps(32)=   "Column(7).Order=8"
            Splits(0)._ColumnProps(33)=   "Column(8).Width=2064"
            Splits(0)._ColumnProps(34)=   "Column(8).DividerColor=0"
            Splits(0)._ColumnProps(35)=   "Column(8)._WidthInPix=1984"
            Splits(0)._ColumnProps(36)=   "Column(8)._ColStyle=2"
            Splits(0)._ColumnProps(37)=   "Column(8).Order=9"
            Splits(0)._ColumnProps(38)=   "Column(9).Width=1879"
            Splits(0)._ColumnProps(39)=   "Column(9).DividerColor=0"
            Splits(0)._ColumnProps(40)=   "Column(9)._WidthInPix=1799"
            Splits(0)._ColumnProps(41)=   "Column(9)._ColStyle=2"
            Splits(0)._ColumnProps(42)=   "Column(9).Order=10"
            Splits(0)._ColumnProps(43)=   "Column(10).Width=2725"
            Splits(0)._ColumnProps(44)=   "Column(10).DividerColor=0"
            Splits(0)._ColumnProps(45)=   "Column(10)._WidthInPix=2646"
            Splits(0)._ColumnProps(46)=   "Column(10)._ColStyle=1"
            Splits(0)._ColumnProps(47)=   "Column(10).Order=11"
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
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=74,.parent=13"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=71,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=72,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=73,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=32,.parent=13"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=29,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=30,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=31,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=46,.parent=13"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=43,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=44,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=45,.parent=17"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=50,.parent=13"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=47,.parent=14"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=48,.parent=15"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=49,.parent=17"
            _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=54,.parent=13"
            _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=51,.parent=14"
            _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=52,.parent=15"
            _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=53,.parent=17"
            _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=58,.parent=13"
            _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=55,.parent=14"
            _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=56,.parent=15"
            _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=57,.parent=17"
            _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=62,.parent=13,.alignment=1"
            _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=59,.parent=14"
            _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=60,.parent=15"
            _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=61,.parent=17"
            _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=66,.parent=13,.alignment=1"
            _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=63,.parent=14"
            _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=64,.parent=15"
            _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=65,.parent=17"
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
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Nro. Reg :"
            ForeColor       =   &H00FFFFFF&
            Height          =   225
            Left            =   10515
            TabIndex        =   19
            Top             =   3000
            Width           =   930
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "0 "
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   11565
            TabIndex        =   18
            Top             =   2985
            Width           =   1140
         End
      End
      Begin VB.CommandButton Cmd_ElimConta 
         Caption         =   "&Eliminar de Contabilidad"
         Height          =   375
         Left            =   -72525
         TabIndex        =   10
         Top             =   8610
         Width           =   2385
      End
      Begin VB.CommandButton Cmd_PasaConta 
         Caption         =   "&Traspasar a Contabilidad"
         Height          =   360
         Left            =   -74880
         TabIndex        =   9
         Top             =   8640
         Width           =   2235
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00808080&
         Height          =   3330
         Left            =   -74850
         ScaleHeight     =   3270
         ScaleWidth      =   12780
         TabIndex        =   6
         Top             =   5220
         Width           =   12840
         Begin TrueOleDBGrid70.TDBGrid TDBG_siConta 
            Height          =   2715
            Left            =   120
            TabIndex        =   15
            Top             =   120
            Width           =   12585
            _ExtentX        =   22199
            _ExtentY        =   4789
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Nº Recibo"
            Columns(0).DataField=   "cabrec_numrecibo"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "I/E"
            Columns(1).DataField=   "cabrec_ingsal"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Cod. Caja"
            Columns(2).DataField=   "cajacodigo"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Razón Social"
            Columns(3).DataField=   "clientecodigo"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Fech Doc."
            Columns(4).DataField=   "cabrec_fechadocumento"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "T.OP"
            Columns(5).DataField=   "operacioncodigo"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "Moneda"
            Columns(6).DataField=   "monedacodigo"
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(7)._VlistStyle=   0
            Columns(7)._MaxComboItems=   5
            Columns(7).Caption=   "Soles"
            Columns(7).DataField=   "cabrec_totsoles"
            Columns(7).NumberFormat=   "###,###,###.00"
            Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(8)._VlistStyle=   0
            Columns(8)._MaxComboItems=   5
            Columns(8).Caption=   "Dolares"
            Columns(8).DataField=   "cabrec_totdolares"
            Columns(8).NumberFormat=   "###,###,###.00"
            Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(9)._VlistStyle=   0
            Columns(9)._MaxComboItems=   5
            Columns(9).Caption=   "Nº C. Contable"
            Columns(9).DataField=   "Comprobconta"
            Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   10
            Splits(0)._UserFlags=   0
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=10"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=1402"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1323"
            Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(5)=   "Column(1).Width=900"
            Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=820"
            Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(9)=   "Column(2).Width=1693"
            Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=1614"
            Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(13)=   "Column(3).Width=3625"
            Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=3545"
            Splits(0)._ColumnProps(16)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(17)=   "Column(4).Width=1773"
            Splits(0)._ColumnProps(18)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(19)=   "Column(4)._WidthInPix=1693"
            Splits(0)._ColumnProps(20)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(21)=   "Column(5).Width=582"
            Splits(0)._ColumnProps(22)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(23)=   "Column(5)._WidthInPix=503"
            Splits(0)._ColumnProps(24)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(25)=   "Column(6).Width=2355"
            Splits(0)._ColumnProps(26)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(27)=   "Column(6)._WidthInPix=2275"
            Splits(0)._ColumnProps(28)=   "Column(6).Order=7"
            Splits(0)._ColumnProps(29)=   "Column(7).Width=2064"
            Splits(0)._ColumnProps(30)=   "Column(7).DividerColor=0"
            Splits(0)._ColumnProps(31)=   "Column(7)._WidthInPix=1984"
            Splits(0)._ColumnProps(32)=   "Column(7)._ColStyle=2"
            Splits(0)._ColumnProps(33)=   "Column(7).Order=8"
            Splits(0)._ColumnProps(34)=   "Column(8).Width=1879"
            Splits(0)._ColumnProps(35)=   "Column(8).DividerColor=0"
            Splits(0)._ColumnProps(36)=   "Column(8)._WidthInPix=1799"
            Splits(0)._ColumnProps(37)=   "Column(8)._ColStyle=2"
            Splits(0)._ColumnProps(38)=   "Column(8).Order=9"
            Splits(0)._ColumnProps(39)=   "Column(9).Width=2725"
            Splits(0)._ColumnProps(40)=   "Column(9).DividerColor=0"
            Splits(0)._ColumnProps(41)=   "Column(9)._WidthInPix=2646"
            Splits(0)._ColumnProps(42)=   "Column(9)._ColStyle=1"
            Splits(0)._ColumnProps(43)=   "Column(9).Order=10"
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
            _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=62,.parent=13,.alignment=1"
            _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=59,.parent=14"
            _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=60,.parent=15"
            _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=61,.parent=17"
            _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=66,.parent=13,.alignment=1"
            _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=63,.parent=14"
            _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=64,.parent=15"
            _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=65,.parent=17"
            _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=78,.parent=13,.alignment=2,.bgcolor=&HFFFFB7&"
            _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=75,.parent=14,.bgcolor=&HFFFF00&"
            _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=76,.parent=15"
            _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=77,.parent=17"
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
         Begin VB.Label lbreg2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "0 "
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   11565
            TabIndex        =   8
            Top             =   2985
            Width           =   1140
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Nro. Reg :"
            ForeColor       =   &H00FFFFFF&
            Height          =   225
            Left            =   10515
            TabIndex        =   7
            Top             =   3000
            Width           =   930
         End
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00808080&
         Height          =   3330
         Left            =   -74850
         ScaleHeight     =   3270
         ScaleWidth      =   12795
         TabIndex        =   2
         Top             =   1470
         Width           =   12855
         Begin TrueOleDBGrid70.TDBGrid TDBG_Noconta 
            Height          =   2715
            Left            =   90
            TabIndex        =   3
            Top             =   135
            Width           =   12585
            _ExtentX        =   22199
            _ExtentY        =   4789
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Nº Recibo"
            Columns(0).DataField=   "cabrec_numrecibo"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "I/E"
            Columns(1).DataField=   "cabrec_ingsal"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Cod. Caja"
            Columns(2).DataField=   "cajacodigo"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Razón Social"
            Columns(3).DataField=   "clientecodigo"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Fech Doc."
            Columns(4).DataField=   "cabrec_fechadocumento"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "T.OP"
            Columns(5).DataField=   "operacioncodigo"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "Moneda"
            Columns(6).DataField=   "monedacodigo"
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(7)._VlistStyle=   0
            Columns(7)._MaxComboItems=   5
            Columns(7).Caption=   "Soles"
            Columns(7).DataField=   "cabrec_totsoles"
            Columns(7).NumberFormat=   "###,###,###.00"
            Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(8)._VlistStyle=   0
            Columns(8)._MaxComboItems=   5
            Columns(8).Caption=   "Dolares"
            Columns(8).DataField=   "cabrec_totdolares"
            Columns(8).NumberFormat=   "###,###,###.00"
            Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(9)._VlistStyle=   0
            Columns(9)._MaxComboItems=   5
            Columns(9).Caption=   "Nº C. Contable"
            Columns(9).DataField=   "Comprobconta"
            Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   10
            Splits(0)._UserFlags=   0
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=10"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=1402"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1323"
            Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(5)=   "Column(1).Width=900"
            Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=820"
            Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(9)=   "Column(2).Width=1693"
            Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=1614"
            Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(13)=   "Column(3).Width=3625"
            Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=3545"
            Splits(0)._ColumnProps(16)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(17)=   "Column(4).Width=1773"
            Splits(0)._ColumnProps(18)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(19)=   "Column(4)._WidthInPix=1693"
            Splits(0)._ColumnProps(20)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(21)=   "Column(5).Width=582"
            Splits(0)._ColumnProps(22)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(23)=   "Column(5)._WidthInPix=503"
            Splits(0)._ColumnProps(24)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(25)=   "Column(6).Width=2355"
            Splits(0)._ColumnProps(26)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(27)=   "Column(6)._WidthInPix=2275"
            Splits(0)._ColumnProps(28)=   "Column(6).Order=7"
            Splits(0)._ColumnProps(29)=   "Column(7).Width=2064"
            Splits(0)._ColumnProps(30)=   "Column(7).DividerColor=0"
            Splits(0)._ColumnProps(31)=   "Column(7)._WidthInPix=1984"
            Splits(0)._ColumnProps(32)=   "Column(7)._ColStyle=2"
            Splits(0)._ColumnProps(33)=   "Column(7).Order=8"
            Splits(0)._ColumnProps(34)=   "Column(8).Width=1879"
            Splits(0)._ColumnProps(35)=   "Column(8).DividerColor=0"
            Splits(0)._ColumnProps(36)=   "Column(8)._WidthInPix=1799"
            Splits(0)._ColumnProps(37)=   "Column(8)._ColStyle=2"
            Splits(0)._ColumnProps(38)=   "Column(8).Order=9"
            Splits(0)._ColumnProps(39)=   "Column(9).Width=2725"
            Splits(0)._ColumnProps(40)=   "Column(9).DividerColor=0"
            Splits(0)._ColumnProps(41)=   "Column(9)._WidthInPix=2646"
            Splits(0)._ColumnProps(42)=   "Column(9)._ColStyle=1"
            Splits(0)._ColumnProps(43)=   "Column(9).Order=10"
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
            _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=62,.parent=13,.alignment=1"
            _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=59,.parent=14"
            _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=60,.parent=15"
            _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=61,.parent=17"
            _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=66,.parent=13,.alignment=1"
            _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=63,.parent=14"
            _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=64,.parent=15"
            _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=65,.parent=17"
            _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=78,.parent=13,.alignment=2,.bgcolor=&HFFFFB7&"
            _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=75,.parent=14,.bgcolor=&HFFFF00&"
            _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=76,.parent=15"
            _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=77,.parent=17"
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
         Begin VB.Label lbreg1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "0 "
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   11550
            TabIndex        =   5
            Top             =   2985
            Width           =   1140
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Nro. Reg :"
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Left            =   10500
            TabIndex        =   4
            Top             =   2985
            Width           =   930
         End
      End
      Begin VB.Label Label12 
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
         Height          =   285
         Left            =   4080
         TabIndex        =   27
         Top             =   720
         Width           =   2850
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFFFFF&
         Height          =   390
         Left            =   120
         TabIndex        =   26
         Top             =   600
         Width           =   12900
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   4560
         Picture         =   "FrmContabilizaTesoreria.frx":0038
         Top             =   0
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   -74880
         Picture         =   "FrmContabilizaTesoreria.frx":0E7A
         Top             =   120
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
         Height          =   285
         Left            =   -70440
         TabIndex        =   14
         Top             =   600
         Width           =   2850
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00FFFFFF&
         Height          =   15
         Left            =   -74880
         Top             =   1275
         Width           =   12900
      End
      Begin VB.Shape Shape1 
         Height          =   15
         Left            =   -74880
         Top             =   900
         Width           =   12900
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Height          =   390
         Left            =   -74880
         TabIndex        =   13
         Top             =   480
         Width           =   12900
      End
      Begin VB.Label Label2 
         BackColor       =   &H00DDF7F9&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Provisiones Contabilizadas en el Mes"
         Height          =   315
         Left            =   -74850
         TabIndex        =   12
         Top             =   4860
         Width           =   12825
      End
      Begin VB.Label Label1 
         BackColor       =   &H00DDF7F9&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Provisiones no Contabilizadas en el Mes"
         Height          =   315
         Left            =   -74880
         TabIndex        =   11
         Top             =   975
         Width           =   12855
      End
   End
   Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuempresa 
      Height          =   375
      Left            =   1050
      TabIndex        =   0
      Top             =   60
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
      BackColor       =   &H80000009&
      Caption         =   "Empresa :"
      Height          =   195
      Left            =   240
      TabIndex        =   28
      Top             =   150
      Width           =   705
   End
End
Attribute VB_Name = "FrmContabilizaTesoreria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsNoConta As ADODB.Recordset
Dim RsSiConta As ADODB.Recordset
Dim RsNoContaTransf As ADODB.Recordset
Dim RsSiContatransf As ADODB.Recordset
Dim rsparimpo As ADODB.Recordset
Public rrsql As New ADODB.Recordset

Private Sub Cmd_ElimConta_Click()
If Ctr_Ayuempresa.xclave = "" Then
  MsgBox (" Ingrese codigo de Empresa ")
  Ctr_Ayuempresa.SetFocus
  Exit Sub
End If

Screen.MousePointer = 11
Dim rsql As New ADODB.Recordset

sqlcad = " select asiento from ct_importartesoreria where tipooperacion<>'T' "
Set rsql = VGCNx.Execute(sqlcad)
sqlcad1 = "("
Do Until rsql.EOF
   sqlcad1 = sqlcad1 & "'" & rsql!asiento & "',"
   rsql.MoveNext
Loop
sqlcad1 = Left(sqlcad1, Len(sqlcad1) - 1) + ")"


SQL = " Update te_cabecerarecibos Set comprobconta='' Where month(cabrec_fechadocumento)=" & VGParamSistem.MesProceso
SQL = SQL & "  and year(cabrec_fechadocumento)=" & VGParamSistem.AnoProceso & " and isnull(cabrec_transferenciaautomatico,0)=0"
If Ctr_Ayuempresa.xclave <> "" Then
 SQL = SQL & " and empresacodigo='" & Ctr_Ayuempresa.xclave & "'"
End If
VGCNx.Execute SQL
    
SQL = " Delete ct_cabcomprob" & VGParamSistem.AnoProceso & " Where cabcomprobmes=" & VGParamSistem.MesProceso
SQL = SQL & " and asientocodigo in  " & sqlcad1 & "   and subasientocodigo='0099' And substring(cabcomprobnprovi,4,1)<>'0' and cabcomprobnprovi is not null "
If Ctr_Ayuempresa.xclave <> "" Then
 SQL = SQL & " and empresacodigo='" & Ctr_Ayuempresa.xclave & "'"
End If
VGCnxCT.Execute SQL
Call CargarDatos
Screen.MousePointer = 1
End Sub

Private Sub Cmd_ElimContaTransf_Click()
If Ctr_Ayuempresa.xclave = "" Then
  MsgBox (" Ingrese codigo de Empresa ")
  Ctr_Ayuempresa.SetFocus
  Exit Sub
End If
Screen.MousePointer = 11
Dim rsql As New ADODB.Recordset

sqlcad = " select asiento from ct_importartesoreria where tipooperacion='T' "
Set rsql = VGCNx.Execute(sqlcad)
sqlcad1 = "("
Do Until rsql.EOF
   sqlcad1 = sqlcad1 & "'" & rsql!asiento & "',"
   rsql.MoveNext
Loop
sqlcad1 = Left(sqlcad1, Len(sqlcad1) - 1) + ")"

SQL = " Update te_cabecerarecibos Set comprobconta='' Where month(cabrec_fechadocumento)=" & VGParamSistem.MesProceso
SQL = SQL & "  and year(cabrec_fechadocumento)=" & VGParamSistem.AnoProceso & " and isnull(cabrec_transferenciaautomatico,0)=1"
If Ctr_Ayuempresa.xclave <> "" Then
 SQL = SQL & " and empresacodigo='" & Ctr_Ayuempresa.xclave & "'"
End If
VGCNx.Execute SQL
    
SQL = " Delete ct_cabcomprob" & VGParamSistem.AnoProceso & " Where cabcomprobmes=" & VGParamSistem.MesProceso
SQL = SQL & " and asientocodigo in " & sqlcad1 & " and subasientocodigo='0099' And substring(cabcomprobnprovi,4,1)='0' and cabcomprobnprovi is not null "
If Ctr_Ayuempresa.xclave <> "" Then
 SQL = SQL & " and empresacodigo='" & Ctr_Ayuempresa.xclave & "'"
End If
VGCnxCT.Execute SQL
Call CargarDatos
Screen.MousePointer = 1
End Sub

Private Sub Cmd_PasaConta_Click()
If Ctr_Ayuempresa.xclave = "" Then
  MsgBox (" Ingrese codigo de Empresa ")
  Ctr_Ayuempresa.SetFocus
  Exit Sub
End If
Set rrsql = Nothing
Set rrsql = VGCnxCT.Execute("select top 1 sistemactaajustedeb,sistemactaajustehab from ct_sistema")
If RsNoConta.RecordCount() > 1 Then
   RsNoConta.MoveFirst
End If
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
SQL = "Select * From  ct_importartesoreria Where Left(Upper(tipooperacion),1) <>'T' AND Left(Upper(tipocajabanco),1) ='" & RsNoConta!detrec_tipocajabanco & "' And monedacodigo='" & RsNoConta!monedacodigo & "' "
Set rsparimpo = VGCnxCT.Execute(SQL)
If rsparimpo.RecordCount() = 0 Then
   Screen.MousePointer = 1
   MsgBox "Verifique el parametro del asiento de " & RsNoConta!cabrec_ingsal & " en ct_importartesoreria ", vbInformation, "Sistema de Tesoreria"
   Exit Sub
End If
VGCNx.BeginTrans
Set VGCommandoSP = New ADODB.Command
Set VGvardllgen = New dllgeneral.dll_general
    comprobconta = Format(VGParamSistem.MesProceso, "00") + rsparimpo!asiento + Format(nn, "00000")
    Set Comando = New ADODB.Command
    With Comando
         .CommandType = adCmdStoredProc
         .CommandText = "te_GeneraAsientosTesoreriaLinea_pro"
         .CommandTimeout = 0
         .ActiveConnection = VGgeneral
         .Parameters.Refresh
         .Parameters("@BaseConta") = VGCnxCT.DefaultDatabase
         .Parameters("@BaseVenta") = VGCNx.DefaultDatabase
         .Parameters("@empresa") = RsNoConta!empresacodigo
         .Parameters("@Asiento") = rsparimpo!asiento
         .Parameters("@SubAsiento") = rsparimpo!SubAsiento
         .Parameters("@Libro") = rsparimpo!libro
         .Parameters("@Mes") = Format(VGParamSistem.MesProceso, "00")
         .Parameters("@Ano") = VGParamSistem.AnoProceso
         .Parameters("@tipanal") = "001"
         .Parameters("@Compu") = VGComputer
         .Parameters("@Usuario") = VGusuario
         .Parameters("@TipoMov") = RsNoConta!cabrec_ingsal
         .Parameters("@Nrecibo") = RsNoConta!cabrec_numrecibo
         .Parameters("@op") = 1
         .Parameters("@comprobconta") = comprobconta
         .Parameters("@ajustehaber") = RTrim(rrsql!sistemactaajustehab)
         .Parameters("@ajustedebe") = RTrim(rrsql!sistemactaajustedeb)
         
         .Execute
     End With
If n = 0 Then VGCNx.CommitTrans
Exit Sub
genasiento:
    n = 1
    Screen.MousePointer = 1
    VGCNx.RollbackTrans
    MsgBox "Hubo Errores al momento que se genero el recibo " & RsNoConta!cabrec_numrecibo & Chr(13) & Err.Description
    Resume Next
End Sub
Private Sub grabaTransf(nn As Integer)
Dim Comando As ADODB.Command
On Error GoTo genasiento
Screen.MousePointer = 11
VGgeneral.BeginTrans
    comprobconta = Format(VGParamSistem.MesProceso, "00") + rsparimpo!asiento + Format(nn, "00000")
    Set Comando = New ADODB.Command
    With Comando
         .CommandType = adCmdStoredProc
         .CommandText = "te_GeneraAsientosTesoreriaTransflinea_pro"
         .CommandTimeout = 0
         .ActiveConnection = VGgeneral
         .Parameters.Refresh
         .Parameters("@BaseVenta") = VGCnxCT.DefaultDatabase
         .Parameters("@BaseConta") = VGCNx.DefaultDatabase
         .Parameters("@Asiento") = rsparimpo!asiento
         .Parameters("@SubAsiento") = rsparimpo!SubAsiento
         .Parameters("@Libro") = rsparimpo!libro
         .Parameters("@Mes") = Format(VGParamSistem.MesProceso, "00")
         .Parameters("@Ano") = VGParamSistem.AnoProceso
         .Parameters("@Compu") = VGComputer
         .Parameters("@Usuario") = VGusuario
         .Parameters("@Ntransfer") = RsNoContaTransf!cabrec_numreciboegreso
         .Parameters("@empresa") = RsNoContaTransf!empresacodigo
         .Parameters("@ajustehaber") = RTrim(rrsql!sistemactaajustehab)
         .Parameters("@ajustedebe") = RTrim(rrsql!sistemactaajustedeb)
         
         .Execute
     End With
If n = 0 Then VGgeneral.CommitTrans
Exit Sub
genasiento:
    nn = 1
    Screen.MousePointer = 1
    VGgeneral.RollbackTrans
    MsgBox "Hubo Errores al momento que se genero el recibo " & RsNoConta!cabrec_numrecibo & Chr(13) & Err.Description
    Resume Next
End Sub

Private Sub Cmd_PasaContaTransf_Click()
If Ctr_Ayuempresa.xclave = "" Then
  MsgBox (" Ingrese codigo de Empresa ")
  Ctr_Ayuempresa.SetFocus
  Exit Sub
End If
Dim n As Integer
Dim nn As Integer
Set rrsql = Nothing
Set rrsql = VGCnxCT.Execute("select top 1 sistemactaajustedeb,sistemactaajustehab from ct_sistema")
 RsNoContaTransf.MoveFirst
 nn = 0
 Do While Not RsNoContaTransf.EOF
    n = 0
    nn = nn + 1
   Set rsparimpo = New ADODB.Recordset
   Set rsparimpo = VGCnxCT.Execute("Select * From  ct_importartesoreria Where Left(Upper(tipooperacion),1) ='T' And Left(Upper(tipocajabanco),1) ='" & RsNoContaTransf!detrec_tipocajabanco & "' And monedacodigo='" & RsNoContaTransf!monedacodigo & "' ")
   If rsparimpo.RecordCount() = 0 Then
      Screen.MousePointer = 1
      MsgBox "Verifique parametro en ct_importartesoreria del Numero de transferencia   " & RsNoContaTransf!cabrec_numreciboegreso, vbInformation, "Sistema de Tesoreria"
      Exit Sub
   End If
    Call grabaTransf(nn)
    RsNoContaTransf.MoveNext
 Loop
 Screen.MousePointer = 1
 MsgBox "La Operacion se Realizo Satisfactoriamente", vbInformation, "Sistema de Tesoreria"
 Call CargarDatos
End Sub

    Private Sub Ctr_Ayuempresa_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
 Call CargarDatos
 Cmd_PasaConta.Enabled = Valida("Contabilidad")
 Cmd_ElimConta.Enabled = Valida("Contabilidad")

 Cmd_PasaContaTransf.Enabled = Valida("Contabilidad")
 Cmd_ElimContaTransf.Enabled = Valida("Contabilidad")

End Sub

Private Sub Form_Load()
Call Ctr_Ayuempresa.conexion(VGCNx)
Call CargarDatos
End Sub
Private Sub CargarDatos()
    Set RsNoConta = New ADODB.Recordset
    Set RsSiConta = New ADODB.Recordset
    SQL = "Select distinct a.*,b.detrec_tipocajabanco from te_cabecerarecibos a Inner Join te_detallerecibos b On a.cabrec_numrecibo=b.cabrec_numrecibo Where month(a.cabrec_fechadocumento)=" & CInt(VGParamSistem.MesProceso)
    SQL = SQL & " and ltrim(rtrim(isnull(a.comprobConta,'')))='' and year(a.cabrec_fechadocumento)=" & VGParamSistem.AnoProceso
    SQL = SQL & " and isnull(a.cabrec_transferenciaautomatico,0)<>1 and cabrec_estadoreg<>1 "
    If Ctr_Ayuempresa.xclave <> "" Then SQL = SQL & " and empresacodigo='" & Ctr_Ayuempresa.xclave & "'"
    RsNoConta.Open (SQL), VGCNx, adOpenKeyset, adLockReadOnly
    Set TDBG_Noconta.DataSource = RsNoConta
    TDBG_Noconta.Refresh
    lbreg1.Caption = Format(RsNoConta.RecordCount, "0 ")
    
    SQL = "Select * from te_cabecerarecibos Where month(cabrec_fechadocumento)=" & CInt(VGParamSistem.MesProceso)
    SQL = SQL & " and ltrim(rtrim(isnull(comprobConta,'')))<>'' and year(cabrec_fechadocumento)=" & VGParamSistem.AnoProceso
    SQL = SQL & " and isnull(cabrec_transferenciaautomatico,0)<>1 and cabrec_estadoreg<>1"
    If Ctr_Ayuempresa.xclave <> "" Then SQL = SQL & " and empresacodigo='" & Ctr_Ayuempresa.xclave & "'"
    RsSiConta.Open SQL, VGCNx, adOpenKeyset, adLockReadOnly
    Set TDBG_siConta.DataSource = RsSiConta
    TDBG_siConta.Refresh
    lbreg2.Caption = Format(RsSiConta.RecordCount, "0 ")
    
    Set RsNoContaTransf = New ADODB.Recordset
    Set RsSiContatransf = New ADODB.Recordset
    SQL = "select z.* from ( Select a.*,b.detrec_tipocajabanco from te_cabecerarecibos a Inner Join te_detallerecibos b On a.cabrec_numrecibo=b.cabrec_numrecibo where isnull(a.cabrec_transferenciaautomatico,0)=1 and a.cabrec_ingsal='E'"
'   SQL = "select z.* from ( Select a.*,b.detrec_tipocajabanco from te_cabecerarecibos a Inner Join te_detallerecibos b On a.cabrec_numrecibo=b.cabrec_numrecibo where and cabrec_numreciboegreso <>'' and numerodocxrendir<>'' a.cabrec_ingsal='E'"
    SQL = SQL & " union all Select a.*,b.detrec_tipocajabanco from te_cabecerarecibos a Inner Join te_detallerecibos b On a.cabrec_numrecibo=b.cabrec_numrecibo where isnull(a.cabrec_transferenciaautomatico,0)=1 and a.cabrec_ingsal='I'"
    SQL = SQL & " and a.cabrec_numreciboegreso in ( select cabrec_numreciboegreso from te_cabecerarecibos where "
    SQL = SQL & " empresacodigo<>'" & Ctr_Ayuempresa.xclave & "' and cabrec_ingsal ='E' ) ) z "
    SQL = SQL & " Where Month(z.cabrec_fechadocumento) = " & CInt(VGParamSistem.MesProceso) & ""
    SQL = SQL & " and ltrim(rtrim(isnull(z.comprobConta,'')))='' and year(z.cabrec_fechadocumento)=" & VGParamSistem.AnoProceso
    SQL = SQL & " and   z.cabrec_estadoreg<>1 "
    If Ctr_Ayuempresa.xclave <> "" Then SQL = SQL & " and z.empresacodigo='" & Ctr_Ayuempresa.xclave & "'"
    RsNoContaTransf.Open (SQL), VGCNx, adOpenKeyset, adLockReadOnly
    Set TDBG_NoContaTransf.DataSource = RsNoContaTransf
    lbreg1.Caption = Format(RsNoContaTransf.RecordCount, "0 ")
    
    SQL = "select z.* from ( Select a.*,b.detrec_tipocajabanco from te_cabecerarecibos a Inner Join te_detallerecibos b On a.cabrec_numrecibo=b.cabrec_numrecibo where isnull(cabrec_transferenciaautomatico,0)=1 and cabrec_ingsal='E'"
'   SQL = "select z.* from ( Select * from te_cabecerarecibos where isnull(cabrec_transferenciaautomatico,0)=1 and cabrec_ingsal='E'"
    SQL = SQL & " union all Select a.*,b.detrec_tipocajabanco from te_cabecerarecibos a Inner Join te_detallerecibos b On a.cabrec_numrecibo=b.cabrec_numrecibo where isnull(cabrec_transferenciaautomatico,0)=1 and cabrec_ingsal='I'"
    SQL = SQL & " and cabrec_numreciboegreso in ( select cabrec_numreciboegreso from te_cabecerarecibos where "
    SQL = SQL & " empresacodigo<>'" & Ctr_Ayuempresa.xclave & "' and cabrec_ingsal ='E' ) ) z "
    SQL = SQL & " Where Month(z.cabrec_fechadocumento) = " & CInt(VGParamSistem.MesProceso) & ""
    SQL = SQL & " and ltrim(rtrim(isnull(z.comprobConta,'')))<>'' and year(z.cabrec_fechadocumento)=" & VGParamSistem.AnoProceso
    SQL = SQL & " and   z.cabrec_estadoreg<>1 "
    If Ctr_Ayuempresa.xclave <> "" Then SQL = SQL & " and z.empresacodigo='" & Ctr_Ayuempresa.xclave & "'"
    RsSiContatransf.Open SQL, VGCNx, adOpenKeyset, adLockReadOnly
    Set TDBG_SiContaTransf.DataSource = RsSiContatransf
    lbreg2.Caption = Format(RsSiContatransf.RecordCount, "0 ")
    
    TDBG_NoContaTransf.FetchRowStyle = True

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

