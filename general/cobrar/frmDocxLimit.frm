VERSION 5.00
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frmDocxLimit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Documentos x Grupos de Limite de Credito"
   ClientHeight    =   6615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6705
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   6705
   Begin TabDlg.SSTab SSTab1 
      Height          =   6525
      Left            =   30
      TabIndex        =   1
      Top             =   15
      Width           =   6645
      _ExtentX        =   11721
      _ExtentY        =   11509
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Consulta x Grupos Limite Credito"
      TabPicture(0)   =   "frmDocxLimit.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblNumReg"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "TDBGrid1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Mantenimiento"
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "frmbotones"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cCancela"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "cAcepta"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Frame1"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "lblMensaje"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).ControlCount=   5
      Begin VB.Frame frmbotones 
         Height          =   555
         Left            =   -74490
         TabIndex        =   10
         Top             =   5460
         Width           =   5730
         Begin VB.CommandButton cmdBotones 
            Caption         =   "&Nuevo"
            Height          =   330
            Index           =   0
            Left            =   60
            TabIndex        =   15
            Top             =   165
            Width           =   1080
         End
         Begin VB.CommandButton cmdBotones 
            Caption         =   "E&ditar"
            Height          =   330
            Index           =   1
            Left            =   1185
            TabIndex        =   14
            Top             =   165
            Width           =   1080
         End
         Begin VB.CommandButton cmdBotones 
            Caption         =   "&Eliminar"
            Height          =   330
            Index           =   2
            Left            =   2310
            TabIndex        =   13
            Top             =   165
            Width           =   1080
         End
         Begin VB.CommandButton cmdBotones 
            Caption         =   "&Salir"
            Height          =   330
            Index           =   4
            Left            =   4560
            TabIndex        =   12
            Top             =   165
            Width           =   1080
         End
         Begin VB.CommandButton cmdBotones 
            Caption         =   "&Imprimir"
            Height          =   330
            Index           =   3
            Left            =   3435
            TabIndex        =   11
            Top             =   165
            Width           =   1080
         End
      End
      Begin VB.CommandButton cCancela 
         Caption         =   "&Cancelar"
         Height          =   345
         Left            =   -71535
         TabIndex        =   9
         Top             =   5040
         Width           =   1140
      End
      Begin VB.CommandButton cAcepta 
         Caption         =   "&Aceptar"
         Height          =   345
         Left            =   -72975
         TabIndex        =   8
         Top             =   5040
         Width           =   1140
      End
      Begin VB.Frame Frame1 
         Height          =   4530
         Left            =   -74955
         TabIndex        =   2
         Top             =   330
         Width           =   6540
         Begin TrueOleDBGrid70.TDBGrid TDBGrid2 
            Height          =   3270
            Left            =   30
            TabIndex        =   16
            Top             =   1095
            Width           =   6405
            _ExtentX        =   11298
            _ExtentY        =   5768
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Código"
            Columns(0).DataField=   "coddoc"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Descripción"
            Columns(1).DataField=   "tdocumentodescripcion"
            Columns(1).DataWidth=   1700
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   2
            Splits(0)._UserFlags=   0
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=2"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(5)=   "Column(1).Width=2725"
            Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2646"
            Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
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
            _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=84,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.alignment=3,.bold=0,.fontsize=825"
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
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=106,.parent=13"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=103,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=104,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=105,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=102,.parent=13"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=99,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=100,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=101,.parent=17"
            _StyleDefs(44)  =   "Named:id=33:Normal"
            _StyleDefs(45)  =   ":id=33,.parent=0"
            _StyleDefs(46)  =   "Named:id=34:Heading"
            _StyleDefs(47)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(48)  =   ":id=34,.wraptext=-1"
            _StyleDefs(49)  =   "Named:id=35:Footing"
            _StyleDefs(50)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(51)  =   "Named:id=36:Selected"
            _StyleDefs(52)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(53)  =   "Named:id=37:Caption"
            _StyleDefs(54)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(55)  =   "Named:id=38:HighlightRow"
            _StyleDefs(56)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(57)  =   "Named:id=39:EvenRow"
            _StyleDefs(58)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(59)  =   "Named:id=40:OddRow"
            _StyleDefs(60)  =   ":id=40,.parent=33"
            _StyleDefs(61)  =   "Named:id=41:RecordSelector"
            _StyleDefs(62)  =   ":id=41,.parent=34"
            _StyleDefs(63)  =   "Named:id=42:FilterBar"
            _StyleDefs(64)  =   ":id=42,.parent=33"
         End
         Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuda1 
            Height          =   315
            Left            =   1890
            TabIndex        =   0
            Top             =   150
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   556
            XcodMaxLongitud =   2
            xcodwith        =   500
            NomTabla        =   "cc_limcredgrupo"
            TituloAyuda     =   "Ayuda de Grupo de Limite de Credito"
            ListaCampos     =   "codgrup(1),descgrup(1)"
            XcodCampo       =   "codgrup"
            XListCampo      =   "descgrup"
            ListaCamposDescrip=   "Código,Descripción"
            ListaCamposText =   "codgrup,descgrup"
         End
         Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuda2 
            Height          =   315
            Left            =   1890
            TabIndex        =   18
            Top             =   465
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   556
            XcodMaxLongitud =   2
            xcodwith        =   500
            NomTabla        =   "cc_tipodocumento"
            TituloAyuda     =   "Ayuda de Tipo de Documentos"
            ListaCampos     =   "tdocumentocodigo(1),tdocumentodescripcion(1)"
            XcodCampo       =   "tdocumentocodigo"
            XListCampo      =   "tdocumentodescripcion"
            ListaCamposDescrip=   "Código,Descripción"
            ListaCamposText =   "tdocumentocodigo,tdocumentodescripcion"
         End
         Begin VB.Label lbl 
            Caption         =   "Tipo de Documentos"
            Height          =   285
            Index           =   1
            Left            =   120
            TabIndex        =   7
            Top             =   510
            Width           =   1590
         End
         Begin VB.Label lbl 
            Caption         =   "Grupo de Limite Cred."
            Height          =   285
            Index           =   0
            Left            =   135
            TabIndex        =   6
            Top             =   210
            Width           =   1665
         End
      End
      Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
         Height          =   3525
         Left            =   30
         TabIndex        =   3
         Top             =   735
         Width           =   6570
         _ExtentX        =   11589
         _ExtentY        =   6218
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
         Columns.Count   =   2
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   12632256
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=2"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(5)=   "Column(1).Width=2725"
         Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2646"
         Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
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
         MultipleLines   =   1
         CellTipsWidth   =   0
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
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=64,.bold=0,.fontsize=825,.italic=0"
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
         _StyleDefs(44)  =   "Named:id=33:Normal"
         _StyleDefs(45)  =   ":id=33,.parent=0"
         _StyleDefs(46)  =   "Named:id=34:Heading"
         _StyleDefs(47)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(48)  =   ":id=34,.wraptext=-1"
         _StyleDefs(49)  =   "Named:id=35:Footing"
         _StyleDefs(50)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(51)  =   "Named:id=36:Selected"
         _StyleDefs(52)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(53)  =   "Named:id=37:Caption"
         _StyleDefs(54)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(55)  =   "Named:id=38:HighlightRow"
         _StyleDefs(56)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(57)  =   "Named:id=39:EvenRow"
         _StyleDefs(58)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(59)  =   "Named:id=40:OddRow"
         _StyleDefs(60)  =   ":id=40,.parent=33"
         _StyleDefs(61)  =   "Named:id=41:RecordSelector"
         _StyleDefs(62)  =   ":id=41,.parent=34"
         _StyleDefs(63)  =   "Named:id=42:FilterBar"
         _StyleDefs(64)  =   ":id=42,.parent=33"
      End
      Begin VB.Label lblMensaje 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   -72630
         TabIndex        =   19
         Top             =   6165
         Width           =   1845
      End
      Begin VB.Label Label3 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Seleccionar un Grupo de Limite de Credito"
         ForeColor       =   &H00800000&
         Height          =   345
         Left            =   30
         TabIndex        =   17
         Top             =   420
         Width           =   6570
      End
      Begin VB.Label Label1 
         Caption         =   "Nº Registros"
         Height          =   270
         Left            =   4740
         TabIndex        =   5
         Top             =   4350
         Width           =   900
      End
      Begin VB.Label lblNumReg 
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblNumReg"
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   5685
         TabIndex        =   4
         Top             =   4335
         Width           =   915
      End
   End
End
Attribute VB_Name = "frmDocxLimit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim modoinsert As Boolean
Dim modoedit As Boolean
Dim rs As New ADODB.Recordset
Dim rsAsiento As ADODB.Recordset

Private Sub Ctr_Ayuda2_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
    If Ctr_Ayuda2.Enabled Then cAcepta.Enabled = True
End Sub

Private Sub Ctr_Ayuda2_AlNoDevolverNada()
    cAcepta.Enabled = False
End Sub

Private Sub Form_Load()
  Call ConfiguraForm
  Call MuestraDatosAsiento
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set rs = Nothing
  Set rsAsiento = Nothing
  Set VGvardllgen = Nothing
End Sub

Sub ConfiguraForm()
  SSTab1.Tab = 0
  SSTab1.TabEnabled(1) = False
  Ctr_Ayuda1.conexion cn
  Ctr_Ayuda2.conexion cn
  
  'Ctr_Ayuda2.Filtro = "monedacodigo<>'00'"
  cAcepta.Enabled = False
  lblNumReg.Caption = Empty
  Me.Width = 6825
  Me.Height = 6990
End Sub

Sub MuestraDatosAsiento()
 Dim SQL  As String
    Set rsAsiento = New ADODB.Recordset
    SQL = "select codgrup as codigo ,descgrup as Descripción  from cc_limcredgrupo where codgrup<>'00' order by 1"
    Set rsAsiento = VGCNx.Execute(SQL)
    Set TDBGrid1.DataSource = rsAsiento
    TDBGrid1.Columns(0).Width = 1400
    TDBGrid1.Columns(1).Width = 4800
    lblNumReg.Caption = rsAsiento.RecordCount
End Sub

Private Sub Ctr_Ayuda1_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
  Call MuestraDatosSubAsiento
End Sub

'FIXIT: Declare 'MuestraDatosSubAsiento' con un tipo de datos de enlace en tiempo de compilación     FixIT90210ae-R1672-R1B8ZE
Public Function MuestraDatosSubAsiento()
 Dim SQL As String
 
  SQL = " select coddoc,tdocumentodescripcion " & _
       " from c_docxlimicred A " & _
       " inner join dbo.cc_tipodocumento B " & _
       " on A.coddoc=B.tdocumentocodigo " & _
       " where codgrup <> '00' and codgrup='" & RTrim$(Ctr_Ayuda1.xclave) & "' " & _
       " ORDER BY 1,2  "
  Set rs = VGCNx.Execute(SQL)
  Set TDBGrid2.DataSource = rs
  Call ConfiguraGridSubAsientos
  If rs.RecordCount <= 0 Then Call LimpiarForm(frmDocxLimit, "ctr_ayuda1")
  
End Function
Private Sub cmdBotones_Click(Index As Integer)
  Dim J As Integer
  Dim spos As Integer
  Dim SQL As String
  
  On Error GoTo X
  
  Select Case Index
     Case 0   'nuevo
        'SSTab1.TabEnabled(2) = True
        SSTab1.Tab = 1
        'Call LimpiarValores
        
        Call LimpiarForm(frmDocxLimit, "Ctr_Ayuda1")
        
        Call ModoEditable(True, frmDocxLimit, "Ctr_Ayuda1")
        frmbotones.Visible = False
        modoinsert = True
        lblMensaje.Caption = "Nuevo"
        
     Case 1   'modificar
        If TDBGrid1.Row < 0 Then
          Exit Sub
        End If
        'SSTab1.TabEnabled(2) = True
        SSTab1.Tab = 1
        modoedit = True
        frmbotones.Visible = False
        Call ModoEditable(True, frmDocxLimit, "Ctr_Ayuda1")
        lblMensaje.Caption = "Editar"
      'codgrup,coddoc
     Case 2   'eliminar
       If MsgBox("Desea eliminar el registro de Documento codigo= " & TDBGrid2.Columns(0).Value & "?", vbYesNo + vbDefaultButton2, "AVISO") = vbYes Then
          SQL = "DELETE FROM c_docxlimicred WHERE coddoc='" & RTrim$(TDBGrid2.Columns(0).Value) & "' AND "
          SQL = SQL & "codgrup='" & RTrim$(Ctr_Ayuda1.xclave) & "'"
          VGCNx.Execute (SQL)
          Call MuestraDatosSubAsiento
       End If
        
     Case 3   'imprimir
       'Call Impresion("rptSubAsiento.rpt")
     
     Case 4  ' salir
       Unload Me
  End Select
  
  Exit Sub
   
X:
  If Index = 2 And Err.Number = -2147217873 Then
    MsgBox "Registro no podrá Eliminarse mientras exista Información en la Tablas Relacionadas", vbInformation, Caption
  Else
    MsgBox "Error inesperado: " & Err.Description & "  " & Err.Number, vbInformation, Caption
  End If
   
End Sub

Private Sub cAcepta_Click()
  Dim SQL As String
  On Error GoTo X
  
  Set VGvardllgen = New dllgeneral.dll_general
  VGCNx.BeginTrans
  
  If modoinsert = True Then
    SQL = "INSERT INTO c_docxlimicred(codgrup,coddoc)" & _
          "VALUES ('" & Ctr_Ayuda1.xclave & "','" & Ctr_Ayuda2.xclave & "')"
    
  ElseIf modoedit = True Then
    SQL = "UPDATE c_docxlimicred SET coddoc='" & Ctr_Ayuda2.xclave & "'" & _
          "WHERE  coddoc='" & TDBGrid2.Columns(0).Value & "' AND codgrup='" & Ctr_Ayuda1.xclave & "'"
  End If
  
  VGCNx.Execute (SQL)
  VGCNx.CommitTrans
  
  Set VGvardllgen = Nothing
  frmbotones.Visible = True
  modoinsert = False: modoedit = False: lblMensaje.Caption = Empty
  Call MuestraDatosSubAsiento
  cAcepta.Enabled = False
  Set VGvardllgen = Nothing
  Call ModoEditable(False, frmDocxLimit, "")
  Exit Sub

X:
  If Err.Number = -2147217873 Then
    MsgBox "Esta intentando registrar Código de documento Existente ", vbInformation, Caption
    
  Else
    MsgBox "Error inesperado: " & Err.Number & " " & Err.Description, vbInformation, Caption
  End If
  VGCNx.RollbackTrans
     
End Sub

Private Sub cCancela_Click()
  frmbotones.Visible = True
  modoinsert = False: modoedit = False: lblMensaje.Caption = Empty
  cAcepta.Enabled = False
  SSTab1.Tab = 0
  SSTab1.TabEnabled(0) = True
  SSTab1.TabEnabled(1) = False
  
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
  If PreviousTab = 0 Then SSTab1.TabEnabled(PreviousTab) = False
End Sub

Private Sub TDBGrid1_HeadClick(ByVal ColIndex As Integer)
On Error Resume Next
    If rs.Sort = Empty Then
        rs.Sort = TDBGrid1.Columns.Item(ColIndex).DataField & " asc"
     ElseIf Right$(rs.Sort, 3) = "asc" Then
        rs.Sort = TDBGrid1.Columns.Item(ColIndex).DataField & " desc"
     ElseIf Right$(rs.Sort, 4) = "desc" Then
        rs.Sort = TDBGrid1.Columns.Item(ColIndex).DataField & " asc"
    End If
    TDBGrid1.Refresh
End Sub

Private Sub TDBGrid1_DblClick()
 If rsAsiento.RecordCount > 0 Then
   SSTab1.TabEnabled(1) = True
   SSTab1.Tab = 1
   Ctr_Ayuda1.xclave = TDBGrid1.Columns(0).Text: Ctr_Ayuda1.Ejecutar
   Ctr_Ayuda1.Enabled = False
   Call ModoEditable(False, frmDocxLimit, "Ctr_Ayuda1")
   cAcepta.Enabled = False
 End If
End Sub

'FIXIT: Declare 'LastRow' con un tipo de datos de enlace en tiempo de compilación          FixIT90210ae-R1672-R1B8ZE
Private Sub TDBGrid2_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
  Call EditarSubAsiento
End Sub

Private Sub txt_Change(Index As Integer)
 If modoinsert = True Or modoedit = True Then
   cAcepta.Enabled = ValidaDataIngreso()
 End If
End Sub

Private Sub chk_Click()
    If modoinsert = True Or modoedit = True Then cAcepta.Enabled = ValidaDataIngreso()
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = 13 And Index = 15 Then
    cAcepta.SetFocus
    Call cAcepta_Click
  End If
End Sub
Sub EditarSubAsiento()
 Dim i As Integer
 
 If rs.RecordCount > 0 Then
    With TDBGrid2
        Ctr_Ayuda2.xclave = .Columns(0).Value: Ctr_Ayuda2.Ejecutar
    End With
 
 End If
End Sub
Sub ConfiguraGridSubAsientos()
 Dim i As Integer
 With TDBGrid2
   .Columns(0).Width = 700
   .Columns(1).Width = 2500
 End With
End Sub

Function ValidaDataIngreso() As Boolean
 Dim i As Integer
  If Ctr_Ayuda1.xclave = Empty Then
    ValidaDataIngreso = False
    Exit Function
  End If
   
  If Ctr_Ayuda2.xclave = Empty Then
    ValidaDataIngreso = False
    Exit Function
  End If
  
  ValidaDataIngreso = True
End Function

