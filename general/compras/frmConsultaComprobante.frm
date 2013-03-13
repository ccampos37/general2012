VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Object = "{4D137D9C-00A6-4458-9B46-7E95DB76D55B}#9.0#0"; "textfer.ocx"
Begin VB.Form frmConsultaComprobante 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Plantilla de Sub Asientos"
   ClientHeight    =   7512
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8028
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7512
   ScaleWidth      =   8028
   Begin TabDlg.SSTab SSTab1 
      Height          =   7248
      Left            =   36
      TabIndex        =   10
      Top             =   0
      Width           =   7812
      _ExtentX        =   13780
      _ExtentY        =   12785
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Consulta"
      TabPicture(0)   =   "frmConsultaComprobante.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblNumRegAsientos"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblNumRegSubAs"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame3"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame2"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Frame4"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "Mantenimiento"
      TabPicture(1)   =   "frmConsultaComprobante.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblMensaje"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame1"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "cAcepta"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "cCancela"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "frmbotones"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).ControlCount=   5
      Begin VB.Frame Frame4 
         Height          =   432
         Left            =   48
         TabIndex        =   42
         Top             =   384
         Width           =   7596
         Begin VB.CheckBox Check2 
            Height          =   240
            Left            =   1308
            TabIndex        =   43
            Top             =   456
            Width           =   576
         End
         Begin TextFer.TxFer txt 
            Height          =   312
            Index           =   4
            Left            =   6204
            TabIndex        =   44
            Top             =   132
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   550
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   ""
            ColorIlumina    =   -2147483624
            SaltarAlEnter   =   -1  'True
            Valor           =   ""
            NoCaracteres    =   "0123456789"
            MarcarTextoAlEnfoque=   -1  'True
            NoRangoCadena   =   -1  'True
         End
         Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuda7 
            Height          =   312
            Left            =   1332
            TabIndex        =   45
            Top             =   132
            Width           =   4044
            _ExtentX        =   7133
            _ExtentY        =   550
            Enabled         =   0   'False
            XcodMaxLongitud =   3
            xcodwith        =   500
            NomTabla        =   "ct_asiento"
            ListaCampos     =   "asientocodigo(1),asientodescripcion(1)"
            XcodCampo       =   "asientocodigo"
            XListCampo      =   "asientodescripcion"
            ListaCamposDescrip=   "Código,Descripción"
            ListaCamposText =   "asientocodigo,asientodescripcion"
         End
         Begin VB.Label lbl 
            Caption         =   "Proveedor"
            Height          =   285
            Index           =   15
            Left            =   135
            TabIndex        =   48
            Top             =   180
            Width           =   2310
         End
         Begin VB.Label lbl 
            Caption         =   "RUC"
            Height          =   288
            Index           =   12
            Left            =   5424
            TabIndex        =   47
            Top             =   192
            Width           =   828
         End
         Begin VB.Label lbl 
            Caption         =   "Cuenta Inafecta"
            Height          =   288
            Index           =   9
            Left            =   120
            TabIndex        =   46
            Top             =   480
            Width           =   2280
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Seleccionar Asientos"
         Height          =   3096
         Left            =   0
         TabIndex        =   36
         Top             =   960
         Width           =   7632
         Begin TrueOleDBGrid70.TDBGrid TDBGrid3 
            Height          =   2796
            Left            =   36
            TabIndex        =   37
            Top             =   216
            Width           =   7500
            _ExtentX        =   13229
            _ExtentY        =   4932
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Código"
            Columns(0).DataField=   "asientocodigo"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Descripción"
            Columns(1).DataField=   "asientodescripcion"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   2
            Splits(0)._UserFlags=   0
            Splits(0).RecordSelectorWidth=   508
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=2"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2731"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(5)=   "Column(1).Width=2731"
            Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2646"
            Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
            Splits.Count    =   1
            PrintInfos(0)._StateFlags=   3
            PrintInfos(0).Name=   "piInternal 0"
            PrintInfos(0).PageHeaderFont=   "Size=7.8,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageFooterFont=   "Size=7.8,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
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
            _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=34,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=780,.italic=0"
            _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
            _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
            _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=780,.italic=0"
            _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
            _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=780,.italic=0"
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
            _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=18,.parent=6,.namedParent=36"
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
      End
      Begin VB.Frame Frame3 
         Caption         =   "Seleccionar SubAsientos"
         Height          =   2400
         Left            =   15
         TabIndex        =   34
         Top             =   4404
         Width           =   7692
         Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
            Height          =   2112
            Left            =   36
            TabIndex        =   35
            Top             =   216
            Width           =   7524
            _ExtentX        =   13272
            _ExtentY        =   3725
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Código"
            Columns(0).DataField=   "subasientocodigo"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Descripción"
            Columns(1).DataField=   "subasientodescripcion"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   2
            Splits(0)._UserFlags=   0
            Splits(0).RecordSelectorWidth=   508
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=2"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2731"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(5)=   "Column(1).Width=2731"
            Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2646"
            Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
            Splits.Count    =   1
            PrintInfos(0)._StateFlags=   3
            PrintInfos(0).Name=   "piInternal 0"
            PrintInfos(0).PageHeaderFont=   "Size=7.8,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageFooterFont=   "Size=7.8,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
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
            _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=64,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=780,.italic=0"
            _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
            _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
            _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=780,.italic=0"
            _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
            _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=780,.italic=0"
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
      End
      Begin VB.Frame frmbotones 
         Height          =   555
         Left            =   -74415
         TabIndex        =   22
         Top             =   6570
         Width           =   5730
         Begin VB.CommandButton cmdBotones 
            Caption         =   "&Nuevo"
            Height          =   330
            Index           =   0
            Left            =   60
            TabIndex        =   27
            Top             =   165
            Width           =   1080
         End
         Begin VB.CommandButton cmdBotones 
            Caption         =   "E&ditar"
            Height          =   330
            Index           =   1
            Left            =   1185
            TabIndex        =   26
            Top             =   165
            Width           =   1080
         End
         Begin VB.CommandButton cmdBotones 
            Caption         =   "&Eliminar"
            Height          =   330
            Index           =   2
            Left            =   2310
            TabIndex        =   25
            Top             =   165
            Width           =   1080
         End
         Begin VB.CommandButton cmdBotones 
            Caption         =   "&Salir"
            Height          =   330
            Index           =   4
            Left            =   4560
            TabIndex        =   24
            Top             =   165
            Width           =   1080
         End
         Begin VB.CommandButton cmdBotones 
            Caption         =   "&Imprimir"
            Height          =   330
            Index           =   3
            Left            =   3435
            TabIndex        =   23
            Top             =   165
            Width           =   1080
         End
      End
      Begin VB.CommandButton cCancela 
         Caption         =   "&Cancelar"
         Height          =   345
         Left            =   -71610
         TabIndex        =   19
         Top             =   6060
         Width           =   1140
      End
      Begin VB.CommandButton cAcepta 
         Caption         =   "&Aceptar"
         Height          =   345
         Left            =   -73050
         TabIndex        =   18
         Top             =   6060
         Width           =   1140
      End
      Begin VB.Frame Frame1 
         Height          =   5610
         Left            =   -74940
         TabIndex        =   11
         Top             =   315
         Width           =   6585
         Begin VB.CheckBox ChkAjuste 
            Height          =   240
            Left            =   6285
            TabIndex        =   32
            Top             =   2055
            Width           =   240
         End
         Begin TrueOleDBGrid70.TDBGrid TDBGrid2 
            Height          =   2445
            Left            =   45
            TabIndex        =   21
            Top             =   3105
            Width           =   6465
            _ExtentX        =   11409
            _ExtentY        =   4318
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Codigo Asiento"
            Columns(0).DataField=   "asientocodigo"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Codigo SubAsiento"
            Columns(1).DataField=   "subasientocodigo"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Correlativo"
            Columns(2).DataField=   "plantillaasientocorrela"
            Columns(2).DataWidth=   800
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Operacion"
            Columns(3).DataField=   "operacioncodigo"
            Columns(3).DataWidth=   1200
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Cuenta"
            Columns(4).DataField=   "cuentacodigo"
            Columns(4).DataWidth=   800
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Debe/Haber"
            Columns(5).DataField=   "iddebeohaber"
            Columns(5).DataWidth=   800
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   4
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "Inafecto"
            Columns(6).DataField=   "plantillaasientoinafecto"
            Columns(6).DataWidth=   700
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(7)._VlistStyle=   0
            Columns(7)._MaxComboItems=   5
            Columns(7).Caption=   "Cuenta IGV"
            Columns(7).DataField=   "plantillaasientocuentaigv"
            Columns(7).DataWidth=   1100
            Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(8)._VlistStyle=   0
            Columns(8)._MaxComboItems=   5
            Columns(8).Caption=   "Valor IGV (%)"
            Columns(8).DataField=   "plantillaasientovalorigv"
            Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(9)._VlistStyle=   0
            Columns(9)._MaxComboItems=   5
            Columns(9).Caption=   "Comodin"
            Columns(9).DataField=   "plantillaasientocomodin"
            Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(10)._VlistStyle=   4
            Columns(10)._MaxComboItems=   5
            Columns(10).Caption=   "Cuenta Ajuste"
            Columns(10).DataField=   "plantillaasientoctaajuste"
            Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   11
            Splits(0)._UserFlags=   0
            Splits(0).RecordSelectorWidth=   508
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=11"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2731"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(5)=   "Column(1).Width=2731"
            Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2646"
            Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(9)=   "Column(2).Width=2731"
            Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=2646"
            Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(13)=   "Column(3).Width=2731"
            Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=2646"
            Splits(0)._ColumnProps(16)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(17)=   "Column(4).Width=2731"
            Splits(0)._ColumnProps(18)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(19)=   "Column(4)._WidthInPix=2646"
            Splits(0)._ColumnProps(20)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(21)=   "Column(5).Width=2731"
            Splits(0)._ColumnProps(22)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(23)=   "Column(5)._WidthInPix=2646"
            Splits(0)._ColumnProps(24)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(25)=   "Column(6).Width=2731"
            Splits(0)._ColumnProps(26)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(27)=   "Column(6)._WidthInPix=2646"
            Splits(0)._ColumnProps(28)=   "Column(6)._ColStyle=1"
            Splits(0)._ColumnProps(29)=   "Column(6).Order=7"
            Splits(0)._ColumnProps(30)=   "Column(7).Width=2731"
            Splits(0)._ColumnProps(31)=   "Column(7).DividerColor=0"
            Splits(0)._ColumnProps(32)=   "Column(7)._WidthInPix=2646"
            Splits(0)._ColumnProps(33)=   "Column(7).Order=8"
            Splits(0)._ColumnProps(34)=   "Column(8).Width=2731"
            Splits(0)._ColumnProps(35)=   "Column(8).DividerColor=0"
            Splits(0)._ColumnProps(36)=   "Column(8)._WidthInPix=2646"
            Splits(0)._ColumnProps(37)=   "Column(8).Order=9"
            Splits(0)._ColumnProps(38)=   "Column(9).Width=2731"
            Splits(0)._ColumnProps(39)=   "Column(9).DividerColor=0"
            Splits(0)._ColumnProps(40)=   "Column(9)._WidthInPix=2646"
            Splits(0)._ColumnProps(41)=   "Column(9).Order=10"
            Splits(0)._ColumnProps(42)=   "Column(10).Width=2053"
            Splits(0)._ColumnProps(43)=   "Column(10).DividerColor=0"
            Splits(0)._ColumnProps(44)=   "Column(10)._WidthInPix=1969"
            Splits(0)._ColumnProps(45)=   "Column(10)._ColStyle=1"
            Splits(0)._ColumnProps(46)=   "Column(10).Order=11"
            Splits.Count    =   1
            PrintInfos(0)._StateFlags=   3
            PrintInfos(0).Name=   "piInternal 0"
            PrintInfos(0).PageHeaderFont=   "Size=7.8,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageFooterFont=   "Size=7.8,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
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
            _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=2,.bold=0,.fontsize=825,.italic=0"
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
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=62,.parent=13"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=59,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=60,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=61,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=58,.parent=13"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=55,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=56,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=57,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=54,.parent=13"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=51,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=52,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=53,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=50,.parent=13"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=46,.parent=13"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=43,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=44,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=45,.parent=17"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=28,.parent=13"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=25,.parent=14"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=26,.parent=15"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=27,.parent=17"
            _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=32,.parent=13,.alignment=2"
            _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=29,.parent=14"
            _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=30,.parent=15"
            _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=31,.parent=17"
            _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=66,.parent=13"
            _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=63,.parent=14"
            _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=64,.parent=15"
            _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=65,.parent=17"
            _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=70,.parent=13"
            _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=67,.parent=14"
            _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=68,.parent=15"
            _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=69,.parent=17"
            _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=74,.parent=13"
            _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=71,.parent=14"
            _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=72,.parent=15"
            _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=73,.parent=17"
            _StyleDefs(76)  =   "Splits(0).Columns(10).Style:id=78,.parent=13,.alignment=2"
            _StyleDefs(77)  =   "Splits(0).Columns(10).HeadingStyle:id=75,.parent=14"
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
         Begin VB.CheckBox chk 
            Height          =   240
            Left            =   2505
            TabIndex        =   7
            Top             =   2085
            Width           =   525
         End
         Begin TextFer.TxFer txt 
            Height          =   315
            Index           =   0
            Left            =   2505
            TabIndex        =   2
            Top             =   750
            Width           =   4005
            _ExtentX        =   7070
            _ExtentY        =   550
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   ""
            ColorIlumina    =   -2147483624
            SaltarAlEnter   =   -1  'True
            Valor           =   ""
            NoCaracteres    =   "0123456789"
            MarcarTextoAlEnfoque=   -1  'True
            NoRangoCadena   =   -1  'True
         End
         Begin TextFer.TxFer txt 
            Height          =   300
            Index           =   1
            Left            =   2475
            TabIndex        =   5
            Top             =   1695
            Width           =   315
            _ExtentX        =   550
            _ExtentY        =   529
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   ""
            ColorIlumina    =   -2147483624
            SaltarAlEnter   =   -1  'True
            Valor           =   ""
            NoCaracteres    =   "DH"
            MarcarTextoAlEnfoque=   -1  'True
            NoRangoCadena   =   -1  'True
         End
         Begin TextFer.TxFer txt 
            Height          =   300
            Index           =   2
            Left            =   2475
            TabIndex        =   9
            Top             =   2655
            Width           =   1125
            _ExtentX        =   1990
            _ExtentY        =   529
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   ""
            ColorIlumina    =   -2147483624
            SaltarAlEnter   =   -1  'True
            Valor           =   ""
            TipoDato        =   1
            SignoNegativo   =   0   'False
            MarcarTextoAlEnfoque=   -1  'True
         End
         Begin TextFer.TxFer txt 
            Height          =   300
            Index           =   3
            Left            =   3900
            TabIndex        =   6
            Top             =   1695
            Width           =   2610
            _ExtentX        =   4614
            _ExtentY        =   529
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   ""
            ColorIlumina    =   -2147483624
            SaltarAlEnter   =   -1  'True
            Valor           =   ""
            NoCaracteres    =   "0123456789%"
            MarcarTextoAlEnfoque=   -1  'True
            NoRangoCadena   =   -1  'True
         End
         Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuda2 
            Height          =   495
            Left            =   2535
            TabIndex        =   1
            Top             =   450
            Width           =   3990
            _ExtentX        =   7049
            _ExtentY        =   868
            Enabled         =   0   'False
            XcodMaxLongitud =   4
            xcodwith        =   500
            NomTabla        =   "ct_subasiento"
            ListaCampos     =   "subasientocodigo(1),subasientodescripcion(1)"
            XcodCampo       =   "subasientocodigo"
            XListCampo      =   "subasientodescripcion"
            ListaCamposDescrip=   "Código,Descripción"
            ListaCamposText =   "subasientocodigo,subasientodescripcion"
         End
         Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuda1 
            Height          =   315
            Left            =   2535
            TabIndex        =   0
            Top             =   135
            Width           =   3990
            _ExtentX        =   7049
            _ExtentY        =   550
            Enabled         =   0   'False
            XcodMaxLongitud =   3
            xcodwith        =   500
            NomTabla        =   "ct_asiento"
            ListaCampos     =   "asientocodigo(1),asientodescripcion(1)"
            XcodCampo       =   "asientocodigo"
            XListCampo      =   "asientodescripcion"
            ListaCamposDescrip=   "Código,Descripción"
            ListaCamposText =   "asientocodigo,asientodescripcion"
         End
         Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuda4 
            Height          =   360
            Left            =   2505
            TabIndex        =   4
            Top             =   1365
            Width           =   4020
            _ExtentX        =   7091
            _ExtentY        =   635
            XcodMaxLongitud =   20
            xcodwith        =   800
            NomTabla        =   "ct_cuenta"
            ListaCampos     =   "cuentacodigo(1),cuentadescripcion(1)"
            XcodCampo       =   "cuentacodigo"
            XListCampo      =   "cuentadescripcion"
            ListaCamposDescrip=   "Cuenta,Descripción"
            ListaCamposText =   "cuentacodigo,cuentadescripcion"
         End
         Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuda3 
            Height          =   345
            Left            =   2505
            TabIndex        =   3
            Top             =   1065
            Width           =   4020
            _ExtentX        =   7091
            _ExtentY        =   614
            XcodMaxLongitud =   2
            xcodwith        =   500
            NomTabla        =   "ct_operacion"
            ListaCampos     =   "operacioncodigo(1),operaciondescripcion(1)"
            XcodCampo       =   "operacioncodigo"
            XListCampo      =   "operaciondescripcion"
            ListaCamposDescrip=   "Código,Descripción"
            ListaCamposText =   "operacioncodigo,operaciondescripcion"
         End
         Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuda5 
            Height          =   360
            Left            =   2490
            TabIndex        =   8
            Top             =   2340
            Width           =   4035
            _ExtentX        =   7112
            _ExtentY        =   635
            XcodMaxLongitud =   20
            xcodwith        =   800
            NomTabla        =   "ct_cuenta"
            ListaCampos     =   "cuentacodigo(1),cuentadescripcion(1)"
            XcodCampo       =   "cuentacodigo"
            XListCampo      =   "cuentadescripcion"
            ListaCamposDescrip=   "Cuenta,Descripción"
            ListaCamposText =   "cuentacodigo,cuentadescripcion"
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Cuenta de Ajuste "
            Height          =   195
            Index           =   7
            Left            =   4710
            TabIndex        =   33
            Top             =   2070
            Width           =   1260
         End
         Begin VB.Label Label5 
            Caption         =   "Cta Comodin"
            Height          =   300
            Left            =   2955
            TabIndex        =   31
            Top             =   1770
            Width           =   900
         End
         Begin VB.Label Label4 
            Caption         =   "IGV (%)"
            Height          =   270
            Left            =   135
            TabIndex        =   30
            Top             =   2790
            Width           =   1410
         End
         Begin VB.Label Label3 
            Caption         =   "Cuenta con IGV"
            Height          =   300
            Left            =   120
            TabIndex        =   29
            Top             =   2445
            Width           =   1350
         End
         Begin VB.Label lbl 
            Caption         =   "Cuenta Inafecta"
            Height          =   285
            Index           =   6
            Left            =   120
            TabIndex        =   20
            Top             =   2115
            Width           =   2280
         End
         Begin VB.Label lbl 
            Caption         =   "Debe/Haber"
            Height          =   285
            Index           =   5
            Left            =   135
            TabIndex        =   17
            Top             =   1785
            Width           =   2310
         End
         Begin VB.Label lbl 
            Caption         =   "Cuenta"
            Height          =   285
            Index           =   4
            Left            =   135
            TabIndex        =   16
            Top             =   1470
            Width           =   2280
         End
         Begin VB.Label lbl 
            Caption         =   "Operación"
            Height          =   285
            Index           =   3
            Left            =   135
            TabIndex        =   15
            Top             =   1155
            Width           =   2310
         End
         Begin VB.Label lbl 
            Caption         =   "Correlativo Plantilla"
            Height          =   285
            Index           =   2
            Left            =   135
            TabIndex        =   14
            Top             =   825
            Width           =   2310
         End
         Begin VB.Label lbl 
            Caption         =   "Sub Asiento"
            Height          =   285
            Index           =   1
            Left            =   135
            TabIndex        =   13
            Top             =   495
            Width           =   2310
         End
         Begin VB.Label lbl 
            Caption         =   "Asiento"
            Height          =   285
            Index           =   0
            Left            =   135
            TabIndex        =   12
            Top             =   180
            Width           =   2310
         End
      End
      Begin VB.Label lblNumRegSubAs 
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblNumReg"
         ForeColor       =   &H0000FFFF&
         Height          =   288
         Left            =   6912
         TabIndex        =   41
         Top             =   6864
         Width           =   840
      End
      Begin VB.Label Label1 
         Caption         =   "Nº SubAsientos x Asiento"
         Height          =   276
         Left            =   4920
         TabIndex        =   40
         Top             =   6876
         Width           =   1860
      End
      Begin VB.Label Label2 
         Caption         =   "Nº Asientos"
         Height          =   252
         Left            =   5892
         TabIndex        =   39
         Top             =   4140
         Width           =   948
      End
      Begin VB.Label lblNumRegAsientos 
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblNumReg"
         ForeColor       =   &H0080FFFF&
         Height          =   276
         Left            =   6780
         TabIndex        =   38
         Top             =   4128
         Width           =   912
      End
      Begin VB.Label lblMensaje 
         Alignment       =   2  'Center
         BackColor       =   &H80000016&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   300
         Left            =   -72315
         TabIndex        =   28
         Top             =   6780
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmConsultaComprobante"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim modoinsert As Boolean
Dim modoedit As Boolean
Dim mododelete As Boolean
Dim i_filaorigen As Integer
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim rsPlantilla As ADODB.Recordset
Dim dllgen As dllgeneral.dll_general

Private Sub ChkAjuste_Click()
    If modoinsert = True Or modoedit = True Then cAcepta.Enabled = ValidaDataIngreso()
End Sub

Private Sub Form_Load()
  Call ConfiguraForm
  Call MuestraDatos1
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set rs = Nothing
End Sub

Sub ConfiguraForm()
  SSTab1.Tab = 0
  Ctr_Ayuda1.conexion VGcnx
  Ctr_Ayuda2.conexion VGcnx
  Ctr_Ayuda3.conexion VGcnx
  Ctr_Ayuda4.conexion VGcnx
  Ctr_Ayuda4.Filtro = "cuentanivel=" & VGnumniveles & " or cuentacodigo='00'"
  Ctr_Ayuda5.conexion VGcnx
  Ctr_Ayuda5.Filtro = "cuentanivel=" & VGnumniveles & " or cuentacodigo='00'"
  cAcepta.Enabled = False
  lblNumRegAsientos.Caption = Empty
  lblNumRegSubAs.Caption = Empty
  Me.Width = 6870
  Me.Height = 7650
  Set dllgen = New dllgeneral.dll_general
End Sub

Sub MuestraDatos1()
  Dim SQL As String
   SQL = "SELECT ct_asiento.asientocodigo, ct_asiento.asientodescripcion "
   SQL = SQL & "FROM ct_asiento "
   SQL = SQL & "WHERE ct_asiento.asientocodigo<>'00' "
   SQL = SQL & "ORDER BY 1,2"
  
  Set rs = VGcnx.Execute(SQL)
  If rs.RecordCount <= 0 Then
        MsgBox "No se han Registrado los Asientos, Ingresar por opción de Asientos", vbInformation, Caption
        Exit Sub
  Else
    Set TDBGrid3.DataSource = rs
    Call ConfiguraTdbgrid
    lblNumRegAsientos.Caption = rs.RecordCount
    SSTab1.Tab = 0
    With TDBGrid3
        .Columns(0).Width = 1500
        .Columns(1).Width = 4500
    End With
 End If
End Sub

Sub MuestraDatos2()
  Dim SQL As String
    SQL = "SELECT ct_subasiento.subasientocodigo, ct_subasiento.subasientodescripcion "
    SQL = SQL & "FROM ct_subasiento "
    SQL = SQL & "WHERE ct_subasiento.asientocodigo='" & TDBGrid3.Columns(0).Text & "' AND ct_subasiento.subasientocodigo<>'00' "
    SQL = SQL & "ORDER BY 1,2"
    
    Set rs1 = VGcnx.Execute(SQL)
    Set TDBGrid1.DataSource = rs1
    lblNumRegSubAs.Caption = rs1.RecordCount
    SSTab1.Tab = 0
    
    With TDBGrid1
      .Columns(0).Width = 1500
      .Columns(1).Width = 4600
    End With

End Sub

Private Sub cCancela_Click()
  frmbotones.Visible = True
  modoinsert = False:  modoedit = False:  mododelete = False: lblMensaje.Caption = Empty
  SSTab1.Tab = 0
  i_filaorigen = -1
End Sub

Private Sub cmdBotones_Click(Index As Integer)
  Dim J As Integer
  Dim spos As Integer
  Dim SQL As String
  
  'On Error GoTo X
  SSTab1.TabEnabled(1) = True
  
  Select Case Index
     Case 0   'nuevo
        frmbotones.Visible = False
        modoinsert = True
        lblMensaje.Caption = "Nuevo"
        Call ModoPlantilla(True)
        Ctr_Ayuda1.xclave = TDBGrid3.Columns(0).Value: Ctr_Ayuda1.Ejecutar
        Ctr_Ayuda2.xclave = TDBGrid1.Columns(0).Value: Ctr_Ayuda2.Ejecutar
        txt(0).SetFocus
        Call LimpiarValores
        
     Case 1   'modificar
        If TDBGrid1.Row < 0 Then
          Exit Sub
        End If
        frmbotones.Visible = False
        modoedit = True
        lblMensaje.Caption = "Editar"
        Call ModoPlantilla(True)
        txt(0).Enabled = False
      
     Case 2   'eliminar
        If MsgBox("Desea eliminar el registro con Correlativo Nº " & CInt(TDBGrid2.Columns(2).Text) & " ?", vbYesNo + vbDefaultButton2, "AVISO") = vbYes Then
           SQL = "DELETE FROM CT_PLANTILLAASIENTO WHERE subasientocodigo = '" & Trim(Ctr_Ayuda2.xclave) & "' AND asientocodigo = '" & Trim(Ctr_Ayuda1.xclave) & "' AND plantillaasientocorrela =" & CInt(TDBGrid2.Columns(2).Text)
           VGcnx.Execute (SQL)
           Call MuestraGrid2
        End If
        
     Case 3   'imprimir
       Call Impresion
     
     Case 4  ' salir
       Unload Me
  End Select
  Exit Sub
   
X:
   If Err Then
      Err = 0
      Resume Next
   End If
   
End Sub

Sub EditarValores()
 Dim i As Integer
  With TDBGrid2
    txt(0).Text = Trim(.Columns(2).Text)
    Ctr_Ayuda3.xclave = Trim(.Columns(3).Text): Ctr_Ayuda3.Ejecutar
    Ctr_Ayuda4.xclave = Trim(.Columns(4).Text): Ctr_Ayuda4.Ejecutar
    Ctr_Ayuda5.xclave = Trim(.Columns(7).Text): Ctr_Ayuda5.Ejecutar
    txt(1).Text = Trim(.Columns(5).Text)
    chk.Value = IIf((.Columns(6).Text) = 0, 0, 1)
    ChkAjuste.Value = IIf((.Columns(10).Text) = 0, 0, 1)
    txt(2).Text = .Columns(8).Text
    txt(3).Text = .Columns(9).Text
  End With
  If modoinsert = True Then Call LimpiarValores
  
End Sub

Public Function LimpiarValores()
 Dim i As Integer
  Ctr_Ayuda3.xclave = Empty: Ctr_Ayuda3.Ejecutar
  Ctr_Ayuda4.xclave = Empty: Ctr_Ayuda4.Ejecutar
  Ctr_Ayuda5.xclave = Empty: Ctr_Ayuda5.Ejecutar
  For i = 0 To 3
     txt(i).Text = Empty
  Next
  chk.Value = 0
  
End Function

Private Sub txt_Change(Index As Integer)
     If modoinsert = True Or modoedit = True Then cAcepta.Enabled = ValidaDataIngreso()
End Sub

Private Sub Ctr_Ayuda3_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
     If modoinsert = True Or modoedit = True Then cAcepta.Enabled = ValidaDataIngreso()
End Sub

Private Sub Ctr_Ayuda4_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
    If modoinsert = True Or modoedit = True Then
       cAcepta.Enabled = ValidaDataIngreso()
       If Ctr_Ayuda4.xclave = "00" Then
          txt(3).Enabled = True
          txt(3).SetFocus
       Else
          txt(3).Enabled = False
       End If
    End If
    
End Sub

Private Sub Ctr_Ayuda4_LostFocus()
  If modoinsert = True Or modoedit = True Then
     If Ctr_Ayuda4.xclave = "00" Then
         txt(3).Enabled = True
         txt(3).SetFocus
     Else
         txt(3).Enabled = False
     End If
  End If
End Sub

Private Sub Ctr_Ayuda4_KeyPress(KeyAscii As Integer)
 If modoinsert = True Or modoedit = True Then
    If KeyAscii = 13 Then
      If Ctr_Ayuda4.xclave = "00" Then
        txt(3).SetFocus
      End If
    End If
 End If
End Sub

Private Sub chk_Click()
    If modoinsert = True Or modoedit = True Then cAcepta.Enabled = ValidaDataIngreso()
    Select Case chk.Value
        Case 0:
            Ctr_Ayuda5.Enabled = True
            txt(2).Enabled = True
            txt(2).BackColor = &H80000005
        Case 1:
            Ctr_Ayuda5.Enabled = False
            txt(2).Enabled = False
            txt(2).BackColor = ColorDesHabilitado
    End Select
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = 13 And Index = 1 Then
    chk.SetFocus
  End If
End Sub

Private Sub cAcepta_Click()
 If ValidaData = True Then
    Call GrabarData
    Call MuestraGrid2
    Call LimpiarValores
    frmbotones.Visible = True
    cAcepta.Enabled = False
    modoinsert = False:  modoedit = False:  mododelete = False: lblMensaje.Caption = Empty
    Call LimpiarValores
    Call ModoPlantilla(False)
 End If
End Sub

Private Sub Ctr_Ayuda2_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
  Call MuestraGrid2
End Sub

Sub MuestraGrid2()
    Dim SQL As String
    SQL = "SELECT asientocodigo,subasientocodigo,plantillaasientocorrela,operacioncodigo,cuentacodigo,iddebeohaber,plantillaasientoinafecto,plantillaasientocuentaigv,plantillaasientovalorigv,plantillaasientocomodin,plantillaasientoctaajuste "
    SQL = SQL & "FROM ct_plantillaasiento WHERE asientocodigo='" & Ctr_Ayuda1.xclave & "' AND subasientocodigo='" & Ctr_Ayuda2.xclave & "' "
    SQL = SQL & "ORDER BY 3"
    Set rsPlantilla = New ADODB.Recordset
    Set rsPlantilla = VGcnx.Execute(SQL)
    Set TDBGrid2.DataSource = rsPlantilla
    Call ConfigGrid2
End Sub

Sub ModoPlantilla(Flag_Normal As Boolean)
  txt(0).Enabled = Flag_Normal
  Ctr_Ayuda3.Enabled = Flag_Normal
  Ctr_Ayuda4.Enabled = Flag_Normal
  Ctr_Ayuda5.Enabled = Flag_Normal
  txt(1).Enabled = Flag_Normal
  chk.Enabled = Flag_Normal
  ChkAjuste.Enabled = Flag_Normal
  txt(2).Enabled = Flag_Normal
  txt(3).Enabled = Flag_Normal
  
End Sub

Sub ConfigGrid2()
    With TDBGrid2
       .Columns(0).Visible = False
       .Columns(1).Visible = False
       .Columns(2).Width = 1000
       .Columns(3).Width = 1100
       .Columns(4).Width = 1000
       .Columns(5).Width = 1100
       .Columns(6).Width = 820
       .Columns(7).Width = 1100
       .Columns(8).Visible = False
    End With
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    Set VGvardllgen = New dllgeneral.dll_general
    Ctr_Ayuda1.xclave = TDBGrid3.Columns(0).Value: Ctr_Ayuda1.Ejecutar
    Ctr_Ayuda2.Filtro = "asientocodigo='" & Ctr_Ayuda1.xclave & "'"
    Ctr_Ayuda2.xclave = VGvardllgen.ESNULO(TDBGrid1.Columns(0).Value, ""): Ctr_Ayuda2.Ejecutar
    Call LimpiarValores
    Call ModoPlantilla(False)
End Sub

Private Sub TDBGrid1_HeadClick(ByVal ColIndex As Integer)
On Error Resume Next
  If rs.Sort = Empty Then
      rs.Sort = TDBGrid1.Columns.Item(ColIndex).DataField & " asc"
   ElseIf Right(rs.Sort, 3) = "asc" Then
      rs.Sort = TDBGrid1.Columns.Item(ColIndex).DataField & " desc"
   ElseIf Right(rs.Sort, 4) = "desc" Then
      rs.Sort = TDBGrid1.Columns.Item(ColIndex).DataField & " asc"
  End If
  Call ConfiguraTdbgrid
  TDBGrid1.Refresh
End Sub

Private Sub TDBGrid1_DblClick()
 If ValidaSubAsientos(TDBGrid3.Columns(0).Text) = True Then
    Call SSTab1_Click(0)
    SSTab1.Tab = 1
 End If
End Sub

Private Sub ConfiguraTdbgrid()
    With TDBGrid3
        .Columns(0).Width = 1100
        .Columns(1).Width = 3500
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
  
  If Ctr_Ayuda3.xclave = Empty Then
    ValidaDataIngreso = False
    Exit Function
  End If
  
  If Ctr_Ayuda4.xclave = Empty Then
    ValidaDataIngreso = False
    Exit Function
  End If
  
  'For i = 0 To 1
  ' If txt(i).Text = Empty Then
  '   ValidaDataIngreso = False
  '   Exit Function
  ' End If
  'Next
  
  ValidaDataIngreso = True
End Function

Private Sub TDBGrid2_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If rsPlantilla.RecordCount = 0 Then Exit Sub
    If modoinsert = False Or modoedit = False Then Call EditarValores
End Sub

Private Sub TDBGrid3_Click()
'If rs.RecordCount > 0 Then
'   Call MuestraDatos2
' End If
End Sub

Private Sub TDBGrid3_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
 If rs.RecordCount > 0 Then
   Call MuestraDatos2
   SSTab1.TabEnabled(1) = (rs1.RecordCount > 0)
 End If
End Sub

Function ValidaData() As Boolean
  Dim SQL As String
    SQL = "SELECT cuentacodigo FROM ct_plantillaasiento WHERE asientocodigo='" & Trim(Ctr_Ayuda1.xclave) & "' AND "
    SQL = SQL & "subasientocodigo='" & Trim(Ctr_Ayuda2.xclave) & "' AND cuentacodigo='" & Trim(Ctr_Ayuda4.xclave) & "'"
    
    If modoinsert = True And dllgen.VerificaDatoExistente(VGcnx, SQL) > 0 Then
      MsgBox "La cuenta se va a duplicar en la Plantilla Actual"
      ValidaData = True
      'Deja pasar esta validación
    End If
    
    SQL = "SELECT cuentacodigo FROM ct_plantillaasiento WHERE asientocodigo='" & Trim(Ctr_Ayuda1.xclave) & "' AND "
    SQL = SQL & "subasientocodigo='" & Trim(Ctr_Ayuda2.xclave) & "' AND "
    SQL = SQL & "plantillaasientocorrela=" & txt(0).Text
    If modoinsert = True And dllgen.VerificaDatoExistente(VGcnx, SQL) > 0 Then
      MsgBox "El Correlativo Nº " & txt(0).Text & " existe en la Plantilla actual", vbInformation, Caption
      txt(0).SetFocus
      ValidaData = False
      Exit Function
    End If
    
    Set VGvardllgen = New dllgeneral.dll_general
    If Ctr_Ayuda5.xclave <> Empty And (CLng(VGvardllgen.ESNULO(txt(2).Text, 0)) <= 0) = True Then
      MsgBox "Falta la Cuenta con IGV  ó Valor de IGV(%)", vbInformation, Caption
      Ctr_Ayuda5.SetFocus
      ValidaData = False
      Exit Function
    End If
    
    If txt(3).Text <> Empty Then
       If Right(txt(3).Text, 1) <> "%" Then
           MsgBox "La Cuenta Comodín debe terminar con un Caracter (%)", vbInformation, Caption
           ValidaData = False
           txt(3).Text = txt(3).Text & "%"
           Exit Function
       End If
        
       If VerificaCriterioComodin(txt(3).Text) = False Then
           ValidaData = False
           txt(3).SetFocus
           Exit Function
       End If
    End If
    
    ValidaData = True
End Function

Sub GrabarData()
 On Error GoTo X
  
  Dim SQL As String
  If modoinsert = True Then
        SQL = "INSERT INTO ct_plantillaasiento (subasientocodigo,asientocodigo,plantillaasientocorrela,operacioncodigo,"
        SQL = SQL & "cuentacodigo,iddebeohaber,plantillaasientoinafecto,plantillaasientocuentaigv,plantillaasientovalorigv,plantillaasientocomodin,plantillaasientoctaajuste,usuariocodigo,fechaact) "
        SQL = SQL & "VALUES ('" & Ctr_Ayuda2.xclave & "','" & Ctr_Ayuda1.xclave & "','" & txt(0).Text & "','"
        SQL = SQL & Ctr_Ayuda3.xclave & "','" & Ctr_Ayuda4.xclave & "','" & UCase(txt(1).Text) & "'," & chk.Value & ",'" & Ctr_Ayuda5.xclave & "',"
        SQL = SQL & VGvardllgen.ESNULO(txt(2).Text, 0) & ",'" & txt(3).Text & "'," & ChkAjuste.Value & ",'" & VGusuario & "','" & Format(Date, "dd/mm/yyyy") & "')"
                  
  ElseIf modoedit = True Then
        SQL = "UPDATE ct_plantillaasiento SET operacioncodigo='" & Ctr_Ayuda3.xclave & "',cuentacodigo='" & Ctr_Ayuda4.xclave & "',"
        SQL = SQL & "iddebeohaber='" & txt(1).Text & "',plantillaasientoinafecto=" & chk.Value & ",plantillaasientocuentaigv='" & Ctr_Ayuda5.xclave & "',"
        SQL = SQL & "plantillaasientovalorigv=" & VGvardllgen.ESNULO(txt(2).Text, 0) & ",plantillaasientocomodin='" & txt(3).Text & "',plantillaasientoctaajuste=" & ChkAjuste.Value & ",usuariocodigo='" & VGusuario & "',fechaact='" & Format(Date, "dd/mm/yyyy") & "'"
        SQL = SQL & "WHERE subasientocodigo='" & Trim(Ctr_Ayuda2.xclave) & "' AND asientocodigo = '" & Trim(Ctr_Ayuda1.xclave) & "' AND plantillaasientocorrela =" & UCase(txt(0).Text)
  End If
    
  VGcnx.BeginTrans
  VGcnx.Execute (SQL)
  VGcnx.CommitTrans
  
  Exit Sub

X:
  If Err.Number = -2147217873 Then
    MsgBox "Esta intentando registrar uno Existente " & Err.Description, vbInformation, Caption
  Else
    MsgBox "Error inesperado: " & Err.Number & " " & Err.Description
  End If
  VGcnx.RollbackTrans

End Sub

Private Sub txt_LostFocus(Index As Integer)
  txt(Index).Text = UCase(txt(Index).Text)
End Sub

Private Sub Ctr_Ayuda5_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
   cAcepta.Enabled = ValidaDataIngreso()
End Sub

Private Sub Ctr_Ayuda5_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
      cAcepta.SetFocus
  End If
End Sub

Private Function VerificaCriterioComodin(cad As String) As Boolean
  Dim pos As Integer
  Dim SQL As String
  Dim valor As String
  Dim flag As Boolean
    pos = 1
    flag = True
    Set VGvardllgen = New dllgeneral.dll_general
    Do While pos <> 0
        pos = InStr(1, cad, "%", vbTextCompare)
        If pos > 0 Then
          valor = Left(cad, pos - 1)
          SQL = "SELECT cuentacodigo FROM ct_cuenta WHERE cuentacodigo='" & valor & "'"
          If VGvardllgen.VerificaDatoExistente(VGcnx, SQL) <= 0 Then
              MsgBox "La Cuenta Contable Nº " & valor & " no Existe en el Plan de Cuentas", vbInformation, Caption
              flag = False
              Exit Do
          End If
          cad = Right(cad, (Len(cad) - pos))
        End If
    Loop
    VerificaCriterioComodin = flag
    
End Function

Sub Impresion()
 Dim arrparm(4) As Variant
 Dim arrform(2) As Variant
    arrparm(0) = VGParamSistem.BDEmpresa
    arrparm(1) = VGParamSistem.Anoproceso
    arrparm(2) = IIf(Ctr_Ayuda1.xclave = Empty, "%%", Ctr_Ayuda1.xclave)
    arrparm(3) = IIf(Ctr_Ayuda2.xclave = Empty, "%%", Ctr_Ayuda2.xclave)
    Set VGvardllgen = New dllgeneral.dll_general
    arrform(0) = "@TituloReporte='" & "Plantilla de SubAsientos - Asiento: " & Ctr_Ayuda1.xclave & " " & Ctr_Ayuda1.xnombre & "'"
    arrform(1) = "@Mes='" & VGvardllgen.DESMES(VGParamSistem.Mesproceso) & "'"
    Call ImpresionRptProc("rptPlantillaSubAsiento.rpt", arrform, arrparm)

End Sub
