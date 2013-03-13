VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmPedidoxGuias 
   Caption         =   "Pedidos x Guias"
   ClientHeight    =   7230
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10845
   LinkTopic       =   "Form2"
   ScaleHeight     =   7230
   ScaleWidth      =   10845
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab SSTab1 
      Height          =   6732
      Left            =   0
      TabIndex        =   0
      Top             =   48
      Width           =   10644
      _ExtentX        =   18785
      _ExtentY        =   11880
      _Version        =   393216
      TabHeight       =   420
      TabCaption(0)   =   "Guias de remision"
      TabPicture(0)   =   "FrmPedidoxGuias.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Fr1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Fr2(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "CmdGrabardata"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Detalle de Guias"
      TabPicture(1)   =   "FrmPedidoxGuias.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame6"
      Tab(1).Control(1)=   "Frame5(1)"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Totales de Guias"
      TabPicture(2)   =   "FrmPedidoxGuias.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame1"
      Tab(2).Control(1)=   "Frame5(2)"
      Tab(2).ControlCount=   2
      Begin VB.CommandButton CmdGrabardata 
         Caption         =   "Grabar"
         Height          =   492
         Left            =   3792
         TabIndex        =   48
         Top             =   7440
         Width           =   1068
      End
      Begin VB.Frame Fr2 
         Height          =   645
         Index           =   0
         Left            =   132
         TabIndex        =   43
         Top             =   480
         Width           =   9204
         Begin VB.TextBox aBusca 
            Height          =   285
            Index           =   1
            Left            =   7032
            MaxLength       =   8
            TabIndex        =   46
            Top             =   210
            Width           =   885
         End
         Begin VB.TextBox aBusca 
            Height          =   285
            Index           =   0
            Left            =   6408
            MaxLength       =   3
            TabIndex        =   45
            Top             =   210
            Width           =   435
         End
         Begin VB.CommandButton cBusca 
            BackColor       =   &H80000008&
            Caption         =   "&Buscar"
            Height          =   375
            Left            =   7968
            TabIndex        =   44
            Top             =   180
            Width           =   1095
         End
         Begin VB.Label Label6 
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Left            =   6912
            TabIndex        =   47
            Top             =   240
            Width           =   192
         End
      End
      Begin VB.Frame Fr1 
         Height          =   6960
         Left            =   96
         TabIndex        =   35
         Top             =   1104
         Width           =   10416
         Begin VB.CommandButton CmdNuevo 
            Caption         =   "Nuevo"
            Height          =   588
            Left            =   9024
            TabIndex        =   51
            Top             =   3360
            Width           =   972
         End
         Begin VB.CommandButton CmdSalir 
            Caption         =   "Salir"
            Height          =   588
            Left            =   9024
            TabIndex        =   50
            Top             =   4848
            Width           =   972
         End
         Begin VB.CommandButton Cmdgrabar 
            Caption         =   "Grabar"
            Height          =   636
            Left            =   9024
            TabIndex        =   49
            Top             =   4080
            Width           =   972
         End
         Begin VB.Frame Frame5 
            Height          =   585
            Index           =   0
            Left            =   7644
            TabIndex        =   36
            Top             =   2532
            Width           =   2265
            Begin MSMask.MaskEdBox totreg 
               Height          =   372
               Index           =   0
               Left            =   1104
               TabIndex        =   37
               Top             =   144
               Width           =   1080
               _ExtentX        =   1905
               _ExtentY        =   635
               _Version        =   393216
               Appearance      =   0
               BackColor       =   12648447
               ForeColor       =   16777215
               Enabled         =   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               PromptChar      =   "_"
            End
            Begin VB.Label Label3 
               Caption         =   "Total Reg."
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   228
               Index           =   0
               Left            =   156
               TabIndex        =   38
               Top             =   192
               Width           =   1032
            End
         End
         Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
            Height          =   2256
            Left            =   144
            TabIndex        =   39
            Top             =   384
            Width           =   9792
            _ExtentX        =   17277
            _ExtentY        =   3969
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
            Columns.Count   =   6
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=6"
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
            Splits(0)._ColumnProps(11)=   "Column(2).Width=3043"
            Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=2963"
            Splits(0)._ColumnProps(14)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(16)=   "Column(3).Width=3043"
            Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=2963"
            Splits(0)._ColumnProps(19)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(21)=   "Column(4).Width=3043"
            Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=2963"
            Splits(0)._ColumnProps(24)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(26)=   "Column(5).Width=3043"
            Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=2963"
            Splits(0)._ColumnProps(29)=   "Column(5)._EditAlways=0"
            Splits(0)._ColumnProps(30)=   "Column(5).Order=6"
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
            _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
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
            _StyleDefs(60)  =   "Named:id=33:Normal"
            _StyleDefs(61)  =   ":id=33,.parent=0"
            _StyleDefs(62)  =   "Named:id=34:Heading"
            _StyleDefs(63)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(64)  =   ":id=34,.wraptext=-1"
            _StyleDefs(65)  =   "Named:id=35:Footing"
            _StyleDefs(66)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(67)  =   "Named:id=36:Selected"
            _StyleDefs(68)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(69)  =   "Named:id=37:Caption"
            _StyleDefs(70)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(71)  =   "Named:id=38:HighlightRow"
            _StyleDefs(72)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(73)  =   "Named:id=39:EvenRow"
            _StyleDefs(74)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(75)  =   "Named:id=40:OddRow"
            _StyleDefs(76)  =   ":id=40,.parent=33"
            _StyleDefs(77)  =   "Named:id=41:RecordSelector"
            _StyleDefs(78)  =   ":id=41,.parent=34"
            _StyleDefs(79)  =   "Named:id=42:FilterBar"
            _StyleDefs(80)  =   ":id=42,.parent=33"
         End
         Begin TrueOleDBGrid70.TDBGrid TDBGrid2 
            Height          =   2160
            Left            =   144
            TabIndex        =   40
            Top             =   3360
            Width           =   8640
            _ExtentX        =   15240
            _ExtentY        =   3810
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
            Splits(0)._ColumnProps(11)=   "Column(2).Width=3043"
            Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=2963"
            Splits(0)._ColumnProps(14)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(16)=   "Column(3).Width=3043"
            Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=2963"
            Splits(0)._ColumnProps(19)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(21)=   "Column(4).Width=3043"
            Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=2963"
            Splits(0)._ColumnProps(24)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(26)=   "Column(5).Width=3043"
            Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=2963"
            Splits(0)._ColumnProps(29)=   "Column(5)._EditAlways=0"
            Splits(0)._ColumnProps(30)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(31)=   "Column(6).Width=3043"
            Splits(0)._ColumnProps(32)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(33)=   "Column(6)._WidthInPix=2963"
            Splits(0)._ColumnProps(34)=   "Column(6)._EditAlways=0"
            Splits(0)._ColumnProps(35)=   "Column(6).Order=7"
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
            _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
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
            Alignment       =   2  'Center
            Caption         =   "GENERACION DE PEDIDOS"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   348
            Index           =   0
            Left            =   0
            TabIndex        =   42
            Top             =   3024
            Width           =   9528
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   "Guia de Remisiom"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   348
            Left            =   240
            TabIndex        =   41
            Top             =   96
            Width           =   9528
         End
      End
      Begin VB.Frame Frame6 
         Height          =   972
         Left            =   -74904
         TabIndex        =   26
         Top             =   336
         Width           =   10284
         Begin VB.Label almacendescr 
            BackColor       =   &H00800000&
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
            ForeColor       =   &H00C0FFC0&
            Height          =   288
            Index           =   1
            Left            =   7488
            TabIndex        =   34
            Top             =   180
            Width           =   2664
         End
         Begin VB.Label clienterazon 
            BackColor       =   &H00800000&
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
            ForeColor       =   &H00C0FFC0&
            Height          =   288
            Index           =   1
            Left            =   1440
            TabIndex        =   33
            Top             =   576
            Width           =   8724
         End
         Begin VB.Label umerodoc 
            BackColor       =   &H00800000&
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
            ForeColor       =   &H00C0FFC0&
            Height          =   288
            Index           =   0
            Left            =   4272
            TabIndex        =   32
            Top             =   240
            Width           =   2088
         End
         Begin VB.Label fechadoc 
            BackColor       =   &H00800000&
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
            ForeColor       =   &H00C0FFC0&
            Height          =   285
            Index           =   0
            Left            =   1440
            TabIndex        =   31
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "Almacen"
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
            Height          =   228
            Index           =   9
            Left            =   6384
            TabIndex        =   30
            Top             =   240
            Width           =   768
         End
         Begin VB.Label Label1 
            Caption         =   "Cliente"
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
            Height          =   225
            Index           =   10
            Left            =   330
            TabIndex        =   29
            Top             =   600
            Width           =   765
         End
         Begin VB.Label Label1 
            Caption         =   "No.Documento"
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
            Height          =   225
            Index           =   11
            Left            =   2880
            TabIndex        =   28
            Top             =   270
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "Fecha Ped"
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
            Height          =   225
            Index           =   12
            Left            =   330
            TabIndex        =   27
            Top             =   270
            Width           =   975
         End
      End
      Begin VB.Frame Frame5 
         Height          =   5340
         Index           =   1
         Left            =   -74904
         TabIndex        =   18
         Top             =   1176
         Width           =   10392
         Begin VB.Frame Fr2 
            BackColor       =   &H00800000&
            BorderStyle     =   0  'None
            Height          =   765
            Index           =   1
            Left            =   5316
            TabIndex        =   22
            Top             =   4344
            Width           =   2055
            Begin MSMask.MaskEdBox totreg 
               Height          =   375
               Index           =   1
               Left            =   300
               TabIndex        =   23
               Top             =   90
               Width           =   1605
               _ExtentX        =   2831
               _ExtentY        =   661
               _Version        =   393216
               Appearance      =   0
               BackColor       =   12648447
               ForeColor       =   16777215
               Enabled         =   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               PromptChar      =   "_"
            End
            Begin VB.Label Label3 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Cantidad"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H0080FF80&
               Height          =   255
               Index           =   5
               Left            =   360
               TabIndex        =   24
               Top             =   495
               Width           =   1335
            End
         End
         Begin VB.Frame Frame4 
            Height          =   930
            Left            =   7968
            TabIndex        =   19
            Top             =   4176
            Width           =   2010
            Begin VB.CommandButton cmdBotones 
               Caption         =   "&Acepta"
               Height          =   690
               Index           =   11
               Left            =   90
               Picture         =   "FrmPedidoxGuias.frx":0054
               Style           =   1  'Graphical
               TabIndex        =   21
               Top             =   180
               Width           =   870
            End
            Begin VB.CommandButton cmdBotones 
               Caption         =   "&Cancelar"
               Height          =   690
               Index           =   12
               Left            =   1140
               Picture         =   "FrmPedidoxGuias.frx":0496
               Style           =   1  'Graphical
               TabIndex        =   20
               Top             =   180
               Width           =   825
            End
         End
         Begin TrueOleDBGrid70.TDBGrid TDBGrid3 
            Height          =   3468
            Left            =   156
            TabIndex        =   25
            Top             =   276
            Width           =   10056
            _ExtentX        =   17727
            _ExtentY        =   6112
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
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=2"
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
            Splits.Count    =   1
            PrintInfos(0)._StateFlags=   3
            PrintInfos(0).Name=   "piInternal 0"
            PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageHeaderHeight=   0
            PrintInfos(0).PageFooterHeight=   0
            PrintInfos.Count=   1
            AllowUpdate     =   0   'False
            DefColWidth     =   0
            HeadLines       =   1
            FootLines       =   1
            MultipleLines   =   0
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
            _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HC0FFFF&,.bold=0,.fontsize=780"
            _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
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
      Begin VB.Frame Frame1 
         Height          =   1452
         Left            =   -75000
         TabIndex        =   9
         Top             =   336
         Width           =   10284
         Begin VB.TextBox Text1 
            Enabled         =   0   'False
            Height          =   288
            Index           =   1
            Left            =   4560
            TabIndex        =   55
            Text            =   "0"
            Top             =   912
            Width           =   972
         End
         Begin VB.TextBox Text1 
            Enabled         =   0   'False
            Height          =   288
            Index           =   0
            Left            =   1776
            TabIndex        =   54
            Text            =   "1"
            Top             =   912
            Width           =   252
         End
         Begin VB.Label Label1 
            Caption         =   "Precio Unitario General"
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
            Height          =   228
            Index           =   6
            Left            =   2544
            TabIndex        =   53
            Top             =   996
            Width           =   1956
         End
         Begin VB.Label Label1 
            Caption         =   "Lista de Precios"
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
            Height          =   228
            Index           =   5
            Left            =   192
            TabIndex        =   52
            Top             =   996
            Width           =   1524
         End
         Begin VB.Label Label1 
            Caption         =   "Fecha Ped"
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
            Height          =   225
            Index           =   1
            Left            =   330
            TabIndex        =   17
            Top             =   270
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "No.Documento"
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
            Height          =   225
            Index           =   2
            Left            =   2880
            TabIndex        =   16
            Top             =   270
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "Cliente"
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
            Height          =   225
            Index           =   3
            Left            =   330
            TabIndex        =   15
            Top             =   600
            Width           =   765
         End
         Begin VB.Label Label1 
            Caption         =   "Almacen"
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
            Height          =   228
            Index           =   4
            Left            =   6384
            TabIndex        =   14
            Top             =   240
            Width           =   768
         End
         Begin VB.Label fechadoc 
            BackColor       =   &H00800000&
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
            ForeColor       =   &H00C0FFC0&
            Height          =   285
            Index           =   1
            Left            =   1440
            TabIndex        =   13
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label numerodoc 
            BackColor       =   &H00800000&
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
            ForeColor       =   &H00C0FFC0&
            Height          =   288
            Index           =   1
            Left            =   4272
            TabIndex        =   12
            Top             =   240
            Width           =   2088
         End
         Begin VB.Label clienterazon 
            BackColor       =   &H00800000&
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
            ForeColor       =   &H00C0FFC0&
            Height          =   288
            Index           =   0
            Left            =   1440
            TabIndex        =   11
            Top             =   576
            Width           =   8724
         End
         Begin VB.Label almacendescr 
            BackColor       =   &H00800000&
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
            ForeColor       =   &H00C0FFC0&
            Height          =   288
            Index           =   0
            Left            =   7488
            TabIndex        =   10
            Top             =   180
            Width           =   2664
         End
      End
      Begin VB.Frame Frame5 
         Height          =   5340
         Index           =   2
         Left            =   -75000
         TabIndex        =   1
         Top             =   1200
         Width           =   10392
         Begin VB.Frame Frame2 
            Height          =   930
            Left            =   7968
            TabIndex        =   5
            Top             =   4176
            Width           =   2010
            Begin VB.CommandButton cmdSalirFinal 
               Caption         =   "&Cancelar"
               Height          =   690
               Index           =   0
               Left            =   1140
               Picture         =   "FrmPedidoxGuias.frx":08D8
               Style           =   1  'Graphical
               TabIndex        =   7
               Top             =   180
               Width           =   825
            End
            Begin VB.CommandButton cmdGrabaFinal 
               Caption         =   "&Acepta"
               Height          =   690
               Left            =   90
               Picture         =   "FrmPedidoxGuias.frx":0D1A
               Style           =   1  'Graphical
               TabIndex        =   6
               Top             =   180
               Width           =   870
            End
         End
         Begin VB.Frame Fr2 
            BackColor       =   &H00800000&
            BorderStyle     =   0  'None
            Height          =   765
            Index           =   2
            Left            =   5316
            TabIndex        =   2
            Top             =   4344
            Width           =   2055
            Begin MSMask.MaskEdBox totreg 
               Height          =   375
               Index           =   2
               Left            =   300
               TabIndex        =   3
               Top             =   90
               Width           =   1605
               _ExtentX        =   2831
               _ExtentY        =   661
               _Version        =   393216
               Appearance      =   0
               BackColor       =   12648447
               ForeColor       =   16777215
               Enabled         =   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               PromptChar      =   "_"
            End
            Begin VB.Label Label3 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Cantidad"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H0080FF80&
               Height          =   255
               Index           =   1
               Left            =   360
               TabIndex        =   4
               Top             =   495
               Width           =   1335
            End
         End
         Begin TrueOleDBGrid70.TDBGrid TDBGrid4 
            Height          =   3468
            Left            =   156
            TabIndex        =   8
            Top             =   708
            Width           =   10056
            _ExtentX        =   17727
            _ExtentY        =   6112
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
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=2"
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
            Splits.Count    =   1
            PrintInfos(0)._StateFlags=   3
            PrintInfos(0).Name=   "piInternal 0"
            PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageHeaderHeight=   0
            PrintInfos(0).PageFooterHeight=   0
            PrintInfos.Count=   1
            AllowUpdate     =   0   'False
            DefColWidth     =   0
            HeadLines       =   1
            FootLines       =   1
            MultipleLines   =   0
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
            _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HC0FFFF&,.bold=0,.fontsize=780"
            _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
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
   End
End
Attribute VB_Name = "FrmPedidoxGuias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsacumula As New ADODB.Recordset
Dim rsdeta As New ADODB.Recordset
Dim csql As New ADODB.Recordset
Dim SQL As New ADODB.Recordset
Dim adll As New dllgeneral.dll_general
Dim dllgeneral As New dllgeneral.dll_general
Dim vt_tempo As String, vt_tempo1 As String
Dim xsql, xalma, xtipo, xnumero As String

Dim acepta As Integer
Dim nLongicampo(1) As Integer
Private Sub aBusca_KeyPress(Index As Integer, KeyAscii As Integer)
  Dim ldato As String
  If KeyAscii = 13 And Index = 1 Then
     TDBGrid1.ClearFields
     Set TDBGrid1.DataSource = Nothing
     aBusca(0) = Right("0000000000" & Trim(aBusca(0)), aBusca(0).MaxLength)
     aBusca(1) = Right("0000000000" & Trim(aBusca(1)), aBusca(1).MaxLength)
     If (Val(Trim(aBusca(1).Text)) = 0 And Val(Trim(aBusca(1).Text)) = 0) Then
       Listado
     Else
'       If adll.ComboDato(Combo1.Text) = g_tipoped Then
'          Call adll.ListarEnTDBGRID(VGcnx, "vt_pedido", TDBGrid1, "CASE  pedidocondicionfactura WHEN '1' THEN '*' ELSE '' END,pedidofecha,pedidonumero,pedidotipofac,pedidonrofact,clienterazonsocial,pedidototneto", "pedidofecha", nLongicampo, "pedidonumero='" & Trim(aBusca(0) & aBusca(1)) & "'")
'       Else
'          Call adll.ListarEnTDBGRID(VGcnx, "vt_pedido", TDBGrid1, "CASE  pedidocondicionfactura WHEN '1' THEN '*' ELSE '' END,pedidofecha,pedidonumero,pedidotipofac,pedidonrofact,clienterazonsocial,pedidototneto", "pedidofecha", nLongicampo, "pedidonrofact='" & Trim(aBusca(0) & aBusca(1)) & "' and pedidotipofac='" & adll.ComboDato(Combo1.Text) & "'")
'       End If
     End If
     ConfiguraGrid
  
  ElseIf KeyAscii = 13 Then
      SendKeys "{tab}"
      Exit Sub
  End If
  
End Sub

Private Sub aBusca_LostFocus(Index As Integer)
  If Index = 0 Then
     aBusca(0) = Right("0000000000" & Trim(aBusca(0)), aBusca(0).MaxLength)
  Else
     aBusca(1) = Right("0000000000" & Trim(aBusca(1)), aBusca(1).MaxLength)
  End If
End Sub

Private Sub cmdGrabaFinal_Click()
    Dim nume As String
    Dim nsql As String
    Dim J As Double
    Dim Precio As Double
    Dim nrs As New ADODB.Recordset
    Dim nrb As New ADODB.Recordset
    Dim igv As Double
    Dim rsdeta As New ADODB.Recordset
    
    On Error GoTo nerror
    If MsgBox("Desea Grabar Las Guias?", vbYesNo, MsgTitle) = vbYes Then

        If ExisteElem(0, VGcnx, "jtempo") Then
           VGcnx.Execute ("delete from jtempo")
         Else
            MsgBox "No existe la Tabla Temporal jtempo...Verifique!!!", vbInformation, MsgTitle
            Exit Sub
        End If
        
        If ExisteElem(0, VGcnx, "jdetatempo") Then
            VGcnx.Execute "delete from jtempo"
          Else
            MsgBox "No existe la Tabla Temporal jdetatempo...Verifique!!!", vbInformation, MsgTitle
            Exit Sub
        End If
       
       nsql = " insert Into jtempo ( pedidonumero,puntovtacodigo, clienteruc, clientecodigo, clienterazonsocial )"
       nsql = nsql & "  values('1','" & g_ptoventa & "','" & Trim(csql!CARUC) & "' ,'" & Trim(csql!CACODCLI) & "','" & Trim(csql!CANOMCLI) & "' )  "
       VGcnx.Execute (nsql)
       
        nume = g_pedserie & Right("000000000000" & TraeDataSerie("select puntovtadoccorr from vt_puntovtadocumento where documentocodigo='" & g_tipoped & "' and puntovtadocserie='" & g_pedserie & "'", VGcnx), 8)
        nsql = "Update vt_puntovtadocumento " & _
                " set puntovtadoccorr='" & Right("00000000" & Trim(Str(CDbl(nume) + 1)), 8) & "'" & _
                " Where documentocodigo='" & g_tipoped & "' and puntovtacodigo='" & g_ptoventa & "' and puntovtadocserie='" & g_pedserie & "'"
        VGcnx.Execute nsql
        Set nrs = VGcnx.Execute("select * from vt_parametroventa where empresacodigo='" & g_Empresa & "' ")
        
        nsql = "Update jtempo Set pedidonumero='" & nume & "',"
        nsql = nsql & "pedidofecha='" & Date & "', pedidoobserva='' , pedidolistaprec='" & RTrim(Text1(0).Text) & "', "
        nsql = nsql & " pedidomoneda='" & nrs!monedacodigo & "',almacencodigo='" & nrs!almacencodigo & "', "
        nsql = nsql & " pedidotipcambio='1' "
        VGcnx.Execute nsql
        nsql = "select * from " & vt_tempo1
        Set rsdeta = Nothing
        Set rsdeta = VGcnx.Execute(nsql)
        igv = 1 + nrs!paramvtaporcigv
        J = 0
        If rsdeta.RecordCount > 0 Then
           J = 1
           rsdeta.MoveFirst
           Do Until rsdeta.EOF()
                If Text1(1).Text > 0 Then
                    Precio = numero(Text1(1).Text)
                 Else
                     nsql = " select productoprecvta from listapre1 where productocodigo='" & rsdeta!productocodigo & "' "
                     nsql = nsql & " and almacencodigo = '" & nrs!almacencodigo & "'"
                     Set nrb = VGcnx.Execute(nsql)
                     Precio = numero(nrb!productoprecvta)
                End If
                nsql = " Insert Into jdetatempo (pedidonumero,detpeditem,detpedcantpedida,productocodigo ,unidadcodigo,detpedmontoprecvta ) "
                nsql = nsql & " Values('1', " & J & " ,'" & rsdeta!productocantidad & "','" & rsdeta!productocodigo & "','' , "
                nsql = nsql & " " & Precio & "/" & igv & ")"
                VGcnx.Execute nsql
                J = J + 1
                rsdeta.MoveNext
            Loop
        End If
        
        VGcnx.Execute "Update jdetatempo " & _
             " Set pedidonumero='" & nume & "'"
        rsdeta.Close
        nrs.Close
                
        VGcnx.BeginTrans
        VGcnx.Execute "insert into tempopedido" & g_ptoventa & "  Select * from jtempo"
        
        Set nrb = VGcnx.Execute("select * from jdetatempo")
        If nrb.RecordCount > 0 Then
            nrs.Open "tempodetallepedido" & g_ptoventa, VGcnx, adOpenDynamic, adLockOptimistic
            nrb.MoveFirst
            Do Until nrb.EOF
                nrs.AddNew
                For J = 0 To nrb.Fields.Count - 1
                    nrs.Fields(J) = nrb.Fields(J)
                Next J
                nrs.Update
                nrb.MoveNext
            Loop
            Set nrs = Nothing
            MsgBox "Numero de Pedido => " & nume, vbInformation, MsgTitle
        End If
        nrb.Close
        
        Set nrb = Nothing
        
      VGcnx.CommitTrans
      VGcnx.Execute "delete from jdetatempo"
      VGcnx.Execute "delete from jtempo"
      'VGcnx.Execute "Drop Table jTempo"
      'VGcnx.Execute "Drop Table jdetatempo"
      
    End If
    
nerror:
 If Err Then
    MsgBox "Comunicarse con  el Sistema" & Chr(13) & Chr(10) & Err.Number & "-" & Err.Description, vbInformation, MsgTitle
    Err = 0
    VGcnx.RollbackTrans
  
    Exit Sub
 End If
    
End Sub
Private Sub CmdNuevo_Click()
   inicializaarchivo
   Listado
End Sub

Private Sub CmdSalir_Click()
  Unload Me
End Sub

Private Sub CmdGrabar_Click()
    Call dllgeneral.ActivaTab(2, 2, SSTab1)
    Text1(0).Enabled = True
    Text1(1).Enabled = True
End Sub

Private Sub cmdSalirFinal_Click(Index As Integer)
   Call dllgeneral.ActivaTab(0, 1, SSTab1)
End Sub

Private Sub Form_Activate()
  Listado
End Sub

Private Sub Form_Load()
   vt_tempo = "##" & ComputerName & "vt_p" & g_ptoventa
   vt_tempo1 = "##" & ComputerName & "vt_p1" & g_ptoventa
  nLongicampo(1) = 0
  inicializaarchivo
  Call dllgeneral.ActivaTab(0, 1, SSTab1)
  Listado
  ConfiguraGrid
End Sub
Private Sub inicializaarchivo()

   If ExisteElem(0, VGcnx, vt_tempo) Then VGcnx.Execute ("drop table " & vt_tempo)
      
   xsql = " CREATE TABLE " & vt_tempo & " (vt_tipdoc nvarchar (10) ,vt_numdoc nvarchar  (20) ,clientecodigo nvarchar (20),"
   xsql = xsql & " clienterazonsocial nvarchar (50) ,documentoreferencia nvarchar  (10) , numeroreferencia nvarchar (20), "
   xsql = xsql & " almacencodigo nvarchar (2) )  "
    VGcnx.Execute (xsql)

   Set TDBGrid2.DataSource = Nothing
    Call adll.ListarEnTDBGRID(VGcnx, vt_tempo, TDBGrid2, "vt_tipdoc ,vt_numdoc, clientecodigo ,clienterazonsocial ,documentoreferencia, numeroreferencia , almacencodigo", "almacencodigo", nLongicampo, "")
    
    Set TDBGrid2.DataSource = Nothing
    Call adll.ListarEnTDBGRID(VGcnx, vt_tempo, TDBGrid2, "vt_tipdoc ,vt_numdoc, clientecodigo ,clienterazonsocial ,documentoreferencia, numeroreferencia , almacencodigo", "almacencodigo", nLongicampo, "")
    
   If ExisteElem(0, VGcnx, vt_tempo1) Then VGcnx.Execute ("drop table " & vt_tempo1)
      
   xsql = " CREATE TABLE " & vt_tempo1 & " (productocodigo nvarchar (20) ,productodescripcion nvarchar  (100) ,productocantidad float, "
   xsql = xsql & " almacencodigo nvarchar (2) )  "
   VGcnx.Execute (xsql)
    
    Set TDBGrid4.DataSource = Nothing
    Call adll.ListarEnTDBGRID(VGcnx, vt_tempo1, TDBGrid4, "productocodigo, productodescripcion, productocantidad, almacencodigo", "productocodigo", nLongicampo, "")
    Configuradocumento
End Sub

Public Function Listado()

  Set TDBGrid1.DataSource = Nothing
  Set TDBGrid2.DataSource = Nothing
  Set TDBGrid4.DataSource = Nothing
  TDBGrid1.ClearFields
  TDBGrid1.Refresh
  Call adll.ListarEnTDBGRID(VGcnx, "movalmcab", TDBGrid1, "carftdoc,carfndoc,caalma,CATD,CANUMDOC, CAFECDOC,CACODCLI,CARUC, CANOMCLI", "carfndoc", nLongicampo, " catd='NI' and catipmov='I' ")
  Call adll.ListarEnTDBGRID(VGcnx, vt_tempo, TDBGrid2, "vt_tipdoc ,vt_numdoc, clientecodigo ,clienterazonsocial ,documentoreferencia, numeroreferencia , almacencodigo", "almacencodigo", nLongicampo, "")
  Call adll.ListarEnTDBGRID(VGcnx, vt_tempo1, TDBGrid4, "productocodigo, productodescripcion, productocantidad, almacencodigo", "productocodigo", nLongicampo, "")
  TDBGrid2.Refresh
  TDBGrid4.Refresh
  totreg(0) = Format(TDBGrid1.ApproxCount, "#####0")
  totreg(1) = Format(TDBGrid3.ApproxCount, "#####0")
  totreg(2) = Format(TDBGrid4.ApproxCount, "#####0")
  ConfiguraGrid
  Configuradocumento
End Function

Public Function ConfiguraGrid()

   With TDBGrid1
       .Columns(0).Caption = "GR"
       .Columns(0).Width = 400
       .Columns(0).HeadAlignment = dbgCenter
       .Columns(1).Caption = "Nro.Guia"
       .Columns(1).Width = 1000
       .Columns(1).HeadAlignment = dbgCenter
       .Columns(2).Caption = "TD sist."
       .Columns(2).Width = 600
       .Columns(2).HeadAlignment = dbgCenter
       .Columns(3).Caption = "TD sist."
       .Columns(3).Width = 600
       .Columns(3).HeadAlignment = dbgCenter
       .Columns(4).Caption = "Nro.Sistema."
       .Columns(4).Width = 1000
       .Columns(4).HeadAlignment = dbgCenter
       .Columns(5).Caption = "Fecha"
       .Columns(5).Width = 1300
       .Columns(5).HeadAlignment = dbgCenter
       .Columns(6).Caption = "Cod.Cliente"
       .Columns(6).Width = 1200
       .Columns(6).HeadAlignment = dbgCenter
       .Columns(7).Caption = "RUC"
       .Columns(7).Width = 1200
       .Columns(7).HeadAlignment = dbgCenter
       .Columns(8).Caption = "Razon Social"
       .Columns(8).Width = 2200
       .Columns(8).HeadAlignment = dbgCenter
       .Refresh
   End With
   
   
End Function

Public Function Configuradocumento()
   
   With TDBGrid2
       .Columns(0).Caption = "GR"
       .Columns(0).Width = 400
       .Columns(0).HeadAlignment = dbgCenter
       .Columns(1).Caption = "Nro.Guia"
       .Columns(1).Width = 1000
       .Columns(1).HeadAlignment = dbgCenter
       .Columns(2).Caption = "TD sist."
       .Columns(2).Width = 600
       .Columns(2).HeadAlignment = dbgCenter
       .Columns(3).Caption = "Nro.Sistema."
       .Columns(3).Width = 1000
       .Columns(3).HeadAlignment = dbgCenter
       .Columns(4).Caption = "Cod. Cliente"
       .Columns(4).Width = 1300
       .Columns(4).HeadAlignment = dbgCenter
       .Columns(5).Caption = "Razon social"
       .Columns(5).Width = 1200
       .Columns(5).HeadAlignment = dbgCenter
       .Columns(6).Caption = "Almacen"
       .Columns(6).Width = 1200
       .Columns(6).HeadAlignment = dbgCenter
       .Refresh
   End With
   
 With TDBGrid4
       .Columns(0).Caption = "Producto"
       .Columns(0).Width = 1200
       .Columns(0).HeadAlignment = dbgCenter
       .Columns(1).Caption = "Descripcion"
       .Columns(1).Width = 6500
       .Columns(1).HeadAlignment = dbgCenter
       .Columns(2).Caption = "Cantidad"
       .Columns(2).Width = 1000
       .Columns(3).Caption = "Almacen."
       .Columns(3).Width = 1200
       .Columns(3).HeadAlignment = dbgCenter
       .Columns(2).NumberFormat = "###,##0.00"
       .Refresh
   End With
  
End Function

Private Sub TDBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
     If TDBGrid1.ApproxCount > 0 Then
        xalma = TDBGrid1.Columns(2).Text
        xtipo = TDBGrid1.Columns(3).Text
        xnumero = TDBGrid1.Columns(4).Text
        dBusca
        Call dllgeneral.ActivaTab(1, 1, SSTab1)
        Listado
     End If
  End If

End Sub
Private Sub TDBGrid1_DblClick()
     If TDBGrid1.ApproxCount > 0 Then
        xalma = TDBGrid1.Columns(2).Text
        xtipo = TDBGrid1.Columns(3).Text
        xnumero = TDBGrid1.Columns(4).Text
        dBusca
        Call dllgeneral.ActivaTab(1, 1, SSTab1)
        Listado
    End If

End Sub


Private Sub dBusca()
    Dim csqld As New ADODB.Recordset
    Dim acliente As New ADODB.Recordset
    Dim nsql As String
    Dim J As Integer
    
   ' Call Limpiartexto(MBox2, 6, 10)
  '  Call Limpiartexto(Label2, 0, 8)
    Call CargaGrilla
    
    xsql = " select * from movalmcab where caalma ='" & xalma & "' and catd='" & xtipo & "'  and canumdoc='" & xnumero & "'  "
'    nvalor = ""
    Set csql = VGcnx.Execute(xsql)
    If csql.RecordCount > 0 Then

        Set acliente = VGcnx.Execute("select * from vt_cliente where clientecodigo='" & Escadena(csql!CACODCLI) & "'")
        If acliente.RecordCount > 0 Then
           clienterazon(0) = Escadena(acliente!clienterazonsocial)
           clienterazon(1) = Escadena(acliente!clienterazonsocial)
        End If
        acliente.Close
        Set acliente = Nothing
        Set acliente = VGcnx.Execute("select * from vt_almacen where almacencodigo='" & Escadena(csql!CAALMA) & "'")
        If acliente.RecordCount > 0 Then
           almacendescr(0) = Escadena(acliente!almacencodigo) & " - " & Escadena(acliente!almacendescripcion)
           almacendescr(1) = Escadena(acliente!almacencodigo) & " - " & Escadena(acliente!almacendescripcion)
        End If
        acliente.Close
        Set acliente = Nothing
        
    Else
        MsgBox "No existe Informacion del Documento...Verifique!!", vbInformation, MsgTitle
        csql.Close
        Set csql = Nothing
        Exit Sub
    End If
       
    Set csqld = VGcnx.Execute("select DEITEM,A.decodigo,b.adescri,b.aunidad," & _
                          "DECANTID " & _
                          "from movalmdet A inner join " & _
                          "[" & VGcnx.DefaultDatabase & "].dbo.maeart B" & _
                          " ON A.decodigo=b.acodigo COLLATE Modern_Spanish_CI_AI " & _
                          "where dealma='" & xalma & "' and detd='" & xtipo & "' and denumdoc='" & xnumero & "'  ")
    
    Set rsdeta = Nothing
    Call CargaGrilla

    Do Until csqld.EOF
       rsdeta.AddNew
       rsdeta.Fields(0) = Escadena(csqld!deitem)
       rsdeta.Fields(1) = Escadena(csqld!decodigo)
       rsdeta.Fields(2) = Escadena(csqld!adescri)
       rsdeta.Fields(3) = Escadena(csqld!aunidad)
       rsdeta.Fields(4) = numero(csqld!decantid)
       rsdeta.Update
       csqld.MoveNext
    Loop
    csqld.Close
    Call ConfigGrid
End Sub

Public Function CargaGrilla()

   Set rsdeta = Nothing
   Call rsdeta.Fields.Append("Item", adInteger)
   Call rsdeta.Fields.Append("Codigo", adChar, 20)
   Call rsdeta.Fields.Append("Descripcion", adChar, 100)
   Call rsdeta.Fields.Append("UM", adChar, 3)
   Call rsdeta.Fields.Append("Cant", adDouble)
   rsdeta.Open
   ConfigGrid
End Function

Public Function ConfigGrid()

   Set TDBGrid3.DataSource = rsdeta
   With TDBGrid3
      .Columns(0).Width = 600
      .Columns(0).Caption = "Item"
      .Columns(1).Width = 1100
      .Columns(1).Caption = "Codigo"
      .Columns(2).Width = 5500
      .Columns(2).Caption = "Descripcion"
      .Columns(3).Width = 800
      .Columns(3).Caption = "UM"
      .Columns(4).Width = 1200
      .Columns(4).Caption = "Cant"
      .Columns(4).NumberFormat = "###,##0.00"
   End With
   TDBGrid3.Refresh

End Function

Private Sub cmdBotones_Click(Index As Integer)
  On Error GoTo nerror
  Select Case Index
  Case 11
    acumulaguias
    Call dllgeneral.ActivaTab(0, 1, SSTab1)
    TDBGrid2.Refresh
   
  Case 12
    Call dllgeneral.ActivaTab(0, 1, SSTab1)
  
  End Select
  
nerror:
   If Err Then
       MsgBox Err.Description & "-" & Err.Description, vbInformation, MsgTitle
       Err = 0
       Resume Next
       Exit Sub
   End If
End Sub

Private Sub acumulaguias()
    xsql = " Insert " & vt_tempo & " (vt_tipdoc,vt_numdoc,clientecodigo,clienterazonsocial,documentoreferencia,numeroreferencia,almacencodigo) "
    xsql = xsql & " values( '" & Escadena(csql!carftdoc) & "', '" & Escadena(csql!carfndoc) & "','" & csql!CACODCLI & "' , "
    xsql = xsql & " '" & csql!CANOMCLI & "', '" & csql!CATD & "', '" & csql!CANUMDOC & "','" & csql!CAALMA & "' )"
    VGcnx.Execute xsql
    
    If rsdeta.RecordCount > 0 Then
       rsdeta.MoveFirst
       Do Until rsdeta.EOF()
            Set SQL = VGcnx.Execute(" Select *  from " & vt_tempo1 & " where productocodigo = '" & rsdeta!codigo & "' ")
            If SQL.RecordCount > 0 Then
                xsql = " Update " & vt_tempo1 & " SET productocantidad = productocantidad + " & rsdeta!cant & ""
                xsql = xsql & " Where productocodigo='" & Trim(rsdeta!codigo) & "' "
              Else
                 xsql = " Insert " & vt_tempo1 & " (productocodigo, productodescripcion, productocantidad ) "
                 xsql = xsql & " values( '" & Escadena(rsdeta!codigo) & "', '" & Escadena(rsdeta!descripcion) & "', " & rsdeta!cant & ") "
           End If
           VGcnx.Execute xsql
           rsdeta.MoveNext
       Loop
    End If
    Listado
End Sub

