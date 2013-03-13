VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form FrmPendientespedidos 
   Caption         =   "Entrega a Clientes"
   ClientHeight    =   6840
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   10770
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6840
   ScaleWidth      =   10770
   Begin TabDlg.SSTab SSTab1 
      Height          =   6735
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   10650
      _ExtentX        =   18785
      _ExtentY        =   11880
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   420
      TabCaption(0)   =   "Pendientes de Entrega"
      TabPicture(0)   =   "FrmPendientespedidos.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Fr1"
      Tab(0).Control(1)=   "Fr2(0)"
      Tab(0).Control(2)=   "CmdGrabardata"
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Detalle de Guias"
      TabPicture(1)   =   "FrmPendientespedidos.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame6"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame5(1)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin VB.Frame Frame5 
         Height          =   5340
         Index           =   1
         Left            =   96
         TabIndex        =   25
         Top             =   1176
         Width           =   10392
         Begin VB.Frame Frame4 
            Height          =   930
            Left            =   7968
            TabIndex        =   29
            Top             =   4176
            Width           =   2010
            Begin VB.CommandButton cmdBotones 
               Caption         =   "&Cancelar"
               Height          =   690
               Index           =   12
               Left            =   1140
               Picture         =   "FrmPendientespedidos.frx":0038
               Style           =   1  'Graphical
               TabIndex        =   31
               Top             =   180
               Width           =   825
            End
            Begin VB.CommandButton cmdBotones 
               Caption         =   "&Acepta"
               Height          =   690
               Index           =   11
               Left            =   90
               Picture         =   "FrmPendientespedidos.frx":047A
               Style           =   1  'Graphical
               TabIndex        =   30
               Top             =   180
               Width           =   870
            End
         End
         Begin VB.Frame Fr2 
            BackColor       =   &H00800000&
            BorderStyle     =   0  'None
            Height          =   765
            Index           =   1
            Left            =   5316
            TabIndex        =   26
            Top             =   4344
            Width           =   2055
            Begin MSMask.MaskEdBox totreg 
               Height          =   375
               Index           =   1
               Left            =   300
               TabIndex        =   27
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
               TabIndex        =   28
               Top             =   495
               Width           =   1335
            End
         End
         Begin TrueOleDBGrid70.TDBGrid TDBGrid2 
            Height          =   3468
            Left            =   156
            TabIndex        =   32
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
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HC0FFFF&,.bold=0,.fontsize=825"
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
      Begin VB.Frame Frame6 
         Height          =   972
         Left            =   96
         TabIndex        =   16
         Top             =   336
         Width           =   10284
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
            TabIndex        =   24
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
            Index           =   11
            Left            =   2880
            TabIndex        =   23
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
            Index           =   10
            Left            =   330
            TabIndex        =   22
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
            Index           =   9
            Left            =   6384
            TabIndex        =   21
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
            Index           =   0
            Left            =   1440
            TabIndex        =   20
            Top             =   240
            Width           =   1335
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
            TabIndex        =   19
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
            Height          =   285
            Left            =   1440
            TabIndex        =   18
            Top             =   540
            Width           =   8730
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
            Left            =   7488
            TabIndex        =   17
            Top             =   180
            Width           =   2664
         End
      End
      Begin VB.Frame Fr1 
         Height          =   6960
         Left            =   -74904
         TabIndex        =   10
         Top             =   1104
         Width           =   10416
         Begin VB.Frame Frame5 
            Height          =   585
            Index           =   0
            Left            =   7644
            TabIndex        =   11
            Top             =   4965
            Width           =   2265
            Begin MSMask.MaskEdBox totreg 
               Height          =   372
               Index           =   0
               Left            =   1104
               TabIndex        =   12
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
               TabIndex        =   13
               Top             =   192
               Width           =   1032
            End
         End
         Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
            Height          =   4350
            Left            =   240
            TabIndex        =   14
            Top             =   480
            Width           =   9975
            _ExtentX        =   17595
            _ExtentY        =   7673
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
            Columns(7)._VlistStyle=   0
            Columns(7)._MaxComboItems=   5
            Columns(7).DataField=   ""
            Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(8)._VlistStyle=   0
            Columns(8)._MaxComboItems=   5
            Columns(8).DataField=   ""
            Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(9)._VlistStyle=   0
            Columns(9)._MaxComboItems=   5
            Columns(9).DataField=   ""
            Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(10)._VlistStyle=   0
            Columns(10)._MaxComboItems=   5
            Columns(10).DataField=   ""
            Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   11
            Splits(0)._UserFlags=   0
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   13160660
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=11"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(5)=   "Column(1).Width=2725"
            Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2646"
            Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(9)=   "Column(2).Width=2725"
            Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=2646"
            Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(13)=   "Column(3).Width=2725"
            Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=2646"
            Splits(0)._ColumnProps(16)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(17)=   "Column(4).Width=2725"
            Splits(0)._ColumnProps(18)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(19)=   "Column(4)._WidthInPix=2646"
            Splits(0)._ColumnProps(20)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(21)=   "Column(5).Width=2725"
            Splits(0)._ColumnProps(22)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(23)=   "Column(5)._WidthInPix=2646"
            Splits(0)._ColumnProps(24)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(25)=   "Column(6).Width=2725"
            Splits(0)._ColumnProps(26)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(27)=   "Column(6)._WidthInPix=2646"
            Splits(0)._ColumnProps(28)=   "Column(6).Order=7"
            Splits(0)._ColumnProps(29)=   "Column(7).Width=2725"
            Splits(0)._ColumnProps(30)=   "Column(7).DividerColor=0"
            Splits(0)._ColumnProps(31)=   "Column(7)._WidthInPix=2646"
            Splits(0)._ColumnProps(32)=   "Column(7).Order=8"
            Splits(0)._ColumnProps(33)=   "Column(8).Width=2725"
            Splits(0)._ColumnProps(34)=   "Column(8).DividerColor=0"
            Splits(0)._ColumnProps(35)=   "Column(8)._WidthInPix=2646"
            Splits(0)._ColumnProps(36)=   "Column(8).Order=9"
            Splits(0)._ColumnProps(37)=   "Column(9).Width=2725"
            Splits(0)._ColumnProps(38)=   "Column(9).DividerColor=0"
            Splits(0)._ColumnProps(39)=   "Column(9)._WidthInPix=2646"
            Splits(0)._ColumnProps(40)=   "Column(9).Order=10"
            Splits(0)._ColumnProps(41)=   "Column(10).Width=2725"
            Splits(0)._ColumnProps(42)=   "Column(10).DividerColor=0"
            Splits(0)._ColumnProps(43)=   "Column(10)._WidthInPix=2646"
            Splits(0)._ColumnProps(44)=   "Column(10).Order=11"
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
            _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=70,.parent=13"
            _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=67,.parent=14"
            _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=68,.parent=15"
            _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=69,.parent=17"
            _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=74,.parent=13"
            _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=71,.parent=14"
            _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=72,.parent=15"
            _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=73,.parent=17"
            _StyleDefs(76)  =   "Splits(0).Columns(10).Style:id=78,.parent=13"
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
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   "Pedidos"
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
            TabIndex        =   15
            Top             =   96
            Width           =   9528
         End
      End
      Begin VB.Frame Fr2 
         Height          =   645
         Index           =   0
         Left            =   -74868
         TabIndex        =   2
         Top             =   480
         Width           =   10290
         Begin VB.CommandButton cBusca 
            BackColor       =   &H80000008&
            Caption         =   "&Buscar"
            Height          =   375
            Left            =   8925
            TabIndex        =   7
            Top             =   180
            Width           =   1095
         End
         Begin VB.TextBox aBusca 
            Height          =   285
            Index           =   0
            Left            =   7365
            MaxLength       =   3
            TabIndex        =   6
            Top             =   210
            Width           =   435
         End
         Begin VB.TextBox aBusca 
            Height          =   285
            Index           =   1
            Left            =   7995
            MaxLength       =   8
            TabIndex        =   5
            Top             =   210
            Width           =   885
         End
         Begin VB.ComboBox Combo2 
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   1710
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   240
            Width           =   1845
         End
         Begin VB.CheckBox ChkTodos 
            Caption         =   "Incluir Todos"
            Height          =   375
            Left            =   4680
            TabIndex        =   3
            Top             =   120
            Width           =   1575
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
            TabIndex        =   9
            Top             =   240
            Width           =   192
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Pto.  Venta"
            ForeColor       =   &H000000FF&
            Height          =   255
            Index           =   7
            Left            =   600
            TabIndex        =   8
            Top             =   270
            Width           =   825
         End
      End
      Begin VB.CommandButton CmdGrabardata 
         Caption         =   "Grabar"
         Height          =   492
         Left            =   -71208
         TabIndex        =   1
         Top             =   7440
         Width           =   1068
      End
   End
End
Attribute VB_Name = "FrmPendientespedidos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsacumula As New ADODB.Recordset
Dim rsdeta As New ADODB.Recordset
Dim csql As New ADODB.Recordset
Dim SQL As New ADODB.Recordset
Dim VGDllGeneral As New dllgeneral.dll_general
Dim dllgeneral As New dllgeneral.dll_general
Dim vt_tempo As String, vt_tempo1 As String
Dim xsql, xAlma, xtipo, xnumero As String
Dim g_tipoped As String
Dim g_pedserie As String
Dim acepta As Integer
Dim nLongicampo(1) As Integer
Private Sub aBusca_KeyPress(Index As Integer, KeyAscii As Integer)
  Dim ldato As String
  If KeyAscii = 13 And Index = 1 Then
     TDBGrid1.ClearFields
     Set TDBGrid1.DataSource = Nothing
     TDBGrid1.Refresh
     aBusca(0) = Right("0000000000" & Trim(aBusca(0)), aBusca(0).MaxLength)
     aBusca(1) = Right("0000000000" & Trim(aBusca(1)), aBusca(1).MaxLength)
     
     If (Val(Trim(aBusca(1).text)) = 0 And Val(Trim(aBusca(1).text)) = 0) Then
       Listado
     Else
'      If VGDllGeneral.ComboDato(Combo1.text) = g_tipoped Then
'          Call VGDllGeneral.ListarEnTDBGRID(VGCNx, "al_liquidacionCompra", TDBGrid1, "CASE  pedidocondicionfactura WHEN '1' THEN '*' ELSE '' END,pedidofecha,pedidonumero,pedidotipofac,pedidonrofact,clienterazonsocial,pedidototneto", "pedidofecha", nLongicampo, "pedidonumero='" & Trim(aBusca(0) & aBusca(1)) & "'")
'       Else
'          Call VGDllGeneral.ListarEnTDBGRID(VGcnx, "al_liquidacionCompra", TDBGrid1, "CASE  pedidocondicionfactura WHEN '1' THEN '*' ELSE '' END,pedidofecha,pedidonumero,pedidotipofac,pedidonrofact,clienterazonsocial,pedidototneto", "pedidofecha", nLongicampo, "pedidonrofact='" & Trim(aBusca(0) & aBusca(1)) & "' and pedidotipofac='" & VGDllGeneral.ComboDato(Combo1.Text) & "'")
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


Private Sub ChkTodos_Click()
If ChkTodos = 1 Then
   Call VGDllGeneral.ListarEnTDBGRID(VGCNx, "movalmcab", TDBGrid1, "carftdoc,carfndoc,caalma,CATD,CANUMDOC, CAFECDOC,CACODCLI,CARUC, CANOMCLI", "cafecdoc", nLongicampo, " catd='NI' and catipmov='I' and carftdoc in('NC')")
Else
   Call VGDllGeneral.ListarEnTDBGRID(VGCNx, "movalmcab", TDBGrid1, "carftdoc,carfndoc,caalma,CATD,CANUMDOC, CAFECDOC,CACODCLI,CARUC, CANOMCLI", "cafecdoc", nLongicampo, " catd='NI' and catipmov='I' and carftdoc in('NC') and isnull(canroped,0)=0")
End If
End Sub

Private Sub cmdNuevo_Click()
   Listado
End Sub

Private Sub CmdSalir_Click()
  Unload Me
End Sub


Private Sub cmdSalirFinal_Click(Index As Integer)
   Call dllgeneral.ActivaTab(0, 1, SSTab1)
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub Combo2_LostFocus()
 Dim rst As New ADODB.Recordset
   Set rst = VGCNx.Execute("select * from vt_puntovtadocumento where puntovtacodigo='" & Left(Combo2.text, 2) & "' and documentocodigo='04'")
      If rst.RecordCount > 0 Then
         g_ptoventa = Left(Combo2.text, 2)
         g_pedserie = rst!puntovtadocserie
         g_tipoped = "04"
         rst.Close
    End If
  
 Set rst = Nothing
End Sub

Private Sub Form_Activate()
  Listado
End Sub

Private Sub Form_Load()
   vt_tempo = "##" & ComputerName & "vt_p" & g_ptoventa
   vt_tempo1 = "##" & ComputerName & "vt_p1" & g_ptoventa
  nLongicampo(1) = 0
  Call VGDllGeneral.llenacombo(Combo2, "select puntovtacodigo,puntovtadescripcion from vt_puntoventa", VGCNx)
  Call dllgeneral.ActivaTab(0, 1, SSTab1)
  Listado
  ConfiguraGrid
End Sub

Public Function Listado()

  Set TDBGrid1.DataSource = Nothing
  Set TDBGrid2.DataSource = Nothing
  TDBGrid1.ClearFields
  TDBGrid1.Refresh
  Call VGDllGeneral.ListarEnTDBGRID(VGCNx, "movalmcab", TDBGrid1, "carftdoc,carfndoc,caalma,CATD,CANUMDOC, CAFECDOC,CACODCLI,CARUC, CANOMCLI", "cafecdoc", nLongicampo, " catd='GR' and casitgui='F' ")
  TDBGrid1.Refresh
  totreg(0) = Format(TDBGrid1.ApproxCount, "#####0")
  totreg(1) = Format(TDBGrid2.ApproxCount, "#####0")
  ConfiguraGrid
  End Function

Public Function ConfiguraGrid()

   With TDBGrid1
       .Columns(0).Caption = "TD"
       .Columns(0).Width = 200
       .Columns(0).HeadAlignment = dbgCenter
       .Columns(1).Caption = "Nro.Doc"
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
       .Columns(6).Caption = "Cod.Proveedor"
       .Columns(6).Width = 1200
       .Columns(6).HeadAlignment = dbgCenter
       .Columns(7).Caption = "RUC"
       .Columns(7).Width = 1200
       .Columns(7).HeadAlignment = dbgCenter
       .Columns(8).Caption = "Razon Social"
       .Columns(8).Width = 1200
       .Columns(8).HeadAlignment = dbgCenter
       .Refresh
   End With
   
   
End Function

Private Sub tdbgrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
     If TDBGrid1.ApproxCount > 0 Then
        xAlma = TDBGrid1.Columns(2).text
        xtipo = TDBGrid1.Columns(3).text
        xnumero = TDBGrid1.Columns(4).text
        dBusca
        Call dllgeneral.ActivaTab(1, 1, SSTab1)
        Listado
     End If
  End If

End Sub
Private Sub TDBGrid1_DblClick()
     If TDBGrid1.ApproxCount > 0 Then
        xAlma = TDBGrid1.Columns(2).text
        xtipo = TDBGrid1.Columns(3).text
        xnumero = TDBGrid1.Columns(4).text
        umerodoc(0) = TDBGrid1.Columns(1).text
        dBusca
        Call dllgeneral.ActivaTab(1, 1, SSTab1)
    End If

End Sub


Private Sub dBusca()
    Dim csqld As New ADODB.Recordset
    Dim acliente As New ADODB.Recordset
    Dim nsql As String
    Dim j As Integer
    
   ' Call Limpiartexto(MBox2, 6, 10)
  '  Call Limpiartexto(Label2, 0, 8)
    Call CargaGrilla
    
    xsql = " select * from movalmcab where caalma ='" & xAlma & "' and catd='" & xtipo & "'  and canumdoc='" & xnumero & "'  "
'    nvalor = ""
    Set csql = VGCNx.Execute(xsql)
    If csql.RecordCount > 0 Then
        fechadoc(0) = csql!cafecdoc
        Set acliente = VGCNx.Execute("select * from vt_cliente where clientecodigo='" & Escadena(csql!CACODcli) & "'")
        If acliente.RecordCount > 0 Then
           clienterazon = Escadena(acliente!clienterazonsocial)
        End If
        acliente.Close
        Set acliente = Nothing
        Set acliente = VGCNx.Execute("select * from vt_almacen where almacencodigo='" & Escadena(csql!CAALMA) & "'")
        If acliente.RecordCount > 0 Then
           almacendescr = Escadena(acliente!almacencodigo) & " - " & Escadena(acliente!almacendescripcion)
        End If
        acliente.Close
        Set acliente = Nothing
        
        'umerodoc(0) = xnumero
        
    Else
        MsgBox "No existe Informacion del Documento...Verifique!!", vbInformation, MsgTitle
        csql.Close
        Set csql = Nothing
        Exit Sub
    End If
    nsql = "select DEITEM,A.decodigo,b.adescri,b.aunidad,DECANTID,"
    nsql = nsql & "decodmon,deprecio,detipcam from movalmdet A "
    nsql = nsql & " inner join [" & VGCNx.DefaultDatabase & "].dbo.maeart b"
    nsql = nsql & " ON A.decodigo=b.acodigo where dealma='" & xAlma & "' and detd='" & xtipo & "' and denumdoc='" & xnumero & "'  "
    Set csqld = VGCNx.Execute(nsql)
    
    Set rsdeta = Nothing
    Call CargaGrilla

    Do Until csqld.EOF
       rsdeta.AddNew
       rsdeta.Fields(0) = Escadena(csqld!DEITEM)
       rsdeta.Fields(1) = Escadena(csqld!decodigo)
       rsdeta.Fields(2) = Escadena(csqld!ADESCRI)
       rsdeta.Fields(3) = Escadena(csqld!aunidad)
       rsdeta.Fields(4) = numero(csqld!DECANTID)
       rsdeta.Fields(5) = Escadena(csqld!DECODMON)
       rsdeta.Fields(6) = numero(csqld!DEPRECIO)
       rsdeta.Fields(7) = numero(csqld!DETIPCAM)
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
   Call rsdeta.Fields.Append("Moneda", adChar, 2)
   Call rsdeta.Fields.Append("Precio", adDouble)
   Call rsdeta.Fields.Append("TipodeCambio", adDouble)
   rsdeta.Open
   ConfigGrid
End Function

Public Function ConfigGrid()

   Set TDBGrid2.DataSource = rsdeta
   With TDBGrid2
      .Columns(0).Width = 400
      .Columns(0).Caption = "Item"
      .Columns(1).Width = 1100
      .Columns(1).Caption = "Codigo"
      .Columns(2).Width = 3000
      .Columns(2).Caption = "Descripcion"
      .Columns(3).Width = 400
      .Columns(3).Caption = "UM"
      .Columns(4).Width = 1200
      .Columns(4).Caption = "Cant"
      .Columns(4).NumberFormat = "###,##0.00"
      .Columns(5).Caption = "Moneda"
      .Columns(5).Width = 400
      .Columns(6).Caption = "precio"
      .Columns(6).Width = 900
      .Columns(6).NumberFormat = "###,##0.000"
      .Columns(7).Caption = "T.Cambio"
      .Columns(7).Width = 900
      .Columns(7).NumberFormat = "###,##0.000"
   End With
   TDBGrid2.Refresh

End Function

Private Sub cmdBotones_Click(Index As Integer)
  On Error GoTo nerror
  Select Case Index
  Case 11
    actualizastock
    Call dllgeneral.ActivaTab(0, 1, SSTab1)
    TDBGrid2.Refresh
   
    Call imprimirfacturas
    
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

Private Sub actualizastock()
rsdeta.MoveFirst
Dim xrsql As New ADODB.Recordset
Dim acmd As New ADODB.Command
Do Until rsdeta.EOF()
   Set acmd.ActiveConnection = VGgeneral
   acmd.CommandType = adCmdStoredProc
   acmd.CommandTimeout = 0
   acmd.CommandText = "vt_actualizoalma_pro"
   acmd.Prepared = True
   With acmd
       .Parameters("@basedes") = CStr(VGCNx.DefaultDatabase)
       .Parameters("@almacen") = xAlma
       .Parameters("@tipo") = "4"
       .Parameters("@articulo") = rsdeta!codigo
       .Parameters("@cantidad") = rsdeta!cant
   End With
   acmd.Execute
   Set acmd = Nothing
   rsdeta.MoveNext
Loop
' SQL = " update movalmcab set casitgui='V' where caalma='" & xAlma & "' and catd='GR' and canumdoc='" & xnumero & "'"
End Sub

Private Sub acumulaguias()
    xsql = " Insert " & vt_tempo & " (vt_tipdoc,vt_numdoc,clientecodigo,clienterazonsocial,documentoreferencia,numeroreferencia,almacencodigo,fecha) "
    xsql = xsql & " values( '" & Escadena(csql!CARFTDOC) & "', '" & Escadena(csql!CARFNDOC) & "','" & csql!CACODcli & "' , "
    xsql = xsql & " '" & csql!CANOMCLI & "', '" & csql!CATD & "', '" & csql!CANUMDOC & "','" & csql!CAALMA & "','" & csql!cafecdoc & "')"
    VGCNx.Execute xsql
    
    If rsdeta.RecordCount > 0 Then
       rsdeta.MoveFirst
       Do Until rsdeta.EOF()
            Set SQL = VGCNx.Execute(" Select *  from " & vt_tempo1 & " where productocodigo = '" & rsdeta!codigo & "' ")
            If SQL.RecordCount > 0 Then
                xsql = " Update " & vt_tempo1 & " SET productocantidad = productocantidad + " & rsdeta!cant & ""
                xsql = xsql & " Where productocodigo='" & Trim(rsdeta!codigo) & "' "
              Else
                 xsql = " Insert " & vt_tempo1 & " (productocodigo, productodescripcion, productocantidad,precio,moneda,tipodecambio) "
                 xsql = xsql & " values( '" & Escadena(rsdeta!codigo) & "', '" & Escadena(rsdeta!descripcion) & "', " & rsdeta!cant & "," & rsdeta!precio & ",'"
                 xsql = xsql & rsdeta!moneda & "'," & rsdeta!TipoDeCambio & ") "
           End If
           VGCNx.Execute xsql
           rsdeta.MoveNext
       Loop
    End If
    Listado
End Sub


Private Sub imprimirfacturas()
 Dim formulas(13) As Variant
 Dim Param(2) As Variant
 Dim reporte As String
 
'      If cOpc2(0).Value Then
'          formulas(0) = "nro='" & aBusca(0) & "'"
'       ElseIf cOpc2(1).Value Then
'          formulas(0) = "nro='" & MBox(3) & "'"
'       ElseIf cOpc2(2).Value Then
'          formulas(0) = "nro='" & MBox(4) & "'"
'       End If
'       formulas(1) = "cliente='" & Trim(MBox3(1)) & "'"
'       formulas(2) = "fecha='" & CStr(Day(CDate(MBox(10)))) & "   " & Format(Month(CDate(MBox(10))), "00") & "  " & Right(CStr(Year(CDate(MBox(10)))), 2) & "'"
'       formulas(3) = "direccion='" & "" & Trim(MBox3(3)) & "'"
'       formulas(4) = "dni='" & "" & Trim(MBox3(2)) & "'"
'       If cOpc2(0).Value Or cOpc2(1).Value Or cOpc2(2).Value Then
'          If cOpc2(0).Value Then
'            formulas(5) = "letras= '" & "SON : " & dllgeneral.NUMLET(numero(Round(CDbl(MBox2(10)), 2))) & IIf(dllgeneral.ComboDato(Combo1.text) = g_TipoSol, "Nuevos Soles", "Dolares Americanos") & "'"
'          Else
'            formulas(5) = "letras= '" & "SON : " & dllgeneral.NUMLET(numero(Round(CDbl(MBox2(10)), 2))) & IIf(dllgeneral.ComboDato(Combo1.text) = g_TipoSol, "Nuevos Soles", "Dolares Americanos") & "'"
'          End If
'       End If
'       formulas(5) = "letras= '" & "SON : " & dllgeneral.NUMLET(Round(CDbl(MBox2(10)), 2)) & IIf(dllgeneral.ComboDato(Combo1.text) = g_TipoSol, "Nuevos Soles", "Dolares Americanos") & "'"
'       formulas(6) = "guias='" & guias_num & "'"
'       formulas(7) = "vendedor='" & Escadena(Ctr_Ayuda2.xnombre) & "'"
'       formulas(8) = "bruto='" & Round(numero(MBox2(7)), 2) & "'"
'       formulas(9) = "dscto='" & Round(MBox2(8), 2) & "'"
'       formulas(10) = "igv='" & Round(MBox2(9), 2) & "'"
'       formulas(11) = "ruc='" & MBox3(2) & "'"
'       formulas(12) = "detraccion='" & Detraccion & "'"
'
'       'End If
       Param(0) = VGCNx.DefaultDatabase
       Param(1) = umerodoc(0)
'       Param(1) = VGParametros.empresacodigo
       'If VGparametros.multifacturas Then
       '   reporte = "vt_guiaimpresa_" & VGCNx.DefaultDatabase & VGparametros.empresacodigo & ".rpt"
       ' Else
          reporte = "vt_guiaimpresa_" & VGCNx.DefaultDatabase & ".rpt"
        
       'End If
       Call ImpresionRptProc(reporte, formulas, Param, , "impresion de Guias de Remision")

End Sub


