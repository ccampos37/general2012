VERSION 5.00
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form FrmDetallePedidoxGuias 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detalle Documento "
   ClientHeight    =   7185
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11850
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7185
   ScaleWidth      =   11850
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame6 
      Height          =   1065
      Left            =   48
      TabIndex        =   47
      Top             =   96
      Width           =   10905
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
         TabIndex        =   55
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
         TabIndex        =   54
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
         TabIndex        =   53
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
         Height          =   225
         Index           =   9
         Left            =   6810
         TabIndex        =   52
         Top             =   240
         Width           =   765
      End
      Begin VB.Label Label2 
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
         Index           =   12
         Left            =   1440
         TabIndex        =   51
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label2 
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
         Index           =   11
         Left            =   4260
         TabIndex        =   50
         Top             =   240
         Width           =   2475
      End
      Begin VB.Label Label2 
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
         Index           =   10
         Left            =   1440
         TabIndex        =   49
         Top             =   570
         Width           =   9255
      End
      Begin VB.Label Label2 
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
         Index           =   9
         Left            =   7920
         TabIndex        =   48
         Top             =   180
         Width           =   2805
      End
   End
   Begin VB.Frame Frame5 
      Height          =   4815
      Left            =   0
      TabIndex        =   42
      Top             =   1080
      Width           =   10965
      Begin VB.Frame Fr2 
         BackColor       =   &H00800000&
         BorderStyle     =   0  'None
         Height          =   765
         Index           =   0
         Left            =   6030
         TabIndex        =   43
         Top             =   3960
         Width           =   2055
         Begin MSMask.MaskEdBox MBox2 
            Height          =   375
            Index           =   0
            Left            =   300
            TabIndex        =   44
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
            TabIndex        =   45
            Top             =   495
            Width           =   1335
         End
      End
      Begin TrueOleDBGrid70.TDBGrid TDBGrid2 
         Height          =   3615
         Left            =   150
         TabIndex        =   46
         Top             =   270
         Width           =   10635
         _ExtentX        =   18759
         _ExtentY        =   6376
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
   Begin VB.Frame Frame4 
      Height          =   930
      Left            =   4620
      TabIndex        =   39
      Top             =   5790
      Width           =   2010
      Begin VB.CommandButton cmdBotones 
         Caption         =   "&Cancelar"
         Height          =   690
         Index           =   12
         Left            =   1140
         Picture         =   "FrmDetalleiPedidoxguias.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   180
         Width           =   825
      End
      Begin VB.CommandButton cmdBotones 
         Caption         =   "&Acepta"
         Height          =   690
         Index           =   11
         Left            =   90
         Picture         =   "FrmDetalleiPedidoxguias.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   180
         Width           =   870
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4215
      Left            =   0
      TabIndex        =   19
      Top             =   1620
      Width           =   11805
      Begin VB.Frame Fr2 
         BackColor       =   &H00800000&
         BorderStyle     =   0  'None
         Height          =   765
         Index           =   2
         Left            =   150
         TabIndex        =   26
         Top             =   3240
         Width           =   11535
         Begin MSMask.MaskEdBox MBox2 
            Height          =   375
            Index           =   6
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
         Begin MSMask.MaskEdBox MBox2 
            Height          =   375
            Index           =   7
            Left            =   2400
            TabIndex        =   28
            Top             =   90
            Width           =   1875
            _ExtentX        =   3307
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
         Begin MSMask.MaskEdBox MBox2 
            Height          =   375
            Index           =   8
            Left            =   4800
            TabIndex        =   29
            Top             =   90
            Width           =   1905
            _ExtentX        =   3360
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
         Begin MSMask.MaskEdBox MBox2 
            Height          =   375
            Index           =   9
            Left            =   7290
            TabIndex        =   30
            Top             =   90
            Width           =   1935
            _ExtentX        =   3413
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
         Begin MSMask.MaskEdBox MBox2 
            Height          =   375
            Index           =   10
            Left            =   9540
            TabIndex        =   31
            Top             =   90
            Width           =   1905
            _ExtentX        =   3360
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
         Begin VB.Line Line2 
            BorderColor     =   &H00808080&
            Index           =   7
            X1              =   9340
            X2              =   9340
            Y1              =   120
            Y2              =   1215
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00808080&
            Index           =   6
            X1              =   6940
            X2              =   6940
            Y1              =   120
            Y2              =   1215
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00808080&
            Index           =   5
            X1              =   4420
            X2              =   4420
            Y1              =   120
            Y2              =   1215
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00808080&
            Index           =   4
            X1              =   2160
            X2              =   2160
            Y1              =   120
            Y2              =   1215
         End
         Begin VB.Line Line2 
            BorderColor     =   &H80000005&
            Index           =   3
            X1              =   9360
            X2              =   9360
            Y1              =   120
            Y2              =   1215
         End
         Begin VB.Line Line2 
            BorderColor     =   &H80000005&
            Index           =   2
            X1              =   6960
            X2              =   6960
            Y1              =   120
            Y2              =   1215
         End
         Begin VB.Line Line2 
            BorderColor     =   &H80000005&
            Index           =   1
            X1              =   4440
            X2              =   4440
            Y1              =   120
            Y2              =   1215
         End
         Begin VB.Line Line2 
            BorderColor     =   &H80000005&
            Index           =   0
            X1              =   2175
            X2              =   2175
            Y1              =   120
            Y2              =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Neto Factura"
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
            Index           =   4
            Left            =   9840
            TabIndex        =   36
            Top             =   495
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Total I.G.V."
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
            Index           =   3
            Left            =   7680
            TabIndex        =   35
            Top             =   495
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Total Dctos"
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
            Index           =   2
            Left            =   5160
            TabIndex        =   34
            Top             =   495
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Total Bruto"
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
            Left            =   2760
            TabIndex        =   33
            Top             =   495
            Width           =   1215
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
            Index           =   0
            Left            =   360
            TabIndex        =   32
            Top             =   495
            Width           =   1335
         End
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Height          =   1185
         Left            =   5616
         TabIndex        =   20
         Top             =   1536
         Visible         =   0   'False
         Width           =   4695
         Begin VB.CommandButton cAceptaA 
            BackColor       =   &H0000C0C0&
            Caption         =   "&Acepta"
            Height          =   345
            Left            =   3120
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   315
            Width           =   1155
         End
         Begin VB.CommandButton cCerrarA 
            BackColor       =   &H0000C0C0&
            Caption         =   "&Cancela"
            Height          =   345
            Left            =   3120
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   690
            Width           =   1155
         End
         Begin MSMask.MaskEdBox MFSerie 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "d/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   3
            EndProperty
            Height          =   312
            Left            =   1728
            TabIndex        =   23
            Top             =   288
            Width           =   1188
            _ExtentX        =   2090
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   3
            Mask            =   "###"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MFnumero 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "d/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   3
            EndProperty
            Height          =   312
            Left            =   1728
            TabIndex        =   24
            Top             =   672
            Width           =   1188
            _ExtentX        =   2090
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   8
            Mask            =   "########"
            PromptChar      =   "_"
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00800000&
            BackStyle       =   1  'Opaque
            FillColor       =   &H0080FF80&
            FillStyle       =   0  'Solid
            Height          =   1008
            Index           =   1
            Left            =   0
            Shape           =   4  'Rounded Rectangle
            Top             =   156
            Width           =   4632
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00800000&
            BackStyle       =   1  'Opaque
            FillColor       =   &H00800000&
            FillStyle       =   0  'Solid
            Height          =   885
            Index           =   0
            Left            =   90
            Shape           =   4  'Rounded Rectangle
            Top             =   210
            Width           =   4455
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Imgrese Serie                       Nro Guia"
            ForeColor       =   &H00C0FFC0&
            Height          =   576
            Left            =   444
            TabIndex        =   25
            Top             =   372
            Width           =   1020
         End
      End
      Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
         Height          =   2895
         Left            =   150
         TabIndex        =   37
         Top             =   270
         Width           =   11475
         _ExtentX        =   20241
         _ExtentY        =   5106
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
   Begin VB.Frame Frame2 
      Height          =   1665
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11745
      Begin VB.Label Label1 
         Caption         =   "No. Pedido"
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
         Left            =   8760
         TabIndex        =   18
         Top             =   270
         Width           =   1125
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo Cambio"
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
         Index           =   8
         Left            =   4470
         TabIndex        =   17
         Top             =   1260
         Width           =   1245
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
         Index           =   0
         Left            =   330
         TabIndex        =   16
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
         Index           =   1
         Left            =   4320
         TabIndex        =   15
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
         TabIndex        =   14
         Top             =   600
         Width           =   765
      End
      Begin VB.Label Label1 
         Caption         =   "Vendedor"
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
         Index           =   4
         Left            =   330
         TabIndex        =   13
         Top             =   930
         Width           =   855
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
         Height          =   225
         Index           =   5
         Left            =   4530
         TabIndex        =   12
         Top             =   960
         Width           =   765
      End
      Begin VB.Label Label1 
         Caption         =   "Moneda"
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
         Index           =   7
         Left            =   360
         TabIndex        =   11
         Top             =   1260
         Width           =   975
      End
      Begin VB.Label Label2 
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
         TabIndex        =   10
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label2 
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
         Left            =   5700
         TabIndex        =   9
         Top             =   240
         Width           =   2475
      End
      Begin VB.Label Label2 
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
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   2
         Left            =   9810
         TabIndex        =   8
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label2 
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
         Index           =   3
         Left            =   1440
         TabIndex        =   7
         Top             =   570
         Width           =   9735
      End
      Begin VB.Label Label2 
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
         Index           =   4
         Left            =   1440
         TabIndex        =   6
         Top             =   900
         Width           =   2715
      End
      Begin VB.Label Label2 
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
         Index           =   5
         Left            =   5640
         TabIndex        =   5
         Top             =   900
         Width           =   2805
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
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
         Index           =   6
         Left            =   10080
         TabIndex        =   4
         Top             =   900
         Width           =   1068
      End
      Begin VB.Label Label2 
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
         Index           =   7
         Left            =   1440
         TabIndex        =   3
         Top             =   1230
         Width           =   2745
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
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
         Index           =   8
         Left            =   5640
         TabIndex        =   2
         Top             =   1230
         Width           =   1155
      End
      Begin VB.Label Label1 
         Caption         =   "Codigo Cliente"
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
         Index           =   6
         Left            =   8820
         TabIndex        =   1
         Top             =   960
         Width           =   1245
      End
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   1728
      Top             =   4224
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSComctlLib.StatusBar Panel 
      Align           =   2  'Align Bottom
      Height          =   312
      Left            =   0
      TabIndex        =   38
      Top             =   6864
      Width           =   11844
      _ExtentX        =   20902
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmDetallePedidoxGuias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsdeta As New ADODB.Recordset
Dim csql As New ADODB.Recordset
Dim RSDETA2 As New ADODB.Recordset
Dim adll As New dllgeneral.dll_general
Dim xtipo As String
Dim xAlma As String
Dim xDocu As String
Dim xcliente As String
Dim xmonto As Double
Dim nvalor1, nvalor2, nvalor3 As String
Dim vt_tempo As String
Dim xmoneda As String

Private Sub cBusca()
    Dim csqld As New ADODB.Recordset
    Dim acliente As New ADODB.Recordset
    Dim nsql As String
    Dim J As Integer
    
    Call Limpiartexto(MBox2, 6, 10)
    Call Limpiartexto(Label2, 0, 8)
    Call CargaGrilla
    
    nsql = " select * from movalmcab where caalma ='" & xAlma & "' and catd='" & xtipo & "'  and canumdoc='" & xDocu & "'  "
'    nvalor = ""
    Set csql = VGCNx.Execute(nsql)
    If csql.RecordCount > 0 Then
        nvalor1 = Escadena(csql!CAALMA)
        nvalor2 = Escadena(csql!CATD)
        nvalor3 = Escadena(csql!CANUMDOC)

        Set acliente = VGCNx.Execute("select * from vt_cliente where clientecodigo='" & Escadena(csql!CACODCLI) & "'")
        If acliente.RecordCount > 0 Then
           Label2(4) = Label2(4) & "-" & Escadena(acliente!clienterazonsocial)
        Else
            Label2(4) = Label2(4)
        End If
        acliente.Close
        Set acliente = Nothing
        Set acliente = VGCNx.Execute("select * from vt_almacen where almacencodigo='" & Escadena(csql!CAALMA) & "'")
        If acliente.RecordCount > 0 Then
           Label2(5) = Label2(5) & "-" & Escadena(acliente!almacendescripcion)
        Else
            Label2(5) = Label2(5)
        End If
        acliente.Close
        Set acliente = Nothing
        
    Else
        MsgBox "No existe Informacion del Documento...Verifique!!", vbInformation, MsgTitle
        csql.Close
        Set csql = Nothing
        Exit Sub
    End If
       
    Set csqld = VGCNx.Execute("select DEITEM,A.decodigo,b.adescri,b.aunidad," & _
                          "DECANTID " & _
                          "from movalmdet A inner join " & _
                          "[" & VGCNx.DefaultDatabase & "].dbo.maeart B" & _
                          " ON A.decodigo=b.acodigo COLLATE Modern_Spanish_CI_AI " & _
                          "where dealma='" & nvalor1 & "' and detd='" & nvalor2 & "' and denumdoc='" & nvalor3 & "'  ")
    
    Set rsdeta = Nothing
    Call CargaGrilla

    Do Until csqld.EOF
       rsdeta.AddNew
       rsdeta.Fields(0) = Escadena(csqld!DEITEM)
       rsdeta.Fields(1) = Escadena(csqld!decodigo)
       rsdeta.Fields(2) = Escadena(csqld!adescri)
       rsdeta.Fields(3) = Escadena(csqld!aunidad)
       rsdeta.Fields(4) = numero(csqld!DECANTID)
       rsdeta.Update
       csqld.MoveNext
    Loop
    csqld.Close
    Call ConfigGrid
End Sub

Private Sub acumulagrilla()
    Dim csql As New ADODB.Recordset
    Dim acliente As New ADODB.Recordset
    Dim nvalor1, nvalor2, nvalor3 As String
    Dim nsql As String
    Dim J As Integer
    
 '   Call Limpiartexto(MBox2, 6, 10)
 '   Call Limpiartexto(Label2, 0, 8)
    Call CargaGrillaTotal
    
'    If xtipo = g_tipobol Then
'       nsql = "select * from vt_pedido where pedidonrofact='" & xDocu & "' and pedidotipofac='" & g_tipobol & "'"
'    ElseIf xtipo = g_tipofac Then
'        nsql = "select * from vt_pedido where pedidonrofact='" & xDocu & "' and pedidotipofac='" & g_tipofac & "'"
'    ElseIf xtipo = g_tipoped Then
'        nsql = "select * from vt_pedido where pedidonumero='" & xDocu & "'"
'    Else
'        nsql = "select * from vt_pedido where pedidonrofact='" & xDocu & "' and pedidotipofac='" & xtipo & "'"
'    End If
'    nsql = " select * from movalmcab where carftdoc = 'GR' "
'    nvalor = ""
'    Set csql = VGcnx.Execute(nsql)
'    If csql.RecordCount > 0 Then
'        nvalor1 = Escadena(csql!caalma)
'        nvalor2 = Escadena(csql!catd)
'        nvalor3 = Escadena(csql!canumdoc)

 '       Label2(0) = Format(csql!pedidofecha, "dd/mm/yyyy")
'        If Not IsNull(csql!pedidonrofact) And Trim(csql!pedidonrofact) <> "0" Then
 '            Label2(1) = Escadena(csql!pedidotipofac) & "-" & Escadena(csql!pedidonrofact)
'        Else
'            Label2(0) = Format(csql!pedidofecha, "dd/mm/yyyy")
'            Label2(1) = g_tipoped & "-" & Escadena(csql!pedidonumero)
'        End If
'        Label2(2) = Escadena(csql!pedidonumero)
'        xcliente = csql!clientecodigo: xmonto = csql!pedidototneto: xmoneda = csql!pedidomoneda
'
'        Label2(3) = Escadena(csql!clientecodigo) & "-" & Escadena(csql!clienterazonsocial)
'        Label2(4) = Escadena(csql!vendedorcodigo)
'        Set acliente = VGcnx.Execute("select * from vt_cliente where clientecodigo='" & Escadena(csql!cacodcli) & "'")
'        If acliente.RecordCount > 0 Then
'           Label2(4) = Label2(4) & "-" & Escadena(acliente!clienterazonsocial)
'        Else
'            Label2(4) = Label2(4)
'        End If
'        acliente.Close
'        Set acliente = Nothing
 '       Label2(5) = Escadena(csql!almacencodigo)
 '       Set acliente = VGcnx.Execute("select * from vt_almacen where almacencodigo='" & Escadena(csql!caalma) & "'")
 '       If acliente.RecordCount > 0 Then
 '          Label2(5) = Label2(5) & "-" & Escadena(acliente!almacendescripcion)
 '       Else
 '           Label2(5) = Label2(5)
 '       End If
 '       acliente.Close
 '       Set acliente = Nothing
'        Label2(6) = Escadena(csql!clientecodigo)
' Label2(7) = Escadena(csql!pedidomoneda)
'        Set acliente = VGcnx.Execute("select * from gr_moneda where monedacodigo='" & Escadena(csql!pedidomoneda) & "'")
'        If acliente.RecordCount > 0 Then
'           Label2(7) = Label2(7) & "-" & Escadena(acliente!monedadescripcion)
'        Else
'           Label2(7) = Label2(7)
'        End If
'        acliente.Close
'        Set acliente = Nothing
'        Label2(8) = numero(csql!pedidotipcambio)
'        MBox2(6) = numero(csql!pedidototitem)
'        MBox2(7) = Format(csql!pedidototbruto, "##,###,##0.0000")
'        MBox2(8) = numero(csql!pedidomontodsctoglobal + csql!pedidomontodsctocliente + csql!pedidomontodsctoppago + csql!pedidomontodsctovtaoficina + csql!pedidototaldsctoxitem + csql!pedidototaldsctoxlinea + csql!pedidototaldsctoxprom)
'        MBox2(9) = numero(csql!pedidototimpuesto)
'        MBox2(10) = numero(csql!pedidototneto)
        
'    Else
'        MsgBox "No existe Informacion del Documento...Verifique!!", vbInformation, MsgTitle
'        csql.Close
'        Set csql = Nothing
'        Exit Sub
'    End If
'    csql.Close
    'DETD DENUMDOC    DEITEM                                               DECANTENT                                             DECANREF                                              DECANFAC                                              DEORDEN DEPREUNI                                              DEPRECIO                                              DEPRECI1                                              DEDESCTO                                              DESTOCK                                            DEIGV                                                 DEIMPMN                                               DEIMPUS                                               DESERIE              DESITUA DEFECDOC                                               DECENCOS DERFALMA DETR DEESTADO DECODMOV DEVALTOT                                              DECOMPRO DECODMON DETIPO DETIPCAM                                              DEPREVTA
    '                               DEMONVTA DEFECVEN                                               DEDEVOL                                               DESOLI DEDESCRI                                                                                             DEPORDES                                              DEIGVPOR                                              DEDESCLI                                              DEDESESP                                              DENUMFAC   DELOTE               DEUNIDAD DEEPQ                                              DEORDFAB   DEQUIPO    DECANREF1
       
    Set csql = VGCNx.Execute("select DEITEM,A.decodigo,b.adescri,b.aunidad," & _
                          "DECANTID " & _
                          "from movalmdet A inner join " & _
                          "[" & VGCNx.DefaultDatabase & "].dbo.maeart B" & _
                          " ON A.decodigo=b.acodigo COLLATE Modern_Spanish_CI_AI " & _
                          "where dealma='" & nvalor1 & "' and detd='" & nvalor2 & "' and denumdoc='" & nvalor3 & "'  ")
    
    Set rsdeta = Nothing
  '  Call CargaGrillaTotal

    Do Until csql.EOF
       RSDETA2.AddNew
       RSDETA2.Fields(0) = Escadena(csql!DEITEM)
       RSDETA2.Fields(1) = Escadena(csql!decodigo)
       RSDETA2.Fields(2) = Escadena(csql!adescri)
       RSDETA2.Fields(3) = Escadena(csql!aunidad)
       RSDETA2.Fields(4) = numero(csql!DECANTID)
       RSDETA2.Update
       csql.MoveNext
    Loop
    csql.Close
    Call ConfigGrid
    Set csql = Nothing
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

Public Function CargaGrillaTotal()

   Set RSDETA2 = Nothing
   
   
   Call RSDETA2.Fields.Append("Item", adInteger)
   Call RSDETA2.Fields.Append("Codigo", adChar, 20)
   Call RSDETA2.Fields.Append("Descripcion", adChar, 100)
   Call RSDETA2.Fields.Append("UM", adChar, 3)
   Call RSDETA2.Fields.Append("Cant", adDouble)
   Call RSDETA2.Fields.Append("Precio_Vta", adDouble)
   Call RSDETA2.Fields.Append("Dscto(%)", adDouble)
   Call RSDETA2.Fields.Append("Total", adDouble)
   Call RSDETA2.Fields.Append("%", adDouble)
   
   RSDETA2.Open
   ConfigGridTotal

End Function

Public Function ConfigGrid()

   Set TDBGrid2.DataSource = rsdeta
   With TDBGrid2
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
   TDBGrid2.Refresh
End Function

Public Function ConfigGridTotal()
   Set TDBGrid1.DataSource = RSDETA2
   With TDBGrid1
      .Columns(0).Width = 600
      .Columns(0).Caption = "Item"
      .Columns(1).Width = 1100
      .Columns(1).Caption = "Codigo"
      .Columns(2).Width = 3500
      .Columns(2).Caption = "Descripcion"
      .Columns(3).Width = 600
      .Columns(3).Caption = "UM"
      .Columns(4).Width = 1000
      .Columns(4).Caption = "Cant"
      .Columns(5).Width = 1000
      .Columns(5).Caption = "Precio_Vta"
      .Columns(6).Width = 1000
      .Columns(6).Caption = "Dscto(%)"
      .Columns(7).Width = 1000
      .Columns(7).Caption = "Total"
      .Columns(8).Width = 1000
      .Columns(8).Caption = "%"
      .Columns(5).NumberFormat = "###,##0.0000"
      .Columns(6).NumberFormat = "###,##0.00"
      .Columns(7).NumberFormat = "###,##0.0000"
      .Columns(8).NumberFormat = "###,##0.00"
   End With
   TDBGrid1.Refresh

End Function

Private Sub cAceptaA_Click()
  Dim ntipo, nnume As String
  Dim rs As New ADODB.Recordset
  Dim acmd As New ADODB.Command
    If adll.ComboDato(Label2(1)) = g_tipofac Then
        ntipo = g_tipofac
        nnume = Mid(Label2(1), Len(g_tipofac) + 2, Len(Trim(Label2(1))))
    ElseIf adll.ComboDato(Label2(1)) = g_tipobol Then
        ntipo = g_tipobol
        nnume = Mid(Label2(1), Len(g_tipobol) + 2, Len(Trim(Label2(1))))
    ElseIf adll.ComboDato(Label2(1)) = g_tipoguia Then
        ntipo = g_tipoguia
        nnume = Mid(Label2(1), Len(g_tipoguia) + 2, Len(Trim(Label2(1))))
    Else
       ntipo = Left(Label2(1), 2)
       nnume = Mid(Label2(1), Len(g_tipofac) + 2, Len(Trim(Label2(1))))
    End If
     imprimirguias
     Frame3.Visible = False
End Sub

Private Sub imprimirguias()

Dim nguia As String
Dim ntabla As String
Dim busca As New dll_apisgen.dll_apis
Dim rb As New ADODB.Recordset
Dim rb1 As New ADODB.Recordset
Dim contador As Double
Dim contador1 As Double
Dim numguias As Integer, TCant As Integer, nflag As Integer
Dim SQL As String
Dim inicio As Integer
Dim fin As Integer
Dim J As Integer
Dim numero As String, razonsocial As String
Dim ruc As String, direccion As String, distrito As String
Dim num_guias As String

ntabla = "vt_detallepedido"
contador = 0

VGCNx.Execute "delete from gtempfile"
VGCNx.Execute "delete from tempfile"
VGCNx.Execute "INSERT into gtempfile" & _
         " Select a.detpedcantpedida,a.productocodigo,b.adescri,(a.detpedimpbruto/a.detpedcantpedida),a.detpedimpbruto,a.detpeddsctoxitem,isnull(a.detpedcantpedidaref,0), case ltrim(rtrim(a.productocodigo)) when '000' then '' else a.unidadcodigo end" & _
         " From " & ntabla & " A inner join " & _
         "[" & VGCNx.DefaultDatabase & "].dbo.maeart B" & _
         " ON A.productocodigo=b.acodigo COLLATE Modern_Spanish_CI_AI " & _
         " Where pedidonumero='" & CStr(Label2(2)) & "'"

Set rb1 = VGCNx.Execute("select * from vt_cliente where clientecodigo='" & Label2(6) & "' ")
If rb1.RecordCount > 0 Then
   razonsocial = Escadena(rb1!clienterazonsocial)
   ruc = Escadena(rb1!clienteruc)
   direccion = Escadena(rb1!clientedireccion)
   distrito = Escadena(rb1!clientedistrito)
End If
rb1.Close
Set rb1 = Nothing

Set rb = VGCNx.Execute("select * from gtempfile inner join maeart on productocodigo=acodigo order by afamilia,alinea,agrupo,acodigo")
If rb.RecordCount > 0 Then
   If rb.RecordCount Mod 50 > 0 Then
       numguias = Int(rb.RecordCount / 50) + 1
    Else
        numguias = Int(rb.RecordCount / 50)
   End If
   numero = MFnumero
   rb.MoveFirst
   Do While contador < numguias
          numero = Right("000000000" + Trim(Str(Val(MFnumero) + contador)), 9)
         contador = contador + 1
          inicio = (contador - 1) * 50 + 1
          If contador * 50 > rb.RecordCount Then
             fin = rb.RecordCount
           Else
             fin = contador * 50
          End If
      
          nguia = Right("000000000000" & Trim(MFSerie) + Trim(numero), 12)
          num_guias = num_guias + nguia + "/"
          VGCNx.Execute "UPDATE vt_pedido set pedidoobserva='" & RTrim(num_guias) & "'" & _
               " Where pedidonumero='" & Label2(2) & "'"
          contador1 = 0
          If fin > rb.RecordCount Then
             fin = rb.RecordCount - inicio
          End If
          VGCNx.Execute "delete from gtempfile2filas"
          For J = inicio To fin
                 contador1 = contador1 + 1
                 If contador1 <= 25 Then
                     SQL = "INSERT INTO gtempfile2filas(item,producto1,descripcion1,cantidad1,importe1,"
                     SQL = SQL & "cantidad2,importe2) "
                     SQL = SQL & " VALUES ( '" & contador1 & "','" & RTrim(rb!productocodigo) & "','" & RTrim(rb!productodescripcion) & "','" & rb!detpedcantpedida & "','" & rb!detpedimpbruto & "',0,0)"
                  Else
                     TCant = contador1 - 25
                      SQL = "UPDATE gtempfile2filas set producto2 ='" & RTrim(rb!productocodigo) & "',"
                      SQL = SQL & " descripcion2='" & RTrim(rb!productodescripcion) & "',"
                      SQL = SQL & "cantidad2='" & rb!detpedcantpedida & "',"
                        SQL = SQL & "importe2= '" & rb!detpedimpbruto & "'"
                        SQL = SQL & " where item = " & TCant & ""
                 End If
                 VGCNx.Execute SQL
                 rb.MoveNext
          Next J
          CrystalReport1.Reset
          CrystalReport1.ReportFileName = VGParamSistem.Rutareport & "Repguiaimpresa.rpt"
          CrystalReport1.LogOnServer "pdssql.dll", _
                   busca.LeerIni(App.Path & "\MARFICE.INI", "CONEXION", "SERVIDOR", ""), _
                    busca.LeerIni(App.Path & "\MARFICE.INI", "CONEXION", "BDDATOS", ""), _
                    busca.LeerIni(App.Path & "\MARFICE.INI", "CONEXION", "USUARIO", ""), _
                    busca.LeerIni(App.Path & "\MARFICE.INI", "CONEXION", "PASSW", "")
          CrystalReport1.Connect = _
                   "DSN=" & busca.LeerIni(App.Path & "\MARFICE.INI", "CONEXION", "SERVIDOR", "") & ";" & _
                   "DSQ=" & busca.LeerIni(App.Path & "\MARFICE.INI", "CONEXION", "BDDATOS", "") & ";" & _
                   "UID=" & busca.LeerIni(App.Path & "\MARFICE.INI", "CONEXION", "USUARIO", "") & ";" & _
                   "PWD=" & busca.LeerIni(App.Path & "\MARFICE.INI", "CONEXION", "PASSW", "")
                
          CrystalReport1.Destination = crptToWindow
          CrystalReport1.WindowState = crptMaximized
          CrystalReport1.DiscardSavedData = True
          With CrystalReport1
                   .formulas(0) = "nro='" & Label2(1) & "'"
                   .formulas(1) = "cliente='" & razonsocial & "'"
                   .formulas(2) = "fecha='" & CStr(Day(CDate(Label2(0)))) & "     " & adll.DesMes(Month(CDate(Label2(0)))) & "                       " & Right(CStr(Year(CDate(Label2(0)))), 4) & "'"
                   .formulas(3) = "direccion='" & direccion & "'"
                   .formulas(4) = "dni='" & ruc & "'"
                   .formulas(5) = "opedido='" & Label2(0) & "'"
              '     .Formulas(6) = "ocompra='" & MBox(17) & "'"
                   .formulas(7) = "guia='" & nguia & "'"
                   .formulas(8) = "distrito='" & distrito & "'"
                   .formulas(9) = "destino='" & direccion & "'"
                   Set rb1 = VGCNx.Execute("select * from gr_empresa where empresacodigo='" & VGParametros.empresacodigo & "'")
                   If rb1.RecordCount > 0 Then
                      .formulas(10) = "partida='" & Escadena(rb1!empresadireccion) & "'"
                    Else
                      .formulas(10) = "partida=''"
                   End If
                    If .Status <> 2 Then .Action = 1
          End With
          SQL = nguia
          MsgBox "Proceda a imprimir la GUIA DE REMISION .", vbInformation, SQL
    Loop
End If
rb.Close

  
nerror:
   If Err Then
      If nflag = 1 Then
         VGCNx.RollbackTrans
      End If
      MsgBox "Comunicarse con Sistemas ...!!" & Chr(13) & Chr(10) & Err.Number & "-" & Err.Description
      Exit Sub
   End If
  
End Sub

Private Sub cCerrarA_Click()
  Frame3.Visible = False
End Sub

Private Sub Command2_Click()

End Sub

Private Sub Form_Load()
    vt_tempo = "##" & ComputerName & "vt_p" & g_ptoventa
    MostrarFormVentas Me, "C"
    DoEvents
    Call Limpiartexto(MBox2, 6, 10)
    Call CargaGrilla
    Call cBusca
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  Set rsdeta = Nothing
End Sub

Public Property Let Balma(pdata)
  xAlma = pdata
End Property

Public Property Let Btipo(pdata)
  xtipo = pdata
End Property
Public Property Let BNumero(pdata)
  xDocu = pdata
End Property

Private Sub cmdBotones_Click(Index As Integer)
  On Error GoTo nerror
  Select Case Index
  Case 11
    Frame3.Visible = True
    Frame1.Visible = True
'    MFSerie = Format(Date, "dd/mm/yyyy")
    acumulaguias
    Unload Me
  Case 12
    Frame3.Visible = False
    Unload Me
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
Dim SQL As String
    SQL = " Insert " & vt_tempo & " (vt_tipdoc,vt_numdoc,clientecodigo,clienterazonsocial,documentoreferencia,numeroreferencia,almacencodigo) "
    SQL = SQL & " values( '" & Escadena(csql!carftdoc) & "', '" & Escadena(csql!carfndoc) & "','" & csql!CACODCLI & "' , "
    SQL = SQL & " '" & csql!CANOMCLI & "', '" & csql!CATD & "', '" & csql!CANUMDOC & "','" & csql!CAALMA & "' )"
    VGCNx.Execute SQL

End Sub

Private Sub TDBGrid1_GotFocus()
  Frame3.Visible = False
End Sub
