VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.UserControl UserControl1 
   ClientHeight    =   8505
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9615
   EditAtDesignTime=   -1  'True
   HitBehavior     =   0  'None
   ScaleHeight     =   8505
   ScaleWidth      =   9615
   Begin VB.Frame frmbotones 
      Height          =   1050
      Left            =   2070
      TabIndex        =   0
      Top             =   7245
      Width           =   5655
      Begin VB.CommandButton cmdBotones 
         Caption         =   "&Nuevo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Index           =   0
         Left            =   225
         Picture         =   "UserControl1.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   180
         Width           =   915
      End
      Begin VB.CommandButton cmdBotones 
         Caption         =   "E&ditar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Index           =   1
         Left            =   1305
         Picture         =   "UserControl1.ctx":0442
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   180
         Width           =   915
      End
      Begin VB.CommandButton cmdBotones 
         Caption         =   "&Eliminar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Index           =   2
         Left            =   2385
         Picture         =   "UserControl1.ctx":0884
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   180
         Width           =   870
      End
      Begin VB.CommandButton cmdBotones 
         Caption         =   "&Salir"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Index           =   4
         Left            =   4590
         Picture         =   "UserControl1.ctx":0CC6
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   180
         Width           =   915
      End
      Begin VB.CommandButton cmdBotones 
         Caption         =   "&Imprimir"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Index           =   3
         Left            =   3510
         Picture         =   "UserControl1.ctx":1108
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   180
         Width           =   870
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6720
      Left            =   315
      TabIndex        =   6
      Top             =   360
      Width           =   8925
      _ExtentX        =   15743
      _ExtentY        =   11853
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Consulta"
      TabPicture(0)   =   "UserControl1.ctx":154A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "TDBGrid1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Mantenimiento"
      TabPicture(1)   =   "UserControl1.ctx":1566
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(1)=   "cAcepta"
      Tab(1).Control(2)=   "cCancela"
      Tab(1).ControlCount=   3
      Begin VB.Frame Frame1 
         Height          =   5415
         Left            =   -74325
         TabIndex        =   9
         Top             =   540
         Width           =   7065
         Begin VB.TextBox txt 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   2790
            TabIndex        =   10
            Top             =   405
            Width           =   2040
         End
         Begin VB.TextBox txt 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   2790
            TabIndex        =   12
            Top             =   855
            Width           =   4020
         End
         Begin VB.TextBox txt 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   2
            Left            =   2790
            TabIndex        =   14
            Top             =   1305
            Width           =   4020
         End
         Begin VB.TextBox txt 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   3
            Left            =   2790
            TabIndex        =   16
            Top             =   1800
            Width           =   4020
         End
         Begin VB.TextBox txt 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   4
            Left            =   2790
            TabIndex        =   18
            Top             =   2295
            Width           =   4020
         End
         Begin VB.TextBox txt 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Index           =   5
            Left            =   2790
            TabIndex        =   20
            Top             =   2790
            Width           =   4065
         End
         Begin VB.TextBox txt 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Index           =   6
            Left            =   2790
            TabIndex        =   22
            Top             =   3375
            Width           =   4065
         End
         Begin VB.TextBox txt 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   7
            Left            =   2790
            TabIndex        =   24
            Top             =   3960
            Width           =   4065
         End
         Begin VB.TextBox txt 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   8
            Left            =   2790
            TabIndex        =   26
            Top             =   4500
            Width           =   4065
         End
         Begin VB.TextBox txt 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   9
            Left            =   2790
            TabIndex        =   28
            Top             =   4950
            Width           =   4065
         End
         Begin VB.Label lbl 
            Caption         =   "lbl0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   180
            TabIndex        =   29
            Top             =   495
            Width           =   2400
         End
         Begin VB.Label lbl 
            Caption         =   "lbl1"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   180
            TabIndex        =   27
            Top             =   900
            Width           =   2355
         End
         Begin VB.Label lbl 
            Caption         =   "lbl2"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   2
            Left            =   180
            TabIndex        =   25
            Top             =   1350
            Width           =   2445
         End
         Begin VB.Label lbl 
            Caption         =   "lbl3"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   3
            Left            =   180
            TabIndex        =   23
            Top             =   1845
            Width           =   2535
         End
         Begin VB.Label lbl 
            Caption         =   "lbl4"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   4
            Left            =   180
            TabIndex        =   21
            Top             =   2385
            Width           =   2535
         End
         Begin VB.Label lbl 
            Caption         =   "lbl5"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   5
            Left            =   180
            TabIndex        =   19
            Top             =   2925
            Width           =   2490
         End
         Begin VB.Label lbl 
            Caption         =   "lbl6"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   6
            Left            =   180
            TabIndex        =   17
            Top             =   3465
            Width           =   2445
         End
         Begin VB.Label lbl 
            Caption         =   "lbl7"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   7
            Left            =   180
            TabIndex        =   15
            Top             =   4005
            Width           =   2625
         End
         Begin VB.Label lbl 
            Caption         =   "lbl8"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   8
            Left            =   225
            TabIndex        =   13
            Top             =   4545
            Width           =   2445
         End
         Begin VB.Label lbl 
            Caption         =   "lbl9"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   9
            Left            =   225
            TabIndex        =   11
            Top             =   4995
            Width           =   2445
         End
      End
      Begin VB.CommandButton cAcepta 
         Caption         =   "&Aceptar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -72525
         TabIndex        =   8
         Top             =   6075
         Width           =   1335
      End
      Begin VB.CommandButton cCancela 
         Caption         =   "&Cancelar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -70680
         TabIndex        =   7
         Top             =   6075
         Width           =   1335
      End
      Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
         Height          =   5955
         Left            =   225
         TabIndex        =   30
         Top             =   540
         Width           =   8340
         _ExtentX        =   14711
         _ExtentY        =   10504
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
         Splits(0).DividerColor=   13160660
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=2"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2752"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=2752"
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
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=975,.italic=0"
         _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=Arial"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
         _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=975,.italic=0"
         _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(12)  =   ":id=2,.fontname=Arial"
         _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=975,.italic=0"
         _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(15)  =   ":id=3,.fontname=Arial"
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
Attribute VB_Name = "UserControl1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Option Base 0
Dim m_nombretabla As String
Dim m_Condicion As String
Dim m_Orden As String
Dim uTitle As String
Dim uCant As Integer
Dim cdb As New ADODB.Connection
'****************************************************************
Dim modoinsert As Boolean
Dim modoedit As Boolean
Dim i_filaorigen As Integer
'****************************************************************
Public g_usuario As String
Dim s_cadenacampos As String
Dim s_cadenaclaves As String
Dim s_cadenacamposvisibles As String
Dim a_Arreglo(0 To 8, 0 To 20)

Public Property Let Arreglo(ByRef valor)
  Dim k, j As Integer
  For k = 0 To 8
    For j = 0 To 20
       a_Arreglo(k, j) = valor(k, j)
    Next j
  Next k
PropertyChanged "Arreglo"
End Property

Public Property Let NombreTabla(ByVal valor As String)
  m_nombretabla = valor
  PropertyChanged "nombretabla"
End Property

Public Property Let Title(ByVal valor As String)
    uTitle = valor
    PropertyChanged "Title"
End Property

Public Property Let NumCampos(ByVal valor As Integer)
    uCant = valor
    PropertyChanged "NumCampos"
End Property

Public Property Let Condicion(ByVal valor As String)
  m_Condicion = valor
  PropertyChanged "Condicion"
End Property

Public Property Let Orden(ByVal valor As String)
  m_Orden = valor
  PropertyChanged "Orden"
End Property

Public Property Let Conexion(valor As ADODB.Connection)
   Set cdb = valor
   PropertyChanged "Conexion"
End Property
Public Function cargar_datos()
  Dim sql As String
  Dim rs As New ADODB.Recordset
  Dim i As Integer
     
  If Len(Trim(s_cadenacamposvisibles)) > 0 And Len(Trim(m_nombretabla)) > 0 Then
     sql = "SELECT " & s_cadenacamposvisibles & " FROM " & m_nombretabla
     Set rs = cdb.Execute(sql)
     Set TDBGrid1.DataSource = rs
    
     For i = 0 To TDBGrid1.Columns.Count - 1
        If Len(a_Arreglo(1, i)) > a_Arreglo(3, i) Then
            TDBGrid1.Columns(i).Width = Len(a_Arreglo(1, i)) * 120
        Else
            TDBGrid1.Columns(i).Width = a_Arreglo(3, i) * 120
        End If
     Next i
     TDBGrid1.Refresh
    
     UserControl.Refresh
  End If
  Set rs = Nothing
  SSTab1.Tab = 0
  
End Function

'''''   UPDATE DE CAMPOS CLAVE:

'            s_set = Null
'            s_cadenaclaves = Null
'            For j = 0 To (UBound(a_Arreglo, 1) + 1)
'              If (a_Arreglo(0, j) <> "") Then            'si existe campo
'                If a_Arreglo(4, j) = False Then          'si no es campo clave:
'                    If (a_Arreglo(5, j) = "") Then       'si no existe valor ingresado por el sistema
'                       If (a_Arreglo(2, j) = "C" Or a_Arreglo(2, j) = "D") Then   'si es tipo char
'                          s_set = s_set & a_Arreglo(0, j) & "='" & Trim(txt(j)) & "',"
'                       Else
'                          s_set = s_set & a_Arreglo(0, j) & "=" & txt(j) & ","
'                       End If
'                    Else
'                       If (a_Arreglo(2, j) = "C" Or a_Arreglo(2, j) = "D") Then  ' si es tipo char
'                          s_set = s_set & a_Arreglo(0, j) & "='" & a_Arreglo(5, j) & "',"
'                       Else
'                           s_set = s_set & a_Arreglo(0, j) & "=" & a_Arreglo(5, j) & ","
'                       End If
'                    End If
'                End If
'              End If
'            Next j
'            s_set = Left(s_set, Len(Trim(s_set)) - 1)
'
'            sql = "Update " & m_nombretabla & _
'                     " Set " & s_set & " Where " & s_cadenaclaves
'            cdb.Execute sql

Private Sub cCancela_Click()
    SSTab1.TabEnabled(0) = True
    SSTab1.Tab = 0
    SSTab1.SetFocus
    frmbotones.Visible = True
    '''''''''
      modoinsert = False
      modoedit = False
      i_filaorigen = -1
    '''''''''
End Sub

Private Sub cmdBotones_Click(Index As Integer)
  Dim j As Integer
  Dim spos As Integer
  Dim sql As String
  
  On Error GoTo nerror
  '''''
  SSTab1.TabEnabled(1) = True
  '''''
  uCant = 1
  Select Case Index
     Case 0   'nuevo
        SSTab1.Tab = 1
        txt(0).SetFocus
        '''''
        Limpia_textos
        frmbotones.Visible = False
        '''''''''
        modoinsert = True
        '''''''''
        
     Case 1   'modificar
        
               For j = 0 To (UBound(a_Arreglo, 1) + 1)
                  If a_Arreglo(0, j) <> "" Then       ' si existe campo
                     If a_Arreglo(1, j) <> "" Then    ' si es visible
                        txt(j) = Trim(TDBGrid1.Columns(j).Text)
                     End If
                  End If
                Next j
                
        frmbotones.Visible = False
        SSTab1.Tab = 1
        txt(1).SetFocus
         '''''''''
        i_filaorigen = TDBGrid1.Row
        modoedit = True
        Obtener_Claves (1)
        '''''''''
      
     Case 2   'eliminar
       If MsgBox("Desea eliminar el registro?", vbYesNo + vbDefaultButton2, "Mensaje") = vbYes Then
          Obtener_Claves (2)
          sql = "Delete From " & m_nombretabla & " where " & s_cadenaclaves
          cdb.Execute sql
          cargar_datos
       End If
        
     Case 3   'imprimir
        If uCant > 0 Then
            Printer.Print Tab((60 - Len(uTitle)) / 2); UCase(uTitle)
        
            For j = 0 To uCant
                Printer.Print Left(lbl(j) & Space(13), 13); Tab(15); ":"; Tab(18); Left(txt(j) & Space(30), 30)
            Next j
            Printer.EndDoc
        End If
     Case 4  ' salir
       Unload Parent
  End Select
   
nerror:
   If Err Then
      Err = 0
      Resume Next
   End If
   
End Sub

'**********************************************************************
Public Function Limpia_textos()
 'Dim OBJ As Object
 '  For Each OBJ In n_form.Controls
 '     If TypeOf OBJ Is TextBox Then OBJ.Text = ""
 'Next
 Dim j As Integer
   For j = 0 To (UBound(a_Arreglo, 2) + 1)
      txt(j) = ""
   Next j
End Function

Private Sub cAcepta_Click()

   Dim rs As New ADODB.Recordset
   Dim spos As Integer
   Dim sql As String
   On Error GoTo nerror
   '''''
   Dim i_cont As Integer
   Dim j As Integer
   Dim s_set As String
   Dim s_where As String
   Dim s_nombrescampos As String
   Dim a_MainArr As Variant
   Dim s_value As Variant
   ''''''''
   SSTab1.TabEnabled(0) = True
   ''''''''
   
   If modoinsert = True Then
         If Validar_CodigosDuplicados(-1) = True Then
            MsgBox "Código ya existe", vbCritical, "Error"
            cAcepta.Enabled = False
            Exit Sub
          End If
       
          s_value = Null
          For j = 0 To (UBound(a_Arreglo, 1) + 1)
              If (a_Arreglo(0, j) <> "") Then                  'si existe campo
                  If (a_Arreglo(5, j) = "") Then               'si no existe valor ingresado por el sistema
                      If (a_Arreglo(2, j) = "C" Or a_Arreglo(2, j) = "D") Then   'si es tipo char
                          s_value = s_value & "'" & Trim(txt(j)) & "',"
                       Else
                          s_value = s_value & txt(j) & ","
                      End If
                   Else
                      If (a_Arreglo(2, j) = "C" Or a_Arreglo(2, j) = "D") Then   ' si no es tipo char
                          s_value = s_value & "'" & a_Arreglo(5, j) & "',"
                      Else
                          s_value = s_value & a_Arreglo(5, j) & ","
                      End If
                   End If
               End If
           Next j
           s_value = Left(s_value, Len(Trim(s_value)) - 1)
               
          sql = "Insert Into " & m_nombretabla & _
               "(" & s_cadenacampos & ")" & " Values (" & s_value & ")"
          cdb.Execute sql
                   
   ElseIf modoedit = True Then
   
             If Validar_CodigosDuplicados(i_filaorigen) = True Then
               MsgBox "Código ya existe", vbCritical, "Error"
               cAcepta.Enabled = False
               Exit Sub
             End If
   
            s_set = ""
            's_cadenaclaves = ""
            
'            For j = 0 To (UBound(a_Arreglo, 1) + 1)
'              If (a_Arreglo(0, j) <> "") Then            'si existe campo
'                If a_Arreglo(4, j) = False Then          'si no es campo clave:
'                    If (a_Arreglo(5, j) = "") Then       'si no existe valor ingresado por el sistema
'                       If (a_Arreglo(2, j) = "C" Or a_Arreglo(2, j) = "D") Then   'si es tipo char
'                          s_set = s_set & a_Arreglo(0, j) & "='" & Trim(txt(j)) & "',"
'                       Else
'                          s_set = s_set & a_Arreglo(0, j) & "=" & txt(j) & ","
'                       End If
'                    Else
'                       If (a_Arreglo(2, j) = "C" Or a_Arreglo(2, j) = "D") Then  ' si es tipo char
'                          s_set = s_set & a_Arreglo(0, j) & "='" & a_Arreglo(5, j) & "',"
'                       Else
'                           s_set = s_set & a_Arreglo(0, j) & "=" & a_Arreglo(5, j) & ","
'                       End If
'                    End If
'                End If
'              End If
'            Next j
'            s_set = Left(s_set, Len(Trim(s_set)) - 1)

             For j = 0 To (UBound(a_Arreglo, 1) + 1)
              If (a_Arreglo(0, j) <> "") Then            'si existe campo
                
                    If (a_Arreglo(5, j) = "") Then       'si no existe valor ingresado por el sistema
                       If (a_Arreglo(2, j) = "C" Or a_Arreglo(2, j) = "D") Then   'si es tipo char
                          s_set = s_set & a_Arreglo(0, j) & "='" & Trim(txt(j)) & "',"
                       Else
                          s_set = s_set & a_Arreglo(0, j) & "=" & txt(j) & ","
                       End If
                    Else
                       If (a_Arreglo(2, j) = "C" Or a_Arreglo(2, j) = "D") Then  ' si es tipo char
                          s_set = s_set & a_Arreglo(0, j) & "='" & a_Arreglo(5, j) & "',"
                       Else
                           s_set = s_set & a_Arreglo(0, j) & "=" & a_Arreglo(5, j) & ","
                       End If
                    End If
                
              End If
            Next j
            s_set = Left(s_set, Len(Trim(s_set)) - 1)
                          
            sql = "Update " & m_nombretabla & _
                     " Set " & s_set & " Where " & s_cadenaclaves
            cdb.Execute sql
              
 '******************************************************************************************
        
 End If
 rs.Close
 Set rs = Nothing
 TDBGrid1.Refresh
      
 cargar_datos
 frmbotones.Visible = True
 UserControl.Refresh
 '''''''''
      modoinsert = False
      modoedit = False
      i_filaorigen = -1
 '''''''''
nerror:
   If Err Then
      Err = 0
      Resume Next
   End If
     
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
   SSTab1.TabEnabled(PreviousTab) = False
   cAcepta.Enabled = False
End Sub

Private Sub txt_Change(Index As Integer)
 cAcepta.Enabled = Validar_Ingreso()
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    cAcepta.Enabled = Validar_Ingreso()
End Sub

Private Sub UserControl_Initialize()
   SSTab1.TabEnabled(1) = False
   cAcepta.Enabled = False
   g_usuario = "elozano"
End Sub

Public Function Setear_Controles()
Dim j As Integer

      For j = 0 To (UBound(a_Arreglo, 1) + 1)
            If (a_Arreglo(1, j) <> "") Then
               lbl(j).Visible = True
               lbl(j).Caption = a_Arreglo(1, j)
               txt(j).Visible = True
               txt(j).MaxLength = a_Arreglo(3, j)
               Parent.Caption = "Mantenimiento de " & StrConv(m_nombretabla, vbProperCase)
            Else
               lbl(j).Visible = False
               lbl(j).Caption = ""
               txt(j).Visible = False
            End If
      Next j
             
End Function

Public Function Obtener_Campos()
Dim j As Integer

 s_cadenacampos = ""
 s_cadenacamposvisibles = ""
        
    For j = 0 To (UBound(a_Arreglo, 1) + 1)
       If a_Arreglo(0, j) <> "" Then     ' si existe campo
          s_cadenacampos = Trim(s_cadenacampos) & Trim(a_Arreglo(0, j)) & ","
            If a_Arreglo(1, j) <> "" Then
               s_cadenacamposvisibles = Trim(s_cadenacamposvisibles) & _
               Trim(a_Arreglo(0, j)) & " AS '" & Trim(a_Arreglo(1, j)) & "' ,"
            End If
        End If
    Next j
    s_cadenacampos = Left(s_cadenacampos, Len(Trim(s_cadenacampos)) - 1)
    s_cadenacamposvisibles = Left(s_cadenacamposvisibles, Len(Trim(s_cadenacamposvisibles)) - 1)
               
End Function

Private Function Obtener_Claves(tipooperacion As Integer)
Dim j As Integer

 s_cadenaclaves = ""
 
    For j = 0 To (UBound(a_Arreglo, 1) + 1)
      If (a_Arreglo(0, j) <> "") Then        ' si existe campo
         If a_Arreglo(4, j) = True Then      ' si es campo clave
            If a_Arreglo(2, j) = "C" Or a_Arreglo(2, j) = "D" Then    ' si es tipo char
                 Select Case tipooperacion
                   Case 1 'Update
                   s_cadenaclaves = s_cadenaclaves & a_Arreglo(0, j) & "='" & Trim(txt(j)) & "' And"
                   Case 2 'Delete
                   TDBGrid1.Col = j
                   s_cadenaclaves = s_cadenaclaves & a_Arreglo(0, j) & "='" & TDBGrid1.Text & "' And"
                 End Select
            Else
                 Select Case tipooperacion
                 Case 1  'Update
                   s_cadenaclaves = s_cadenaclaves & a_Arreglo(0, j) & "=" & txt(j) & " And"
                 Case 2  'Delete
                  TDBGrid1.Col = j
                  s_cadenaclaves = s_cadenaclaves & a_Arreglo(0, j) & "=" & TDBGrid1.Text & " And"
                 End Select
            End If
         End If
      End If
   Next j
   s_cadenaclaves = Left(s_cadenaclaves, Len(Trim(s_cadenaclaves)) - 3)
               
End Function

Private Function Validar_Ingreso() As Boolean
Dim j As Integer

               For j = 0 To (UBound(a_Arreglo, 1) + 1)
                If (a_Arreglo(0, j) <> "") Then             ' si existe campo
                     If a_Arreglo(1, j) <> "" Then          ' si es visible
                         If a_Arreglo(6, j) = False Then    ' si no permite nulos
                               If Trim(txt(j)) = "" Then
                                 Validar_Ingreso = False
                                 Exit Function
                               End If
                        End If
                     End If
                 End If
               Next j
               Validar_Ingreso = True

End Function

Private Function Validar_CodigosDuplicados(filaorigen As Integer) As Boolean
Dim j As Integer
Dim i As Integer
Dim fila As Integer

               fila = -1
               Validar_CodigosDuplicados = False
               For j = 0 To (UBound(a_Arreglo, 1) + 1)
                If (a_Arreglo(0, j) <> "") Then             ' si existe campo
                    If a_Arreglo(4, j) = True Then          ' si es clave
                        TDBGrid1.Col = j
                        TDBGrid1.MoveFirst
                        Do Until TDBGrid1.EOF
                            If (Trim(txt(j)) = Trim(TDBGrid1.Text)) And _
                               (fila = -1 Or TDBGrid1.Row = fila) And _
                                  (TDBGrid1.Row <> filaorigen) Then
                                  
                                     fila = TDBGrid1.Row
                                     Validar_CodigosDuplicados = True
                                     Exit Do
                                     
                            End If
                            TDBGrid1.MoveNext
                            If TDBGrid1.EOF = True Then
                              Validar_CodigosDuplicados = False
                            End If
                        Loop
                    End If
                 End If
               Next j
End Function

