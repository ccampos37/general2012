VERSION 5.00
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form FrmCopiaPedido 
   BackColor       =   &H00C9955A&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Copia Documentos"
   ClientHeight    =   8760
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12165
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8760
   ScaleWidth      =   12165
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox aBusca 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   1
      Left            =   5040
      MaxLength       =   8
      TabIndex        =   2
      Top             =   120
      Width           =   1740
   End
   Begin VB.TextBox aBusca 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   0
      Left            =   3870
      MaxLength       =   3
      TabIndex        =   1
      Top             =   120
      Width           =   840
   End
   Begin VB.CommandButton cBusca 
      BackColor       =   &H80000009&
      Caption         =   "&Buscar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1140
      Left            =   4635
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7470
      Width           =   1380
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1755
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   135
      Width           =   1965
   End
   Begin VB.CommandButton cCopia 
      BackColor       =   &H80000009&
      Caption         =   "&Copia a Pedido"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1140
      Left            =   6375
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7470
      Width           =   1380
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   4635
      Left            =   135
      TabIndex        =   24
      Top             =   2640
      Width           =   11805
      Begin VB.Frame Fr2 
         BackColor       =   &H00C9955A&
         BorderStyle     =   0  'None
         Height          =   885
         Index           =   2
         Left            =   120
         TabIndex        =   25
         Top             =   3630
         Width           =   11535
         Begin MSMask.MaskEdBox MBox2 
            Height          =   375
            Index           =   6
            Left            =   300
            TabIndex        =   26
            Top             =   240
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
            TabIndex        =   27
            Top             =   240
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
            TabIndex        =   28
            Top             =   240
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
            TabIndex        =   29
            Top             =   240
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
            TabIndex        =   30
            Top             =   240
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
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   35
            Top             =   675
            Width           =   1335
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
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   1
            Left            =   2760
            TabIndex        =   34
            Top             =   680
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
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   2
            Left            =   5160
            TabIndex        =   33
            Top             =   680
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
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   3
            Left            =   7680
            TabIndex        =   32
            Top             =   680
            Width           =   1215
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
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   4
            Left            =   9840
            TabIndex        =   31
            Top             =   680
            Width           =   1215
         End
         Begin VB.Line Line2 
            BorderColor     =   &H80000005&
            Index           =   0
            X1              =   2175
            X2              =   2175
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
            Index           =   2
            X1              =   6960
            X2              =   6960
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
            BorderColor     =   &H00808080&
            Index           =   4
            X1              =   2160
            X2              =   2160
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
            Index           =   6
            X1              =   6940
            X2              =   6940
            Y1              =   120
            Y2              =   1215
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00808080&
            Index           =   7
            X1              =   9340
            X2              =   9340
            Y1              =   120
            Y2              =   1215
         End
      End
      Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
         Height          =   3225
         Left            =   150
         TabIndex        =   36
         Top             =   270
         Width           =   11475
         _ExtentX        =   20241
         _ExtentY        =   5689
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
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HFFFFFF&,.bold=0,.fontsize=825"
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
      BackColor       =   &H00C9955A&
      BorderStyle     =   0  'None
      Height          =   1875
      Left            =   135
      TabIndex        =   5
      Top             =   765
      Width           =   11805
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   0
         Left            =   330
         TabIndex        =   23
         Top             =   270
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nro. Doc :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   4500
         TabIndex        =   22
         Top             =   270
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nro. Pedido :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   8670
         TabIndex        =   21
         Top             =   270
         Width           =   1035
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   3
         Left            =   330
         TabIndex        =   20
         Top             =   660
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vendedor :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   4
         Left            =   330
         TabIndex        =   19
         Top             =   1050
         Width           =   900
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Almacen :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   5
         Left            =   4485
         TabIndex        =   18
         Top             =   1080
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lista Precios :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   6
         Left            =   8685
         TabIndex        =   17
         Top             =   1080
         Width           =   1155
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Moneda :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   7
         Left            =   330
         TabIndex        =   16
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo Cambio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   8
         Left            =   4470
         TabIndex        =   15
         Top             =   1440
         Width           =   1035
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   0
         Left            =   1440
         TabIndex        =   14
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   1
         Left            =   5610
         TabIndex        =   13
         Top             =   240
         Width           =   2475
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   2
         Left            =   9855
         TabIndex        =   12
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   3
         Left            =   1440
         TabIndex        =   11
         Top             =   630
         Width           =   9735
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   4
         Left            =   1440
         TabIndex        =   10
         Top             =   1020
         Width           =   2715
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   5
         Left            =   5640
         TabIndex        =   9
         Top             =   1020
         Width           =   2805
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   6
         Left            =   9900
         TabIndex        =   8
         Top             =   1020
         Width           =   495
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   7
         Left            =   1440
         TabIndex        =   7
         Top             =   1410
         Width           =   2745
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   8
         Left            =   5640
         TabIndex        =   6
         Top             =   1410
         Width           =   1155
      End
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   105
      Left            =   90
      Top             =   630
      Width           =   11895
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4815
      TabIndex        =   38
      Top             =   135
      Width           =   135
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo Doc :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   225
      TabIndex        =   37
      Top             =   180
      Width           =   1455
   End
End
Attribute VB_Name = "FrmCopiaPedido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsdeta As New ADODB.Recordset
Dim adll As New dllgeneral.dll_general
Dim Tabla As String, Tabladetalle As String

Private Sub aBusca_Change(Index As Integer)
  If Len(Trim(aBusca(Index))) = 0 Then
     If Index = 0 Then
        aBusca(1) = ""
     End If
     Call Limpiartexto(MBox2, 6, 10)
     Call Limpiartexto(Label2, 0, 8)
     Call CargaGrilla
  End If
  
End Sub

Private Sub aBusca_GotFocus(Index As Integer)
     Call Limpiartexto(MBox2, 6, 10)
     Call Limpiartexto(Label2, 0, 8)
     Call CargaGrilla
  
End Sub

Private Sub aBusca_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  Dim nsql As String
  
  If KeyCode = 112 Then  ' Ayuda de Productos
       If adll.ComboDato(Combo2.Text) = g_tipobol Then
            nsql = "CASE pedidocondicionfactura WHEN '1' THEN '*' ELSE '' END,pedidofecha as Fecha,pedidonrofact as Boleta,clientecodigo as Codigo,clienterazonsocial as Cliente,pedidototneto as Total"
       ElseIf adll.ComboDato(Combo2.Text) = g_tipofac Then
            nsql = "CASE pedidocondicionfactura WHEN '1' THEN '*' ELSE '' END,pedidofecha as Fecha,pedidonrofact as Factura,clientecodigo as Codigo,clienterazonsocial as Cliente,pedidototneto as Total"
       ElseIf adll.ComboDato(Combo2.Text) = g_tipoped Then
            nsql = "CASE pedidoestado WHEN '1' THEN '*' ELSE '' END,pedidofecha as Fecha,pedidonumero as Pedido,clientecodigo as Codigo,clienterazonsocial as Cliente,pedidototneto as Total"
       End If
       Dim sfiltra(1 To 4, 1 To 2) As String
       sfiltra(1, 1) = "Cliente": sfiltra(1, 2) = "clienterazonsocial"
       sfiltra(2, 1) = "Ruc": sfiltra(2, 2) = "clienteruc"
       sfiltra(3, 1) = "Doc.Venta": sfiltra(3, 2) = "pedidonrofact"
       sfiltra(4, 1) = "Pedido": sfiltra(4, 2) = "pedidonumero"
       FrmAyudaPedidos.TipoForma = 2
       FrmAyudaPedidos.BConexion = VGCNx
       FrmAyudaPedidos.BTabla = "vt_pedido"
       FrmAyudaPedidos.BCampos = nsql
       If adll.ComboDato(Combo2.Text) = g_tipobol Then
            FrmAyudaPedidos.BCondi = "pedidotipofac='" & g_tipobol & "'"
            FrmAyudaPedidos.BOrden = "pedidonrofact"
       ElseIf adll.ComboDato(Combo2.Text) = g_tipofac Then
            FrmAyudaPedidos.BCondi = "pedidotipofac='" & g_tipofac & "'"
            FrmAyudaPedidos.BOrden = "pedidonrofact"
       ElseIf adll.ComboDato(Combo2.Text) = g_tipoped Then
            FrmAyudaPedidos.BCondi = ""
            FrmAyudaPedidos.BOrden = "pedidonumero"
       End If
       'FrmAyudaPedidos.BCondi = ""
       FrmAyudaPedidos.BFiltro = sfiltra
       FrmAyudaPedidos.Show 1
       aBusca(0) = Left(nAyuda, aBusca(0).MaxLength)
       aBusca(1) = Right(nAyuda, aBusca(1).MaxLength)
       nAyuda = "": nDetalle = ""
   ElseIf KeyCode = 13 Then
       SendKeys "{tab}"
   End If
End Sub

Private Sub aBusca_LostFocus(Index As Integer)
    If Index = 0 Then
       aBusca(Index) = Right("000000000000" & aBusca(Index), aBusca(Index).MaxLength)
    ElseIf Index = 1 Then
       aBusca(Index) = Right("0000000000000" & aBusca(Index), aBusca(Index).MaxLength)
    End If
    
End Sub

Private Sub cBusca_Click()
Dim csql As New ADODB.Recordset
Dim acliente As New ADODB.Recordset
Dim nvalor As String
Dim nsql As String
Dim J As Integer
    
Call Limpiartexto(MBox2, 6, 10)
Call Limpiartexto(Label2, 0, 8)
Call CargaGrilla

Tabladetalle = IIf(Left(Combo2.Text, 2) = "PE", "vt_detallepedido", "detallecotizalibre")

If adll.ComboDato(Combo2.Text) = g_tipobol Then
   nsql = "select * from " & Tabla & " where pedidonrofact='" & Trim(aBusca(0) & aBusca(1)) & "' and pedidotipofac='" & g_tipobol & "' and puntovtacodigo='" & VGParametros.puntovta & "' and empresacodigo='" & VGParametros.empresacodigo & "' and pedidocondicionfactura<>'2'"
ElseIf adll.ComboDato(Combo2.Text) = g_tipofac Then
    nsql = "select * from " & Tabla & " where pedidonrofact='" & Trim(aBusca(0) & aBusca(1)) & "' and pedidotipofac='" & g_tipofac & "' and puntovtacodigo='" & VGParametros.puntovta & "' and empresacodigo='" & VGParametros.empresacodigo & "' and pedidocondicionfactura<>'2'"
ElseIf adll.ComboDato(Combo2.Text) = g_tipoped Then
    nsql = "select * from " & Tabla & " where pedidonumero='" & Trim(aBusca(0) & aBusca(1)) & "' and puntovtacodigo='" & VGParametros.puntovta & "' and empresacodigo='" & VGParametros.empresacodigo & "' and pedidocondicionfactura<>'2'"
ElseIf adll.ComboDato(Combo2.Text) = g_tipocot Then
    nsql = "select * from " & Tabla & " where pedidonumero='" & Trim(aBusca(0) & aBusca(1)) & "'  and puntovtacodigo='" & VGParametros.puntovta & "' and empresacodigo='" & VGParametros.empresacodigo & "' and pedidocondicionfactura<>'2'"
Else
    Exit Sub
End If

nvalor = ""
    
    Set csql = VGCNx.Execute(nsql)
    If csql.RecordCount > 0 Then
        nvalor = Escadena(csql!pedidonumero)
        If adll.ComboDato(Combo2.Text) = g_tipobol Then
            Label2(0) = Format(csql!pedidofechafact, "dd/mm/yyyy")
            Label2(1) = g_tipobol & "-" & Escadena(csql!pedidonrofact)
        ElseIf adll.ComboDato(Combo2.Text) = g_tipofac Then
            Label2(0) = Format(csql!pedidofechafact, "dd/mm/yyyy")
            Label2(1) = g_tipofac & "-" & Escadena(csql!pedidonrofact)
        ElseIf adll.ComboDato(Combo2.Text) = g_tipoped Then
            Label2(0) = Format(csql!pedidofecha, "dd/mm/yyyy")
            Label2(1) = g_tipoped & "-" & Escadena(csql!pedidonumero)
        ElseIf adll.ComboDato(Combo2.Text) = g_tipocot Then
            Label2(0) = Format(csql!pedidofecha, "dd/mm/yyyy")
            Label2(1) = g_tipocot & "-" & Escadena(csql!pedidonumero)
        End If
        
        Label2(2) = Escadena(csql!pedidonumero)
        Label2(3) = Escadena(csql!clientecodigo) & "-" & Escadena(csql!clienterazonsocial)
        Label2(4) = Escadena(csql!vendedorcodigo)
        Set acliente = VGCNx.Execute("select * from vt_vendedor where vendedorcodigo='" & Escadena(csql!vendedorcodigo) & "'")
        If acliente.RecordCount > 0 Then
           Label2(4) = Label2(4) & "-" & Escadena(acliente!vendedornombres)
        Else
            Label2(4) = Label2(4)
        End If
        acliente.Close
        Set acliente = Nothing
        Label2(5) = Escadena(csql!almacencodigo)
        Set acliente = VGCNx.Execute("select * from tabalm where taalma='" & Escadena(csql!almacencodigo) & "'")
        If acliente.RecordCount > 0 Then
           Label2(5) = Label2(5) & "-" & Escadena(acliente!tadescri)
        Else
            Label2(5) = Label2(5)
        End If
        acliente.Close
        Set acliente = Nothing
        
        Label2(6) = Escadena(csql!pedidolistaprec)
        Label2(7) = Escadena(csql!pedidomoneda)
        Set acliente = VGCNx.Execute("select * from gr_moneda where monedacodigo='" & Escadena(csql!pedidomoneda) & "'")
        If acliente.RecordCount > 0 Then
           Label2(7) = Label2(7) & "-" & Escadena(acliente!monedadescripcion)
        Else
           Label2(7) = Label2(7)
        End If
        acliente.Close
        Set acliente = Nothing
        Label2(8) = numero(csql!pedidotipcambio)
        MBox2(6) = numero(csql!pedidototitem)
        MBox2(7) = Format(csql!pedidototbruto, "##,###,##0.0000")
        MBox2(8) = numero(csql!pedidomontodsctoglobal + csql!pedidomontodsctocliente + csql!pedidomontodsctoppago + csql!pedidomontodsctovtaoficina + csql!pedidototaldsctoxitem + csql!pedidototaldsctoxlinea + csql!pedidototaldsctoxprom)
        MBox2(9) = numero(csql!pedidototimpuesto)
        MBox2(10) = numero(csql!pedidototneto)
        
    Else
        MsgBox "No existe Informacion del Documento...Verifique!!", vbInformation, MsgTitle
        csql.Close
        Set csql = Nothing
        Exit Sub
    End If
    csql.Close
       
    Set csql = VGCNx.Execute("select detpeditem,A.productocodigo,b.adescri,a.unidadcodigo," & _
                          "detpedcantpedida,detpedmontoprecvta,detpeddsctoxitem,detpedimpbruto," & _
                          " detpedporccomis,detpedcantpedidaref,detpedfactorconv " & _
                          "from " & Tabladetalle & " A " & _
                          "inner Join " & _
                          "[" & VGCNx.DefaultDatabase & "].dbo.maeart B " & _
                          " ON A.productocodigo=b.acodigo COLLATE Modern_Spanish_CI_AI " & _
                          "where pedidonumero='" & nvalor & "'")
    
    Set rsdeta = Nothing
    Call CargaGrilla
   
    Do Until csql.EOF
       rsdeta.AddNew
       rsdeta.Fields(0) = Escadena(csql!detpeditem)
       rsdeta.Fields(1) = Escadena(csql!productocodigo)
       rsdeta.Fields(2) = Escadena(csql!adescri)
       rsdeta.Fields(3) = Escadena(csql!unidadcodigo)
       rsdeta.Fields(4) = numero(csql!detpedcantpedida)
       rsdeta.Fields(5) = numero(IIf(IsNull(csql!detpedmontoprecvta), 0, csql!detpedmontoprecvta)) * numero(csql!detpedcantpedida) 'numero(csql!detpedimpbruto)   'bruto
       rsdeta.Fields(6) = numero(csql!detpeddsctoxitem)
       rsdeta.Fields(7) = numero(IIf(IsNull(csql!detpedmontoprecvta), 0, csql!detpedmontoprecvta))  'numero(csql!detpedimpbruto)
       rsdeta.Fields(8) = numero(csql!detpedporccomis)
       rsdeta.Fields(9) = numero(IIf(IsNull(csql!detpedcantpedidaref), 0, csql!detpedcantpedidaref))
       rsdeta.Fields(10) = numero(IIf(IsNull(csql!detpedfactorconv), 0, csql!detpedfactorconv))
       rsdeta.Update
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
   Call rsdeta.Fields.Append("Precio_Vta", adDouble)
   Call rsdeta.Fields.Append("Dscto(%)", adDouble)
   Call rsdeta.Fields.Append("Total", adDouble)
   Call rsdeta.Fields.Append("%", adDouble)
   Call rsdeta.Fields.Append("CantRef", adDouble)
   Call rsdeta.Fields.Append("Factor", adDouble)
   Call rsdeta.Fields.Append("%P", adDouble)
   
   rsdeta.Open
   ConfigGrid

End Function

Public Function ConfigGrid()
   Set TDBGrid1.DataSource = Nothing
   
   Set TDBGrid1.DataSource = rsdeta
   With TDBGrid1
      .Columns(0).Width = 600
      .Columns(0).Caption = "Item"
      .Columns(1).Width = 1100
      .Columns(1).Caption = "Codigo"
      .Columns(2).Width = 3000
      .Columns(2).Caption = "Descripcion"
      .Columns(3).Width = 600
      .Columns(3).Caption = "UM"
      .Columns(4).Width = 1000
      .Columns(4).Caption = "Cant"
      .Columns(5).Width = 1000
      .Columns(5).Caption = "Precio_Vta"
      .Columns(6).Width = 1000
      .Columns(6).Caption = "Dscto(%)"
      .Columns(7).Width = 800
      .Columns(7).Caption = "Total"
      .Columns(8).Width = 1000
      .Columns(8).Caption = "%"
      .Columns(5).NumberFormat = "###,##0.0000"
      .Columns(6).NumberFormat = "###,##0.00"
      .Columns(7).NumberFormat = "###,##0.0000"
      .Columns(8).NumberFormat = "###,##0.00"
      .Columns(9).Width = 800
      .Columns(9).Caption = "Cant.Ref"
      .Columns(9).NumberFormat = "###,##0"
      .Columns(10).Width = 600
      .Columns(10).Caption = "Factor"
      .Columns(10).NumberFormat = "###,##0.00"
      .Columns(11).Width = 0
      .Columns(11).NumberFormat = "###,##0.00"
   End With
   TDBGrid1.Refresh

End Function


Private Sub cCopia_Click()
    Dim nume As String
    Dim nsql As String
    Dim J As Double
    Dim nrs As New ADODB.Recordset
    Dim nrb As New ADODB.Recordset
    
    On Error GoTo nerror
    
    If (Len(Trim(Label2(2))) = 0 Or Len(Trim(Label2(3))) = 0) And TDBGrid1.ApproxCount = 0 Then Exit Sub
        
    If MsgBox("Desea Copiar el Documento?", vbYesNo + vbQuestion, MsgTitle) = vbYes Then

        If adll.VerificaDatoExistente(VGCNx, "select * from sysobjects where name Like 'jtempo%'") = 0 Then
            MsgBox "No existe la Tabla Temporal jtempo...Verifique!!!", vbInformation, MsgTitle
            Exit Sub
            'VGcnx.Execute "Select * into xtempo from vt_pedido where pedidonumero='" & Trim(aBusca(0) & aBusca(1)) & "'"
        Else
           If ExisteElem(0, VGCNx, "jtempo") Then VGCNx.Execute "drop table jtempo"
           VGCNx.Execute "select * into jtempo from " & Tabla & " where pedidonumero='" & Trim(aBusca(0) & aBusca(1)) & "'"
        End If
        'VGcnx.Execute "delete from xtempo"
        
        If adll.VerificaDatoExistente(VGCNx, "select * from sysobjects where name Like 'jdetatempo%'") = 0 Then
            MsgBox "No existe la Tabla Temporal jdetatempo...Verifique!!!", vbInformation, MsgTitle
            Exit Sub
            
        Else
           If ExisteElem(0, VGCNx, "jdetatempo") Then VGCNx.Execute "drop table jdetatempo"
            VGCNx.Execute "Select * into jdetatempo from " & Tabladetalle & " where pedidonumero='" & Trim(aBusca(0) & aBusca(1)) & "'"
        End If
           
        nume = g_pedserie & Right("000000000000" & TraeDataSerie("select puntovtadoccorr from vt_puntovtadocumento where documentocodigo='" & g_tipoped & "' and puntovtadocserie='" & g_pedserie & "'", VGCNx), 8)
        
        nsql = "Update jtempo Set pedidonumero='" & nume & "',"
        nsql = nsql & "pedidofecha='" & Date & "',pedidoobserva='' "
        nsql = nsql & " Where pedidonumero='" & Trim(aBusca(0) & aBusca(1)) & "'"
        
        VGCNx.Execute nsql

        
        VGCNx.Execute "Update jdetatempo " & _
                   " Set pedidonumero='" & nume & "'" & _
                   " Where pedidonumero='" & Trim(aBusca(0) & aBusca(1)) & "'"
        
        nsql = "Update vt_puntovtadocumento " & _
                " set puntovtadoccorr='" & Right("00000000" & Trim(Str(CDbl(nume) + 1)), 8) & "'" & _
                " Where documentocodigo='" & g_tipoped & "' and puntovtacodigo='" & g_ptoventa & "' and puntovtadocserie='" & g_pedserie & "' and empresacodigo='" & VGParametros.empresacodigo & "'"
                
        VGCNx.Execute nsql
        
      VGCNx.BeginTrans
        
        If VGParamSistem.stockcomp = 1 Then
           Dim rssql As New ADODB.Recordset
           Dim acmd As New ADODB.Command
           Dim asql As String
           Set rssql = VGCNx.Execute("select almacencodigo from jtempo")
           asql = Trim(rssql.Fields(0))
           Set rssql = VGCNx.Execute("select productocodigo,detpedcantpedida from jdetatempo")
            If rssql.RecordCount > 0 Then
               rssql.MoveFirst
               Do Until rssql.EOF
                     Set acmd.ActiveConnection = VGgeneral
                     acmd.CommandType = adCmdStoredProc
                     acmd.CommandTimeout = 0
                     acmd.CommandText = "vt_actualizoalma_pro"
                     acmd.Prepared = True
                     With acmd
                         .Parameters("@basedes") = CStr(VGCNx.DefaultDatabase)
                         .Parameters("@almacen") = asql
                         .Parameters("@tipo") = "3"
                         .Parameters("@articulo") = Trim(rssql.Fields(0))
                         .Parameters("@cantidad") = rssql.Fields(1)
                    End With
                    acmd.Execute
                    Set acmd = Nothing
                    rssql.MoveNext
              Loop
           End If
        End If
        rssql.Close
        
        VGCNx.Execute "insert into vt_tempopedido" & g_ptoventa & "  Select * from jtempo"
        
        Set nrb = VGCNx.Execute("select * from jdetatempo")
        If nrb.RecordCount > 0 Then
            nrs.Open "vt_tempodetallepedido" & g_ptoventa, VGCNx, adOpenDynamic, adLockOptimistic
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
                
            VGCNx.Execute "update " & Tabla & " set pedidocondicionfactura='2' " _
            & " where pedidonumero='" & aBusca(0) & aBusca(1) & "' and " _
            & " puntovtacodigo='" & VGParametros.puntovta & "' and " _
            & " empresacodigo='" & VGParametros.empresacodigo & "'"
            
            MsgBox "Numero de Pedido => " & nume, vbInformation, MsgTitle
        
        End If
        nrb.Close
        
        Set nrb = Nothing
        
      VGCNx.CommitTrans
      VGCNx.Execute "delete from jdetatempo"
      VGCNx.Execute "delete from jtempo"
      'VGcnx.Execute "Drop Table jTempo"
      'VGcnx.Execute "Drop Table jdetatempo"
      
    End If
 
nerror:
 If Err Then
    MsgBox "Comunicarse con  el Sistema" & Chr(13) & Chr(10) & Err.Number & "-" & Err.Description, vbInformation, MsgTitle
    'Resume
    Err = 0
    VGCNx.RollbackTrans
  
    Exit Sub
 End If
    
End Sub

Private Sub Combo2_Click()
aBusca(0) = ""
aBusca(1) = ""
Call Limpiartexto(MBox2, 6, 10)
Call Limpiartexto(Label2, 0, 8)
Call CargaGrilla

Tabla = IIf(Left(Combo2.Text, 2) = "PE", "vt_pedido", "cotizalibre")


End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
Seguir Combo2, KeyAscii
End Sub

Private Sub Form_Load()
MostrarFormVentas Me, "C"

Call Limpiartexto(MBox2, 6, 10)
Combo2.Clear
Combo2.AddItem g_tipoped & "-Pedidos"
Combo2.AddItem g_tipocot & "-Cotizacion"
Combo2.ListIndex = 0
Call CargaGrilla
    
Tabla = "vt_pedido"

cBusca.Picture = MDIPrincipal.ImageList3.ListImages.Item("Buscar").Picture
cCopia.Picture = MDIPrincipal.ImageList3.ListImages.Item("Copiar").Picture

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  Set rsdeta = Nothing
End Sub



