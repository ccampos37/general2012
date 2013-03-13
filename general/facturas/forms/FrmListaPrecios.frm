VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form FrmListaPrecios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lista Precios"
   ClientHeight    =   7770
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11265
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7770
   ScaleWidth      =   11265
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   7650
      Left            =   90
      TabIndex        =   0
      Top             =   45
      Width           =   11115
      _ExtentX        =   19606
      _ExtentY        =   13494
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Consulta"
      TabPicture(0)   =   "FrmListaPrecios.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "oCrystalReport"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Fr4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdBotones(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdBotones(1)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdBotones(2)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdBotones(4)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdBotones(3)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "Mantenimiento"
      TabPicture(1)   =   "FrmListaPrecios.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label7"
      Tab(1).Control(1)=   "Label5"
      Tab(1).Control(2)=   "TDBGrid2"
      Tab(1).Control(3)=   "Frame5(0)"
      Tab(1).Control(4)=   "Frame3"
      Tab(1).Control(5)=   "aBusca(1)"
      Tab(1).Control(6)=   "aBusca(0)"
      Tab(1).Control(7)=   "cBusca"
      Tab(1).Control(8)=   "cmdBotones(12)"
      Tab(1).Control(9)=   "cmdBotones(11)"
      Tab(1).ControlCount=   10
      Begin VB.CommandButton cmdBotones 
         Caption         =   "&Imprimir"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Index           =   3
         Left            =   6075
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   6570
         Width           =   1050
      End
      Begin VB.CommandButton cmdBotones 
         Caption         =   "&Salir"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Index           =   4
         Left            =   7200
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   6570
         Width           =   1050
      End
      Begin VB.CommandButton cmdBotones 
         Caption         =   "&Eliminar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Index           =   2
         Left            =   4950
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   6570
         Width           =   1050
      End
      Begin VB.CommandButton cmdBotones 
         Caption         =   "E&ditar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Index           =   1
         Left            =   3825
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   6570
         Width           =   1050
      End
      Begin VB.CommandButton cmdBotones 
         Caption         =   "&Nuevo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Index           =   0
         Left            =   2700
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   6570
         Width           =   1050
      End
      Begin VB.Frame Fr4 
         BackColor       =   &H00C9955A&
         BorderStyle     =   0  'None
         Height          =   4185
         Left            =   3375
         TabIndex        =   20
         Top             =   1080
         Visible         =   0   'False
         Width           =   4815
         Begin VB.CommandButton cBoton 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Acepta"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   0
            Left            =   1020
            Style           =   1  'Graphical
            TabIndex        =   32
            Top             =   3420
            Width           =   1395
         End
         Begin VB.CommandButton cBoton 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Cancela"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   1
            Left            =   2580
            Style           =   1  'Graphical
            TabIndex        =   31
            Top             =   3420
            Width           =   1395
         End
         Begin VB.Frame Fr2 
            BackColor       =   &H00C9955A&
            Height          =   2475
            Index           =   1
            Left            =   390
            TabIndex        =   21
            Top             =   600
            Width           =   4245
            Begin VB.OptionButton Opt 
               BackColor       =   &H00C9955A&
               Caption         =   "Crea a Partir de Otra Tabla"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   0
               Left            =   330
               TabIndex        =   25
               Top             =   330
               Width           =   2325
            End
            Begin VB.OptionButton Opt 
               BackColor       =   &H00C9955A&
               Caption         =   "Crea a Partir de Maestro de Productos"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   1
               Left            =   330
               TabIndex        =   24
               Top             =   720
               Width           =   3045
            End
            Begin VB.ComboBox Combo1 
               Height          =   315
               Left            =   2730
               Style           =   2  'Dropdown List
               TabIndex        =   23
               Top             =   270
               Width           =   1125
            End
            Begin MSMask.MaskEdBox MBox3 
               Height          =   255
               Index           =   0
               Left            =   2910
               TabIndex        =   22
               Top             =   1260
               Width           =   1035
               _ExtentX        =   1826
               _ExtentY        =   450
               _Version        =   393216
               Appearance      =   0
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MBox3 
               Height          =   255
               Index           =   1
               Left            =   2910
               TabIndex        =   26
               Top             =   1620
               Width           =   1035
               _ExtentX        =   1826
               _ExtentY        =   450
               _Version        =   393216
               Appearance      =   0
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MBox3 
               Height          =   255
               Index           =   2
               Left            =   2910
               TabIndex        =   27
               Top             =   1980
               Width           =   1035
               _ExtentX        =   1826
               _ExtentY        =   450
               _Version        =   393216
               Appearance      =   0
               PromptChar      =   "_"
            End
            Begin VB.Line Line1 
               BorderColor     =   &H80000005&
               Index           =   2
               X1              =   30
               X2              =   4230
               Y1              =   1080
               Y2              =   1080
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackColor       =   &H00C9955A&
               Caption         =   "Factor Nuevo Precio :"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   0
               Left            =   210
               TabIndex        =   30
               Top             =   1290
               Width           =   1560
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackColor       =   &H00C9955A&
               Caption         =   "Factor para Dscto. Vta Oficina :"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   1
               Left            =   210
               TabIndex        =   29
               Top             =   1620
               Width           =   2235
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackColor       =   &H00C9955A&
               Caption         =   "Factor para Dscto. Vta Reparto :"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   2
               Left            =   210
               TabIndex        =   28
               Top             =   1980
               Width           =   2310
            End
         End
         Begin MSMask.MaskEdBox Lpre 
            Height          =   315
            Left            =   3000
            TabIndex        =   33
            Top             =   270
            Width           =   450
            _ExtentX        =   794
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            AllowPrompt     =   -1  'True
            MaxLength       =   1
            PromptChar      =   "_"
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H80000009&
            BorderWidth     =   3
            Height          =   4110
            Left            =   45
            Top             =   45
            Width           =   4740
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Nueva Lista de Precios :"
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
            Height          =   165
            Index           =   0
            Left            =   750
            TabIndex        =   34
            Top             =   300
            Width           =   2025
         End
      End
      Begin VB.CommandButton cmdBotones 
         Caption         =   "&Acepta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1005
         Index           =   11
         Left            =   -66765
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   630
         Width           =   1140
      End
      Begin VB.CommandButton cmdBotones 
         Caption         =   "&Cancelar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1005
         Index           =   12
         Left            =   -65460
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   630
         Width           =   1140
      End
      Begin VB.CommandButton cBusca 
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
         Height          =   1005
         Left            =   -68025
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   630
         Width           =   1140
      End
      Begin VB.TextBox aBusca 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   -73395
         MaxLength       =   8
         TabIndex        =   11
         Top             =   855
         Width           =   975
      End
      Begin VB.TextBox aBusca 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   -73425
         TabIndex        =   10
         Top             =   1215
         Width           =   4590
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Height          =   675
         Left            =   -74775
         TabIndex        =   6
         Top             =   6210
         Width           =   3075
         Begin VB.TextBox TPrecio 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1440
            MaxLength       =   10
            TabIndex        =   7
            Top             =   240
            Width           =   1485
         End
         Begin VB.Label Label4 
            Caption         =   "Precio Venta"
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
            Left            =   180
            TabIndex        =   8
            Top             =   240
            Width           =   1185
         End
      End
      Begin VB.Frame Frame5 
         BorderStyle     =   0  'None
         Height          =   585
         Index           =   0
         Left            =   -66540
         TabIndex        =   3
         Top             =   6210
         Width           =   2265
         Begin VB.TextBox TReg 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Index           =   1
            Left            =   1170
            TabIndex        =   4
            Top             =   210
            Width           =   975
         End
         Begin VB.Label Label3 
            Caption         =   "Total Reg."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   0
            Left            =   150
            TabIndex        =   5
            Top             =   240
            Width           =   1035
         End
      End
      Begin Crystal.CrystalReport oCrystalReport 
         Left            =   480
         Top             =   6840
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.Frame Frame2 
         Height          =   5925
         Left            =   180
         TabIndex        =   1
         Top             =   540
         Width           =   10755
         Begin VB.TextBox TReg 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Index           =   0
            Left            =   9645
            TabIndex        =   40
            Top             =   5535
            Width           =   975
         End
         Begin VB.Frame Frame6 
            BackColor       =   &H00C9955A&
            BorderStyle     =   0  'None
            Height          =   1545
            Left            =   3420
            TabIndex        =   17
            Top             =   2115
            Visible         =   0   'False
            Width           =   4245
            Begin VB.CommandButton cBoton 
               BackColor       =   &H00FFFFFF&
               Caption         =   "&Cancela"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Index           =   2
               Left            =   2130
               Style           =   1  'Graphical
               TabIndex        =   19
               Top             =   840
               Width           =   1395
            End
            Begin VB.CommandButton cBoton 
               BackColor       =   &H00FFFFFF&
               Caption         =   "&Acepta"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Index           =   3
               Left            =   570
               Style           =   1  'Graphical
               TabIndex        =   18
               Top             =   840
               Width           =   1395
            End
            Begin VB.Shape Shape2 
               BorderColor     =   &H80000009&
               BorderWidth     =   3
               Height          =   1455
               Left            =   45
               Top             =   45
               Width           =   4155
            End
         End
         Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
            Height          =   5040
            Left            =   210
            TabIndex        =   2
            Top             =   270
            Width           =   10395
            _ExtentX        =   18336
            _ExtentY        =   8890
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
            _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=240,.bold=0,.fontsize=825,.italic=0"
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
            Height          =   225
            Index           =   1
            Left            =   8505
            TabIndex        =   41
            Top             =   5565
            Width           =   1035
         End
      End
      Begin TrueOleDBGrid70.TDBGrid TDBGrid2 
         Height          =   4380
         Left            =   -74730
         TabIndex        =   9
         Top             =   1800
         Width           =   10575
         _ExtentX        =   18653
         _ExtentY        =   7726
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
         AllowAddNew     =   -1  'True
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
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=255,.bold=0,.fontsize=825,.italic=0"
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
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Descripcion :"
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
         Left            =   -74595
         TabIndex        =   16
         Top             =   1260
         Width           =   1065
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Codigo :"
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
         Left            =   -74595
         TabIndex        =   13
         Top             =   870
         Width           =   660
      End
   End
End
Attribute VB_Name = "FrmListaPrecios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim adll As New dllgeneral.dll_general
Dim rsb As New ADODB.Recordset
Dim archilista As String
Public PTalmacen As String


Private Sub aBusca_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = 13 Then
      SendKeys "{tab}"
  End If
End Sub

Private Sub cBoton_Click(Index As Integer)
  Dim asql As String
  On Error Resume Next
  If Index = 0 Then
    If adll.VerificaDatoExistente(VGCNx, "select * from sysobjects where name='" & "listapre" & Trim(Lpre.text) & "'") = 1 Then
       MsgBox "Ya existe la Lista de Precios No.: " & Trim(Lpre), vbInformation, MsgTitle
       Call adll.Enfoquetexto(Lpre)
       Exit Sub
    End If
    If Not adll.ValidaCadena(MBox3(0), "N") Then
       MsgBox "Factores Incompletos...", vbInformation, MsgTitle
       Call adll.Enfoquetexto(MBox3(0))
       Exit Sub
    End If
    If Not adll.ValidaCadena(MBox3(1), "N") Then
       MsgBox "Factores Incompletos...", vbInformation, MsgTitle
       Call adll.Enfoquetexto(MBox3(1))
       Exit Sub
    End If
    If Not adll.ValidaCadena(MBox3(2), "N") Then
       MsgBox "Factores Incompletos...", vbInformation, MsgTitle
       Call adll.Enfoquetexto(MBox3(2))
       Exit Sub
    End If
    If Combo1.ListCount <= 0 And Opt(0).Value Then
       MsgBox "Debe Ingresar la lista ...", vbInformation, MsgTitle
       Combo1.SetFocus
       Exit Sub
    End If
    Fr4.Visible = True
    VGCNx.Execute "CREATE TABLE Listapre" & Trim(Lpre.text) & _
               "(productocodigo char(20),productodescripcion char(100),productoprecvta float," & _
               "productodescrcorta char(30), grupovtacodigo char(2),productofamiliacodigo varchar(5)," & _
               "productocategoriacodigo char(3),productotipo char(1),productoporcimpto float," & _
               "productoestunidreferencia bit,unidadfactorconv float," & _
               "productoprecvtaofi float,productoprecvtareparto float,unidadcodigo varchar(5)," & _
               "unidadreferencial varchar(5),monedacodigo char(2),empresacodigo char(2)," & _
               "CONSTRAINT ID_PK" & Trim(Lpre.text) & " PRIMARY KEY (productocodigo) ) "
               
    DoEvents
    
    If Opt(0).Value Then
      If Combo1.ListCount > 0 Then
       FrmProcesolista.Btipo = 1
       FrmProcesolista.Bmodo = 0
       FrmProcesolista.bfactorprecio = CDbl(MBox3(0))
       FrmProcesolista.bfactordvta = CDbl(MBox3(1))
       FrmProcesolista.bfactordreparto = CDbl(MBox3(2))
       FrmProcesolista.bsqlini = "Select * From Listapre" & Trim(adll.ComboDato(Combo1.text))
      Else
        Combo1.SetFocus
        Exit Sub
      End If
    Else
       FrmProcesolista.Btipo = 1
       FrmProcesolista.Bmodo = 1
       FrmProcesolista.bfactorprecio = CDbl(MBox3(0))
       FrmProcesolista.bfactordvta = CDbl(MBox3(1))
       FrmProcesolista.bfactordreparto = CDbl(MBox3(2))
       FrmProcesolista.bsqlini = "select * From maeart where isnull(afstock,1)=1 "
    End If
    FrmProcesolista.bsqlfin = "Listapre" & Trim(Lpre.text)
    FrmProcesolista.Show 1
    'Unload FrmProcesolista
  ElseIf Index = 3 Then
     Frame6.Visible = False
     SQL = "select productocodigo,productodescripcion,productoprecvta,almacencodigo='  ' from " & archilista & ""
     
     Call Listado(SQL)
     Call adll.ActivaTab(1, 1, SSTab1)
     Configura
     TDBGrid2.SetFocus
     Exit Sub
  ElseIf Index = 2 Then
     Frame6.Visible = False
     Exit Sub
  End If
  Configura
  Fr4.Visible = False
  cmdBotones(0).SetFocus
End Sub

Private Sub cBusca_Click()
  Dim asql As String
  If Len(Trim(aBusca(0).text)) > 0 Then
     asql = "select productocodigo,productodescripcion,productoprecvta,almacencodigo from " & TDBGrid1.Columns(1).text & " Where productocodigo like '" & Trim(aBusca(0).text) & "%'"
     Call Listado(asql)
  ElseIf Len(Trim(aBusca(1))) > 0 Then
     asql = "select productocodigo,productodescripcion,productoprecvta,almacencodigo from " & TDBGrid1.Columns(1).text & " Where productodescripcion like '%" & Trim(aBusca(1).text) & "%'"
     Call Listado(asql)
  End If
  
End Sub

Private Sub cmdBotones_Click(Index As Integer)
   Dim asql As String
   Dim busca As New dll_apisgen.dll_apis
   Dim acmd As New ADODB.Command
  
   On Error GoTo nerror
   
   Select Case Index
    Case 0
       Fr4.Visible = True
       Lpre.text = 0
       MBox3(0) = "0.0000": MBox3(1) = "0.0000": MBox3(2) = "0.0000"
       Lpre.SetFocus
    Case 1
       If TDBGrid1.Row < 0 Then
          Exit Sub
       End If
       TPrecio = numero(0)
       Set rsb = Nothing
       archilista = TDBGrid1.Columns(1).text
       Frame6.Visible = True
    Case 2
        If TDBGrid1.Row >= 0 Then
            If MsgBox("Desea Eliminar la Lista de Precios?", vbYesNo, MsgTitle) = vbYes Then
               VGCNx.Execute "drop table " & TDBGrid1.Columns(1).text
               TReg(0) = Format(TReg(0) - 1, "#######0")
            End If
            Configura
        End If
    Case 3
       oCrystalReport.ReportFileName = VGParamSistem.RutaReport & "RepListaprecio.rpt"
       oCrystalReport.LogOnServer "pdssql.dll", _
               busca.LeerIni(App.Path & "\MARFICE.INI", "CONEXION", "SERVIDOR", ""), _
               busca.LeerIni(App.Path & "\MARFICE.INI", "CONEXION", "BDDATOS", ""), _
               busca.LeerIni(App.Path & "\MARFICE.INI", "CONEXION", "USUARIO", ""), _
               busca.LeerIni(App.Path & "\MARFICE.INI", "CONEXION", "PASSW", "")
       oCrystalReport.Connect = _
              "DSN=" & busca.LeerIni(App.Path & "\MARFICE.INI", "CONEXION", "SERVIDOR", "") & ";" & _
              "DSQ=" & busca.LeerIni(App.Path & "\MARFICE.INI", "CONEXION", "BDDATOS", "") & ";" & _
              "UID=" & busca.LeerIni(App.Path & "\MARFICE.INI", "CONEXION", "USUARIO", "") & ";" & _
              "PWD=" & busca.LeerIni(App.Path & "\MARFICE.INI", "CONEXION", "PASSW", "")
               
       oCrystalReport.Destination = crptToWindow
       oCrystalReport.WindowState = crptMaximized
       oCrystalReport.formulas(0) = "empresa='" & Trim(VGParametros.NomEmpresa) & "'"
       oCrystalReport.DiscardSavedData = True
       oCrystalReport.StoredProcParam(0) = CStr(VGCNx.DefaultDatabase)
       oCrystalReport.StoredProcParam(1) = CStr(TDBGrid1.Columns(1).text)
       oCrystalReport.Action = 1
    Case 4
       Unload Me
    Case 11
        If MsgBox("Desea Grabar los Cambios?", vbYesNo + vbQuestion, MsgTitle) = vbYes Then
           If TDBGrid2.Row >= 0 Then
               FrmProcesolista.Btipo = 2
               'FrmProcesolista.bsqlfin = TDBGrid1.Columns(1).Text
               FrmProcesolista.bsqlfin = archilista
               FrmProcesolista.Show
               Unload FrmProcesolista
           End If
        End If
        Configura
        Call adll.ActivaTab(0, 1, SSTab1)
        TDBGrid1.SetFocus
    Case 12
       Call adll.ActivaTab(0, 1, SSTab1)
   End Select
   
nerror:
   If Err <> 0 Then
     MsgBox "Error Inesperado : " & Err.Number & "-" & Err.Description, vbInformation, MsgTitle
     Err = 0
     Exit Sub
   End If
End Sub

Public Function Listado(asql)
    Set rsb = Nothing
    rsb.Open asql, VGCNx, adOpenDynamic, adLockOptimistic
    Set rsb.ActiveConnection = Nothing
    
    Set TDBGrid2.DataSource = rsb
    With TDBGrid2
       .Columns(0).HeadAlignment = dbgCenter
       .Columns(0).Caption = "Codigo"
       .Columns(0).Width = 1000
       .Columns(1).HeadAlignment = dbgCenter
       .Columns(1).Caption = "Descripcion"
       .Columns(1).Width = 5500
       .Columns(2).HeadAlignment = dbgCenter
       .Columns(2).Caption = "Precio Vta"
       .Columns(2).NumberFormat = "##,###,##0.0000"
    End With
    'Label5 = TDBGrid1.Columns(0).Text
    TReg(1) = Format(rsb.RecordCount, "#######0")
    TDBGrid2.Refresh
End Function




Private Sub Combo1_GotFocus()
   If Combo1.ListCount - 1 >= 0 Then
     Call adll.llenacombo(Combo1, "select right(name,1),'Precios' from sysobjects where name like 'listapre%'", VGCNx)
     Combo1.ListIndex = 0
   End If

End Sub

Private Sub Form_Load()
Me.Show
MostrarFormVentas Me, "C"

Configura
    
Call adll.ActivaTab(0, 1, SSTab1)

cmdBotones(0).Picture = MDIPrincipal.ImageList2.ListImages.item("Nuevo").Picture
cmdBotones(1).Picture = MDIPrincipal.ImageList2.ListImages.item("Modificar").Picture
cmdBotones(2).Picture = MDIPrincipal.ImageList2.ListImages.item("Eliminar").Picture
cmdBotones(3).Picture = MDIPrincipal.ImageList2.ListImages.item("Imprimir").Picture
cmdBotones(4).Picture = MDIPrincipal.ImageList2.ListImages.item("Retornar").Picture

'cBusca.Picture = MDIPrincipal.ImageList3.ListImages.item("Buscar").Picture
'cmdBotones(11).Picture = MDIPrincipal.ImageList3.ListImages.item("Copiar").Picture
'cmdBotones(12).Picture = MDIPrincipal.ImageList2.ListImages.item("Retornar").Picture

End Sub

Public Function Configura()
    Dim rs As New ADODB.Recordset
    Set rs = VGCNx.Execute("select 'Lista de Precios ' + right(name,1),name as Codigo  from sysobjects where name like 'listapre%'")
    Set TDBGrid1.DataSource = Nothing
    TDBGrid1.ClearFields
    TDBGrid1.Refresh
    Set TDBGrid1.DataSource = rs
    
    With TDBGrid1
        .Columns(0).HeadAlignment = dbgCenter
        .Columns(0).Caption = "Descripcion"
        .Columns(0).Width = 6500
    End With
    TReg(0) = Format(rs.RecordCount, "#######0")
    
    Call adll.llenacombo(Combo1, "select right(name,1),'Precios' from sysobjects where name like 'listapre%'", VGCNx)
    TDBGrid1.Refresh
    Set rs = Nothing
End Function

Private Sub Lpre_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then
    If adll.VerificaDatoExistente(VGCNx, "select * from sysobjects where name='" & "listapre" & Trim(Lpre.text) & "'") = 1 Then
       MsgBox "Ya existe la Lista de Precios No.: " & Trim(Lpre), vbInformation, MsgTitle
    Else
      Fr2(1).Enabled = True
      Opt(0).SetFocus
      Exit Sub
    End If
  End If
End Sub

Private Sub MBox3_GotFocus(Index As Integer)
   Call adll.Enfoquetexto(MBox3(Index))
End Sub

Private Sub MBox3_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = 13 Then
     MBox3(Index) = Format(MBox3(Index), "###0.0000")
     SendKeys "{tab}"
  End If
End Sub

Private Sub MBox3_LostFocus(Index As Integer)
   If adll.ValidaCadena(MBox3(Index), "N") Then
      MBox3(Index) = Format(MBox3(Index), "###0.0000")
      Exit Sub
   End If
End Sub

Private Sub Opt_Click(Index As Integer)
  If Index = 0 And Opt(0).Value Then
     Combo1.SetFocus
  End If
End Sub

Private Sub Opt_DblClick(Index As Integer)
   TPrecio = numero(TDBGrid2.Columns(2).text)
   Call adll.Enfoquetexto(TPrecio)
 
End Sub

Private Sub TDBGrid2_Click()
  TPrecio = numero(TDBGrid2.Columns(2).text)
   Call adll.Enfoquetexto(TPrecio)
End Sub

Private Sub TDBGrid2_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 13 Then
   TPrecio = numero(TDBGrid2.Columns(2).text)
   Call adll.Enfoquetexto(TPrecio)
   Exit Sub
 End If
 
 
End Sub

Private Sub TPrecio_KeyPress(KeyAscii As Integer)
  
   If KeyAscii = 13 Then
      If Not adll.ValidaCadena(TPrecio, "N") Then
        MsgBox Msg29, vbInformation, MsgTitle
        Call adll.Enfoquetexto(TPrecio)
        Exit Sub
      End If
      TPrecio = numero(TPrecio)
      
      rsb.Fields(2) = numero(TPrecio)
      rsb.Update
      TDBGrid2.SetFocus
      Exit Sub
   End If

End Sub
