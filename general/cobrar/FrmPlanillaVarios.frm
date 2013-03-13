VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Begin VB.Form FrmPlanillaVarios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Planilla de Documentos Varios"
   ClientHeight    =   8715
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11940
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8715
   ScaleWidth      =   11940
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   7845
      Left            =   150
      TabIndex        =   5
      Top             =   150
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   13838
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Datos Generales"
      TabPicture(0)   =   "FrmPlanillaVarios.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame4"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Detalle"
      TabPicture(1)   =   "FrmPlanillaVarios.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "frmbotones"
      Tab(1).Control(1)=   "Frame3"
      Tab(1).ControlCount=   2
      Begin VB.Frame Frame4 
         Height          =   930
         Left            =   5310
         TabIndex        =   37
         Top             =   5400
         Width           =   1980
         Begin VB.CommandButton cmdBotones 
            Caption         =   "&Acepta"
            Height          =   690
            Index           =   11
            Left            =   90
            Picture         =   "FrmPlanillaVarios.frx":0038
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   180
            Width           =   870
         End
         Begin VB.CommandButton cmdBotones 
            Caption         =   "&Cancelar"
            Height          =   690
            Index           =   12
            Left            =   1050
            Picture         =   "FrmPlanillaVarios.frx":047A
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   180
            Width           =   855
         End
      End
      Begin VB.Frame frmbotones 
         Height          =   930
         Left            =   -71040
         TabIndex        =   14
         Top             =   6660
         Width           =   4290
         Begin VB.CommandButton cmdBotones 
            Caption         =   "&Nuevo"
            Height          =   690
            Index           =   0
            Left            =   180
            Picture         =   "FrmPlanillaVarios.frx":08BC
            Style           =   1  'Graphical
            TabIndex        =   24
            Top             =   180
            Width           =   870
         End
         Begin VB.CommandButton cmdBotones 
            Caption         =   "&Grabar"
            Height          =   690
            Index           =   1
            Left            =   1260
            Picture         =   "FrmPlanillaVarios.frx":0CFE
            Style           =   1  'Graphical
            TabIndex        =   25
            Top             =   180
            Width           =   870
         End
         Begin VB.CommandButton cmdBotones 
            Caption         =   "&Eliminar"
            Height          =   690
            Index           =   2
            Left            =   2280
            Picture         =   "FrmPlanillaVarios.frx":1140
            Style           =   1  'Graphical
            TabIndex        =   26
            Top             =   180
            Width           =   825
         End
         Begin VB.CommandButton cmdBotones 
            Caption         =   "&Salir"
            Height          =   690
            Index           =   4
            Left            =   3255
            Picture         =   "FrmPlanillaVarios.frx":1582
            Style           =   1  'Graphical
            TabIndex        =   28
            Top             =   180
            Width           =   870
         End
      End
      Begin VB.Frame Frame3 
         Height          =   6045
         Left            =   -74790
         TabIndex        =   12
         Top             =   570
         Width           =   11175
         Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
            Height          =   4305
            Left            =   150
            TabIndex        =   38
            Top             =   330
            Width           =   10845
            _ExtentX        =   19129
            _ExtentY        =   7594
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
            PrintInfos(0).PageHeaderFont=   "Size=9.75,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=System"
            PrintInfos(0).PageFooterFont=   "Size=9.75,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=System"
            PrintInfos(0).PageHeaderHeight=   0
            PrintInfos(0).PageFooterHeight=   0
            PrintInfos.Count=   1
            Appearance      =   2
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
            _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=43,.bold=0,.fontsize=825,.italic=0"
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
         Begin VB.Frame Frame5 
            Height          =   1305
            Left            =   150
            TabIndex        =   13
            Top             =   4650
            Width           =   10875
            Begin VB.CommandButton cAyuda 
               Caption         =   "..."
               Height          =   315
               Index           =   1
               Left            =   10530
               TabIndex        =   42
               Top             =   480
               Width           =   195
            End
            Begin VB.CommandButton cAyuda 
               Caption         =   "..."
               Height          =   315
               Index           =   0
               Left            =   2130
               TabIndex        =   41
               Top             =   450
               Width           =   195
            End
            Begin VB.CommandButton cAyuda 
               Caption         =   "..."
               Height          =   315
               Index           =   2
               Left            =   1470
               TabIndex        =   40
               Top             =   450
               Width           =   150
            End
            Begin VB.TextBox Text1 
               Height          =   285
               Index           =   6
               Left            =   9930
               MaxLength       =   2
               TabIndex        =   23
               Top             =   480
               Width           =   585
            End
            Begin VB.TextBox Text1 
               Height          =   285
               Index           =   5
               Left            =   8430
               MaxLength       =   14
               TabIndex        =   22
               Top             =   480
               Width           =   1365
            End
            Begin VB.TextBox Text1 
               Height          =   285
               Index           =   4
               Left            =   7560
               MaxLength       =   2
               TabIndex        =   21
               Top             =   480
               Width           =   705
            End
            Begin MSMask.MaskEdBox MBox2 
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "dd/MM/yyyy"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   10250
                  SubFormatType   =   3
               EndProperty
               Height          =   315
               Index           =   0
               Left            =   4530
               TabIndex        =   19
               Top             =   450
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   556
               _Version        =   393216
               MaxLength       =   10
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/####"
               PromptChar      =   "_"
            End
            Begin VB.TextBox Text1 
               Height          =   285
               Index           =   3
               Left            =   3060
               MaxLength       =   10
               TabIndex        =   18
               Top             =   450
               Width           =   1275
            End
            Begin VB.TextBox Text1 
               Height          =   285
               Index           =   2
               Left            =   2340
               MaxLength       =   4
               TabIndex        =   17
               Top             =   450
               Width           =   525
            End
            Begin VB.TextBox Text1 
               Height          =   285
               Index           =   1
               Left            =   1650
               MaxLength       =   2
               TabIndex        =   16
               Top             =   450
               Width           =   465
            End
            Begin VB.TextBox Text1 
               Height          =   285
               Index           =   0
               Left            =   180
               MaxLength       =   11
               TabIndex        =   15
               Top             =   450
               Width           =   1275
            End
            Begin MSMask.MaskEdBox MBox2 
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "dd/MM/yyyy"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   10250
                  SubFormatType   =   3
               EndProperty
               Height          =   315
               Index           =   1
               Left            =   6000
               TabIndex        =   20
               Top             =   450
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   556
               _Version        =   393216
               MaxLength       =   10
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/####"
               PromptChar      =   "_"
            End
            Begin VB.Label Label3 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   180
               TabIndex        =   39
               Top             =   840
               Width           =   10395
            End
            Begin VB.Label Label2 
               Caption         =   "Banco"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   8
               Left            =   9960
               TabIndex        =   36
               Top             =   240
               Width           =   645
            End
            Begin VB.Label Label2 
               Alignment       =   2  'Center
               Caption         =   "Importe"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   7
               Left            =   8610
               TabIndex        =   35
               Top             =   240
               Width           =   1035
            End
            Begin VB.Label Label2 
               Alignment       =   2  'Center
               Caption         =   "Mon."
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   6
               Left            =   7650
               TabIndex        =   34
               Top             =   210
               Width           =   585
            End
            Begin VB.Label Label2 
               Alignment       =   2  'Center
               Caption         =   "F. Vcmto."
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   5
               Left            =   6060
               TabIndex        =   33
               Top             =   180
               Width           =   1095
            End
            Begin VB.Label Label2 
               Alignment       =   2  'Center
               Caption         =   "F. Emision"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   4
               Left            =   4590
               TabIndex        =   32
               Top             =   180
               Width           =   1095
            End
            Begin VB.Label Label2 
               Alignment       =   2  'Center
               Caption         =   "Numero"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   3
               Left            =   3090
               TabIndex        =   31
               Top             =   180
               Width           =   1155
            End
            Begin VB.Label Label2 
               Alignment       =   2  'Center
               Caption         =   "Serie"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   2
               Left            =   2340
               TabIndex        =   30
               Top             =   180
               Width           =   615
            End
            Begin VB.Label Label2 
               Caption         =   "TD"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   1
               Left            =   1740
               TabIndex        =   29
               Top             =   180
               Width           =   315
            End
            Begin VB.Label Label2 
               Alignment       =   2  'Center
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
               Height          =   180
               Index           =   0
               Left            =   300
               TabIndex        =   27
               Top             =   180
               Width           =   1095
            End
         End
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H80000008&
         Height          =   3015
         Left            =   2190
         TabIndex        =   6
         Top             =   2250
         Width           =   7695
         Begin VB.Frame Frame1 
            BackColor       =   &H00FF8080&
            BorderStyle     =   0  'None
            Height          =   2655
            Left            =   150
            TabIndex        =   7
            Top             =   240
            Width           =   7425
            Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuda2 
               Height          =   285
               Left            =   3450
               TabIndex        =   2
               Top             =   1380
               Width           =   3645
               _ExtentX        =   6429
               _ExtentY        =   503
               XcodMaxLongitud =   3
               xcodwith        =   200
               NomTabla        =   "vt_vendedor"
               ListaCampos     =   "vendedorcodigo(1),vendedornombres(1)"
               XcodCampo       =   "vendedorcodigo"
               XListCampo      =   "vendedornombres"
               ListaCamposDescrip=   "Codigo,Descripcion"
               ListaCamposText =   "vendedorcodigo,vendedornombres"
               Requerido       =   0   'False
            End
            Begin MSMask.MaskEdBox MBox1 
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "MM/dd/yyyy"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   10250
                  SubFormatType   =   3
               EndProperty
               Height          =   330
               Left            =   3450
               TabIndex        =   1
               Top             =   870
               Width           =   1185
               _ExtentX        =   2090
               _ExtentY        =   582
               _Version        =   393216
               MaxLength       =   10
               Mask            =   "##/##/####"
               PromptChar      =   "_"
            End
            Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuda1 
               Height          =   285
               Left            =   3450
               TabIndex        =   0
               Top             =   390
               Width           =   3645
               _ExtentX        =   6429
               _ExtentY        =   503
               XcodMaxLongitud =   2
               xcodwith        =   150
               NomTabla        =   "cc_tipoplanilla"
               TituloAyuda     =   "Ayuda de Tipo de Planilla"
               ListaCampos     =   "tplanillacodigo(1),tplanilladesccorta(1)"
               XcodCampo       =   "tplanillacodigo"
               XListCampo      =   "tplanilladesccorta"
               ListaCamposDescrip=   "Codigo,Descripcion"
               ListaCamposText =   "tplanillacodigo,tplanilladesccorta"
               Requerido       =   0   'False
            End
            Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuempresa 
               Height          =   315
               Left            =   3360
               TabIndex        =   43
               Top             =   1920
               Visible         =   0   'False
               Width           =   3615
               _ExtentX        =   6376
               _ExtentY        =   556
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
            Begin VB.Label Lblempresa 
               BackStyle       =   0  'Transparent
               Caption         =   "EMPRESA"
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
               Height          =   225
               Left            =   1200
               TabIndex        =   44
               Top             =   2040
               Visible         =   0   'False
               Width           =   2085
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "TIPO DE PLANILLA"
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
               Height          =   225
               Index           =   0
               Left            =   960
               TabIndex        =   10
               Top             =   420
               Width           =   2085
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "FECHA DE PLANILLA"
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
               Height          =   225
               Index           =   1
               Left            =   960
               TabIndex        =   9
               Top             =   960
               Width           =   2085
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "OFICINA / VENDEDOR"
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
               Height          =   225
               Index           =   2
               Left            =   960
               TabIndex        =   8
               Top             =   1500
               Width           =   2085
            End
         End
      End
   End
   Begin MSComctlLib.StatusBar Panel 
      Align           =   2  'Align Bottom
      Height          =   348
      Left            =   0
      TabIndex        =   11
      Top             =   8364
      Width           =   11940
      _ExtentX        =   21061
      _ExtentY        =   609
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
Attribute VB_Name = "FrmPlanillaVarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim adll As New dllgeneral.dll_general
Dim rsdetav As New ADODB.Recordset


Private Sub cAyuda_Click(Index As Integer)
 
  nAyuda = "": nDetalle = ""
 If Index = 0 Then
    'If adll.VerificaDatoExistente(vgcnx, "select * from cc_tipodocumento where tdocumentotipo='C' and tdocumentoingplan='1'") = 1 Then
    If adll.VerificaDatoExistente(VGCNx, "select * from cc_tipodocumento where tdocumentoingplan='1'") = 1 Then
        Dim sfiltra(1, 2) As String
        sfiltra(1, 1) = "Documento": sfiltra(1, 2) = "tdocumentocodigo"
        FrmAyudaCli.TipoForma = 1
        FrmAyudaCli.BConexion = VGCNx
        FrmAyudaCli.Bdata = "0"
        FrmAyudaCli.BTabla = "cc_tipodocumento"
        FrmAyudaCli.BCampos = "tdocumentocodigo as Codigo,tdocumentodescripcion as Descripcion"
        FrmAyudaCli.BOrden = "tdocumentocodigo"
        FrmAyudaCli.BCondi = "tdocumentoingplan='1'"   'tdocumentotipo='C'"
        FrmAyudaCli.BFiltro = sfiltra
        FrmAyudaCli.Show 1
        Text1(1) = nAyuda
        Text1(2).SetFocus
     Else
         nAyuda = "": nDetalle = ""
         MsgBox "No existen Documentos...", vbInformation, MsgTitle
         Exit Sub
     End If
 ElseIf Index = 1 Then
    'If adll.VerificaDatoExistente(vgcnx, "select * from gr_banco where bancocodigo='" & Text1(6) & "'") = 1 Then
         Dim tfiltra(1, 2) As String
         tfiltra(1, 1) = "Banco": tfiltra(1, 2) = "bancodescripcion"
         FrmAyudaCli.TipoForma = 1
         FrmAyudaCli.BConexion = VGCNx
         FrmAyudaCli.Bdata = "0"
         FrmAyudaCli.BTabla = "gr_banco"
         FrmAyudaCli.BCampos = "bancocodigo as Codigo,bancodescripcion as Descripcion"
         FrmAyudaCli.BOrden = "bancocodigo"
         FrmAyudaCli.BCondi = ""
         FrmAyudaCli.BFiltro = tfiltra
         FrmAyudaCli.Show 1
         Text1(6) = nAyuda
     'Else
'         nAyuda = "": nDetalle = ""
'         MsgBox "No existen Bancos...", vbInformation, MsgTitle
'         Exit Sub
'     End If
     
 ElseIf Index = 2 Then
        Dim hfiltra(1, 2) As String
        hfiltra(1, 1) = "Cliente": hfiltra(1, 2) = "clienterazonsocial"
        FrmAyudaCli.TipoForma = 4
        FrmAyudaCli.BConexion = VGCNx
        FrmAyudaCli.Bdata = "2"
        FrmAyudaCli.Bdato = Escadena(Text1(0))
        FrmAyudaCli.BTabla = "vt_cliente"
        FrmAyudaCli.BCampos = "clientecodigo as Codigo,clienterazonsocial as Descripcion"
        FrmAyudaCli.BOrden = "clienterazonsocial"
        FrmAyudaCli.BCondi = ""
        FrmAyudaCli.BFiltro = hfiltra
        FrmAyudaCli.Show 1
        Text1(0) = nAyuda
        Label3 = nDetalle
        Text1(1).SetFocus
   End If
   nAyuda = "": nDetalle = ""

End Sub

Private Sub cmdBotones_Click(Index As Integer)
  Dim acmd As New ADODB.Command
'FIXIT: Declare 'xcargo' and 'xzona' and 'xmone' con un tipo de datos de enlace en tiempo de compilación     FixIT90210ae-R1672-R1B8ZE
  Dim xcargo, xzona, xmone, xcuenta As String
'FIXIT: Declare 'xnumplan' and 'ximpsol' and 'xtcam' con un tipo de datos de enlace en tiempo de compilación     FixIT90210ae-R1672-R1B8ZE
  Dim xnumplan, ximpsol, xtcam, xnumpag As Double
  Dim nflag As Integer
  Dim rb As New ADODB.Recordset
  
  On Error GoTo nerror
  nflag = 0
  Select Case Index
    Case 0
       Limpiartexto Text1, 0, 6
       Text1(0).SetFocus
    Case 1   'GRABAR DATOS
        cmdBotones(1).Enabled = False
        
        If rsdetav.RecordCount > 0 Then
            nflag = 1
            VGCNx.BeginTrans
            
            Set rb = VGCNx.Execute("select * from cc_tipoplanilla where tplanillacodigo='" & Escadena(Ctr_Ayuda1.xclave) & "'")
            If rb.RecordCount > 0 Then
               xnumplan = Val(RTrim$(rb!tplanillanumerador)) + 1
            Else
               xnumplan = 1
            End If
            rb.Close
            Set rb = Nothing
            
            VGCNx.Execute "update cc_tipoplanilla " & _
                       " set tplanillanumerador='" & xnumplan & "' " & _
                       " where tplanillacodigo='" & Escadena(Ctr_Ayuda1.xclave) & "'"
        
        
           rsdetav.MoveFirst
           Do Until rsdetav.EOF
            Set rb = VGCNx.Execute("select * from cc_tipodocumento where tdocumentocodigo='" & rsdetav.Fields("TD") & "'")
            If rb.RecordCount > 0 Then
               xcargo = rb!tdocumentotipo
               
               If rsdetav.Fields("moneda") = g_tiposol Then
                  xcuenta = "" & RTrim$(rb!tdocumentocuentasoles)
               Else
                  xcuenta = "" & RTrim$(rb!tdocumentocuentadolares)
               End If
            Else
               xcargo = "": xcuenta = ""
            End If
            rb.Close
            Set rb = Nothing
                        
          
              Set acmd.ActiveConnection = VGGeneral
              acmd.CommandText = "cc_ingresavarios_pro"
              acmd.CommandType = adCmdStoredProc
              acmd.Prepared = True
              With acmd
                .Parameters("@base") = VGCNx.DefaultDatabase
                .Parameters("@tipo") = "1"
                .Parameters("@tabla") = "vt_cargo"
                .Parameters("@tipodocu") = Escadena(rsdetav.Fields("td"))
                .Parameters("@numero") = Escadena(rsdetav.Fields("serie") & rsdetav.Fields("numero"))
                .Parameters("@cliente") = Escadena(RTrim$(rsdetav.Fields("cliente")))
                .Parameters("@vendedor") = Escadena(Ctr_Ayuda2.xclave)
                .Parameters("@zona") = "01"
                .Parameters("@apefecemi") = rsdetav.Fields("femision")
                .Parameters("@moneda") = rsdetav.Fields("moneda")
                .Parameters("@apeimppag") = CDbl(rsdetav.Fields("importe"))
                .Parameters("@usuario") = g_usuario
                .Parameters("@tipocambio") = 0
                .Parameters("@fechaact") = Now
                .Parameters("@flagcancel") = 0
                .Parameters("@tipoplanilla") = Ctr_Ayuda1.xclave
                .Parameters("@planilla") = xnumplan
                .Parameters("@vencimiento") = rsdetav.Fields("FVencimiento")
                .Parameters("@fechaplani") = MBox1.text
                .Parameters("@banco") = rsdetav.Fields("banco")
                .Parameters("@cargoabono") = xcargo
                .Parameters("@empresa") = Ctr_Ayuempresa.xclave
              End With
              acmd.Execute
              Set acmd = Nothing
              DoEvents
              rsdetav.MoveNext
            Loop
            VGCNx.CommitTrans
            nflag = 0
        End If
       
        rsdetav.Close
        Set rsdetav = Nothing
       
        MsgBox "Los datos han sido grabados satisfactoriamente...!!!", vbInformation, MsgTitle
       
        cmdBotones(1).Enabled = True
        Call adll.ActivaTab(0, 1, SSTab1)
    Case 2
       If TDBGrid1.Row >= 0 Then
         TDBGrid1.Delete
         TDBGrid1.Update
         TDBGrid1.Refresh
       End If
    
    Case 11
      If Len(RTrim$(Ctr_Ayuda1.xclave)) = 0 Then
        MsgBox "Falta Ingresar Tipo de Planilla...Verifique!!", vbInformation, MsgTitle
        Exit Sub
      End If
      If Len(RTrim$(Ctr_Ayuda2.xclave)) = 0 Then
        MsgBox "Falta Ingresar Oficina/Vendedor...Verifique!!", vbInformation, MsgTitle
        Exit Sub
      End If
      If adll.VerificaDatoExistente(VGCNx, "select * from cc_tipoplanilla where tplanilladocvarios='1' and tplanillacodigo='" & Escadena(Ctr_Ayuda1.xclave) & "' ") = 0 Then
            MsgBox "La planilla no es valida para realizar los registros...Verifique!!!", vbInformation, MsgTitle
            Ctr_Ayuda1.SetFocus
            Exit Sub
      End If

      
      Set rsdetav = Nothing
      TDBGrid1.ClearFields
      Set TDBGrid1.DataSource = Nothing
      Call cargar_grilla
       
      Limpiartexto Text1, 0, 6
      Call adll.ActivaTab(1, 1, SSTab1)
      Text1(0).SetFocus
    Case 4
       Call adll.ActivaTab(0, 1, SSTab1)
    Case 12
      Unload Me
  End Select
  
nerror:
  If Err <> 0 Then
    If nflag = 1 Then
      VGCNx.RollbackTrans
    End If
    MsgBox "Error : " & Err.Number & "-" & Err.Description, vbInformation, MsgTitle
    Err = 0
    Exit Sub
    Resume
  End If
End Sub



Private Sub Form_Load()
  MostrarForm Me, "C"
  Call cargar_grilla
  MBox1 = Format(Date, "DD/MM/YYYY")
  Call Ctr_Ayuda1.Conexion(VGCNx)
  Ctr_Ayuda1.filtro = "tplanilladocvarios='1'"
  Call Ctr_Ayuempresa.Conexion(VGCNx)
  Call Ctr_Ayuda2.Conexion(VGCNx)
  If VGParametros.sistemamultiempresas = True Then
     Lblempresa.Visible = True
     Ctr_Ayuempresa.Visible = True
   Else
     Ctr_Ayuempresa.xclave = VGParametros.empresacodigo
  End If
  Call adll.ActivaTab(0, 1, SSTab1)
    
End Sub

Public Sub ConfigGrid()
    With TDBGrid1
        .Columns(0).Width = 1500
        .Columns(1).Width = 3000
        .Columns(2).Width = 700
        .Columns(3).Width = 700
        .Columns(4).Width = 1300
        .Columns(5).Width = 1200
        .Columns(6).Width = 1200
        .Columns(7).Width = 1000
        .Columns(8).Width = 1300
        .Columns(8).NumberFormat = "######,##0.00"
        .Columns(9).Width = 1000
        .Refresh
    End With
End Sub

Public Sub cargar_grilla()
   Set rsdetav = Nothing
   
   Call rsdetav.Fields.Append("Cliente", adChar, 11)
   Call rsdetav.Fields.Append("Descripcion", adChar, 80)
   Call rsdetav.Fields.Append("TD", adChar, 2)
   Call rsdetav.Fields.Append("Serie", adChar, 4)
   Call rsdetav.Fields.Append("Numero", adChar, 10)
   Call rsdetav.Fields.Append("Femision", adDate)
   Call rsdetav.Fields.Append("FVencimiento", adDate)
   Call rsdetav.Fields.Append("Moneda", adChar, 2)
   Call rsdetav.Fields.Append("Importe", adDouble)
   Call rsdetav.Fields.Append("Banco", adChar, 2)
   
   rsdetav.Open
   Set TDBGrid1.DataSource = rsdetav
   TDBGrid1.Refresh
   ConfigGrid
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
'  rsdeta.Close
'  Set rsdeta = Nothing
End Sub

Private Sub MBox1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then
    SendKeys "{tab}"
  End If
End Sub

Private Sub MBox2_GotFocus(Index As Integer)
  Call adll.Enfoquetexto(MBox2(Index))
End Sub

Private Sub MBox2_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then
     If Index Like "[01]" Then
'        If Len(rtrim$(MBox2(1).ClipText)) <> 8 Then
'          MsgBox "Fecha no válida......Verifique!!!", vbInformation, MsgTitle
'          Exit Sub
'        End If
        
     End If
     SendKeys "{tab}"
  End If
End Sub



Private Sub Text1_GotFocus(Index As Integer)
 Dim rb As New ADODB.Recordset
 On Error Resume Next
 
 If Index = 4 Then
   Text1(3) = Right$("000000000" & RTrim$(Text1(3)), Text1(3).MaxLength)
    
    Set rb = VGCNx.Execute("select * from vt_cargo where documentocargo='" & Text1(1) & "' and cargonumdoc='" & RTrim$(Text1(2) & Text1(3)) & "'")
    If rb.RecordCount > 0 Then
      MsgBox "Ya existe el documento...Verifique!!", vbInformation, MsgTitle
      rb.Close
      Set rb = Nothing
      Text1(3).SetFocus
      Exit Sub
    End If
    rb.Close
  End If
  Set rb = Nothing

End Sub
Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
  Dim rb As New ADODB.Recordset
  Dim rclo As New ADODB.Recordset
  Dim ncad As String
'FIXIT: Declare 'J' con un tipo de datos de enlace en tiempo de compilación                FixIT90210ae-R1672-R1B8ZE
  Dim J, flag As Double
  If KeyAscii = 13 Then
     If Index = 6 Then
        If rsdetav.RecordCount > 0 Then
            ncad = rsdetav.Fields(2).Name & "='" & Escadena(Text1(1)) & "' And " & rsdetav.Fields(3).Name & "='" & Escadena(Text1(2)) & "' And " & rsdetav.Fields(4).Name & "='" & Escadena(Text1(3)) & "'"
            Set rclo = rsdetav.Clone(adLockReadOnly)
            rclo.Filter = ncad
            If rclo.RecordCount > 0 Then
                    MsgBox "Ya ingreso el documento....", vbInformation, MsgTitle
                    Limpiartexto Text1, 0, 6
                    Text1(0).SetFocus
                    Set rclo = Nothing
                    Exit Sub
            End If
            Set rclo = Nothing
        End If
        
        Set rb = VGCNx.Execute("select * from gr_banco where bancocodigo='" & RTrim$(Escadena(Text1(Index))) & "'")
        If rb.RecordCount = 0 Then
            MsgBox "Codigo de Banco no existe....Verifique!!!", vbInformation, MsgTitle
            Text1(6) = ""
        End If
        rb.Close
        Set rb = Nothing
        flag = 0
        For J = 0 To 5
          If Len(RTrim$(Text1(J))) = 0 Then
             flag = 1
             Exit For
          End If
        Next J
        If flag = 1 Then
           MsgBox "Falta completar datos...Verifique!!!", vbInformation, MsgTitle
           Exit Sub
        End If
        
        If Len(RTrim$(MBox2(0).ClipText)) <> 8 Or Len(RTrim$(MBox2(1).ClipText)) <> 8 Then
          MsgBox "Fecha no válida......Verifique!!!", vbInformation, MsgTitle
          Exit Sub
        End If
        
        If CDate(MBox2(1)) < CDate(MBox2(0)) Or Len(RTrim$(MBox2(1))) = 0 Then
          MsgBox " La fecha de vencimiento no puede ser menor ....Verifique!!", vbInformation, MsgTitle
          Exit Sub
        End If
        
        rsdetav.AddNew
        rsdetav.Fields(0) = Escadena(Text1(0))
        rsdetav.Fields(1) = Left$(RTrim$(Escadena(Label3)), 80)
        rsdetav.Fields(2) = Escadena(Text1(1))
        rsdetav.Fields(3) = Escadena(Text1(2))
        rsdetav.Fields(4) = Escadena(Text1(3))
        rsdetav.Fields(5) = IIf(IsNull(MBox2(0)), "", MBox2(0))
        rsdetav.Fields(6) = IIf(IsNull(MBox2(1)), "", MBox2(1))
        rsdetav.Fields(7) = Escadena(Text1(4))
        rsdetav.Fields(8) = numero(IIf(IsNull(Text1(5)) Or Len(RTrim$(Text1(5))) = 0, 0, Text1(5)))
        rsdetav.Fields(9) = Escadena(Text1(6))
        rsdetav.Update
        Limpiartexto Text1, 0, 6
        Text1(0).SetFocus
        Exit Sub
     ElseIf Index = 2 Then
         If IsNumeric(Text1(Index)) Then
            Text1(Index) = Right$("000000000" & RTrim$(Text1(Index)), Text1(Index).MaxLength)
         End If
     ElseIf Index = 1 Then
       'Set rb = vgcnx.Execute("select * from cc_tipodocumento where tdocumentocodigo='" & rtrim$(Escadena(Text1(Index))) & "' and tdocumentotipo='C' and tdocumentoingplan='1'")
       Set rb = VGCNx.Execute("select * from cc_tipodocumento where tdocumentocodigo='" & RTrim$(Escadena(Text1(Index))) & "' and tdocumentoingplan='1'")
       If rb.RecordCount = 0 Then
           MsgBox "Documento no válido...Verifique!!!", vbInformation, MsgTitle
           rb.Close
           Set rb = Nothing
           Exit Sub
       End If
       rb.Close
       Set rb = Nothing
     
     ElseIf Index = 0 Then
       If Val(Text1(0)) = 0 And Len(RTrim$(Text1(0))) > 0 Then
         Call cAyuda_Click(2)
         Text1(1).SetFocus
         Exit Sub
       End If
       Set rb = VGCNx.Execute("select * from vt_cliente where clientecodigo='" & RTrim$(Escadena(Text1(Index))) & "'")
       If rb.RecordCount = 0 Then
           MsgBox "Cliente No existe...Verifique!!!", vbInformation, MsgTitle
           rb.Close
           Set rb = Nothing
           Exit Sub
       Else
          Label3 = rb.Fields("clienterazonsocial")
       End If
       rb.Close
       Set rb = Nothing
     ElseIf Index = 3 Then
       If IsNumeric(Text1(Index)) Then
            Text1(Index) = Right$("000000000" & RTrim$(Text1(Index)), Text1(Index).MaxLength)
       End If
       Set rb = VGCNx.Execute("select * from vt_cargo where documentocargo='" & Text1(1) & "' and cargonumdoc='" & RTrim$(Text1(2) & Text1(3)) & "'")
       If rb.RecordCount > 0 Then
         MsgBox "Ya existe el documento...Verifique!!", vbInformation, MsgTitle
         rb.Close
         Set rb = Nothing
         Exit Sub
       End If
       rb.Close
       Set rb = Nothing
     ElseIf Index = 4 Then
       Set rb = VGCNx.Execute("select * from gr_moneda where monedacodigo='" & RTrim$(Escadena(Text1(Index))) & "'")
       If rb.RecordCount = 0 Then
           MsgBox "La moneda no existe....Verifique!!!", vbInformation, MsgTitle
           rb.Close
           Set rb = Nothing
           Exit Sub
       End If
       rb.Close
       Set rb = Nothing
    ElseIf Index = 5 Then
        Text1(Index) = numero(Text1(Index))
    End If
     SendKeys "{tab}"
  End If
  Set rb = Nothing
End Sub

Private Sub Text1_LostFocus(Index As Integer)
 If Index = 3 Or Index = 2 Then
    If IsNumeric(Text1(Index)) Then
        Text1(Index) = Right$("0000000000" & RTrim$(Text1(Index)), Text1(Index).MaxLength)
    End If
 End If
End Sub
