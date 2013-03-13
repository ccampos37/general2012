VERSION 5.00
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form FrmAnulaPllaCanjes 
   Caption         =   "Form1"
   ClientHeight    =   9570
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12330
   LinkTopic       =   "Form1"
   ScaleHeight     =   9570
   ScaleWidth      =   12330
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   9045
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   15954
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   0
      TabCaption(0)   =   "Datos Generales"
      TabPicture(0)   =   "FrmAnulaPllaCanjes.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Detalle"
      TabPicture(1)   =   "FrmAnulaPllaCanjes.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).Control(1)=   "Frame7"
      Tab(1).ControlCount=   2
      Begin VB.Frame Frame1 
         BackColor       =   &H8000000B&
         BorderStyle     =   0  'None
         Caption         =   "sssssssssss"
         ForeColor       =   &H000000FF&
         Height          =   3600
         Left            =   2280
         TabIndex        =   18
         Top             =   2040
         Width           =   7545
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   1
            EndProperty
            Height          =   405
            Left            =   3480
            MaxLength       =   6
            TabIndex        =   28
            Top             =   1320
            Width           =   975
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
            Height          =   300
            Left            =   3480
            TabIndex        =   19
            Top             =   1920
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   0   'False
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin ctrlayuda_f.Ctr_Ayuda Ctr_Tipoplla 
            Height          =   285
            Left            =   3480
            TabIndex        =   20
            Top             =   870
            Width           =   3645
            _ExtentX        =   6429
            _ExtentY        =   503
            XcodMaxLongitud =   2
            xcodwith        =   150
            NomTabla        =   "cp_tipoplanilla"
            TituloAyuda     =   "Ayuda de Tipo de Planilla"
            ListaCampos     =   "tplanillacodigo(1),tplanilladesccorta(1)"
            XcodCampo       =   "tplanillacodigo"
            XListCampo      =   "tplanilladesccorta"
            ListaCamposDescrip=   "Codigo,Descripcion"
            ListaCamposText =   "tplanillacodigo,tplanilladesccorta"
            Requerido       =   0   'False
         End
         Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuda3 
            Height          =   285
            Left            =   3450
            TabIndex        =   21
            Top             =   2385
            Width           =   3645
            _ExtentX        =   6429
            _ExtentY        =   503
            Enabled         =   0   'False
            XcodMaxLongitud =   3
            xcodwith        =   200
            NomTabla        =   "cp_oficina"
            ListaCampos     =   "vendedorcodigo(1),vendedornombres(1)"
            XcodCampo       =   "vendedorcodigo"
            XListCampo      =   "vendedornombres"
            ListaCamposDescrip=   "Codigo,Descripcion"
            ListaCamposText =   "vendedorcodigo,vendedornombres"
            Requerido       =   0   'False
         End
         Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuempresa 
            Height          =   315
            Left            =   3480
            TabIndex        =   22
            Top             =   360
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
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "NUMRERO DE  PLANILLA"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   225
            Index           =   4
            Left            =   960
            TabIndex        =   27
            Top             =   1470
            Width           =   2325
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
            ForeColor       =   &H00FF0000&
            Height          =   225
            Index           =   1
            Left            =   960
            TabIndex        =   26
            Top             =   2040
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
            ForeColor       =   &H00FF0000&
            Height          =   225
            Index           =   0
            Left            =   960
            TabIndex        =   25
            Top             =   1020
            Width           =   2085
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "OFICINA"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   225
            Index           =   3
            Left            =   960
            TabIndex        =   24
            Top             =   2475
            Width           =   2085
         End
         Begin VB.Label Label1 
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
            ForeColor       =   &H00FF0000&
            Height          =   225
            Index           =   2
            Left            =   1080
            TabIndex        =   23
            Top             =   480
            Width           =   1845
         End
      End
      Begin VB.Frame Frame7 
         BorderStyle     =   0  'None
         Height          =   4425
         Left            =   -74670
         TabIndex        =   9
         Top             =   4440
         Width           =   11295
         Begin VB.Frame Frame6 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   840
            Left            =   4440
            TabIndex        =   11
            Top             =   3000
            Width           =   2670
            Begin VB.CommandButton cmdBotones 
               Caption         =   "&Salir"
               Height          =   735
               Index           =   7
               Left            =   1440
               Picture         =   "FrmAnulaPllaCanjes.frx":0038
               Style           =   1  'Graphical
               TabIndex        =   13
               Top             =   120
               Width           =   975
            End
            Begin VB.CommandButton cmdBotones 
               Caption         =   "&Grabar"
               Height          =   735
               Index           =   5
               Left            =   240
               Picture         =   "FrmAnulaPllaCanjes.frx":047A
               Style           =   1  'Graphical
               TabIndex        =   12
               Top             =   120
               Width           =   975
            End
         End
         Begin TrueOleDBGrid70.TDBGrid TDBGrid2 
            Height          =   2565
            Left            =   135
            TabIndex        =   10
            Top             =   390
            Width           =   10845
            _ExtentX        =   19129
            _ExtentY        =   4524
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
            Columns.Count   =   10
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=10"
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
            Splits(0)._ColumnProps(11)=   "Column(2).Width=2725"
            Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=2646"
            Splits(0)._ColumnProps(14)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(16)=   "Column(3).Width=2725"
            Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=2646"
            Splits(0)._ColumnProps(19)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(21)=   "Column(4).Width=2725"
            Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=2646"
            Splits(0)._ColumnProps(24)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(26)=   "Column(5).Width=2725"
            Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=2646"
            Splits(0)._ColumnProps(29)=   "Column(5)._EditAlways=0"
            Splits(0)._ColumnProps(30)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(31)=   "Column(6).Width=2725"
            Splits(0)._ColumnProps(32)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(33)=   "Column(6)._WidthInPix=2646"
            Splits(0)._ColumnProps(34)=   "Column(6)._EditAlways=0"
            Splits(0)._ColumnProps(35)=   "Column(6).Order=7"
            Splits(0)._ColumnProps(36)=   "Column(7).Width=2725"
            Splits(0)._ColumnProps(37)=   "Column(7).DividerColor=0"
            Splits(0)._ColumnProps(38)=   "Column(7)._WidthInPix=2646"
            Splits(0)._ColumnProps(39)=   "Column(7)._EditAlways=0"
            Splits(0)._ColumnProps(40)=   "Column(7).Order=8"
            Splits(0)._ColumnProps(41)=   "Column(8).Width=2725"
            Splits(0)._ColumnProps(42)=   "Column(8).DividerColor=0"
            Splits(0)._ColumnProps(43)=   "Column(8)._WidthInPix=2646"
            Splits(0)._ColumnProps(44)=   "Column(8)._EditAlways=0"
            Splits(0)._ColumnProps(45)=   "Column(8).Order=9"
            Splits(0)._ColumnProps(46)=   "Column(9).Width=2725"
            Splits(0)._ColumnProps(47)=   "Column(9).DividerColor=0"
            Splits(0)._ColumnProps(48)=   "Column(9)._WidthInPix=2646"
            Splits(0)._ColumnProps(49)=   "Column(9)._EditAlways=0"
            Splits(0)._ColumnProps(50)=   "Column(9).Order=10"
            Splits.Count    =   1
            PrintInfos(0)._StateFlags=   3
            PrintInfos(0).Name=   "piInternal 0"
            PrintInfos(0).PageHeaderFont=   "Size=9.75,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=System"
            PrintInfos(0).PageFooterFont=   "Size=9.75,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=System"
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
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=32,.parent=13"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=46,.parent=13"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=43,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=44,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=45,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=50,.parent=13"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=47,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=48,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=49,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=54,.parent=13"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=51,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=52,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=53,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=58,.parent=13"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=55,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=56,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=57,.parent=17"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=62,.parent=13"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=59,.parent=14"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=60,.parent=15"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=61,.parent=17"
            _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=66,.parent=13"
            _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=63,.parent=14"
            _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=64,.parent=15"
            _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=65,.parent=17"
            _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=70,.parent=13"
            _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=67,.parent=14"
            _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=68,.parent=15"
            _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=69,.parent=17"
            _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=74,.parent=13"
            _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=71,.parent=14"
            _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=72,.parent=15"
            _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=73,.parent=17"
            _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=78,.parent=13"
            _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=75,.parent=14"
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
         Begin VB.Shape Shape1 
            BorderColor     =   &H0000C0C0&
            BorderWidth     =   2
            FillColor       =   &H00FFFFC0&
            FillStyle       =   0  'Solid
            Height          =   4035
            Index           =   1
            Left            =   0
            Shape           =   4  'Rounded Rectangle
            Top             =   210
            Width           =   11175
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Total"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   225
            Index           =   1
            Left            =   9300
            TabIndex        =   16
            Top             =   3210
            Width           =   645
         End
         Begin VB.Label Label6 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   1
            Left            =   9840
            TabIndex        =   15
            Top             =   3180
            Width           =   1155
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   "DOCUMENTOS CANJEADOS"
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
            Height          =   195
            Index           =   1
            Left            =   180
            TabIndex        =   14
            Top             =   0
            Width           =   10785
         End
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Height          =   3795
         Left            =   -74640
         TabIndex        =   4
         Top             =   360
         Width           =   11415
         Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
            Height          =   2535
            Left            =   150
            TabIndex        =   5
            Top             =   540
            Width           =   11085
            _ExtentX        =   19553
            _ExtentY        =   4471
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
            Columns.Count   =   10
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=10"
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
            Splits(0)._ColumnProps(11)=   "Column(2).Width=2725"
            Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=2646"
            Splits(0)._ColumnProps(14)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(16)=   "Column(3).Width=2725"
            Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=2646"
            Splits(0)._ColumnProps(19)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(21)=   "Column(4).Width=2725"
            Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=2646"
            Splits(0)._ColumnProps(24)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(26)=   "Column(5).Width=2725"
            Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=2646"
            Splits(0)._ColumnProps(29)=   "Column(5)._EditAlways=0"
            Splits(0)._ColumnProps(30)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(31)=   "Column(6).Width=2725"
            Splits(0)._ColumnProps(32)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(33)=   "Column(6)._WidthInPix=2646"
            Splits(0)._ColumnProps(34)=   "Column(6)._EditAlways=0"
            Splits(0)._ColumnProps(35)=   "Column(6).Order=7"
            Splits(0)._ColumnProps(36)=   "Column(7).Width=2725"
            Splits(0)._ColumnProps(37)=   "Column(7).DividerColor=0"
            Splits(0)._ColumnProps(38)=   "Column(7)._WidthInPix=2646"
            Splits(0)._ColumnProps(39)=   "Column(7)._EditAlways=0"
            Splits(0)._ColumnProps(40)=   "Column(7).Order=8"
            Splits(0)._ColumnProps(41)=   "Column(8).Width=2725"
            Splits(0)._ColumnProps(42)=   "Column(8).DividerColor=0"
            Splits(0)._ColumnProps(43)=   "Column(8)._WidthInPix=2646"
            Splits(0)._ColumnProps(44)=   "Column(8)._EditAlways=0"
            Splits(0)._ColumnProps(45)=   "Column(8).Order=9"
            Splits(0)._ColumnProps(46)=   "Column(9).Width=2725"
            Splits(0)._ColumnProps(47)=   "Column(9).DividerColor=0"
            Splits(0)._ColumnProps(48)=   "Column(9)._WidthInPix=2646"
            Splits(0)._ColumnProps(49)=   "Column(9)._EditAlways=0"
            Splits(0)._ColumnProps(50)=   "Column(9).Order=10"
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
         Begin VB.Shape Shape1 
            BorderColor     =   &H0000C0C0&
            BorderWidth     =   2
            FillColor       =   &H00FFFFC0&
            FillStyle       =   0  'Solid
            Height          =   2925
            Index           =   0
            Left            =   0
            Shape           =   4  'Rounded Rectangle
            Top             =   360
            Width           =   11325
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   "DOCUMENTOS A CANJEAR"
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
            Height          =   195
            Index           =   0
            Left            =   210
            TabIndex        =   8
            Top             =   120
            Width           =   10755
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Total"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   225
            Index           =   0
            Left            =   9270
            TabIndex        =   7
            Top             =   3450
            Width           =   645
         End
         Begin VB.Label Label6 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   0
            Left            =   9840
            TabIndex        =   6
            Top             =   3420
            Width           =   1155
         End
      End
      Begin VB.Frame Frame4 
         Height          =   930
         Left            =   5310
         TabIndex        =   1
         Top             =   6360
         Width           =   1980
         Begin VB.CommandButton cmdBotones 
            Caption         =   "&Cancelar"
            Height          =   690
            Index           =   12
            Left            =   1050
            Picture         =   "FrmAnulaPllaCanjes.frx":08BC
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   180
            Width           =   855
         End
         Begin VB.CommandButton cmdBotones 
            Caption         =   "&Acepta"
            Height          =   690
            Index           =   11
            Left            =   90
            Picture         =   "FrmAnulaPllaCanjes.frx":0CFE
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   180
            Width           =   870
         End
      End
   End
   Begin MSComctlLib.StatusBar Panel 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   17
      Top             =   9225
      Width           =   12330
      _ExtentX        =   21749
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
Attribute VB_Name = "FrmAnulaPllaCanjes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim adll As New dllgeneral.dll_general
Dim rsdetac1 As New ADODB.Recordset
Dim rsdetac2 As New ADODB.Recordset

Public Function Cargar_grilla2()
   Set rsdetac2 = Nothing
   SQL = " select Cliente=clientecodigo,Descripcion=' ',td=documentocargo ,serie=left$(cargonumdoc,3 ),"
   SQL = SQL & " numero=right$(cargonumdoc,8 ) ,FEmision=cargoapefecemi ,FVencimiento=cargoapefecvct ,"
   SQL = SQL & " Moneda=monedacodigo,Importe=cargoapeimpape from cp_cargo where empresacodigo='" & Ctr_Ayuempresa.xclave & "' and "
   SQL = SQL & " abonotipoplanilla='" & Ctr_Tipoplla.xclave & "' and abononumplanilla='" & Format(Text3.Text, "000000") & "'"
   Set rsdetac2 = VGCNx.Execute(SQL)
   Set TDBGrid2.DataSource = rsdetac2
   Call ConfigGrid2
   TDBGrid2.Refresh

End Function

Public Function ConfigGrid2()
    With TDBGrid2
       .Columns(0).Width = 1200
       .Columns(1).Width = 2800
       .Columns(2).Width = 500
       .Columns(3).Width = 700
       .Columns(4).Width = 1100
       .Columns(5).Width = 1100
       .Columns(6).Width = 1200
       .Columns(7).Width = 700
       .Columns(8).Width = 1200
       .Columns(8).NumberFormat = "###,###,##0.00"
       .Refresh
    End With
    
End Function


Private Sub cmdBotones_Click(Index As Integer)
  Dim acmd As New ADODB.Command
  Dim rb As New ADODB.Recordset
  Dim xabono, xzona, xmone, xcuenta, xcargo, xcance As String
  Dim xparcial, xtipo As String
  Dim xnumplan, ximpsol, xtcam, xnumpag As Double
  
  On Error GoTo nerror
  
  Select Case Index

    Case 1   'Eliminar Datos
      If TDBGrid1.ApproxCount > 0 Then
         TDBGrid1.Delete
         TDBGrid1.Update
         TDBGrid1.Refresh
         Call PlanillaTotales(rsdetac1, "importe", Label6(0))
      End If
    Case 2   'Grabar Datos de Documentos a Canjear
      If TDBGrid1.ApproxCount > 0 Then
         Frame7.Enabled = True
         'Text2(0).SetFocus
      Else
         Call adll.ActivaTab(0, 1, SSTab1)
      End If
    Case 5  'Grabar Datos
       'Grabar datos a canjear
         
        If rsdetac1.RecordCount > 0 Then
             rsdetac1.MoveFirst
            Do Until rsdetac1.EOF
                xmone = rsdetac1!moneda
                Set rb = VGCNx.Execute("select * from cp_cargo where empresacodigo='" & Ctr_Ayuempresa.xclave & "' and  documentocargo='" & rsdetac1.Fields(2) & "' and cargonumdoc='" & rsdetac1.Fields(3) & rsdetac1.Fields(4) & "' and clientecodigo='" & rsdetac1.Fields(0) & "'")
                If rb.RecordCount > 0 Then
                  xnumpag = rb!cargoapenumpag
                 Else
                   xzona = "01": xmone = g_TipoSol: xnumpag = 1: xparcial = ""
                End If
                
                ximpsol = CDbl(rsdetac1.Fields("importe"))
                xtcam = DatoTipoCambio(VGcnxCT, MBox1.Text)               'TraeTipoCambio(Date, VGcnx)
                If rsdetac1.Fields("moneda") <> xmone Then
                   If rsdetac1.Fields("moneda") = g_TipoSol Then
                      ximpsol = CDbl(rsdetac1.Fields("importe")) / CDbl(xtcam)
                   Else
                      ximpsol = CDbl(rsdetac1.Fields("importe")) * CDbl(xtcam)
                   End If
                End If

                Set acmd = Nothing
                DoEvents

                '**** Actualizamos Saldos de documento pendiente
                If rsdetac1.Fields("moneda") = g_TipoDolar Then
                   If xmone = g_TipoSol Then
                       VGCNx.Execute "Update  cp_cargo Set cargoapeimppag=isnull(cargoapeimppag,0) -" & CDbl(rsdetac1.Fields("importe") / xtcam) & "," & _
                                   " cargoapenumpag='" & xnumpag - 1 & "'" & _
                                  " Where empresacodigo='" & Ctr_Ayuempresa.xclave & "' and documentocargo='" & rsdetac1.Fields(2) & "' and cargonumdoc='" & Trim$(rsdetac1.Fields(3) & rsdetac1.Fields(4)) & "' and " & _
                                  " clientecodigo='" & Trim$(rsdetac1.Fields("Cliente")) & "'"
                   Else
                       VGCNx.Execute "Update  cp_cargo Set cargoapeimppag=isnull(cargoapeimppag,0) -" & CDbl(rsdetac1.Fields("importe")) & "," & _
                                  " cargoapenumpag='" & xnumpag - 1 & "'" & _
                                  " Where  empresacodigo='" & Ctr_Ayuempresa.xclave & "' and  documentocargo='" & rsdetac1.Fields(2) & "' and cargonumdoc='" & Trim$(rsdetac1.Fields(3) & rsdetac1.Fields(4)) & "' and " & _
                                  " clientecodigo='" & Trim$(rsdetac1.Fields("Cliente")) & "'"
                   End If
                ElseIf rsdetac1.Fields("moneda") = g_TipoSol Then
                   If xmone = g_TipoDolar Then
                       VGCNx.Execute "Update  cp_cargo Set cargoapeimppag=isnull(cargoapeimppag,0) - " & CDbl(rsdetac1.Fields("importe") * xtcam) & "," & _
                                  " cargoapenumpag='" & xnumpag - 1 & "'" & _
                                  " Where  empresacodigo='" & Ctr_Ayuempresa.xclave & "' and  documentocargo='" & rsdetac1.Fields(2) & "' and cargonumdoc='" & Trim$(rsdetac1.Fields(3) & rsdetac1.Fields(4)) & "' and " & _
                                  " clientecodigo='" & Trim$(rsdetac1.Fields("Cliente")) & "'"
                   Else
                       VGCNx.Execute "Update  cp_cargo Set cargoapeimppag=isnull(cargoapeimppag,0) - " & CDbl(rsdetac1.Fields("importe")) & "," & _
                                  " cargoapenumpag='" & xnumpag - 1 & "'" & _
                                  " Where  empresacodigo='" & Ctr_Ayuempresa.xclave & "' and documentocargo='" & rsdetac1.Fields(2) & "' and cargonumdoc='" & Trim$(rsdetac1.Fields(3) & rsdetac1.Fields(4)) & "' and " & _
                                  " clientecodigo='" & Trim$(rsdetac1.Fields("Cliente")) & "'"
                   End If
                End If

                VGCNx.Execute "Update  cp_cargo " & _
                            " Set cargoapeflgcan= CASE isnull(cargoapeimpape,0)-isnull(cargoapeimppag,0) WHEN 0 THEN '1' ELSE '0' END ," & _
                            "   cargoapefeccan='" & Date & "'" & _
                            " Where  empresacodigo='" & Ctr_Ayuempresa.xclave & "' and documentocargo='" & rsdetac1.Fields(2) & "' and cargonumdoc='" & Trim$(rsdetac1.Fields(3) & rsdetac1.Fields(4)) & "' and " & _
                            " clientecodigo='" & Trim$(rsdetac1.Fields("Cliente")) & "'"
                            
                SQL = "delete  cp_abono Where  empresacodigo='" & Ctr_Ayuempresa.xclave & "' and documentoabono='" & rsdetac1.Fields(2) & "' and abononumdoc='" & Trim$(rsdetac1.Fields(3) & rsdetac1.Fields(4)) & "' and  "
                SQL = SQL & " abonocancli='" & Trim$(rsdetac1.Fields("Cliente")) & "'"
                Set rb = VGCNx.Execute(SQL)
                rsdetac1.MoveNext
           Loop
        Else
            MsgBox "No existen datos...Verifique!!", vbInformation, MsgTitle
            Exit Sub
        End If
        
       'Grabar datos de Documentos Canjeados
        If rsdetac2.RecordCount > 0 Then
           rsdetac2.MoveFirst
           Do Until rsdetac2.EOF
                SQL = "delete  cp_cargo Where  empresacodigo='" & Ctr_Ayuempresa.xclave & "' and documentocargo='" & rsdetac2.Fields(2) & "' and cargonumdoc='" & Trim$(rsdetac2.Fields(3) & rsdetac2.Fields(4)) & "' and  "
                SQL = SQL & " clientecodigo='" & Trim$(rsdetac2.Fields("Cliente")) & "'"
                Set rb = VGCNx.Execute(SQL)
            rsdetac2.MoveNext
            Loop
        End If
        
        rsdetac1.Close
        Set rsdetac1 = Nothing
       
        rsdetac2.Close
        Set rsdetac2 = Nothing
       
        MsgBox "Los datos han sido grabados satisfactoriamente...!!!", vbInformation, MsgTitle
        
        Call adll.ActivaTab(0, 1, SSTab1)
    

    Case 7
        Call adll.ActivaTab(0, 1, SSTab1)
    Case 11
      If Len(Trim$(Ctr_Tipoplla.xclave)) = 0 Then
        MsgBox "Falta Ingresar Tipo de Planilla...Verifique!!", vbInformation, MsgTitle
        Exit Sub
      End If
      'If Len(trim$(Ctr_Ayuda2.xclave)) = 0 Then
      '  MsgBox "Falta Ingresar Oficina/Vendedor...Verifique!!", vbInformation, MsgTitle
      '  Exit Sub
      'End If
      If Len(Trim$(Ctr_Ayuda3.xclave)) = 0 Then
        MsgBox "Falta Ingresar Oficina/Vendedor...Verifique!!", vbInformation, MsgTitle
        Exit Sub
      End If
      If adll.VerificaDatoExistente(VGCNx, "select * from cp_tipoplanilla where tplanillacanjes='1' and tplanillacodigo='" & Escadena(Ctr_Tipoplla.xclave) & "' ") = 0 Then
            MsgBox "La planilla no es valida para realizar los canjes...Verifique!!!", vbInformation, MsgTitle
            Ctr_Tipoplla.SetFocus
            Exit Sub
      End If

      Set rsdetac1 = Nothing
      TDBGrid1.ClearFields
      Set TDBGrid1.DataSource = Nothing
      Call cargar_grilla
      
      Set rsdetac2 = Nothing
      TDBGrid2.ClearFields
      Set TDBGrid2.DataSource = Nothing
      Call Cargar_grilla2
       
      Label6(0) = "": Label6(1) = ""
      
      Call adll.ActivaTab(1, 1, SSTab1)
      'Text1(0).SetFocus
    Case 12
      Unload Me
  End Select
Exit Sub
nerror:
  If Err Then
    'MsgBox "Error : " & Err.Number & "-" & Err.Description, vbInformation, MsgTitle
    Err = 0
    Exit Sub
  End If
End Sub

Private Sub Form_Load()
  MostrarForm Me, "C"
  
  MBox1 = Format(Date, "DD/MM/YYYY")
  Call Ctr_Tipoplla.conexion(VGCNx)
  Call Ctr_Ayuda3.conexion(VGCNx)
  Call Ctr_Ayuempresa.conexion(VGCNx)
  Ctr_Tipoplla.Filtro = "tplanillacanjes='1' or tplanillarenovar='1'"
  If VGparametros.sistemamultiempresas = True Then
     Ctr_Ayuempresa.Visible = True
     Label1(3).Visible = True
   Else
     Ctr_Ayuempresa.xclave = "01"
     Ctr_Ayuempresa.Visible = False
     Label1(3).Visible = False
  End If
  Call adll.ActivaTab(0, 1, SSTab1)
End Sub

Public Sub ConfigGrid()
   With TDBGrid1
       .Columns(0).Width = 1200
       .Columns(1).Width = 2800
       .Columns(2).Width = 500
       .Columns(3).Width = 700
       .Columns(4).Width = 1100
       .Columns(5).Width = 1100
       .Columns(6).Width = 1200
       .Columns(7).Width = 700
       .Columns(8).Width = 1200
       .Columns(8).NumberFormat = "###,###,##0.00"
       .Refresh
   End With
End Sub

Public Sub cargar_grilla()
   Set rsdetac1 = Nothing
   SQL = " select Cliente=abonocancli,Descripcion=' ',td=documentoabono ,serie=left$(abononumdoc,3 ),"
   SQL = SQL & " numero=right$(abononumdoc,8 ) ,FEmision=abonocanfecan ,FVencimiento=abonocanfecan ,"
   SQL = SQL & " Moneda=abonocanmoncan,Importe= abonocanimpcan  from cp_abono where empresacodigo='" & Ctr_Ayuempresa.xclave & "' and "
   SQL = SQL & " abonotipoplanilla='" & Ctr_Tipoplla.xclave & "' and abononumplanilla='" & Format(Text3.Text, "000000") & "'"
   
   Set rsdetac1 = VGCNx.Execute(SQL)
   Set TDBGrid1.DataSource = rsdetac1
   Call ConfigGrid
   TDBGrid1.Refresh
End Sub
Private Sub MBox1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then
     SendKeys "{tab}"
  End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
Dim rss As New ADODB.Recordset
If KeyAscii = 13 Then
   SQL = " select top 1 * from cp_abono where empresacodigo='" & Ctr_Ayuempresa.xclave & "' and "
   SQL = SQL & " abonotipoplanilla='" & Ctr_Tipoplla.xclave & "' and abononumplanilla='" & Format(Text3.Text, "000000") & "'"
   Set rss = VGCNx.Execute(SQL)
   If rss.RecordCount = 0 Then
      MsgBox (" No existe Numero de Planilla ")
      Exit Sub
   End If
   MBox1 = rss!abonocanfecpla
   Ctr_Ayuda3.xclave = rss!vendedorcodigo: Ctr_Ayuda3.Ejecutar
End If
End Sub
